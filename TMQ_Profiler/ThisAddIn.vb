Imports System.IO
Imports System.Net
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Data.OleDb
Imports System.Data
Imports System.Collections
Imports System.Text.RegularExpressions

Public Module GlobalVariables

    Public SVBI_Execute_String, SVDG_Execute_String, ChosenLoop, dangerzone, ControlLocation, RequestItemCheck, codeword, currentfile, ErrorMessage, adds, notadds, Item_Found, mv, mvbi, senderemail, ST, Subject, StartTimer, EndTimer As String
    Public SVBI_Connection, SVBI_Connection_Alt, SVBI_Connection_Alt_Two, SVBI_Connection_Persistant, SVDG_Connection As New ADODB.Connection
    Public destfolder As Outlook.MAPIFolder
    Public attachments As Outlook.Attachments
    Public rs, rs_alt, rs_alt2, rs_alt3, rsbi As ADODB.Recordset
    Public length, pricesum As Integer
    Public Desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\"
    Public outlookNameSpace As Outlook.NameSpace
    Public inbox, OBox As Outlook.MAPIFolder
    Public reply, reply_DQ As Outlook.MailItem

End Module

Public Module PresetValues

    Public Deactivate_Trigger As String = "No"
    Public Alt_Trigger As String = "TMQDATA"
    Public Version As String = "22/10/2018 - Clean up"
    Public price As Integer = CInt(Math.Ceiling(Rnd() * 10)) + 1
    Public desktopchanger As String = vbYes
    Public DisplayModel As String = vbNo
    Public EmailName As String = "Yes"
    'Public SVTS_CP As String = "ClaimsProcessor@edumail.vic.gov.au"
    Public offer As String = "offer"
    Public everything As String = "everything"
    Public PDP2018 As String = "PDP18"
    Public profile As String = "profile"
    Public case_history As String = "casehistory"
    Public WaitingOnSend As String = "Okay"
    Public risk As String = "Risk"
    Public facts As String = "facts"
    Public TestS As String = vbNo
    Public FastFix As String = vbNo
    Public TestStatus As String = "Failed on VBNET error"
    Public program As String = "program"
    Public Uplift As String = "Uplift"
    Public RTOprogram As String = "RTOProgram"
    Public DestFolderName As String = "TMSProfiler"
    Public SVBI_Risk As String = "Provider='SQLOLEDB';" _
                                  & "Data Source='svbidb01\svbidb01';" _
                                  & "Initial Catalog ='SVTS_Risk';" _
                                  & "Integrated Security ='SSPI';"
    Public SVDG_CaseTracker As String = "Provider='SQLOLEDB';" _
                                      & "Data Source='PRWDSVDG01,1435';" _
                                      & "Initial Catalog ='VPMS';" _
                                      & "Integrated Security ='SSPI';"
    Public ProfileOnly As String = "1"
    Public OnTime As String = "<B>Turned on:</b> " & Now.ToString
End Module

Public Class ThisAddIn

    Public PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    Public WithEvents Items As Outlook.Items
    Public WithEvents Inspectors As Outlook.Inspectors

    Shared Function UNameWindows() As String

        UNameWindows = Environ("USERNAME")

    End Function

    Public Sub WaitOnSend()
        WaitingOnSend = "Yes"
        Dim Counter As Long
        For Counter = 0 To 10000
            System.Windows.Forms.Application.DoEvents()
        Next Counter
        WaitingOnSend = "No"

    End Sub

    Public Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
Resetter:
        If TestS = vbYes Then Call TestZone()
        If TestS = vbYes Then MsgBox("Test Status: " + TestStatus)
        TestS = vbNo

        Dim oForm As TMQProfilerSetup
        oForm = New TMQProfilerSetup
        oForm.Name = "TMSProfileSetUp"
        oForm.ShowDialog()
        oForm.Activate()
        oForm = Nothing
        Inspectors = Me.Application.Inspectors
        outlookNameSpace = Me.Application.GetNamespace("MAPI")
        inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        Items = inbox.Items
        Dim destfolder = inbox.Folders(DestFolderName)
        On Error Resume Next
        inbox.Folders.Add(DestFolderName)
        If TestS = vbYes Then GoTo Resetter
    End Sub
    Private Sub myOlApp_ItemSend(ByVal Item As Object, Cancel As Boolean)

        WaitingOnSend = "No"

    End Sub


    Public Sub TestZone()
        'Use this to test new functionality, DB pulls etc

        reply_DQ = Application.CreateItem(Outlook.OlItemType.olMailItem)
        reply_DQ.DeleteAfterSubmit = True
        reply_DQ.To = UNameWindows()
        reply_DQ.Subject = "[Unclassified: For Official Use Only] TMS Profiler: Test email"
        reply_DQ.HTMLBody = "Diagnostics: <br><br>" & OnTime & "<br><br>" & "<b>Current trigger word: </b>" & codeword
        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>Version: " & Version
        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "This automated email is a test email regarding new functionality to be rolled into the next release of the TMS Profiler"
        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><br>"

        'Start test code

        'End test code
        reply_DQ.Send()
        Call WaitOnSend()
        TestStatus = "Test passed!"
    End Sub
    Public Sub Items_ItemAdd(ByVal item As Object) Handles Items.ItemAdd
        If WaitingOnSend = "Yes" Then Exit Sub
        If FastFix = vbYes Then Exit Sub
        On Error GoTo earlyexit
        destfolder = inbox.Folders(DestFolderName)
        StartTimer = "<B>Start:</b> " & Now.ToString
        EndTimer = "<BR><BR><B>End:</b> "

        If TypeOf (item) IsNot Outlook.MailItem Then Exit Sub
        Dim mail As Outlook.MailItem = item
        'If mail.MessageClass IsNot IPM.Note Then Exit Sub

        Dim sender As Outlook.AddressEntry = mail.Sender
        Dim exchUser As Outlook.ExchangeUser = sender.GetExchangeUser()

        On Error GoTo exiter
        If mail.Subject.Length = 0 Then
exiter:
            Exit Sub
        End If

        On Error GoTo earlyexit

        Subject = RTrim(LTrim(Replace(Replace(Subject, "<", ""), ">", "")))

        If mail.SenderEmailType = "EX" Then
            If sender IsNot Nothing Then
                If sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry OrElse sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
                    If exchUser IsNot Nothing Then
                        senderemail = exchUser.PrimarySmtpAddress
                    Else
                        senderemail = Nothing
                    End If
                Else
                    senderemail = TryCast(sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS), String)
                End If
            Else
                senderemail = Nothing
            End If
        Else
            senderemail = mail.SenderEmailAddress
        End If

        If mail.Subject.ToUpper.Contains("TMS_OFF") Then
            Call TMS_Off(mail)
            Exit Sub
        ElseIf mail.Subject.ToUpper.Contains("TMS_ON") Then
            Call TMS_On(mail)
            Exit Sub
        End If

        If Deactivate_Trigger = "Yes" Then Exit Sub

        If ProfileOnly = 1 Then GoTo NoScan

        'First trigger a check for automated emails needed (always do this)
        If Today.DayOfWeek = DayOfWeek.Monday And Now().Hour >= 8 And Now().Minute > 30 Then
            Call SVTS_BDACheck()
        ElseIf Today.DayOfWeek = DayOfWeek.Friday And Now().Hour >= 8 And Now().Minute > 30 Then
            Call SVTS_SUPLift()
        End If
        If Now().Hour >= 8 Then Call SVTS_PaymentSpike()

        Dim CR As String
        CR = "renshaw.carol.l@edumail.vic.gov.au"
        Dim Siam As String = "Luangpithathorn"
        Dim ASIC As String = "Alerts"
        Dim ASIC_Alt As String = "Digest"

        'ASQA Notification
        Dim ASQAString As String = "From: Commissioners"
        Dim ASQAString_Alt As String = "ASQA"
        If mail.Body.ToUpper.Contains(ASQAString.ToUpper) And mail.Body.ToUpper.Contains(ASQAString_Alt.ToUpper) And senderemail.ToUpper.Contains(CR.ToUpper) Then
            Call ASQA_CTInsert(mail)
            Exit Sub
        ElseIf mail.Body.ToUpper.Contains(ASIC.ToUpper) And mail.Body.ToUpper.Contains(ASIC_Alt.ToUpper) And senderemail.ToUpper.Contains(Siam.ToUpper) Then
            Call ASIC_CTInsert(mail)
            Exit Sub
        End If
        Alt_Trigger = Replace(codeword.ToUpper, "S", "Q")
NoScan:

        If Right(senderemail, 19) <> "@edumail.vic.gov.au" Then Exit Sub

        If Not (mail.Subject.ToUpper.Contains(codeword.ToUpper) Or mail.Subject.ToUpper.Contains(Alt_Trigger.ToUpper)) Then Exit Sub

        'Create the reply

        reply = Application.CreateItem(Outlook.OlItemType.olMailItem)
        reply.DeleteAfterSubmit = True
        reply.To = senderemail
        reply.CC = UNameWindows()
        reply.Subject = "[Unclassified: For Official Use Only] Automated reply: TMS Profiler"

        'Check individual user permissions now

        If SVDG_Connection.State = 1 Then SVDG_Connection.Close()
        SVDG_Connection.Open(SVDG_CaseTracker)

        SVDG_Execute_String = "SELECT COUNT(*) as Matched FROM VPMS.dbo.Officers where EmailAddress = '" & senderemail & "'"
        rs = SVDG_Connection.Execute(SVDG_Execute_String)
        mv = rs("Matched").Value

        If mv > 0 Then GoTo NoProblem
        'The below bounces back due to failed authentication

        reply.HTMLBody =
                                         "<br>Request: " + mail.Subject +
                                         "<br>Recieved: " & mail.ReceivedTime &
                                         "<br>Sent: " & DateTime.Today.ToString &
                                         "<br>Logged: SVTS_Risk.Printer.Requestlog<br><br>"
        reply.HTMLBody = reply.HTMLBody + "<P STYLE='font-family:Calbri;font-size:12'>Thank you for your email.<br><br>" _
                                 & "It appears that you have insufficient access rights to recieve the data you requested. This is controlled by HESG staff as listed in the case tracker (the list resides at VPMS.DBO.Officers).<br><br>" +
                                 "<br>If you feel this is in error, please let me know. <br><br>Cheers, <br><br>Lance<br>"

        Call WaitOnSend()
        reply.Send()
        mail.Move(destfolder)

        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
        SVBI_Connection.Open(SVBI_Risk)
        SVBI_Execute_String = "INSERT INTO SVTS_Risk.Printer.RequestLog
							( RequestID,
							  Requestor ,
							  Request_Text,
							  Request_Date
									  )
							  VALUES( '" & mv + 1 & "', '" & senderemail & "' , -- Requestor - varchar(50)
							 Request was refused - requestor had no permission. " & Subject & "' , -- Request_Text - varchar(max)
							 CAST(GETDATE() As Date)  -- Request_Date - Date
							)"
        SVBI_Connection.Execute(SVBI_Execute_String)
        Call WriteExitReason()
        Exit Sub

NoProblem:
        'Mail type, sender etc all check out.

        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
        SVBI_Connection.Open(SVBI_Risk)
        SVBI_Execute_String = "INSERT INTO SVTS_Risk.Printer.RequestPrice
											  ( Requestor,
												Price,
												BilledDate
											  )
											  VALUES  ( '" & senderemail & "' , -- Requestor
														'" & price & "' , -- Request_Text - varchar(max)
														   CAST(GETDATE() as Date)  -- BilledDate - date
											  )"
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        rs = Nothing
        SVBI_Execute_String = "SELECT SUM(Price) as Price from SVTS_Risk.Printer.RequestPrice WHERE Requestor = '" & senderemail & "'"
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        pricesum = rs("Price").Value

        Item_Found = Trim(StrReverse(Left(Trim(StrReverse(mail.Subject.ToUpper)), InStr(Trim(StrReverse(mail.Subject.ToUpper)), " "))))

        mv = vbEmpty
        rs = Nothing

        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()

        SVBI_Connection.Open(SVBI_Risk)

        reply.To = mail.Sender.Address

        reply.HTMLBody = "<P STYLE='font-family:Calbri;font-size:12'>Re: " & mail.Subject & "<br>Recieved: " & mail.ReceivedTime &
                                     "<br>Logged: SVTS_Risk.Printer.Requestlog<br><hr><br>You requested the following files:"

        SVBI_Execute_String = "Select max([RequestID]) RID from SVTS_Risk.Printer.RequestLog"
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        mv = rs("RID").Value
        SVBI_Execute_String = "INSERT INTO SVTS_Risk.Printer.RequestLog
										  ( RequestID,
											Requestor ,
											Request_Text ,
											Request_Date
										  )
										  VALUES  ( '" & mv + 1 & "', '" & senderemail & "' , -- Requestor - varchar(50)
										  '" & Subject & "' , -- Request_Text - varchar(max)
											CAST(GETDATE() as Date)  -- Request_Date - date
										  )"
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        rs = Nothing
        mail.Move(destfolder)
        Threading.Thread.Sleep(100)
        mail.UnRead = False
        Threading.Thread.Sleep(100)

        Dim VRQA As String = "VRQA"

        If mail.Subject.ToUpper.Contains(Uplift.ToUpper) Then
            CreateUplift(Item_Found)
        ElseIf mail.Subject.ToUpper.Contains(VRQA.ToUpper) Then
            CreateVRQA(LTrim(RTrim(Item_Found)))
        ElseIf mail.Subject.ToUpper.Contains(RTOprogram.ToUpper) Then
            CreateRTOProgram(Item_Found)

            'For PSP when needed
            'ElseIf mail.Subject.ToUpper.Contains(profile.ToUpper) Then
            '     Create2018Profile(Item_Found, EmailName)

        ElseIf mail.Subject.ToUpper.Contains(case_history.ToUpper) Then
            CreateCaseHistories(Item_Found, senderemail)
        ElseIf mail.Subject.ToUpper.Contains(program.ToUpper) Then
            CreateProgram(Item_Found)
        ElseIf mail.Subject.ToUpper.Contains(risk.ToUpper) Then
            CreateRisk(LTrim(RTrim(Item_Found)))
        ElseIf mail.Subject.ToUpper.Contains(PDP2018.ToUpper) Then
            CreatePDP2018(LTrim(RTrim(Item_Found)))
        End If
        Exit Sub

earlyexit:
        Call WriteExitReason()
        reply = Nothing
        reply = Application.CreateItem(Outlook.OlItemType.olMailItem)
        reply.DeleteAfterSubmit = True
        reply.To = UNameWindows()
        reply.Subject = "[Unclassified: For Official Use Only] Automated reply: TMS Profiler"
        reply.HTMLBody = "Error in TMS profiler application. Details follow " & Err.Description.ToString & ""
        reply.HTMLBody = reply.HTMLBody & "<br><br>" & Version
        reply.Send()
        Call WaitOnSend()
        Exit Sub

    End Sub
    Sub SVTS_BDACheck()
        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
        SVBI_Connection.Open(SVBI_Risk)
        SVBI_Execute_String = "SELECT COUNT(*) as Matched FROM SVTS_Risk.printer.BDAFlags ADQ WHERE informed = 'No' AND ADQ.FlaggingReason != ''"
        mv = 0
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        mv = rs("Matched").Value

        If mv > 0 Then

            SVBI_Execute_String = "SELECT Targets FROM SVTS_Risk.Printer.DistList WHERE ListName = 'PaymentWatch'"
            rs = SVBI_Connection.Execute(SVBI_Execute_String)
            mv = rs("Targets").Value
            reply_DQ = Application.CreateItem(Outlook.OlItemType.olMailItem)
            reply_DQ.DeleteAfterSubmit = True
            reply_DQ.To = mv
            reply_DQ.CC = UNameWindows()
            reply_DQ.Subject = "[Unclassified: For Official Use Only] Automated Data Alert: TMS Profiler"
            reply_DQ.HTMLBody = "<P STYLE='font-family:Calbri; font-size:12'>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "This automated email is about provider growth ahead of BDA..."
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><br>" & "Note: Comparative payments in the below are calendar (training) year on year." & "<br><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>" & "<a href=" & "\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & ">Risk reports are here!</a>" & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>"

            Call BDA_Loop()

            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>If you have any questions, feel free to hit up Lance!"
            reply_DQ.Send()
            Call WaitOnSend()
            rs = SVBI_Connection.Execute(SVBI_Execute_String)

        Else
        End If
        SVBI_Execute_String = "Update SVTS_Risk.Printer.BDAFlags SET informed = 'Yes' WHERE informed like '%No%'"
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()


    End Sub

    Sub BDA_Loop()
        'HERE
        SVBI_Execute_String = "SELECT b.* FROM SVTS_Risk.Printer.BDAFlags b left join svts_risk.ref.TAFE T on t.toid = b.toid WHERE T.Toid is null and b.informed LIKE '%no%' and b.Status like '%monitored%' and b.FlaggingReason != '' Order by b.Status asc, b.TOID asc"
        If SVBI_Connection_Persistant.State = 1 Then SVBI_Connection_Persistant.Close()
        SVBI_Connection_Persistant.Open(SVBI_Risk)
        rs_alt2 = SVBI_Connection_Persistant.Execute(SVBI_Execute_String)
        'Watch
        If Not rs_alt2.EOF Then
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14'><font color = 'darkred'><b>Non-TAFES of interest - currently flagged as 'Watch':</font>"
        End If

        'On Error GoTo ErrorCatcher
        If 1 = 1 Then

            'On Error Resume Next
            Do While Not rs_alt2.EOF
                RiskQuickRun(rs_alt2.Fields("TOID").Value.ToString())
                SVBI_Execute_String = "SELECT TOP 1 * FROM SVTS_Risk.Printer.PaymentFlags ADQ where toid = " & rs_alt2.Fields("TOID").Value.ToString() & " and logdate = '" & rs_alt2.Fields("LogDate").Value.ToString() & "'"
                If SVBI_Connection_Alt.State = 1 Then SVBI_Connection_Alt.Close()
                SVBI_Connection_Alt.Open(SVBI_Risk)
                rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "</b><P STYLE='font-family:Calbri; font-size:13'><b>TOID: " & "<a href=\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & rs_alt.Fields("TOID").Value.ToString() & "_Profile_2018.pdf>" & rs_alt.Fields("TOID").Value.ToString() & " - " & rs_alt.Fields("ShortName").Value.ToString() & "</a></b><br>"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<font color = 'darkred'>Provider status: " & rs_alt.Fields("Status").Value.ToString() & "<br></font>"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & rs_alt2.Fields("FlaggingReason").Value.ToString() & "<br>"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated due for payment by next BDA (2018): " & Convert.ToDecimal(rs_alt2.Fields("PayPaid_CY").Value.ToString()).ToString("c") & " exc GST<br>"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated due for payment by next BDA (2017): " & Convert.ToDecimal(rs_alt2.Fields("PaidToDate_PY").Value.ToString()).ToString("c") & " exc GST<br>"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Change year on year: " & Convert.ToDecimal(rs_alt2.Fields("Change_YoY").Value.ToString()).ToString("c") & " (" & rs_alt2.Fields("ChangeShare_YoY").Value.ToString() & "%) exc GST<br><br>"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pending (active claims): " & Convert.ToDecimal(rs_alt.Fields("Payments - Pending").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Pending").Value.ToString()).ToString("N0") & " hours)<br>"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>"
                rs_alt2.MoveNext()
            Loop
        End If
        'Normal Status
        If SVBI_Connection_Persistant.State = 1 Then SVBI_Connection_Persistant.Close()
        SVBI_Connection_Persistant.Open(SVBI_Risk)
        SVBI_Execute_String = "SELECT b.* FROM SVTS_Risk.Printer.BDAFlags b left join svts_risk.ref.TAFE T on t.toid = b.toid WHERE T.Toid is null and b.informed LIKE '%no%' and b.Status like '%system payment%' and b.FlaggingReason != '' Order by b.Status asc,  b.TOID asc"
        rs_alt2 = SVBI_Connection_Persistant.Execute(SVBI_Execute_String)

        If Not rs_alt2.EOF Then
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14'><b>Non-TAFES of (potential) interest - no current payment flags:"
        End If
        'MsgBox("Non-TAFE - N")
        Do While Not rs_alt2.EOF
            RiskQuickRun(rs_alt2.Fields("TOID").Value.ToString())
            SVBI_Execute_String = "SELECT TOP 1 * FROM SVTS_Risk.Printer.PaymentFlags ADQ where toid = " & rs_alt2.Fields("TOID").Value.ToString() & " and logdate = '" & rs_alt2.Fields("LogDate").Value.ToString() & "'"
            If SVBI_Connection_Alt.State = 1 Then SVBI_Connection_Alt.Close()
            SVBI_Connection_Alt.Open(SVBI_Risk)
            rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "</b><P STYLE='font-family:Calbri; font-size:13'><b>TOID: " & "<a href=\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & rs_alt2.Fields("TOID").Value.ToString() & "_Profile_2018.pdf>" & rs_alt2.Fields("TOID").Value.ToString() & " - " & rs_alt2.Fields("ShortName").Value.ToString() & "</a></b><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<font color = 'DARKGREEN'>Provider status: " & rs_alt2.Fields("Status").Value.ToString() & "<br></font>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & rs_alt2.Fields("FlaggingReason").Value.ToString() & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated due for payment by next BDA (2018): " & Convert.ToDecimal(rs_alt2.Fields("PayPaid_CY").Value.ToString()).ToString("c") & " exc GST<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated due for payment by next BDA (2017): " & Convert.ToDecimal(rs_alt2.Fields("PaidToDate_PY").Value.ToString()).ToString("c") & " exc GST<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Change year on year: " & Convert.ToDecimal(rs_alt2.Fields("Change_YoY").Value.ToString()).ToString("c") & " (" & rs_alt2.Fields("ChangeShare_YoY").Value.ToString() & "%) exc GST<br><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pending (active claims): " & Convert.ToDecimal(rs_alt.Fields("Payments - Pending").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Pending").Value.ToString()).ToString("N0") & " hours)<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>"
            rs_alt2.MoveNext()
        Loop

        If SVBI_Connection_Persistant.State = 1 Then SVBI_Connection_Persistant.Close()
        SVBI_Connection_Persistant.Open(SVBI_Risk)
        SVBI_Execute_String = "SELECT b.* FROM SVTS_Risk.Printer.BDAFlags b left join svts_risk.ref.TAFE T on t.toid = b.toid WHERE T.Toid is null and b.informed LIKE '%no%' and Status not like '%monitored%' and Status  != 'Defaults to system payment rules' and b.FlaggingReason != '' Order by b.Status asc,  b.TOID asc"
        rs_alt2 = SVBI_Connection_Persistant.Execute(SVBI_Execute_String)

        'Weird Status
        If Not rs_alt2.EOF Then
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14'><font color = 'Red'><b>Non-TAFES of (potential) interest with a flagged status:</font>"
        End If

        'MsgBox("Non-TAFE - Weird")
        Do While Not rs_alt2.EOF
            RiskQuickRun(rs_alt2.Fields("TOID").Value.ToString())
            SVBI_Execute_String = "SELECT TOP 1 * FROM SVTS_Risk.Printer.PaymentFlags ADQ where toid = " & rs_alt2.Fields("TOID").Value.ToString() & " and logdate = '" & rs_alt2.Fields("LogDate").Value.ToString() & "'"
            If SVBI_Connection_Alt.State = 1 Then SVBI_Connection_Alt.Close()
            SVBI_Connection_Alt.Open(SVBI_Risk)
            rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "</b><P STYLE='font-family:Calbri; font-size:13'><b>TOID: " & "<a href=\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & rs_alt.Fields("TOID").Value.ToString() & "_Profile_2018.pdf>" & rs_alt.Fields("TOID").Value.ToString() & " - " & rs_alt.Fields("ShortName").Value.ToString() & "</a></b><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<font color = 'darkred'>Provider status: " & rs_alt2.Fields("Status").Value.ToString() & "<br></font>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & rs_alt2.Fields("FlaggingReason").Value.ToString() & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated due for payment by next BDA (2018): " & Convert.ToDecimal(rs_alt2.Fields("PayPaid_CY").Value.ToString()).ToString("c") & " exc GST<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated due for payment by next BDA (2017): " & Convert.ToDecimal(rs_alt2.Fields("PaidToDate_PY").Value.ToString()).ToString("c") & " exc GST<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Change year on year: " & Convert.ToDecimal(rs_alt2.Fields("Change_YoY").Value.ToString()).ToString("c") & " (" & rs_alt2.Fields("ChangeShare_YoY").Value.ToString() & "%) exc GST<br><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pending (active claims): " & Convert.ToDecimal(rs_alt.Fields("Payments - Pending").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Pending").Value.ToString()).ToString("N0") & " hours)<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>"
            rs_alt2.MoveNext()
        Loop

        'TAFES
        If SVBI_Connection_Persistant.State = 1 Then SVBI_Connection_Persistant.Close()
        SVBI_Connection_Persistant.Open(SVBI_Risk)
        SVBI_Execute_String = "SELECT b.* FROM SVTS_Risk.Printer.BDAFlags b left join svts_risk.ref.TAFE T on t.toid = b.toid WHERE T.Toid is not null and b.informed LIKE '%no%' and b.FlaggingReason != '' Order by b.Status asc, b.FlaggingReason ASc, b.TOID asc"
        rs_alt2 = SVBI_Connection_Persistant.Execute(SVBI_Execute_String)

        If Not rs_alt2.EOF Then
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14'><font color = 'navy'><b>TAFEs of (potential) interest:</font>"
        End If
        'MsgBox("TAFE - N")
        Do While Not rs_alt2.EOF
            RiskQuickRun(rs_alt2.Fields("TOID").Value.ToString())
            SVBI_Execute_String = "SELECT TOP 1 * FROM SVTS_Risk.Printer.PaymentFlags ADQ where toid = " & rs_alt2.Fields("TOID").Value.ToString() & " and logdate = '" & rs_alt2.Fields("LogDate").Value.ToString() & "'"
            If SVBI_Connection_Alt.State = 1 Then SVBI_Connection_Alt.Close()
            SVBI_Connection_Alt.Open(SVBI_Risk)
            rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "</b><P STYLE='font-family:Calbri; font-size:13'><b>TOID: " & "<a href=\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & rs_alt2.Fields("TOID").Value.ToString() & "_Profile_2018.pdf>" & rs_alt2.Fields("TOID").Value.ToString() & " - " & rs_alt2.Fields("ShortName").Value.ToString() & "</a></b><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<font color = 'navy'>Provider status: " & rs_alt.Fields("Status").Value.ToString() & "<br></font>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & rs_alt2.Fields("FlaggingReason").Value.ToString() & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated due for payment by next BDA (2018): " & Convert.ToDecimal(rs_alt2.Fields("PayPaid_CY").Value.ToString()).ToString("c") & " exc GST<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated due for payment by next BDA (2017): " & Convert.ToDecimal(rs_alt2.Fields("PaidToDate_PY").Value.ToString()).ToString("c") & " exc GST<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Change year on year: " & Convert.ToDecimal(rs_alt2.Fields("Change_YoY").Value.ToString()).ToString("c") & " (" & rs_alt2.Fields("ChangeShare_YoY").Value.ToString() & "%) exc GST<br><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pending (active claims): " & Convert.ToDecimal(rs_alt.Fields("Payments - Pending").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Pending").Value.ToString()).ToString("N0") & " hours)<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>"
            rs_alt2.MoveNext()
        Loop

    End Sub
    Sub SVTS_SUPLift()

        'HERE
        Dim BaseString As String
        BaseString = "SELECT DISTINCT SL.TOID, SL.TradingName FROM SVTS_Risk.Printer.SupLift SL WHERE sl.Notified LIKE '%No%'"
        If SVBI_Connection_Persistant.State = 1 Then SVBI_Connection_Persistant.Close()
        SVBI_Connection_Persistant.Open(SVBI_Risk)
        rs_alt2 = SVBI_Connection_Persistant.Execute(BaseString)

        If rs_alt2.EOF Then Exit Sub

        reply_DQ = Application.CreateItem(Outlook.OlItemType.olMailItem)
        reply_DQ.DeleteAfterSubmit = True
        reply_DQ.To = "purcell.myra.e@edumail.vic.gov.au"
        reply_DQ.CC = UNameWindows()
        reply_DQ.Subject = "[Unclassified: For Official Use Only] TMS Automated Summary: SUPLift"
        reply_DQ.HTMLBody = "<P STYLE='font-family:Calbri; font-size:12'>"
        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "This automated email provides a break down of new (and old) activity under SUPLift <Br><BR>"

        If SVBI_Connection_Alt.State = 1 Then SVBI_Connection_Alt.Close()
        SVBI_Execute_String = "SELECT COUNT(DISTINCT SL.TOID) AS Providers, COUNT(DISTINCT SL.RTOReferenceStudentID) AS Students, SUM(SL.Paid_Pay) AS Paid_Pay, SUM(sl.Pending) AS Pending FROM SVTS_Risk.Printer.SupLift SL where SL.Type != 'Summary' "
        SVBI_Connection_Alt.Open(SVBI_Risk)
        rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)

        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<B>Program summary: <Br></b>"
        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Providers: " & rs_alt.Fields("Providers").Value.ToString() & ". Students: " & rs_alt.Fields("Students").Value.ToString() & "<BR>"
        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Paid: " & Convert.ToDecimal(rs_alt.Fields("Paid_Pay").Value.ToString()).ToString("c") & ". Pending: " & Convert.ToDecimal(rs_alt.Fields("Pending").Value.ToString()).ToString("c") & "."
        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><BR><hr>"

        Do While Not rs_alt2.EOF
            If SVBI_Connection_Alt.State = 1 Then SVBI_Connection_Alt.Close()
            SVBI_Execute_String = "SELECT * FROM SVTS_Risk.Printer.SupLift SL WHERE SL.Type = 'Summary' AND SL.TOID = '" & rs_alt2.Fields("TOID").Value.ToString & "'"
            SVBI_Connection_Alt.Open(SVBI_Risk)
            rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)
            Dim StudentFinder As String
            StudentFinder = "SELECT SL.TOID, COUNT(Distinct SL.RTOReferenceStudentID) as Students FROM SVTS_Risk.Printer.SupLift SL WHERE SL.Type != 'Summary' AND SL.TOID = '" & rs_alt2.Fields("TOID").Value.ToString & "' GROUP BY SL.TOID"
            If SVBI_Connection_Alt_Two.State = 1 Then SVBI_Connection_Alt_Two.Close()
            SVBI_Connection_Alt_Two.Open(SVBI_Risk)
            rs_alt3 = SVBI_Connection_Alt_Two.Execute(StudentFinder)

            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<b>Overall summary for " & "" & rs_alt.Fields("TOID").Value.ToString() & " - " & rs_alt.Fields("TradingName").Value.ToString() & "</a></b><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Students: " & rs_alt3.Fields("Students").Value.ToString() & ". Total hours: " & rs_alt.Fields("ScheduledHours").Value.ToString() & "<br>" & "Paid (to date): " & Convert.ToDecimal(rs_alt.Fields("Paid_Pay").Value.ToString()).ToString("c") & ". Pending: " & Convert.ToDecimal(rs_alt.Fields("Pending").Value.ToString()).ToString("c") & "."
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>"

            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><hr><br>"
            rs_alt2.MoveNext()
        Loop

        reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>If you have any questions, feel free to hit up Lance!"
        reply_DQ.Send()
        Call WaitOnSend()

        'Clean Up
        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
        SVBI_Connection.Open(SVBI_Risk)
        SVBI_Execute_String = "Update SVTS_Risk.Printer.SupLift SET Notified = 'Yes' WHERE Notified like '%No%'"
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
        If SVBI_Connection_Persistant.State = 1 Then SVBI_Connection_Persistant.Close()


    End Sub
    Sub SVTS_PaymentSpike()

        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
        SVBI_Connection.Open(SVBI_Risk)
        SVBI_Execute_String = "SELECT COUNT(*) as Matched FROM SVTS_Risk.Printer.PaymentFlags ADQ WHERE informed = 'No' and flaggingreason != ''"
        mv = 0
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        mv = rs("Matched").Value

        If mv > 0 Then
            If SVBI_Connection_Alt.State = 1 Then SVBI_Connection_Alt.Close()
            SVBI_Connection_Alt.Open(SVBI_Risk)
            SVBI_Execute_String = "SELECT Targets FROM SVTS_Risk.Printer.DistList WHERE ListName = 'PaymentWatch'"
            rs = SVBI_Connection.Execute(SVBI_Execute_String)
            mv = rs("Targets").Value

            reply_DQ = Application.CreateItem(Outlook.OlItemType.olMailItem)
            reply_DQ.DeleteAfterSubmit = True
            reply_DQ.To = mv
            reply_DQ.CC = UNameWindows()
            reply_DQ.Subject = "[Unclassified: For Official Use Only] Automated Data Alert: TMS Profiler"
            reply_DQ.HTMLBody = "<P STYLE='font-family:Calbri; font-size:12'>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "This automated email is being sent as one or more providers have appeared to have had something interesting occur when claims processor ran over their last upload for yesterday."
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><br>" & "Note: Payments in the below are calendar (training) year on year." & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br>" & "<a href=" & "\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & ">Risk reports are here!</a>" & "<br>"

            SVBI_Execute_String = "SELECT * FROM SVTS_Risk.Printer.TotalPay ORDER BY Sector asc"
            rs = SVBI_Connection.Execute(SVBI_Execute_String)

            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14; margin:0'><b>Next BDA payment data summary (" & rs.Fields("Monthname").Value.ToString() & "): </b><P STYLE='font-family:Calbri; font-size:12'; margin:0>"
            Dim X As String

            Do While Not rs.EOF
                reply_DQ.HTMLBody = "" & reply_DQ.HTMLBody & rs.Fields("Sector").Value.ToString() & "&nbsp;"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Estimated Pay: " & Convert.ToDecimal(rs.Fields("PayValue").Value.ToString()).ToString("c") & " exc GST ("
                If InStr(rs.Fields("Change").Value.ToString(), "-") > 0 Then
                    X = "<font color='red'>" & rs.Fields("Change").Value.ToString() & "% </font> year on year)<br>"
                Else
                    X = "<font color='green'>" & rs.Fields("Change").Value.ToString() & "% </font> year on year)<br>"
                End If
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & X
                rs.MoveNext()
            Loop
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<hr>"

            Dim SVBI_Table As String
            SVBI_Table = "SVTS_Risk.Printer.PaymentFlags"
            Call PaymentReportLoop(SVBI_Table)

            SVBI_Execute_String = "SELECT * FROM SVTS_Risk.Printer.TotalTYD ORDER BY Sector asc"
            rs = SVBI_Connection.Execute(SVBI_Execute_String)

            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<hr>" & "<P STYLE='font-family:Calbri; font-size:14'><b>Estimated year on year payment summary (following " & rs.Fields("Monthname").Value.ToString() & " BDA): </b><P STYLE='font-family:Calbri; font-size:12'>"

            Do While Not rs.EOF
                reply_DQ.HTMLBody = "" & reply_DQ.HTMLBody & rs.Fields("Sector").Value.ToString() & "&nbsp;"
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Paid and Pay (YTD): " & Convert.ToDecimal(rs.Fields("PayValue").Value.ToString()).ToString("c") & " exc GST ("
                If InStr(rs.Fields("Change").Value.ToString(), "-") > 0 Then
                    X = "<font color='red'>" & rs.Fields("Change").Value.ToString() & "% </font> year on year)<br>"
                Else
                    X = "<font color='green'>" & rs.Fields("Change").Value.ToString() & "% </font> year on year)<br>"
                End If
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & X
                rs.MoveNext()
            Loop

            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<BR><hr><P STYLE='font-family:Calbri; font-size:14'><b>New addition: Providers with registration ending soon </b><P STYLE='font-family:Calbri; font-size:12'><br>"
            If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
            SVBI_Connection.Open(SVBI_Risk)
            SVBI_Execute_String = "SELECT 'Rating ' + CASE WHEN o.RegistrationStatus NOT IN ('Current','Current (Re-registration pending)') THEN 'High' 
WHEN o.RegistrationStatus IN ('Current (Re-registration pending)') THEN 'Medium'
ELSE 'Low' END + ': ' +
	   E.TradingName + ' (' + CAST(E.TOID AS VARCHAR(MAX)) +')<br>' +'
	   Status: ' + RegistrationStatus + ' - (end date: ' +  CAST(E.TGA_ExpiryDate AS VARCHAR(MAX)) + ', ' + CAST(DaysToExpiry AS VARCHAR(MAX)) + ')<br> 
	   Current students: ' + CAST(E.CurrentStudents AS VARCHAR(MAX)) + '<br>
	   Last contract: ' + CAST(E.LatestContractYear AS VARCHAR(MAX))+ ' ('+ E.ContractStatus +') <BR><br>' AS CareFactor
FROM SVTS_Risk.dbo.vw_CaseTracker_ExpiredRegistrations E
INNER JOIN SVTS_TGA.dbo.Organisation o ON o.OrganisationCode = e.TOID
LEFT JOIN [SVTS_TGA].[dbo].[RTORegistrationPeriod] rrp ON rrp.OrganisationID = o.OrganisationID AND rrp.EndDate = e.TGA_ExpiryDate
ORDER BY CASE WHEN o.RegistrationStatus NOT IN ('Current','Current (Re-registration pending)') THEN 1
WHEN o.RegistrationStatus IN ('Current (Re-registration pending)') THEN 2
ELSE 3 END  ASC"
            rs = SVBI_Connection.Execute(SVBI_Execute_String)

            Do While Not rs.EOF

                reply_DQ.HTMLBody = reply_DQ.HTMLBody & rs.Fields("CareFactor").Value.ToString()
                rs.MoveNext()
            Loop

            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<hr><P STYLE='font-family:Calbri; font-size:12'><br>If you have any questions, feel free to hit up Lance!"
            reply_DQ.Send()
            Call WaitOnSend()

        Else
        End If
        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()

    End Sub

    Sub PaymentReportLoop(SVBI_Table As String)

        SVBI_Execute_String = "SELECT * FROM " & SVBI_Table & " ADQ JOIN svts_Risk.ref.NonTAFE t ON t.TOID = ADQ.TOID WHERE informed = 'No' and Status like '%monitored%' and FlaggingReason != '' Order by Status asc, ADQ.FlaggingReason ASc, t.TOID asc"
        If SVBI_Connection_Alt.State = 1 Then SVBI_Connection_Alt.Close()
        SVBI_Connection_Alt.Open(SVBI_Risk)
        rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)

        'Non-TAFES - Watch Status

        If Not rs_alt.EOF Then
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14'><font color = 'darkred'><b>Non-TAFES of interest - currently flagged as 'Watch':</font>"
        End If

        Do While Not rs_alt.EOF
            RiskQuickRun(rs_alt.Fields("TOID").Value.ToString())
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "</b><P STYLE='font-family:Calbri; font-size:13'><b>TOID: " & "<a href=\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & rs_alt.Fields("TOID").Value.ToString() & "_Profile_2018.pdf>" & rs_alt.Fields("TOID").Value.ToString() & " - " & rs_alt.Fields("ShortName").Value.ToString() & "</a></b><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<font color = 'darkred'>Provider status: " & rs_alt.Fields("Status").Value.ToString() & "<br></font>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Flagging reason: " & rs_alt.Fields("FlaggingReason").Value.ToString() & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Last processed: " & rs_alt.Fields("LogDate").Value.ToString() & " (compare date: " & rs_alt.Fields("COmpareDate").Value.ToString() & ")<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Paid (active claims): " & "" & Convert.ToDecimal(rs_alt.Fields("Payments - Paid").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Paid").Value.ToString()).ToString("N0") & " hours)<br>"
            If Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") > 0 Then
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pay (active claims): <font color='green'>" & Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") & "</font> exc GST (<font color='green'>" & Convert.ToDouble(rs_alt.Fields("Hours - Pay").Value.ToString()).ToString("N0") & " </font>hours)<br>"
            Else
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pay (active claims): <font color='red'>" & Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") & "</font> exc GST (<font color='red'>" & Convert.ToDouble(rs_alt.Fields("Hours - Pay").Value.ToString()).ToString("N0") & " </font>hours)<br>"
            End If
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pending (active claims): " & Convert.ToDecimal(rs_alt.Fields("Payments - Pending").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Pending").Value.ToString()).ToString("N0") & " hours)<br>"

            If Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") > 0 Then
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<i>Estimated pay (as of next BDA): <font color='green'>" & Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") & "</font> exc GST (<font color='green'>" & Convert.ToDouble(rs_alt.Fields("Next - Hours").Value.ToString()).ToString("N0") & "</font> hours)</i>"
            Else
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<i>Estimated pay (as of next BDA): <font color='red'>" & Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") & "</font> exc GST (<font color='red'>" & Convert.ToDouble(rs_alt.Fields("Next - Hours").Value.ToString()).ToString("N0") & "</font> hours)</i>"
            End If
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><font color='black'>"
            rs_alt.MoveNext()
        Loop

        'Non-TAFES - Normal Status

        SVBI_Execute_String = "SELECT * FROM " & SVBI_Table & " ADQ JOIN svts_Risk.ref.NonTAFE t ON t.TOID = ADQ.TOID WHERE informed = 'No' and Status = 'Defaults to system payment rules' and FlaggingReason != '' Order by Status asc, ADQ.FlaggingReason ASc, t.TOID asc"
        rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)

        If Not rs_alt.EOF Then
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14'><b>Non-TAFES of (potential) interest - no current payment flags:"
        End If

        Do While Not rs_alt.EOF
            RiskQuickRun(rs_alt.Fields("TOID").Value.ToString())
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "</b><P STYLE='font-family:Calbri; font-size:13'><b>TOID: " & "<a href=\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & rs_alt.Fields("TOID").Value.ToString() & "_Profile_2018.pdf>" & rs_alt.Fields("TOID").Value.ToString() & " - " & rs_alt.Fields("ShortName").Value.ToString() & "</a></b><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Flagging reason: " & rs_alt.Fields("FlaggingReason").Value.ToString() & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Last processed: " & rs_alt.Fields("LogDate").Value.ToString() & " (compare date: " & rs_alt.Fields("COmpareDate").Value.ToString() & ")<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Paid (active claims): " & "" & Convert.ToDecimal(rs_alt.Fields("Payments - Paid").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Paid").Value.ToString()).ToString("N0") & " hours)<br>"
            If Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") > 0 Then
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pay (active claims): <font color='green'>" & Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") & "</font> exc GST (<font color='green'>" & Convert.ToDouble(rs_alt.Fields("Hours - Pay").Value.ToString()).ToString("N0") & " </font>hours)<br>"
            Else
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pay (active claims): <font color='red'>" & Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") & "</font> exc GST (<font color='red'>" & Convert.ToDouble(rs_alt.Fields("Hours - Pay").Value.ToString()).ToString("N0") & " </font>hours)<br>"
            End If
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pending (active claims): " & Convert.ToDecimal(rs_alt.Fields("Payments - Pending").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Pending").Value.ToString()).ToString("N0") & " hours)<br>"

            If Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") > 0 Then
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<i>Estimated pay (as of next BDA): <font color='green'>" & Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") & "</font> exc GST (<font color='green'>" & Convert.ToDouble(rs_alt.Fields("Next - Hours").Value.ToString()).ToString("N0") & "</font> hours)</i>"
            Else
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<i>Estimated pay (as of next BDA): <font color='red'>" & Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") & "</font> exc GST (<font color='red'>" & Convert.ToDouble(rs_alt.Fields("Next - Hours").Value.ToString()).ToString("N0") & "</font> hours)</i>"
            End If
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><font color='black'>"
            rs_alt.MoveNext()
        Loop
        SVBI_Execute_String = "SELECT * FROM " & SVBI_Table & " ADQ JOIN svts_Risk.ref.NonTAFE t ON t.TOID = ADQ.TOID WHERE informed = 'No' and Status not like '%monitored%' and Status  != 'Defaults to system payment rules'  and FlaggingReason != '' Order by Status desc, ADQ.FlaggingReason ASc, t.TOID asc"
        rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)

        'Non-TAFES - Weird Status
        If Not rs_alt.EOF Then
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14'><font color = 'Red'><b>Non-TAFES of (potential) interest with a flagged status:</font>"
        End If

        Do While Not rs_alt.EOF
            RiskQuickRun(rs_alt.Fields("TOID").Value.ToString())
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "</b><P STYLE='font-family:Calbri; font-size:13'><b>TOID: " & "<a href=\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & rs_alt.Fields("TOID").Value.ToString() & "_Profile_2018.pdf>" & rs_alt.Fields("TOID").Value.ToString() & " - " & rs_alt.Fields("ShortName").Value.ToString() & "</a></b><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Provider status:<font color='FireBrick'> " & rs_alt.Fields("Status").Value.ToString() & "</font><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Flagging reason: " & rs_alt.Fields("FlaggingReason").Value.ToString() & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Last processed: " & rs_alt.Fields("LogDate").Value.ToString() & " (compare date: " & rs_alt.Fields("COmpareDate").Value.ToString() & ")<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Paid (active claims): " & "" & Convert.ToDecimal(rs_alt.Fields("Payments - Paid").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Paid").Value.ToString()).ToString("N0") & " hours)<br>"
            If Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") > 0 Then
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pay (active claims): <font color='green'>" & Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") & "</font> exc GST (<font color='green'>" & Convert.ToDouble(rs_alt.Fields("Hours - Pay").Value.ToString()).ToString("N0") & " </font>hours)<br>"
            Else
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pay (active claims): <font color='red'>" & Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") & "</font> exc GST (<font color='red'>" & Convert.ToDouble(rs_alt.Fields("Hours - Pay").Value.ToString()).ToString("N0") & " </font>hours)<br>"
            End If
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pending (active claims): " & Convert.ToDecimal(rs_alt.Fields("Payments - Pending").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Pending").Value.ToString()).ToString("N0") & " hours)<br>"

            If Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") > 0 Then
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<i>Estimated pay (as of next BDA): <font color='green'>" & Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") & "</font> exc GST (<font color='green'>" & Convert.ToDouble(rs_alt.Fields("Next - Hours").Value.ToString()).ToString("N0") & "</font> hours)</i>"
            Else
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<i>Estimated pay (as of next BDA): <font color='red'>" & Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") & "</font> exc GST (<font color='red'>" & Convert.ToDouble(rs_alt.Fields("Next - Hours").Value.ToString()).ToString("N0") & "</font> hours)</i>"
            End If
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><font color='black'>"
            rs_alt.MoveNext()
        Loop

        'TAFES

        SVBI_Execute_String = "SELECT * FROM " & SVBI_Table & " ADQ JOIN svts_Risk.ref.TAFE t ON t.TOID = ADQ.TOID WHERE informed = 'No' and FlaggingReason != '' Order by ADQ.FlaggingReason ASc, t.TOID asc"
        rs_alt = SVBI_Connection_Alt.Execute(SVBI_Execute_String)

        If Not rs_alt.EOF Then
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<P STYLE='font-family:Calbri; font-size:14'><font color = 'navy'><b>TAFEs of (potential) interest:</font>"
        End If

        Do While Not rs_alt.EOF
            RiskQuickRun(rs_alt.Fields("TOID").Value.ToString())
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "</b><P STYLE='font-family:Calbri; font-size:13'><b>TOID: " & "<a href=\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\" & rs_alt.Fields("TOID").Value.ToString() & "_Profile_2018.pdf>" & rs_alt.Fields("TOID").Value.ToString() & " - " & rs_alt.Fields("ShortName").Value.ToString() & "</a></b><br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Flagging reason: " & rs_alt.Fields("FlaggingReason").Value.ToString() & "<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Last processed: " & rs_alt.Fields("LogDate").Value.ToString() & " (compare date: " & rs_alt.Fields("COmpareDate").Value.ToString() & ")<br>"
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Paid (active claims): " & "" & Convert.ToDecimal(rs_alt.Fields("Payments - Paid").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Paid").Value.ToString()).ToString("N0") & " hours)<br>"
            If Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") > 0 Then
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pay (active claims): <font color='green'>" & Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") & "</font> exc GST (<font color='green'>" & Convert.ToDouble(rs_alt.Fields("Hours - Pay").Value.ToString()).ToString("N0") & " </font>hours)<br>"
            Else
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pay (active claims): <font color='red'>" & Convert.ToDecimal(rs_alt.Fields("Payments - Pay").Value.ToString()).ToString("c") & "</font> exc GST (<font color='red'>" & Convert.ToDouble(rs_alt.Fields("Hours - Pay").Value.ToString()).ToString("N0") & " </font>hours)<br>"
            End If
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "Pending (active claims): " & Convert.ToDecimal(rs_alt.Fields("Payments - Pending").Value.ToString()).ToString("c") & " exc GST (" & Convert.ToDouble(rs_alt.Fields("Hours - Pending").Value.ToString()).ToString("N0") & " hours)<br>"

            If Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") > 0 Then
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<i>Estimated pay (as of next BDA): <font color='green'>" & Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") & "</font> exc GST (<font color='green'>" & Convert.ToDouble(rs_alt.Fields("Next - Hours").Value.ToString()).ToString("N0") & "</font> hours)</i>"
            Else
                reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<i>Estimated pay (as of next BDA): <font color='red'>" & Convert.ToDecimal(rs_alt.Fields("Next - Payments").Value.ToString()).ToString("c") & "</font> exc GST (<font color='red'>" & Convert.ToDouble(rs_alt.Fields("Next - Hours").Value.ToString()).ToString("N0") & "</font> hours)</i>"
            End If
            reply_DQ.HTMLBody = reply_DQ.HTMLBody & "<br><font color='black'>"
            rs_alt.MoveNext()
        Loop

        SVBI_Execute_String = "Update " & SVBI_Table & " SET informed = 'Yes' WHERE informed like '%No%'"
        rs = SVBI_Connection_Alt.Execute(SVBI_Execute_String)

    End Sub
    Sub RiskQuickRun(Item_Found As String)

        EmailName = "No"

        CreateRisk(LTrim(RTrim(Item_Found)))
        Dim NewLocation As String
        Dim FolderString As String
        FolderString = "\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\" & "DailyEmail\" & Today.ToShortDateString & "\"
        NewLocation = FolderString & Item_Found & "_Profile_2018.pdf"

        If Not Directory.Exists(FolderString) Then
            Directory.CreateDirectory(FolderString)
        End If

        If File.Exists("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\All\" & Item_Found & "_Profile_2018.pdf") And Not File.Exists(NewLocation) Then
            File.Copy("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\All\" & Item_Found & "_Profile_2018.pdf", NewLocation)
        End If

        EmailName = "Yes"

    End Sub

    Sub WriteExitReason()

        If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
        SVBI_Connection.Open(SVBI_Risk)
        SVBI_Execute_String = "INSERT INTO svts_risk.Printer.VSTOError 
							( ErrorMessage,
							  EduMailUser ,
							  VersionNumber
							)
							  VALUES( 
							  'Requested by: " & senderemail & "' 
							, '" & UNameWindows.ToString & "'
							, '" & Version & " - " & Err.Description.ToString & "'
							)"
        SVBI_Connection.Execute(SVBI_Execute_String)


    End Sub
    Sub CreateCaseHistories(RequestItem As String, EmailName As String)

        Dim FromPath As String = "\\education.vic.gov.au\SHARE\TMO\Vet\Division VET\RTO Case Tracker\Production\"
        Dim AccessDB As String = "VPMS.accde"
        Dim Report As String = "OutputRpt"
        Dim RemoveDB As String = "Yes"
        Dim ReplaceDB As String = "No"

        SVDG_Execute_String = "SELECT DISTINCT TOID FROM VPMS.dbo.RTOCase_Case WHERE TOID = " & RequestItem
        Call SVDG_Check(SVDG_Execute_String)
        If RequestItemCheck = RequestItem Then
            Call CreateAccessReports(FromPath & AccessDB, AccessDB, Report, RequestItem, RemoveDB, ReplaceDB)
        End If
    End Sub

    Sub CreateASQANotification()

        'Note this is point to DEV

        Dim FromPath As String = "\\education.vic.gov.au\SHARE\TMO\Vet\Division VET\RTO Case Tracker\Production\"
        Dim AccessDB As String = "VPMS.accde"
        Dim Report As String = "ASQANotifications"
        Dim RemoveDB As String = "Yes"
        Dim ReplaceDB As String = "No"
        Dim RequestItem As String = ""

        Call CreateAccessReports(FromPath & AccessDB, AccessDB, Report, RequestItem, RemoveDB, ReplaceDB)

    End Sub
    Sub CreateCERS(RequestItem As String)

        Dim FromPath As String = "\\education.vic.gov.au\share\TMO\Projects\CERS\4_Backup\"
        Dim AccessDB As String = "CERS_2018.accde"
        Dim Report As String = "OutputRpt"
        Dim RemoveDB As String = "Yes"
        Dim ReplaceDB As String = "\\education.vic.gov.au\share\TMO\Projects\CERS\1_Production\CERS_2018.accde"

        Call CreateAccessReports(FromPath & AccessDB, AccessDB, Report, RequestItem, RemoveDB, ReplaceDB)

    End Sub

    Sub CreateAccessReports(FromPath As String, AccessDB As String, Report As String, RequestItem As String, RemoveDB As String, ReplaceDB As String)

        Dim sKill As String
        Dim nnn As Integer = 0
        Dim FSO As Object = CreateObject("scripting.filesystemobject")

        Do While nnn < 2
            sKill = "TASKKILL /F /IM MSACCESS.EXE"
            Shell(sKill, vbHide)
            nnn = nnn + 1
        Loop

        If File.Exists(Desktop & AccessDB) Then Kill(Desktop & AccessDB)
        Dim toPath As String = Desktop & AccessDB
        FSO.CopyFile(Source:=FromPath, Destination:=toPath)

        nnn = 0

        Threading.Thread.Sleep(500)
        On Error Resume Next
        Dim objAccess = CreateObject("Access.Application")
        objAccess.OpenCurrentDatabase(Desktop & AccessDB, False)
        objAccess.SetWarnings = False
        objAccess.DisplayAlerts = False
        objAccess.Run(Report, RequestItem)
        objAccess.CloseCurrentDatabase
        objAccess.application.Quit
        objAccess = Nothing
        Threading.Thread.Sleep(500)

AccessProblem:

        Do While nnn < 2
            sKill = "TASKKILL /F /IM MSACCESS.EXE"
            Shell(sKill, vbHide)
            nnn = nnn + 1
        Loop

        Threading.Thread.Sleep(500)

        On Error Resume Next
        If RemoveDB = "Yes" Then
            If File.Exists(Desktop & AccessDB) Then Kill(Desktop & AccessDB)
        End If

        If ReplaceDB = "No" Then Exit Sub
        If File.Exists(ReplaceDB) Then Exit Sub
        FSO.CopyFile(Source:=FromPath, Destination:=ReplaceDB)
    End Sub

    Sub ExcelFileLooper(RequestItem As String, CurrentFile As String, ControlLocation As String)

        Dim sKill As String
        Dim xxx As Integer = 0
        Dim FromPath As String
        Dim FSO As Object = CreateObject("scripting.filesystemobject")

        If File.Exists(Desktop & CurrentFile) Then Kill(Desktop & CurrentFile)
        If File.Exists(ControlLocation & CurrentFile) Then
            FSO.CopyFile(Source:=ControlLocation & CurrentFile, Destination:=Desktop & CurrentFile)
        Else
            Exit Sub
        End If


        Do While xxx < 2
            sKill = "TASKKILL /F /IM EXCEL.EXE"
            Shell(sKill, vbHide)
            xxx = xxx + 1
        Loop

        Threading.Thread.Sleep(500)

        On Error Resume Next

        Dim objExcel = CreateObject("Excel.Application")
        objExcel.Application.Run("'" & Desktop & CurrentFile & "'!Module1.Automate", RequestItem)
        objExcel.DisplayAlerts = False
        objExcel.Application.Quit
        objExcel = Nothing

        Threading.Thread.Sleep(500)

        xxx = 0
        Do While xxx < 2
            sKill = "TASKKILL /F /IM EXCEL.EXE"
            Shell(sKill, vbHide)
            xxx = xxx + 1
        Loop

        Threading.Thread.Sleep(500)

        If File.Exists(Desktop & CurrentFile) Then Kill(Desktop & CurrentFile)

    End Sub

    Sub SVTS_Check(SVBI_Execute_String)
        If SVBI_Connection.State = 1 Then
            SVBI_Connection.Close()
        Else
        End If

        SVBI_Connection.Open(SVBI_Risk)
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        Do While Not rs.EOF
            RequestItemCheck = rs.Fields(0).Value.ToString()
            rs.MoveNext()
        Loop
    End Sub

    Sub SVDG_Check(SVDG_Execute_String)
        If SVDG_Connection.State = 1 Then
            SVDG_Connection.Close()
        Else
        End If

        SVDG_Connection.Open(SVDG_CaseTracker)
        rs = SVDG_Connection.Execute(SVDG_Execute_String)
        Do While Not rs.EOF
            RequestItemCheck = rs.Fields(0).Value.ToString()
            rs.MoveNext()
        Loop
    End Sub

    'Sub Create2018Profile(RequestItem As String, EmailName As String)
    '
    '    CreateCaseHistories(RequestItem, "NoEmail")
    '    CreateCERS(RequestItem)
    '    ControlLocation = "\\education.vic.gov.au\SHARE\TMO\Projects\2017_PSP\ControlItems\"
    '
    '    currentfile = "2018_PSP_Control.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_Regionality.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_Regionality_PDPs.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_StateData.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_SurveyDetails.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_One.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_Two.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_Three.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_Four.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_Five.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_Six.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    currentfile = "2018_PSP_Combiner.xlsm"
    '    Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
    '    If EmailName = "Yes" Then
    '        attachments = reply.Attachments
    '        If File.Exists("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\2017_PSP\Reports\All\" & Item_Found & "_Profile_2018.pdf") Then
    '            attachments.Add("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\2017_PSP\Reports\All\" & Item_Found & "_Profile_2018.pdf")
    '        End If
    '        reply.HTMLBody = reply.HTMLBody + "<br>Thank you for requesting a 2018 Provider Selection Process profile. <br><br>" &
    '                             "Please find attached the 2018 selection profile, as it currently sits, for " & Item_Found & ".<br> <br>Your account has been billed: 
    '									 $" & price & ". Your current balance is now $" & pricesum & ".<br><br><hr><P STYLE='font-family:Calbri;
    '									 font-size:12'>If there was something else you were after, or if you have any suggestions - chat to Lance 
    '									 Snell <br><br>"
    '        reply.Send()
    '        Call WaitOnSend()
    '    End If
    '
    '    If DisplayModel = vbYes Then
    '        currentfile = "SVTS_Reports.xlsm"
    '        Call ReportingGo(RequestItem, currentfile, ControlLocation)
    '
    '    End If
    'End Sub

    Sub ReportingGo(RequestItem, CurrentFile, ControlLocation)

        Dim sKill As String
        Dim xxx As Integer = 0
        Dim FromPath As String
        Dim FSO As Object = CreateObject("scripting.filesystemobject")

        If File.Exists(Desktop & CurrentFile) Then Kill(Desktop & CurrentFile)
        FSO.CopyFile(Source:=ControlLocation & CurrentFile, Destination:=Desktop & CurrentFile)

        Do While xxx < 2
            sKill = "TASKKILL /F /IM EXCEL.EXE"
            Shell(sKill, vbHide)
            xxx = xxx + 1
        Loop

        Threading.Thread.Sleep(500)

        On Error Resume Next

        Dim objExcel = CreateObject("Excel.Application")
        objExcel.Visable = True
        objExcel.WindowState = vbMaximizedFocus
        objExcel.Application.Workbooks.Open(Desktop & CurrentFile)
        objExcel.DisplayAlerts = True

        objExcel = Nothing

        Threading.Thread.Sleep(500)

    End Sub

    Sub CreateUplift(RequestItem As String)

        ControlLocation = "\\education.vic.gov.au\SHARE\TMO\Projects\RTOUplift\"

        currentfile = "TMQUplift_Starter.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
        currentfile = "TMQUplift_Facts.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
        currentfile = "TMQUplift_Controller.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        attachments = reply.Attachments
        If File.Exists("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOUplift\Reports\" & Item_Found & "_Uplift.pdf") Then
            attachments.Add("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOUplift\Reports\" & Item_Found & "_Uplift.pdf")
        End If
        reply.HTMLBody = reply.HTMLBody + "<br>Please find attached a profile covering details of the Skills Uplift (for " & Item_Found & ") level activity for 2018 created by Lana Dalidowicz.<br>
										Your account has been billed: $" & price & ". Your current balance is now $" & pricesum & ".<br> <br> " +
                        "<hr><P STYLE='font-family:Calbri;font-size:12'>If there was something else you were after, or if you have any suggestions - chat to Lance Snell <br><br>"
        reply.Send()
        Call WaitOnSend()

    End Sub

    Sub CreateProgram(RequestItem As String)

        ControlLocation = "\\education.vic.gov.au\SHARE\TMO\Projects\ProgramInformer\ControlItems\"

        currentfile = "Program_Starter.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
        currentfile = "ProgramFacts.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
        currentfile = "Program_Controller.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        attachments = reply.Attachments

        If File.Exists("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\ProgramInformer\Reports\" & Item_Found & "_Profile.pdf") Then
            attachments.Add("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\ProgramInformer\Reports\" & Item_Found & "_Profile.pdf")
        End If

        reply.HTMLBody = reply.HTMLBody + "<br>Please find attached a profile covering Program level activity for 2017/18 created by Lana Dalidowicz.<br>
										Your account has been billed: $" & price & ". Your current balance is now $" & pricesum & ".<br> <br> " +
                        "<hr><P STYLE='font-family:Calbri;font-size:12'>If there was something else you were after, or if you have any suggestions - chat to Lance Snell <br><br>"
        reply.Send()
        Call WaitOnSend()

    End Sub

    Sub CreateRisk(RequestItem As String)

        SVBI_Execute_String = "SELECT TOID FROM SVTS_Risk.ppm.ControlTOIDS WHERE TOID = " + RequestItem

        Call SVTS_Check(SVBI_Execute_String)
        If RequestItemCheck <> RequestItem Then
            ErrorMessage =
            "It appears that no information for this TOID/Request; OR the SVBI server is unresponsive." &
            "<Br><br>Please check your request and try again."
            Call ErrorInEmails(RequestItem, ErrorMessage)
            Exit Sub
        End If

        Call CreateCaseHistories(RequestItem, "")

        ControlLocation = "\\education.vic.gov.au\SHARE\TMO\Projects\RTOInformer\ControlItems_TMS\"

        currentfile = "TMSRisk_Controller.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        currentfile = "TMSRisk_Facts.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        currentfile = "TMSRisk_Leagues.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        Directory.GetFiles(ControlLocation)

        Dim txtFiles = Directory.GetFiles(ControlLocation, "*.xlsm", SearchOption.TopDirectoryOnly).
        [Select](Function(nm) Path.GetFileName(nm))
        Dim BulkFiles As String
        BulkFiles = "TMSRisk_RisksIR"
        For Each filenm As String In txtFiles
            If filenm.ToUpper.Contains(BulkFiles.ToUpper) Then
                currentfile = filenm
                Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
            End If
        Next

        currentfile = "TMSRisk_StateData.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        currentfile = "TMSRisk_Combiner.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        On Error Resume Next
        If EmailName = "Yes" Then
            attachments = reply.Attachments
            If File.Exists("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\All\" & Item_Found & "_Profile_2018.pdf") Then
                attachments.Add("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer\Reports\All\" & Item_Found & "_Profile_2018.pdf")
            End If
            reply.HTMLBody = reply.HTMLBody + "<br>Thank you for requesting a the 2018 TMS Risk Profile. <br><br>" &
                                 "Please find attached the profile, as it currently sits, for " & Item_Found & ".<br> <br>Your account has been billed: 
										 $" & price & ". Your current balance is now $" & pricesum & ".<br><br><hr><P STYLE='font-family:Calbri;
										 font-size:12'>If there was something else you were after, or if you have any suggestions - chat to Lance 
										 Snell <br><br>"
            reply.Send()
            Call WaitOnSend()
        End If

    End Sub

    Sub CreateVRQA(RequestItem As String)

        SVBI_Execute_String = "SELECT TOID FROM SVTS_Risk.VRQA.ControlTOIDS WHERE TOID = " + RequestItem

        Call SVTS_Check(SVBI_Execute_String)
        If RequestItemCheck <> RequestItem Then
            ErrorMessage =
            "It appears that no information for this TOID/Request; OR the TOID is not related to the VRQA; OR the SVBI server is unresponsive." &
            "<Br><br>Please check your request and try again."
            Call ErrorInEmails(RequestItem, ErrorMessage)
            Exit Sub
        End If

        ControlLocation = "\\education.vic.gov.au\SHARE\TMO\Projects\RTOInformer_VRQA\ControlItems_TMS\"

        currentfile = "TMSRisk_Controller.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        currentfile = "TMSRisk_Facts.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        currentfile = "TMSRisk_Leagues.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        Directory.GetFiles(ControlLocation)

        Dim txtFiles = Directory.GetFiles(ControlLocation, "*.xlsm", SearchOption.TopDirectoryOnly).
        [Select](Function(nm) Path.GetFileName(nm))
        Dim BulkFiles As String
        BulkFiles = "TMSRisk_RisksIR"
        For Each filenm As String In txtFiles
            If filenm.ToUpper.Contains(BulkFiles.ToUpper) Then
                currentfile = filenm
                Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
            End If
        Next

        currentfile = "TMSRisk_Combiner.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        On Error Resume Next
        If EmailName = "Yes" Then
            attachments = reply.Attachments
            If File.Exists("\\education.vic.gov.au\SHARE\TMO\Projects\RTOInformer_VRQA\Reports\All\" & Item_Found & "_VRQA_2018.pdf") Then
                attachments.Add("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOInformer_VRQA\Reports\All\" & Item_Found & "_VRQA_2018.pdf")
            End If
            reply.HTMLBody = reply.HTMLBody + "<br>Thank you for requesting a copy of the 2018 VRQA Risk Profile. <br><br>" &
                                 "Please find attached the profile, as it currently sits, for " & Item_Found & ".<br> <br>Your account has been billed: 
										 $" & price & ". Your current balance is now $" & pricesum & ".<br><br><hr><P STYLE='font-family:Calbri;
										 font-size:12'>If there was something else you were after, or if you have any suggestions - chat to Lance 
										 Snell <br><br>"
            reply.Send()
            Call WaitOnSend()
        End If

    End Sub

    Sub CreatePDP2018(RequestItem As String)
        ControlLocation = "C:\Users\09221031\Desktop\TMSProfile\" 'TESTING - local CD

        currentfile = "TMSProfile_Controller.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        Directory.GetFiles(ControlLocation)

        Dim txtFiles = Directory.GetFiles(ControlLocation, "*.xlsm", SearchOption.TopDirectoryOnly).
        [Select](Function(nm) Path.GetFileName(nm))
        Dim BulkFiles As String
        BulkFiles = "TMSProfile"
        For Each filenm As String In txtFiles
            If filenm.ToUpper.Contains(BulkFiles.ToUpper) And Not filenm.ToUpper.Contains("CONTROLLER") And Not filenm.ToUpper.Contains("COMBINER") Then
                currentfile = filenm
                Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
            End If
        Next

        currentfile = "TMSProfile_Combiner.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        attachments = reply.Attachments
        If File.Exists("C:\Users\09221031\Desktop\Test\All\" & Item_Found & "_Profile_2018.pdf") Then
            attachments.Add("C:\Users\09221031\Desktop\Test\All\" & Item_Found & "_Profile_2018.pdf")
        End If
        reply.HTMLBody = reply.HTMLBody + "<br>2018 Profile assessement tool kit is attached for " & Item_Found & "!<br> <br>Your account has been billed: 
										 $" & price & ". Your current balance is now $" & pricesum & ".<br><br><hr><P STYLE='font-family:Calbri;
										 font-size:12'>If there was something else you were after, or if you have any suggestions - chat to Lance 
										 Snell <br><br>"
        reply.Send()
        Call WaitOnSend()
    End Sub
    Sub CreateRTOProgram(RequestItem As String)

        ControlLocation = "\\education.vic.gov.au\SHARE\TMO\Projects\RTOProgramInformer\ControlItems\"

        currentfile = "RTOProgramFacts_Starter.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
        currentfile = "RTOProgramFacts_Regions.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
        currentfile = "RTOProgramFacts_One.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
        currentfile = "RTOProgramFacts_Two.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)
        currentfile = "RTOProgramFacts_Controller.xlsm"
        Call ExcelFileLooper(RequestItem, currentfile, ControlLocation)

        attachments = reply.Attachments
        If File.Exists("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOProgramInformer\Reports\" & Item_Found & "_Profile.pdf") Then
            attachments.Add("\\Education.Vic.Gov.Au\SHARE\TMO\Projects\RTOProgramInformer\Reports\" & Item_Found & "_Profile.pdf")
        End If
        reply.HTMLBody = reply.HTMLBody + "<br>Thanks for requesting an RTO/Program hybrid delivery profile for " & Item_Found & ".<br> <br>Your account has been billed: 
										 $" & price & ". Your current balance is now $" & pricesum & ".<br><br><hr><P STYLE='font-family:Calbri;
										 font-size:12'>If there was something else you were after, or if you have any suggestions - chat to Lance 
										 Snell <br><br>"
        reply.Send()
        Call WaitOnSend()

    End Sub

    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Integer, ByVal uParam As Integer, ByVal lpvParam As String, ByVal fuWinIni As Integer) As Integer

    Private Const SETDESKWALLPAPER = 20
    Private Const UPDATEINIFILE = &H1
    Public Sub SetWallpaper(path)

        Dim sKill As String
        Dim xxx As Integer = 0
        Dim FromPath As String
        Dim FSO As Object = CreateObject("scripting.filesystemobject")
        Dim Currentfile As String = "BackGroundMaker.xlsm"
        ControlLocation = "\\education.vic.gov.au\SHARE\TMO\Projects\2017_PSP\ControlItems\"

        If File.Exists(Desktop & Currentfile) Then Kill(Desktop & Currentfile)
        FSO.CopyFile(Source:=ControlLocation & Currentfile, Destination:=Desktop & Currentfile)


        Do While xxx < 2
            sKill = "TASKKILL /F /IM EXCEL.EXE"
            Shell(sKill, vbHide)
            xxx = xxx + 1
        Loop

        Threading.Thread.Sleep(500)

        On Error Resume Next

        Dim objExcel = CreateObject("Excel.Application")
        objExcel.Application.Run("'" & Desktop & Currentfile & "'!Module1.Automate")
        objExcel.DisplayAlerts = False
        objExcel.Application.Quit
        objExcel = Nothing

        Threading.Thread.Sleep(500)

        xxx = 0

        Do While xxx < 2
            sKill = "TASKKILL /F /IM EXCEL.EXE"
            Shell(sKill, vbHide)
            xxx = xxx + 1
        Loop

        Threading.Thread.Sleep(500)

        If File.Exists(Desktop & Currentfile) Then Kill(Desktop & Currentfile)

        SystemParametersInfo(SETDESKWALLPAPER, 0, path, UPDATEINIFILE)

    End Sub
    Sub Dangerzonesub(ChosenLoop As String)
        'Note: Allows lots of profiles to be created in a row...
        'Can loop to make efficient if needed
        If dangerzone = vbNo Then Exit Sub
        Dim rs As ADODB.Recordset
        Dim SVBI_Connection As New ADODB.Connection
        Dim RequestItem As String
        ChosenLoop = 0

        Dim start As String = Now()

        SVBI_Execute_String = "SELECT DISTINCT TOID FROM SVTS_Risk.PPM.ControlTOIDS"
        If SVBI_Connection.State = 1 Then
            SVBI_Connection.Close()
        Else
        End If

        SVBI_Connection.Open(SVBI_Risk)
        rs = SVBI_Connection.Execute(SVBI_Execute_String)
        Dim Email As String = "No"
        'Do this while there is still data (rows) being retrieved from the SQL script.
        Do While Not rs.EOF
            RequestItem = rs.Fields(0).Value.ToString()
            Call CreateRisk(RequestItem)
            rs.MoveNext()
        Loop


    End Sub

    Sub ErrorInEmails(RequestItem, ErrorMessage)
        reply.HTMLBody =
        "<Br><Br>Oho - something went wrong with your request... :(" &
        "<br><br>You had requested something for: " & RequestItem & "." &
        "<br><br>" & ErrorMessage
        reply.Send()
        Call WaitOnSend()
    End Sub

    Public Sub Rewrite()
        'Note: Not currently used
        Dim Ok As String
        Ok = Desktop & "ASIC.csv"
        If File.Exists(Ok) Then
            Dim CSVlinesIn As New ArrayList
            Dim CSVout As New List(Of String)
            CSVlinesIn.AddRange(IO.File.ReadAllLines(Ok))
            Dim XY As Int16 = 1
            For Each line As String In CSVlinesIn
                XY = XY + 1
                If XY > 3 Then
                    Dim nameANDnumber As String() = line.Split(","c)
                    If nameANDnumber(0).Trim <> "" And nameANDnumber(1).Trim <> "" And nameANDnumber(2).Trim <> "" And nameANDnumber(3).Trim <> "" And nameANDnumber(6).Trim <> "" Then CSVout.Add(line)
                End If
            Next
            IO.File.WriteAllLines(Ok, CSVout.ToArray)
        Else MsgBox("No file at " & Ok)
        End If
    End Sub

    Public Sub CSVDataReader(mail As MailItem)
        'Note: Not currently used
        Call Rewrite()
        Dim Ok As String
        Ok = Desktop & "ASIC.csv"
        Dim fi As New FileInfo(Ok)
        Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Text;Data Source=" & fi.DirectoryName

        Dim conn As New OleDbConnection(connectionString)
        conn.Open()
        Dim cmdSelect As New OleDbCommand("SELECT Date, ABN, Action, [Sub Type] FROM " & fi.Name, conn)

        Dim adapter1 As New OleDbDataAdapter
        adapter1.SelectCommand = cmdSelect

        Dim ds As New DataSet
        adapter1.Fill(ds, "DATA")
        Dim T As String = 0

        Dim thisTable As DataTable
        For Each thisTable In ds.Tables
            On Error GoTo BadEnd
            reply = Nothing
            reply = Application.CreateItem(Outlook.OlItemType.olMailItem)
            reply.DeleteAfterSubmit = True
            reply.To = senderemail
            reply.CC = UNameWindows() & "; crest.pamela.p@edumail.vic.gov.au"
            reply.Subject = "[Unclassified: For Official Use Only] Automated ASIC Update"
            reply.HTMLBody = "<P STYLE='font-family:Calbri;font-size:11'>"

            For Each row As DataRow In thisTable.Rows
                If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
                SVBI_Connection.Open(SVBI_Risk)
                'MsgBox(row.Field(Of String)(0).ToString())
                'MsgBox(row.Field(Of String)(1).ToString())
                'MsgBox(Regex.Replace(row.Field(Of String)(2).ToString(), "'", ""))
                'MsgBox(Regex.Replace(row.Field(Of String)(3).ToString(), "'", ""))
                SVBI_Execute_String = "
BEGIN TRY DROP TABLE #ASIC END TRY BEGIN CATCH END CATCH;
CREATE TABLE #ASIC (ActionDate DATETIME, ABN VARCHAR(MAX), Notification_Type VARCHAR(MAX), Actioned VARCHAR(MAX))

INSERT INTO #ASIC (ActionDate,ABN,Notification_Type,Actioned)
SELECT  convert(datetime, '" & row.Field(Of String)(0).ToString() & "', 103)
	  , '" & row.Field(Of String)(1).ToString() & "'
	  , '" & Regex.Replace(row.Field(Of String)(2).ToString(), "'", "") & "' + ' ' + '" & Regex.Replace(row.Field(Of String)(3).ToString(), "'", "") & "'
	  , 'No'

IF OBJECT_ID('TempDB..#ContractDetails','U') IS NOT NULL DROP TABLE #ContractDetails
SELECT R.TOID, COUNT(DISTINCT ces.CourseEnrolmentSupersededID) AS ContEnrolments, MAX(cy.ContractYear) MaxYear, MAX(ib.ProcessingDate) AS LastPayment 
INTO #ContractDetails
FROM #ASIC A
LEFT JOIN svts.dbo.rto r ON r.ABN = A.ABN
LEFT JOIN svts.dbo.CourseEnrolmentSuperseded ces ON ces.TOID = r.TOID AND ces.LastActivityDate > GETDATE() AND ces.PaidAmount IS NOT NULL
LEFT JOIN svts.dbo.contract c ON c.TOID = r.TOID
LEFT JOIN svts.dbo.ContractYear cy ON cy.ContractYearID = c.ContractYearID
LEFT JOIN svts.dbo.Invoice i ON i.ContractID = c.ContractID
LEFT JOIN SVTS.dbo.InvoiceBatch ib ON ib.InvoiceBatchID = i.InvoiceBatchID
GROUP BY r.TOID


INSERT INTO SVTS_Risk.printer.ASIC (ActionDate,ABN,Notification_Type,Actioned,Students,LastContract, LastPayment)
SELECT  I.ActionDate
	  , I.ABN
	  , I.Notification_Type
	  , I.Actioned
	  , CD.ContEnrolments
	  , ISNULL(CD.MaxYear,0)
	  , CD.LastPayment
FROM #ASIC I
LEFT JOIN SVTS_Risk.printer.ASIC A ON A.ABN = I.ABN AND A.ActionDate = I.ActionDate AND A.Notification_Type = I.Notification_Type
LEFT JOIN svts.dbo.rto r ON r.ABN = I.ABN
LEFT JOIN #ContractDetails CD ON CD.TOID = r.TOID
WHERE 1 = 1
AND A.ActionDate IS NULL"

                rs = SVBI_Connection.Execute(SVBI_Execute_String)
                rs = Nothing

            Next

            If SVBI_Connection.State = 1 Then SVBI_Connection.Close()
            SVBI_Connection.Open(SVBI_Risk)
            SVBI_Execute_String = "IF OBJECT_ID('TempDB..#HolderPlace', 'U') IS NOT NULL DROP TABLE #HolderPlace; 
SELECT MAX(ActionDate) AS ActionDate ,
	   ABN ,
	   '' Notification_Type ,
	   'No' Actioned ,
	   MAX(Students) AS Students,
	   MAX(LastContract) AS LastContract,
	   MAX(LastPayment) AS LastPayment
INTO #HolderPlace
FROM SVTS_Risk.Printer.ASIC A
WHERE A.Actioned = 'No'
GROUP BY ABN

IF OBJECT_ID('TempDB..#Holder', 'U') IS NOT NULL DROP TABLE #Holder; 
SELECT A.ActionDate ,
	   A.ABN ,
	   A.Notification_Type ,
	   A.Actioned,
	   ROW_NUMBER() OVER (PARTITION BY A.ABN ORDER BY A.ABN) AS TempOrder
INTO #Holder
FROM  SVTS_Risk.Printer.ASIC A
WHERE a.Actioned = 'No'

DELETE FROM SVTS_Risk.printer.ASIC WHERE Actioned = 'No'

ALTER TABLE #HolderPlace ALTER COLUMN Notification_Type VARCHAR(MAX)

DECLARE @MaxOrder AS INT = (SELECT MAX(Temporder) FROM #Holder)
WHILE @MaxOrder > 0
BEGIN
UPDATE HP
SET hp.Notification_Type = hp.Notification_Type + '
' + h.Notification_Type FROM #HolderPlace HP
JOIN #Holder H ON H.ABN = HP.ABN
WHERE H.TempOrder = @MaxOrder
SET @MaxOrder -= 1
END

INSERT INTO SVTS_Risk.printer.ASIC (ActionDate,ABN,Notification_Type,Actioned,Students,LastContract, LastPayment)
SELECT  I.ActionDate
	  , I.ABN
	  , I.Notification_Type
	  , I.Actioned
	  , ISNULL(I.Students,0)
	  , ISNULL(I.LastContract,0)
	  , ISNULL(I.LastPayment,0)
FROM #HolderPlace I
LEFT JOIN SVTS_Risk.printer.ASIC A ON A.ABN = I.ABN AND A.ActionDate = I.ActionDate AND A.Notification_Type = I.Notification_Type
WHERE 1 = 1
AND A.ActionDate IS NULL"
            rs = SVBI_Connection.Execute(SVBI_Execute_String)
            SVBI_Execute_String = "SELECT COUNT(*) AS Loops FROM SVTS_Risk.printer.ASIC WHERE Actioned = 'No'"
            rs = SVBI_Connection.Execute(SVBI_Execute_String)
            Dim Looped As String = rs.Fields("Loops").Value.ToString()

            Do While Looped > 0
                SVBI_Execute_String = "SELECT TOP 1 * from SVTS_Risk.printer.ASIC A left join svts.dbo.rto r on r.abn = a.abn where Actioned = 'No'"
                rs = SVBI_Connection.Execute(SVBI_Execute_String)
                If rs.EOF = False Then
                    Dim TOID As String
                    TOID = rs.Fields("TOID").Value.ToString()
                    Dim Notification_Type As String
                    Notification_Type = rs.Fields("Notification_Type").Value.ToString()
                    Dim TradingName As String
                    TradingName = rs.Fields("TradingName").Value.ToString()
                    Dim Contract As String
                    Contract = rs.Fields("LastContract").Value.ToString()
                    Dim Students As String
                    Students = rs.Fields("Students").Value.ToString()
                    Dim LastPayment As String
                    LastPayment = rs.Fields("LastPayment").Value.ToString()
                    Dim Dated As String
                    Dated = rs.Fields("ActionDate").Value.ToString()
                    If SVDG_Connection.State = 1 Then SVDG_Connection.Close()
                    SVDG_Connection.Open(SVDG_CaseTracker)
                    SVDG_Execute_String = "IF OBJECT_ID('TempDB..#CheckOne', 'U') IS NOT NULL DROP TABLE #CheckOne; 
SELECT C.CaseID
INTO #CheckOne
FROM VPMS_DEV.dbo.RTOCase_Case C
WHERE c.TOID = " & TOID & " AND C.CaseBrief = 'ASIC Notice' AND C.CaseDate =  '" & Dated & "' AND CaseEntryOfficer = 09221031

IF (SELECT CASEID FROM #CheckOne) IS NULL
BEGIN
INSERT INTO VPMS_DEV.dbo.RTOCase_Case (TOID, CaseBrief, CaseDate, CaseAudit, CaseRecordAddedDate, CaseSummary, CaseTriggerId, CaseEntryOfficer, CaseOutcome)
SELECT " & TOID & ", 'ASIC Notice', '" & Dated & "', 'Bulk upload, ASIC notification: Test insertion by addon - 09221031', GETDATE(), '" & Notification_Type & "', 15, 09221031, 'NFA'
END

IF OBJECT_ID('TempDB..#CheckTwo', 'U') IS NOT NULL DROP TABLE #CheckTwo; 
SELECT C.CaseID
INTO #CheckTwo
FROM VPMS_DEV.dbo.RTOCase_Case C
JOIN VPMS_DEV.dbo.RTOCase_CaseStatus cs ON cs.CaseID = C.CaseID
WHERE c.TOID = " & TOID & " AND C.CaseBrief = 'ASIC Notice' AND C.CaseDate =  '" & Dated & "' AND CaseEntryOfficer = 09221031 

IF (SELECT CASEID FROM #CheckTwo) IS NULL
BEGIN
INSERT INTO VPMS_DEV.dbo.RTOCase_CaseStatus (CaseId, CaseStatusId, CaseManagerOfficerId, TeamId, Comments)
SELECT C.CaseID, 4, 09221031, 1, 'AutoStatus'
FROM VPMS_DEV.dbo.RTOCase_Case C
WHERE c.TOID = " & TOID & " AND C.CaseBrief = 'ASIC Notice' AND C.CaseDate =  '" & Dated & "' AND CaseEntryOfficer = 09221031 
END
"
                    rs = SVDG_Connection.Execute(SVDG_Execute_String)
                    rs = Nothing
                    SVBI_Execute_String = "UPDATE SVTS_Risk.printer.ASIC SET Actioned = 'Ye' where Actioned = 'No' and actiondate = '" & Dated & "' and Notification_Type = '" & Notification_Type & "'"
                    rs = SVBI_Connection.Execute(SVBI_Execute_String)
                    If Contract > 0 Then
                        If Students = 0 Then
                            rs = SVDG_Connection.Execute(SVDG_Execute_String)
                            reply.HTMLBody = reply.HTMLBody & TOID & " - " & TradingName
                            reply.HTMLBody = reply.HTMLBody & "<br><B>Last contract: </b>" & Contract
                            reply.HTMLBody = reply.HTMLBody & "<Br>Active students: </b>" & Students
                            reply.HTMLBody = reply.HTMLBody & "<Br><b>Last payment: </b></font>" & LastPayment
                            reply.HTMLBody = reply.HTMLBody & "<Br><B>Specifics: </B><br>" & Notification_Type & "<br><br>"
                        Else
                            rs = SVDG_Connection.Execute(SVDG_Execute_String)
                            reply.HTMLBody = "<B>Specifics: </B>" & "" & Notification_Type & "<br><br>" & reply.HTMLBody
                            reply.HTMLBody = "<font color = 'darkred'><b>Last payment: </b>" & LastPayment & "</font><br>" & reply.HTMLBody
                            reply.HTMLBody = "<font color = 'darkred'><b>Active students: </b>" & Students & "</font><br>" & reply.HTMLBody
                            reply.HTMLBody = "<B>Last contract: </b>" & Contract & "<br>" & reply.HTMLBody
                            reply.HTMLBody = TOID & " - " & TradingName & "<br>" & reply.HTMLBody

                        End If
                        T = 1
                    End If
                End If

                Looped = Looped - 1
            Loop

        Next
        reply.HTMLBody = "<P STYLE='font-family:Calbri;font-size:11'>Automated email about: " & mail.Subject.ToString() & "<BR><br>" & reply.HTMLBody
        reply.HTMLBody = "<P STYLE='font-family:Calbri;font-size:11'><br><b>ASIC update received</b><Br><br>All below data inputed into Case Tracker (DEV) - located <a href=\\education.vic.gov.au\SHARE\TMO\Vet\Division VET\RTO Case Tracker\Testing\>here</a><br>" & reply.HTMLBody
        reply.HTMLBody = "<BR><font color = 'darkred'><B>BETA</B></font>" & reply.HTMLBody
        If T > 0 Then
            reply.Send()
            Call WaitOnSend()
        End If
BadEnd:
        reply = Nothing
    End Sub

    Public Sub TMS_Off(Mail As MailItem)
        Deactivate_Trigger = "Yes"
        reply = Nothing
        reply = Application.CreateItem(Outlook.OlItemType.olMailItem)
        reply.DeleteAfterSubmit = True
        reply.To = Mail.SenderEmailAddress
        reply.CC = UNameWindows()
        reply.Subject = "[Unclassified: For Official Use Only] TMS Profiler - Off"
        reply.HTMLBody = OnTime & "<br><br>" & "<b>Current trigger word: </b>" & codeword
        reply.HTMLBody = reply.HTMLBody & "<br><br>" & Version
        reply.Send()
        Call WaitOnSend()
        Mail.UnRead = False
        Mail.Move(destfolder)
    End Sub

    Public Sub TMS_On(Mail As MailItem)
        Deactivate_Trigger = "No"
        reply = Nothing
        reply = Application.CreateItem(Outlook.OlItemType.olMailItem)
        reply.DeleteAfterSubmit = True
        reply.To = Mail.SenderEmailAddress
        reply.CC = UNameWindows()
        reply.Subject = "[Unclassified: For Official Use Only] TMS Profiler - On"
        reply.HTMLBody = OnTime & "<br><br>" & "<b>Current trigger word: </b>" & codeword
        reply.HTMLBody = reply.HTMLBody & "<br><br>" & Version
        reply.HTMLBody = reply.HTMLBody & "<br><br>" & StartTimer
        reply.Send()
        Call WaitOnSend()
        Mail.UnRead = False
        Mail.Move(destfolder)
    End Sub

    Public Sub ASQA_CTInsert(mail As MailItem)
        Dim PDFAddress As String = "\\education.vic.gov.au\share\TMO\All_Units\Case_Tracker\Work_Email\ASQA_"
        Dim iAttachCnt As Integer
        Dim sFileName As String
        Dim msg As String
        Dim varname As String

        varname = New Random().Next(0, 1000000)

        mail.SaveAs(PDFAddress & varname & "_" & Now().ToString("yyyyMMdd") & ".txt", 0)
        msg = mail.HTMLBody.ToString()

        With mail.Attachments
            iAttachCnt = .Count
            If iAttachCnt > 0 Then
                For iCtr = 1 To iAttachCnt
                    sFileName = .Item(iCtr).FileName.ToString
                    If sFileName.ToUpper.Contains("PDF") Then
                        .Item(iCtr).SaveAsFile(PDFAddress & varname & "_" & Now().ToString("yyyyMMdd") & ".pdf")
                        Threading.Thread.Sleep(500)
                    End If
                    Threading.Thread.Sleep(500)
                Next iCtr
            Else
            End If
        End With

        mail.UnRead = False
        mail.Move(destfolder)
        Threading.Thread.Sleep(500)
        Call CreateASQANotification()
    End Sub

    Public Sub ASIC_CTInsert(mail As MailItem)
        'Note: Not currently used!
        Dim PDFAddress As String = "\\education.vic.gov.au\share\TMO\All_Units\Case_Tracker\Work_Email\ASQA_"
        Dim iAttachCnt As Integer
        Dim sFileName As String
        Dim varname As String

        varname = "ASIC.CSV"

        With mail.Attachments
            iAttachCnt = .Count
            If iAttachCnt > 0 Then
                For iCtr = 1 To iAttachCnt
                    sFileName = .Item(iCtr).FileName.ToString
                    If sFileName.ToUpper.Contains("CSV") Then
                        .Item(iCtr).SaveAsFile(Desktop & varname)
                        Threading.Thread.Sleep(500)
                    End If
                    Threading.Thread.Sleep(500)
                Next iCtr
            Else
            End If
        End With

        mail.UnRead = False
        mail.Move(destfolder)
        Threading.Thread.Sleep(500)
        Call CSVDataReader(mail)
        If File.Exists(Desktop & varname) Then File.Delete(Desktop & varname)
    End Sub
    Public Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class