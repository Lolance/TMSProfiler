Public Class ControlForm
    Private Sub Stopper_Click(sender As Object, e As EventArgs) Handles Stopper.Click
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ControlLocation = RTOInformer_Location.Text.ToString
        ExcelFileRefresher(ControlLocation)
    End Sub

    Sub ExcelFileRefresher(ControlLocation As String)

        Dim sKill As String
        Dim xxx As Integer = 0
        Dim FSO As Object = CreateObject("scripting.filesystemobject")

        Do While xxx < 2
            sKill = "TASKKILL /F /IM EXCEL.EXE"
            Shell(sKill, vbHide)
            xxx = xxx + 1
        Loop

        Dim XLApp As Object
        Dim wr As Object
        Threading.Thread.Sleep(100)
        Dim fname As String = Dir(ControlLocation & "*.xlsm")
        Do While Len(fname) > 0
            XLApp = CreateObject("Excel.Application")
            XLApp.visible = False   ' not required, you do not need to see this happening
            wr = XLApp.workbooks.Open(ControlLocation & fname)
            wr.refreshall()
            'wr.doevents()
            Threading.Thread.Sleep(1000)
            wr.Save()
            Threading.Thread.Sleep(1000)
            fname = Dir()
            xxx = 0
            Do While xxx < 2
                sKill = "TASKKILL /F /IM EXCEL.EXE"
                Shell(sKill, vbHide)
                xxx = xxx + 1
            Loop
            wr = Nothing
            XLApp = Nothing
            Threading.Thread.Sleep(1000)
        Loop
        MsgBox("Update completed")
        Threading.Thread.Sleep(500)

    End Sub

    Private Sub ControlForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'ActiveForm.BringToFront()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles Profile_Only.CheckedChanged
        If Profile_Only.Checked = True Then
            ProfileOnly = "1"
        Else
            ProfileOnly = "2"
        End If
    End Sub

End Class