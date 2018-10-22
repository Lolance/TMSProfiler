Public Class TMQProfilerSetup

    Public Sub TMQProfilerSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If TestS = vbYes Then MsgBox("Warning: Test mode engaged")
    End Sub

    Public Sub Stopper_Click(sender As Object, e As EventArgs) Handles Stopper.Click
        TestS = vbNo
        FastFix = vbNo
        Close()
    End Sub

    Public Sub Launch_Click(sender As Object, e As EventArgs) Handles Launch.Click
        If DesktopChanger_Button.Checked = True Then
            desktopchanger = vbYes
        Else
            desktopchanger = vbNo
        End If
        If SVTSReports_Button.Checked = True Then
            DisplayModel = vbYes
        Else
            DisplayModel = vbNo
        End If
        If LooperCheck.Checked = True Then
            dangerzone = vbYes
        Else dangerzone = vbno
        End If
        codeword = CodeMaker_Input.Text.ToString.ToUpper
        'MsgBox("Codeword: " & codeword & vbCrLf & vbCrLf & "Desktop changer: " & desktopchanger & vbCrLf & vbCrLf & "Displaymodel: " & DisplayModel)
        Close()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs)

        Close()

    End Sub

    Private Sub CodeMaker_Input_TextChanged(sender As Object, e As EventArgs) Handles CodeMaker_Input.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles LooperCheck.CheckedChanged

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        MsgBox(Version)
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oFormTwo As ControlForm
        oFormTwo = New ControlForm
        oFormTwo.BringToFront()
        oFormTwo.ShowDialog()
        oFormTwo.Activate()
    End Sub

    Private Sub Test_Click(sender As Object, e As EventArgs) Handles Test.Click
        TestS = vbYes
        Close()

    End Sub
End Class