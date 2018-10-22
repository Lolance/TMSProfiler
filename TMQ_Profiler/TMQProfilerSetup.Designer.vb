<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TMQProfilerSetup
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DesktopChanger_Button = New System.Windows.Forms.CheckBox()
        Me.SVTSReports_Button = New System.Windows.Forms.CheckBox()
        Me.Launch = New System.Windows.Forms.Button()
        Me.Stopper = New System.Windows.Forms.Button()
        Me.CodeMaker_Input = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.LooperCheck = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Test = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'DesktopChanger_Button
        '
        Me.DesktopChanger_Button.AutoSize = True
        Me.DesktopChanger_Button.Location = New System.Drawing.Point(13, 209)
        Me.DesktopChanger_Button.Margin = New System.Windows.Forms.Padding(4)
        Me.DesktopChanger_Button.Name = "DesktopChanger_Button"
        Me.DesktopChanger_Button.Size = New System.Drawing.Size(221, 20)
        Me.DesktopChanger_Button.TabIndex = 1
        Me.DesktopChanger_Button.Text = "Generate a new desktop image?"
        Me.DesktopChanger_Button.UseVisualStyleBackColor = True
        '
        'SVTSReports_Button
        '
        Me.SVTSReports_Button.AutoSize = True
        Me.SVTSReports_Button.Location = New System.Drawing.Point(13, 237)
        Me.SVTSReports_Button.Margin = New System.Windows.Forms.Padding(4)
        Me.SVTSReports_Button.Name = "SVTSReports_Button"
        Me.SVTSReports_Button.Size = New System.Drawing.Size(290, 20)
        Me.SVTSReports_Button.TabIndex = 2
        Me.SVTSReports_Button.Text = "Display SVTS status after profiles generate?"
        Me.SVTSReports_Button.UseVisualStyleBackColor = True
        '
        'Launch
        '
        Me.Launch.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Maroon
        Me.Launch.Location = New System.Drawing.Point(317, 318)
        Me.Launch.Margin = New System.Windows.Forms.Padding(4)
        Me.Launch.Name = "Launch"
        Me.Launch.Size = New System.Drawing.Size(186, 41)
        Me.Launch.TabIndex = 6
        Me.Launch.Text = "Go"
        Me.Launch.UseVisualStyleBackColor = True
        '
        'Stopper
        '
        Me.Stopper.FlatAppearance.BorderColor = System.Drawing.Color.Maroon
        Me.Stopper.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Maroon
        Me.Stopper.Location = New System.Drawing.Point(13, 318)
        Me.Stopper.Margin = New System.Windows.Forms.Padding(4)
        Me.Stopper.Name = "Stopper"
        Me.Stopper.Size = New System.Drawing.Size(92, 33)
        Me.Stopper.TabIndex = 5
        Me.Stopper.Text = "Do not go"
        Me.Stopper.UseVisualStyleBackColor = True
        '
        'CodeMaker_Input
        '
        Me.CodeMaker_Input.Location = New System.Drawing.Point(152, 286)
        Me.CodeMaker_Input.Margin = New System.Windows.Forms.Padding(4)
        Me.CodeMaker_Input.MaxLength = 15
        Me.CodeMaker_Input.Name = "CodeMaker_Input"
        Me.CodeMaker_Input.Size = New System.Drawing.Size(135, 22)
        Me.CodeMaker_Input.TabIndex = 4
        Me.CodeMaker_Input.Text = "TMSData"
        Me.CodeMaker_Input.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 289)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Command text:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(290, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Reminder: You have loaded in the TMS Profiler."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(394, 16)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "This addon will, amongst other things occasionally take control of:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(45, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 16)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Outlook;"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(45, 69)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 16)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Access; and"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(45, 87)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 16)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Excel."
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 106)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(330, 16)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "It will also copy files to your desktop and send emails..."
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 126)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 16)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Before you run this:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(45, 165)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(458, 16)
        Me.Label10.TabIndex = 15
        Me.Label10.Text = "2. Open Access and Excel to set Trust Center settings to trust everything; and"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(45, 146)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(404, 16)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "1. Install PDFCreator 2.4 (or later) including the COM addin for excel;"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(45, 185)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(398, 16)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "3. Have permissions to SVTS, VMPS and to the TMO Share Drive."
        '
        'LooperCheck
        '
        Me.LooperCheck.AutoSize = True
        Me.LooperCheck.Location = New System.Drawing.Point(13, 265)
        Me.LooperCheck.Margin = New System.Windows.Forms.Padding(4)
        Me.LooperCheck.Name = "LooperCheck"
        Me.LooperCheck.Size = New System.Drawing.Size(131, 20)
        Me.LooperCheck.TabIndex = 3
        Me.LooperCheck.Text = "Run The Looper?"
        Me.LooperCheck.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(380, 292)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 16)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Version (Click)"
        '
        'Button1
        '
        Me.Button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Maroon
        Me.Button1.Location = New System.Drawing.Point(440, 9)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(63, 24)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "Admin"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Test
        '
        Me.Test.Location = New System.Drawing.Point(440, 40)
        Me.Test.Name = "Test"
        Me.Test.Size = New System.Drawing.Size(63, 25)
        Me.Test.TabIndex = 20
        Me.Test.Text = "Test"
        Me.Test.UseVisualStyleBackColor = True
        '
        'TMQProfilerSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(506, 364)
        Me.Controls.Add(Me.Test)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.LooperCheck)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CodeMaker_Input)
        Me.Controls.Add(Me.Stopper)
        Me.Controls.Add(Me.Launch)
        Me.Controls.Add(Me.SVTSReports_Button)
        Me.Controls.Add(Me.DesktopChanger_Button)
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(61, 53)
        Me.Name = "TMQProfilerSetup"
        Me.Text = "TMS Profiler Setup"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DesktopChanger_Button As System.Windows.Forms.CheckBox
    Friend WithEvents SVTSReports_Button As System.Windows.Forms.CheckBox
    Friend WithEvents Launch As System.Windows.Forms.Button
    Friend WithEvents Stopper As System.Windows.Forms.Button
    Friend WithEvents CodeMaker_Input As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents LooperCheck As System.Windows.Forms.CheckBox
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents Test As Windows.Forms.Button
End Class
