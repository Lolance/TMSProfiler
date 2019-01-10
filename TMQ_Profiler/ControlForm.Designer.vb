<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ControlForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Stopper = New System.Windows.Forms.Button()
        Me.RTOInformer_Location = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Profile_Only = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Stopper
        '
        Me.Stopper.FlatAppearance.BorderColor = System.Drawing.Color.Maroon
        Me.Stopper.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Maroon
        Me.Stopper.Location = New System.Drawing.Point(456, 56)
        Me.Stopper.Margin = New System.Windows.Forms.Padding(4)
        Me.Stopper.Name = "Stopper"
        Me.Stopper.Size = New System.Drawing.Size(92, 33)
        Me.Stopper.TabIndex = 6
        Me.Stopper.Text = "Close"
        Me.Stopper.UseVisualStyleBackColor = True
        '
        'RTOInformer_Location
        '
        Me.RTOInformer_Location.Location = New System.Drawing.Point(13, 4)
        Me.RTOInformer_Location.Margin = New System.Windows.Forms.Padding(4)
        Me.RTOInformer_Location.MaxLength = 255
        Me.RTOInformer_Location.Name = "RTOInformer_Location"
        Me.RTOInformer_Location.Size = New System.Drawing.Size(419, 20)
        Me.RTOInformer_Location.TabIndex = 7
        Me.RTOInformer_Location.Text = "\\education.vic.gov.au\SHARE\TMO\Projects\RTOInformer\ControlItems_TMS\"
        Me.RTOInformer_Location.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Button1
        '
        Me.Button1.FlatAppearance.BorderColor = System.Drawing.Color.Maroon
        Me.Button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Maroon
        Me.Button1.Location = New System.Drawing.Point(457, 3)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(92, 20)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Refresh"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Profile_Only
        '
        Me.Profile_Only.AutoSize = True
        Me.Profile_Only.Checked = True
        Me.Profile_Only.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Profile_Only.Location = New System.Drawing.Point(472, 32)
        Me.Profile_Only.Name = "Profile_Only"
        Me.Profile_Only.Size = New System.Drawing.Size(77, 17)
        Me.Profile_Only.TabIndex = 9
        Me.Profile_Only.Text = "Profile only"
        Me.Profile_Only.UseVisualStyleBackColor = True
        '
        'ControlForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.ClientSize = New System.Drawing.Size(554, 92)
        Me.Controls.Add(Me.Profile_Only)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.RTOInformer_Location)
        Me.Controls.Add(Me.Stopper)
        Me.Name = "ControlForm"
        Me.RightToLeftLayout = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Admin Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Stopper As Windows.Forms.Button
    Friend WithEvents RTOInformer_Location As Windows.Forms.TextBox
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents Profile_Only As Windows.Forms.CheckBox
End Class
