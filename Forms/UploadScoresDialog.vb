

Namespace Forms
    Public Class UploadScoresDialog
        Inherits Form

#Region " Windows Form Designer generated code "

        Public Sub New(ByVal defaultUsername As String, ByVal comp_dates As ArrayList)
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call
            Username.Text = defaultUsername
            Username.Select(0, 0)

            Dim i As Integer
            For i = 0 To comp_dates.Count - 1
                CompDate.Items.Add(comp_dates.Item(i))
                CompDate.SelectedIndex = 0
            Next
        End Sub

        'Form overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        Friend WithEvents Label1 As Label
        Friend WithEvents Label2 As Label
        Friend WithEvents Label3 As Label
        Friend WithEvents Username As System.Windows.Forms.TextBox
        Friend WithEvents Password As System.Windows.Forms.TextBox
        Friend WithEvents CompDate As System.Windows.Forms.ComboBox
        Friend WithEvents OKbutton As System.Windows.Forms.Button
        Friend WithEvents Cancel_Button As System.Windows.Forms.Button
        Friend WithEvents Upload_digital_scores As System.Windows.Forms.CheckBox
        Friend WithEvents Upload_print_scores As System.Windows.Forms.CheckBox

        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Dim resources As System.ComponentModel.ComponentResourceManager =
                    New System.ComponentModel.ComponentResourceManager(GetType(UploadScoresDialog))
            Me.Label1 = New Label()
            Me.Label2 = New Label()
            Me.Label3 = New Label()
            Me.Username = New System.Windows.Forms.TextBox()
            Me.Password = New System.Windows.Forms.TextBox()
            Me.CompDate = New System.Windows.Forms.ComboBox()
            Me.OKbutton = New System.Windows.Forms.Button()
            Me.Cancel_Button = New System.Windows.Forms.Button()
            Me.Upload_digital_scores = New System.Windows.Forms.CheckBox()
            Me.Upload_print_scores = New System.Windows.Forms.CheckBox()
            Me.SuspendLayout()
            '
            'Label1
            '
            Me.Label1.Font = New System.Drawing.Font("Segoe UI",
                                                     9.0!,
                                                     System.Drawing.FontStyle.Regular,
                                                     System.Drawing.GraphicsUnit.Point,
                                                     CType(0, Byte))
            Me.Label1.Location = New Point(11, 14)
            Me.Label1.Margin = New System.Windows.Forms.Padding(0)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New Size(105, 23)
            Me.Label1.TabIndex = 0
            Me.Label1.Text = "Username:"
            Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            '
            'Label2
            '
            Me.Label2.Font = New System.Drawing.Font("Segoe UI",
                                                     9.0!,
                                                     System.Drawing.FontStyle.Regular,
                                                     System.Drawing.GraphicsUnit.Point,
                                                     CType(0, Byte))
            Me.Label2.Location = New Point(11, 46)
            Me.Label2.Margin = New System.Windows.Forms.Padding(0)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New Size(105, 23)
            Me.Label2.TabIndex = 1
            Me.Label2.Text = "Password:"
            Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            '
            'Label3
            '
            Me.Label3.Font = New System.Drawing.Font("Segoe UI",
                                                     9.0!,
                                                     System.Drawing.FontStyle.Regular,
                                                     System.Drawing.GraphicsUnit.Point,
                                                     CType(0, Byte))
            Me.Label3.Location = New Point(11, 74)
            Me.Label3.Margin = New System.Windows.Forms.Padding(0)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New Size(105, 23)
            Me.Label3.TabIndex = 2
            Me.Label3.Text = "Competition Date:"
            Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            '
            'Username
            '
            Me.Username.Font = New System.Drawing.Font("Segoe UI",
                                                       9.0!,
                                                       System.Drawing.FontStyle.Regular,
                                                       System.Drawing.GraphicsUnit.Point,
                                                       CType(0, Byte))
            Me.Username.Location = New Point(119, 14)
            Me.Username.Margin = New System.Windows.Forms.Padding(5, 3, 11, 3)
            Me.Username.Name = "Username"
            Me.Username.Size = New Size(170, 23)
            Me.Username.TabIndex = 3
            '
            'Password
            '
            Me.Password.Font = New System.Drawing.Font("Segoe UI",
                                                       9.0!,
                                                       System.Drawing.FontStyle.Regular,
                                                       System.Drawing.GraphicsUnit.Point,
                                                       CType(0, Byte))
            Me.Password.Location = New Point(119, 47)
            Me.Password.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
            Me.Password.Name = "Password"
            Me.Password.PasswordChar = ChrW(42)
            Me.Password.Size = New Size(170, 23)
            Me.Password.TabIndex = 4
            '
            'CompDate
            '
            Me.CompDate.Font = New System.Drawing.Font("Segoe UI",
                                                       9.0!,
                                                       System.Drawing.FontStyle.Regular,
                                                       System.Drawing.GraphicsUnit.Point,
                                                       CType(0, Byte))
            Me.CompDate.ItemHeight = 15
            Me.CompDate.Location = New Point(119, 80)
            Me.CompDate.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
            Me.CompDate.Name = "CompDate"
            Me.CompDate.Size = New Size(170, 23)
            Me.CompDate.TabIndex = 5
            '
            'OKbutton
            '
            Me.OKbutton.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.OKbutton.Font = New System.Drawing.Font("Segoe UI",
                                                       9.0!,
                                                       System.Drawing.FontStyle.Regular,
                                                       System.Drawing.GraphicsUnit.Point,
                                                       CType(0, Byte))
            Me.OKbutton.Location = New Point(77, 171)
            Me.OKbutton.Name = "OKbutton"
            Me.OKbutton.Size = New Size(75, 23)
            Me.OKbutton.TabIndex = 6
            Me.OKbutton.Text = "OK"
            '
            'Cancel_Button
            '
            Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Cancel_Button.Font = New System.Drawing.Font("Segoe UI",
                                                            9.0!,
                                                            System.Drawing.FontStyle.Regular,
                                                            System.Drawing.GraphicsUnit.Point,
                                                            CType(0, Byte))
            Me.Cancel_Button.Location = New Point(158, 171)
            Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(3, 11, 3, 11)
            Me.Cancel_Button.Name = "Cancel_Button"
            Me.Cancel_Button.Size = New Size(75, 23)
            Me.Cancel_Button.TabIndex = 7
            Me.Cancel_Button.Text = "Cancel"
            '
            'Upload_digital_scores
            '
            Me.Upload_digital_scores.Checked = True
            Me.Upload_digital_scores.CheckState = System.Windows.Forms.CheckState.Checked
            Me.Upload_digital_scores.Font = New System.Drawing.Font("Segoe UI",
                                                                    8.25!,
                                                                    System.Drawing.FontStyle.Regular,
                                                                    System.Drawing.GraphicsUnit.Point,
                                                                    CType(0, Byte))
            Me.Upload_digital_scores.Location = New Point(119, 117)
            Me.Upload_digital_scores.Margin = New System.Windows.Forms.Padding(3, 11, 3, 3)
            Me.Upload_digital_scores.Name = "Upload_digital_scores"
            Me.Upload_digital_scores.Size = New Size(144, 17)
            Me.Upload_digital_scores.TabIndex = 8
            Me.Upload_digital_scores.Text = "Upload Digital Scores"
            '
            'Upload_print_scores
            '
            Me.Upload_print_scores.Checked = True
            Me.Upload_print_scores.CheckState = System.Windows.Forms.CheckState.Checked
            Me.Upload_print_scores.Font = New System.Drawing.Font("Segoe UI",
                                                                  8.25!,
                                                                  System.Drawing.FontStyle.Regular,
                                                                  System.Drawing.GraphicsUnit.Point,
                                                                  CType(0, Byte))
            Me.Upload_print_scores.Location = New Point(119, 140)
            Me.Upload_print_scores.Name = "Upload_print_scores"
            Me.Upload_print_scores.Size = New Size(144, 17)
            Me.Upload_print_scores.TabIndex = 9
            Me.Upload_print_scores.Text = "Upload Print Scores"
            '
            'UploadScoresDialog
            '
            Me.AcceptButton = Me.OKbutton
            Me.AutoScaleBaseSize = New Size(6, 16)
            Me.CancelButton = Me.Cancel_Button
            Me.ClientSize = New Size(313, 209)
            Me.Controls.Add(Me.Upload_print_scores)
            Me.Controls.Add(Me.Upload_digital_scores)
            Me.Controls.Add(Me.Cancel_Button)
            Me.Controls.Add(Me.OKbutton)
            Me.Controls.Add(Me.CompDate)
            Me.Controls.Add(Me.Password)
            Me.Controls.Add(Me.Username)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.Label1)
            Me.Font = New System.Drawing.Font("Segoe UI",
                                              9.0!,
                                              System.Drawing.FontStyle.Regular,
                                              System.Drawing.GraphicsUnit.Point,
                                              CType(0, Byte))
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.Name = "UploadScoresDialog"
            Me.Padding = New System.Windows.Forms.Padding(11)
            Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
            Me.Text = "Upload Scores"
            Me.ResumeLayout(False)
            Me.PerformLayout()
        End Sub

#End Region

        Private Sub Upload_Scores_Dialog_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) _
            Handles MyBase.Closing
            Dim is_invalid_input As Boolean = False

            Dim err_msg As String = "Please correct the following errors" + vbCrLf + vbCrLf

            If DialogResult = DialogResult.OK Then
                If Username.Text = "" Then
                    err_msg = err_msg + "* Username is empty" + vbCrLf
                    is_invalid_input = True
                End If
                If Password.Text = "" Then
                    err_msg = err_msg + "* Password is empty" + vbCrLf
                    is_invalid_input = True
                End If
                If (Not Upload_digital_scores.Checked) And (Not Upload_print_scores.Checked) Then
                    err_msg = err_msg + "* No competition type selected" + vbCrLf
                    is_invalid_input = True
                End If
                If is_invalid_input = True Then
                    MessageBox.Show(err_msg, "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    e.Cancel = True
                Else
                    e.Cancel = False
                End If
            End If
        End Sub
    End Class
End Namespace
