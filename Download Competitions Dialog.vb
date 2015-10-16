Imports System.ComponentModel

Public Class Download_Competitions_Dialog
    Inherits Form

    Dim theMainForm As MainForm
#Region " Windows Form Designer generated code "

    Public Sub New(ByVal myMainForm As MainForm, ByVal defaultUsername As String, ByVal comp_dates As ArrayList, ByVal rootFolder As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        theMainForm = myMainForm

        Username.Text = defaultUsername
        Username.Select(0, 0)

        Dim i As Integer
        For i = 0 To comp_dates.Count - 1
            CompetitionDate.Items.Add(comp_dates.Item(i))
        Next
        CompetitionDate.SelectedIndex = 0

        OutputFolder.Text = rootFolder
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Username As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Password As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CompetitionDate As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents OutputFolder As System.Windows.Forms.TextBox
    Friend WithEvents BrowseButton As System.Windows.Forms.Button
    Friend WithEvents OkButton As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Download_digital As System.Windows.Forms.CheckBox
    Friend WithEvents Download_prints As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Download_Competitions_Dialog))
        Me.Label1 = New System.Windows.Forms.Label
        Me.Username = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Password = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.CompetitionDate = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.OutputFolder = New System.Windows.Forms.TextBox
        Me.BrowseButton = New System.Windows.Forms.Button
        Me.OkButton = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Label5 = New System.Windows.Forms.Label
        Me.Download_digital = New System.Windows.Forms.CheckBox
        Me.Download_prints = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(144, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Username:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Username
        '
        Me.Username.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Username.Location = New System.Drawing.Point(176, 16)
        Me.Username.Name = "Username"
        Me.Username.Size = New System.Drawing.Size(208, 23)
        Me.Username.TabIndex = 1
        Me.Username.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(144, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Password:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Password
        '
        Me.Password.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Password.Location = New System.Drawing.Point(176, 48)
        Me.Password.Name = "Password"
        Me.Password.PasswordChar = ChrW(42)
        Me.Password.Size = New System.Drawing.Size(208, 23)
        Me.Password.TabIndex = 3
        Me.Password.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 24)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Competition Date:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CompetitionDate
        '
        Me.CompetitionDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CompetitionDate.Location = New System.Drawing.Point(176, 80)
        Me.CompetitionDate.Name = "CompetitionDate"
        Me.CompetitionDate.Size = New System.Drawing.Size(208, 24)
        Me.CompetitionDate.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 160)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(144, 24)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Output Folder:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OutputFolder
        '
        Me.OutputFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OutputFolder.Location = New System.Drawing.Point(176, 160)
        Me.OutputFolder.Name = "OutputFolder"
        Me.OutputFolder.Size = New System.Drawing.Size(336, 23)
        Me.OutputFolder.TabIndex = 7
        Me.OutputFolder.Text = ""
        '
        'BrowseButton
        '
        Me.BrowseButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BrowseButton.Location = New System.Drawing.Point(520, 160)
        Me.BrowseButton.Name = "BrowseButton"
        Me.BrowseButton.Size = New System.Drawing.Size(72, 24)
        Me.BrowseButton.TabIndex = 8
        Me.BrowseButton.Text = "Browse..."
        '
        'OkButton
        '
        Me.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OkButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OkButton.Location = New System.Drawing.Point(216, 200)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(80, 24)
        Me.OkButton.TabIndex = 9
        Me.OkButton.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancel_Button.Location = New System.Drawing.Point(320, 200)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(72, 24)
        Me.Cancel_Button.TabIndex = 10
        Me.Cancel_Button.Text = "Cancel"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(144, 24)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Competitions:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Download_digital
        '
        Me.Download_digital.Checked = True
        Me.Download_digital.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Download_digital.Location = New System.Drawing.Point(176, 112)
        Me.Download_digital.Name = "Download_digital"
        Me.Download_digital.Size = New System.Drawing.Size(88, 16)
        Me.Download_digital.TabIndex = 12
        Me.Download_digital.Text = "Digital"
        '
        'Download_prints
        '
        Me.Download_prints.Checked = True
        Me.Download_prints.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Download_prints.Location = New System.Drawing.Point(176, 136)
        Me.Download_prints.Name = "Download_prints"
        Me.Download_prints.Size = New System.Drawing.Size(88, 16)
        Me.Download_prints.TabIndex = 13
        Me.Download_prints.Text = "Prints"
        '
        'Download_Competitions_Dialog
        '
        Me.AcceptButton = Me.OkButton
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(608, 240)
        Me.Controls.Add(Me.Download_prints)
        Me.Controls.Add(Me.Download_digital)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.OkButton)
        Me.Controls.Add(Me.BrowseButton)
        Me.Controls.Add(Me.OutputFolder)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.CompetitionDate)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Password)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Username)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Download_Competitions_Dialog"
        Me.Text = "Download Competitions"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub BrowseButton_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BrowseButton.Click
        'Dim fileOpenDialog As New OpenFileDialog

        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.SelectedPath = theMainForm.images_root_folder
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            OutputFolder.Text = FolderBrowserDialog1.SelectedPath
        End If

    End Sub

    Private Sub Download_Competitions_Dialog_Closing(ByVal sender As Object, ByVal e As CancelEventArgs) Handles MyBase.Closing
        Dim InvalidInput As Boolean = False
        Dim errMsg As String = "Please correct the following errors" + vbCrLf + vbCrLf

        If Me.DialogResult = DialogResult.OK Then
            If Username.Text = "" Then
                errMsg = errMsg + "* Username is empty" + vbCrLf
                InvalidInput = True
            End If
            If Password.Text = "" Then
                errMsg = errMsg + "* Password is empty" + vbCrLf
                InvalidInput = True
            End If
            If (Not Download_digital.Checked) And (Not Download_prints.Checked) Then
                errMsg = errMsg + "* No competition type selected" + vbCrLf
                InvalidInput = True
            End If
            If OutputFolder.Text = "" Then
                errMsg = errMsg + "* Output Folder is empty" + vbCrLf
                InvalidInput = True
            End If
            If InvalidInput = True Then
                MessageBox.Show(errMsg, "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                e.Cancel = True
            Else
                e.Cancel = False
            End If
        End If

    End Sub
End Class
