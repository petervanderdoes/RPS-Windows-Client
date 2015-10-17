Imports System.ComponentModel

Public Class DownloadCompetitionsDialog
    Inherits Form

    Private ReadOnly _the_main_form As MainForm

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal myMainForm As MainForm,
                   ByVal defaultUsername As String,
                   ByVal comp_dates As ArrayList,
                   ByVal rootFolder As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _the_main_form = myMainForm

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

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DownloadCompetitionsDialog))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Username = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Password = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CompetitionDate = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OutputFolder = New System.Windows.Forms.TextBox()
        Me.BrowseButton = New System.Windows.Forms.Button()
        Me.OkButton = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Download_digital = New System.Windows.Forms.CheckBox()
        Me.Download_prints = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(14, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Username:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Username
        '
        Me.Username.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Username.Location = New System.Drawing.Point(124, 12)
        Me.Username.Margin = New System.Windows.Forms.Padding(5, 11, 3, 3)
        Me.Username.Name = "Username"
        Me.Username.Size = New System.Drawing.Size(170, 23)
        Me.Username.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(14, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Password:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Password
        '
        Me.Password.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Password.Location = New System.Drawing.Point(124, 45)
        Me.Password.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
        Me.Password.Name = "Password"
        Me.Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Password.Size = New System.Drawing.Size(170, 23)
        Me.Password.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(14, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 23)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Competition Date:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CompetitionDate
        '
        Me.CompetitionDate.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CompetitionDate.Location = New System.Drawing.Point(124, 78)
        Me.CompetitionDate.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
        Me.CompetitionDate.Name = "CompetitionDate"
        Me.CompetitionDate.Size = New System.Drawing.Size(170, 23)
        Me.CompetitionDate.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(14, 139)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 23)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Output Folder:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'OutputFolder
        '
        Me.OutputFolder.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OutputFolder.Location = New System.Drawing.Point(124, 139)
        Me.OutputFolder.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
        Me.OutputFolder.Name = "OutputFolder"
        Me.OutputFolder.Size = New System.Drawing.Size(336, 23)
        Me.OutputFolder.TabIndex = 7
        '
        'BrowseButton
        '
        Me.BrowseButton.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BrowseButton.Location = New System.Drawing.Point(470, 139)
        Me.BrowseButton.Margin = New System.Windows.Forms.Padding(7, 3, 3, 3)
        Me.BrowseButton.Name = "BrowseButton"
        Me.BrowseButton.Size = New System.Drawing.Size(75, 23)
        Me.BrowseButton.TabIndex = 8
        Me.BrowseButton.Text = "Browse..."
        '
        'OkButton
        '
        Me.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OkButton.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OkButton.Location = New System.Drawing.Point(224, 176)
        Me.OkButton.Margin = New System.Windows.Forms.Padding(7, 11, 3, 3)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(75, 23)
        Me.OkButton.TabIndex = 9
        Me.OkButton.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancel_Button.Location = New System.Drawing.Point(309, 176)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(7, 3, 3, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(75, 23)
        Me.Cancel_Button.TabIndex = 10
        Me.Cancel_Button.Text = "Cancel"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(14, 108)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(104, 23)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Competitions:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Download_digital
        '
        Me.Download_digital.Checked = True
        Me.Download_digital.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Download_digital.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Download_digital.Location = New System.Drawing.Point(124, 111)
        Me.Download_digital.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
        Me.Download_digital.Name = "Download_digital"
        Me.Download_digital.Size = New System.Drawing.Size(65, 17)
        Me.Download_digital.TabIndex = 12
        Me.Download_digital.Text = "Digital"
        '
        'Download_prints
        '
        Me.Download_prints.Checked = True
        Me.Download_prints.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Download_prints.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Download_prints.Location = New System.Drawing.Point(195, 112)
        Me.Download_prints.Name = "Download_prints"
        Me.Download_prints.Size = New System.Drawing.Size(88, 17)
        Me.Download_prints.TabIndex = 13
        Me.Download_prints.Text = "Prints"
        '
        'DownloadCompetitionsDialog
        '
        Me.AcceptButton = Me.OkButton
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 16)
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(560, 213)
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
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DownloadCompetitionsDialog"
        Me.Padding = New System.Windows.Forms.Padding(11)
        Me.Text = "Download Competitions"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private Sub BrowseButton_Click(sender As Object, e As EventArgs) Handles BrowseButton.Click
        'Dim fileOpenDialog As New OpenFileDialog

        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.SelectedPath = _the_main_form.images_root_folder
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            OutputFolder.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Download_Competitions_Dialog_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing
        Dim invalid_input As Boolean = False
        Dim err_msg As String = "Please correct the following errors" + vbCrLf + vbCrLf

        If DialogResult = DialogResult.OK Then
            If Username.Text = "" Then
                err_msg = err_msg + "* Username is empty" + vbCrLf
                invalid_input = True
            End If
            If Password.Text = "" Then
                err_msg = err_msg + "* Password is empty" + vbCrLf
                invalid_input = True
            End If
            If (Not Download_digital.Checked) And (Not Download_prints.Checked) Then
                err_msg = err_msg + "* No competition type selected" + vbCrLf
                invalid_input = True
            End If
            If OutputFolder.Text = "" Then
                err_msg = err_msg + "* Output Folder is empty" + vbCrLf
                invalid_input = True
            End If
            If invalid_input = True Then
                MessageBox.Show(err_msg, "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                e.Cancel = True
            Else
                e.Cancel = False
            End If
        End If
    End Sub
End Class
