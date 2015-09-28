Public Class PreferencesDialog
    Inherits System.Windows.Forms.Form

    Dim theMainForm As MainForm

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal myMainForm As MainForm)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        theMainForm = myMainForm

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
    Friend WithEvents tbImagesRootFolder As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbDatabaseFileName As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowseRootFolder As System.Windows.Forms.Button
    Friend WithEvents btnDatabaseFileName As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbReportsOutputFolder As System.Windows.Forms.TextBox
    Friend WithEvents btnReportsOutputFolder As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbServerName As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cbCameraClubName As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbServerScriptDir As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbNumJudges As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PreferencesDialog))
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbImagesRootFolder = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbDatabaseFileName = New System.Windows.Forms.TextBox
        Me.btnBrowseRootFolder = New System.Windows.Forms.Button
        Me.btnDatabaseFileName = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbReportsOutputFolder = New System.Windows.Forms.TextBox
        Me.btnReportsOutputFolder = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbServerName = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cbCameraClubName = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbServerScriptDir = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.cbNumJudges = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(201, 28)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Root Folder for Images"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tbImagesRootFolder
        '
        Me.tbImagesRootFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbImagesRootFolder.Location = New System.Drawing.Point(221, 9)
        Me.tbImagesRootFolder.Name = "tbImagesRootFolder"
        Me.tbImagesRootFolder.Size = New System.Drawing.Size(345, 26)
        Me.tbImagesRootFolder.TabIndex = 1
        Me.tbImagesRootFolder.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(10, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(201, 28)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Database File Name"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tbDatabaseFileName
        '
        Me.tbDatabaseFileName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbDatabaseFileName.Location = New System.Drawing.Point(221, 46)
        Me.tbDatabaseFileName.Name = "tbDatabaseFileName"
        Me.tbDatabaseFileName.Size = New System.Drawing.Size(345, 26)
        Me.tbDatabaseFileName.TabIndex = 3
        Me.tbDatabaseFileName.Text = ""
        '
        'btnBrowseRootFolder
        '
        Me.btnBrowseRootFolder.Location = New System.Drawing.Point(576, 9)
        Me.btnBrowseRootFolder.Name = "btnBrowseRootFolder"
        Me.btnBrowseRootFolder.Size = New System.Drawing.Size(75, 28)
        Me.btnBrowseRootFolder.TabIndex = 4
        Me.btnBrowseRootFolder.Text = "Browse..."
        '
        'btnDatabaseFileName
        '
        Me.btnDatabaseFileName.Location = New System.Drawing.Point(576, 46)
        Me.btnDatabaseFileName.Name = "btnDatabaseFileName"
        Me.btnDatabaseFileName.Size = New System.Drawing.Size(75, 28)
        Me.btnDatabaseFileName.TabIndex = 5
        Me.btnDatabaseFileName.Text = "Browse..."
        '
        'btnSave
        '
        Me.btnSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnSave.Location = New System.Drawing.Point(223, 272)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(105, 28)
        Me.btnSave.TabIndex = 6
        Me.btnSave.Text = "Save"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(348, 272)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(96, 28)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(10, 83)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(201, 28)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Reports Output Folder"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tbReportsOutputFolder
        '
        Me.tbReportsOutputFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbReportsOutputFolder.Location = New System.Drawing.Point(221, 83)
        Me.tbReportsOutputFolder.Name = "tbReportsOutputFolder"
        Me.tbReportsOutputFolder.Size = New System.Drawing.Size(345, 26)
        Me.tbReportsOutputFolder.TabIndex = 9
        Me.tbReportsOutputFolder.Text = ""
        '
        'btnReportsOutputFolder
        '
        Me.btnReportsOutputFolder.Location = New System.Drawing.Point(576, 83)
        Me.btnReportsOutputFolder.Name = "btnReportsOutputFolder"
        Me.btnReportsOutputFolder.Size = New System.Drawing.Size(75, 28)
        Me.btnReportsOutputFolder.TabIndex = 10
        Me.btnReportsOutputFolder.Text = "Browse..."
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(10, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(201, 28)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Server Host Name"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tbServerName
        '
        Me.tbServerName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbServerName.Location = New System.Drawing.Point(221, 120)
        Me.tbServerName.Name = "tbServerName"
        Me.tbServerName.Size = New System.Drawing.Size(345, 26)
        Me.tbServerName.TabIndex = 12
        Me.tbServerName.Text = ""
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(10, 192)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(201, 28)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Camera Club"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbCameraClubName
        '
        Me.cbCameraClubName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbCameraClubName.Location = New System.Drawing.Point(221, 192)
        Me.cbCameraClubName.Name = "cbCameraClubName"
        Me.cbCameraClubName.Size = New System.Drawing.Size(345, 28)
        Me.cbCameraClubName.TabIndex = 16
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(10, 157)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(201, 28)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Server script virt. dir."
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tbServerScriptDir
        '
        Me.tbServerScriptDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbServerScriptDir.Location = New System.Drawing.Point(221, 157)
        Me.tbServerScriptDir.Name = "tbServerScriptDir"
        Me.tbServerScriptDir.Size = New System.Drawing.Size(345, 26)
        Me.tbServerScriptDir.TabIndex = 18
        Me.tbServerScriptDir.Text = ""
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(10, 227)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(201, 28)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Number of Judges"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbNumJudges
        '
        Me.cbNumJudges.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbNumJudges.Items.AddRange(New Object() {"1", "2", "3"})
        Me.cbNumJudges.Location = New System.Drawing.Point(221, 227)
        Me.cbNumJudges.Name = "cbNumJudges"
        Me.cbNumJudges.Size = New System.Drawing.Size(64, 28)
        Me.cbNumJudges.TabIndex = 20
        '
        'PreferencesDialog
        '
        Me.AcceptButton = Me.btnSave
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(666, 311)
        Me.Controls.Add(Me.cbNumJudges)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.tbServerScriptDir)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cbCameraClubName)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.tbServerName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnReportsOutputFolder)
        Me.Controls.Add(Me.tbReportsOutputFolder)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnDatabaseFileName)
        Me.Controls.Add(Me.btnBrowseRootFolder)
        Me.Controls.Add(Me.tbDatabaseFileName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbImagesRootFolder)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "PreferencesDialog"
        Me.Text = "Preferences"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnBrowseRootFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseRootFolder.Click
        Dim folderBrowser As New FolderBrowserDialog

        Try
            folderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
            folderBrowser.SelectedPath = theMainForm.imagesRootFolder
            folderBrowser.ShowDialog()
            tbImagesRootFolder.Text = folderBrowser.SelectedPath
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in btnBrowseRootFolder_Click()")
        End Try
    End Sub

    Private Sub btnDatabaseFileName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDatabaseFileName.Click
        Dim fileOpenDialog As New OpenFileDialog
        Try
            fileOpenDialog.InitialDirectory = Application.StartupPath
            fileOpenDialog.Filter = "Access Databases (*.mdb)|*.mdb"
            fileOpenDialog.ShowDialog()
            tbDatabaseFileName.Text = fileOpenDialog.FileName
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in btnDatabaseFileName_Click()")
        End Try
    End Sub

    Private Sub btnReportsOutputFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReportsOutputFolder.Click
        Dim folderBrowser As New FolderBrowserDialog

        Try
            folderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
            folderBrowser.SelectedPath = Application.StartupPath
            folderBrowser.ShowDialog()
            tbReportsOutputFolder.Text = folderBrowser.SelectedPath
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in btnReportsOutputFolder_Click()")
        End Try

    End Sub


End Class
