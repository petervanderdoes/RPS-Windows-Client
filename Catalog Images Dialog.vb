Public Class Catalog_Images_Dialog
    Inherits System.Windows.Forms.Form

    Dim theMainForm As MainForm
    Dim imageSelectionMode As String
    Dim fileNames As ArrayList

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal myMainForm As MainForm, ByVal mode As String, ByVal names As ArrayList)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        theMainForm = myMainForm
        imageSelectionMode = mode
        fileNames = names

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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents dpCompetitionDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents tbTheme As System.Windows.Forms.TextBox
    Friend WithEvents cbMedium As System.Windows.Forms.ComboBox
    Friend WithEvents cbClassification As System.Windows.Forms.ComboBox
    Friend WithEvents tbNewImageFolder As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Catalog_Images_Dialog))
        Me.Label1 = New System.Windows.Forms.Label
        Me.dpCompetitionDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbTheme = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cbMedium = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cbClassification = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbNewImageFolder = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(21, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(160, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Competition Date"
        '
        'dpCompetitionDate
        '
        Me.dpCompetitionDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dpCompetitionDate.Location = New System.Drawing.Point(192, 10)
        Me.dpCompetitionDate.Name = "dpCompetitionDate"
        Me.dpCompetitionDate.Size = New System.Drawing.Size(224, 26)
        Me.dpCompetitionDate.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(21, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(150, 28)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Theme"
        '
        'tbTheme
        '
        Me.tbTheme.Location = New System.Drawing.Point(192, 48)
        Me.tbTheme.Name = "tbTheme"
        Me.tbTheme.Size = New System.Drawing.Size(395, 26)
        Me.tbTheme.TabIndex = 3
        Me.tbTheme.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(21, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(150, 28)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Medium"
        '
        'cbMedium
        '
        Me.cbMedium.Location = New System.Drawing.Point(192, 86)
        Me.cbMedium.Name = "cbMedium"
        Me.cbMedium.Size = New System.Drawing.Size(224, 28)
        Me.cbMedium.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(21, 124)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(150, 28)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Classification"
        '
        'cbClassification
        '
        Me.cbClassification.Location = New System.Drawing.Point(192, 124)
        Me.cbClassification.Name = "cbClassification"
        Me.cbClassification.Size = New System.Drawing.Size(224, 28)
        Me.cbClassification.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(21, 162)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(150, 28)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Folder Path"
        '
        'tbNewImageFolder
        '
        Me.tbNewImageFolder.Location = New System.Drawing.Point(192, 162)
        Me.tbNewImageFolder.Name = "tbNewImageFolder"
        Me.tbNewImageFolder.Size = New System.Drawing.Size(395, 26)
        Me.tbNewImageFolder.TabIndex = 9
        Me.tbNewImageFolder.Text = ""
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(619, 162)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(106, 28)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Browse..."
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button2.Location = New System.Drawing.Point(235, 209)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(117, 29)
        Me.Button2.TabIndex = 11
        Me.Button2.Text = "OK"
        '
        'Button3
        '
        Me.Button3.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button3.Location = New System.Drawing.Point(384, 209)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(117, 29)
        Me.Button3.TabIndex = 12
        Me.Button3.Text = "Cancel"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'Catalog_Images_Dialog
        '
        Me.AcceptButton = Me.Button2
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.CancelButton = Me.Button3
        Me.ClientSize = New System.Drawing.Size(738, 256)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.tbNewImageFolder)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cbClassification)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cbMedium)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbTheme)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dpCompetitionDate)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Catalog_Images_Dialog"
        Me.Text = "Catalog Images"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim folderName As String
        Dim fileOpenDialog As New OpenFileDialog
        'Dim imageFileName As String
        Dim i As Integer
        'Dim dir As Directory
        'Dim workingDir As String

        Try
            If imageSelectionMode = "Folder" Then
                ' Get the folder name containing the new images
                FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
                FolderBrowserDialog1.SelectedPath = theMainForm.images_root_folder
                If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
                    tbNewImageFolder.Text = FolderBrowserDialog1.SelectedPath
                End If
            Else ' Select individual images
                fileOpenDialog.InitialDirectory = theMainForm.images_root_folder
                fileOpenDialog.Filter = "JPEG Images (*.jpg)|*.jpg|All files (*.*)|*.*"
                fileOpenDialog.Multiselect = True
                fileOpenDialog.RestoreDirectory = True
                If fileOpenDialog.ShowDialog() = DialogResult.Cancel Then
                    Exit Sub
                End If
                ' Copy the returned list of files to the ArrayList that was passed in
                ' to the dialog
                For i = LBound(fileOpenDialog.FileNames) To UBound(fileOpenDialog.FileNames)
                    fileNames.Add(fileOpenDialog.FileNames(i))
                    tbNewImageFolder.Text = tbNewImageFolder.Text + fileOpenDialog.FileNames(i) + ";"
                Next
                ' Strip off the trailing semicolon
                tbNewImageFolder.Text = Mid(tbNewImageFolder.Text, 1, Len(tbNewImageFolder.Text) - 1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in Catalog_Images_Dialog.Button1_Click()")
        End Try
    End Sub

    Private Sub tbTheme_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbTheme.Validated
        If tbTheme.Text = "" Then
            ErrorProvider1.SetError(tbTheme, "Please provide a value for Theme")
        Else
            ErrorProvider1.SetError(tbTheme, "")
        End If
    End Sub

    Private Sub tbNewImageFolder_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbNewImageFolder.Validated
        If tbNewImageFolder.Text = "" Then
            ErrorProvider1.SetError(tbNewImageFolder, "Please provide a value for Theme")
        Else
            ErrorProvider1.SetError(tbNewImageFolder, "")
        End If
    End Sub

    Private Sub Catalog_Images_Dialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim InvalidInput As Boolean = False
        Dim errMsg As String = "Please correct the following errors" + vbCrLf + vbCrLf

        If Me.DialogResult = DialogResult.OK Then
            If tbTheme.Text = "" Then
                errMsg = errMsg + "* Theme is empty" + vbCrLf
                InvalidInput = True
            End If
            If tbNewImageFolder.Text = "" Then
                errMsg = errMsg + "* Folder Path is empty" + vbCrLf
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


    Private Sub Catalog_Images_Dialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer

        If imageSelectionMode = "Files" Then
            Label5.Text = "Files"
        End If

        ' Populate the list of classifications
        cbClassification.Items.Clear()
        For i = 0 To theMainForm.classifications.Count - 1
            cbClassification.Items.Add(theMainForm.classifications.Item(i))
        Next
        cbClassification.SelectedItem = theMainForm.classifications.Item(0)

        ' Populate the list of mediums
        cbMedium.Items.Clear()
        For i = 0 To theMainForm.mediums.Count - 1
            cbMedium.Items.Add(theMainForm.mediums.Item(i))
        Next
        cbMedium.SelectedItem = theMainForm.mediums.Item(0)
    End Sub
End Class
