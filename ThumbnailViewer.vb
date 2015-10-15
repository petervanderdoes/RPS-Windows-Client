Public Class ThumbnailViewer
    Inherits System.Windows.Forms.Form

    Dim theMainForm As MainForm
    Dim ImageList As IList
    Dim numImages As Integer
    Dim currentIndex As Integer
    Dim currentFileName As String
    Dim currentRow As CompetitionEntry
    Dim fullSizeFileName As String
    Dim Zoomed As Boolean
    Dim thumbnailViewTitle As String

    ' API parameters for setting border select style for the listview
    Public Const LVM_FIRST As Integer = &H1000
    Public Const LVM_GETCOUNTPERPAGE As Integer = LVM_FIRST + 40
    Public Const WM_SETREDRAW As Integer = &HB

    Public Enum LVS_EX
        LVS_EX_GRIDLINES = &H1
        LVS_EX_SUBITEMIMAGES = &H2
        LVS_EX_CHECKBOXES = &H4
        LVS_EX_TRACKSELECT = &H8
        LVS_EX_HEADERDRAGDROP = &H10
        LVS_EX_FULLROWSELECT = &H20
        LVS_EX_ONECLICKACTIVATE = &H40
        LVS_EX_TWOCLICKACTIVATE = &H80
        LVS_EX_FLATSB = &H100
        LVS_EX_REGIONAL = &H200
        LVS_EX_INFOTIP = &H400
        LVS_EX_UNDERLINEHOT = &H800
        LVS_EX_UNDERLINECOLD = &H1000
        LVS_EX_MULTIWORKAREAS = &H2000
        LVS_EX_LABELTIP = &H4000
        LVS_EX_BORDERSELECT = &H8000
        LVS_EX_DOUBLEBUFFER = &H10000
        LVS_EX_HIDELABELS = &H20000
        LVS_EX_SINGLEROW = &H40000
        LVS_EX_SNAPTOGRID = &H80000
        LVS_EX_SIMPLESELECT = &H100000
    End Enum 'LVS_EX

    Public Enum LVM
        LVM_FIRST = &H1000
        LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
        LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
    End Enum 'LVM

    ' Windows API for setting border select style on the listview
    Public Overloads Declare Auto Function SendMessage Lib "User32.dll" (ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal myMainForm As MainForm, ByVal ds As IList, ByVal screenTitle As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        theMainForm = myMainForm
        ImageList = ds
        thumbnailViewTitle = screenTitle
        setSizes()

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
    Friend WithEvents ThumbnailListView As System.Windows.Forms.ListView
    Friend WithEvents ThumbnailImageList As System.Windows.Forms.ImageList
    Friend WithEvents ZoomedImage As System.Windows.Forms.PictureBox
    Friend WithEvents ThumbnailViewTitleBar As System.Windows.Forms.Label
    Friend WithEvents Divider As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ThumbnailViewer))
        Me.ThumbnailListView = New System.Windows.Forms.ListView()
        Me.ThumbnailImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.ZoomedImage = New System.Windows.Forms.PictureBox()
        Me.ThumbnailViewTitleBar = New System.Windows.Forms.Label()
        Me.Divider = New System.Windows.Forms.Label()
        CType(Me.ZoomedImage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ThumbnailListView
        '
        Me.ThumbnailListView.BackColor = System.Drawing.Color.Black
        Me.ThumbnailListView.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ThumbnailListView.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ThumbnailListView.ForeColor = System.Drawing.Color.White
        Me.ThumbnailListView.LargeImageList = Me.ThumbnailImageList
        Me.ThumbnailListView.Location = New System.Drawing.Point(0, 43)
        Me.ThumbnailListView.Margin = New System.Windows.Forms.Padding(0)
        Me.ThumbnailListView.MultiSelect = False
        Me.ThumbnailListView.Name = "ThumbnailListView"
        Me.ThumbnailListView.Size = New System.Drawing.Size(1024, 725)
        Me.ThumbnailListView.TabIndex = 0
        Me.ThumbnailListView.UseCompatibleStateImageBehavior = False
        '
        'ThumbnailImageList
        '
        Me.ThumbnailImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit
        Me.ThumbnailImageList.ImageSize = New System.Drawing.Size(256, 256)
        Me.ThumbnailImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'ZoomedImage
        '
        Me.ZoomedImage.BackColor = System.Drawing.Color.Black
        Me.ZoomedImage.Location = New System.Drawing.Point(0, 0)
        Me.ZoomedImage.Margin = New System.Windows.Forms.Padding(0)
        Me.ZoomedImage.Name = "ZoomedImage"
        Me.ZoomedImage.Size = New System.Drawing.Size(1024, 768)
        Me.ZoomedImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.ZoomedImage.TabIndex = 1
        Me.ZoomedImage.TabStop = False
        Me.ZoomedImage.Visible = False
        '
        'ThumbnailViewTitleBar
        '
        Me.ThumbnailViewTitleBar.BackColor = System.Drawing.Color.Black
        Me.ThumbnailViewTitleBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ThumbnailViewTitleBar.ForeColor = System.Drawing.Color.Yellow
        Me.ThumbnailViewTitleBar.Location = New System.Drawing.Point(0, 0)
        Me.ThumbnailViewTitleBar.Margin = New System.Windows.Forms.Padding(0)
        Me.ThumbnailViewTitleBar.Name = "ThumbnailViewTitleBar"
        Me.ThumbnailViewTitleBar.Size = New System.Drawing.Size(1024, 37)
        Me.ThumbnailViewTitleBar.TabIndex = 2
        Me.ThumbnailViewTitleBar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Divider
        '
        Me.Divider.BackColor = System.Drawing.Color.White
        Me.Divider.Location = New System.Drawing.Point(0, 37)
        Me.Divider.Margin = New System.Windows.Forms.Padding(0)
        Me.Divider.Name = "Divider"
        Me.Divider.Size = New System.Drawing.Size(1024, 2)
        Me.Divider.TabIndex = 3
        '
        'ThumbnailViewer
        '
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(1024, 768)
        Me.ControlBox = False
        Me.Controls.Add(Me.Divider)
        Me.Controls.Add(Me.ThumbnailViewTitleBar)
        Me.Controls.Add(Me.ZoomedImage)
        Me.Controls.Add(Me.ThumbnailListView)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(1024, 768)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1024, 768)
        Me.Name = "ThumbnailViewer"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "ThumbnailViewer"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.ZoomedImage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ThumbnailListView_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ThumbnailListView.KeyDown
        Dim datasetRowNum As Integer
        Dim newIndex As Integer

        Try
            Select Case e.KeyCode
                Case Keys.Escape, Keys.X
                    Me.Close()
                Case Keys.F1
                    If ThumbnailListView.SelectedIndices.Count > 0 And
                       theMainForm.awards.Count >= 1 Then
                        ThumbnailListView.SelectedItems(0).Text = theMainForm.awards.Item(0)
                        currentIndex = ThumbnailListView.SelectedIndices(0)
                        datasetRowNum = ThumbnailListView.Items(currentIndex).Tag
                        currentRow = ImageList.Item(datasetRowNum)
                        currentRow.Award = theMainForm.awards.Item(0)
                    End If
                Case Keys.F2
                    If ThumbnailListView.SelectedIndices.Count > 0 And
                       theMainForm.awards.Count >= 2 Then
                        ThumbnailListView.SelectedItems(0).Text = theMainForm.awards.Item(1)
                        currentIndex = ThumbnailListView.SelectedIndices(0)
                        datasetRowNum = ThumbnailListView.Items(currentIndex).Tag
                        currentRow = ImageList.Item(datasetRowNum)
                        currentRow.Award = theMainForm.awards.Item(1)
                    End If
                Case Keys.F3
                    If ThumbnailListView.SelectedIndices.Count > 0 And
                       theMainForm.awards.Count >= 3 Then
                        ThumbnailListView.SelectedItems(0).Text = theMainForm.awards.Item(2)
                        currentIndex = ThumbnailListView.SelectedIndices(0)
                        datasetRowNum = ThumbnailListView.Items(currentIndex).Tag
                        currentRow = ImageList.Item(datasetRowNum)
                        currentRow.Award = theMainForm.awards.Item(2)
                    End If
                Case Keys.F4
                    If ThumbnailListView.SelectedIndices.Count > 0 And
                       theMainForm.awards.Count >= 4 Then
                        ThumbnailListView.SelectedItems(0).Text = theMainForm.awards.Item(3)
                        currentIndex = ThumbnailListView.SelectedIndices(0)
                        datasetRowNum = ThumbnailListView.Items(currentIndex).Tag
                        currentRow = ImageList.Item(datasetRowNum)
                        currentRow.Award = theMainForm.awards.Item(3)
                    End If
                Case Keys.Delete, Keys.Back
                    If ThumbnailListView.SelectedIndices.Count > 0 Then
                        ThumbnailListView.SelectedItems(0).Text = ""
                        currentIndex = ThumbnailListView.SelectedIndices(0)
                        datasetRowNum = ThumbnailListView.Items(currentIndex).Tag
                        currentRow = ImageList.Item(datasetRowNum)
                        currentRow.Award = Nothing
                    End If
                Case Keys.R
                    If ThumbnailListView.SelectedIndices.Count > 0 Then
                        ' Remove the currently selected item from the ListView
                        currentIndex = ThumbnailListView.SelectedIndices(0)
                        ThumbnailListView.Items.RemoveAt(currentIndex)
                        ' Make sure a thumbnail is still selected
                        If ThumbnailListView.Items.Count > 0 Then
                            newIndex = Math.Min(currentIndex, ThumbnailListView.Items.Count - 1)
                            ThumbnailListView.Items(newIndex).Selected = True
                        End If

                    End If
                Case Keys.Z
                    If ThumbnailListView.SelectedIndices.Count > 0 Then
                        If Not Zoomed Then
                            ' Retrieve the file name of the currently selected image from the
                            ' database
                            currentIndex = ThumbnailListView.SelectedIndices(0)
                            datasetRowNum = ThumbnailListView.Items(currentIndex).Tag
                            currentRow = ImageList.Item(datasetRowNum)
                            fullSizeFileName = currentRow.Image_File_Name
                            ' If it's a relative path, convert to an absolute path
                            If Not InStr(1, fullSizeFileName, ":\") = 2 Then
                                fullSizeFileName = theMainForm.imagesRootFolder + "\" + fullSizeFileName
                            End If
                            ' Load it into the PictureBox
                            ZoomedImage.Image = Image.FromFile(fullSizeFileName)
                            ' Hide the ListView and reveal the PictureBox
                            ThumbnailListView.Visible = False
                            ThumbnailViewTitleBar.Visible = False
                            Divider.Visible = False
                            ZoomedImage.Visible = True
                            Zoomed = True
                        Else
                            ThumbnailListView.Visible = True
                            ThumbnailViewTitleBar.Visible = True
                            Divider.Visible = True
                            ZoomedImage.Visible = False
                            Zoomed = False
                        End If
                    End If
                    e.Handled = True
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in ThumbnailListView_KeyDown()")
        End Try
    End Sub

    Private Sub ThumbnailViewer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim imageFileName As String
        Dim posn As Integer
        Dim path As String
        Dim fileName As String
        Dim thumbFileName As String
        Dim thumbImage As System.Drawing.Image
        Dim numImages As Integer
        Dim thumb As Thumbnail
        Dim thisItem As ListViewItem

        Try
            ThumbnailViewTitleBar.Text = thumbnailViewTitle
            SetListViewBorderSelect()

            ' The thumbnails are rendered to 256 x 256 pixels.  Two rows of
            ' three fit nicely on the 1024 x 768 screen.  If there are more
            ' than 6 thumnails, dynamically downsize them to fit three rows
            ' of four.
            numImages = ImageList.Count
            If numImages > 6 Then
                ThumbnailImageList.ImageSize = New System.Drawing.Size(185, 185)
            End If

            ' Iterate through the dataset to load the thumbnails into the imagelist
            Dim i As Int16 = 0
            For Each entry As CompetitionEntry In ImageList

                ' Compute the file name of the thumbnail image
                imageFileName = entry.Image_File_Name
                posn = InStrRev(imageFileName, "\")
                If posn = 0 Then
                    path = "."
                Else
                    path = Mid(imageFileName, 1, posn - 1)
                End If
                fileName = Mid(imageFileName, posn + 1)
                thumbFileName = path + "\Thumbnails\" + fileName
                ' If it's a relative path, convert to an absolute path
                If Not InStr(1, thumbFileName, ":\") = 2 Then
                    thumbFileName = theMainForm.imagesRootFolder + "\" + thumbFileName
                End If
                ' If the thumbnail file doesn't exist, render it now
                If Not File.Exists(thumbFileName) Then
                    thumb = New Thumbnail(theMainForm)
                    thumb.imageFile = imageFileName
                    thumb.Render()
                End If

                ' Store the thumbnail image in the ImageList
                thumbImage = Image.FromFile(thumbFileName)
                ThumbnailImageList.Images.Add(thumbImage)

                ' Add the image to the ListView
                thisItem = ThumbnailListView.Items.Add(entry.Award, i)
                thisItem.Tag = i    ' Store the corresponding run number of the dataset in the listview item
                i = i + 1
            Next

            ' Make sure the first item is selected
            If ThumbnailListView.Items.Count > 0 Then
                ThumbnailListView.Items(0).Selected = True
            End If

            Zoomed = False

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in ThumbnailViewer_Load()")
        Finally
            thumb = Nothing
        End Try
    End Sub

    Private Sub SetListViewBorderSelect()
        Dim styles As Integer

        Try
            styles = SendMessage(ThumbnailListView.Handle, LVM.LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
            styles = styles Or LVS_EX.LVS_EX_BORDERSELECT
            SendMessage(ThumbnailListView.Handle, LVM.LVM_SETEXTENDEDLISTVIEWSTYLE, 0, styles)
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in SetListViewBorderStyle()")
        End Try
    End Sub

    Friend Sub setSizes()
        Dim I As RpsImageSize = New RpsImageSize
        ClientSize = New System.Drawing.Size(I.getFullWidth(), I.getFullHeight())
        MaximumSize = New Size(I.getFullWidth(), I.getFullHeight())
        MinimumSize = New Size(I.getFullWidth(), I.getFullHeight())

        ThumbnailListView.Location = New Point(0, 43)
        ThumbnailListView.Size = New Size(I.getFullWidth(), I.getFullHeight() - 37 - 2)

        ThumbnailImageList.ImageSize = New Size(256, 256)

        ZoomedImage.Location = New Point(0, 0)
        ZoomedImage.Size = New Size(I.getFullWidth(), I.getFullHeight())

        ThumbnailViewTitleBar.Location = New Point(0, 0)
        ThumbnailViewTitleBar.Size = New Size(I.getFullWidth(), 37)

        Divider.Location = New Point(0, 37)
        Divider.Size = New Size(I.getFullWidth(), 2)


    End Sub

End Class
