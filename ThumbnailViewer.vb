Imports System.Reflection

Public Class ThumbnailViewer
    Inherits Form

    Private ReadOnly _the_main_form As MainForm
    Private ReadOnly _image_list As IList
    Private _current_index As Integer
    Private _current_row As CompetitionEntry
    Private _full_size_file_name As String
    Private _is_zoomed As Boolean
    Private ReadOnly _thumbnail_view_title As String

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

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal myMainForm As MainForm, ByVal ds As IList, ByVal screenTitle As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _the_main_form = myMainForm
        _image_list = ds
        _thumbnail_view_title = screenTitle
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

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
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
        Me.ThumbnailListView.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.ThumbnailViewTitleBar.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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

    Private Sub ThumbnailListView_KeyDown(sender As Object, e As KeyEventArgs) Handles ThumbnailListView.KeyDown
        Dim dataset_row_num As Integer
        Dim new_index As Integer

        Try
            Select Case e.KeyCode
                Case Keys.Escape, Keys.X
                    Close()
                Case Keys.F1
                    If ThumbnailListView.SelectedIndices.Count > 0 And
                       _the_main_form.awards.Count >= 1 Then
                        ThumbnailListView.SelectedItems(0).Text = _the_main_form.awards.Item(0)
                        _current_index = ThumbnailListView.SelectedIndices(0)
                        dataset_row_num = ThumbnailListView.Items(_current_index).Tag
                        _current_row = _image_list.Item(dataset_row_num)
                        _current_row.Award = _the_main_form.awards.Item(0)
                    End If
                Case Keys.F2
                    If ThumbnailListView.SelectedIndices.Count > 0 And
                       _the_main_form.awards.Count >= 2 Then
                        ThumbnailListView.SelectedItems(0).Text = _the_main_form.awards.Item(1)
                        _current_index = ThumbnailListView.SelectedIndices(0)
                        dataset_row_num = ThumbnailListView.Items(_current_index).Tag
                        _current_row = _image_list.Item(dataset_row_num)
                        _current_row.Award = _the_main_form.awards.Item(1)
                    End If
                Case Keys.F3
                    If ThumbnailListView.SelectedIndices.Count > 0 And
                       _the_main_form.awards.Count >= 3 Then
                        ThumbnailListView.SelectedItems(0).Text = _the_main_form.awards.Item(2)
                        _current_index = ThumbnailListView.SelectedIndices(0)
                        dataset_row_num = ThumbnailListView.Items(_current_index).Tag
                        _current_row = _image_list.Item(dataset_row_num)
                        _current_row.Award = _the_main_form.awards.Item(2)
                    End If
                Case Keys.F4
                    If ThumbnailListView.SelectedIndices.Count > 0 And
                       _the_main_form.awards.Count >= 4 Then
                        ThumbnailListView.SelectedItems(0).Text = _the_main_form.awards.Item(3)
                        _current_index = ThumbnailListView.SelectedIndices(0)
                        dataset_row_num = ThumbnailListView.Items(_current_index).Tag
                        _current_row = _image_list.Item(dataset_row_num)
                        _current_row.Award = _the_main_form.awards.Item(3)
                    End If
                Case Keys.Delete, Keys.Back
                    If ThumbnailListView.SelectedIndices.Count > 0 Then
                        ThumbnailListView.SelectedItems(0).Text = ""
                        _current_index = ThumbnailListView.SelectedIndices(0)
                        dataset_row_num = ThumbnailListView.Items(_current_index).Tag
                        _current_row = _image_list.Item(dataset_row_num)
                        _current_row.Award = Nothing
                    End If
                Case Keys.R
                    If ThumbnailListView.SelectedIndices.Count > 0 Then
                        ' Remove the currently selected item from the ListView
                        _current_index = ThumbnailListView.SelectedIndices(0)
                        ThumbnailListView.Items.RemoveAt(_current_index)
                        ' Make sure a thumbnail is still selected
                        If ThumbnailListView.Items.Count > 0 Then
                            new_index = Math.Min(_current_index, ThumbnailListView.Items.Count - 1)
                            ThumbnailListView.Items(new_index).Selected = True
                        End If

                    End If
                Case Keys.Z
                    If ThumbnailListView.SelectedIndices.Count > 0 Then
                        If Not _is_zoomed Then
                            ' Retrieve the file name of the currently selected image from the
                            ' database
                            _current_index = ThumbnailListView.SelectedIndices(0)
                            dataset_row_num = ThumbnailListView.Items(_current_index).Tag
                            _current_row = _image_list.Item(dataset_row_num)
                            _full_size_file_name = _current_row.Image_File_Name
                            ' If it's a relative path, convert to an absolute path
                            If Not InStr(1, _full_size_file_name, ":\") = 2 Then
                                _full_size_file_name = _the_main_form.images_root_folder + "\" + _full_size_file_name
                            End If
                            ' Load it into the PictureBox
                            ZoomedImage.Image = Image.FromFile(_full_size_file_name)
                            ' Hide the ListView and reveal the PictureBox
                            ThumbnailListView.Visible = False
                            ThumbnailViewTitleBar.Visible = False
                            Divider.Visible = False
                            ZoomedImage.Visible = True
                            _is_zoomed = True
                        Else
                            ThumbnailListView.Visible = True
                            ThumbnailViewTitleBar.Visible = True
                            Divider.Visible = True
                            ZoomedImage.Visible = False
                            _is_zoomed = False
                        End If
                    End If
                    e.Handled = True
            End Select
        Catch exception As Exception
            MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
        End Try
    End Sub

    Private Sub ThumbnailViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim image_file_name As String
        Dim posn As Integer
        Dim path As String
        Dim file_name As String
        Dim thumb_file_name As String
        Dim thumb_image As Image
        Dim num_images As Integer
        Dim thumb As Thumbnail
        Dim this_item As ListViewItem

        Try
            ThumbnailViewTitleBar.Text = _thumbnail_view_title
            setListViewBorderSelect()

            ' The thumbnails are rendered to 256 x 256 pixels.  Two rows of
            ' three fit nicely on the 1024 x 768 screen.  If there are more
            ' than 6 thumnails, dynamically downsize them to fit three rows
            ' of four.
            num_images = _image_list.Count
            If num_images > 6 Then
                ThumbnailImageList.ImageSize = New Size(185, 185)
            End If

            ' Iterate through the dataset to load the thumbnails into the imagelist
            Dim i As Int16 = 0
            For Each entry As CompetitionEntry In _image_list

                ' Compute the file name of the thumbnail image
                image_file_name = entry.Image_File_Name
                posn = InStrRev(image_file_name, "\")
                If posn = 0 Then
                    path = "."
                Else
                    path = Mid(image_file_name, 1, posn - 1)
                End If
                file_name = Mid(image_file_name, posn + 1)
                thumb_file_name = path + "\Thumbnails\" + file_name
                ' If it's a relative path, convert to an absolute path
                If Not InStr(1, thumb_file_name, ":\") = 2 Then
                    thumb_file_name = _the_main_form.images_root_folder + "\" + thumb_file_name
                End If
                ' If the thumbnail file doesn't exist, render it now
                If Not File.Exists(thumb_file_name) Then
                    thumb = New Thumbnail(_the_main_form)
                    thumb.imageFile = image_file_name
                    thumb.doRender()
                End If

                ' Store the thumbnail image in the ImageList
                thumb_image = Image.FromFile(thumb_file_name)
                ThumbnailImageList.Images.Add(thumb_image)

                ' Add the image to the ListView
                this_item = ThumbnailListView.Items.Add(entry.Award, i)
                this_item.Tag = i    ' Store the corresponding run number of the dataset in the listview item
                i = i + 1
            Next

            ' Make sure the first item is selected
            If ThumbnailListView.Items.Count > 0 Then
                ThumbnailListView.Items(0).Selected = True
            End If

            _is_zoomed = False

        Catch exception As Exception
            MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
        Finally
        End Try
    End Sub

    Private Sub setListViewBorderSelect()
        Dim styles As Integer

        Try
            styles = NativeMethods.SendMessage(ThumbnailListView.Handle, LVM.LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
            styles = styles Or LVS_EX.LVS_EX_BORDERSELECT
            NativeMethods.SendMessage(ThumbnailListView.Handle, LVM.LVM_SETEXTENDEDLISTVIEWSTYLE, 0, styles)
        Catch exception As Exception
            MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
        End Try
    End Sub

    Friend Sub setSizes()
        Dim I As RpsImageSize = New RpsImageSize
        ClientSize = New Size(I.getFullWidth(), I.getFullHeight())
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
