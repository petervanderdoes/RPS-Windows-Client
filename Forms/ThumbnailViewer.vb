Namespace Forms
    Public Class ThumbnailViewer
        Inherits Form

        Private ReadOnly the_main_form As MainForm
        Private ReadOnly image_list As IList
        Private current_index As Integer
        Private current_row As Entities.CompetitionEntry
        Private full_size_file_name As String
        Private is_zoomed As Boolean
        Private ReadOnly thumbnail_view_title As String

        ' API parameters for setting border select style for the listview
        Private Const LVM_FIRST As Integer = &H1000
        Public Const LVM_GETCOUNTPERPAGE As Integer = LVM_FIRST + 40
        Public Const WM_SETREDRAW As Integer = &HB

        Private Enum LVS_EX
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

        Private Enum LVM
            LVM_FIRST = &H1000
            LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
            LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
        End Enum 'LVM

#Region " Windows Form Designer generated code "

        Public Sub New(ByVal myMainForm As MainForm,
                       ByVal ds As IList,
                       ByVal screenTitle As String,
                       ByVal image_width As Integer,
                       ByVal image_height As Integer)
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call
            the_main_form = myMainForm
            image_list = ds
            thumbnail_view_title = screenTitle
            SetSizes(image_width, image_height)
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
        Friend WithEvents ZoomedImage As PictureBox
        Friend WithEvents ThumbnailViewTitleBar As Label
        Friend WithEvents Divider As Label

        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Dim resources As System.ComponentModel.ComponentResourceManager =
                    New System.ComponentModel.ComponentResourceManager(GetType(ThumbnailViewer))
            Me.ThumbnailListView = New System.Windows.Forms.ListView()
            Me.ThumbnailImageList = New System.Windows.Forms.ImageList(Me.components)
            Me.ZoomedImage = New PictureBox()
            Me.ThumbnailViewTitleBar = New Label()
            Me.Divider = New Label()
            CType(Me.ZoomedImage, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'ThumbnailListView
            '
            Me.ThumbnailListView.BackColor = Color.Black
            Me.ThumbnailListView.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.ThumbnailListView.Font = New System.Drawing.Font("Segoe UI",
                                                                18.0!,
                                                                System.Drawing.FontStyle.Bold,
                                                                System.Drawing.GraphicsUnit.Point,
                                                                CType(0, Byte))
            Me.ThumbnailListView.ForeColor = Color.White
            Me.ThumbnailListView.LargeImageList = Me.ThumbnailImageList
            Me.ThumbnailListView.Location = New Point(0, 43)
            Me.ThumbnailListView.Margin = New System.Windows.Forms.Padding(0)
            Me.ThumbnailListView.MultiSelect = False
            Me.ThumbnailListView.Name = "ThumbnailListView"
            Me.ThumbnailListView.Size = New Size(1024, 725)
            Me.ThumbnailListView.TabIndex = 0
            Me.ThumbnailListView.UseCompatibleStateImageBehavior = False
            '
            'ThumbnailImageList
            '
            Me.ThumbnailImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit
            Me.ThumbnailImageList.ImageSize = New Size(256, 256)
            Me.ThumbnailImageList.TransparentColor = Color.Transparent
            '
            'ZoomedImage
            '
            Me.ZoomedImage.BackColor = Color.Black
            Me.ZoomedImage.Location = New Point(0, 0)
            Me.ZoomedImage.Margin = New System.Windows.Forms.Padding(0)
            Me.ZoomedImage.Name = "ZoomedImage"
            Me.ZoomedImage.Size = New Size(1024, 768)
            Me.ZoomedImage.SizeMode = PictureBoxSizeMode.CenterImage
            Me.ZoomedImage.TabIndex = 1
            Me.ZoomedImage.TabStop = False
            Me.ZoomedImage.Visible = False
            '
            'ThumbnailViewTitleBar
            '
            Me.ThumbnailViewTitleBar.BackColor = Color.Black
            Me.ThumbnailViewTitleBar.Font = New System.Drawing.Font("Segoe UI",
                                                                    18.0!,
                                                                    System.Drawing.FontStyle.Bold,
                                                                    System.Drawing.GraphicsUnit.Point,
                                                                    CType(0, Byte))
            Me.ThumbnailViewTitleBar.ForeColor = Color.Yellow
            Me.ThumbnailViewTitleBar.Location = New Point(0, 0)
            Me.ThumbnailViewTitleBar.Margin = New System.Windows.Forms.Padding(0)
            Me.ThumbnailViewTitleBar.Name = "ThumbnailViewTitleBar"
            Me.ThumbnailViewTitleBar.Size = New Size(1024, 37)
            Me.ThumbnailViewTitleBar.TabIndex = 2
            Me.ThumbnailViewTitleBar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
            '
            'Divider
            '
            Me.Divider.BackColor = Color.White
            Me.Divider.Location = New Point(0, 37)
            Me.Divider.Margin = New System.Windows.Forms.Padding(0)
            Me.Divider.Name = "Divider"
            Me.Divider.Size = New Size(1024, 2)
            Me.Divider.TabIndex = 3
            '
            'ThumbnailViewer
            '
            Me.BackColor = Color.Black
            Me.ClientSize = New Size(1024, 768)
            Me.ControlBox = False
            Me.Controls.Add(Me.Divider)
            Me.Controls.Add(Me.ThumbnailViewTitleBar)
            Me.Controls.Add(Me.ZoomedImage)
            Me.Controls.Add(Me.ThumbnailListView)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.MaximizeBox = False
            Me.MaximumSize = New Size(1024, 768)
            Me.MinimizeBox = False
            Me.MinimumSize = New Size(1024, 768)
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
                           the_main_form.awards.Count >= 1 Then
                            ThumbnailListView.SelectedItems(0).Text = the_main_form.awards.Item(0)
                            current_index = ThumbnailListView.SelectedIndices(0)
                            dataset_row_num = ThumbnailListView.Items(current_index).Tag
                            current_row = image_list.Item(dataset_row_num)
                            current_row.Award = the_main_form.awards.Item(0)
                        End If
                    Case Keys.F2
                        If ThumbnailListView.SelectedIndices.Count > 0 And
                           the_main_form.awards.Count >= 2 Then
                            ThumbnailListView.SelectedItems(0).Text = the_main_form.awards.Item(1)
                            current_index = ThumbnailListView.SelectedIndices(0)
                            dataset_row_num = ThumbnailListView.Items(current_index).Tag
                            current_row = image_list.Item(dataset_row_num)
                            current_row.Award = the_main_form.awards.Item(1)
                        End If
                    Case Keys.F3
                        If ThumbnailListView.SelectedIndices.Count > 0 And
                           the_main_form.awards.Count >= 3 Then
                            ThumbnailListView.SelectedItems(0).Text = the_main_form.awards.Item(2)
                            current_index = ThumbnailListView.SelectedIndices(0)
                            dataset_row_num = ThumbnailListView.Items(current_index).Tag
                            current_row = image_list.Item(dataset_row_num)
                            current_row.Award = the_main_form.awards.Item(2)
                        End If
                    Case Keys.F4
                        If ThumbnailListView.SelectedIndices.Count > 0 And
                           the_main_form.awards.Count >= 4 Then
                            ThumbnailListView.SelectedItems(0).Text = the_main_form.awards.Item(3)
                            current_index = ThumbnailListView.SelectedIndices(0)
                            dataset_row_num = ThumbnailListView.Items(current_index).Tag
                            current_row = image_list.Item(dataset_row_num)
                            current_row.Award = the_main_form.awards.Item(3)
                        End If
                    Case Keys.Delete, Keys.Back
                        If ThumbnailListView.SelectedIndices.Count > 0 Then
                            ThumbnailListView.SelectedItems(0).Text = ""
                            current_index = ThumbnailListView.SelectedIndices(0)
                            dataset_row_num = ThumbnailListView.Items(current_index).Tag
                            current_row = image_list.Item(dataset_row_num)
                            current_row.Award = Nothing
                        End If
                    Case Keys.R
                        If ThumbnailListView.SelectedIndices.Count > 0 Then
                            ' Remove the currently selected item from the ListView
                            current_index = ThumbnailListView.SelectedIndices(0)
                            ThumbnailListView.Items.RemoveAt(current_index)
                            ' Make sure a thumbnail is still selected
                            If ThumbnailListView.Items.Count > 0 Then
                                new_index = Math.Min(current_index, ThumbnailListView.Items.Count - 1)
                                ThumbnailListView.Items(new_index).Selected = True
                            End If

                        End If
                    Case Keys.Z
                        If ThumbnailListView.SelectedIndices.Count > 0 Then
                            If Not is_zoomed Then
                                ' Retrieve the file name of the currently selected image from the
                                ' database
                                current_index = ThumbnailListView.SelectedIndices(0)
                                dataset_row_num = ThumbnailListView.Items(current_index).Tag
                                current_row = image_list.Item(dataset_row_num)
                                full_size_file_name = current_row.Image_File_Name
                                ' If it's a relative path, convert to an absolute path
                                If Not InStr(1, full_size_file_name, ":\") = 2 Then
                                    full_size_file_name = the_main_form.images_root_folder + "\" + full_size_file_name
                                End If
                                ' Load it into the PictureBox
                                ZoomedImage.Image = Drawing.Image.FromFile(full_size_file_name)
                                ' Hide the ListView and reveal the PictureBox
                                ThumbnailListView.Visible = False
                                ThumbnailViewTitleBar.Visible = False
                                Divider.Visible = False
                                ZoomedImage.Visible = True
                                is_zoomed = True
                            Else
                                ThumbnailListView.Visible = True
                                ThumbnailViewTitleBar.Visible = True
                                Divider.Visible = True
                                ZoomedImage.Visible = False
                                is_zoomed = False
                            End If
                        End If
                        e.Handled = True
                End Select
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub ThumbnailViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            Dim image_file_name As String
            Dim posn As Integer
            Dim path As String
            Dim file_name As String
            Dim thumb_file_name As String
            Dim thumb_image As Drawing.Image
            Dim num_images As Integer
            Dim thumb As Image.Thumbnail
            Dim this_item As ListViewItem

            Try
                ThumbnailViewTitleBar.Text = thumbnail_view_title
                setListViewBorderSelect()

                ' The thumbnails are rendered to 256 x 256 pixels.  Two rows of
                ' three fit nicely on the 1024 x 768 screen.  If there are more
                ' than 6 thumnails, dynamically downsize them to fit three rows
                ' of four.
                num_images = image_list.Count
                If num_images > 6 Then
                    ThumbnailImageList.ImageSize = New Size(185, 185)
                End If

                ' Iterate through the dataset to load the thumbnails into the imagelist
                Dim i As Int16 = 0
                For Each entry As Entities.CompetitionEntry In image_list

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
                        thumb_file_name = the_main_form.images_root_folder + "\" + thumb_file_name
                    End If
                    ' If the thumbnail file doesn't exist, render it now
                    If Not File.Exists(thumb_file_name) Then
                        thumb = New Image.Thumbnail(the_main_form)
                        thumb.imageFile = image_file_name
                        thumb.doRender()
                    End If

                    ' Store the thumbnail image in the ImageList
                    thumb_image = Drawing.Image.FromFile(thumb_file_name)
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

                is_zoomed = False

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
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
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub SetSizes(ByVal image_width As Integer, ByVal image_height As Integer)
            Dim I As Image.RpsImageSize = New Image.RpsImageSize
            I.ImageWidth = image_width
            I.ImageHeight = image_height
            ClientSize = New Size(I.ImageWidth(), I.ImageHeight())
            MaximumSize = New Size(I.ImageWidth(), I.ImageHeight())
            MinimumSize = New Size(I.ImageWidth(), I.ImageHeight())

            ThumbnailListView.Location = New Point(0, 43)
            ThumbnailListView.Size = New Size(I.ImageWidth(), I.ImageHeight() - 37 - 2)

            ThumbnailImageList.ImageSize = New Size(256, 256)

            ZoomedImage.Location = New Point(0, 0)
            ZoomedImage.Size = New Size(I.ImageWidth(), I.ImageHeight())

            ThumbnailViewTitleBar.Location = New Point(0, 0)
            ThumbnailViewTitleBar.Size = New Size(I.ImageWidth(), 37)

            Divider.Location = New Point(0, 37)
            Divider.Size = New Size(I.ImageWidth(), 2)
        End Sub
    End Class
End Namespace
