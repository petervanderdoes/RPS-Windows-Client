Namespace Forms
    Public Class ImageViewer
        Inherits Form

        ReadOnly the_main_form As MainForm
        ReadOnly image_list As IList
        Private num_images As Integer
        Private current_index As Integer
        Private current_file_name As String
        Private current_row As Entities.CompetitionEntry
        Private current_score_str As String
        Private status_bar_visible As Integer = 0
        Private splash_screen_visible As Boolean = True
        Private in_delay_loop As Boolean = False
        Private thumb As Image.Thumbnail
        Private thumbnail_thread As Thread

        Private Const SCORE As Integer = 0
        Private Const AWARD As Integer = 1

#Region " Windows Form Designer generated code "

        Public Sub New(ByVal myMainForm As MainForm,
                       ByVal ds As IList,
                       ByVal idx As Integer,
                       ByVal splash As Boolean,
                       ByVal statusBarState As Integer,
                       ByVal image_width As Integer,
                       ByVal image_height As Integer)
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call
            the_main_form = myMainForm
            image_list = ds
            current_index = idx
            splash_screen_visible = splash
            status_bar_visible = statusBarState
            SetSizes(image_width, image_height)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.
        'Do not modify it using the code editor.
        Friend WithEvents picShowPicture As PictureBox
        Friend WithEvents ofdSelectPicture As OpenFileDialog
        Friend WithEvents ScorePopUp As Label
        Friend WithEvents StatusBar As Panel
        Friend WithEvents StatusBarTitle As Label
        Friend WithEvents StatusBarScore As Label
        Friend WithEvents StatusBarAward As Label
        Friend WithEvents splashClub As Label
        Friend WithEvents splashTheme As Label
        Friend WithEvents splashClassification As Label

        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.picShowPicture = New System.Windows.Forms.PictureBox()
            Me.ofdSelectPicture = New System.Windows.Forms.OpenFileDialog()
            Me.ScorePopUp = New System.Windows.Forms.Label()
            Me.StatusBar = New System.Windows.Forms.Panel()
            Me.StatusBarAward = New System.Windows.Forms.Label()
            Me.StatusBarScore = New System.Windows.Forms.Label()
            Me.StatusBarTitle = New System.Windows.Forms.Label()
            Me.splashClub = New System.Windows.Forms.Label()
            Me.splashTheme = New System.Windows.Forms.Label()
            Me.splashClassification = New System.Windows.Forms.Label()
            CType(Me.picShowPicture, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.StatusBar.SuspendLayout()
            Me.SuspendLayout()
            '
            'picShowPicture
            '
            Me.picShowPicture.BackColor = System.Drawing.Color.Black
            Me.picShowPicture.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
            Me.picShowPicture.Location = New System.Drawing.Point(0, 0)
            Me.picShowPicture.Margin = New System.Windows.Forms.Padding(0)
            Me.picShowPicture.Name = "picShowPicture"
            Me.picShowPicture.Size = New System.Drawing.Size(1024, 768)
            Me.picShowPicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
            Me.picShowPicture.TabIndex = 2
            Me.picShowPicture.TabStop = False
            '
            'ofdSelectPicture
            '
            Me.ofdSelectPicture.DefaultExt = "jpg"
            Me.ofdSelectPicture.Filter = "Windows Bitmaps|*.BMP|JPEG Files|*.JPG"
            Me.ofdSelectPicture.FilterIndex = 2
            Me.ofdSelectPicture.Title = "Select Picture"
            '
            'ScorePopUp
            '
            Me.ScorePopUp.Anchor =
                CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right),
                      System.Windows.Forms.AnchorStyles)
            Me.ScorePopUp.BackColor = System.Drawing.Color.Black
            Me.ScorePopUp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.ScorePopUp.Font = New System.Drawing.Font("Segoe UI",
                                                         56.25!,
                                                         System.Drawing.FontStyle.Bold,
                                                         System.Drawing.GraphicsUnit.Point,
                                                         CType(0, Byte))
            Me.ScorePopUp.ForeColor = System.Drawing.Color.White
            Me.ScorePopUp.Location = New System.Drawing.Point(803, 29)
            Me.ScorePopUp.Margin = New System.Windows.Forms.Padding(0)
            Me.ScorePopUp.Name = "ScorePopUp"
            Me.ScorePopUp.Size = New System.Drawing.Size(192, 115)
            Me.ScorePopUp.TabIndex = 3
            Me.ScorePopUp.Text = "3rd"
            Me.ScorePopUp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            Me.ScorePopUp.Visible = False
            '
            'StatusBar
            '
            Me.StatusBar.BackColor = System.Drawing.Color.Black
            Me.StatusBar.Controls.Add(Me.StatusBarAward)
            Me.StatusBar.Controls.Add(Me.StatusBarScore)
            Me.StatusBar.Controls.Add(Me.StatusBarTitle)
            Me.StatusBar.Location = New System.Drawing.Point(0, 750)
            Me.StatusBar.Margin = New System.Windows.Forms.Padding(0)
            Me.StatusBar.Name = "StatusBar"
            Me.StatusBar.Size = New System.Drawing.Size(1024, 28)
            Me.StatusBar.TabIndex = 4
            Me.StatusBar.Visible = False
            '
            'StatusBarAward
            '
            Me.StatusBarAward.BackColor = System.Drawing.Color.Black
            Me.StatusBarAward.Font = New System.Drawing.Font("Verdana",
                                                             12.0!,
                                                             System.Drawing.FontStyle.Bold,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(0, Byte))
            Me.StatusBarAward.ForeColor = System.Drawing.Color.White
            Me.StatusBarAward.Location = New System.Drawing.Point(0, 0)
            Me.StatusBarAward.Margin = New System.Windows.Forms.Padding(0)
            Me.StatusBarAward.Name = "StatusBarAward"
            Me.StatusBarAward.Size = New System.Drawing.Size(134, 28)
            Me.StatusBarAward.TabIndex = 2
            Me.StatusBarAward.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'StatusBarScore
            '
            Me.StatusBarScore.BackColor = System.Drawing.Color.Black
            Me.StatusBarScore.Font = New System.Drawing.Font("Verdana",
                                                             12.0!,
                                                             System.Drawing.FontStyle.Bold,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(0, Byte))
            Me.StatusBarScore.ForeColor = System.Drawing.Color.White
            Me.StatusBarScore.Location = New System.Drawing.Point(889, 0)
            Me.StatusBarScore.Margin = New System.Windows.Forms.Padding(0)
            Me.StatusBarScore.Name = "StatusBarScore"
            Me.StatusBarScore.Size = New System.Drawing.Size(135, 28)
            Me.StatusBarScore.TabIndex = 1
            Me.StatusBarScore.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'StatusBarTitle
            '
            Me.StatusBarTitle.BackColor = System.Drawing.Color.Black
            Me.StatusBarTitle.Font = New System.Drawing.Font("Verdana",
                                                             12.0!,
                                                             System.Drawing.FontStyle.Bold,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(0, Byte))
            Me.StatusBarTitle.ForeColor = System.Drawing.Color.White
            Me.StatusBarTitle.Location = New System.Drawing.Point(134, 0)
            Me.StatusBarTitle.Margin = New System.Windows.Forms.Padding(0)
            Me.StatusBarTitle.Name = "StatusBarTitle"
            Me.StatusBarTitle.Size = New System.Drawing.Size(755, 28)
            Me.StatusBarTitle.TabIndex = 0
            Me.StatusBarTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            Me.StatusBarTitle.UseMnemonic = False
            '
            'splashClub
            '
            Me.splashClub.BackColor = System.Drawing.Color.Black
            Me.splashClub.Font = New System.Drawing.Font("Segoe UI",
                                                         39.75!,
                                                         System.Drawing.FontStyle.Bold,
                                                         System.Drawing.GraphicsUnit.Point,
                                                         CType(0, Byte))
            Me.splashClub.ForeColor = System.Drawing.Color.Yellow
            Me.splashClub.Location = New System.Drawing.Point(0, 200)
            Me.splashClub.Margin = New System.Windows.Forms.Padding(0)
            Me.splashClub.Name = "splashClub"
            Me.splashClub.Size = New System.Drawing.Size(1024, 101)
            Me.splashClub.TabIndex = 5
            Me.splashClub.Text = "Club Name"
            Me.splashClub.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            Me.splashClub.UseMnemonic = False
            '
            'splashTheme
            '
            Me.splashTheme.BackColor = System.Drawing.Color.Black
            Me.splashTheme.Font = New System.Drawing.Font("Segoe UI",
                                                          36.0!,
                                                          System.Drawing.FontStyle.Bold,
                                                          System.Drawing.GraphicsUnit.Point,
                                                          CType(0, Byte))
            Me.splashTheme.ForeColor = System.Drawing.Color.White
            Me.splashTheme.Location = New System.Drawing.Point(0, 334)
            Me.splashTheme.Margin = New System.Windows.Forms.Padding(0)
            Me.splashTheme.Name = "splashTheme"
            Me.splashTheme.Size = New System.Drawing.Size(1024, 101)
            Me.splashTheme.TabIndex = 6
            Me.splashTheme.Text = "Theme Name"
            Me.splashTheme.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            Me.splashTheme.UseMnemonic = False
            '
            'splashClassification
            '
            Me.splashClassification.BackColor = System.Drawing.Color.Black
            Me.splashClassification.Font = New System.Drawing.Font("Segoe UI",
                                                                   48.0!,
                                                                   System.Drawing.FontStyle.Bold,
                                                                   System.Drawing.GraphicsUnit.Point,
                                                                   CType(0, Byte))
            Me.splashClassification.ForeColor = System.Drawing.Color.Red
            Me.splashClassification.Location = New System.Drawing.Point(0, 468)
            Me.splashClassification.Margin = New System.Windows.Forms.Padding(0)
            Me.splashClassification.Name = "splashClassification"
            Me.splashClassification.Size = New System.Drawing.Size(1024, 101)
            Me.splashClassification.TabIndex = 7
            Me.splashClassification.Text = "Classification"
            Me.splashClassification.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            Me.splashClassification.UseMnemonic = False
            '
            'ImageViewer
            '
            Me.BackColor = System.Drawing.Color.Maroon
            Me.ClientSize = New System.Drawing.Size(1024, 768)
            Me.ControlBox = False
            Me.Controls.Add(Me.ScorePopUp)
            Me.Controls.Add(Me.splashClassification)
            Me.Controls.Add(Me.splashTheme)
            Me.Controls.Add(Me.splashClub)
            Me.Controls.Add(Me.StatusBar)
            Me.Controls.Add(Me.picShowPicture)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Me.MaximizeBox = False
            Me.MaximumSize = New System.Drawing.Size(1024, 768)
            Me.MinimizeBox = False
            Me.MinimumSize = New System.Drawing.Size(1024, 768)
            Me.Name = "ImageViewer"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
            Me.Text = "Picture Viewer"
            Me.TopMost = True
            Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
            CType(Me.picShowPicture, System.ComponentModel.ISupportInitialize).EndInit()
            Me.StatusBar.ResumeLayout(False)
            Me.ResumeLayout(False)
        End Sub

#End Region

        Private Sub fclsViewer_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown

            Try
                ' Don't accept any more keystrokes until the rendering thread is complete
                If Not thumbnail_thread Is Nothing Then
                    While thumbnail_thread.IsAlive
                        Application.DoEvents()
                    End While
                End If

                ' If a key was pressed while in the middle of a delay loop (e.g. while
                ' the status bar is visible), terminate the delay loop before processing
                ' the key press
                If in_delay_loop Then
                    in_delay_loop = False
                    Application.DoEvents()
                End If

                ' Examine which key was pressed and act accordingly
                Select Case e.KeyCode
                    Case Keys.Escape ' Esc to exit
                        thumb = Nothing
                        Close()
                    Case Keys.PageDown, Keys.Down, Keys.Right ' PgDn, Down Arrow, Right Arrow = move to next image
                        If splash_screen_visible Then
                            hideSplashScreen()
                            moveToImage(current_index)
                        Else
                            ScorePopUp.Visible = False
                            moveToImage(current_index + 1)
                        End If
                    Case Keys.PageUp, Keys.Up, Keys.Left ' PgUp, Up Arrow, Left Arrow = move to previous image
                        ScorePopUp.Visible = False
                        moveToImage(current_index - 1)
                    Case Keys.Home ' Home = Move to first image
                        ScorePopUp.Visible = False
                        moveToImage(0)
                    Case Keys.End ' End = Move to last image
                        ScorePopUp.Visible = False
                        moveToImage(num_images - 1)
                    Case Keys.Delete ' Del key = clear score and award
                        current_row.Score_1 = Nothing
                        current_row.Award = Nothing
                        StatusBarScore.Text = ""
                        StatusBarAward.Text = ""
                        ScorePopUp.Visible = False
                    Case Keys.Back ' Backspace Key = Delete one character from score
                        If Len(current_score_str) > 0 Then
                            current_score_str = Mid(current_score_str, 1, Len(current_score_str) - 1) _
                            ' Remove rightmost character
                            doScorePopUp(SCORE, current_score_str, current_row)
                        End If
                    Case Keys.F1 ' F1 Key
                        If the_main_form.awards.Count >= 1 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(0), current_row)
                        End If
                    Case Keys.F2 ' F2 Key
                        If the_main_form.awards.Count >= 2 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(1), current_row)
                        End If
                    Case Keys.F3 ' F3 Key
                        If the_main_form.awards.Count >= 3 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(2), current_row)
                        End If
                    Case Keys.F4 ' F4 Key
                        If the_main_form.awards.Count >= 4 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(3), current_row)
                        End If
                    Case Keys.F5 ' F5 Key
                        If the_main_form.awards.Count >= 5 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(4), current_row)
                        End If
                    Case Keys.F6 ' F6 Key
                        If the_main_form.awards.Count >= 6 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(5), current_row)
                        End If
                    Case Keys.F7 ' F7 Key
                        If the_main_form.awards.Count >= 7 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(6), current_row)
                        End If
                    Case Keys.F8 ' F8 Key
                        If the_main_form.awards.Count >= 8 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(7), current_row)
                        End If
                    Case Keys.F9 ' F9 Key
                        If the_main_form.awards.Count >= 9 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(8), current_row)
                        End If
                    Case Keys.F10 ' F10 Key
                        If the_main_form.awards.Count >= 10 Then
                            doScorePopUp(AWARD, the_main_form.awards.Item(9), current_row)
                        End If
                    Case 48 ' 0 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "0"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 49 ' 1 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "1"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 50 ' 2 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "2"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 51 ' 3 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "3"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 52 ' 4 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "4"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 53 ' 5 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "5"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 54 ' 6 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "6"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 55 ' 7 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "7"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 56 ' 8 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "8"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case 57 ' 9 Key
                        If Len(current_score_str) < 2 Then
                            current_score_str += "9"
                        End If
                        doScorePopUp(SCORE, current_score_str, current_row)
                    Case Keys.S ' S key = toggle the status bar
                        If status_bar_visible = 0 Then
                            StatusBar.Visible = True
                            status_bar_visible = 1
                        ElseIf status_bar_visible = 1 Then
                            StatusBarTitle.Text = """" + StatusBarTitle.Text + """  by  " + current_row.Maker
                            status_bar_visible = 2
                        ElseIf status_bar_visible = 2 Then
                            StatusBarTitle.Text = current_row.Title
                            StatusBar.Visible = False
                            status_bar_visible = 0
                        End If
                    Case Keys.C ' C Key = hide the status bar and score
                        status_bar_visible = 0
                        StatusBar.Visible = False
                        ScorePopUp.Visible = False
                End Select
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub ImageViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            Try
                ' Create an instance of the Thumbnail class to render thumbnail as necessary
                thumb = New Image.Thumbnail(the_main_form)

                ' How many images are there in this slideshow?
                num_images = image_list.Count

                ' Configure the splash screen
                splashClub.Text = the_main_form.camera_club_name
                splashTheme.Text = ""
                If the_main_form.EnableTheme.CheckState = CheckState.Checked Then
                    splashTheme.Text = the_main_form.SelectTheme.Text
                End If

                splashClassification.Text = ""
                If the_main_form.EnableClassification.CheckState = CheckState.Checked Then
                    splashClassification.Text += the_main_form.SelectClassification.Text
                End If
                If the_main_form.EnableMedium.CheckState = CheckState.Checked Then
                    splashClassification.Text += " " + the_main_form.SelectMedium.Text
                End If

                ' Display the Splash Screen or the selected image
                If splash_screen_visible Then
                    showSplashScreen()
                Else
                    hideSplashScreen()
                    moveToImage(current_index)
                End If
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub doDelay(ms As Long)
            Try
                Dim i As Integer = 0
                Dim num_iterations As Integer

                num_iterations = ms / 100
                in_delay_loop = True
                While i < num_iterations And in_delay_loop
                    Thread.Sleep(100)
                    i = i + 1
                    Application.DoEvents()
                End While
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            Finally
                in_delay_loop = False
            End Try
        End Sub

        Private Sub doScorePopUp(mode As Integer, value As String, ByRef competition_entry As Entities.CompetitionEntry)
            Try
                ' When entering scores, allow the user to enter up to two digit
                If mode = SCORE Then
                    If Len(value) > 0 Then
                        competition_entry.Score_1 = CType(value, Integer)
                        StatusBarScore.Text = value + " Points"
                        If CType(value, Integer) >= (the_main_form.min_score_for_award * the_main_form.num_judges) Then
                            renderThumbnail(competition_entry.Image_File_Name)
                        End If
                    Else
                        competition_entry.Score_1 = Nothing
                        StatusBarScore.Text = ""
                    End If
                    ScorePopUp.Text = value
                    ScorePopUp.Visible = True

                    ' When entering awards, just allow one key
                Else
                    competition_entry.Award = value
                    ScorePopUp.Text = value
                    StatusBarAward.Text = value
                    ScorePopUp.Visible = True
                End If

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub doStatusBarPopUp()
            Try
                StatusBar.Visible = True
                doDelay(2000)
                StatusBar.Visible = False
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub moveToImage(index As Integer)
            Try
                If index >= 0 And index <= num_images - 1 Then
                    ' Get rid of the splash screen
                    If splash_screen_visible Then
                        hideSplashScreen()
                    End If

                    ' Get the selected image file name from the database
                    current_index = index
                    current_row = image_list(current_index)
                    current_file_name = current_row.Image_File_Name

                    ' If it's a relative path, convert to an absolute path
                    If Not InStr(1, current_file_name, ":\") = 2 Then
                        current_file_name = the_main_form.images_root_folder + "\" + current_file_name
                    End If

                    ' Prepare the Status Bar
                    StatusBarTitle.Text = current_row.Title
                    If status_bar_visible = 2 Then 'include Maker name
                        StatusBarTitle.Text = """" + StatusBarTitle.Text + """  by  " + current_row.Maker
                    End If
                    current_score_str = ""
                    StatusBarScore.Text = current_row.Score_1.ToString()
                    If StatusBarScore.Text > "" Then
                        StatusBarScore.Text = StatusBarScore.Text + " Points"
                    End If
                    If IsNothing(current_row.Award) Then
                        StatusBarAward.Text = ""
                    Else
                        StatusBarAward.Text = current_row.Award
                    End If

                    ' Show the image
                    picShowPicture.Image = Drawing.Image.FromFile(current_file_name)

                    ' Show the Status Bar if necessary
                    If status_bar_visible = 0 Then
                        doStatusBarPopUp()
                    Else
                        StatusBar.Visible = True
                    End If
                Else
                    Beep()
                End If
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub hideSplashScreen()
            splashClub.Visible = False
            splashTheme.Visible = False
            splashClassification.Visible = False
            splash_screen_visible = False
        End Sub

        Private Sub showSplashScreen()
            splashClub.Visible = True
            splashTheme.Visible = True
            splashClassification.Visible = True
            splash_screen_visible = True
        End Sub

        Private Sub renderThumbnail(file_name As String)
            ' Set the name of the image from which to create a thumbnail
            thumb.imageFile = file_name

            ' Spawn a separate thread to render the thumbnail image
            thumbnail_thread = New Thread(AddressOf thumb.doRender)
            thumbnail_thread.Start()
        End Sub

        Private Sub SetSizes(image_width As Integer, image_height As Integer)
            Dim I As Image.RpsImageSize = New Image.RpsImageSize
            Dim splash_location_y As Integer
            I.ImageWidth = image_width
            I.ImageHeight = image_height
            ClientSize = New Size(I.ImageWidth(), I.ImageHeight())
            MaximumSize = New Size(I.ImageWidth(), I.ImageHeight())
            MinimumSize = New Size(I.ImageWidth(), I.ImageHeight())

            picShowPicture.Size = New Size(I.ImageWidth(), I.ImageHeight())

            ScorePopUp.Location = New Point(I.ImageWidth() - 221, 29)
            ScorePopUp.Size = New Size(192, 115)

            StatusBar.Location = New Point(0, I.ImageHeight() - 28)
            StatusBar.Size = New Size(I.ImageWidth(), 28)

            StatusBarAward.Location = New Point(0, 0)
            StatusBarAward.Size = New Size(134, StatusBar.Size.Height)

            StatusBarScore.Location = New Point(I.ImageWidth() - 135, 0)
            StatusBarScore.Size = New Size(135, StatusBar.Size.Height)

            StatusBarTitle.Location = New Point(134, 0)
            StatusBarTitle.Size = New Size(I.ImageWidth() - StatusBarAward.Size.Width - StatusBarScore.Size.Width,
                                           StatusBar.Size.Height)

            splash_location_y = (I.ImageHeight - (3 * 110)) / 3
            splashClub.Location = New Point(0, splash_location_y)
            splashClub.Size = New Size(I.ImageWidth(), 100)

            splashTheme.Location = New Point(0, splash_location_y * 2)
            splashTheme.Size = New Size(I.ImageWidth(), 100)

            splashClassification.Location = New Point(0, splash_location_y * 3)
            splashClassification.Size = New Size(I.ImageWidth(), 100)
        End Sub
    End Class
End Namespace
