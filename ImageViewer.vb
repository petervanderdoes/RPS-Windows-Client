Imports System.Linq
Imports System.Data.Entity.Core.EntityClient
Imports System.Data.Entity.Core

Public Class ImageViewer
    Inherits System.Windows.Forms.Form

    Dim theMainForm As MainForm
    Dim ImageList As IList
    Dim numImages As Integer
    Dim currentIndex As Integer
    Dim currentFileName As String
    Dim currentRow As CompetitionEntry
    Dim currentScoreStr As String
    Dim currentScoreSubStr As String
    Dim statusBarVisible As Integer = 0
    Dim showPhotogName As Boolean = False
    Dim splashScreenVisible As Boolean = True
    Dim inDelayLoop As Boolean = False
    Dim thumb As Thumbnail
    Dim thumbnailThread As Thread

    Private Const SCORE As Integer = 0
    Private Const AWARD As Integer = 1

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal myMainForm As MainForm, ByVal ds As IList, ByVal idx As Integer, ByVal splash As Boolean, ByVal statusBarState As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        theMainForm = myMainForm
        ImageList = ds
        currentIndex = idx
        splashScreenVisible = splash
        statusBarVisible = statusBarState
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
    Friend WithEvents picShowPicture As System.Windows.Forms.PictureBox
    Friend WithEvents ofdSelectPicture As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ScorePopUp As System.Windows.Forms.Label
    Friend WithEvents StatusBar As System.Windows.Forms.Panel
    Friend WithEvents StatusBarTitle As System.Windows.Forms.Label
    Friend WithEvents StatusBarScore As System.Windows.Forms.Label
    Friend WithEvents StatusBarAward As System.Windows.Forms.Label
    Friend WithEvents splashClub As System.Windows.Forms.Label
    Friend WithEvents splashTheme As System.Windows.Forms.Label
    Friend WithEvents splashClassification As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
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
        Me.ScorePopUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ScorePopUp.BackColor = System.Drawing.Color.Black
        Me.ScorePopUp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.ScorePopUp.Font = New System.Drawing.Font("Microsoft Sans Serif", 56.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.StatusBarAward.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.StatusBarScore.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.StatusBarTitle.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.splashClub.Font = New System.Drawing.Font("Microsoft Sans Serif", 40.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.splashClub.ForeColor = System.Drawing.Color.Yellow
        Me.splashClub.Location = New System.Drawing.Point(0, 157)
        Me.splashClub.Margin = New System.Windows.Forms.Padding(0)
        Me.splashClub.Name = "splashClub"
        Me.splashClub.Size = New System.Drawing.Size(1024, 101)
        Me.splashClub.TabIndex = 5
        Me.splashClub.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.splashClub.UseMnemonic = False
        '
        'splashTheme
        '
        Me.splashTheme.BackColor = System.Drawing.Color.Black
        Me.splashTheme.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.splashTheme.ForeColor = System.Drawing.Color.White
        Me.splashTheme.Location = New System.Drawing.Point(0, 351)
        Me.splashTheme.Margin = New System.Windows.Forms.Padding(0)
        Me.splashTheme.Name = "splashTheme"
        Me.splashTheme.Size = New System.Drawing.Size(1024, 101)
        Me.splashTheme.TabIndex = 6
        Me.splashTheme.Text = "Open and Oldies"
        Me.splashTheme.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.splashTheme.UseMnemonic = False
        '
        'splashClassification
        '
        Me.splashClassification.BackColor = System.Drawing.Color.Black
        Me.splashClassification.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.splashClassification.ForeColor = System.Drawing.Color.Red
        Me.splashClassification.Location = New System.Drawing.Point(0, 535)
        Me.splashClassification.Margin = New System.Windows.Forms.Padding(0)
        Me.splashClassification.Name = "splashClassification"
        Me.splashClassification.Size = New System.Drawing.Size(1024, 102)
        Me.splashClassification.TabIndex = 7
        Me.splashClassification.Text = "Beginner"
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

    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Unload Form
        Me.Close()
    End Sub


    Private Sub fclsViewer_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        Try
            ' Don't accept any more keystrokes until the rendering thread is complete
            If Not thumbnailThread Is Nothing Then
                While thumbnailThread.IsAlive
                    Application.DoEvents()
                End While
            End If

            ' If a key was pressed while in the middle of a delay loop (e.g. while
            ' the status bar is visible), terminate the delay loop before processing
            ' the key press
            If inDelayLoop Then
                inDelayLoop = False
                Application.DoEvents()
            End If

            ' Examine which key was pressed and act accordingly
            Select Case e.KeyCode
                Case Keys.Escape    ' Esc to exit
                    thumb = Nothing
                    Me.Close()
                Case Keys.PageDown, Keys.Down, Keys.Right   ' PgDn, Down Arrow, Right Arrow = move to next image
                    If splashScreenVisible Then
                        HideSplashScreen()
                        MoveToImage(currentIndex)
                    Else
                        ScorePopUp.Visible = False
                        MoveToImage(currentIndex + 1)
                    End If
                Case Keys.PageUp, Keys.Up, Keys.Left   ' PgUp, Up Arrow, Left Arrow = move to previous image
                    ScorePopUp.Visible = False
                    MoveToImage(currentIndex - 1)
                Case Keys.Home                          ' Home = Move to first image
                    ScorePopUp.Visible = False
                    MoveToImage(0)
                Case Keys.End                           ' End = Move to last image
                    ScorePopUp.Visible = False
                    MoveToImage(numImages - 1)
                Case Keys.Delete                        ' Del key = clear score and award
                    currentRow.Score_1 = Nothing
                    currentRow.Award = Nothing
                    StatusBarScore.Text = ""
                    StatusBarAward.Text = ""
                    ScorePopUp.Visible = False
                Case Keys.Back                          ' Backspace Key = Delete one character from score
                    If Len(currentScoreStr) > 0 Then
                        currentScoreStr = Mid(currentScoreStr, 1, Len(currentScoreStr) - 1) ' Remove rightmost character
                        DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                    End If
                Case Keys.F1             ' F1 Key
                    If theMainForm.awards.Count >= 1 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(0), currentRow)
                    End If
                Case Keys.F2             ' F2 Key
                    If theMainForm.awards.Count >= 2 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(1), currentRow)
                    End If
                Case Keys.F3             ' F3 Key
                    If theMainForm.awards.Count >= 3 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(2), currentRow)
                    End If
                Case Keys.F4             ' F4 Key
                    If theMainForm.awards.Count >= 4 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(3), currentRow)
                    End If
                Case Keys.F5             ' F5 Key
                    If theMainForm.awards.Count >= 5 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(4), currentRow)
                    End If
                Case Keys.F6             ' F6 Key
                    If theMainForm.awards.Count >= 6 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(5), currentRow)
                    End If
                Case Keys.F7             ' F7 Key
                    If theMainForm.awards.Count >= 7 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(6), currentRow)
                    End If
                Case Keys.F8             ' F8 Key
                    If theMainForm.awards.Count >= 8 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(7), currentRow)
                    End If
                Case Keys.F9             ' F9 Key
                    If theMainForm.awards.Count >= 9 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(8), currentRow)
                    End If
                Case Keys.F10           ' F10 Key
                    If theMainForm.awards.Count >= 10 Then
                        DoScorePopUp_2(AWARD, theMainForm.awards.Item(9), currentRow)
                    End If
                Case 48             ' 0 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "0"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 49             ' 1 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "1"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 50             ' 2 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "2"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 51             ' 3 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "3"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 52             ' 4 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "4"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 53             ' 5 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "5"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 54             ' 6 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "6"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 55             ' 7 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "7"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 56             ' 8 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "8"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case 57             ' 9 Key
                    If Len(currentScoreStr) < 2 Then
                        currentScoreStr += "9"
                    End If
                    DoScorePopUp_2(SCORE, currentScoreStr, currentRow)
                Case Keys.S         ' S key = toggle the status bar
                    If statusBarVisible = 0 Then
                        StatusBar.Visible = True
                        statusBarVisible = 1
                    ElseIf statusBarVisible = 1 Then
                        showPhotogName = True
                        StatusBarTitle.Text = """" + StatusBarTitle.Text + """  by  " + currentRow.Maker
                        statusBarVisible = 2
                    ElseIf statusBarVisible = 2 Then
                        showPhotogName = False
                        StatusBarTitle.Text = currentRow.Title
                        StatusBar.Visible = False
                        statusBarVisible = 0
                    End If
                Case Keys.C         ' C Key = hide the status bar and score
                    statusBarVisible = 0
                    StatusBar.Visible = False
                    ScorePopUp.Visible = False
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in fclsViewer_KeyDown()")
        End Try
    End Sub

    Private Sub ImageViewer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim row As CompetitionEntry
        Try
            ' Create an instance of the Thumbnail class to render thumbnail as necessary
            thumb = New Thumbnail(theMainForm)

            ' How many images are there in this slideshow?
            numImages = ImageList.Count

            ' Grab the first row of the table to get some values
            'row = ImageList.First

            ' Configure the splash screen
            splashClub.Text = theMainForm.camera_club_name
            splashTheme.Text = ""
            If theMainForm.EnableTheme.CheckState = CheckState.Checked Then
                splashTheme.Text = theMainForm.SelectTheme.Text
            End If

            splashClassification.Text = ""
            If theMainForm.EnableClassification.CheckState = CheckState.Checked Then
                splashClassification.Text += theMainForm.SelectClassification.Text
            End If
            If theMainForm.EnableMedium.CheckState = CheckState.Checked Then
                splashClassification.Text += " " + theMainForm.SelectMedium.Text
            End If

            ' Display the Splash Screen or the selected image
            If splashScreenVisible Then
                ShowSplashScreen()
            Else
                HideSplashScreen()
                MoveToImage(currentIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in ImageViewer_Load()")
        End Try
    End Sub

    Private Sub Delay(ByVal ms As Long)
        Try
            Dim i As Integer = 0
            Dim numIterations As Integer

            numIterations = ms / 100
            inDelayLoop = True
            While i < numIterations And inDelayLoop
                System.Threading.Thread.Sleep(100)
                i = i + 1
                Application.DoEvents()
            End While
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in Delay()")
        Finally
            inDelayLoop = False
        End Try
    End Sub

    Private Sub DoScorePopUp(ByVal s As String)
        Try
            ScorePopUp.Text = s
            ScorePopUp.Visible = True
            ''Delay(800)
            ''ScorePopUp.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in DoScorePopUp()")
        End Try
    End Sub
    Private Sub DoScorePopUp_2(ByVal mode As Integer, ByVal value As String, ByRef currentRow As CompetitionEntry)
        Try
            ' When entering scores, allow the user to enter up to two digit
            If mode = SCORE Then
                If Len(value) > 0 Then
                    currentRow.Score_1 = CType(value, Integer)
                    StatusBarScore.Text = value + " Points"
                    If CType(value, Integer) >= (theMainForm.min_score_for_award * theMainForm.num_judges) Then
                        RenderThumbnail(currentRow.Image_File_Name)
                    End If
                Else
                    currentRow.Score_1 = Nothing
                    StatusBarScore.Text = ""
                End If
                ScorePopUp.Text = value
                ScorePopUp.Visible = True

                ' When entering awards, just allow one key
            Else
                currentRow.Award = value
                ScorePopUp.Text = value
                StatusBarAward.Text = value
                ScorePopUp.Visible = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in DoScorePopUp()")
        End Try
    End Sub
    Private Sub DoStatusBarPopUp()
        Try
            StatusBar.Visible = True
            Delay(2000)
            StatusBar.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in DoStatusBarPopUp()")
        End Try
    End Sub

    Private Sub MoveToImage(ByVal index As Integer)
        Try
            If index >= 0 And index <= numImages - 1 Then
                ' Get rid of the splash screen
                If splashScreenVisible Then
                    HideSplashScreen()
                End If

                ' Get the selected image file name from the database
                currentIndex = index
                currentRow = ImageList(currentIndex)
                currentFileName = currentRow.Image_File_Name

                ' If it's a relative path, convert to an absolute path
                If Not InStr(1, currentFileName, ":\") = 2 Then
                    currentFileName = theMainForm.images_root_folder + "\" + currentFileName
                End If

                ' Prepare the Status Bar
                StatusBarTitle.Text = currentRow.Title
                If statusBarVisible = 2 Then 'include Maker name
                    StatusBarTitle.Text = """" + StatusBarTitle.Text + """  by  " + currentRow.Maker
                End If
                currentScoreStr = ""
                StatusBarScore.Text = currentRow.Score_1
                If StatusBarScore.Text > "" Then
                    StatusBarScore.Text = StatusBarScore.Text + " Points"
                End If
                StatusBarAward.Text = currentRow.Award

                ' Show the image
                picShowPicture.Image = Image.FromFile(currentFileName)

                ' Show the Status Bar if necessary
                If statusBarVisible = 0 Then
                    DoStatusBarPopUp()
                Else
                    StatusBar.Visible = True
                End If
            Else
                Beep()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in MoveToImage()")
        End Try

    End Sub

    Private Sub HideSplashScreen()
        splashClub.Visible = False
        splashTheme.Visible = False
        splashClassification.Visible = False
        splashScreenVisible = False
    End Sub

    Private Sub ShowSplashScreen()
        splashClub.Visible = True
        splashTheme.Visible = True
        splashClassification.Visible = True
        splashScreenVisible = True
    End Sub

    Private Sub RenderThumbnail(ByVal fileName As String)
        ' Set the name of the image from which to create a thumbnail
        thumb.imageFile = fileName

        ' Spawn a separate thread to render the thumbnail image
        thumbnailThread = New Thread(AddressOf thumb.Render)
        thumbnailThread.Start()
    End Sub

    Friend Sub setSizes()
        Dim I As RpsImageSize = New RpsImageSize
        Dim splash_location_y As Integer
        ClientSize = New System.Drawing.Size(I.getFullWidth(), I.getFullHeight())
        MaximumSize = New Size(I.getFullWidth(), I.getFullHeight())
        MinimumSize = New Size(I.getFullWidth(), I.getFullHeight())

        picShowPicture.Size = New Size(I.getFullWidth(), I.getFullHeight())

        ScorePopUp.Location = New Point(I.getFullWidth() - 221, 29)
        ScorePopUp.Size = New Size(192, 115)

        StatusBar.Location = New Point(0, I.getFullHeight() - 28)
        StatusBar.Size = New Size(I.getFullWidth(), 28)

        StatusBarAward.Location = New Point(0, 0)
        StatusBarAward.Size = New Size(134, StatusBar.Size.Height)

        StatusBarScore.Location = New Point(I.getFullWidth() - 135, 0)
        StatusBarScore.Size = New Size(135, StatusBar.Size.Height)

        StatusBarTitle.Location = New Point(134, 0)
        StatusBarTitle.Size = New Size(I.getFullWidth() - StatusBarAward.Size.Width - StatusBarScore.Size.Width, StatusBar.Size.Height)

        splash_location_y = (I.getFullHeight - (3 * 110)) / 3
        splashClub.Location = New Point(0, splash_location_y)
        splashClub.Size = New Size(I.getFullWidth(), 100)

        splashTheme.Location = New Point(0, splash_location_y * 2)
        splashTheme.Size = New Size(I.getFullWidth(), 100)

        splashClassification.Location = New Point(0, splash_location_y * 3)
        splashClassification.Size = New Size(I.getFullWidth(), 100)

    End Sub
End Class

