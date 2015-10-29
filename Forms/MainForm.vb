Imports System.Collections.Generic
Imports System.Data.SQLite
Imports System.Globalization
Imports System.Linq
Imports System.Reflection

Namespace Forms

    Public Class MainForm
        Inherits Form

        Private ReadOnly _data_folder As String = My.Computer.FileSystem.GetParentPath(Application.LocalUserAppDataPath)
        Private _database_file_name As String = _data_folder + "\rps.db"

        ' Database
        Public ef_setup As SqLiteConfiguration = New SqLiteConfiguration
        Public rps_context As Entities.rpsEntities
        Private _query As Object
        Private _record As Object
        Public entries As IList(Of Entities.CompetitionEntry)

        ' User Preferences (defaults)
        Public reports_output_folder As String = _data_folder + "\Reports"
        Public images_root_folder As String = _data_folder + "\Photos"
        Private _server_name As String = "localhost"
        Private _server_script_dir As String = "/"
        Public camera_club_name As String = "Raritan Photographic Society"
        Public camera_club_id As Integer = 1
        Public classifications As New ArrayList
        Public mediums As New ArrayList
        Public awards As New ArrayList
        Public themes As New ArrayList
        Public min_score As Integer
        Public max_score As Integer
        Public min_score_for_award As Integer
        Public num_judges As Integer = 1
        Private _all_scores_selected As Boolean
        Private _eights_and_awards_selected As Boolean
        Private _selected_score As Integer
        Private _selected_avg_score As Integer
        Public last_admin_username As String

        Public status_bar As New ProgressStatus

        ' For the thumbnail view
        Private _nine_point_thumb_view_title As String
        Private _eight_point_thumb_view_title As String
        Friend WithEvents data_grid_text_box_column3 As DataGridTextBoxColumn
        Friend WithEvents data_grid_text_box_column2 As DataGridTextBoxColumn
        Friend WithEvents data_grid_text_box_column1 As DataGridTextBoxColumn
        Friend WithEvents data_grid_entries_view As DataGridView
        Friend WithEvents grid_caption As Label
        Private _seven_point_thumb_view_title As String
        Friend WithEvents grid_entries_score As DataGridViewTextBoxColumn
        Friend WithEvents grid_entries_award As DataGridViewTextBoxColumn
        Friend WithEvents grid_entries_title As DataGridViewTextBoxColumn
        Private ReadOnly _center_cell_style As New DataGridViewCellStyle

        Private Sub initializeStatusBar()
            Dim info As StatusBarPanel = New StatusBarPanel
            Dim progress As StatusBarPanel = New StatusBarPanel

            'info.Text = "Ready"
            'info.Width = 592
            progress.Width = 200
            info.Width = Width - progress.Width

            'progress.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
            info.AutoSize = StatusBarPanelAutoSize.Spring

            With status_bar
                .Panels.Add(info)
                .Panels.Add(progress)
                .ShowPanels = True
                .SizingGrip = False
                .setProgressBar = 1
                .progress_bar.Minimum = 0
                .progress_bar.Maximum = 100
            End With

            Controls.Add(status_bar)
        End Sub

#Region " Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()
            Application.EnableVisualStyles()
            'This call is required by the Windows Form Designer.
            InitializeComponent()
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer
        Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
        Friend WithEvents FileExitMenu As System.Windows.Forms.MenuItem
        Friend WithEvents ReportsResultsReportMenu As System.Windows.Forms.MenuItem
        Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
        Friend WithEvents HelpAboutMenu As System.Windows.Forms.MenuItem
        Friend WithEvents btnSlideShow As System.Windows.Forms.Button
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents RecalcAwards As System.Windows.Forms.Button
        Friend WithEvents ReportsMenu As System.Windows.Forms.MenuItem
        Friend WithEvents FileMenu As System.Windows.Forms.MenuItem
        Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
        Friend WithEvents AwardsTableTitleBar As System.Windows.Forms.Label
        Friend WithEvents ReportsScoreSheetMenu As System.Windows.Forms.MenuItem
        Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
        Friend WithEvents ViewSlideShowMenu As System.Windows.Forms.MenuItem
        Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
        Friend WithEvents FilePreferencesMenu As System.Windows.Forms.MenuItem
        Friend WithEvents tbEligibleNines As System.Windows.Forms.TextBox
        Friend WithEvents tbEligibleEights As System.Windows.Forms.TextBox
        Friend WithEvents tbEligibleSevens As System.Windows.Forms.TextBox
        Friend WithEvents Label4 As System.Windows.Forms.Label
        Friend WithEvents SelectAward As System.Windows.Forms.ComboBox
        Friend WithEvents Label5 As System.Windows.Forms.Label
        Friend WithEvents EnableClassification As System.Windows.Forms.CheckBox
        Friend WithEvents EnableMedium As System.Windows.Forms.CheckBox
        Friend WithEvents EnableTheme As System.Windows.Forms.CheckBox
        Friend WithEvents EnableAward As System.Windows.Forms.CheckBox
        Friend WithEvents ViewThumbnailsMenu As System.Windows.Forms.MenuItem
        Friend WithEvents btnThumbnails As System.Windows.Forms.Button
        Friend WithEvents NumNinesHeadingButton As System.Windows.Forms.Button
        Friend WithEvents NumEightsHeadingButton As System.Windows.Forms.Button
        Friend WithEvents NumSevensHeadingButton As System.Windows.Forms.Button
        Friend WithEvents Label6 As System.Windows.Forms.Label
        Friend WithEvents SelectScore As System.Windows.Forms.ComboBox
        Friend WithEvents SelectTheme As System.Windows.Forms.ComboBox
        Friend WithEvents SelectClassification As System.Windows.Forms.ComboBox
        Friend WithEvents SelectMedium As System.Windows.Forms.ComboBox
        Friend WithEvents SelectDate As System.Windows.Forms.ComboBox
        Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
        Friend WithEvents CompCatalogImagesDownload As System.Windows.Forms.MenuItem
        Friend WithEvents CompUploadScores As System.Windows.Forms.MenuItem

        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
            Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
            Me.FileMenu = New System.Windows.Forms.MenuItem()
            Me.FilePreferencesMenu = New System.Windows.Forms.MenuItem()
            Me.MenuItem3 = New System.Windows.Forms.MenuItem()
            Me.FileExitMenu = New System.Windows.Forms.MenuItem()
            Me.MenuItem5 = New System.Windows.Forms.MenuItem()
            Me.MenuItem1 = New System.Windows.Forms.MenuItem()
            Me.CompCatalogImagesDownload = New System.Windows.Forms.MenuItem()
            Me.CompUploadScores = New System.Windows.Forms.MenuItem()
            Me.MenuItem2 = New System.Windows.Forms.MenuItem()
            Me.ViewSlideShowMenu = New System.Windows.Forms.MenuItem()
            Me.ViewThumbnailsMenu = New System.Windows.Forms.MenuItem()
            Me.ReportsMenu = New System.Windows.Forms.MenuItem()
            Me.ReportsScoreSheetMenu = New System.Windows.Forms.MenuItem()
            Me.ReportsResultsReportMenu = New System.Windows.Forms.MenuItem()
            Me.MenuItem7 = New System.Windows.Forms.MenuItem()
            Me.HelpAboutMenu = New System.Windows.Forms.MenuItem()
            Me.SelectClassification = New System.Windows.Forms.ComboBox()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.SelectMedium = New System.Windows.Forms.ComboBox()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.RecalcAwards = New System.Windows.Forms.Button()
            Me.AwardsTableTitleBar = New System.Windows.Forms.Label()
            Me.NumNinesHeadingButton = New System.Windows.Forms.Button()
            Me.NumEightsHeadingButton = New System.Windows.Forms.Button()
            Me.NumSevensHeadingButton = New System.Windows.Forms.Button()
            Me.tbEligibleNines = New System.Windows.Forms.TextBox()
            Me.tbEligibleEights = New System.Windows.Forms.TextBox()
            Me.tbEligibleSevens = New System.Windows.Forms.TextBox()
            Me.SelectAward = New System.Windows.Forms.ComboBox()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.SelectTheme = New System.Windows.Forms.ComboBox()
            Me.EnableClassification = New System.Windows.Forms.CheckBox()
            Me.EnableMedium = New System.Windows.Forms.CheckBox()
            Me.EnableTheme = New System.Windows.Forms.CheckBox()
            Me.EnableAward = New System.Windows.Forms.CheckBox()
            Me.SelectScore = New System.Windows.Forms.ComboBox()
            Me.Label6 = New System.Windows.Forms.Label()
            Me.SelectDate = New System.Windows.Forms.ComboBox()
            Me.btnThumbnails = New System.Windows.Forms.Button()
            Me.btnSlideShow = New System.Windows.Forms.Button()
            Me.data_grid_text_box_column3 = New System.Windows.Forms.DataGridTextBoxColumn()
            Me.data_grid_text_box_column2 = New System.Windows.Forms.DataGridTextBoxColumn()
            Me.data_grid_text_box_column1 = New System.Windows.Forms.DataGridTextBoxColumn()
            Me.data_grid_entries_view = New System.Windows.Forms.DataGridView()
            Me.grid_entries_score = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.grid_entries_award = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.grid_entries_title = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.grid_caption = New System.Windows.Forms.Label()
            CType(Me.data_grid_entries_view, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'MainMenu1
            '
            Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.FileMenu, Me.MenuItem5, Me.MenuItem2, Me.ReportsMenu, Me.MenuItem7})
            '
            'FileMenu
            '
            Me.FileMenu.Index = 0
            Me.FileMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.FilePreferencesMenu, Me.MenuItem3, Me.FileExitMenu})
            Me.FileMenu.Text = "&File"
            '
            'FilePreferencesMenu
            '
            Me.FilePreferencesMenu.Index = 0
            Me.FilePreferencesMenu.Text = "&Preferences..."
            '
            'MenuItem3
            '
            Me.MenuItem3.Index = 1
            Me.MenuItem3.Text = "-"
            '
            'FileExitMenu
            '
            Me.FileExitMenu.Index = 2
            Me.FileExitMenu.Text = "E&xit"
            '
            'MenuItem5
            '
            Me.MenuItem5.Index = 1
            Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.CompCatalogImagesDownload, Me.CompUploadScores})
            Me.MenuItem5.Text = "Competitions"
            '
            'MenuItem1
            '
            Me.MenuItem1.Index = 0
            Me.MenuItem1.Text = "-"
            '
            'CompCatalogImagesDownload
            '
            Me.CompCatalogImagesDownload.Index = 1
            Me.CompCatalogImagesDownload.Text = "Download Images from Server..."
            '
            'CompUploadScores
            '
            Me.CompUploadScores.Index = 2
            Me.CompUploadScores.Text = "Upload Scores..."
            '
            'MenuItem2
            '
            Me.MenuItem2.Index = 2
            Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ViewSlideShowMenu, Me.ViewThumbnailsMenu})
            Me.MenuItem2.Text = "&View"
            '
            'ViewSlideShowMenu
            '
            Me.ViewSlideShowMenu.Index = 0
            Me.ViewSlideShowMenu.Text = "&Slide Show"
            '
            'ViewThumbnailsMenu
            '
            Me.ViewThumbnailsMenu.Index = 1
            Me.ViewThumbnailsMenu.Text = "Thumbnails"
            '
            'ReportsMenu
            '
            Me.ReportsMenu.Index = 3
            Me.ReportsMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ReportsScoreSheetMenu, Me.ReportsResultsReportMenu})
            Me.ReportsMenu.Text = "&Reports"
            '
            'ReportsScoreSheetMenu
            '
            Me.ReportsScoreSheetMenu.Index = 0
            Me.ReportsScoreSheetMenu.Text = "&Score Sheet"
            '
            'ReportsResultsReportMenu
            '
            Me.ReportsResultsReportMenu.Index = 1
            Me.ReportsResultsReportMenu.Text = "Res&ults Report"
            '
            'MenuItem7
            '
            Me.MenuItem7.Index = 4
            Me.MenuItem7.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.HelpAboutMenu})
            Me.MenuItem7.Text = "&Help"
            '
            'HelpAboutMenu
            '
            Me.HelpAboutMenu.Index = 0
            Me.HelpAboutMenu.Text = "&About"
            '
            'SelectClassification
            '
            Me.SelectClassification.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.SelectClassification.Location = New System.Drawing.Point(38, 143)
            Me.SelectClassification.Margin = New System.Windows.Forms.Padding(7, 5, 3, 0)
            Me.SelectClassification.Name = "SelectClassification"
            Me.SelectClassification.Size = New System.Drawing.Size(174, 23)
            Me.SelectClassification.TabIndex = 20
            '
            'Label3
            '
            Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label3.Location = New System.Drawing.Point(38, 122)
            Me.Label3.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(168, 16)
            Me.Label3.TabIndex = 19
            Me.Label3.Text = "Classification"
            '
            'SelectMedium
            '
            Me.SelectMedium.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.SelectMedium.Location = New System.Drawing.Point(38, 194)
            Me.SelectMedium.Margin = New System.Windows.Forms.Padding(7, 5, 3, 0)
            Me.SelectMedium.Name = "SelectMedium"
            Me.SelectMedium.Size = New System.Drawing.Size(174, 23)
            Me.SelectMedium.TabIndex = 18
            '
            'Label2
            '
            Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label2.Location = New System.Drawing.Point(38, 173)
            Me.Label2.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(175, 16)
            Me.Label2.TabIndex = 17
            Me.Label2.Text = "Medium"
            '
            'Label1
            '
            Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label1.Location = New System.Drawing.Point(38, 20)
            Me.Label1.Margin = New System.Windows.Forms.Padding(3, 11, 3, 0)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(175, 16)
            Me.Label1.TabIndex = 15
            Me.Label1.Text = "Competition Date"
            '
            'RecalcAwards
            '
            Me.RecalcAwards.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.RecalcAwards.Location = New System.Drawing.Point(61, 407)
            Me.RecalcAwards.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
            Me.RecalcAwards.Name = "RecalcAwards"
            Me.RecalcAwards.Size = New System.Drawing.Size(129, 23)
            Me.RecalcAwards.TabIndex = 25
            Me.RecalcAwards.Text = "Recalc &Awards"
            '
            'AwardsTableTitleBar
            '
            Me.AwardsTableTitleBar.BackColor = System.Drawing.SystemColors.ActiveCaption
            Me.AwardsTableTitleBar.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.AwardsTableTitleBar.ForeColor = System.Drawing.Color.White
            Me.AwardsTableTitleBar.Location = New System.Drawing.Point(38, 330)
            Me.AwardsTableTitleBar.Margin = New System.Windows.Forms.Padding(0, 11, 0, 0)
            Me.AwardsTableTitleBar.Name = "AwardsTableTitleBar"
            Me.AwardsTableTitleBar.Size = New System.Drawing.Size(175, 24)
            Me.AwardsTableTitleBar.TabIndex = 27
            Me.AwardsTableTitleBar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'NumNinesHeadingButton
            '
            Me.NumNinesHeadingButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.NumNinesHeadingButton.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.NumNinesHeadingButton.Location = New System.Drawing.Point(38, 354)
            Me.NumNinesHeadingButton.Margin = New System.Windows.Forms.Padding(0)
            Me.NumNinesHeadingButton.Name = "NumNinesHeadingButton"
            Me.NumNinesHeadingButton.Size = New System.Drawing.Size(59, 24)
            Me.NumNinesHeadingButton.TabIndex = 34
            Me.NumNinesHeadingButton.Text = "9s"
            '
            'NumEightsHeadingButton
            '
            Me.NumEightsHeadingButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.NumEightsHeadingButton.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.NumEightsHeadingButton.Location = New System.Drawing.Point(96, 354)
            Me.NumEightsHeadingButton.Margin = New System.Windows.Forms.Padding(0)
            Me.NumEightsHeadingButton.Name = "NumEightsHeadingButton"
            Me.NumEightsHeadingButton.Size = New System.Drawing.Size(59, 24)
            Me.NumEightsHeadingButton.TabIndex = 35
            Me.NumEightsHeadingButton.Text = "8s"
            '
            'NumSevensHeadingButton
            '
            Me.NumSevensHeadingButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.NumSevensHeadingButton.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.NumSevensHeadingButton.Location = New System.Drawing.Point(154, 354)
            Me.NumSevensHeadingButton.Margin = New System.Windows.Forms.Padding(0)
            Me.NumSevensHeadingButton.Name = "NumSevensHeadingButton"
            Me.NumSevensHeadingButton.Size = New System.Drawing.Size(59, 24)
            Me.NumSevensHeadingButton.TabIndex = 36
            Me.NumSevensHeadingButton.Text = "7s"
            '
            'tbEligibleNines
            '
            Me.tbEligibleNines.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.tbEligibleNines.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.tbEligibleNines.Location = New System.Drawing.Point(38, 377)
            Me.tbEligibleNines.Margin = New System.Windows.Forms.Padding(0)
            Me.tbEligibleNines.Name = "tbEligibleNines"
            Me.tbEligibleNines.Size = New System.Drawing.Size(59, 23)
            Me.tbEligibleNines.TabIndex = 37
            Me.tbEligibleNines.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
            '
            'tbEligibleEights
            '
            Me.tbEligibleEights.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.tbEligibleEights.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.tbEligibleEights.Location = New System.Drawing.Point(96, 377)
            Me.tbEligibleEights.Margin = New System.Windows.Forms.Padding(0)
            Me.tbEligibleEights.Name = "tbEligibleEights"
            Me.tbEligibleEights.Size = New System.Drawing.Size(59, 23)
            Me.tbEligibleEights.TabIndex = 38
            Me.tbEligibleEights.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
            '
            'tbEligibleSevens
            '
            Me.tbEligibleSevens.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.tbEligibleSevens.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.tbEligibleSevens.Location = New System.Drawing.Point(154, 377)
            Me.tbEligibleSevens.Margin = New System.Windows.Forms.Padding(0)
            Me.tbEligibleSevens.Name = "tbEligibleSevens"
            Me.tbEligibleSevens.Size = New System.Drawing.Size(59, 23)
            Me.tbEligibleSevens.TabIndex = 39
            Me.tbEligibleSevens.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
            '
            'SelectAward
            '
            Me.SelectAward.Enabled = False
            Me.SelectAward.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.SelectAward.Location = New System.Drawing.Point(38, 296)
            Me.SelectAward.Margin = New System.Windows.Forms.Padding(7, 5, 3, 0)
            Me.SelectAward.Name = "SelectAward"
            Me.SelectAward.Size = New System.Drawing.Size(175, 23)
            Me.SelectAward.TabIndex = 40
            '
            'Label4
            '
            Me.Label4.Enabled = False
            Me.Label4.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label4.Location = New System.Drawing.Point(38, 275)
            Me.Label4.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(176, 16)
            Me.Label4.TabIndex = 41
            Me.Label4.Text = "Award"
            '
            'Label5
            '
            Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label5.Location = New System.Drawing.Point(38, 71)
            Me.Label5.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(175, 16)
            Me.Label5.TabIndex = 42
            Me.Label5.Text = "Theme"
            '
            'SelectTheme
            '
            Me.SelectTheme.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.SelectTheme.Location = New System.Drawing.Point(38, 92)
            Me.SelectTheme.Margin = New System.Windows.Forms.Padding(7, 5, 3, 0)
            Me.SelectTheme.Name = "SelectTheme"
            Me.SelectTheme.Size = New System.Drawing.Size(174, 23)
            Me.SelectTheme.TabIndex = 43
            '
            'EnableClassification
            '
            Me.EnableClassification.Checked = True
            Me.EnableClassification.CheckState = System.Windows.Forms.CheckState.Checked
            Me.EnableClassification.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.EnableClassification.Location = New System.Drawing.Point(11, 146)
            Me.EnableClassification.Name = "EnableClassification"
            Me.EnableClassification.Size = New System.Drawing.Size(17, 17)
            Me.EnableClassification.TabIndex = 45
            '
            'EnableMedium
            '
            Me.EnableMedium.Checked = True
            Me.EnableMedium.CheckState = System.Windows.Forms.CheckState.Checked
            Me.EnableMedium.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.EnableMedium.Location = New System.Drawing.Point(11, 197)
            Me.EnableMedium.Name = "EnableMedium"
            Me.EnableMedium.Size = New System.Drawing.Size(17, 17)
            Me.EnableMedium.TabIndex = 46
            '
            'EnableTheme
            '
            Me.EnableTheme.Checked = True
            Me.EnableTheme.CheckState = System.Windows.Forms.CheckState.Checked
            Me.EnableTheme.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.EnableTheme.Location = New System.Drawing.Point(11, 95)
            Me.EnableTheme.Name = "EnableTheme"
            Me.EnableTheme.Size = New System.Drawing.Size(17, 17)
            Me.EnableTheme.TabIndex = 47
            '
            'EnableAward
            '
            Me.EnableAward.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.EnableAward.Location = New System.Drawing.Point(11, 299)
            Me.EnableAward.Name = "EnableAward"
            Me.EnableAward.Size = New System.Drawing.Size(17, 17)
            Me.EnableAward.TabIndex = 48
            '
            'SelectScore
            '
            Me.SelectScore.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.SelectScore.Location = New System.Drawing.Point(38, 245)
            Me.SelectScore.Margin = New System.Windows.Forms.Padding(7, 5, 3, 0)
            Me.SelectScore.Name = "SelectScore"
            Me.SelectScore.Size = New System.Drawing.Size(101, 23)
            Me.SelectScore.TabIndex = 50
            '
            'Label6
            '
            Me.Label6.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label6.Location = New System.Drawing.Point(38, 224)
            Me.Label6.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
            Me.Label6.Name = "Label6"
            Me.Label6.Size = New System.Drawing.Size(80, 16)
            Me.Label6.TabIndex = 51
            Me.Label6.Text = "Score"
            '
            'SelectDate
            '
            Me.SelectDate.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.SelectDate.Location = New System.Drawing.Point(38, 41)
            Me.SelectDate.Margin = New System.Windows.Forms.Padding(3, 5, 3, 0)
            Me.SelectDate.Name = "SelectDate"
            Me.SelectDate.Size = New System.Drawing.Size(174, 23)
            Me.SelectDate.TabIndex = 52
            '
            'btnThumbnails
            '
            Me.btnThumbnails.FlatAppearance.BorderSize = 0
            Me.btnThumbnails.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.btnThumbnails.Image = Global.RPS_Competition_Client.My.Resources.Resources.pictures_32x32
            Me.btnThumbnails.Location = New System.Drawing.Point(180, 236)
            Me.btnThumbnails.Margin = New System.Windows.Forms.Padding(0)
            Me.btnThumbnails.Name = "btnThumbnails"
            Me.btnThumbnails.Size = New System.Drawing.Size(32, 32)
            Me.btnThumbnails.TabIndex = 49
            '
            'btnSlideShow
            '
            Me.btnSlideShow.BackColor = System.Drawing.SystemColors.Control
            Me.btnSlideShow.FlatAppearance.BorderSize = 0
            Me.btnSlideShow.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.btnSlideShow.Image = Global.RPS_Competition_Client.My.Resources.Resources.single_32x32
            Me.btnSlideShow.Location = New System.Drawing.Point(145, 236)
            Me.btnSlideShow.Name = "btnSlideShow"
            Me.btnSlideShow.Size = New System.Drawing.Size(32, 32)
            Me.btnSlideShow.TabIndex = 22
            Me.btnSlideShow.UseVisualStyleBackColor = False
            '
            'data_grid_text_box_column3
            '
            Me.data_grid_text_box_column3.Format = ""
            Me.data_grid_text_box_column3.FormatInfo = Nothing
            Me.data_grid_text_box_column3.HeaderText = "Title"
            Me.data_grid_text_box_column3.MappingName = "Title"
            Me.data_grid_text_box_column3.NullText = ""
            Me.data_grid_text_box_column3.Width = 270
            '
            'data_grid_text_box_column2
            '
            Me.data_grid_text_box_column2.Alignment = System.Windows.Forms.HorizontalAlignment.Center
            Me.data_grid_text_box_column2.Format = ""
            Me.data_grid_text_box_column2.FormatInfo = Nothing
            Me.data_grid_text_box_column2.HeaderText = "Award"
            Me.data_grid_text_box_column2.MappingName = "Award"
            Me.data_grid_text_box_column2.NullText = ""
            Me.data_grid_text_box_column2.Width = 50
            '
            'data_grid_text_box_column1
            '
            Me.data_grid_text_box_column1.Alignment = System.Windows.Forms.HorizontalAlignment.Center
            Me.data_grid_text_box_column1.Format = ""
            Me.data_grid_text_box_column1.FormatInfo = Nothing
            Me.data_grid_text_box_column1.HeaderText = "Score"
            Me.data_grid_text_box_column1.MappingName = "Score_1"
            Me.data_grid_text_box_column1.NullText = ""
            Me.data_grid_text_box_column1.Width = 50
            '
            'data_grid_entries_view
            '
            Me.data_grid_entries_view.AllowUserToAddRows = False
            Me.data_grid_entries_view.AllowUserToDeleteRows = False
            Me.data_grid_entries_view.AllowUserToResizeColumns = False
            Me.data_grid_entries_view.AllowUserToResizeRows = False
            Me.data_grid_entries_view.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.data_grid_entries_view.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
            Me.data_grid_entries_view.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
            Me.data_grid_entries_view.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
            DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
            DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
            DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control
            DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.WindowText
            DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
            Me.data_grid_entries_view.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
            Me.data_grid_entries_view.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.data_grid_entries_view.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.grid_entries_score, Me.grid_entries_award, Me.grid_entries_title})
            DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
            DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
            DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Window
            DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText
            DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
            Me.data_grid_entries_view.DefaultCellStyle = DataGridViewCellStyle2
            Me.data_grid_entries_view.EnableHeadersVisualStyles = False
            Me.data_grid_entries_view.Location = New System.Drawing.Point(225, 42)
            Me.data_grid_entries_view.Name = "data_grid_entries_view"
            Me.data_grid_entries_view.ReadOnly = True
            Me.data_grid_entries_view.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
            Me.data_grid_entries_view.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
            DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.data_grid_entries_view.RowsDefaultCellStyle = DataGridViewCellStyle3
            Me.data_grid_entries_view.Size = New System.Drawing.Size(1030, 387)
            Me.data_grid_entries_view.TabIndex = 53
            '
            'grid_entries_score
            '
            Me.grid_entries_score.DataPropertyName = "Score_1"
            Me.grid_entries_score.FillWeight = 21.80233!
            Me.grid_entries_score.HeaderText = "Score"
            Me.grid_entries_score.Name = "grid_entries_score"
            Me.grid_entries_score.ReadOnly = True
            Me.grid_entries_score.Width = 60
            '
            'grid_entries_award
            '
            Me.grid_entries_award.DataPropertyName = "Award"
            Me.grid_entries_award.FillWeight = 21.80233!
            Me.grid_entries_award.HeaderText = "Award"
            Me.grid_entries_award.Name = "grid_entries_award"
            Me.grid_entries_award.ReadOnly = True
            Me.grid_entries_award.Width = 65
            '
            'grid_entries_title
            '
            Me.grid_entries_title.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.grid_entries_title.DataPropertyName = "Title"
            Me.grid_entries_title.FillWeight = 256.3954!
            Me.grid_entries_title.HeaderText = "Title"
            Me.grid_entries_title.Name = "grid_entries_title"
            Me.grid_entries_title.ReadOnly = True
            '
            'grid_caption
            '
            Me.grid_caption.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.grid_caption.BackColor = System.Drawing.SystemColors.ActiveCaption
            Me.grid_caption.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.grid_caption.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.grid_caption.ForeColor = System.Drawing.Color.Black
            Me.grid_caption.Location = New System.Drawing.Point(225, 20)
            Me.grid_caption.Margin = New System.Windows.Forms.Padding(0)
            Me.grid_caption.Name = "grid_caption"
            Me.grid_caption.Size = New System.Drawing.Size(1030, 24)
            Me.grid_caption.TabIndex = 54
            Me.grid_caption.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'MainForm
            '
            Me.AutoScroll = True
            Me.AutoScrollMinSize = New System.Drawing.Size(640, 480)
            Me.ClientSize = New System.Drawing.Size(1273, 491)
            Me.Controls.Add(Me.grid_caption)
            Me.Controls.Add(Me.data_grid_entries_view)
            Me.Controls.Add(Me.SelectDate)
            Me.Controls.Add(Me.Label6)
            Me.Controls.Add(Me.SelectScore)
            Me.Controls.Add(Me.btnThumbnails)
            Me.Controls.Add(Me.EnableAward)
            Me.Controls.Add(Me.EnableTheme)
            Me.Controls.Add(Me.EnableMedium)
            Me.Controls.Add(Me.EnableClassification)
            Me.Controls.Add(Me.SelectTheme)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.SelectAward)
            Me.Controls.Add(Me.tbEligibleSevens)
            Me.Controls.Add(Me.tbEligibleEights)
            Me.Controls.Add(Me.tbEligibleNines)
            Me.Controls.Add(Me.NumSevensHeadingButton)
            Me.Controls.Add(Me.NumEightsHeadingButton)
            Me.Controls.Add(Me.NumNinesHeadingButton)
            Me.Controls.Add(Me.AwardsTableTitleBar)
            Me.Controls.Add(Me.RecalcAwards)
            Me.Controls.Add(Me.btnSlideShow)
            Me.Controls.Add(Me.SelectClassification)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.SelectMedium)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.Label1)
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.Menu = Me.MainMenu1
            Me.Name = "MainForm"
            Me.Padding = New System.Windows.Forms.Padding(11)
            Me.Text = "RPS Digital Competition Viewer"
            Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
            CType(Me.data_grid_entries_view, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

#End Region

        Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            Dim dirinfo As DirectoryInfo
            status_bar.progress_bar.Hide()
            Try
                ' Set up the connection string for the database connection
                Set_database_name(_database_file_name)
                ' Load the user preferences from the registry
                getPreferences()

                rps_context = New Entities.rpsEntities(New SQLiteConnectionStringBuilder() With {
                                                 .DataSource = _database_file_name,
                                                 .ForeignKeys = True
                                                 }.ConnectionString)
                If Not File.Exists(_database_file_name) Then
                    SQLiteConnection.CreateFile(_database_file_name)
                    initializeDatabase()
                End If

                dirinfo = New DirectoryInfo(images_root_folder)
                If Not dirinfo.Exists Then
                    dirinfo.Create()
                End If
                dirinfo = New DirectoryInfo(reports_output_folder)
                If Not dirinfo.Exists Then
                    dirinfo.Create()
                End If

                ' Setup Datagrid Styles
                _center_cell_style.Alignment = DataGridViewContentAlignment.MiddleCenter

                getClubRules()

                ' Initialize the StatusBar and ProgressBar
                initializeStatusBar()
                status_bar.progress_bar.Value = 0
                ' Load the unique competition dates into the Competition Date combobox
                setCompetitionDatesCombobox()

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub initializeDatabase()
            _query = "CREATE TABLE `medium` (`name` TEXT, `id` INTEGER Not NULL PRIMARY KEY AUTOINCREMENT UNIQUE);" &
                     "CREATE TABLE 'club_medium' (`club_id` INTEGER Not NULL,`medium_id` INTEGER Not NULL,`sort_key` INTEGER DEFAULT 0, PRIMARY KEY(club_id, medium_id), FOREIGN KEY(`club_id`) REFERENCES club ( id ), FOREIGN KEY(`medium_id`) REFERENCES medium ( id ));" &
                     "CREATE TABLE 'club_classification' (`club_id` INTEGER Not NULL,`classification_id` INTEGER Not NULL,`sort_key` INTEGER DEFAULT 0,PRIMARY KEY(club_id, classification_id),FOREIGN KEY(`club_id`) REFERENCES club ( id ),FOREIGN KEY(`classification_id`) REFERENCES classification ( id ));" &
                     "CREATE TABLE 'club_award' (`club_id` INTEGER Not NULL,`award_id` INTEGER Not NULL,`points` INTEGER,`sort_key` INTEGER,PRIMARY KEY(club_id, award_id),FOREIGN KEY(`club_id`) REFERENCES club ( id ),FOREIGN KEY(`award_id`) REFERENCES award ( id ));" &
                     "CREATE TABLE 'club' (`short_name` TEXT,`name` TEXT,`id` INTEGER Not NULL PRIMARY KEY AUTOINCREMENT UNIQUE,`min_score` INTEGER DEFAULT 0,`max_score` INTEGER DEFAULT 0,`min_score_for_award` INTEGER DEFAULT 0);" &
                     "CREATE TABLE 'classification' (`name` TEXT,`id` INTEGER Not NULL PRIMARY KEY AUTOINCREMENT UNIQUE);" &
                     "CREATE TABLE 'award' (`name` TEXT,`id` INTEGER Not NULL PRIMARY KEY AUTOINCREMENT UNIQUE);" &
                     "CREATE TABLE 'CompetitionEntries' (`Photo_ID` INTEGER Not NULL PRIMARY KEY AUTOINCREMENT,`Title` TEXT,`Maker` TEXT,`Classification` TEXT,`Medium` TEXT,`Theme` TEXT,`Competition_Date_1` TEXT,`Score_1` INTEGER DEFAULT 0,`Award` TEXT,`Image_File_Name` TEXT,`Display_Sequence` INTEGER DEFAULT 0,`Server_Entry_ID` INTEGER DEFAULT 0);"
            rps_context.Database.ExecuteSqlCommand(_query)

            _query = "INSERT INTO `medium` (name,id) VALUES ('Color Digital',1);" &
                     "INSERT INTO `medium` (name,id) VALUES ('Color Prints',2);" &
                     "INSERT INTO `medium` (name,id) VALUES ('B&W Digital',3);" &
                     "INSERT INTO `medium` (name,id) VALUES ('B&W Prints',4);" &
                     "INSERT INTO `club` (short_name,name,id,min_score,max_score,min_score_for_award) VALUES ('RPS','Raritan Photographic Society',1,5,9,7);" &
                     "INSERT INTO `classification` (name,id) VALUES ('Beginner',1);" &
                     "INSERT INTO `classification` (name,id) VALUES ('Advanced',2);" &
                     "INSERT INTO `classification` (name,id) VALUES ('Salon',3);" &
                     "INSERT INTO `award` (name,id) VALUES ('1st',1);" &
                     "INSERT INTO `award` (name,id) VALUES ('2nd',2);" &
                     "INSERT INTO `award` (name,id) VALUES ('3rd',3);" &
                     "INSERT INTO `award` (name,id) VALUES ('HM',4);" &
                     "INSERT INTO `club_medium` (club_id,medium_id,sort_key) VALUES (1,1,1);" &
                     "INSERT INTO `club_medium` (club_id,medium_id,sort_key) VALUES (1,2,3);" &
                     "INSERT INTO `club_medium` (club_id,medium_id,sort_key) VALUES (1,3,0);" &
                     "INSERT INTO `club_medium` (club_id,medium_id,sort_key) VALUES (1,4,2);" &
                     "INSERT INTO `club_classification` (club_id,classification_id,sort_key) VALUES (1,1,0);" &
                     "INSERT INTO `club_classification` (club_id,classification_id,sort_key) VALUES (1,2,1);" &
                     "INSERT INTO `club_classification` (club_id,classification_id,sort_key) VALUES (1,3,2);" &
                     "INSERT INTO `club_award` (club_id,award_id,points,sort_key) VALUES (1,1,NULL,1);" &
                     "INSERT INTO `club_award` (club_id,award_id,points,sort_key) VALUES (1,2,NULL,2);" &
                     "INSERT INTO `club_award` (club_id,award_id,points,sort_key) VALUES (1,3,NULL,3);" &
                     "INSERT INTO `club_award` (club_id,award_id,points,sort_key) VALUES (1,4,NULL,4);"
            rps_context.Database.ExecuteSqlCommand(_query)
        End Sub


        Private Sub btnSlideShow_Click(sender As Object, e As EventArgs) _
            Handles btnSlideShow.Click
            doSlideShow()
        End Sub

        Private Sub btnThumbnails_Click(sender As Object, e As EventArgs) _
            Handles btnThumbnails.Click
            doPickAwards()
        End Sub

        Private Sub FileCatalogImagesDownload_Click(sender As Object, e As EventArgs) _
            Handles CompCatalogImagesDownload.Click
            downloadCompetitionImages()
            setCompetitionDatesCombobox()
        End Sub

        Private Sub FileUploadScores_Click(sender As Object, e As EventArgs) _
            Handles CompUploadScores.Click
            uploadScores()
        End Sub

        Private Sub FileExitMenu_Click(sender As Object, e As EventArgs) _
            Handles FileExitMenu.Click
            Close()
        End Sub


        Private Sub SelectMedium_SelectedIndexChanged(sender As Object, e As EventArgs) _
            Handles SelectMedium.SelectedIndexChanged
            SelectScore.SelectedItem = "All"
            _all_scores_selected = True
            _eights_and_awards_selected = False
            getSelectedEntries()
        End Sub

        Private Sub SelectClassification_SelectedIndexChanged(sender As Object, e As EventArgs) _
            Handles SelectClassification.SelectedIndexChanged
            SelectScore.SelectedItem = "All"
            _all_scores_selected = True
            _eights_and_awards_selected = False
            getSelectedEntries()
        End Sub

        Private Sub SelectTheme_SelectedIndexChanged(sender As Object, e As EventArgs) _
            Handles SelectTheme.SelectedIndexChanged
            SelectScore.SelectedItem = "All"
            _all_scores_selected = True
            _eights_and_awards_selected = False
            getSelectedEntries()
        End Sub

        Private Sub SelectAward_SelectedIndexChanged(sender As Object, e As EventArgs) _
            Handles SelectAward.SelectedIndexChanged
            SelectScore.SelectedItem = "All"
            _all_scores_selected = True
            _eights_and_awards_selected = False
            getSelectedEntries()
        End Sub

        Private Sub EnableMedium_CheckedChanged(sender As Object, e As EventArgs) _
            Handles EnableMedium.CheckedChanged
            If EnableMedium.CheckState = CheckState.Checked Then
                SelectMedium.Enabled = True
                Label2.Enabled = True
            Else
                SelectMedium.Enabled = False
                Label2.Enabled = False
            End If
            getSelectedEntries()
        End Sub

        Private Sub EnableClassification_CheckedChanged(sender As Object, e As EventArgs) _
            Handles EnableClassification.CheckedChanged
            If EnableClassification.CheckState = CheckState.Checked Then
                SelectClassification.Enabled = True
                Label3.Enabled = True
            Else
                SelectClassification.Enabled = False
                Label3.Enabled = False
            End If
            getSelectedEntries()
        End Sub

        Private Sub EnableTheme_CheckedChanged(sender As Object, e As EventArgs) _
            Handles EnableTheme.CheckedChanged
            If EnableTheme.CheckState = CheckState.Checked Then
                SelectTheme.Enabled = True
                Label5.Enabled = True
            Else
                SelectTheme.Enabled = False
                Label5.Enabled = False
            End If
            getSelectedEntries()
        End Sub

        Private Sub EnableAward_CheckedChanged(sender As Object, e As EventArgs) _
            Handles EnableAward.CheckedChanged
            If EnableAward.CheckState = CheckState.Checked Then
                SelectAward.Enabled = True
                Label4.Enabled = True
            Else
                SelectAward.Enabled = False
                Label4.Enabled = False
            End If
            getSelectedEntries()
        End Sub

        Private Sub RecalcAwards_Click(sender As Object, e As EventArgs) _
            Handles RecalcAwards.Click
            doCalculateAwards()
        End Sub

        Private Sub HelpAboutMenu_Click(sender As Object, e As EventArgs) _
            Handles HelpAboutMenu.Click
            Dim about_form As New About
            about_form.Show()
        End Sub

        Private Sub ViewSlideShowMenu_Click(sender As Object, e As EventArgs) _
            Handles ViewSlideShowMenu.Click
            doSlideShow()
        End Sub

        Private Sub ViewThumbnailsMenu_Click(sender As Object, e As EventArgs) _
            Handles ViewThumbnailsMenu.Click
            doPickAwards()
        End Sub

        Private Sub MainForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
            'GridResizeColumns(grdCompetition_Entries, 60, 0, 10, 10, 50, 30)
            'GridResizeColumns(grdCompetition_Entries, 60, 0, 15, 15, 70)
        End Sub

        Private Sub FilePreferencesMenu_Click(sender As Object, e As EventArgs) _
            Handles FilePreferencesMenu.Click
            getUserPreferences()
        End Sub

        Private Sub ReportsResultsReportMenu_Click(sender As Object, e As EventArgs) _
            Handles ReportsResultsReportMenu.Click
            doResultReport(True)
        End Sub

        Private Sub ReportsScoreSheetMenu_Click(sender As Object, e As EventArgs) _
            Handles ReportsScoreSheetMenu.Click
            'ScoreSheet()
            doResultReport(False)
        End Sub

        Private Sub doSlideShow()
            Dim viewer As ImageViewer
            Dim show_splash As Boolean
            Dim status_bar_state As Integer

            ' Bail out if the dataset is empty
            If data_grid_entries_view.RowCount() <= 0 Then
                MsgBox("No competition has been loaded.", MsgBoxStyle.Exclamation, "Error In DoSlideShow()")
                Exit Sub
            End If

            ' Display the splash screen if about to view all images in a competition starting
            ' from the beginning
            'If AllScoresRadioButton.Checked And grdCompetition_Entries.CurrentRowIndex <= 0 Then
            If _all_scores_selected And data_grid_entries_view.CurrentCell.RowIndex <= 0 Then
                show_splash = True
            Else
                show_splash = False
            End If

            ' Set the status bar to show the title and maker name if we're announcing the winners
            'If EightsAndAwardsRadioButton.Checked Then
            If _eights_and_awards_selected Then
                status_bar_state = 2
                ' Set the status bar to show the title only if we're assigning awards
                'ElseIf NineScoreRadioButton.Checked Or EightScoreRadioButton.Checked Or SevenScoreRadioButton.Checked Then
            ElseIf Not _all_scores_selected Then
                status_bar_state = 1
            Else
                status_bar_state = 0
            End If

            viewer = New ImageViewer(Me,
                                     entries,
                                     data_grid_entries_view.CurrentCell.RowIndex,
                                     show_splash,
                                     status_bar_state)
            Cursor.Hide()
            viewer.ShowDialog()
            Cursor.Show()
            Try
                ' Attempt to update the datasource.
                data_grid_entries_view.Refresh()

                For Each entry As Entities.CompetitionEntry In entries
                    _query = "UPDATE CompetitionEntries SET Score_1=@score , Award=@award Where Server_Entry_ID=@key"
                    rps_context.Database.ExecuteSqlCommand(_query,
                                                           New SQLiteParameter("@score", entry.Score_1),
                                                           New SQLiteParameter("@award", entry.Award),
                                                           New SQLiteParameter("@key", entry.Server_Entry_ID)
                                                           )
                Next
                ' If we've just completed entering scores, calculate the eligible awards
                'If AllScoresRadioButton.Checked Then
                If _all_scores_selected Then
                    doCalculateAwards()
                    'PickAwards()
                End If
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub
        '
        ' Convert a date string in the form of dd-MMM-yyyy e.g 19-Nov-2009 to a date
        '
        Private Function parseSelectedDate(s As String) As Date
            Dim d As Date
            Try
                d = Date.ParseExact(s, "dd-MMM-yyyy", CultureInfo.CurrentCulture)
                parseSelectedDate = d
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
                parseSelectedDate = Nothing
            End Try
        End Function

        Private Sub doPickAwards()
            Dim screen_title As String = ""
            Dim viewer As ThumbnailViewer
            Try
                ' Bail out if the dataset is empty
                If entries.Count <= 0 Then
                    MsgBox("No images loaded.", MsgBoxStyle.Exclamation, "Error In PickAwards()")
                    Exit Sub
                End If

                ' Configure the title for the thumbnail screen
                'If AllScoresRadioButton.Checked Then
                '    screenTitle = SelectClassification.Text + " " + SelectMedium.Text
                'ElseIf NineScoreRadioButton.Checked Then
                '    screenTitle = ninePointThumbViewTitle
                'ElseIf EightScoreRadioButton.Checked Then
                '    screenTitle = eightPointThumbViewTitle
                'ElseIf SevenScoreRadioButton.Checked Then
                '    screenTitle = sevenPointThumbViewTitle
                'End If


                ' Configure the title for the thumbnail screen
                If _all_scores_selected Then
                    screen_title = SelectClassification.Text + " " + SelectMedium.Text
                ElseIf _eights_and_awards_selected Then
                    If num_judges > 1 Then
                        screen_title = "Award winners And images averaging 8 points Or more"
                    Else
                        screen_title = "Award winners And images With 8 points Or more"
                    End If
                ElseIf _selected_avg_score > 0 Then
                    If _selected_avg_score = 9 Then
                        screen_title = _nine_point_thumb_view_title
                    ElseIf _selected_avg_score = 8 Then
                        screen_title = _eight_point_thumb_view_title
                    ElseIf _selected_avg_score = 7 Then
                        screen_title = _seven_point_thumb_view_title
                    End If
                Else
                    screen_title = "Images scoring " + CType(_selected_score, String) + " points"
                End If

                ' Launch the thumbnail screen
                viewer = New ThumbnailViewer(Me, entries, screen_title)
                viewer.ShowDialog()

                ' Attempt to update the datasource.
                data_grid_entries_view.Refresh()

                For Each entry As Entities.CompetitionEntry In entries
                    _query = "UPDATE CompetitionEntries SET Score_1=@score , Award=@award Where Server_Entry_ID=@key"
                    rps_context.Database.ExecuteSqlCommand(_query,
                                                           New SQLiteParameter("@score", entry.Score_1),
                                                           New SQLiteParameter("@award", entry.Award),
                                                           New SQLiteParameter("@key", entry.Server_Entry_ID)
                                                           )
                Next

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub addImageToDatabase(
                                       file As FileInfo,
                                       maker As String,
                                       title As String,
                                       score As String,
                                       award As String,
                                       classification As String,
                                       medium As String,
                                       competition_date As Date,
                                       competition_theme As String,
                                       entry_id As String,
                                       sequence As Integer)

            Dim relative_path As String
            Dim entry As Entities.CompetitionEntry = New Entities.CompetitionEntry

            Try
                ' Calculate the relative path to the file.  The path is relative to the
                ' imagesRootFolder set in File -> Preferences
                If InStr(1, file.FullName, images_root_folder) = 1 Then
                    relative_path = Mid(file.FullName, Len(images_root_folder) + 2)
                Else
                    relative_path = file.FullName    ' Store absolute path if can't calculate relative path
                End If

                ' Fill in the field values
                entry.Title = title
                entry.Maker = maker
                entry.Classification = classification
                entry.Medium = medium
                entry.Theme = competition_theme
                entry.Competition_Date_1 = competition_date
                entry.Image_File_Name = relative_path
                If sequence = 0 Then
                    entry.Display_Sequence = file.Length Mod 61
                Else
                    entry.Display_Sequence = sequence
                End If
                If entry_id > "" Then
                    entry.Server_Entry_ID = entry_id
                Else
                    entry.Server_Entry_ID = Nothing
                End If
                If score > "" Then
                    entry.Score_1 = score
                Else
                    entry.Score_1 = Nothing
                End If
                If award > "" Then
                    entry.Award = award
                Else
                    entry.Award = Nothing
                End If

                ' Add it to the database
                rps_context.CompetitionEntries.Add(entry)

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
            rps_context.SaveChanges()
        End Sub

        Private Function setMaxAwards(num_images As Double) As Integer
            Try
                If num_images.Equals(1) Then
                    setMaxAwards = 1
                Else
                    setMaxAwards = Int((num_images / 4) + 0.5)
                End If
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Function

        Private Sub getSelectedEntries()
            Dim num_selected As Double
            Dim select_stmt As String
            Dim where_clause As String
            Dim order_clause As String

            If SelectDate.Text > "" And SelectMedium.Text > "" And SelectClassification.Text > "" Then
                Try
                    ' Build the complete SQL statement to select the records specified
                    ' by the selection criteria on the main form.
                    ' Start with a basic SQL statement that selects records by date.
                    select_stmt = "Select * FROM CompetitionEntries"
                    where_clause = " WHERE Competition_Date_1='" + Format(parseSelectedDate(SelectDate.Text), "M/dd/yyyy") +
                                   "'"
                    order_clause = " ORDER BY Display_Sequence, Title"
                    grid_caption.Text = Format(parseSelectedDate(SelectDate.Text), "MM/dd/yyyy")

                    'Dim q As System.Linq.IQueryable(Of CompetitionEntry)
                    'q = From entry In rpsContext.CompetitionEntries
                    'Select Case entry.Award, entry.Classification, entry.Competition_Date_1, entry.Display_Sequence, entry.Image_File_Name, entry.Maker, entry.Medium, entry.Photo_ID, entry.Score_1, entry.Server_Entry_ID, entry.Theme, entry.Title

                    'q = q.Where(Function(entry) entry.Competition_Date_1 = Format(ParseSelectedDate(SelectDate.Text), "MM/dd/yyyy"))
                    ' If enabled, add the value of the Classification field to the selection criteria
                    If EnableClassification.CheckState = CheckState.Checked Then
                        where_clause += " AND Classification='" + SelectClassification.Text + "'"
                        grid_caption.Text += "  -  " + SelectClassification.Text
                        If EnableMedium.CheckState = CheckState.Checked Then
                            grid_caption.Text += " / "
                        End If
                    End If

                    ' If enabled, add the value of the Medium field to the selection criteria
                    If EnableMedium.CheckState = CheckState.Checked Then
                        If EnableClassification.CheckState = CheckState.Unchecked Then
                            grid_caption.Text += "  -  "
                        End If
                        where_clause += " AND Medium='" + SelectMedium.Text + "'"
                        grid_caption.Text += SelectMedium.Text
                    End If

                    If Not _all_scores_selected Then
                        If _eights_and_awards_selected Then
                            where_clause += " AND ((Award Is Not Null) OR (round(Score_1/" + CType(num_judges, String) +
                                            ", 0) >= 8 AND Award Is Null))"
                            ' "CASE WHEN Award is NULL THEN 0 ELSE 1 END" Ensure the NULL values are shown first.
                            order_clause = " ORDER BY CASE WHEN Award is NULL THEN 0 ELSE 1 END, Award DESC, Score_1 ASC"
                            grid_caption.Text += " (8s and Awards)"
                        ElseIf _selected_avg_score > 0 Then
                            where_clause += " AND round(Score_1/" + CType(num_judges, String) + ", 0) = " +
                                            CType(_selected_avg_score, String)
                            grid_caption.Text += " (Avg of " + CType(_selected_avg_score, String) + " points)"
                        Else
                            where_clause += " AND Score_1=" + CType(_selected_score, String)
                            grid_caption.Text += " (" + CType(_selected_score, String) + " points only)"
                        End If
                    End If

                    ' If enabled, add the value of the Theme field to the selection criteria
                    If EnableTheme.CheckState = CheckState.Checked Then
                        If SelectTheme.Text > "" Then
                            where_clause += " AND Theme='" + SelectTheme.Text + "'"
                        End If
                    End If

                    ' If enabled, add the value of the Award field to the selection criteria
                    If EnableAward.CheckState = CheckState.Checked Then
                        If SelectAward.Text > "" Then
                            where_clause += " AND Award='" + SelectAward.Text + "'"
                            grid_caption.Text += SelectAward.Text + " only"
                        End If
                    End If

                    ' Install the updated SQL SELECT statement
                    _query = select_stmt + where_clause + order_clause
                    entries = rps_context.Database.SqlQuery(Of Entities.CompetitionEntry)(_query).ToList

                    With data_grid_entries_view
                        .Columns("grid_entries_score").DefaultCellStyle = _center_cell_style
                        .Columns("grid_entries_award").DefaultCellStyle = _center_cell_style
                        .AutoGenerateColumns = False
                        .DataSource = entries
                    End With

                    ' Count the number of rows selected and add it to the caption of the DataGrid
                    num_selected = data_grid_entries_view.RowCount()

                    grid_caption.Text += "  -  " + num_selected.ToString + " Images"

                    ' Recalculate the awards
                    'If AllScoresRadioButton.Checked And EnableAward.CheckState = CheckState.Unchecked Then
                    If _all_scores_selected And EnableAward.CheckState = CheckState.Unchecked Then
                        doCalculateAwards()
                    End If

                Catch exception As Exception
                    MsgBox(exception.Message, , "Error In " + MethodBase.GetCurrentMethod().ToString)
                End Try
            End If
        End Sub


        Private Sub doCalculateAwards()
            Dim eligible_scores As New ArrayList
            Dim maximum_awards As Integer
            Dim i As Integer
            Dim num_eligible_nines As Integer
            Dim num_eligible_eights As Integer
            Dim num_eligible_sevens As Integer
            Dim total_num_scores As Integer
            Dim award_names() As String = {"1st", "2nd", "3rd", "HM"}
            Dim delim_9 As String = ""
            Dim delim_8 As String = ""
            Dim delim_7 As String = ""
            Dim num_nine_hm, num_eight_hm, num_seven_hm As Integer
            Dim eligible_score As Double

            Try
                ' What is the maximum number of Awards possible
                maximum_awards = setMaxAwards(CType(entries.Count, Double))

                ' Iterate through the dataset and record all the scores which are eligible for an award
                For Each entry As Entities.CompetitionEntry In From entry1 In entries Where IsNumeric(entry1.Score_1)
                    total_num_scores = total_num_scores + 1
                    If entry.Score_1 >= (min_score_for_award * num_judges) And entry.Score_1 <= (max_score * num_judges) Then
                        eligible_scores.Add(entry.Score_1)
                    End If
                Next
                ' Sort descending
                eligible_scores.Sort()
                eligible_scores.Reverse()

                ' Step through the eligible scores until all possible awards have been allocated
                ' or until the number of eligible scores exhausted.
                _nine_point_thumb_view_title = ""
                _eight_point_thumb_view_title = ""
                _seven_point_thumb_view_title = ""
                For i = 0 To Math.Min(maximum_awards, eligible_scores.Count) - 1
                    ' Count up the number of eligible 9s, 8s and 7s
                    eligible_score = Math.Round(eligible_scores(i) / num_judges, 0)
                    Select Case eligible_score
                        Case 9
                            ' Count the number of 9 point awards for the distribution on the main screen
                            num_eligible_nines = num_eligible_nines + 1
                            ' Build the title for the thumbnail screen
                            If i < 3 Then
                                _nine_point_thumb_view_title += delim_9 + award_names(i)
                            Else
                                num_nine_hm += 1
                            End If
                            delim_9 = ", "
                        Case 8
                            ' Count the number of 8 point awards for the distribution on the main screen
                            num_eligible_eights = num_eligible_eights + 1
                            ' Build the title for the thumbnail screen
                            If i < 3 Then
                                _eight_point_thumb_view_title += delim_8 + award_names(i)
                            Else
                                num_eight_hm += 1
                            End If
                            delim_8 = ", "
                        Case 7
                            ' Count the number of 7 point awards for the distribution on the main screen
                            num_eligible_sevens = num_eligible_sevens + 1
                            ' Build the title for the thumbnail screen
                            If i < 3 Then
                                _seven_point_thumb_view_title += delim_7 + award_names(i)
                            Else
                                num_seven_hm += 1
                            End If
                            delim_7 = ", "
                    End Select
                Next

                ' Update the 9 point thumbnail screen title to include the HMs
                If num_nine_hm > 0 Then
                    If _nine_point_thumb_view_title > "" Then
                        _nine_point_thumb_view_title += " And "
                    End If
                    If num_nine_hm = 1 Then
                        _nine_point_thumb_view_title += "1 HM"
                    Else
                        _nine_point_thumb_view_title += CStr(num_nine_hm) + " HMs"
                    End If
                End If

                ' Update the 8 point thumbnail screen title to include the HMs
                If num_eight_hm > 0 Then
                    If _eight_point_thumb_view_title > "" Then
                        _eight_point_thumb_view_title += " And "
                    End If
                    If num_eight_hm = 1 Then
                        _eight_point_thumb_view_title += "1 HM"
                    Else
                        _eight_point_thumb_view_title += CStr(num_eight_hm) + " HMs"
                    End If
                End If

                ' Update the 7 point thumbnail screen title to include the HMs
                If num_seven_hm > 0 Then
                    If _seven_point_thumb_view_title > "" Then
                        _seven_point_thumb_view_title += " And "
                    End If
                    If num_seven_hm = 1 Then
                        _seven_point_thumb_view_title += "1 HM"
                    Else
                        _seven_point_thumb_view_title += CStr(num_seven_hm) + " HMs"
                    End If
                End If

                ' Add a prefix to the thumbnail screen title
                If _nine_point_thumb_view_title > "" Then
                    _nine_point_thumb_view_title = "Choose " + _nine_point_thumb_view_title
                Else
                    _nine_point_thumb_view_title = "Choose (none)"
                End If
                If _eight_point_thumb_view_title > "" Then
                    _eight_point_thumb_view_title = "Choose " + _eight_point_thumb_view_title
                Else
                    _eight_point_thumb_view_title = "Choose (none)"
                End If
                If _seven_point_thumb_view_title > "" Then
                    _seven_point_thumb_view_title = "Choose " + _seven_point_thumb_view_title
                Else
                    _seven_point_thumb_view_title = "Choose (none)"
                End If

                ' Enter the total awards to be chosen into the "caption"
                If total_num_scores > 0 Then
                    AwardsTableTitleBar.Text = Math.Min(maximum_awards, eligible_scores.Count).ToString + " Awards"
                Else
                    AwardsTableTitleBar.Text = maximum_awards.ToString + " Awards"
                End If
                tbEligibleNines.Text = num_eligible_nines.ToString
                tbEligibleEights.Text = num_eligible_eights.ToString
                tbEligibleSevens.Text = num_eligible_sevens.ToString

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub


        Private Sub doResultReport(display_scores As Boolean)
            Dim temp_file As String
            Dim report_type As String
            Dim competition_date As Date
            Dim entry As Entities.CompetitionEntry
            Dim sw As StreamWriter
            Try
                ' Bail out if the dataset is empty
                If entries.Count <= 0 Then
                    MsgBox("No images selected." + vbCrLf + "Select one Or more images before running the report",
                           MsgBoxStyle.Exclamation,
                           "Error In ResultsReport()")
                    Exit Sub
                End If

                ' Open the output file and write the HTML preamble
                'competitionDate = CType(SelectDate.Text, Date)
                competition_date = parseSelectedDate(SelectDate.Text)
                If display_scores Then
                    report_type = "Results"
                Else
                    report_type = "Scoresheet"
                End If
                temp_file = reports_output_folder + "\" +
                            report_type + "_" + CType(competition_date.Year, String) + "-" +
                            CType(competition_date.Month, String) + "-" +
                            CType(competition_date.Day, String)
                If EnableTheme.Checked Then
                    temp_file += "_" + SelectTheme.Text
                End If
                If EnableClassification.Checked Then
                    temp_file += "_" + SelectClassification.Text
                End If
                If EnableMedium.Checked Then
                    temp_file += "_" + SelectMedium.Text
                End If
                If SelectScore.Text <> "All" Then
                    temp_file += "_" + SelectScore.Text
                End If
                If EnableAward.Checked Then
                    temp_file += "_" + SelectAward.Text
                End If
                temp_file = handleStrMap(temp_file, " ?[]/\=+<>:;"",*|", "_---------------") + ".html"
                sw = File.CreateText(temp_file)
                sw.WriteLine("<html><style type=""text/css"">")
                sw.WriteLine("<!--")
                sw.WriteLine("body {")
                sw.WriteLine("	font-family: Arial, Helvetica, sans-serif;")
                sw.WriteLine("	margin: 0px;")
                sw.WriteLine("	padding: 0px;")
                sw.WriteLine("	font-size: 12pt;")
                sw.WriteLine("}")
                sw.WriteLine("table {")
                sw.WriteLine("	border: none;")
                sw.WriteLine("	border-collapse:collapse;")
                sw.WriteLine("	border-spacing: 0px;")
                sw.WriteLine("	width: 100%;")
                sw.WriteLine("}")
                sw.WriteLine("table.worksheet {")
                sw.WriteLine("	border: 1px solid black;")
                sw.WriteLine("}")
                sw.WriteLine("th {")
                sw.WriteLine("	font-weight: bold;")
                sw.WriteLine("	color: #FFFFFF;")
                sw.WriteLine("	background-color: #666666;")
                sw.WriteLine("	Text-align: center;")
                sw.WriteLine("	border: 1px solid #666666;")
                sw.WriteLine("}")
                sw.WriteLine("td {")
                sw.WriteLine("	border: 1px solid black;")
                sw.WriteLine("}")
                sw.WriteLine("td.header_left {")
                sw.WriteLine("	Text-align: left;")
                sw.WriteLine("  border: none;")
                sw.WriteLine("}")
                sw.WriteLine("td.header_right {")
                sw.WriteLine("	Text-align: right;")
                sw.WriteLine("  border: none;")
                sw.WriteLine("}")
                sw.WriteLine("td.header_center {")
                sw.WriteLine("	Text-align: center;")
                sw.WriteLine("  border: none;")
                sw.WriteLine("}")
                sw.WriteLine("td.score {")
                sw.WriteLine("	Text-align: center;")
                sw.WriteLine("	width: 10%;")
                sw.WriteLine("}")
                sw.WriteLine("td.award {")
                sw.WriteLine("	Text-align: center;")
                sw.WriteLine("	width: 10%;")
                sw.WriteLine("}")
                sw.WriteLine("td.title {")
                sw.WriteLine("	width: 50%;")
                sw.WriteLine("	padding-Left: 5px;")
                sw.WriteLine("}")
                sw.WriteLine("td.photographer {")
                sw.WriteLine("	width: 30%;")
                sw.WriteLine("	padding-Left: 5px;")
                sw.WriteLine("}")
                sw.WriteLine("p {")
                sw.WriteLine("	margin: 0px;")
                'sw.WriteLine("	Text-align: center;")
                sw.WriteLine("}")
                sw.WriteLine("p.title {")
                sw.WriteLine("	font-weight: bold;")
                sw.WriteLine("	font-size: 14pt;")
                sw.WriteLine("}")
                sw.WriteLine("p.subtitle {")
                sw.WriteLine("	font-size: 14pt;")
                sw.WriteLine("}")
                sw.WriteLine("-->")
                sw.WriteLine("</style>")

                sw.WriteLine("<body>")
                sw.WriteLine("<table class=""worksheet"">")
                sw.WriteLine(
                    "<tr><td colspan=""4"" class=""header_center""><p class=""title"">" + camera_club_name +
                    " Competition Results</p></td></tr>")
                If _
                    EnableClassification.Checked Or EnableMedium.Checked Or EnableTheme.Checked Or SelectScore.Text <> "All" Or
                    EnableAward.Checked Then
                    sw.Write("<tr><td colspan=""4"" class=""header_center""><p class=""title"">")
                    If EnableClassification.Checked Then
                        sw.Write(SelectClassification.Text + "&nbsp;")
                    End If
                    If EnableMedium.Checked Then
                        sw.Write(SelectMedium.Text + "&nbsp;")
                    End If
                    If SelectScore.Text <> "All" Then
                        sw.Write(SelectScore.Text + "&nbsp;points&nbsp;")
                    End If
                    If EnableAward.Checked Then
                        sw.Write(SelectAward.Text)
                    End If
                    sw.WriteLine("&nbsp;</p></td></tr>")
                End If
                sw.Write(
                    "<tr><td colspan=""4"" class=""header_center""><p class=""subtitle"">" +
                    competition_date.ToString("MMMM", CultureInfo.CurrentCulture) + " " +
                    competition_date.Year.ToString)
                If EnableTheme.Checked Then
                    sw.Write(" - " + SelectTheme.Text)
                End If
                sw.WriteLine("</p></td></tr>")
                If Not display_scores Then
                    sw.WriteLine(
                        "<tr><td colspan=""4"" class=""header_center""><p class=""subtitle"">(" +
                        setMaxAwards(entries.Count).ToString + " Awards)</p></td></tr>")
                End If
                sw.WriteLine("<tr><td colspan=""4"" class=""header_center""><p class=""subtitle"">&nbsp;</p></td></tr>")
                sw.WriteLine("<tr><th>Score</th><th>Award</th>")
                sw.WriteLine("<th>Title</th><th>Photographer</th></tr>")
                '
                ' Print a row for each image in the category
                For Each entry In entries
                    sw.WriteLine("<tr>")
                    If display_scores Then
                        sw.WriteLine("  <td class=""score"">" + entry.Score_1.ToString + "</td>")
                        sw.WriteLine("  <td class=""award"">" + entry.Award + "</td>")
                    Else
                        sw.WriteLine("  <td class=""score"">&nbsp;</td>")
                        sw.WriteLine("  <td class=""award"">&nbsp;</td>")
                    End If
                    sw.WriteLine("  <td class=""title"">" + entry.Title + "</td>")
                    sw.WriteLine("  <td class=""photographer"">" + entry.Maker + "</td>")
                    sw.WriteLine("</tr>")
                Next entry
                '
                ' Finish off the HTML
                sw.WriteLine("</table>")
                sw.WriteLine("</body>")
                sw.WriteLine("</html>")
                sw.Close()
                '
                ' Launch a browser to display the worksheet
                doShellExecute(temp_file)
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub


        Private Shared Sub doShellExecute(file As String)
            Try
                Dim my_process As New Process
                my_process.StartInfo.FileName = file
                my_process.StartInfo.UseShellExecute = True
                my_process.StartInfo.RedirectStandardOutput = False
                my_process.Start()
                my_process.Dispose()
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub getUserPreferences()
            Dim prefs_dialog As New PreferencesDialog(Me)
            Try
                ' Load the current perferences into the dialog
                prefs_dialog.tbDatabaseFileName.Text = _database_file_name
                If Len(images_root_folder) = 2 And Mid(images_root_folder, 2, 1) = ":" Then
                    prefs_dialog.tbImagesRootFolder.Text = images_root_folder + "\"
                Else
                    prefs_dialog.tbImagesRootFolder.Text = images_root_folder
                End If
                If Len(reports_output_folder) = 2 And Mid(reports_output_folder, 2, 1) = ":" Then
                    prefs_dialog.tbReportsOutputFolder.Text = reports_output_folder + "\"
                Else
                    prefs_dialog.tbReportsOutputFolder.Text = reports_output_folder
                End If
                prefs_dialog.tbServerName.Text = _server_name
                prefs_dialog.tbServerScriptDir.Text = _server_script_dir
                _query = From club In rps_context.clubs
                         Select club.id, club.name

                For Each _record In _query
                    prefs_dialog.cbCameraClubName.Items.Add(New Entities.DataItem(_record.id, _record.name))
                Next
                prefs_dialog.cbCameraClubName.Text = camera_club_name
                prefs_dialog.cbNumJudges.Text = CType(num_judges, String)

                ' Display the dialog
                prefs_dialog.ShowDialog()

                ' Check the results
                Dim irf As String = Trim(prefs_dialog.tbImagesRootFolder.Text)
                Dim dbfn As String = Trim(prefs_dialog.tbDatabaseFileName.Text)
                Dim rof As String = Trim(prefs_dialog.tbReportsOutputFolder.Text)
                Dim sn As String = Trim(prefs_dialog.tbServerName.Text)
                Dim ssd As String = Trim(prefs_dialog.tbServerScriptDir.Text)
                Dim ccn As String = Trim(CType(prefs_dialog.cbCameraClubName.SelectedItem, Entities.DataItem).Value)
                Dim ccid As Integer = CType(prefs_dialog.cbCameraClubName.SelectedItem, Entities.DataItem).ID
                Dim nj As Integer = CType(prefs_dialog.cbNumJudges.Text, Integer)
                If prefs_dialog.DialogResult = DialogResult.OK Then
                    If irf > "" Then
                        ' If necessary, strip off a trailing "\"
                        images_root_folder = Helper.trimTrailingSlash(images_root_folder)
                        ' write it to the registry
                        writeRegistryString("Software\RPS Digital Viewer", "Images Root Folder", images_root_folder)
                    End If
                    If dbfn > "" Then
                        ' update the connection string
                        Set_database_name(dbfn)
                        ' write it to the registry
                        writeRegistryString("Software\RPS Digital Viewer", "Database File Name", _database_file_name)
                    End If
                    If rof > "" Then
                        ' If necessary, strip off a trailing "\"
                        reports_output_folder = Helper.trimTrailingSlash(reports_output_folder)

                        ' write it to the registry
                        writeRegistryString("Software\RPS Digital Viewer", "Reports Output Folder", reports_output_folder)
                    End If
                    If sn > "" Then
                        ' Store the new server name in memory
                        _server_name = sn
                        ' Write it to the registry
                        writeRegistryString("Software\RPS Digital Viewer", "Server Name", sn)
                    End If
                    If ssd > "" Then
                        ' Store the new server script directory in memory
                        _server_script_dir = ssd
                        ' Write it to the registry
                        writeRegistryString("Software\RPS Digital Viewer", "Server Script Directory", ssd)
                    End If
                    If ccid > 0 Then
                        ' Set and store the club id
                        camera_club_id = ccid
                        writeRegistryString("Software\RPS Digital Viewer", "Camera Club ID", CType(ccid, String))
                    End If
                    If ccn > "" Then
                        ' Set and store the club name
                        camera_club_name = ccn
                        writeRegistryString("Software\RPS Digital Viewer", "Camera Club Name", ccn)
                    End If
                    If nj > 0 Then
                        ' Set and store the number of judges
                        num_judges = nj
                        writeRegistryString("Software\RPS Digital Viewer", "Number of Judges", CType(nj, String))
                    End If

                    ' Fetch the list of classifications, mediums and awards from the database
                    ' for the selected club
                    getClubRules()

                End If
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub Set_database_name(file_name As String)
            Try
                _database_file_name = file_name
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub writeRegistryString(key As String, registry_name As String, registry_value As String)
            Dim reg_key As RegistryKey = Nothing

            Try
                reg_key = Registry.CurrentUser.OpenSubKey(key, True)
                If reg_key Is Nothing Then
                    reg_key = Registry.CurrentUser.CreateSubKey(key)
                End If
                reg_key.SetValue(registry_name, registry_value)
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            Finally
                If Not reg_key Is Nothing Then
                    reg_key.Close()
                End If
            End Try
        End Sub

        Private Function getRegistryString(registry_key As String, registry_name As String) As String
            Dim reg_key As RegistryKey = Nothing

            Try
                reg_key = Registry.CurrentUser.OpenSubKey(registry_key, False)
                If reg_key Is Nothing Then
                    getRegistryString = ""
                Else
                    getRegistryString = reg_key.GetValue(registry_name, "")
                End If
            Catch exception As Exception
                getRegistryString = ""
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            Finally
                If Not reg_key Is Nothing Then
                    reg_key.Close()
                End If
            End Try
        End Function

        Private Sub getPreferences()
            Dim value As String

            Try
                value = getRegistryString("Software\RPS Digital Viewer", "Images Root Folder")
                If value > "" Then
                    images_root_folder = value
                End If
                value = getRegistryString("Software\RPS Digital Viewer", "Database File Name")
                If value > "" Then
                    Set_database_name(value)
                End If
                value = getRegistryString("Software\RPS Digital Viewer", "Reports Output Folder")
                If value > "" Then
                    reports_output_folder = value
                End If
                value = getRegistryString("Software\RPS Digital Viewer", "Server Name")
                If value > "" Then
                    _server_name = value
                End If
                value = getRegistryString("Software\RPS Digital Viewer", "Server Script Directory")
                If value > "" Then
                    _server_script_dir = value
                End If
                value = getRegistryString("Software\RPS Digital Viewer", "Camera Club ID")
                If value > "" Then
                    camera_club_id = CType(value, Integer)
                End If
                value = getRegistryString("Software\RPS Digital Viewer", "Camera Club Name")
                If value > "" Then
                    camera_club_name = value
                End If
                value = getRegistryString("Software\RPS Digital Viewer", "Last Admin Username")
                If value > "" Then
                    last_admin_username = value
                End If
                value = getRegistryString("Software\RPS Digital Viewer", "Number of Judges")
                If value > "" Then
                    num_judges = value
                End If

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        '
        ' Query the database to retrieve the list of classifications, mediums and awards for the
        ' selected club.  Store these lists in memory.
        '
        Private Sub getClubRules()
            Try
                ' Fetch the list of club classifications from the database
                classifications.Clear()
                SelectClassification.Items.Clear()     ' remove any items in the classifications combobox
                _query = From c In rps_context.classifications
                         From b In rps_context.club_classification
                         From a In rps_context.clubs
                         Where a.id = camera_club_id AndAlso b.classification_id = c.id
                         Select c.name

                For Each _record In _query
                    classifications.Add(_record)
                    SelectClassification.Items.Add(_record)
                Next
                SelectClassification.SelectedIndex = 0 ' Select the first element in the combobox

                ' Fetch the list of club mediums from the database
                mediums.Clear()
                SelectMedium.Items.Clear()     ' remove any items in the mediums combobox
                _query = From c In rps_context.media
                         From b In rps_context.club_medium
                         From a In rps_context.clubs
                         Where a.id = camera_club_id AndAlso b.medium_id = c.id
                         Order By b.sort_key
                         Select c.name

                For Each _record In _query
                    mediums.Add(_record)
                    SelectMedium.Items.Add(_record)
                Next
                SelectMedium.SelectedIndex = 0 ' Select the first element in the combobox

                ' Fetch the list of club awards from the database
                awards.Clear()
                SelectAward.Items.Clear()
                _query = From c In rps_context.awards
                         From b In rps_context.club_award
                         From a In rps_context.clubs
                         Where a.id = camera_club_id AndAlso b.award_id = c.id
                         Select c.name Distinct

                For Each _record In _query
                    awards.Add(_record)
                    SelectAward.Items.Add(_record)
                Next
                SelectAward.SelectedIndex = 0 ' Select the first element in the combobox

                ' Fetch the club's min and max scores from the database
                _record = (From club In rps_context.clubs
                           Where club.id = camera_club_id
                           Select club.max_score, club.min_score, club.min_score_for_award).SingleOrDefault

                min_score = _record.min_score
                max_score = _record.max_score
                min_score_for_award = _record.min_score_for_award
                ' Fill the SelectScore combobox
                SelectScore.Items.Clear()
                SelectScore.Items.Add("All")
                For i As Integer = (max_score * num_judges) To (min_score * num_judges) Step -1
                    SelectScore.Items.Add(CType(i, String))
                Next
                SelectScore.SelectedIndex = 0
                _all_scores_selected = True

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        '
        ' Load the list of unique competition dates into the Competition Dates combobox
        '
        Private Sub setCompetitionDatesCombobox()

            _query = From entry In rps_context.CompetitionEntries
                     Order By entry.Competition_Date_1
                     Select entry.Competition_Date_1 Distinct

            Try
                ' Empty the list if it's not already empty
                If SelectDate.Items.Count > 0 Then
                    SelectDate.Items.Clear()
                End If

                For Each _record In _query
                    Dim item As DateTime
                    item = Convert.ToDateTime(_record)

                    SelectDate.Items.Add(item.ToString("dd-MMM-yyyy"))
                Next
                ' Select the first item in the list
                If SelectDate.Items.Count > 0 Then
                    SelectDate.SelectedIndex = 0
                End If
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub
        ' Call the REST service on the server to retrieve the list of available
        ' competition dates.
        Private Function getRestCompetitionDates(params As Hashtable) As ArrayList
            Dim navigator As XPathNavigator
            Dim response As XPathDocument
            Dim nodes As XPathNodeIterator
            Dim node As XPathNavigator
            Dim dates As New ArrayList

            Try
                ' Retrieve the list of competition dates from the server
                params.Add("rpswinclient", "getcompdate")
                If doRest(_server_name, _server_script_dir, "GET", params, response) Then
                    navigator = response.CreateNavigator()
                    nodes = navigator.Select("/rsp/Competition_Date")
                    While nodes.MoveNext()
                        node = nodes.Current
                        dates.Add(node.Value)
                    End While
                Else
                    getRestCompetitionDates = dates
                    Exit Function
                End If

                getRestCompetitionDates = dates

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
                getRestCompetitionDates = dates
            End Try
        End Function

        Structure CompEntry
            Implements IComparable

            Public entry_id As String
            Public first_name As String
            Public last_name As String
            Public title As String
            Public score As String
            Public award As String
            Public url As String
            Public bucket As Integer

            Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
                If bucket.CompareTo(CType(obj, CompEntry).bucket) = 0 Then
                    Return String.Compare(first_name, CType(obj, CompEntry).first_name, StringComparison.Ordinal)
                Else
                    Return bucket.CompareTo(CType(obj, CompEntry).bucket)
                End If
            End Function
        End Structure

        Private Sub downloadCompetitionImages()

            Dim download_dialog As DownloadCompetitionsDialog
            Dim username As String
            Dim password As String
            Dim comp_date As String
            Dim comp_classification As String = Nothing
            Dim comp_theme As String = Nothing
            Dim comp_medium As String = Nothing
            Dim local_image_file_name As String
            Dim output_folder As String
            Dim competition_folder As String
            Dim dir_info As DirectoryInfo
            Dim params As New Hashtable
            Dim response As XPathDocument
            Dim comp_dates As ArrayList
            Dim navigator As XPathNavigator
            Dim nodes As XPathNodeIterator
            Dim node As XPathNavigator
            Dim comp_nodes As XPathNodeIterator
            Dim comp_node As XPathNavigator
            Dim entry_nodes As XPathNodeIterator
            Dim entry_node As XPathNavigator
            Dim entry_properties As XPathNodeIterator
            Dim entry_property As XPathNavigator
            Dim sql As String
            Dim entries_array_list As New ArrayList
            Dim this_entry As CompEntry = Nothing
            Dim member As String
            Dim prev_member As String
            Dim bucket As Integer
            Dim sequence_num As Integer
            Dim download_digital As Boolean
            Dim download_prints As Boolean

            Try
                ' Retrieve the list of competition dates from the server
                Cursor.Current = Cursors.WaitCursor
                params.Clear()
                params.Add("closed", "Y")
                params.Add("scored", "N")
                comp_dates = getRestCompetitionDates(params)
                If comp_dates.Count = 0 Then
                    MsgBox("No competitions are available for download", , "Download Competition Images")
                    Exit Sub
                End If
                Cursor.Current = Cursors.Default

                ' Open the Download Competitions dialog
                download_dialog = New DownloadCompetitionsDialog(Me, last_admin_username, comp_dates, images_root_folder)
                download_dialog.ShowDialog(Me)
                If download_dialog.DialogResult = DialogResult.OK Then
                    username = Trim(download_dialog.Username.Text())
                    password = Trim(download_dialog.Password.Text())
                    comp_date = download_dialog.CompetitionDate.Text()
                    download_digital = download_dialog.Download_digital.Checked
                    download_prints = download_dialog.Download_prints.Checked
                    output_folder = Trim(download_dialog.OutputFolder.Text())
                Else
                    Exit Sub
                End If

                ' Save the username in the registry as the default admin username
                last_admin_username = username
                writeRegistryString("Software\RPS Digital Viewer", "Last Admin Username", last_admin_username)

                ' Delete any competitions in the local database that already have this date
                Dim d As Date
                d = Date.ParseExact(comp_date, "yyyy-MM-dd", CultureInfo.CurrentCulture)
                sql = "DELETE FROM CompetitionEntries WHERE Competition_Date_1 = '" +
                      Format(d, "M/dd/yyyy") + "'"
                If download_digital And Not download_prints Then
                    sql += " AND Medium like '%Digital'"
                End If
                If download_prints And Not download_digital Then
                    sql += " AND Medium like '%Prints'"
                End If

                _query = rps_context.Database.ExecuteSqlCommand(sql)
                Application.DoEvents()

                ' Retrieve the competition Manifest from the server
                Cursor.Current = Cursors.WaitCursor
                params.Clear()
                params.Add("rpswinclient", "download")
                params.Add("username", username)
                params.Add("password", password)
                params.Add("comp_date", comp_date)
                If download_digital And Not download_prints Then
                    params.Add("medium", "digital")
                End If
                If download_prints And Not download_digital Then
                    params.Add("medium", "prints")
                End If
                If Not doRest(_server_name, _server_script_dir, "POST", params, response) Then
                    navigator = response.CreateNavigator()
                    nodes = navigator.Select("/rsp/err")
                    nodes.MoveNext()
                    node = nodes.Current
                    MsgBox("Download Failure - " + node.GetAttribute("msg", ""), , "Error in DownloadCompetitionImages()")
                    Exit Sub
                End If

                ' Initialize the progress bar
                status_bar.progress_bar.Minimum = 0
                status_bar.progress_bar.Value = 0
                navigator = response.CreateNavigator()
                nodes = navigator.Select("/descendant::*[name()='Image_URL']") ' Count the images in the manifest
                status_bar.progress_bar.Maximum = nodes.Count

                ' Select the list of Competitions from the manifest
                'navigator = response.CreateNavigator()
                nodes = navigator.Select("/rsp/Competitions/Competition")

                ' Iterate through the list of competitions
                While nodes.MoveNext
                    node = nodes.Current

                    ' Select the subnodes of this competition
                    comp_nodes = node.Select("./*")
                    ' Parse out the properties of this competition from its subnodes
                    While comp_nodes.MoveNext
                        comp_node = comp_nodes.Current
                        Select Case comp_node.Name
                            Case "Date"
                                comp_date = comp_node.Value
                            Case "Theme"
                                comp_theme = comp_node.Value
                            Case "Medium"
                                comp_medium = comp_node.Value
                            Case "Classification"
                                comp_classification = comp_node.Value
                            Case "Entries"
                                ' Select all the child <Entry> nodes of the <Entries> nodes
                                entry_nodes = comp_node.Select("./*")
                        End Select
                    End While

                    ' Now iterate through the <Entry> nodes of this competition
                    prev_member = ""
                    bucket = 0
                    entries_array_list.Clear()
                    While entry_nodes.MoveNext
                        entry_node = entry_nodes.Current
                        ' Select all the properties (i.e. subnodes) of this entry
                        entry_properties = entry_node.Select("./*")
                        ' Parse out the properties of this <Entry> from its subnodes
                        While entry_properties.MoveNext
                            entry_property = entry_properties.Current
                            Select Case entry_property.Name
                                Case "ID"
                                    'entry_id = entry_property.Value
                                    this_entry.entry_id = entry_property.Value
                                Case "First_Name"
                                    'first_name = entry_property.Value
                                    this_entry.first_name = entry_property.Value
                                Case "Last_Name"
                                    'last_name = entry_property.Value
                                    this_entry.last_name = entry_property.Value
                                Case "Title"
                                    'title = entry_property.Value
                                    this_entry.title = entry_property.Value
                                Case "Score"
                                    'score = entry_property.Value
                                    this_entry.score = entry_property.Value
                                Case "Award"
                                    'award = entry_property.Value
                                    this_entry.award = entry_property.Value
                                Case "Image_URL"
                                    'url = entry_property.Value
                                    this_entry.url = entry_property.Value
                            End Select
                        End While

                        ' Assign this member's entry to a bucket or pile
                        ' This assumes that the list of entries is sorted by maker
                        member = this_entry.first_name + " " + this_entry.last_name
                        If member <> prev_member Then
                            bucket = 0
                        Else
                            bucket += 1
                        End If
                        this_entry.bucket = bucket  ' store the bucket number in the record
                        entries_array_list.Add(this_entry)     ' Store the record in a list
                        prev_member = member

                    End While

                    ' Sort the list of entries by bucket
                    entries_array_list.Sort()

                    ' Iterate through all the entries and download the images
                    ' and update the database
                    sequence_num = 1
                    status_bar.progress_bar.Show()
                    For Each entry As CompEntry In entries_array_list
                        ' If necessary, create the folder for this entry
                        competition_folder = output_folder + "\" + comp_date + " " + comp_classification + " " + comp_medium
                        dir_info = New DirectoryInfo(competition_folder)
                        If Not dir_info.Exists Then
                            dir_info.Create()
                        End If

                        ' Fetch the image file from the server
                        local_image_file_name = competition_folder + "\" +
                                                handleStrMap(entry.title, " ?[]/\=+<>:;"",*|", "_---------------") +
                                                "+" + entry.first_name + "_" + entry.last_name + ".jpg"
                        My.Computer.Network.DownloadFile(address:=entry.url,
                                                         destinationFileName:=local_image_file_name,
                                                         userName:=String.Empty,
                                                         password:=String.Empty,
                                                         showUI:=False,
                                                         connectionTimeout:=100000,
                                                         overwrite:=True)

                        ' Insert this image into the database
                        addImageToDatabase(
                            New FileInfo(local_image_file_name),
                            entry.first_name + " " + entry.last_name,
                            entry.title,
                            entry.score,
                            entry.award,
                            comp_classification,
                            comp_medium,
                            comp_date,
                            comp_theme,
                            entry.entry_id,
                            sequence_num)

                        ' Update the Progressbar
                        status_bar.progress_bar.Value = status_bar.progress_bar.Value + 1
                        Application.DoEvents()

                        sequence_num += 1

                    Next

                End While

                ' Update the list of dates in the Competition Date combobox
                setCompetitionDatesCombobox()

                MsgBox("Competition Images Downloaded Successfully", , "Download Competition Images")

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            Finally
                ' Clear the ProgressBar
                status_bar.progress_bar.Hide()
                status_bar.progress_bar.Value = 0
                Cursor.Current = Cursors.Default
            End Try
        End Sub


        Private Function doRest(server As String,
                                operation As String,
                                http_method As String,
                                params As Hashtable,
                                ByRef results As XPathDocument) As Boolean
            Dim request As HttpWebRequest
            Dim response As HttpWebResponse = Nothing
            Dim url As String
            Dim delim As String
            Dim x_path_doc As XPathDocument
            Dim data As New StringBuilder
            Dim byte_data() As Byte
            Dim post_stream As Stream = Nothing
            Dim fs As FileStream
            Dim br As BinaryReader
            Dim ms As New MemoryStream

            Try
                ' Build the URL
                url = "http://" + server + operation
                If http_method = "GET" Then
                    delim = "?"
                    For Each key As String In params.Keys
                        url = url + delim + key + "=" + params.Item(key)
                        delim = "&"
                    Next
                End If

                ' Create the web request  
                request = HttpWebRequest.Create(url)
                If http_method = "POST" Then
                    ' Set type to POST  
                    request.Method = "POST"
                    'request.ContentType = "application/x-www-form-urlencoded"
                    request.ContentType = "multipart/form-data, boundary=AaB03x"
                    ' Build the body of the POST transaction

                    'delim = ""
                    data.Append("--" + "AaB03x" + vbCrLf)
                    For Each param As String In params.Keys
                        If param <> "file" Then
                            'data.Append(delim + param + "=" + HttpUtility.UrlEncode(params(param)))
                            'delim = "&"
                            data.Append("Content-Disposition: form-data; name=""" + param + """" + vbCrLf)
                            data.Append(vbCrLf)
                            data.Append(HttpUtility.UrlEncode(params(param)) + vbCrLf)
                            data.Append("--" + "AaB03x" + vbCrLf)
                        End If
                    Next

                    ' Write the POST header, thus far, to a binary stream
                    byte_data = UTF8Encoding.UTF8.GetBytes(data.ToString())
                    ms.Write(byte_data, 0, byte_data.Length)
                    data.Remove(0, data.Length)

                    ' Attach any files in the param list
                    For Each param As String In From param1 As String In params.Keys Where param1 = "file"
                        data.Append(
                            "Content-Disposition: form-data; name=""file""; filename=""" + params(param) + """" + vbCrLf)
                        data.Append("Content-Transfer-Encoding: binary" + vbCrLf)
                        data.Append("Content-Type: Image/jpeg" + vbCrLf)
                        data.Append(vbCrLf)
                        ' Write the MIME header to a binary stream
                        byte_data = UTF8Encoding.UTF8.GetBytes(data.ToString())
                        ms.Write(byte_data, 0, byte_data.Length)
                        ' Open the file and write it to a binary stream
                        fs = Nothing
                        Try
                            fs = New FileStream(params(param), FileMode.Open, FileAccess.Read)
                            br = New BinaryReader(fs)
                            byte_data = br.ReadBytes(fs.Length)
                            ms.Write(byte_data, 0, fs.Length)
                            br.Close()
                            ' Write the terminating boundry marker
                            data.Remove(0, data.Length)
                            data.Append("--" + "AaB03x" + vbCrLf)
                            byte_data = UTF8Encoding.UTF8.GetBytes(data.ToString())
                            ms.Write(byte_data, 0, byte_data.Length)
                        Finally
                            If fs IsNot Nothing Then
                                fs.Close()
                            End If
                        End Try
                    Next

                    ' Read the entire contents of the memory stream back into a byte array
                    byte_data = ms.ToArray
                    'MsgBox(UTF8Encoding.UTF8.GetString(byteData))

                    ' Set the content length in the request headers
                    request.ContentLength = byte_data.Length

                    ' Write data
                    Try
                        post_stream = request.GetRequestStream()
                        post_stream.Write(byte_data, 0, byte_data.Length)
                    Finally
                        If Not post_stream Is Nothing Then post_stream.Close()
                    End Try
                End If

                ' Get response
                response = request.GetResponse()

                ' Debug: Read the HTTP response into a string and display it
                'If http_method = "POST" Then
                'reader = New StreamReader(response.GetResponseStream())
                'Dim dummy As String
                'dummy = reader.ReadToEnd
                'MsgBox(dummy)
                'End If

                ' Parse the status out of the xml response
                Dim response_stream As Stream = response.GetResponseStream()
                x_path_doc = New XPathDocument(response_stream)
                If getRestStatResponse(x_path_doc) Then
                    doRest = True
                Else
                    doRest = False
                End If
                results = x_path_doc

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
                doRest = False
            Finally
                If Not response Is Nothing Then response.Close()
            End Try
        End Function

        Private Function getRestStatResponse(response As XPathDocument) As Boolean
            Dim navigator As XPathNavigator
            Dim nodes As XPathNodeIterator
            Dim node As XPathNavigator

            Try
                ' Get the response status node with XPath
                navigator = response.CreateNavigator()
                nodes = navigator.Select("/rsp[@stat]")
                nodes.MoveNext()
                node = nodes.Current
                If node.GetAttribute("stat", "") = "ok" Then
                    getRestStatResponse = True
                Else
                    getRestStatResponse = False
                End If

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Function
        ' Map characters within a string.  Replace all occurrances of each character in
        ' MapWhat with the corresponding character in toWhat.
        '
        ' Call it like so
        '       MyString = StrMap(MyString, "àéù", "aeu")
        '
        ' If the MapWhat string is longer than the ToWhat string, the extra characters
        ' in MapWhat will be deleted from the string e.g.
        '       MyString = StrMap(MyString, "àéùxyz", "aeu")
        ' will translate all occurrances of à, é, ù and will delete all occurrances of x,y,z.

        Public Function handleStrMap(original_string As String,
                                     map_what As String,
                                     to_what As String,
                                     Optional ByVal compare As Long = 0) As String

            Dim output_ptr As Long
            Dim input_ptr As Long
            Dim c As String
            Dim map_index As Long

            Try
                For input_ptr = 1 To Len(original_string)
                    ' get the next char in the source string
                    c = Mid$(original_string, input_ptr, 1)
                    ' is this a character to be mapped?
                    map_index = InStr(1, map_what, c, compare)
                    ' No. If the output string is shorter than the input string
                    ' (because previous characters have been deleted) relocate this character
                    ' to the end of the output string.
                    If map_index = 0 Then
                        output_ptr = output_ptr + 1
                        If output_ptr < input_ptr Then
                            Mid$(original_string, output_ptr) = c
                        End If
                    ElseIf map_index <= Len(to_what) Then
                        ' Yes, this character is to be remapped (or deleted).  If there is a
                        ' corresponding character to map it to, replace the original character,
                        ' otherwise do nothing.  Doing nothing will effectively delete the
                        ' character from the output string.
                        output_ptr = output_ptr + 1
                        Mid$(original_string, output_ptr) = Mid$(to_what, map_index, 1)
                    End If
                Next
                ' If the output string is the same length as the input string (i.e. there have
                ' been no deletions) return the original string length.  If there have been 
                ' deletions, return the shortened string.
                If output_ptr = Len(original_string) Then
                    handleStrMap = original_string
                Else
                    handleStrMap = Microsoft.VisualBasic.Left(original_string, output_ptr)
                End If
            Catch exception As Exception
                handleStrMap = original_string
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Function

        Function uploadScores() As Boolean
            Dim upload_dialog As UploadScoresDialog
            Dim username As String
            Dim password As String
            Dim comp_date As String
            Dim comp_dates As ArrayList
            Dim comp_class_list As New ArrayList
            Dim comp_medium_list As New ArrayList
            Dim sql_select As String
            Dim sql_where As String
            Dim recs As Object
            Dim comp_num As Integer
            Dim classification As String
            Dim medium As String
            Dim maker As String
            Dim title As String
            Dim score As String
            Dim award As String
            Dim entry_id As String
            Dim sw As StreamWriter
            Dim file_name As String
            Dim params As New Hashtable
            Dim response As XPathDocument
            Dim navigator As XPathNavigator
            Dim nodes As XPathNodeIterator
            Dim node As XPathNavigator
            Dim info As New StringBuilder
            Dim posn As Integer
            Dim first_name As String
            Dim last_name As String
            Dim upload_digital As Boolean
            Dim upload_prints As Boolean
            Dim selected_medium As String

            Try
                ' Get the list of competition dates from the server
                Cursor.Current = Cursors.WaitCursor
                params.Add("closed", "Y")
                params.Add("scored", "N")
                comp_dates = getRestCompetitionDates(params)
                If comp_dates.Count = 0 Then
                    MsgBox("No competitions available to receive scores", , "Upload Scores")
                    Return False
                End If
                Cursor.Current = Cursors.Default

                ' Open the Upload Scores dialog
                upload_dialog = New UploadScoresDialog(last_admin_username, comp_dates)
                upload_dialog.ShowDialog(Me)
                If upload_dialog.DialogResult = DialogResult.OK Then
                    username = Trim(upload_dialog.Username.Text())
                    password = Trim(upload_dialog.Password.Text())
                    comp_date = upload_dialog.CompDate.Text()
                    upload_digital = upload_dialog.Upload_digital_scores.Checked
                    upload_prints = upload_dialog.Upload_print_scores.Checked
                Else
                    Exit Function
                End If

                ' Save the username in the registry as the default admin username
                last_admin_username = username
                writeRegistryString("Software\RPS Digital Viewer", "Last Admin Username", last_admin_username)

                Cursor.Current = Cursors.WaitCursor
                Application.DoEvents()

                ' Open a local text file to receive the XML
                file_name = reports_output_folder + "\" + "Scores_" + comp_date + ".xml"
                sw = File.CreateText(file_name)
                sw.WriteLine("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                sw.WriteLine("<Competitions>")

                ' Select the unique competitions for the given date
                selected_medium = ""
                If upload_digital And Not upload_prints Then
                    selected_medium = " AND Medium like '%Digital'"
                End If
                If upload_prints And Not upload_digital Then
                    selected_medium = " AND Medium like '%Prints'"
                End If
                Dim d As Date = Date.ParseExact(comp_date, "yyyy-MM-dd", CultureInfo.CurrentCulture)
                sql_select = "SELECT DISTINCT Classification, Medium FROM CompetitionEntries "
                sql_where = "WHERE Competition_Date_1 = '" +
                            Format(d, "M/dd/yyyy") + "'" +
                            selected_medium
                _query = sql_select + sql_where
                recs = rps_context.Database.SqlQuery(Of Entities.ClassificationMedium)(_query).ToList

                For Each record As Entities.ClassificationMedium In recs
                    comp_class_list.Add(record.Classification)
                    comp_medium_list.Add(record.Medium)
                Next

                ' Iterate through all the competition for this date
                For comp_num = 0 To comp_class_list.Count - 1
                    ' Query the database for the entries of this competition
                    classification = comp_class_list.Item(comp_num)
                    medium = comp_medium_list.Item(comp_num)
                    sql_select = "Select * FROM CompetitionEntries "
                    sql_where = "WHERE Competition_Date_1 = '" + Format(d, "M/dd/yyyy") + "' And classification = '" +
                                classification + "' AND Medium = '" + medium + "'"
                    _query = sql_select + sql_where
                    recs = rps_context.Database.SqlQuery(Of Entities.CompetitionEntry)(_query).ToList
                    ' Output the tags that describe this competition
                    sw.WriteLine("  <Competition>")
                    sw.WriteLine("    <Date>{0}</Date>", HttpUtility.HtmlEncode(comp_date))
                    sw.WriteLine("    <Classification>{0}</Classification>", HttpUtility.HtmlEncode(classification))
                    sw.WriteLine("    <Medium>{0}</Medium>", HttpUtility.HtmlEncode(medium))
                    sw.WriteLine("    <Entries>")
                    ' Iterate through all the entries of this competition
                    For Each record As Entities.CompetitionEntry In recs
                        ' Read the entry data from the database
                        maker = record.Maker
                        'fields = Split(maker, " ")
                        posn = InStr(1, maker, " ")
                        first_name = Mid(maker, 1, posn - 1)
                        last_name = Mid(maker, posn + 1)
                        title = record.Title
                        If IsNothing(record.Score_1) Then
                            score = ""
                        Else
                            score = record.Score_1.ToString()
                        End If
                        If IsNothing(record.Award) Then
                            award = ""
                        Else
                            award = record.Award
                        End If
                        If IsNothing(record.Server_Entry_ID) Then
                            entry_id = ""
                        Else
                            entry_id = record.Server_Entry_ID
                        End If
                        ' Write this entry to the xml file
                        sw.WriteLine("      <Entry>")
                        sw.WriteLine("        <ID>{0}</ID>", entry_id)
                        sw.WriteLine("        <First_Name>{0}</First_Name>", HttpUtility.HtmlEncode(first_name))
                        sw.WriteLine("        <Last_Name>{0}</Last_Name>", HttpUtility.HtmlEncode(last_name))
                        sw.WriteLine("        <Title>{0}</Title>", HttpUtility.HtmlEncode(title))
                        sw.WriteLine("        <Score>{0}</Score>", HttpUtility.HtmlEncode(score))
                        sw.WriteLine("        <Award>{0}</Award>", HttpUtility.HtmlEncode(award))
                        sw.WriteLine("      </Entry>")
                    Next
                    ' Close out this competition
                    sw.WriteLine("    </Entries>")
                    sw.WriteLine("  </Competition>")
                Next
                ' Close out the xml file
                sw.WriteLine("</Competitions>")
                sw.Close()

                ' Call the web service to upload the xml file to the server
                Application.DoEvents()
                params.Clear()
                params.Add("date", comp_date)
                params.Add("username", username)
                params.Add("password", password)
                params.Add("file", file_name)
                If Not doRest(_server_name, _server_script_dir + "/?rpswinclient=uploadscore", "POST", params, response) _
                    Then
                    ' If the web service returned an error, display it
                    navigator = response.CreateNavigator()
                    nodes = navigator.Select("/rsp/err")
                    nodes.MoveNext()
                    node = nodes.Current
                    Cursor.Current = Cursors.Default
                    MsgBox("Upload Failure - " + node.GetAttribute("msg", ""), , "Error in UploadScores()")
                    Exit Function
                Else
                    ' Display the results of the web service (may include warnings)
                    navigator = response.CreateNavigator()
                    nodes = navigator.Select("/rsp/info")
                    While nodes.MoveNext()
                        node = nodes.Current
                        info.Append(node.Value + vbCrLf)
                    End While
                    If info.Length > 0 Then
                        Cursor.Current = Cursors.Default
                        MsgBox(info.ToString, , "Upload Scores")
                    End If
                End If

                ' Finally, Delete the .xml file
                If File.Exists(file_name) = True Then
                    File.Delete(file_name)
                End If

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Function
        '
        ' Enter the unique list of competition themes for this date into the Theme combobox
        '
        Private Sub setThemeCombobox()
            Try
                Dim comp_date As String = Format(parseSelectedDate(SelectDate.Text), "M/dd/yyyy")
                themes.Clear()
                SelectTheme.Items.Clear()

                _query = From entry In rps_context.CompetitionEntries
                         Where entry.Competition_Date_1.Equals(comp_date)
                         Select entry.Theme Distinct

                For Each _record In _query
                    SelectTheme.Items.Add(_record)
                Next
                If SelectTheme.Items.Count > 0 Then
                    SelectTheme.SelectedIndex = 0
                Else
                    SelectTheme.Text = ""
                End If
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Sub NumNinesHeadingButton_Click(sender As Object, e As EventArgs) _
            Handles NumNinesHeadingButton.Click
            'NineScoreRadioButton.Checked = True
            _all_scores_selected = False
            _eights_and_awards_selected = False
            _selected_avg_score = 9
            _selected_score = 0
            getSelectedEntries()
            doPickAwards()
        End Sub

        Private Sub NumEightsHeadingButton_Click(sender As Object, e As EventArgs) _
            Handles NumEightsHeadingButton.Click
            'EightScoreRadioButton.Checked = True
            _all_scores_selected = False
            _eights_and_awards_selected = False
            _selected_avg_score = 8
            _selected_score = 0
            getSelectedEntries()
            doPickAwards()
        End Sub

        Private Sub NumSevensHeadingButton_Click(sender As Object, e As EventArgs) _
            Handles NumSevensHeadingButton.Click
            'SevenScoreRadioButton.Checked = True
            _all_scores_selected = False
            _eights_and_awards_selected = False
            _selected_avg_score = 7
            _selected_score = 0
            getSelectedEntries()
            doPickAwards()
        End Sub

        Private Sub AwardsTableTitleBar_Click(sender As Object, e As EventArgs) _
            Handles AwardsTableTitleBar.Click
            'EightsAndAwardsRadioButton.Checked = True
            _eights_and_awards_selected = True
            _all_scores_selected = False
            getSelectedEntries()
            doSlideShow()
        End Sub

        Private Sub SelectScore_SelectedIndexChanged(sender As Object, e As EventArgs) _
            Handles SelectScore.SelectedIndexChanged
            If SelectScore.SelectedIndex = 0 Then
                _all_scores_selected = True
                _eights_and_awards_selected = False
            Else
                _all_scores_selected = False
                _eights_and_awards_selected = False
                _selected_score = CType(SelectScore.SelectedItem(), Integer)
            End If
            _selected_avg_score = 0
            getSelectedEntries()
        End Sub

        Private Sub SelectDate_SelectedIndexChanged(sender As Object, e As EventArgs) _
            Handles SelectDate.SelectedIndexChanged
            setThemeCombobox()
            SelectScore.SelectedItem = "All"
            _all_scores_selected = True
            _eights_and_awards_selected = False
            getSelectedEntries()
        End Sub
    End Class
End Namespace
