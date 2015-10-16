Imports System.Collections.Generic
Imports System.Data.Entity.Core.EntityClient
Imports System.Data.SQLite
Imports System.Globalization
Imports System.Linq

Public Class MainForm
    Inherits Form

    Private ReadOnly _data_folder As String = My.Computer.FileSystem.GetParentPath(Application.LocalUserAppDataPath)
    Private _database_file_name As String = _data_folder + "\rps.db"

    ' Database 
    Dim _connection_string As String
    Public rps_context As rpsEntities
    Private _query As Object
    Private _record As Object
    Public entries As IList(Of CompetitionEntry)

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
    Friend WithEvents score As DataGridViewTextBoxColumn
    Friend WithEvents award As DataGridViewTextBoxColumn
    Friend WithEvents title As DataGridViewTextBoxColumn
    Private ReadOnly _center_cell_style As New DataGridViewCellStyle

    Private Sub initializeStatusBar()
        Dim info As StatusBarPanel = New StatusBarPanel
        Dim progress As StatusBarPanel = New StatusBarPanel

        'info.Text = "Ready"
        'info.Width = 592
        progress.Width = 200
        info.Width = Me.Width - progress.Width

        'progress.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        info.AutoSize = StatusBarPanelAutoSize.Spring

        With status_bar
            .Panels.Add(info)
            .Panels.Add(progress)
            .ShowPanels = True
            .SizingGrip = False
            .setProgressBar = 1
            .progressBar.Minimum = 0
            .progressBar.Maximum = 100
        End With

        Me.Controls.Add(status_bar)
    End Sub

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        Application.EnableVisualStyles()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
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
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnCancelAll As System.Windows.Forms.Button
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
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle =
                New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle =
                New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager =
                New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnCancelAll = New System.Windows.Forms.Button()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.FileMenu = New System.Windows.Forms.MenuItem()
        Me.FilePreferencesMenu = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.FileExitMenu = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.CompCatalogImagesDownload = New System.Windows.Forms.MenuItem()
        Me.CompUploadScores = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
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
        Me.score = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.award = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.title = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.grid_caption = New System.Windows.Forms.Label()
        CType(Me.data_grid_entries_view, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnLoad
        '
        Me.btnLoad.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),
                                  System.Windows.Forms.AnchorStyles)
        Me.btnLoad.Location = New System.Drawing.Point(701, 471)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 24)
        Me.btnLoad.TabIndex = 0
        Me.btnLoad.Text = "&Load"
        '
        'btnUpdate
        '
        Me.btnUpdate.Anchor =
            CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),
                  System.Windows.Forms.AnchorStyles)
        Me.btnUpdate.Location = New System.Drawing.Point(783, 471)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 24)
        Me.btnUpdate.TabIndex = 1
        Me.btnUpdate.Text = "&Update"
        '
        'btnCancelAll
        '
        Me.btnCancelAll.Anchor =
            CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),
                  System.Windows.Forms.AnchorStyles)
        Me.btnCancelAll.Location = New System.Drawing.Point(865, 471)
        Me.btnCancelAll.Name = "btnCancelAll"
        Me.btnCancelAll.Size = New System.Drawing.Size(75, 24)
        Me.btnCancelAll.TabIndex = 2
        Me.btnCancelAll.Text = "Ca&ncel All"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(
            New System.Windows.Forms.MenuItem() _
                                           {Me.FileMenu, Me.MenuItem5, Me.MenuItem2, Me.ReportsMenu, Me.MenuItem7})
        '
        'FileMenu
        '
        Me.FileMenu.Index = 0
        Me.FileMenu.MenuItems.AddRange(
            New System.Windows.Forms.MenuItem() _
                                          {Me.FilePreferencesMenu, Me.MenuItem3, Me.FileExitMenu})
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
        Me.MenuItem5.MenuItems.AddRange(
            New System.Windows.Forms.MenuItem() _
                                           {Me.MenuItem1, Me.CompCatalogImagesDownload, Me.CompUploadScores,
                                            Me.MenuItem6})
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
        'MenuItem6
        '
        Me.MenuItem6.Index = 3
        Me.MenuItem6.Text = "-"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.MenuItems.AddRange(
            New System.Windows.Forms.MenuItem() _
                                           {Me.ViewSlideShowMenu, Me.ViewThumbnailsMenu})
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
        Me.ReportsMenu.MenuItems.AddRange(
            New System.Windows.Forms.MenuItem() _
                                             {Me.ReportsScoreSheetMenu, Me.ReportsResultsReportMenu})
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
        Me.SelectClassification.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                               10.0!,
                                                               System.Drawing.FontStyle.Regular,
                                                               System.Drawing.GraphicsUnit.Point,
                                                               CType(0, Byte))
        Me.SelectClassification.Location = New System.Drawing.Point(28, 136)
        Me.SelectClassification.Name = "SelectClassification"
        Me.SelectClassification.Size = New System.Drawing.Size(174, 24)
        Me.SelectClassification.TabIndex = 20
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                 10.0!,
                                                 System.Drawing.FontStyle.Regular,
                                                 System.Drawing.GraphicsUnit.Point,
                                                 CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(28, 114)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(168, 16)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Classification"
        '
        'SelectMedium
        '
        Me.SelectMedium.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                       10.0!,
                                                       System.Drawing.FontStyle.Regular,
                                                       System.Drawing.GraphicsUnit.Point,
                                                       CType(0, Byte))
        Me.SelectMedium.Location = New System.Drawing.Point(28, 188)
        Me.SelectMedium.Name = "SelectMedium"
        Me.SelectMedium.Size = New System.Drawing.Size(174, 24)
        Me.SelectMedium.TabIndex = 18
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                 10.0!,
                                                 System.Drawing.FontStyle.Regular,
                                                 System.Drawing.GraphicsUnit.Point,
                                                 CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(28, 166)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(175, 16)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Medium"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                 10.0!,
                                                 System.Drawing.FontStyle.Regular,
                                                 System.Drawing.GraphicsUnit.Point,
                                                 CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(28, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(175, 16)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Competition Date"
        '
        'RecalcAwards
        '
        Me.RecalcAwards.Location = New System.Drawing.Point(51, 419)
        Me.RecalcAwards.Name = "RecalcAwards"
        Me.RecalcAwards.Size = New System.Drawing.Size(129, 23)
        Me.RecalcAwards.TabIndex = 25
        Me.RecalcAwards.Text = "Recalc &Awards"
        '
        'AwardsTableTitleBar
        '
        Me.AwardsTableTitleBar.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.AwardsTableTitleBar.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                              10.0!,
                                                              System.Drawing.FontStyle.Bold,
                                                              System.Drawing.GraphicsUnit.Point,
                                                              CType(0, Byte))
        Me.AwardsTableTitleBar.ForeColor = System.Drawing.Color.White
        Me.AwardsTableTitleBar.Location = New System.Drawing.Point(28, 336)
        Me.AwardsTableTitleBar.Margin = New System.Windows.Forms.Padding(0)
        Me.AwardsTableTitleBar.Name = "AwardsTableTitleBar"
        Me.AwardsTableTitleBar.Size = New System.Drawing.Size(175, 24)
        Me.AwardsTableTitleBar.TabIndex = 27
        Me.AwardsTableTitleBar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'NumNinesHeadingButton
        '
        Me.NumNinesHeadingButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.NumNinesHeadingButton.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                                10.0!,
                                                                System.Drawing.FontStyle.Regular,
                                                                System.Drawing.GraphicsUnit.Point,
                                                                CType(0, Byte))
        Me.NumNinesHeadingButton.Location = New System.Drawing.Point(28, 360)
        Me.NumNinesHeadingButton.Margin = New System.Windows.Forms.Padding(0)
        Me.NumNinesHeadingButton.Name = "NumNinesHeadingButton"
        Me.NumNinesHeadingButton.Size = New System.Drawing.Size(59, 24)
        Me.NumNinesHeadingButton.TabIndex = 34
        Me.NumNinesHeadingButton.Text = "9s"
        '
        'NumEightsHeadingButton
        '
        Me.NumEightsHeadingButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.NumEightsHeadingButton.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                                 10.0!,
                                                                 System.Drawing.FontStyle.Regular,
                                                                 System.Drawing.GraphicsUnit.Point,
                                                                 CType(0, Byte))
        Me.NumEightsHeadingButton.Location = New System.Drawing.Point(86, 360)
        Me.NumEightsHeadingButton.Margin = New System.Windows.Forms.Padding(0)
        Me.NumEightsHeadingButton.Name = "NumEightsHeadingButton"
        Me.NumEightsHeadingButton.Size = New System.Drawing.Size(59, 24)
        Me.NumEightsHeadingButton.TabIndex = 35
        Me.NumEightsHeadingButton.Text = "8s"
        '
        'NumSevensHeadingButton
        '
        Me.NumSevensHeadingButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.NumSevensHeadingButton.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                                 10.0!,
                                                                 System.Drawing.FontStyle.Regular,
                                                                 System.Drawing.GraphicsUnit.Point,
                                                                 CType(0, Byte))
        Me.NumSevensHeadingButton.Location = New System.Drawing.Point(144, 360)
        Me.NumSevensHeadingButton.Margin = New System.Windows.Forms.Padding(0)
        Me.NumSevensHeadingButton.Name = "NumSevensHeadingButton"
        Me.NumSevensHeadingButton.Size = New System.Drawing.Size(59, 24)
        Me.NumSevensHeadingButton.TabIndex = 36
        Me.NumSevensHeadingButton.Text = "7s"
        '
        'tbEligibleNines
        '
        Me.tbEligibleNines.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbEligibleNines.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                          10.0!,
                                                          System.Drawing.FontStyle.Bold,
                                                          System.Drawing.GraphicsUnit.Point,
                                                          CType(0, Byte))
        Me.tbEligibleNines.Location = New System.Drawing.Point(28, 383)
        Me.tbEligibleNines.Margin = New System.Windows.Forms.Padding(0)
        Me.tbEligibleNines.Name = "tbEligibleNines"
        Me.tbEligibleNines.Size = New System.Drawing.Size(59, 23)
        Me.tbEligibleNines.TabIndex = 37
        Me.tbEligibleNines.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbEligibleEights
        '
        Me.tbEligibleEights.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbEligibleEights.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                           10.0!,
                                                           System.Drawing.FontStyle.Bold,
                                                           System.Drawing.GraphicsUnit.Point,
                                                           CType(0, Byte))
        Me.tbEligibleEights.Location = New System.Drawing.Point(86, 383)
        Me.tbEligibleEights.Margin = New System.Windows.Forms.Padding(0)
        Me.tbEligibleEights.Name = "tbEligibleEights"
        Me.tbEligibleEights.Size = New System.Drawing.Size(59, 23)
        Me.tbEligibleEights.TabIndex = 38
        Me.tbEligibleEights.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbEligibleSevens
        '
        Me.tbEligibleSevens.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbEligibleSevens.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                           10.0!,
                                                           System.Drawing.FontStyle.Bold,
                                                           System.Drawing.GraphicsUnit.Point,
                                                           CType(0, Byte))
        Me.tbEligibleSevens.Location = New System.Drawing.Point(144, 383)
        Me.tbEligibleSevens.Margin = New System.Windows.Forms.Padding(0)
        Me.tbEligibleSevens.Name = "tbEligibleSevens"
        Me.tbEligibleSevens.Size = New System.Drawing.Size(59, 23)
        Me.tbEligibleSevens.TabIndex = 39
        Me.tbEligibleSevens.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'SelectAward
        '
        Me.SelectAward.Enabled = False
        Me.SelectAward.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                      10.0!,
                                                      System.Drawing.FontStyle.Regular,
                                                      System.Drawing.GraphicsUnit.Point,
                                                      CType(0, Byte))
        Me.SelectAward.Location = New System.Drawing.Point(28, 292)
        Me.SelectAward.Name = "SelectAward"
        Me.SelectAward.Size = New System.Drawing.Size(175, 24)
        Me.SelectAward.TabIndex = 40
        '
        'Label4
        '
        Me.Label4.Enabled = False
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                 10.0!,
                                                 System.Drawing.FontStyle.Regular,
                                                 System.Drawing.GraphicsUnit.Point,
                                                 CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(28, 270)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 16)
        Me.Label4.TabIndex = 41
        Me.Label4.Text = "Award"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                 10.0!,
                                                 System.Drawing.FontStyle.Regular,
                                                 System.Drawing.GraphicsUnit.Point,
                                                 CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(28, 62)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(175, 16)
        Me.Label5.TabIndex = 42
        Me.Label5.Text = "Theme"
        '
        'SelectTheme
        '
        Me.SelectTheme.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                      10.0!,
                                                      System.Drawing.FontStyle.Regular,
                                                      System.Drawing.GraphicsUnit.Point,
                                                      CType(0, Byte))
        Me.SelectTheme.Location = New System.Drawing.Point(28, 84)
        Me.SelectTheme.Name = "SelectTheme"
        Me.SelectTheme.Size = New System.Drawing.Size(174, 24)
        Me.SelectTheme.TabIndex = 43
        '
        'EnableClassification
        '
        Me.EnableClassification.Checked = True
        Me.EnableClassification.CheckState = System.Windows.Forms.CheckState.Checked
        Me.EnableClassification.Location = New System.Drawing.Point(12, 139)
        Me.EnableClassification.Name = "EnableClassification"
        Me.EnableClassification.Size = New System.Drawing.Size(13, 21)
        Me.EnableClassification.TabIndex = 45
        '
        'EnableMedium
        '
        Me.EnableMedium.Checked = True
        Me.EnableMedium.CheckState = System.Windows.Forms.CheckState.Checked
        Me.EnableMedium.Location = New System.Drawing.Point(12, 191)
        Me.EnableMedium.Name = "EnableMedium"
        Me.EnableMedium.Size = New System.Drawing.Size(13, 21)
        Me.EnableMedium.TabIndex = 46
        '
        'EnableTheme
        '
        Me.EnableTheme.Checked = True
        Me.EnableTheme.CheckState = System.Windows.Forms.CheckState.Checked
        Me.EnableTheme.Location = New System.Drawing.Point(12, 87)
        Me.EnableTheme.Name = "EnableTheme"
        Me.EnableTheme.Size = New System.Drawing.Size(13, 21)
        Me.EnableTheme.TabIndex = 47
        '
        'EnableAward
        '
        Me.EnableAward.Location = New System.Drawing.Point(12, 295)
        Me.EnableAward.Name = "EnableAward"
        Me.EnableAward.Size = New System.Drawing.Size(13, 21)
        Me.EnableAward.TabIndex = 48
        '
        'SelectScore
        '
        Me.SelectScore.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                      10.0!,
                                                      System.Drawing.FontStyle.Regular,
                                                      System.Drawing.GraphicsUnit.Point,
                                                      CType(0, Byte))
        Me.SelectScore.Location = New System.Drawing.Point(28, 240)
        Me.SelectScore.Name = "SelectScore"
        Me.SelectScore.Size = New System.Drawing.Size(73, 24)
        Me.SelectScore.TabIndex = 50
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                 10.0!,
                                                 System.Drawing.FontStyle.Regular,
                                                 System.Drawing.GraphicsUnit.Point,
                                                 CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(28, 218)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 16)
        Me.Label6.TabIndex = 51
        Me.Label6.Text = "Score"
        '
        'SelectDate
        '
        Me.SelectDate.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                     10.0!,
                                                     System.Drawing.FontStyle.Regular,
                                                     System.Drawing.GraphicsUnit.Point,
                                                     CType(0, Byte))
        Me.SelectDate.Location = New System.Drawing.Point(28, 32)
        Me.SelectDate.Name = "SelectDate"
        Me.SelectDate.Size = New System.Drawing.Size(174, 24)
        Me.SelectDate.TabIndex = 52
        '
        'btnThumbnails
        '
        Me.btnThumbnails.Image = Global.RPS_Digital_Viewer.My.Resources.Resources.thumbs
        Me.btnThumbnails.Location = New System.Drawing.Point(155, 223)
        Me.btnThumbnails.Margin = New System.Windows.Forms.Padding(0)
        Me.btnThumbnails.Name = "btnThumbnails"
        Me.btnThumbnails.Size = New System.Drawing.Size(40, 41)
        Me.btnThumbnails.TabIndex = 49
        '
        'btnSlideShow
        '
        Me.btnSlideShow.BackColor = System.Drawing.SystemColors.Control
        Me.btnSlideShow.FlatAppearance.BorderSize = 0
        Me.btnSlideShow.Image = Global.RPS_Digital_Viewer.My.Resources.Resources.Single_Photo
        Me.btnSlideShow.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.btnSlideShow.Location = New System.Drawing.Point(112, 223)
        Me.btnSlideShow.Name = "btnSlideShow"
        Me.btnSlideShow.Size = New System.Drawing.Size(40, 41)
        Me.btnSlideShow.TabIndex = 22
        Me.btnSlideShow.UseVisualStyleBackColor = False
        '
        'DataGridTextBoxColumn3
        '
        Me.data_grid_text_box_column3.Format = ""
        Me.data_grid_text_box_column3.FormatInfo = Nothing
        Me.data_grid_text_box_column3.HeaderText = "Title"
        Me.data_grid_text_box_column3.MappingName = "Title"
        Me.data_grid_text_box_column3.NullText = ""
        Me.data_grid_text_box_column3.Width = 270
        '
        'DataGridTextBoxColumn2
        '
        Me.data_grid_text_box_column2.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.data_grid_text_box_column2.Format = ""
        Me.data_grid_text_box_column2.FormatInfo = Nothing
        Me.data_grid_text_box_column2.HeaderText = "Award"
        Me.data_grid_text_box_column2.MappingName = "Award"
        Me.data_grid_text_box_column2.NullText = ""
        Me.data_grid_text_box_column2.Width = 50
        '
        'DataGridTextBoxColumn1
        '
        Me.data_grid_text_box_column1.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.data_grid_text_box_column1.Format = ""
        Me.data_grid_text_box_column1.FormatInfo = Nothing
        Me.data_grid_text_box_column1.HeaderText = "Score"
        Me.data_grid_text_box_column1.MappingName = "Score_1"
        Me.data_grid_text_box_column1.NullText = ""
        Me.data_grid_text_box_column1.Width = 50
        '
        'DataGridView1
        '
        Me.data_grid_entries_view.AllowUserToAddRows = False
        Me.data_grid_entries_view.AllowUserToDeleteRows = False
        Me.data_grid_entries_view.AllowUserToResizeColumns = False
        Me.data_grid_entries_view.AllowUserToResizeRows = False
        Me.data_grid_entries_view.Anchor =
            CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                   Or System.Windows.Forms.AnchorStyles.Right),
                  System.Windows.Forms.AnchorStyles)
        Me.data_grid_entries_view.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        Me.data_grid_entries_view.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.data_grid_entries_view.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                              9.75!,
                                                              System.Drawing.FontStyle.Bold,
                                                              System.Drawing.GraphicsUnit.Point,
                                                              CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.data_grid_entries_view.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.data_grid_entries_view.ColumnHeadersHeightSizeMode =
            System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.data_grid_entries_view.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.score, Me.award, Me.title})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                              9.75!,
                                                              System.Drawing.FontStyle.Regular,
                                                              System.Drawing.GraphicsUnit.Point,
                                                              CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.data_grid_entries_view.DefaultCellStyle = DataGridViewCellStyle2
        Me.data_grid_entries_view.EnableHeadersVisualStyles = False
        Me.data_grid_entries_view.Location = New System.Drawing.Point(209, 32)
        Me.data_grid_entries_view.Name = "data_grid_entries_view"
        Me.data_grid_entries_view.ReadOnly = True
        Me.data_grid_entries_view.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.data_grid_entries_view.RowHeadersWidthSizeMode =
            System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.data_grid_entries_view.Size = New System.Drawing.Size(729, 433)
        Me.data_grid_entries_view.TabIndex = 53
        '
        'Score
        '
        Me.score.DataPropertyName = "Score_1"
        Me.score.FillWeight = 21.80233!
        Me.score.HeaderText = "Score"
        Me.score.Name = "Score"
        Me.score.ReadOnly = True
        Me.score.Width = 73
        '
        'Award
        '
        Me.award.DataPropertyName = "Award"
        Me.award.FillWeight = 21.80233!
        Me.award.HeaderText = "Award"
        Me.award.Name = "Award"
        Me.award.ReadOnly = True
        Me.award.Width = 75
        '
        'Title
        '
        Me.title.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.title.DataPropertyName = "Title"
        Me.title.FillWeight = 256.3954!
        Me.title.HeaderText = "Title"
        Me.title.Name = "Title"
        Me.title.ReadOnly = True
        '
        'GridCaption
        '
        Me.grid_caption.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                                       Or System.Windows.Forms.AnchorStyles.Right),
                                      System.Windows.Forms.AnchorStyles)
        Me.grid_caption.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.grid_caption.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.grid_caption.Font = New System.Drawing.Font("Microsoft Sans Serif",
                                                      9.75!,
                                                      System.Drawing.FontStyle.Bold,
                                                      System.Drawing.GraphicsUnit.Point,
                                                      CType(0, Byte))
        Me.grid_caption.ForeColor = System.Drawing.Color.Black
        Me.grid_caption.Location = New System.Drawing.Point(209, 9)
        Me.grid_caption.Margin = New System.Windows.Forms.Padding(0)
        Me.grid_caption.Name = "grid_caption"
        Me.grid_caption.Size = New System.Drawing.Size(729, 24)
        Me.grid_caption.TabIndex = 54
        Me.grid_caption.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MainForm
        '
        Me.AutoScroll = True
        Me.AutoScrollMinSize = New System.Drawing.Size(640, 480)
        Me.ClientSize = New System.Drawing.Size(950, 523)
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
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnCancelAll)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "MainForm"
        Me.Text = "RPS Digital Competition Viewer"
        CType(Me.data_grid_entries_view, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

#End Region

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dirinfo As DirectoryInfo
        Try
            ' Set up the connection string for the database connection
            SetDatabaseName(_database_file_name)
            ' Load the user preferences from the registry
            LoadPreferences()

            ' Setup Database Variables
            _connection_string = New EntityConnectionStringBuilder() With
                {
                .Metadata = "res://*/RpsModel.csdl|res://*/RpsModel.ssdl|res://*/RpsModel.msl",
                .Provider = "System.Data.SQLite.EF6",
                .ProviderConnectionString = New SQLiteConnectionStringBuilder() With
                    {
                    .DataSource = _database_file_name,
                    .ForeignKeys = True
                    }.ConnectionString
                }.ConnectionString

            rps_context = New rpsEntities(_connection_string)
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

            LoadClubRules()

            ' Initialize the StatusBar and ProgressBar
            initializeStatusBar()
            status_bar.progressBar.Value = 0
            ' Load the unique competition dates into the Competition Date combobox
            LoadCompDates()

        Catch ex As Exception
            MsgBox(ex.Message, , "Error In MainForm_Load()")
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

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Try
            ' Recalculate the awards
            CalculateAwards()
        Catch eUpdate As Exception
            'Add your error handling code here.
            'Display error message, if any.
            MsgBox(eUpdate.Message, , "Error In btnUpdate_Click()")
        End Try
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        SelectImages()
    End Sub


    Private Sub btnCancelAll_Click(sender As Object, e As EventArgs) _
        Handles btnCancelAll.Click
        'Me.objSelectedPhotos.RejectChanges()
    End Sub

    Private Sub btnSlideShow_Click(sender As Object, e As EventArgs) _
        Handles btnSlideShow.Click
        DoSlideShow()
    End Sub

    Private Sub btnThumbnails_Click(sender As Object, e As EventArgs) _
        Handles btnThumbnails.Click
        PickAwards()
    End Sub

    Private Sub FileCatalogImagesDownload_Click(sender As Object, e As EventArgs) _
        Handles CompCatalogImagesDownload.Click
        DownloadCompetitionImages()
        LoadCompDates()
    End Sub

    Private Sub FileUploadScores_Click(sender As Object, e As EventArgs) _
        Handles CompUploadScores.Click
        UploadScores()
    End Sub

    Private Sub FileExitMenu_Click(sender As Object, e As EventArgs) _
        Handles FileExitMenu.Click
        Me.Close()
    End Sub


    Private Sub SelectMedium_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles SelectMedium.SelectedIndexChanged
        SelectScore.SelectedItem = "All"
        _all_scores_selected = True
        _eights_and_awards_selected = False
        SelectImages()
    End Sub

    Private Sub SelectClassification_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles SelectClassification.SelectedIndexChanged
        SelectScore.SelectedItem = "All"
        _all_scores_selected = True
        _eights_and_awards_selected = False
        SelectImages()
    End Sub

    Private Sub SelectTheme_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles SelectTheme.SelectedIndexChanged
        SelectScore.SelectedItem = "All"
        _all_scores_selected = True
        _eights_and_awards_selected = False
        SelectImages()
    End Sub

    Private Sub SelectAward_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles SelectAward.SelectedIndexChanged
        SelectScore.SelectedItem = "All"
        _all_scores_selected = True
        _eights_and_awards_selected = False
        SelectImages()
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
        SelectImages()
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
        SelectImages()
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
        SelectImages()
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
        SelectImages()
    End Sub

    Private Sub RecalcAwards_Click(sender As Object, e As EventArgs) _
        Handles RecalcAwards.Click
        CalculateAwards()
    End Sub

    Private Sub HelpAboutMenu_Click(sender As Object, e As EventArgs) _
        Handles HelpAboutMenu.Click
        Dim aboutForm As New About
        aboutForm.Show()
    End Sub

    Private Sub ViewSlideShowMenu_Click(sender As Object, e As EventArgs) _
        Handles ViewSlideShowMenu.Click
        DoSlideShow()
    End Sub

    Private Sub ViewThumbnailsMenu_Click(sender As Object, e As EventArgs) _
        Handles ViewThumbnailsMenu.Click
        PickAwards()
    End Sub

    Private Sub MainForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        'GridResizeColumns(grdCompetition_Entries, 60, 0, 10, 10, 50, 30)
        'GridResizeColumns(grdCompetition_Entries, 60, 0, 15, 15, 70)
    End Sub

    Private Sub FilePreferencesMenu_Click(sender As Object, e As EventArgs) _
        Handles FilePreferencesMenu.Click
        GetUserPreferences()
    End Sub

    Private Sub ReportsResultsReportMenu_Click(sender As Object, e As EventArgs) _
        Handles ReportsResultsReportMenu.Click
        ResultsReport(True)
    End Sub

    Private Sub ReportsScoreSheetMenu_Click(sender As Object, e As EventArgs) _
        Handles ReportsScoreSheetMenu.Click
        'ScoreSheet()
        ResultsReport(False)
    End Sub

    Private Sub DoSlideShow()
        Dim Viewer As ImageViewer
        Dim showSplash As Boolean
        Dim statusBarState As Integer

        ' Bail out if the dataset is empty
        If data_grid_entries_view.RowCount() <= 0 Then
            MsgBox("No competition has been loaded.", MsgBoxStyle.Exclamation, "Error In DoSlideShow()")
            Exit Sub
        End If

        ' Display the splash screen if about to view all images in a competition starting
        ' from the beginning
        'If AllScoresRadioButton.Checked And grdCompetition_Entries.CurrentRowIndex <= 0 Then
        If _all_scores_selected And data_grid_entries_view.CurrentCell.RowIndex <= 0 Then
            showSplash = True
        Else
            showSplash = False
        End If

        ' Set the status bar to show the title and maker name if we're announcing the winners
        'If EightsAndAwardsRadioButton.Checked Then
        If _eights_and_awards_selected Then
            statusBarState = 2
            ' Set the status bar to show the title only if we're assigning awards
            'ElseIf NineScoreRadioButton.Checked Or EightScoreRadioButton.Checked Or SevenScoreRadioButton.Checked Then
        ElseIf Not _all_scores_selected Then
            statusBarState = 1
        Else
            statusBarState = 0
        End If

        Viewer = New ImageViewer(Me, entries, data_grid_entries_view.CurrentCell.RowIndex, showSplash, statusBarState)
        Cursor.Hide()
        Viewer.ShowDialog()
        Cursor.Show()
        Try
            ' Attempt to update the datasource.
            data_grid_entries_view.Refresh()

            For Each entry As CompetitionEntry In entries
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
                CalculateAwards()
                'PickAwards()
            End If
        Catch eUpdate As Exception
            'Add your error handling code here.
            'Display error message, if any.
            MsgBox(eUpdate.Message, , "Error In DoSlideShow()")
        End Try
    End Sub
    '
    ' Convert a date string in the form of dd-MMM-yyyy e.g 19-Nov-2009 to a date
    '
    Private Function ParseSelectedDate(s) As Date
        Dim d As Date
        Try
            d = Date.ParseExact(s, "dd-MMM-yyyy", CultureInfo.CurrentCulture)
            Return d
        Catch ex As Exception
            MsgBox(ex.Message, , "Error In ParseSelectedDate()")
            Return d
        End Try
    End Function

    Private Sub PickAwards()
        Dim screenTitle As String = ""
        Dim Viewer As ThumbnailViewer
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
                screenTitle = SelectClassification.Text + " " + SelectMedium.Text
            ElseIf _eights_and_awards_selected Then
                If num_judges > 1 Then
                    screenTitle = "Award winners And images averaging 8 points Or more"
                Else
                    screenTitle = "Award winners And images With 8 points Or more"
                End If
            ElseIf _selected_avg_score > 0 Then
                If _selected_avg_score = 9 Then
                    screenTitle = _nine_point_thumb_view_title
                ElseIf _selected_avg_score = 8 Then
                    screenTitle = _eight_point_thumb_view_title
                ElseIf _selected_avg_score = 7 Then
                    screenTitle = _seven_point_thumb_view_title
                End If
            Else
                screenTitle = "Images scoring " + CType(_selected_score, String) + " points"
            End If

            ' Launch the thumbnail screen
            Viewer = New ThumbnailViewer(Me, entries, screenTitle)
            Viewer.ShowDialog()

            ' Attempt to update the datasource.
            data_grid_entries_view.Refresh()

            For Each entry As CompetitionEntry In entries
                _query = "UPDATE CompetitionEntries SET Score_1=@score , Award=@award Where Server_Entry_ID=@key"
                rps_context.Database.ExecuteSqlCommand(_query,
                                                      New SQLiteParameter("@score", entry.Score_1),
                                                      New SQLiteParameter("@award", entry.Award),
                                                      New SQLiteParameter("@key", entry.Server_Entry_ID)
                                                      )
            Next

        Catch ex As Exception
            MsgBox(ex.Message, , "Error In PickAwards()")
        End Try
    End Sub

    Private Function CheckFileName(file As FileInfo) As Boolean
        Dim fileName As String
        Dim filePath As String
        Dim newFileName As String
        Dim fields
        Dim validFileName As Boolean
        Dim cancelled As Boolean

        Try
            Do
                validFileName = True
                cancelled = False
                fileName = file.Name        ' Get the name without the path
                ' Does the file name end in ".jpg"?
                If InStr(1, LCase(fileName), ".jpg") = Len(fileName) - 3 Then
                    ' Yes, strip off the extension
                    fileName = Mid(fileName, 1, InStr(1, LCase(fileName), ".jpg") - 1)
                    ' Does the file name contain a single plus sign?
                    fields = Split(fileName, "+")
                    If UBound(fields) <> 1 Then
                        validFileName = False
                    End If
                Else
                    validFileName = False
                End If

                ' Give the operator an opportunity to rename the file
                If Not validFileName Then
                    newFileName = InputBox("Rename the file?", "Invalid File Name", file.Name)
                    If newFileName > "" Then
                        ' Strip the old name from the path
                        filePath = Mid(file.FullName, 1, InStrRev(file.FullName, "\") - 1)
                        ' Rename it, but make sure you don't move it
                        file.MoveTo(filePath + "\" + newFileName)
                    Else
                        cancelled = True
                    End If
                End If
            Loop While Not validFileName And Not cancelled

            If Not validFileName Then
                CheckFileName = False
            Else
                CheckFileName = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error In CheckFileName()")
            CheckFileName = False
        End Try
    End Function


    Private Sub CatalogOneImage(file As FileInfo,
                                classification As String,
                                medium As String,
                                competitionDate As Date,
                                competitionTheme As String)
        'Dim relativePath As String
        Dim fileName As String
        Dim fields()
        Dim title As String
        Dim maker As String
        Dim score As String
        Dim award As String
        'Dim dRow As DataRow

        Try
            ' Parse the title and maker out of the file name
            fileName = file.Name        ' Get the name without the path
            fileName = Mid(fileName, 1, Len(fileName) - 4) ' Strip off the extension
            fields = Split(fileName, "+")
            title = fields(0)
            maker = fields(1)
            title = title.Replace("_", " ")
            maker = maker.Replace("_", " ")
            score = ""
            award = ""

            ' Insert a new row in to the database table
            InsertImageIntoDatabase(file,
                                    maker,
                                    title,
                                    score,
                                    award,
                                    classification,
                                    medium,
                                    competitionDate,
                                    competitionTheme,
                                    "",
                                    0)
        Catch ex As Exception
            MsgBox(ex.Message, , "Error In CatalogOneImage()")
        End Try
    End Sub

    Private Sub InsertImageIntoDatabase(
                                        file As FileInfo,
                                        maker As String,
                                        title As String,
                                        score As String,
                                        award As String,
                                        classification As String,
                                        medium As String,
                                        competitionDate As Date,
                                        competitionTheme As String,
                                        entry_id As String,
                                        sequence As Integer)

        Dim relativePath As String
        Dim entry As CompetitionEntry = New CompetitionEntry

        Try
            ' Calculate the relative path to the file.  The path is relative to the
            ' imagesRootFolder set in File -> Preferences
            If InStr(1, file.FullName, images_root_folder) = 1 Then
                relativePath = Mid(file.FullName, Len(images_root_folder) + 2)
            Else
                relativePath = file.FullName    ' Store absolute path if can't calculate relative path
            End If

            ' Fill in the field values
            entry.Title = title
            entry.Maker = maker
            entry.Classification = classification
            entry.Medium = medium
            entry.Theme = competitionTheme
            entry.Competition_Date_1 = competitionDate
            entry.Image_File_Name = relativePath
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

        Catch ex As Exception
            MsgBox(ex.Message, , "Error In InsertImageInDatabase()")
        End Try
        rps_context.SaveChanges()
    End Sub

    Private Function MaxAwards(numImages As Double) As Integer
        Try
            If numImages = 1 Then
                MaxAwards = 1
            Else
                MaxAwards = Int((numImages / 4) + 0.5)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in MaxAwards()")
        End Try
    End Function

    Private Sub SelectImages()
        Dim numSelected As Double
        Dim select_stmt As String
        Dim where_clause As String
        Dim order_clause As String

        If SelectDate.Text > "" And SelectMedium.Text > "" And SelectClassification.Text > "" Then
            Try
                ' Build the complete SQL statement to select the records specified
                ' by the selection criteria on the main form.
                ' Start with a basic SQL statement that selects records by date.
                select_stmt = "SELECT * FROM CompetitionEntries"
                where_clause = " WHERE Competition_Date_1='" + Format(ParseSelectedDate(SelectDate.Text), "M/dd/yyyy") +
                               "'"
                order_clause = " ORDER BY Display_Sequence, Title"
                grid_caption.Text = Format(ParseSelectedDate(SelectDate.Text), "MM/dd/yyyy")

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
                entries = rps_context.Database.SqlQuery(Of CompetitionEntry)(_query).ToList

                With data_grid_entries_view
                    .Columns("Score").DefaultCellStyle = _center_cell_style
                    .Columns("Award").DefaultCellStyle = _center_cell_style
                    .AutoGenerateColumns = False
                    .DataSource = entries
                End With

                ' Count the number of rows selected and add it to the caption of the DataGrid
                numSelected = data_grid_entries_view.RowCount()

                grid_caption.Text += "  -  " + numSelected.ToString + " Images"

                ' Recalculate the awards
                'If AllScoresRadioButton.Checked And EnableAward.CheckState = CheckState.Unchecked Then
                If _all_scores_selected And EnableAward.CheckState = CheckState.Unchecked Then
                    CalculateAwards()
                End If

            Catch eLoad As Exception
                'Add your error handling code here.
                'Display error message, if any.
                MsgBox(eLoad.Message, , "Error in SelectImages()")
            End Try
        End If
    End Sub


    Private Sub CalculateAwards()
        Dim eligibleScores As New ArrayList
        Dim maximumAwards As Integer
        Dim dRow As CompetitionEntry
        Dim i As Integer
        Dim numEligibleNines As Integer
        Dim numEligibleEights As Integer
        Dim numEligibleSevens As Integer
        Dim totalNumScores As Integer
        Dim awardNames() As String = {"1St", "2nd", "3Rd", "HM"}
        Dim delim_9, delim_8, delim_7 As String
        Dim numNineHM, numEightHM, numSevenHM As Integer

        Try
            ' What is the maximum number of Awards possible
            maximumAwards = MaxAwards(CType(entries.Count, Double))

            ' Iterate through the dataset and record all the scores which are eligible for an award
            For Each dRow In entries
                If IsNumeric(dRow.Score_1) Then
                    totalNumScores = totalNumScores + 1
                    If dRow.Score_1 >= (min_score_for_award * num_judges) And dRow.Score_1 <= (max_score * num_judges) Then
                        eligibleScores.Add(dRow.Score_1)
                    End If
                End If
            Next
            ' Sort descending
            eligibleScores.Sort()
            eligibleScores.Reverse()

            ' Step through the eligible scores until all possible awards have been allocated
            ' or until the number of eligible scores exhausted.
            _nine_point_thumb_view_title = ""
            _eight_point_thumb_view_title = ""
            _seven_point_thumb_view_title = ""
            For i = 0 To Math.Min(maximumAwards, eligibleScores.Count) - 1
                ' Count up the number of eligible 9s, 8s and 7s
                Select Case Math.Round(eligibleScores(i) / num_judges, 0)
                    Case 9
                        ' Count the number of 9 point awards for the distribution on the main screen
                        numEligibleNines = numEligibleNines + 1
                        ' Build the title for the thumbnail screen
                        If i < 3 Then
                            _nine_point_thumb_view_title += delim_9 + awardNames(i)
                        Else
                            numNineHM += 1
                        End If
                        delim_9 = ", "
                    Case 8
                        ' Count the number of 8 point awards for the distribution on the main screen
                        numEligibleEights = numEligibleEights + 1
                        ' Build the title for the thumbnail screen
                        If i < 3 Then
                            _eight_point_thumb_view_title += delim_8 + awardNames(i)
                        Else
                            numEightHM += 1
                        End If
                        delim_8 = ", "
                    Case 7
                        ' Count the number of 7 point awards for the distribution on the main screen
                        numEligibleSevens = numEligibleSevens + 1
                        ' Build the title for the thumbnail screen
                        If i < 3 Then
                            _seven_point_thumb_view_title += delim_7 + awardNames(i)
                        Else
                            numSevenHM += 1
                        End If
                        delim_7 = ", "
                End Select
            Next

            ' Update the 9 point thumbnail screen title to include the HMs
            If numNineHM > 0 Then
                If _nine_point_thumb_view_title > "" Then
                    _nine_point_thumb_view_title += " And "
                End If
                If numNineHM = 1 Then
                    _nine_point_thumb_view_title += "1 HM"
                Else
                    _nine_point_thumb_view_title += CStr(numNineHM) + " HMs"
                End If
            End If

            ' Update the 8 point thumbnail screen title to include the HMs
            If numEightHM > 0 Then
                If _eight_point_thumb_view_title > "" Then
                    _eight_point_thumb_view_title += " And "
                End If
                If numEightHM = 1 Then
                    _eight_point_thumb_view_title += "1 HM"
                Else
                    _eight_point_thumb_view_title += CStr(numEightHM) + " HMs"
                End If
            End If

            ' Update the 7 point thumbnail screen title to include the HMs
            If numSevenHM > 0 Then
                If _seven_point_thumb_view_title > "" Then
                    _seven_point_thumb_view_title += " And "
                End If
                If numSevenHM = 1 Then
                    _seven_point_thumb_view_title += "1 HM"
                Else
                    _seven_point_thumb_view_title += CStr(numSevenHM) + " HMs"
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
            If totalNumScores > 0 Then
                AwardsTableTitleBar.Text = Math.Min(maximumAwards, eligibleScores.Count).ToString + " Awards"
            Else
                AwardsTableTitleBar.Text = maximumAwards.ToString + " Awards"
            End If
            tbEligibleNines.Text = numEligibleNines.ToString
            tbEligibleEights.Text = numEligibleEights.ToString
            tbEligibleSevens.Text = numEligibleSevens.ToString

        Catch ex As Exception
            MsgBox(ex.Message, , "Error In CalculateAwards()")
        End Try
    End Sub


    Private Sub ResultsReport(displayScores As Boolean)
        Dim tempFile As String
        Dim reportType As String
        Dim competitionDate As Date
        Dim img As CompetitionEntry
        Dim f As File
        Dim sw As StreamWriter
        Try
            ' Bail out if the dataset is empty
            If entries.Count <= 0 Then
                MsgBox("No images selected." + vbCrLf + "Select one Or more images before running the report",
                       MsgBoxStyle.Exclamation,
                       "Error In ResultsReport()")
                Exit Sub
            End If

            ' Grab the first row of the collection to get some values for the report header
            img = entries.First

            ' Open the output file and write the HTML preamble
            'competitionDate = CType(SelectDate.Text, Date)
            competitionDate = ParseSelectedDate(SelectDate.Text)
            If displayScores Then
                reportType = "Results"
            Else
                reportType = "Scoresheet"
            End If
            tempFile = reports_output_folder + "\" +
                       reportType + "_" + CType(competitionDate.Year, String) + "-" +
                       CType(competitionDate.Month, String) + "-" +
                       CType(competitionDate.Day, String)
            If EnableTheme.Checked Then
                tempFile += "_" + SelectTheme.Text
            End If
            If EnableClassification.Checked Then
                tempFile += "_" + SelectClassification.Text
            End If
            If EnableMedium.Checked Then
                tempFile += "_" + SelectMedium.Text
            End If
            If SelectScore.Text <> "All" Then
                tempFile += "_" + SelectScore.Text
            End If
            If EnableAward.Checked Then
                tempFile += "_" + SelectAward.Text
            End If
            tempFile = StrMap(tempFile, " ?[]/\=+<>:;"",*|", "_---------------") + ".html"
            sw = f.CreateText(tempFile)
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
                competitionDate.ToString("MMMM", CultureInfo.CurrentCulture) + " " +
                competitionDate.Year.ToString)
            If EnableTheme.Checked Then
                sw.Write(" - " + SelectTheme.Text)
            End If
            sw.WriteLine("</p></td></tr>")
            If Not displayScores Then
                sw.WriteLine(
                    "<tr><td colspan=""4"" class=""header_center""><p class=""subtitle"">(" +
                    MaxAwards(entries.Count).ToString + " Awards)</p></td></tr>")
            End If
            sw.WriteLine("<tr><td colspan=""4"" class=""header_center""><p class=""subtitle"">&nbsp;</p></td></tr>")
            sw.WriteLine("<tr><th>Score</th><th>Award</th>")
            sw.WriteLine("<th>Title</th><th>Photographer</th></tr>")
            '
            ' Print a row for each image in the category
            For Each img In entries
                sw.WriteLine("<tr>")
                If displayScores Then
                    sw.WriteLine("  <td class=""score"">" + img.Score_1.ToString + "</td>")
                    sw.WriteLine("  <td class=""award"">" + img.Award + "</td>")
                Else
                    sw.WriteLine("  <td class=""score"">&nbsp;</td>")
                    sw.WriteLine("  <td class=""award"">&nbsp;</td>")
                End If
                sw.WriteLine("  <td class=""title"">" + img.Title + "</td>")
                sw.WriteLine("  <td class=""photographer"">" + img.Maker + "</td>")
                sw.WriteLine("</tr>")
            Next img
            '
            ' Finish off the HTML
            sw.WriteLine("</table>")
            sw.WriteLine("</body>")
            sw.WriteLine("</html>")
            sw.Close()
            '
            ' Launch a browser to display the worksheet
            ShellExecute(tempFile)
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in ResultsReport()")
        End Try
    End Sub


    Private Function ShellExecute(File As String) As Boolean
        Try
            Dim myProcess As New Process
            myProcess.StartInfo.FileName = File
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.Start()
            myProcess.Dispose()
        Catch e As Exception
            MsgBox(e.Message, , "Error in ShellExecute()")
        End Try
    End Function

    Public Shared Function GetDBStringField(r As DataRow, name As String, nullStr As String) As String
        Try
            If r(name) Is DBNull.Value Then
                GetDBStringField = nullStr
            Else
                GetDBStringField = CType(r(name), String)
            End If
        Catch e As Exception
            MsgBox(e.Message, , "Error in GetDBStringField()")
        End Try
    End Function


    Private Sub GridResizeColumns(DataGridView As DataGrid,
                                  FixedTotalWidth As Double,
                                  StartCol As Integer,
                                  ByVal ParamArray ColumnWidthInPercent() As Double)

        Dim TotalGridWidth As Long
        Dim intIndex As Integer
        Dim PercentColWidth As Integer
        Dim i As Integer
        Try
            TotalGridWidth = DataGridView.Width - FixedTotalWidth

            For i = 0 To ColumnWidthInPercent.Length - 1
                DataGridView.TableStyles(0).GridColumnStyles(i).Width = ColumnWidthInPercent(i) * TotalGridWidth / 100
            Next
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in GridResizeColumns()")
        End Try
    End Sub


    Private Sub GetUserPreferences()
        Dim prefsDialog As New PreferencesDialog(Me)
        Dim s As String

        Try
            ' Load the current perferences into the dialog
            prefsDialog.tbDatabaseFileName.Text = _database_file_name
            If Len(images_root_folder) = 2 And Mid(images_root_folder, 2, 1) = ":" Then
                prefsDialog.tbImagesRootFolder.Text = images_root_folder + "\"
            Else
                prefsDialog.tbImagesRootFolder.Text = images_root_folder
            End If
            If Len(reports_output_folder) = 2 And Mid(reports_output_folder, 2, 1) = ":" Then
                prefsDialog.tbReportsOutputFolder.Text = reports_output_folder + "\"
            Else
                prefsDialog.tbReportsOutputFolder.Text = reports_output_folder
            End If
            prefsDialog.tbServerName.Text = _server_name
            prefsDialog.tbServerScriptDir.Text = _server_script_dir
            _query = From club In rps_context.clubs
                     Select club.id, club.name

            For Each _record In _query
                prefsDialog.cbCameraClubName.Items.Add(New DataItem(_record.id, _record.name))
            Next
            prefsDialog.cbCameraClubName.Text = camera_club_name
            prefsDialog.cbNumJudges.Text = CType(num_judges, String)

            ' Display the dialog
            prefsDialog.ShowDialog()

            ' Check the results
            Dim irf As String = Trim(prefsDialog.tbImagesRootFolder.Text)
            Dim dbfn As String = Trim(prefsDialog.tbDatabaseFileName.Text)
            Dim rof As String = Trim(prefsDialog.tbReportsOutputFolder.Text)
            Dim sn As String = Trim(prefsDialog.tbServerName.Text)
            Dim ssd As String = Trim(prefsDialog.tbServerScriptDir.Text)
            Dim ccn As String = Trim(CType(prefsDialog.cbCameraClubName.SelectedItem, DataItem).Value)
            Dim ccid As Integer = CType(prefsDialog.cbCameraClubName.SelectedItem, DataItem).ID
            Dim nj As Integer = CType(prefsDialog.cbNumJudges.Text, Integer)
            If prefsDialog.DialogResult = DialogResult.OK Then
                If irf > "" Then
                    ' If necessary, strip off a trailing "\"
                    images_root_folder = Helper.TrimTrailingSlash(images_root_folder)
                    ' write it to the registry
                    WriteRegistryString("Software\RPS Digital Viewer", "Images Root Folder", images_root_folder)
                End If
                If dbfn > "" Then
                    ' update the connection string
                    SetDatabaseName(dbfn)
                    ' write it to the registry
                    WriteRegistryString("Software\RPS Digital Viewer", "Database File Name", _database_file_name)
                End If
                If rof > "" Then
                    ' If necessary, strip off a trailing "\"
                    reports_output_folder = Helper.TrimTrailingSlash(reports_output_folder)

                    ' write it to the registry
                    WriteRegistryString("Software\RPS Digital Viewer", "Reports Output Folder", reports_output_folder)
                End If
                If sn > "" Then
                    ' Store the new server name in memory
                    _server_name = sn
                    ' Write it to the registry
                    WriteRegistryString("Software\RPS Digital Viewer", "Server Name", sn)
                End If
                If ssd > "" Then
                    ' Store the new server script directory in memory
                    _server_script_dir = ssd
                    ' Write it to the registry
                    WriteRegistryString("Software\RPS Digital Viewer", "Server Script Directory", ssd)
                End If
                If ccid > 0 Then
                    ' Set and store the club id
                    camera_club_id = ccid
                    WriteRegistryString("Software\RPS Digital Viewer", "Camera Club ID", CType(ccid, String))
                End If
                If ccn > "" Then
                    ' Set and store the club name
                    camera_club_name = ccn
                    WriteRegistryString("Software\RPS Digital Viewer", "Camera Club Name", ccn)
                End If
                If nj > 0 Then
                    ' Set and store the number of judges
                    num_judges = nj
                    WriteRegistryString("Software\RPS Digital Viewer", "Number of Judges", CType(nj, String))
                End If

                ' Fetch the list of classifications, mediums and awards from the database
                ' for the selected club
                LoadClubRules()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in GetUserPreferences()")
        End Try
    End Sub

    Private Sub SetDatabaseName(fileName As String)
        Try
            _database_file_name = fileName
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in SetDatabasename()")
        End Try
    End Sub

    Private Sub WriteRegistryString(key As String, name As String, value As String)
        Dim regKey As RegistryKey

        Try
            regKey = Registry.CurrentUser.OpenSubKey(key, True)
            If regKey Is Nothing Then
                regKey = Registry.CurrentUser.CreateSubKey(key)
            End If
            regKey.SetValue(name, value)
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in WriteRegistryString()")
        Finally
            If Not regKey Is Nothing Then
                regKey.Close()
            End If
        End Try
    End Sub

    Private Function ReadRegistryString(key As String, name As String) As String
        Dim regKey As RegistryKey

        Try
            regKey = Registry.CurrentUser.OpenSubKey(key, False)
            If regKey Is Nothing Then
                ReadRegistryString = ""
            Else
                ReadRegistryString = regKey.GetValue(name, "")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in ReadRegistryString()")
        Finally
            If Not regKey Is Nothing Then
                regKey.Close()
            End If
        End Try
    End Function

    Private Sub LoadPreferences()
        Dim value As String

        Try
            value = ReadRegistryString("Software\RPS Digital Viewer", "Images Root Folder")
            If value > "" Then
                images_root_folder = value
            End If
            value = ReadRegistryString("Software\RPS Digital Viewer", "Database File Name")
            If value > "" Then
                SetDatabaseName(value)
            End If
            value = ReadRegistryString("Software\RPS Digital Viewer", "Reports Output Folder")
            If value > "" Then
                reports_output_folder = value
            End If
            value = ReadRegistryString("Software\RPS Digital Viewer", "Server Name")
            If value > "" Then
                _server_name = value
            End If
            value = ReadRegistryString("Software\RPS Digital Viewer", "Server Script Directory")
            If value > "" Then
                _server_script_dir = value
            End If
            value = ReadRegistryString("Software\RPS Digital Viewer", "Camera Club ID")
            If value > "" Then
                camera_club_id = CType(value, Integer)
            End If
            value = ReadRegistryString("Software\RPS Digital Viewer", "Camera Club Name")
            If value > "" Then
                camera_club_name = value
            End If
            value = ReadRegistryString("Software\RPS Digital Viewer", "Last Admin Username")
            If value > "" Then
                last_admin_username = value
            End If
            value = ReadRegistryString("Software\RPS Digital Viewer", "Number of Judges")
            If value > "" Then
                num_judges = value
            End If

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in LoadPreferences()")
        End Try
    End Sub

    '
    ' Query the database to retrieve the list of classifications, mediums and awards for the
    ' selected club.  Store these lists in memory.
    '
    Private Sub LoadClubRules()
        Dim rec As OleDbDataReader

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

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in LoadClubRules()")
        End Try
    End Sub

    '
    ' Load the list of unique competition dates into the Competition Dates combobox
    '
    Private Sub LoadCompDates()

        _query = From entries In rps_context.CompetitionEntries
                 Order By entries.Competition_Date_1
                 Select entries.Competition_Date_1 Distinct

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
        Catch ex As Exception
            MsgBox(ex.Message, , "Error In LoadCompDates()")
        End Try
    End Sub
    ' Call the REST service on the server to retrieve the list of available
    ' competition dates.
    Private Function GetCompetitionDates(params As Hashtable) As ArrayList
        Dim navigator As XPathNavigator
        Dim response As XPathDocument
        Dim nodes As XPathNodeIterator
        Dim node As XPathNavigator
        Dim dates As New ArrayList

        Try
            ' Retrieve the list of competition dates from the server
            params.Add("rpswinclient", "getcompdate")
            If REST(_server_name, _server_script_dir, "GET", params, response) Then
                navigator = response.CreateNavigator()
                nodes = navigator.Select("/rsp/Competition_Date")
                While nodes.MoveNext()
                    node = nodes.Current
                    dates.Add(node.Value)
                End While
            Else
                Exit Function
            End If

            Return dates

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in GetCompetitionDates()")
        End Try
    End Function

    Structure comp_entry
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
            If Me.bucket.CompareTo(CType(obj, comp_entry).bucket) = 0 Then
                Return Me.first_name.CompareTo(CType(obj, comp_entry).first_name)
            Else
                Return Me.bucket.CompareTo(CType(obj, comp_entry).bucket)
            End If
        End Function
    End Structure

    Private Sub DownloadCompetitionImages()

        Dim Download_Dialog As Download_Competitions_Dialog
        Dim username As String
        Dim password As String
        Dim comp_date As String
        Dim comp_classification As String
        Dim comp_theme As String
        Dim comp_medium As String
        'Dim first_name As String
        'Dim last_name As String
        'Dim title As String
        'Dim score As String
        'Dim award As String
        'Dim url As String
        'Dim entry_id As String
        Dim localImageFileName As String
        Dim outputFolder As String
        Dim competitionFolder As String
        Dim dirInfo As DirectoryInfo
        Dim params As New Hashtable
        Dim response As XPathDocument
        Dim comp_dates As New ArrayList
        Dim navigator As XPathNavigator
        Dim nodes As XPathNodeIterator
        Dim node As XPathNavigator
        Dim comp_nodes As XPathNodeIterator
        Dim comp_node As XPathNavigator
        Dim entry_nodes As XPathNodeIterator
        Dim entry_node As XPathNavigator
        Dim entry_properties As XPathNodeIterator
        Dim entry_property As XPathNavigator
        Dim sqlCommand As New OleDbCommand
        Dim sql As String
        Dim dateParts() As String
        Dim entries As New ArrayList
        Dim this_entry As comp_entry
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
            comp_dates = GetCompetitionDates(params)
            If comp_dates.Count = 0 Then
                MsgBox("No competitions are available for download", , "Download Competition Images")
                Exit Sub
            End If
            Cursor.Current = Cursors.Default

            ' Open the Download Competitions dialog
            Download_Dialog = New Download_Competitions_Dialog(Me, last_admin_username, comp_dates, images_root_folder)
            Download_Dialog.ShowDialog(Me)
            If Download_Dialog.DialogResult = DialogResult.OK Then
                username = Trim(Download_Dialog.Username.Text())
                password = Trim(Download_Dialog.Password.Text())
                comp_date = Download_Dialog.CompetitionDate.Text()
                download_digital = Download_Dialog.Download_digital.Checked
                download_prints = Download_Dialog.Download_prints.Checked
                outputFolder = Trim(Download_Dialog.OutputFolder.Text())
            Else
                Exit Sub
            End If

            ' Save the username in the registry as the default admin username
            last_admin_username = username
            WriteRegistryString("Software\RPS Digital Viewer", "Last Admin Username", last_admin_username)

            ' Delete any competitions in the local database that already have this date
            dateParts = Split(comp_date, "-")
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
            If Not REST(_server_name, _server_script_dir, "POST", params, response) Then
                navigator = response.CreateNavigator()
                nodes = navigator.Select("/rsp/err")
                nodes.MoveNext()
                node = nodes.Current
                MsgBox("Download Failure - " + node.GetAttribute("msg", ""), , "Error in DownloadCompetitionImages()")
                Exit Sub
            End If

            ' Initialize the progress bar
            status_bar.progressBar.Minimum = 0
            status_bar.progressBar.Value = 0
            navigator = response.CreateNavigator()
            nodes = navigator.Select("/descendant::*[name()='Image_URL']") ' Count the images in the manifest
            status_bar.progressBar.Maximum = nodes.Count

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
                entries.Clear()
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
                    entries.Add(this_entry)     ' Store the record in a list
                    prev_member = member


                    ' If necessary, create the folder for this entry
                    'competitionFolder = outputFolder + "\" + comp_date + " " + comp_classification + " " + comp_medium
                    'dirInfo = New DirectoryInfo(competitionFolder)
                    'If Not dirInfo.Exists Then
                    '   dirInfo.Create()
                    'End If

                    ' Fetch the image file from the server
                    'localImageFileName = competitionFolder + "\" + _
                    '    StrMap(title, " ?[]/\=+<>:;"",*|", "_---------------") + _
                    '    "+" + first_name + "_" + last_name + ".jpg"
                    'DownloadImage(url, localImageFileName)

                    ' Insert this image into the database
                    'InsertImageIntoDatabase( _
                    '    New FileInfo(localImageFileName), _
                    '    first_name + " " + last_name, _
                    '    title, _
                    '    score, _
                    '    award, _
                    '    comp_classification, _
                    '    comp_medium, _
                    '    comp_date, _
                    '    comp_theme, _
                    '    entry_id)

                    ' Update the Progressbar
                    'StatusBar.progressBar.Value = StatusBar.progressBar.Value + 1
                    'Application.DoEvents()

                End While

                ' Sort the list of entries by bucket
                entries.Sort()

                ' Iterate through all the entries and download the images
                ' and update the database
                sequence_num = 1
                For Each entry As comp_entry In entries
                    ' If necessary, create the folder for this entry
                    competitionFolder = outputFolder + "\" + comp_date + " " + comp_classification + " " + comp_medium
                    dirInfo = New DirectoryInfo(competitionFolder)
                    If Not dirInfo.Exists Then
                        dirInfo.Create()
                    End If

                    ' Fetch the image file from the server
                    localImageFileName = competitionFolder + "\" +
                                         StrMap(entry.title, " ?[]/\=+<>:;"",*|", "_---------------") +
                                         "+" + entry.first_name + "_" + entry.last_name + ".jpg"
                    DownloadImage(entry.url, localImageFileName)

                    ' Insert this image into the database
                    InsertImageIntoDatabase(
                        New FileInfo(localImageFileName),
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
                    status_bar.progressBar.Value = status_bar.progressBar.Value + 1
                    Application.DoEvents()

                    sequence_num += 1

                Next

            End While

            ' Update the list of dates in the Competition Date combobox
            LoadCompDates()

            MsgBox("Competition Images Downloaded Successfully", , "Download Competition Images")

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in DownloadCompetitionImages()")
        Finally
            ' Clear the ProgressBar
            status_bar.progressBar.Value = 0
            Cursor.Current = Cursors.Default
        End Try
    End Sub


    Private Function REST(server As String,
                          operation As String,
                          http_method As String,
                          params As Hashtable,
                          ByRef results As XPathDocument) As Boolean
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim URL As String
        Dim delim As String
        Dim XPathDoc As XPathDocument
        Dim data As New StringBuilder
        Dim byteData() As Byte
        Dim postStream As Stream = Nothing
        Dim fs As FileStream
        Dim br As BinaryReader
        Dim ms As New MemoryStream

        Try
            ' Build the URL
            URL = "http://" + server + operation
            If http_method = "GET" Then
                delim = "?"
                For Each key As String In params.Keys
                    URL = URL + delim + key + "=" + params.Item(key)
                    delim = "&"
                Next
            End If

            ' Create the web request  
            request = HttpWebRequest.Create(URL)
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
                byteData = UTF8Encoding.UTF8.GetBytes(data.ToString())
                ms.Write(byteData, 0, byteData.Length)
                data.Remove(0, data.Length)

                ' Attach any files in the param list
                For Each param As String In params.Keys
                    If param = "file" Then
                        ' Build the MIME header
                        data.Append(
                            "Content-Disposition: form-data; name=""file""; filename=""" + params(param) + """" + vbCrLf)
                        data.Append("Content-Transfer-Encoding: binary" + vbCrLf)
                        data.Append("Content-Type: Image/jpeg" + vbCrLf)
                        data.Append(vbCrLf)
                        ' Write the MIME header to a binary stream
                        byteData = UTF8Encoding.UTF8.GetBytes(data.ToString())
                        ms.Write(byteData, 0, byteData.Length)
                        ' Open the file and write it to a binary stream
                        fs = New FileStream(params(param), FileMode.Open, FileAccess.Read)
                        br = New BinaryReader(fs)
                        byteData = br.ReadBytes(fs.Length)
                        ms.Write(byteData, 0, fs.Length)
                        br.Close()
                        fs.Close()
                        ' Write the terminating boundry marker
                        data.Remove(0, data.Length)
                        data.Append("--" + "AaB03x" + vbCrLf)
                        byteData = UTF8Encoding.UTF8.GetBytes(data.ToString())
                        ms.Write(byteData, 0, byteData.Length)
                    End If
                Next

                ' Read the entire contents of the memory stream back into a byte array
                byteData = ms.ToArray
                'MsgBox(UTF8Encoding.UTF8.GetString(byteData))

                ' Set the content length in the request headers  
                request.ContentLength = byteData.Length

                ' Write data  
                Try
                    postStream = request.GetRequestStream()
                    postStream.Write(byteData, 0, byteData.Length)
                Finally
                    If Not postStream Is Nothing Then postStream.Close()
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
            XPathDoc = New XPathDocument(response_stream)
            If responseOK(XPathDoc) Then
                REST = True
            Else
                REST = False
            End If
            results = XPathDoc

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in REST()")
            REST = False
        Finally
            If Not response Is Nothing Then response.Close()
        End Try
    End Function

    Private Function responseOK(response As XPathDocument) As Boolean
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
                responseOK = True
            Else
                responseOK = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in responseOK()")
        End Try
    End Function

    Public Sub DownloadImage(URL As String, outputFileName As String)
        Dim Req As HttpWebRequest
        Dim SourceStream As Stream
        Dim Response As HttpWebResponse
        Dim Buffer(4096) As Byte
        Dim BlockSize As Integer
        Dim outputFile As New FileStream(outputFileName, FileMode.Create, FileAccess.Write)

        Try
            'create a web request to the URL
            Req = HttpWebRequest.Create(URL)

            'get a response from web site
            Response = Req.GetResponse()

            'Source stream with requested document
            SourceStream = Response.GetResponseStream()

            'SourceStream has no ReadAll, so we must read data block-by-block
            Do
                BlockSize = SourceStream.Read(Buffer, 0, 4096)
                If BlockSize > 0 Then outputFile.Write(Buffer, 0, BlockSize)
            Loop While BlockSize > 0

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in DownloadImage()")
        Finally
            SourceStream.Close()
            outputFile.Close()
            Response.Close()
        End Try
    End Sub

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

    Public Function StrMap(S As String,
                           MapWhat As String,
                           ToWhat As String,
                           Optional ByVal Compare As Long = 0) As String

        Dim output_ptr As Long
        Dim input_ptr As Long
        Dim c As String
        Dim map_index As Long

        Try
            For input_ptr = 1 To Len(S)
                ' get the next char in the source string
                c = Mid$(S, input_ptr, 1)
                ' is this a character to be mapped?
                map_index = InStr(1, MapWhat, c, Compare)
                ' No. If the output string is shorter than the input string
                ' (because previous characters have been deleted) relocate this character
                ' to the end of the output string.
                If map_index = 0 Then
                    output_ptr = output_ptr + 1
                    If output_ptr < input_ptr Then
                        Mid$(S, output_ptr) = c
                    End If
                ElseIf map_index <= Len(ToWhat) Then
                    ' Yes, this character is to be remapped (or deleted).  If there is a
                    ' corresponding character to map it to, replace the original character,
                    ' otherwise do nothing.  Doing nothing will effectively delete the
                    ' character from the output string.
                    output_ptr = output_ptr + 1
                    Mid$(S, output_ptr) = Mid$(ToWhat, map_index, 1)
                End If
            Next
            ' If the output string is the same length as the input string (i.e. there have
            ' been no deletions) return the original string length.  If there have been 
            ' deletions, return the shortened string.
            If output_ptr = Len(S) Then
                StrMap = S
            Else
                StrMap = Microsoft.VisualBasic.Left(S, output_ptr)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in StrMap()")
        End Try
    End Function

    Function UploadScores() As Boolean
        Dim Upload_Dialog As Upload_Scores_Dialog
        Dim username As String
        Dim password As String
        Dim comp_date As String
        Dim comp_dates As ArrayList
        Dim comp_class_list As New ArrayList
        Dim comp_medium_list As New ArrayList
        Dim sqlSelect As String
        Dim sqlWhere As String
        Dim sqlCommand As New OleDbCommand
        Dim recs As Object
        Dim compNum As Integer
        Dim classification As String
        Dim medium As String
        Dim maker As String
        Dim title As String
        Dim score As String
        Dim award As String
        Dim entry_id As String
        Dim f As File
        Dim sw As StreamWriter
        Dim fileName As String
        Dim params As New Hashtable
        Dim response As XPathDocument
        Dim navigator As XPathNavigator
        Dim nodes As XPathNodeIterator
        Dim node As XPathNavigator
        Dim info As New StringBuilder
        Dim posn As Integer
        Dim firstName As String
        Dim lastName As String
        Dim upload_digital As Boolean
        Dim upload_prints As Boolean
        Dim selectedMedium As String

        Try
            ' Get the list of competition dates from the server
            Cursor.Current = Cursors.WaitCursor
            params.Add("closed", "Y")
            params.Add("scored", "N")
            comp_dates = GetCompetitionDates(params)
            If comp_dates.Count = 0 Then
                MsgBox("No competitions available to receive scores", , "Upload Scores")
                Return False
            End If
            Cursor.Current = Cursors.Default

            ' Open the Upload Scores dialog
            Upload_Dialog = New Upload_Scores_Dialog(last_admin_username, comp_dates)
            Upload_Dialog.ShowDialog(Me)
            If Upload_Dialog.DialogResult = DialogResult.OK Then
                username = Trim(Upload_Dialog.Username.Text())
                password = Trim(Upload_Dialog.Password.Text())
                comp_date = Upload_Dialog.CompDate.Text()
                upload_digital = Upload_Dialog.Upload_digital_scores.Checked
                upload_prints = Upload_Dialog.Upload_print_scores.Checked
            Else
                Exit Function
            End If

            ' Save the username in the registry as the default admin username
            last_admin_username = username
            WriteRegistryString("Software\RPS Digital Viewer", "Last Admin Username", last_admin_username)

            Cursor.Current = Cursors.WaitCursor
            Application.DoEvents()

            ' Open a local text file to receive the XML
            fileName = reports_output_folder + "\" + "Scores_" + comp_date + ".xml"
            sw = f.CreateText(fileName)
            sw.WriteLine("<?xml version=""1.0"" encoding=""utf-8"" ?>")
            sw.WriteLine("<Competitions>")

            ' Select the unique competitions for the given date
            selectedMedium = ""
            If upload_digital And Not upload_prints Then
                selectedMedium = " AND Medium like '%Digital'"
            End If
            If upload_prints And Not upload_digital Then
                selectedMedium = " AND Medium like '%Prints'"
            End If
            Dim d As Date = Date.ParseExact(comp_date, "yyyy-MM-dd", CultureInfo.CurrentCulture)
            sqlSelect = "SELECT DISTINCT Classification, Medium FROM CompetitionEntries "
            sqlWhere = "WHERE Competition_Date_1 = '" +
                       Format(d, "M/dd/yyyy") + "'" +
                       selectedMedium
            _query = sqlSelect + sqlWhere
            recs = rps_context.Database.SqlQuery(Of UploadEntity_Classification_Medium)(_query).ToList

            For Each record As UploadEntity_Classification_Medium In recs
                comp_class_list.Add(record.Classification)
                comp_medium_list.Add(record.Medium)
            Next

            ' Iterate through all the competition for this date
            For compNum = 0 To comp_class_list.Count - 1
                ' Query the database for the entries of this competition
                classification = comp_class_list.Item(compNum)
                medium = comp_medium_list.Item(compNum)
                sqlSelect = "Select * FROM CompetitionEntries "
                sqlWhere = "WHERE Competition_Date_1 = '" + Format(d, "M/dd/yyyy") + "' And classification = '" +
                           classification + "' AND Medium = '" + medium + "'"
                _query = sqlSelect + sqlWhere
                recs = rps_context.Database.SqlQuery(Of CompetitionEntry)(_query).ToList
                ' Output the tags that describe this competition
                sw.WriteLine("  <Competition>")
                sw.WriteLine("    <Date>{0}</Date>", HttpUtility.HtmlEncode(comp_date))
                sw.WriteLine("    <Classification>{0}</Classification>", HttpUtility.HtmlEncode(classification))
                sw.WriteLine("    <Medium>{0}</Medium>", HttpUtility.HtmlEncode(medium))
                sw.WriteLine("    <Entries>")
                ' Iterate through all the entries of this competition
                For Each record As CompetitionEntry In recs
                    ' Read the entry data from the database
                    maker = record.Maker
                    'fields = Split(maker, " ")
                    posn = InStr(1, maker, " ")
                    firstName = Mid(maker, 1, posn - 1)
                    lastName = Mid(maker, posn + 1)
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
                    sw.WriteLine("        <First_Name>{0}</First_Name>", HttpUtility.HtmlEncode(firstName))
                    sw.WriteLine("        <Last_Name>{0}</Last_Name>", HttpUtility.HtmlEncode(lastName))
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
            params.Add("file", fileName)
            If Not REST(_server_name, _server_script_dir + "/?rpswinclient=uploadscore", "POST", params, response) Then
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
            If File.Exists(fileName) = True Then
                File.Delete(fileName)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, , "Error in UploadScores()")
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Function
    '
    ' Enter the unique list of competition themes for this date into the Theme combobox
    '
    Private Sub LoadUniqueThemes()
        Try
            Dim compDate As String = Format(ParseSelectedDate(SelectDate.Text), "M/dd/yyyy")
            themes.Clear()
            SelectTheme.Items.Clear()

            _query = From entries In rps_context.CompetitionEntries
                     Where entries.Competition_Date_1.Equals(compDate)
                     Select entries.Theme Distinct

            For Each _record In _query
                SelectTheme.Items.Add(_record)
            Next
            If SelectTheme.Items.Count > 0 Then
                SelectTheme.SelectedIndex = 0
            Else
                SelectTheme.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error in LoadUniqueThemes()")
        End Try
    End Sub

    Private Sub NumNinesHeadingButton_Click(sender As Object, e As EventArgs) _
        Handles NumNinesHeadingButton.Click
        'NineScoreRadioButton.Checked = True
        _all_scores_selected = False
        _eights_and_awards_selected = False
        _selected_avg_score = 9
        _selected_score = 0
        SelectImages()
        PickAwards()
    End Sub

    Private Sub NumEightsHeadingButton_Click(sender As Object, e As EventArgs) _
        Handles NumEightsHeadingButton.Click
        'EightScoreRadioButton.Checked = True
        _all_scores_selected = False
        _eights_and_awards_selected = False
        _selected_avg_score = 8
        _selected_score = 0
        SelectImages()
        PickAwards()
    End Sub

    Private Sub NumSevensHeadingButton_Click(sender As Object, e As EventArgs) _
        Handles NumSevensHeadingButton.Click
        'SevenScoreRadioButton.Checked = True
        _all_scores_selected = False
        _eights_and_awards_selected = False
        _selected_avg_score = 7
        _selected_score = 0
        SelectImages()
        PickAwards()
    End Sub

    Private Sub AwardsTableTitleBar_Click(sender As Object, e As EventArgs) _
        Handles AwardsTableTitleBar.Click
        'EightsAndAwardsRadioButton.Checked = True
        _eights_and_awards_selected = True
        _all_scores_selected = False
        SelectImages()
        DoSlideShow()
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
        SelectImages()
    End Sub

    Private Sub SelectDate_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles SelectDate.SelectedIndexChanged
        LoadUniqueThemes()
        SelectScore.SelectedItem = "All"
        _all_scores_selected = True
        _eights_and_awards_selected = False
        SelectImages()
    End Sub

    Private Function CreateDataTable(sourceTable As DataTable, rows As DataRow()) As DataTable
        Dim result As DataTable

        result = sourceTable.Clone()
        For Each row As DataRow In rows
            result.Rows.Add(row.ItemArray)
        Next
        Return result
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
