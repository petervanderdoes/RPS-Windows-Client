Imports System.IO

Public Class RPS_Data_Form
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents OleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents objSelectedPhotos As PictureViewer.SelectedPhotos
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnCancelAll As System.Windows.Forms.Button
    Friend WithEvents grdCompetition_Entries As System.Windows.Forms.DataGrid
    Friend WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn4 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn5 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSlideShow As System.Windows.Forms.Button
    Friend WithEvents SevenScoreRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents EightScoreRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents NineScoreRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents AllScoresRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents AwardsCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDataAdapter1 = New System.Data.OleDb.OleDbDataAdapter
        Me.btnLoad = New System.Windows.Forms.Button
        Me.btnUpdate = New System.Windows.Forms.Button
        Me.btnCancelAll = New System.Windows.Forms.Button
        Me.grdCompetition_Entries = New System.Windows.Forms.DataGrid
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle
        Me.DataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn4 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn5 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SevenScoreRadioButton = New System.Windows.Forms.RadioButton
        Me.EightScoreRadioButton = New System.Windows.Forms.RadioButton
        Me.NineScoreRadioButton = New System.Windows.Forms.RadioButton
        Me.AllScoresRadioButton = New System.Windows.Forms.RadioButton
        Me.btnSlideShow = New System.Windows.Forms.Button
        Me.AwardsCheckBox = New System.Windows.Forms.CheckBox
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        CType(Me.grdCompetition_Entries, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT Award, [Banquet Award], [Banquet Score], Classification, [Competition Date" & _
        " 1], [Competition Date 2], [Image File Name], Maker, Medium, Photo_ID, [Score 1]" & _
        ", [Score 2], Theme, Title FROM [Competition Entries]"
        Me.OleDbSelectCommand1.Connection = Me.OleDbConnection1
        '
        'OleDbConnection1
        '
        Me.OleDbConnection1.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Jet OLEDB:Database Password=;Data Source=""C:\Jake\Access\rps\rps.m" & _
        "db"";Password=;Jet OLEDB:Engine Type=5;Jet OLEDB:Global Bulk Transactions=1;Provi" & _
        "der=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=False;Ext" & _
        "ended Properties=;Mode=Share Deny None;Jet OLEDB:New Database Password=;Jet OLED" & _
        "B:Create System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet " & _
        "OLEDB:Compact Without Replica Repair=False;User ID=Admin;Jet OLEDB:Encrypt Datab" & _
        "ase=False"
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO [Competition Entries] (Award, [Banquet Award], [Banquet Score], Class" & _
        "ification, [Competition Date 1], [Competition Date 2], [Image File Name], Maker," & _
        " Medium, [Score 1], [Score 2], Theme, Title) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, " & _
        "?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.OleDbConnection1
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Award", System.Data.OleDb.OleDbType.VarWChar, 50, "Award"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Banquet_Award", System.Data.OleDb.OleDbType.VarWChar, 50, "Banquet Award"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Banquet_Score", System.Data.OleDb.OleDbType.Integer, 0, "Banquet Score"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Classification", System.Data.OleDb.OleDbType.VarWChar, 50, "Classification"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Competition_Date_1", System.Data.OleDb.OleDbType.DBDate, 0, "Competition Date 1"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Competition_Date_2", System.Data.OleDb.OleDbType.VarWChar, 50, "Competition Date 2"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Image_File_Name", System.Data.OleDb.OleDbType.VarWChar, 128, "Image File Name"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Maker", System.Data.OleDb.OleDbType.VarWChar, 128, "Maker"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Medium", System.Data.OleDb.OleDbType.VarWChar, 50, "Medium"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Score_1", System.Data.OleDb.OleDbType.Integer, 0, "Score 1"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Score_2", System.Data.OleDb.OleDbType.Integer, 0, "Score 2"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Theme", System.Data.OleDb.OleDbType.VarWChar, 128, "Theme"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Title", System.Data.OleDb.OleDbType.VarWChar, 128, "Title"))
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE [Competition Entries] SET Award = ?, [Banquet Award] = ?, [Banquet Score] " & _
        "= ?, Classification = ?, [Competition Date 1] = ?, [Competition Date 2] = ?, [Im" & _
        "age File Name] = ?, Maker = ?, Medium = ?, [Score 1] = ?, [Score 2] = ?, Theme =" & _
        " ?, Title = ? WHERE (Photo_ID = ?) AND (Award = ? OR ? IS NULL AND Award IS NULL" & _
        ") AND ([Banquet Award] = ? OR ? IS NULL AND [Banquet Award] IS NULL) AND ([Banqu" & _
        "et Score] = ? OR ? IS NULL AND [Banquet Score] IS NULL) AND (Classification = ? " & _
        "OR ? IS NULL AND Classification IS NULL) AND ([Competition Date 1] = ? OR ? IS N" & _
        "ULL AND [Competition Date 1] IS NULL) AND ([Competition Date 2] = ? OR ? IS NULL" & _
        " AND [Competition Date 2] IS NULL) AND ([Image File Name] = ? OR ? IS NULL AND [" & _
        "Image File Name] IS NULL) AND (Maker = ? OR ? IS NULL AND Maker IS NULL) AND (Me" & _
        "dium = ? OR ? IS NULL AND Medium IS NULL) AND ([Score 1] = ? OR ? IS NULL AND [S" & _
        "core 1] IS NULL) AND ([Score 2] = ? OR ? IS NULL AND [Score 2] IS NULL) AND (The" & _
        "me = ? OR ? IS NULL AND Theme IS NULL) AND (Title = ? OR ? IS NULL AND Title IS " & _
        "NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.OleDbConnection1
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Award", System.Data.OleDb.OleDbType.VarWChar, 50, "Award"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Banquet_Award", System.Data.OleDb.OleDbType.VarWChar, 50, "Banquet Award"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Banquet_Score", System.Data.OleDb.OleDbType.Integer, 0, "Banquet Score"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Classification", System.Data.OleDb.OleDbType.VarWChar, 50, "Classification"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Competition_Date_1", System.Data.OleDb.OleDbType.DBDate, 0, "Competition Date 1"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Competition_Date_2", System.Data.OleDb.OleDbType.VarWChar, 50, "Competition Date 2"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Image_File_Name", System.Data.OleDb.OleDbType.VarWChar, 128, "Image File Name"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Maker", System.Data.OleDb.OleDbType.VarWChar, 128, "Maker"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Medium", System.Data.OleDb.OleDbType.VarWChar, 50, "Medium"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Score_1", System.Data.OleDb.OleDbType.Integer, 0, "Score 1"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Score_2", System.Data.OleDb.OleDbType.Integer, 0, "Score 2"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Theme", System.Data.OleDb.OleDbType.VarWChar, 128, "Theme"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Title", System.Data.OleDb.OleDbType.VarWChar, 128, "Title"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Photo_ID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Photo_ID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Award", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Award", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Award1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Award", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Banquet_Award", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Banquet Award", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Banquet_Award1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Banquet Award", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Banquet_Score", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Banquet Score", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Banquet_Score1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Banquet Score", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Classification", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Classification", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Classification1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Classification", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Competition_Date_1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Competition Date 1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Competition_Date_11", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Competition Date 1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Competition_Date_2", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Competition Date 2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Competition_Date_21", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Competition Date 2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Image_File_Name", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Image File Name", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Image_File_Name1", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Image File Name", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Maker", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Maker", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Maker1", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Maker", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Medium", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Medium", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Medium1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Medium", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Score_1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Score 1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Score_11", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Score 1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Score_2", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Score 2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Score_21", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Score 2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Theme", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Theme", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Theme1", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Theme", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Title", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Title", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Title1", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Title", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM [Competition Entries] WHERE (Photo_ID = ?) AND (Award = ? OR ? IS NUL" & _
        "L AND Award IS NULL) AND ([Banquet Award] = ? OR ? IS NULL AND [Banquet Award] I" & _
        "S NULL) AND ([Banquet Score] = ? OR ? IS NULL AND [Banquet Score] IS NULL) AND (" & _
        "Classification = ? OR ? IS NULL AND Classification IS NULL) AND ([Competition Da" & _
        "te 1] = ? OR ? IS NULL AND [Competition Date 1] IS NULL) AND ([Competition Date " & _
        "2] = ? OR ? IS NULL AND [Competition Date 2] IS NULL) AND ([Image File Name] = ?" & _
        " OR ? IS NULL AND [Image File Name] IS NULL) AND (Maker = ? OR ? IS NULL AND Mak" & _
        "er IS NULL) AND (Medium = ? OR ? IS NULL AND Medium IS NULL) AND ([Score 1] = ? " & _
        "OR ? IS NULL AND [Score 1] IS NULL) AND ([Score 2] = ? OR ? IS NULL AND [Score 2" & _
        "] IS NULL) AND (Theme = ? OR ? IS NULL AND Theme IS NULL) AND (Title = ? OR ? IS" & _
        " NULL AND Title IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.OleDbConnection1
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Photo_ID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Photo_ID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Award", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Award", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Award1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Award", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Banquet_Award", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Banquet Award", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Banquet_Award1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Banquet Award", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Banquet_Score", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Banquet Score", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Banquet_Score1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Banquet Score", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Classification", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Classification", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Classification1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Classification", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Competition_Date_1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Competition Date 1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Competition_Date_11", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Competition Date 1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Competition_Date_2", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Competition Date 2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Competition_Date_21", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Competition Date 2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Image_File_Name", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Image File Name", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Image_File_Name1", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Image File Name", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Maker", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Maker", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Maker1", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Maker", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Medium", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Medium", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Medium1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Medium", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Score_1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Score 1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Score_11", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Score 1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Score_2", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Score 2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Score_21", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Score 2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Theme", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Theme", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Theme1", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Theme", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Title", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Title", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Title1", System.Data.OleDb.OleDbType.VarWChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Title", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDataAdapter1
        '
        Me.OleDbDataAdapter1.DeleteCommand = Me.OleDbDeleteCommand1
        Me.OleDbDataAdapter1.InsertCommand = Me.OleDbInsertCommand1
        Me.OleDbDataAdapter1.SelectCommand = Me.OleDbSelectCommand1
        Me.OleDbDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Competition Entries", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("Award", "Award"), New System.Data.Common.DataColumnMapping("Banquet Award", "Banquet Award"), New System.Data.Common.DataColumnMapping("Banquet Score", "Banquet Score"), New System.Data.Common.DataColumnMapping("Classification", "Classification"), New System.Data.Common.DataColumnMapping("Competition Date 1", "Competition Date 1"), New System.Data.Common.DataColumnMapping("Competition Date 2", "Competition Date 2"), New System.Data.Common.DataColumnMapping("Image File Name", "Image File Name"), New System.Data.Common.DataColumnMapping("Maker", "Maker"), New System.Data.Common.DataColumnMapping("Medium", "Medium"), New System.Data.Common.DataColumnMapping("Photo_ID", "Photo_ID"), New System.Data.Common.DataColumnMapping("Score 1", "Score 1"), New System.Data.Common.DataColumnMapping("Score 2", "Score 2"), New System.Data.Common.DataColumnMapping("Theme", "Theme"), New System.Data.Common.DataColumnMapping("Title", "Title")})})
        Me.OleDbDataAdapter1.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(16, 328)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.TabIndex = 0
        Me.btnLoad.Text = "&Load"
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(416, 464)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.TabIndex = 1
        Me.btnUpdate.Text = "&Update"
        '
        'btnCancelAll
        '
        Me.btnCancelAll.Location = New System.Drawing.Point(536, 464)
        Me.btnCancelAll.Name = "btnCancelAll"
        Me.btnCancelAll.TabIndex = 2
        Me.btnCancelAll.Text = "Ca&ncel All"
        '
        'grdCompetition_Entries
        '
        Me.grdCompetition_Entries.CaptionFont = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.grdCompetition_Entries.DataMember = "Competition Entries"
        Me.grdCompetition_Entries.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdCompetition_Entries.Location = New System.Drawing.Point(216, 40)
        Me.grdCompetition_Entries.Name = "grdCompetition_Entries"
        Me.grdCompetition_Entries.Size = New System.Drawing.Size(616, 408)
        Me.grdCompetition_Entries.TabIndex = 3
        Me.grdCompetition_Entries.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.DataGrid = Me.grdCompetition_Entries
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn2, Me.DataGridTextBoxColumn3, Me.DataGridTextBoxColumn4, Me.DataGridTextBoxColumn5})
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = "Competition Entries"
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Format = ""
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.HeaderText = "Award"
        Me.DataGridTextBoxColumn2.MappingName = "Award"
        Me.DataGridTextBoxColumn2.NullText = ""
        Me.DataGridTextBoxColumn2.Width = 50
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.HeaderText = "Score"
        Me.DataGridTextBoxColumn3.MappingName = "Score 1"
        Me.DataGridTextBoxColumn3.NullText = ""
        Me.DataGridTextBoxColumn3.Width = 50
        '
        'DataGridTextBoxColumn4
        '
        Me.DataGridTextBoxColumn4.Format = ""
        Me.DataGridTextBoxColumn4.FormatInfo = Nothing
        Me.DataGridTextBoxColumn4.HeaderText = "Title"
        Me.DataGridTextBoxColumn4.MappingName = "Title"
        Me.DataGridTextBoxColumn4.Width = 250
        '
        'DataGridTextBoxColumn5
        '
        Me.DataGridTextBoxColumn5.Format = ""
        Me.DataGridTextBoxColumn5.FormatInfo = Nothing
        Me.DataGridTextBoxColumn5.HeaderText = "Maker"
        Me.DataGridTextBoxColumn5.MappingName = "Maker"
        Me.DataGridTextBoxColumn5.Width = 150
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 16)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Competition Date"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker1.Location = New System.Drawing.Point(16, 64)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(176, 23)
        Me.DateTimePicker1.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(176, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Medium"
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.Items.AddRange(New Object() {"B&W Prints", "Color Prints", "Slides", "Digital"})
        Me.ComboBox1.Location = New System.Drawing.Point(16, 120)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(176, 24)
        Me.ComboBox1.TabIndex = 9
        Me.ComboBox1.Text = "Digital"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(168, 16)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Classification"
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.Items.AddRange(New Object() {"Beginner", "Advanced", "Salon"})
        Me.ComboBox2.Location = New System.Drawing.Point(16, 176)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(176, 24)
        Me.ComboBox2.TabIndex = 11
        Me.ComboBox2.Text = "Beginner"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.SevenScoreRadioButton)
        Me.GroupBox1.Controls.Add(Me.EightScoreRadioButton)
        Me.GroupBox1.Controls.Add(Me.NineScoreRadioButton)
        Me.GroupBox1.Controls.Add(Me.AllScoresRadioButton)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(16, 216)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(176, 64)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Score"
        '
        'SevenScoreRadioButton
        '
        Me.SevenScoreRadioButton.CheckAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.SevenScoreRadioButton.Location = New System.Drawing.Point(128, 16)
        Me.SevenScoreRadioButton.Name = "SevenScoreRadioButton"
        Me.SevenScoreRadioButton.Size = New System.Drawing.Size(32, 32)
        Me.SevenScoreRadioButton.TabIndex = 3
        Me.SevenScoreRadioButton.Text = "7"
        Me.SevenScoreRadioButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'EightScoreRadioButton
        '
        Me.EightScoreRadioButton.CheckAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.EightScoreRadioButton.Location = New System.Drawing.Point(88, 16)
        Me.EightScoreRadioButton.Name = "EightScoreRadioButton"
        Me.EightScoreRadioButton.Size = New System.Drawing.Size(32, 32)
        Me.EightScoreRadioButton.TabIndex = 2
        Me.EightScoreRadioButton.Text = "8"
        Me.EightScoreRadioButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'NineScoreRadioButton
        '
        Me.NineScoreRadioButton.CheckAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.NineScoreRadioButton.Location = New System.Drawing.Point(48, 16)
        Me.NineScoreRadioButton.Name = "NineScoreRadioButton"
        Me.NineScoreRadioButton.Size = New System.Drawing.Size(32, 32)
        Me.NineScoreRadioButton.TabIndex = 1
        Me.NineScoreRadioButton.Text = "9"
        Me.NineScoreRadioButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'AllScoresRadioButton
        '
        Me.AllScoresRadioButton.CheckAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.AllScoresRadioButton.Checked = True
        Me.AllScoresRadioButton.Location = New System.Drawing.Point(8, 16)
        Me.AllScoresRadioButton.Name = "AllScoresRadioButton"
        Me.AllScoresRadioButton.Size = New System.Drawing.Size(33, 32)
        Me.AllScoresRadioButton.TabIndex = 0
        Me.AllScoresRadioButton.TabStop = True
        Me.AllScoresRadioButton.Text = "Any"
        Me.AllScoresRadioButton.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnSlideShow
        '
        Me.btnSlideShow.Location = New System.Drawing.Point(120, 328)
        Me.btnSlideShow.Name = "btnSlideShow"
        Me.btnSlideShow.TabIndex = 13
        Me.btnSlideShow.Text = "Slide Show"
        '
        'AwardsCheckBox
        '
        Me.AwardsCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AwardsCheckBox.Location = New System.Drawing.Point(16, 288)
        Me.AwardsCheckBox.Name = "AwardsCheckBox"
        Me.AwardsCheckBox.Size = New System.Drawing.Size(176, 24)
        Me.AwardsCheckBox.TabIndex = 14
        Me.AwardsCheckBox.Text = "Award Winners"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem4, Me.MenuItem7})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3})
        Me.MenuItem1.Text = "File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Import Images ..."
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "Exit"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5, Me.MenuItem6})
        Me.MenuItem4.Text = "Reports"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Text = "Worksheet"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 1
        Me.MenuItem6.Text = "Results Report"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 2
        Me.MenuItem7.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem8})
        Me.MenuItem7.Text = "Help"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 0
        Me.MenuItem8.Text = "About"
        '
        'MainForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(840, 502)
        Me.Controls.Add(Me.AwardsCheckBox)
        Me.Controls.Add(Me.btnSlideShow)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnCancelAll)
        Me.Controls.Add(Me.grdCompetition_Entries)
        Me.Menu = Me.MainMenu1
        Me.Name = "MainForm"
        Me.Text = "RPS_Digital_Viewer"
        CType(Me.grdCompetition_Entries, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try
            'Attempt to update the datasource.
            Me.UpdateDataSet()
        Catch eUpdate As System.Exception
            'Add your error handling code here.
            'Display error message, if any.
            System.Windows.Forms.MessageBox.Show(eUpdate.Message)
        End Try

    End Sub
    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Dim numSelected As Double
        Dim where_clause As String
        Dim select_stmt As String
        Dim DataGridCaptionText = DateTimePicker1.Text + " - " + ComboBox2.Text + " / " + ComboBox1.Text

        ' Build the complete SQL statement to select the records specified
        ' by the selection criteria on the main form
        select_stmt = "SELECT Award, [Banquet Award], [Banquet Score], Classification, [Competition Date 1], [Competition Date 2], [Image File Name], Maker, Medium, Photo_ID, [Score 1], [Score 2], Theme, Title FROM [Competition Entries]"
        where_clause = " WHERE " + "[Competition Date 1]=#" + DateTimePicker1.Text + "# AND " + _
                        "Medium=""" + ComboBox1.Text + """ AND " + _
                        "Classification=""" + ComboBox2.Text + """"
        If AwardsCheckBox.Checked Then
            where_clause = where_clause + " AND NOT (Award IS NULL)"
            DataGridCaptionText = DataGridCaptionText + " - Award Winners"
        ElseIf Not AllScoresRadioButton.Checked Then
            If NineScoreRadioButton.Checked Then
                where_clause = where_clause + " AND [Score 1]=9"
                DataGridCaptionText = DataGridCaptionText + " - 9 points"
            ElseIf EightScoreRadioButton.Checked Then
                where_clause = where_clause + " AND [Score 1]=8"
                DataGridCaptionText = DataGridCaptionText + " - 8 points"
            ElseIf SevenScoreRadioButton.Checked Then
                where_clause = where_clause + " AND [Score 1]=7"
                DataGridCaptionText = DataGridCaptionText + " - 7 points"
            End If
        End If
        'MsgBox(select_stmt + where_clause)
        ' Install the updated SQL SELECT statement
        OleDbSelectCommand1.CommandText = select_stmt + where_clause

        Try
            'Attempt to load the dataset.
            Me.LoadDataSet()
            'MsgBox("Selected " + CStr(objSelectedPhotos.Tables("Competition Entries").Rows.Count) + " rows")
            'Dim r As DataRow
            'Dim temp As String
            'For Each r In objSelectedPhotos.Tables("Competition Entries").Rows
            '   temp = r("Maker")
            'Next

            ' Count the number of rows selected and add it to the cation of the DataGrid
            numSelected = CDbl(objSelectedPhotos.Tables("Competition Entries").Rows.Count)
            grdCompetition_Entries.CaptionText = DataGridCaptionText + " - " + _
                CStr(Int(numSelected)) + " Entries; " + CStr(Int((numSelected / 4) + 0.5)) + " Awards"
        Catch eLoad As System.Exception
            'Add your error handling code here.
            'Display error message, if any.
            System.Windows.Forms.MessageBox.Show(eLoad.Message)
        End Try

    End Sub
    Private Sub btnCancelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelAll.Click
        Me.objSelectedPhotos.RejectChanges()

    End Sub
    Public Sub UpdateDataSet()
        'Create a new dataset to hold the changes that have been made to the main dataset.
        Dim objDataSetChanges As PictureViewer.SelectedPhotos = New PictureViewer.SelectedPhotos
        'Stop any current edits.
        Me.BindingContext(objSelectedPhotos, "Competition Entries").EndCurrentEdit()
        'Get the changes that have been made to the main dataset.
        objDataSetChanges = CType(objSelectedPhotos.GetChanges, PictureViewer.SelectedPhotos)
        'Check to see if any changes have been made.
        If (Not (objDataSetChanges) Is Nothing) Then
            Try
                'There are changes that need to be made, so attempt to update the datasource by
                'calling the update method and passing the dataset and any parameters.
                Me.UpdateDataSource(objDataSetChanges)
                objSelectedPhotos.Merge(objDataSetChanges)
                objSelectedPhotos.AcceptChanges()
            Catch eUpdate As System.Exception
                'Add your error handling code here.
                Throw eUpdate
            End Try
            'Add your code to check the returned dataset for any errors that may have been
            'pushed into the row object's error.
        End If

    End Sub
    Public Sub LoadDataSet()
        'Create a new dataset to hold the records returned from the call to FillDataSet.
        'A temporary dataset is used because filling the existing dataset would
        'require the databindings to be rebound.
        Dim objDataSetTemp As PictureViewer.SelectedPhotos
        objDataSetTemp = New PictureViewer.SelectedPhotos
        Try
            'Attempt to fill the temporary dataset.
            Me.FillDataSet(objDataSetTemp)
        Catch eFillDataSet As System.Exception
            'Add your error handling code here.
            Throw eFillDataSet
        End Try
        Try
            grdCompetition_Entries.DataSource = Nothing
            'Empty the old records from the dataset.
            objSelectedPhotos.Clear()
            'Merge the records into the main dataset.
            objSelectedPhotos.Merge(objDataSetTemp)
            grdCompetition_Entries.SetDataBinding(objSelectedPhotos, "Competition Entries")
        Catch eLoadMerge As System.Exception
            'Add your error handling code here.
            Throw eLoadMerge
        End Try

    End Sub
    Public Sub UpdateDataSource(ByVal ChangedRows As PictureViewer.SelectedPhotos)
        Try
            'The data source only needs to be updated if there are changes pending.
            If (Not (ChangedRows) Is Nothing) Then
                'Open the connection.
                Me.OleDbConnection1.Open()
                'Attempt to update the data source.
                OleDbDataAdapter1.Update(ChangedRows)
            End If
        Catch updateException As System.Exception
            'Add your error handling code here.
            Throw updateException
        Finally
            'Close the connection whether or not the exception was thrown.
            Me.OleDbConnection1.Close()
        End Try

    End Sub
    Public Sub FillDataSet(ByVal dataSet As PictureViewer.SelectedPhotos)
        'Turn off constraint checking before the dataset is filled.
        'This allows the adapters to fill the dataset without concern
        'for dependencies between the tables.
        dataSet.EnforceConstraints = False
        Try
            'Open the connection.
            Me.OleDbConnection1.Open()
            'Attempt to fill the dataset through the OleDbDataAdapter1.
            Me.OleDbDataAdapter1.Fill(dataSet)
        Catch fillException As System.Exception
            'Add your error handling code here.
            Throw fillException
        Finally
            'Turn constraint checking back on.
            dataSet.EnforceConstraints = True
            'Close the connection whether or not the exception was thrown.
            Me.OleDbConnection1.Close()
        End Try

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Me.Close()
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim Import_Dialog As New Import_Images_Dialog
        Dim competitionDate As String
        Dim competitionTheme As String
        Dim medium As String
        Dim classification As String
        Dim title As String
        Dim maker As String
        Dim folderPath As String
        Dim dRow As DataRow

        ' Open the dialog
        Import_Dialog.ShowDialog(Me)
        If Import_Dialog.DialogResult = DialogResult.OK Then
            competitionDate = Import_Dialog.dpCompetitionDate.Text()
            competitionTheme = Import_Dialog.tbTheme.Text()
            medium = Import_Dialog.cbMedium.Text()
            classification = Import_Dialog.cbClassification.Text()
            folderPath = Import_Dialog.tbNewImageFolder.Text()
        End If
        ' Make sure the folder exists
        Dim dirInfo As New DirectoryInfo(folderPath)
        If Not dirInfo.Exists Then
            MsgBox("Folder """ + folderPath + """ doesn't exist")
            Exit Sub
        End If

        ' Iterate though all the .jpg files in the folder
        Dim files() As FileInfo
        Dim file As FileInfo
        Dim fileFullName As String
        Dim fileName As String
        Dim fields()
        Try
            files = dirInfo.GetFiles
            For Each file In files
                ' Save the full file path and name
                fileFullName = file.FullName
                ' Parse the title and maker out of the file name
                fileName = file.Name
                fileName = Mid(fileName, 1, InStr(1, fileName, ".") - 1) ' Strip off the extension
                fields = Split(fileName, "+")
                title = fields(0)
                maker = fields(1)
                ' Insert a new row in to the database table
                dRow = objSelectedPhotos.Tables("Competition Entries").NewRow
                dRow("Title") = title
                dRow("Maker") = maker
                dRow("Classification") = classification
                dRow("Medium") = medium
                dRow("Theme") = competitionTheme
                dRow("Competition Date 1") = competitionDate
                dRow("Image File Name") = fileFullName
                objSelectedPhotos.Tables("Competition Entries").Rows.Add(dRow)
                OleDbDataAdapter1.Update(objSelectedPhotos)
            Next
        Catch ex As Exception
            MsgBox("Error importing image files" + vbCrLf + ex.ToString)
        End Try
    End Sub

    Private Sub btnSlideShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlideShow.Click
        Dim Viewer As New ImageViewer(objSelectedPhotos)
        Viewer.ShowDialog()
        Try
            'Attempt to update the datasource.
            Me.UpdateDataSet()
        Catch eUpdate As System.Exception
            'Add your error handling code here.
            'Display error message, if any.
            System.Windows.Forms.MessageBox.Show(eUpdate.Message)
        End Try
    End Sub
End Class
