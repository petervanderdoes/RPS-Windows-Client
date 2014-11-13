Public Class BestOfDialog
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents dtSeasonEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lbSelectedAwards As System.Windows.Forms.ListBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbTheme As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(BestOfDialog))
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtSeasonEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtStartDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtEndDate = New System.Windows.Forms.DateTimePicker
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.lbSelectedAwards = New System.Windows.Forms.ListBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbTheme = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 28)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "BestOf Competition Date"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtSeasonEndDate
        '
        Me.dtSeasonEndDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtSeasonEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtSeasonEndDate.Location = New System.Drawing.Point(259, 18)
        Me.dtSeasonEndDate.Name = "dtSeasonEndDate"
        Me.dtSeasonEndDate.Size = New System.Drawing.Size(211, 26)
        Me.dtSeasonEndDate.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(10, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(240, 28)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Starting Competition Date"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtStartDate
        '
        Me.dtStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtStartDate.Location = New System.Drawing.Point(259, 55)
        Me.dtStartDate.Name = "dtStartDate"
        Me.dtStartDate.Size = New System.Drawing.Size(211, 26)
        Me.dtStartDate.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(10, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(240, 28)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Ending Competition Date"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtEndDate
        '
        Me.dtEndDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtEndDate.Location = New System.Drawing.Point(259, 92)
        Me.dtEndDate.Name = "dtEndDate"
        Me.dtEndDate.Size = New System.Drawing.Size(211, 26)
        Me.dtEndDate.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(130, 288)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(96, 28)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(254, 288)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(96, 28)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(10, 166)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(240, 28)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Selected Awards"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbSelectedAwards
        '
        Me.lbSelectedAwards.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbSelectedAwards.ItemHeight = 20
        Me.lbSelectedAwards.Location = New System.Drawing.Point(259, 165)
        Me.lbSelectedAwards.Name = "lbSelectedAwards"
        Me.lbSelectedAwards.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lbSelectedAwards.Size = New System.Drawing.Size(211, 104)
        Me.lbSelectedAwards.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(10, 129)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(240, 28)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "BestOf Theme"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tbTheme
        '
        Me.tbTheme.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbTheme.Location = New System.Drawing.Point(259, 128)
        Me.tbTheme.Name = "tbTheme"
        Me.tbTheme.Size = New System.Drawing.Size(211, 26)
        Me.tbTheme.TabIndex = 12
        Me.tbTheme.Text = "Best Of Month"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(112, 200)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 48)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Use Ctrl-Click to select multiple awards"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BestOfDialog
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(489, 330)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.tbTheme)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lbSelectedAwards)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dtEndDate)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtStartDate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtSeasonEndDate)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "BestOfDialog"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Create BestOf Competition"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub BestOfDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim InvalidInput As Boolean = False
        Dim errMsg As String = "Please correct the following errors" + vbCrLf + vbCrLf

        If Me.DialogResult = DialogResult.OK Then
            If Me.dtStartDate.Value > Me.dtEndDate.Value Then
                errMsg = errMsg + "* Starting competition date must be earlier than end date" + vbCrLf
                InvalidInput = True
            End If
            If Me.dtStartDate.Value > Me.dtSeasonEndDate.Value Then
                errMsg = errMsg + "* Starting Competition Date must be earlier than Year End date" + vbCrLf
                InvalidInput = True
            End If
            If Me.dtEndDate.Value > Me.dtSeasonEndDate.Value Then
                errMsg = errMsg + "* Ending Competition Date must not be later than BestOf date" + vbCrLf
                InvalidInput = True
            End If
            If Me.tbTheme.Text = "" Then
                errMsg = errMsg + "* Theme must not be blank" + vbCrLf
                InvalidInput = True
            End If
            If Me.lbSelectedAwards.SelectedItems.Count < 1 Then
                errMsg = errMsg + "* At least one type of award must be selected" + vbCrLf
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


    Private Sub BestOfDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer

        For i = 0 To theMainForm.awards.Count - 1
            lbSelectedAwards.Items.Add(theMainForm.awards.Item(i))
        Next
    End Sub
End Class
