Namespace Forms
    Public Class About
        Inherits Form

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
        Friend WithEvents Label1 As Label
        Friend WithEvents Button1 As System.Windows.Forms.Button

        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.Label1 = New Label()
            Me.Button1 = New System.Windows.Forms.Button()
            Me.SuspendLayout()
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Font = New System.Drawing.Font("Segoe UI",
                                                     12.0!,
                                                     System.Drawing.FontStyle.Regular,
                                                     System.Drawing.GraphicsUnit.Point,
                                                     CType(0, Byte))
            Me.Label1.Location = New Point(16, 16)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New Size(0, 21)
            Me.Label1.TabIndex = 0
            Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
            '
            'Button1
            '
            Me.Button1.Anchor = AnchorStyles.Bottom
            Me.Button1.Location = New Point(109, 145)
            Me.Button1.Name = "Button1"
            Me.Button1.Size = New Size(75, 23)
            Me.Button1.TabIndex = 1
            Me.Button1.Text = "OK"
            '
            'About
            '
            Me.AutoScaleBaseSize = New Size(6, 16)
            Me.AutoSize = True
            Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.ClientSize = New Size(292, 182)
            Me.Controls.Add(Me.Button1)
            Me.Controls.Add(Me.Label1)
            Me.Font = New System.Drawing.Font("Segoe UI",
                                              9.0!,
                                              System.Drawing.FontStyle.Regular,
                                              System.Drawing.GraphicsUnit.Point,
                                              CType(0, Byte))
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.Name = "About"
            Me.Padding = New System.Windows.Forms.Padding(11)
            Me.Text = "About"
            Me.ResumeLayout(False)
            Me.PerformLayout()
        End Sub

#End Region

        Private Sub About_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            Label1.Text = "RPS Competition Client" + vbCrLf +
                          "Version " + Application.ProductVersion + vbCrLf + vbCrLf +
                          "(c) Jake Chapple, Peter van der Does" + vbCrLf +
                          "Raritan Photographic Society"
        End Sub

        Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
            Close()
        End Sub
    End Class
End Namespace
