Public Class ProgressStatus : Inherits StatusBar
    Public progressBar As New ProgressBar
    Private _progressBar As Integer = -1

    Sub New()
        progressBar.Hide()

        Me.Controls.Add(progressBar)
    End Sub

    Public Property setProgressBar() As Integer
        Get
            Return _progressBar
        End Get
        Set(ByVal Value As Integer)
            _progressBar = Value
            Me.Panels(_progressBar).Style = StatusBarPanelStyle.OwnerDraw
        End Set
    End Property

    Private Sub Reposition(ByVal sender As Object, ByVal sbdevent As StatusBarDrawItemEventArgs) Handles MyBase.DrawItem
        progressBar.Location = New Point(sbdevent.Bounds.X, sbdevent.Bounds.Y)
        progressBar.Size = New Size(sbdevent.Bounds.Width, sbdevent.Bounds.Height)
        progressBar.Show()
    End Sub
End Class