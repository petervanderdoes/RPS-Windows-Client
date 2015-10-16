Public Class ProgressStatus : Inherits StatusBar
    Public progress_bar As New ProgressBar
    Private _progressBar As Integer = -1

    Sub New()
        progress_bar.Hide()

        Me.Controls.Add(progress_bar)
    End Sub

    Public Property setProgressBar As Integer
        Get
            Return _progressBar
        End Get
        Set
            _progressBar = Value
            Me.Panels(_progressBar).Style = StatusBarPanelStyle.OwnerDraw
        End Set
    End Property

    Private Sub Reposition(ByVal sender As Object, ByVal sbdevent As StatusBarDrawItemEventArgs) Handles MyBase.DrawItem
        progress_bar.Location = New Point(sbdevent.Bounds.X, sbdevent.Bounds.Y)
        progress_bar.Size = New Size(sbdevent.Bounds.Width, sbdevent.Bounds.Height)
        progress_bar.Show()
    End Sub
End Class
