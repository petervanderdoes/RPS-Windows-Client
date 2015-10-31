Namespace Components
    Public Class ProgressStatus
        Inherits StatusBar
        Public progress_bar As New ProgressBar
        Private _progress_bar As Integer = -1

        Sub New()
            progress_bar.Hide()

            Controls.Add(progress_bar)
        End Sub

        Public Property setProgressBar As Integer
            Get
                Return _progress_bar
            End Get
            Set
                _progress_bar = Value
                Panels(_progress_bar).Style = StatusBarPanelStyle.OwnerDraw
            End Set
        End Property

        Private Sub Reposition(sender As Object, sbdevent As StatusBarDrawItemEventArgs) Handles MyBase.DrawItem
            progress_bar.Location = New Point(sbdevent.Bounds.X, sbdevent.Bounds.Y)
            progress_bar.Size = New Size(sbdevent.Bounds.Width, sbdevent.Bounds.Height)
            progress_bar.Show()
        End Sub
    End Class
End Namespace
