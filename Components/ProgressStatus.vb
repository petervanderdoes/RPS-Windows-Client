Namespace Components
    Public Class ProgressStatus
        Inherits StatusBar
        Public ReadOnly progress_bar As New ProgressBar
        Private progress_bar_panel As Integer = -1

        Sub New()
            progress_bar.Hide()

            Controls.Add(progress_bar)
        End Sub

        Public WriteOnly Property progressBarPanel As Integer
            Set
                progress_bar_panel = Value
                Panels(progress_bar_panel).Style = StatusBarPanelStyle.OwnerDraw
            End Set
        End Property

        Private Sub Reposition(sender As Object, sbdevent As StatusBarDrawItemEventArgs) Handles MyBase.DrawItem
            progress_bar.Location = New Point(sbdevent.Bounds.X, sbdevent.Bounds.Y)
            progress_bar.Size = New Size(sbdevent.Bounds.Width, sbdevent.Bounds.Height)
            progress_bar.Show()
        End Sub
    End Class
End Namespace
