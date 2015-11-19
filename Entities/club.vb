

Namespace Entities
    <System.ComponentModel.DataAnnotations.Schema.Table("club")>
    Partial Public Class club
        Public Sub New()
            club_award = New Generic.HashSet(Of club_award)()
            club_classification = New Generic.HashSet(Of club_classification)()
            club_medium = New Generic.HashSet(Of club_medium)()
        End Sub

        <System.ComponentModel.DataAnnotations.StringLength(2147483647)>
        Public Property short_name As String

        <System.ComponentModel.DataAnnotations.StringLength(2147483647)>
        Public Property name As String

        Public Property id As Long

        Public Property min_score As Long?

        Public Property max_score As Long?

        Public Property min_score_for_award As Long?

        Public Overridable Property club_award As Generic.ICollection(Of club_award)

        Public Overridable Property club_classification As Generic.ICollection(Of club_classification)

        Public Overridable Property club_medium As Generic.ICollection(Of club_medium)
    End Class
End Namespace
