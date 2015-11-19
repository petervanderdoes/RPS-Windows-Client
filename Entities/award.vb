

Namespace Entities
    <System.ComponentModel.DataAnnotations.Schema.Table("award")>
    Partial Public Class award
        Public Sub New()
            club_award = New Generic.HashSet(Of club_award)()
        End Sub

        <System.ComponentModel.DataAnnotations.StringLength(2147483647)>
        Public Property name As String

        Public Property id As Long

        Public Overridable Property club_award As Generic.ICollection(Of club_award)
    End Class
End Namespace
