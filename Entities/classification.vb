

Namespace Entities
    <System.ComponentModel.DataAnnotations.Schema.Table("classification")>
    Partial Public Class classification
        Public Sub New()
            club_classification = New Generic.HashSet(Of club_classification)()
        End Sub

        <System.ComponentModel.DataAnnotations.StringLength(2147483647)>
        Public Property name As String

        Public Property id As Long

        Public Overridable Property club_classification As Generic.ICollection(Of club_classification)
    End Class
End Namespace
