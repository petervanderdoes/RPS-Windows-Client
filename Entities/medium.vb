

Namespace Entities
    <System.ComponentModel.DataAnnotations.Schema.Table("medium")>
    Partial Public Class medium
        Public Sub New()
            club_medium = New Generic.HashSet(Of club_medium)()
        End Sub

        <System.ComponentModel.DataAnnotations.StringLength(2147483647)>
        Public Property name As String

        Public Property id As Long

        Public Overridable Property club_medium As Generic.ICollection(Of club_medium)
    End Class
End Namespace
