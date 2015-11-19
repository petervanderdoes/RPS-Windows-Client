

Namespace Entities
    <System.ComponentModel.DataAnnotations.Schema.Table("photosize")>
    Partial Public Class photosize
        Public Sub New()
        End Sub

        <System.ComponentModel.DataAnnotations.Key>
        <System.ComponentModel.DataAnnotations.StringLength(2147483647)>
        Public Property Competition_Date_1 As String

        Public Property width As Long?

        Public Property height As Long?
    End Class
End Namespace
