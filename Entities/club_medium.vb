

Namespace Entities
    Partial Public Class club_medium
        Public Property club_id As Long?

        Public Property medium_id As Long?

        <System.ComponentModel.DataAnnotations.Key>
        <
            System.ComponentModel.DataAnnotations.Schema.DatabaseGenerated _
                (System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.None)>
        Public Property sort_key As Long

        Public Overridable Property club As club

        Public Overridable Property medium As medium
    End Class
End Namespace
