

Namespace Entities
    Partial Public Class club_award
        Public Property club_id As Long?

        Public Property award_id As Long?

        Public Property points As Long?

        <System.ComponentModel.DataAnnotations.Key>
        <
            System.ComponentModel.DataAnnotations.Schema.DatabaseGenerated _
                (System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.None)>
        Public Property sort_key As Long

        Public Overridable Property award As award

        Public Overridable Property club As club
    End Class
End Namespace
