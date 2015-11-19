Namespace Entities
    Partial Public Class club_classification
        Public Property club_id As Long?

        Public Property classification_id As Long?

        <System.ComponentModel.DataAnnotations.Key>
        <
            System.ComponentModel.DataAnnotations.Schema.DatabaseGenerated _
                (System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.None)>
        Public Property sort_key As Long

        Public Overridable Property classification As classification

        Public Overridable Property club As club
    End Class
End Namespace
