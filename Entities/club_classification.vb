Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

Namespace Entities

    Partial Public Class club_classification
        Public Property club_id As Long?

        Public Property classification_id As Long?

        <Key>
        <DatabaseGenerated(DatabaseGeneratedOption.None)>
        Public Property sort_key As Long

        Public Overridable Property classification As classification

        Public Overridable Property club As club
    End Class
End Namespace
