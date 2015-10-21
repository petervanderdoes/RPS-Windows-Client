Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

Partial Public Class club_award
    Public Property club_id As Long?

    Public Property award_id As Long?

    Public Property points As Long?

    <Key>
    <DatabaseGenerated(DatabaseGeneratedOption.None)>
    Public Property sort_key As Long

    Public Overridable Property award As award

    Public Overridable Property club As club
End Class
