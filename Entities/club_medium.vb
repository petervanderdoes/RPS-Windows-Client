Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

Partial Public Class club_medium
    Public Property club_id As Long?

    Public Property medium_id As Long?

    <Key>
    <DatabaseGenerated(DatabaseGeneratedOption.None)>
    Public Property sort_key As Long

    Public Overridable Property club As club

    Public Overridable Property medium As medium
End Class
