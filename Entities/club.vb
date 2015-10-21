Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

<Table("club")>
Partial Public Class club
    Public Sub New()
        club_award = New HashSet(Of club_award)()
        club_classification = New HashSet(Of club_classification)()
        club_medium = New HashSet(Of club_medium)()
    End Sub

    <StringLength(2147483647)>
    Public Property short_name As String

    <StringLength(2147483647)>
    Public Property name As String

    Public Property id As Long

    Public Property min_score As Long?

    Public Property max_score As Long?

    Public Property min_score_for_award As Long?

    Public Overridable Property club_award As ICollection(Of club_award)

    Public Overridable Property club_classification As ICollection(Of club_classification)

    Public Overridable Property club_medium As ICollection(Of club_medium)
End Class
