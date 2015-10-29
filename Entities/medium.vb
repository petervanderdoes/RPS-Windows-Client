Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

Namespace Entities

    <Table("medium")>
    Partial Public Class medium
        Public Sub New()
            club_medium = New HashSet(Of club_medium)()
        End Sub

        <StringLength(2147483647)>
        Public Property name As String

        Public Property id As Long

        Public Overridable Property club_medium As ICollection(Of club_medium)
    End Class
End Namespace
