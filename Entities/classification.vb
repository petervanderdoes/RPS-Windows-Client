Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

Namespace Entities

    <Table("classification")>
    Partial Public Class classification
        Public Sub New()
            club_classification = New HashSet(Of club_classification)()
        End Sub

        <StringLength(2147483647)>
        Public Property name As String

        Public Property id As Long

        Public Overridable Property club_classification As ICollection(Of club_classification)
    End Class
End Namespace
