Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

Namespace Entities

    <Table("award")>
    Partial Public Class award
        Public Sub New()
            club_award = New HashSet(Of club_award)()
        End Sub

        <StringLength(2147483647)>
        Public Property name As String

        Public Property id As Long

        Public Overridable Property club_award As ICollection(Of club_award)
    End Class
End Namespace
