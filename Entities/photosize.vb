Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

Namespace Entities

    <Table("photosize")>
    Partial Public Class photosize
        Public Sub New()

        End Sub
        <Key>
        <StringLength(2147483647)>
        Public Property Competition_Date_1 As String

        Public Property width As Long?

        Public Property height As Long?

    End Class
End Namespace
