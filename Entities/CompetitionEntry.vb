Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Entity.Spatial

Partial Public Class CompetitionEntry
    <Key>
    Public Property Photo_ID As Long

    <StringLength(2147483647)>
    Public Property Title As String

    <StringLength(2147483647)>
    Public Property Maker As String

    <StringLength(2147483647)>
    Public Property Classification As String

    <StringLength(2147483647)>
    Public Property Medium As String

    <StringLength(2147483647)>
    Public Property Theme As String

    <StringLength(2147483647)>
    Public Property Competition_Date_1 As String

    Public Property Score_1 As Long?

    <StringLength(2147483647)>
    Public Property Award As String

    <StringLength(2147483647)>
    Public Property Image_File_Name As String

    Public Property Display_Sequence As Long?

    Public Property Server_Entry_ID As Long?
End Class
