'
'   A key-value pair for adding to listboxes and comboboxes
'
Namespace Entities
    Public Class DataItem

        Private _id As Integer
        Private _value As String

        Public Sub New(ByVal id As String, ByVal value As String)
            _id = id
            _value = value
        End Sub

        Public ReadOnly Property ID() As String
            Get
                Return _id
            End Get
        End Property

        Public ReadOnly Property Value() As String
            Get
                Return _value
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Value
        End Function
    End Class
End Namespace
