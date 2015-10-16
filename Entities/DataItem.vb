'
'   A key-value pair for adding to listboxes and comboboxes
'
Namespace Entities
    Public Class DataItem

        Private m_id As Integer
        Private m_value As String

        Public Sub New(ByVal id As String, ByVal value As String)
            m_id = id
            m_value = value
        End Sub

        Public ReadOnly Property ID() As String
            Get
                Return m_id
            End Get
        End Property

        Public ReadOnly Property Value() As String
            Get
                Return m_value
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Value
        End Function
    End Class
End Namespace
