Namespace Entities
    Public Class ClassificationMedium
        Public Property Classification As String
            Get
                Return _classification1
            End Get
            Set
                _classification1 = Value
            End Set
        End Property

        Public Property Medium As String
            Get
                Return _medium1
            End Get
            Set
                _medium1 = Value
            End Set
        End Property

        Private _classification1 As String
        Private _medium1 As String
    End Class
End Namespace
