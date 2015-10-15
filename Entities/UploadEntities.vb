Public Class UploadEntity_Classification_Medium
    Public Property Classification() As String
        Get
            Return m_Classification
        End Get
        Set
            m_Classification = Value
        End Set
    End Property
    Private m_Classification As String
    Public Property Medium() As String
        Get
            Return m_Medium
        End Get
        Set
            m_Medium = Value
        End Set
    End Property
    Private m_Medium As String
End Class
