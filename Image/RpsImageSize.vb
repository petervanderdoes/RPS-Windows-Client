Namespace Image
    Class RpsImageSize
        Private ReadOnly full_width As Int64 = 1440
        Private ReadOnly full_height As Int64 = 900

        Public ReadOnly Property getFullWidth As Integer
            Get
                Return full_width
            End Get
        End Property

        Public ReadOnly Property getFullHeight As Integer
            Get
                Return full_height
            End Get
        End Property
    End Class
End Namespace
