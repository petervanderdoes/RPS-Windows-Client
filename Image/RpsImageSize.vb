Namespace Image
    Class RpsImageSize
        Private get_full_width As Integer = 0
        Private get_full_height As Integer = 0

        Public Property ImageWidth As Integer
            Get
                Return get_full_width
            End Get
            Set
                get_full_width = Value
            End Set
        End Property

        Public Property ImageHeight As Integer
            Get
                Return get_full_height
            End Get
            Set
                get_full_height = Value
            End Set
        End Property
    End Class
End Namespace
