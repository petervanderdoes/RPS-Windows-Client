Class RpsImage

    Private full_width As Integer = 1440
    Private full_height As Integer = 900

    Function setFullSize() As System.Drawing.Size
        Return New Size(full_width, full_height)
    End Function

    Function setSplashClub() As System.Drawing.Size
        Return New Size(full_width, 102)
    End Function

    Function setSplashTheme() As System.Drawing.Size
        Return New Size(full_width, 102)
    End Function

    Function setSplashClassification() As System.Drawing.Size
        Return New Size(full_width, 102)
    End Function
End Class
