Imports System.Drawing.Drawing2D

Public Class Thumbnail
    Dim theMainForm As MainForm
    Dim imageFileName As String
    Dim thumbFileName As String
    Dim img As Bitmap
    Dim thumb As Bitmap
    Dim gr As Graphics
    Dim width As Integer = 256
    Dim height As Integer = 256
    Dim jpgQuality As Long = 90

    Public Sub New(ByVal myMainForm As MainForm)
        MyBase.New()
        theMainForm = myMainForm
    End Sub

    Property imageFile() As String
        Get
            imageFile = imageFileName
        End Get
        Set(ByVal Value As String)
            imageFileName = Value
        End Set

    End Property

    Public Sub Render()
        Dim posn As Integer
        Dim path As String
        Dim fileName As String
        Dim thumbnailPath As String
        Dim jpgEncoder As ImageCodecInfo = GetEncoderInfo("image/jpeg")
        Dim encoderParams As EncoderParameters = New EncoderParameters(1)
        Dim scaleFactor As Double
        Dim thumbWidth As Integer
        Dim thumbHeight As Integer
        Dim thumbX As Integer
        Dim thumbY As Integer

        Try
            ' Convert the image name into an absolute path
            posn = InStrRev(imageFileName, "\")
            If posn = 0 Then
                path = "."
            Else
                path = Mid(imageFileName, 1, posn - 1)
            End If
            fileName = Mid(imageFileName, posn + 1)
            ' If it's a relative path, convert to an absolute path
            If Not InStr(1, path, ":\") = 2 Then
                path = theMainForm.images_root_folder + "\" + path
            End If

            ' Create a source and destination bitmap and a background canvas
            img = New Bitmap(path + "\" + fileName)
            thumb = New Bitmap(width, height)
            gr = Graphics.FromImage(thumb)
            gr.InterpolationMode = InterpolationMode.HighQualityBicubic

            ' Compute the size of the thumbnail to paint on the canvas
            If img.Width > img.Height Then  ' Landscape image
                scaleFactor = width / img.Width
                thumbWidth = width
                thumbHeight = img.Height * scaleFactor
                thumbX = 0
                thumbY = (height - thumbHeight) / 2
            Else
                scaleFactor = height / img.Height
                thumbHeight = height
                thumbWidth = img.Width * scaleFactor
                thumbY = 0
                thumbX = (width - thumbWidth) / 2
            End If

            ' Paint the thumbnail on the canvas
            gr.DrawImage(img, New Rectangle(thumbX, thumbY, thumbWidth, thumbHeight), New Rectangle(0, 0, img.Width, img.Height), GraphicsUnit.Pixel)
            gr.Dispose()

            ' Compute the name of the destination thumbnail file
            thumbnailPath = path + "\Thumbnails"
            ' If the Thumbnails subdirectory doesn't exist, create it
            If Not Directory.Exists(thumbnailPath) Then
                Directory.CreateDirectory(thumbnailPath)
            End If
            thumbFileName = thumbnailPath + "\" + fileName

            ' Write the thumbnail image to the destination file
            encoderParams.Param(0) = New EncoderParameter(Imaging.Encoder.Quality, jpgQuality)
            thumb.Save(thumbFileName, jpgEncoder, encoderParams)
            img.Dispose()
            thumb.Dispose()
            encoderParams.Dispose()

        Catch exception As Exception
            MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
        End Try
    End Sub


    Private Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
        Dim j As Integer
        Dim encoders As ImageCodecInfo()
        encoders = ImageCodecInfo.GetImageEncoders()
        For j = 0 To encoders.Length
            If encoders(j).MimeType = mimeType Then
                Return encoders(j)
            End If
        Next j
        Return Nothing
    End Function
End Class
