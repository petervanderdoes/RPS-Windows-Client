Imports System.Drawing.Drawing2D

Public Class Thumbnail
    Private Const JPG_QUALITY As Long = 90
    Private Const THUMBNAIL_HEIGHT As Integer = 256
    Private Const THUMBNAIL_WIDTH As Integer = 256
    Private ReadOnly _the_main_form As MainForm
    Private _gr As Graphics
    Private _image_file_name As String
    Private _img As Bitmap
    Private _thumb As Bitmap
    Private _thumb_file_name As String

    Public Sub New(my_main_form As MainForm)
        MyBase.New()
        _the_main_form = my_main_form
    End Sub

    Property imageFile As String
        Get
            imageFile = _image_file_name
        End Get
        Set
            _image_file_name = Value
        End Set

    End Property

    Public Sub doRender()
        Dim posn As Integer
        Dim path As String
        Dim file_name As String
        Dim thumbnail_path As String
        Dim jpg_encoder As ImageCodecInfo = GetEncoderInfo("image/jpeg")
        Dim encoder_params As EncoderParameters = New EncoderParameters(1)
        Dim scale_factor As Double
        Dim thumb_width As Integer
        Dim thumb_height As Integer
        Dim thumb_x As Integer
        Dim thumb_y As Integer

        Try
            ' Convert the image name into an absolute path
            posn = InStrRev(_image_file_name, "\")
            If posn = 0 Then
                path = "."
            Else
                path = Mid(_image_file_name, 1, posn - 1)
            End If
            file_name = Mid(_image_file_name, posn + 1)
            ' If it's a relative path, convert to an absolute path
            If Not InStr(1, path, ":\") = 2 Then
                path = _the_main_form.images_root_folder + "\" + path
            End If

            ' Create a source and destination bitmap and a background canvas
            _img = New Bitmap(path + "\" + file_name)
            _thumb = New Bitmap(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT)
            _gr = Graphics.FromImage(_thumb)
            _gr.InterpolationMode = InterpolationMode.HighQualityBicubic

            ' Compute the size of the thumbnail to paint on the canvas
            If _img.Width > _img.Height Then  ' Landscape image
                scale_factor = THUMBNAIL_WIDTH / _img.Width
                thumb_width = THUMBNAIL_WIDTH
                thumb_height = _img.Height * scale_factor
                thumb_x = 0
                thumb_y = (THUMBNAIL_HEIGHT - thumb_height) / 2
            Else
                scale_factor = THUMBNAIL_HEIGHT / _img.Height
                thumb_height = THUMBNAIL_HEIGHT
                thumb_width = _img.Width * scale_factor
                thumb_y = 0
                thumb_x = (THUMBNAIL_WIDTH - thumb_width) / 2
            End If

            ' Paint the thumbnail on the canvas
            _gr.DrawImage(_img, New Rectangle(thumb_x, thumb_y, thumb_width, thumb_height), New Rectangle(0, 0, _img.Width, _img.Height), GraphicsUnit.Pixel)
            _gr.Dispose()

            ' Compute the name of the destination thumbnail file
            thumbnail_path = path + "\Thumbnails"
            ' If the Thumbnails subdirectory doesn't exist, create it
            If Not Directory.Exists(thumbnail_path) Then
                Directory.CreateDirectory(thumbnail_path)
            End If
            _thumb_file_name = thumbnail_path + "\" + file_name

            ' Write the thumbnail image to the destination file
            encoder_params.Param(0) = New EncoderParameter(Imaging.Encoder.Quality, JPG_QUALITY)
            _thumb.Save(_thumb_file_name, jpg_encoder, encoder_params)
            _img.Dispose()
            _thumb.Dispose()
            encoder_params.Dispose()

        Catch exception As Exception
            MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
        End Try
    End Sub


    Private Function GetEncoderInfo(mime_type As String) As ImageCodecInfo
        Dim j As Integer
        Dim encoders As ImageCodecInfo()
        encoders = ImageCodecInfo.GetImageEncoders()
        For j = 0 To encoders.Length
            If encoders(j).MimeType = mime_type Then
                Return encoders(j)
            End If
        Next j
        Return Nothing
    End Function
End Class
