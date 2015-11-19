Namespace Image
    Public Class Thumbnail
        Implements IDisposable
        Private Const JPG_QUALITY As Long = 90
        Private Const THUMBNAIL_HEIGHT As Integer = 256
        Private Const THUMBNAIL_WIDTH As Integer = 256
        Private ReadOnly the_main_form As Forms.MainForm
        Private gr As Graphics
        Private image_file_name As String
        Private img As Bitmap
        Private thumb As Bitmap
        Private thumb_file_name As String
        Private disposed As Boolean = False

        Public Sub New(my_main_form As Forms.MainForm)
            MyBase.New()
            the_main_form = my_main_form
        End Sub

        Property imageFile As String
            Private Get
                imageFile = image_file_name
            End Get
            Set
                image_file_name = Value
            End Set
        End Property

        Public Sub doRender()
            Dim posn As Integer
            Dim path As String
            Dim file_name As String
            Dim thumbnail_path As String
            Dim jpg_encoder As ImageCodecInfo = getEncoderInfo("image/jpeg")
            Dim encoder_params As EncoderParameters = New EncoderParameters(1)
            Dim scale_factor As Double
            Dim thumb_width As Integer
            Dim thumb_height As Integer
            Dim thumb_x As Integer
            Dim thumb_y As Integer

            Try
                ' Convert the image name into an absolute path
                posn = InStrRev(image_file_name, "\")
                If posn = 0 Then
                    path = "."
                Else
                    path = Mid(image_file_name, 1, posn - 1)
                End If
                file_name = Mid(image_file_name, posn + 1)
                ' If it's a relative path, convert to an absolute path
                If Not InStr(1, path, ":\") = 2 Then
                    path = the_main_form.images_root_folder + "\" + path
                End If

                ' Create a source and destination bitmap and a background canvas
                img = New Bitmap(path + "\" + file_name)
                thumb = New Bitmap(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT)
                gr = Graphics.FromImage(thumb)
                gr.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic

                ' Compute the size of the thumbnail to paint on the canvas
                If img.Width > img.Height Then ' Landscape image
                    scale_factor = THUMBNAIL_WIDTH / img.Width
                    thumb_width = THUMBNAIL_WIDTH
                    thumb_height = img.Height * scale_factor
                    thumb_x = 0
                    thumb_y = (THUMBNAIL_HEIGHT - thumb_height) / 2
                Else
                    scale_factor = THUMBNAIL_HEIGHT / img.Height
                    thumb_height = THUMBNAIL_HEIGHT
                    thumb_width = img.Width * scale_factor
                    thumb_y = 0
                    thumb_x = (THUMBNAIL_WIDTH - thumb_width) / 2
                End If

                ' Paint the thumbnail on the canvas
                gr.DrawImage(img,
                             New Rectangle(thumb_x, thumb_y, thumb_width, thumb_height),
                             New Rectangle(0, 0, img.Width, img.Height),
                             GraphicsUnit.Pixel)
                gr.Dispose()

                ' Compute the name of the destination thumbnail file
                thumbnail_path = path + "\Thumbnails"
                ' If the Thumbnails subdirectory doesn't exist, create it
                If Not Directory.Exists(thumbnail_path) Then
                    Directory.CreateDirectory(thumbnail_path)
                End If
                thumb_file_name = thumbnail_path + "\" + file_name

                ' Write the thumbnail image to the destination file
                encoder_params.Param(0) = New EncoderParameter(Imaging.Encoder.Quality, JPG_QUALITY)
                thumb.Save(thumb_file_name, jpg_encoder, encoder_params)
                img.Dispose()
                thumb.Dispose()
                encoder_params.Dispose()

            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            End Try
        End Sub

        Private Function getEncoderInfo(mime_type As String) As ImageCodecInfo
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

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If disposed Then Exit Sub

            ' Dispose of managed resources here.
            If disposing Then
                img.Dispose()
                thumb.Dispose()
            End If

            ' Dispose of any unmanaged resources not wrapped in safe handles.

            disposed = True
        End Sub
    End Class
End Namespace
