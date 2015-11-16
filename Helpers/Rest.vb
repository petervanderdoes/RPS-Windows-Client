Namespace Helpers
    Public Class Rest
        Sub New(server As String)
            Me.Server = server
        End Sub

        Public Property Server As String
        Public Property ErrorMessage As String

        Private Property HttpResponseMessage As Http.HttpResponseMessage

        Public Function DoPost(post_data As Generic.IReadOnlyCollection(Of Generic.KeyValuePair(Of String, String))) As Object
            Dim client As New Http.HttpClient()
            Try
                client.BaseAddress = New Uri("http://" + Server)
                Dim content As Http.HttpContent = New Http.FormUrlEncodedContent(post_data)
                HttpResponseMessage = client.PostAsync(client.BaseAddress, content).Result
                HttpResponseMessage.EnsureSuccessStatusCode()
                DoPost = HttpResponseMessage.Content.ReadAsStringAsync().Result
            Catch exception As HttpException
                DoPost = False
                ErrorMessage = exception.Message
            Finally
                client.Dispose()
            End Try
        End Function

        Public Function DoGet(operation As String, arguments As Hashtable) As Object
            Dim client As New Http.HttpClient()
            Dim delim As String
            Dim uri_string As String

            Try
                uri_string = operation
                delim = "?"
                For Each argument As DictionaryEntry In arguments
                    uri_string = String.Format("{0}{1}{2}={3}", uri_string, delim, argument.Key, argument.Value)
                    delim = "&"
                Next
                client.BaseAddress = New Uri("http://" + Server + uri_string)
                HttpResponseMessage = client.GetAsync(client.BaseAddress).Result
                If HttpResponseMessage.IsSuccessStatusCode Then
                    DoGet = HttpResponseMessage.Content.ReadAsStringAsync().Result
                Else
                    ErrorMessage = HttpResponseMessage.ReasonPhrase
                    DoGet = Nothing
                End If
            Catch exception As Exception
                DoGet = Nothing
                ErrorMessage = exception.Message
            Finally
                client.Dispose()
            End Try
        End Function
    End Class
End Namespace
