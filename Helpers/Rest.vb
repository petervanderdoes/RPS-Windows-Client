Namespace Helpers
    Public Class Rest
        Sub New(server As String)
            Me.Server = server
        End Sub

        Public Property Server As String
        Public Property ErrorMessage As String
        Public Property result As String

        Public Function DoPost(post_data As _
                                      Generic.IReadOnlyCollection(Of Generic.KeyValuePair(Of String, String))) _
            As Boolean
            Dim client As New Http.HttpClient()
            DoPost = True
            Try
                client.BaseAddress = New Uri("http://" + Me.Server)
                Dim content As Http.HttpContent = New Http.FormUrlEncodedContent(post_data)
                Dim response As Http.HttpResponseMessage = client.PostAsync(client.BaseAddress, content).Result
                response.EnsureSuccessStatusCode()
            Catch exception As HttpException
                DoPost = False
                Me.ErrorMessage = exception.Message
            Finally
                client.Dispose()
            End Try
        End Function

        Public Function DoGet(operation As String,
                              arguments As Hashtable) As Object
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
                client.BaseAddress = New Uri("http://" + Me.Server + uri_string)
                Dim response As Http.HttpResponseMessage = client.GetAsync(client.BaseAddress).Result
                DoGet = response.Content.ReadAsStringAsync().Result
            Catch exception As HttpException
                DoGet = Nothing
                Me.ErrorMessage = exception.Message
            Finally
                client.Dispose()
            End Try
        End Function
    End Class
End Namespace
