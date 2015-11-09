Namespace Helpers
    Public Class Rest
        Sub New(server As String)
            Me.Server = server
        End Sub

        Public Property Server As String

        Public Function DoRestGet(server As String, operation As String, params As Hashtable, ByRef results As String) _
            As Boolean
            Dim request As HttpWebRequest
            Dim response As HttpWebResponse = Nothing
            Dim url As String
            Dim delim As String
            Dim reader As StreamReader
            Dim string_builder As StringBuilder
            DoRestGet = False

            Try
                ' Build the URL
                url = "http://" + server + operation
                delim = "?"
                For Each key As String In params.Keys
                    url = url + delim + key + "=" + params.Item(key)
                    delim = "&"
                Next

                ' Create the web request  
                request = DirectCast(WebRequest.Create(url), HttpWebRequest)
                response = DirectCast(request.GetResponse(), HttpWebResponse)
                If request.HaveResponse = True AndAlso Not (response Is Nothing) Then

                    ' Get the response stream  
                    reader = New StreamReader(response.GetResponseStream())

                    ' Read it into a StringBuilder  
                    string_builder = New StringBuilder(reader.ReadToEnd())

                    ' Console application output  
                    results = string_builder.ToString()
                    DoRestGet = True
                End If
            Catch web_exception As WebException
                ' This exception will be raised if the Server didn't return 200 - OK  
                ' Try to retrieve more information about the network error  
                If Not web_exception.Response Is Nothing Then
                    Dim error_response As HttpWebResponse = Nothing
                    Try
                        error_response = DirectCast(web_exception.Response, HttpWebResponse)
                        Dim message As String
                        message = String.Format("The Server returned '{0}' with the status code {1} ({2:d}).",
                                                error_response.StatusDescription,
                                                error_response.StatusCode,
                                                error_response.StatusCode)
                        MsgBox(message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)

                    Finally
                        If Not error_response Is Nothing Then error_response.Close()
                    End Try
                End If
            Finally
                If Not response Is Nothing Then response.Close()
            End Try
        End Function

        Public Async Function DoRestPost(
                                         post_data As _
                                            Generic.IReadOnlyCollection(Of Generic.KeyValuePair(Of String, String))) _
            As Tasks.Task(Of Boolean)

            Dim client As New Http.HttpClient()
            client.BaseAddress = New Uri("http://" + Me.Server)
            Dim content As Http.HttpContent = New Http.FormUrlEncodedContent(post_data)
            Dim response As Http.HttpResponseMessage = Await client.PostAsync(client.BaseAddress, content)
            response.EnsureSuccessStatusCode()
        End Function
    End Class
End Namespace
