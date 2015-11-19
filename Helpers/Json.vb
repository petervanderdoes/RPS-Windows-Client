

Namespace Helpers
    Public Class Json
        Public Shared Function GetErrorMessage(json As Newtonsoft.Json.Linq.JObject) As String

            Dim error_messages As String = "Server Error" + Environment.NewLine

            error_messages = error_messages +
                             Linq.Enumerable.Aggregate(Linq.Enumerable.Cast(Of Object)(json("errors")),
                                                       "",
                                                       Function(current, error_entry) _
                                                          current + error_entry("detail") + Environment.NewLine)
            Return error_messages
        End Function

        Public Shared Function GetFailureMessage(json As Newtonsoft.Json.Linq.JObject) As String

            Dim error_messages As String = "Server Warning" + Environment.NewLine

            error_messages = error_messages +
                             Linq.Enumerable.Aggregate(Linq.Enumerable.Cast(Of Object)(json("errors")),
                                                       "",
                                                       Function(current, error_entry) _
                                                          current + error_entry("detail") + Environment.NewLine)
            Return error_messages
        End Function

        Public Shared Function HasErrors(json As Newtonsoft.Json.Linq.JObject) As Boolean
            Return Not IsNothing(json("errors"))
        End Function

        Public Shared Function IsError(json As Newtonsoft.Json.Linq.JObject) As Boolean
            If json("status") = "error" Then
                Return True
            End If
            Return False
        End Function

        Public Shared Function IsFailed(json As Newtonsoft.Json.Linq.JObject) As Boolean
            If json("status") = "failed" Then
                Return True
            End If
            Return False
        End Function

        Public Shared Function IsSuccess(json As Newtonsoft.Json.Linq.JObject) As Boolean
            If json("status") = "success" Then
                Return True
            End If
            Return False
        End Function
    End Class
End Namespace
