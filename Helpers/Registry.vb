Namespace Helpers
    ' ReSharper disable once ClassNeverInstantiated.Global
    Public Class Registry
        Public Shared Property programRegistryKey As String
        Private Shared Property completeRegistryKey As RegistryKey

        Public Shared Sub writeRegistryString(registry_name As String, registry_value As String)
            completeRegistryKey = Nothing

            Try
                completeRegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(programRegistryKey, True)
                If completeRegistryKey Is Nothing Then
                    completeRegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(programRegistryKey)
                End If
                completeRegistryKey.SetValue(registry_name, registry_value)
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            Finally
                If Not completeRegistryKey Is Nothing Then
                    completeRegistryKey.Close()
                End If
            End Try
        End Sub

        Public Shared Function getRegistryString(registry_name As String) As String
            completeRegistryKey = Nothing

            Try
                completeRegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(programRegistryKey, False)
                If completeRegistryKey Is Nothing Then
                    getRegistryString = ""
                Else
                    getRegistryString = completeRegistryKey.GetValue(registry_name, "")
                End If
            Catch exception As Exception
                getRegistryString = ""
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            Finally
                If Not completeRegistryKey Is Nothing Then
                    completeRegistryKey.Close()
                End If
            End Try
        End Function
    End Class
End Namespace
