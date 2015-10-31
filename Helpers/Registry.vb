Namespace Helpers
    Public Class Registry
        Public Shared Property key As String

        Public Shared Sub writeRegistryString(registry_name As String, registry_value As String)
            Dim reg_key As RegistryKey = Nothing

            Try
                reg_key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key, True)
                If reg_key Is Nothing Then
                    reg_key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(key)
                End If
                reg_key.SetValue(registry_name, registry_value)
            Catch exception As Exception
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            Finally
                If Not reg_key Is Nothing Then
                    reg_key.Close()
                End If
            End Try
        End Sub

        Public Shared Function getRegistryString(registry_name As String) As String
            Dim reg_key As RegistryKey = Nothing

            Try
                reg_key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key, False)
                If reg_key Is Nothing Then
                    getRegistryString = ""
                Else
                    getRegistryString = reg_key.GetValue(registry_name, "")
                End If
            Catch exception As Exception
                getRegistryString = ""
                MsgBox(exception.Message, , "Error in: " + Reflection.MethodBase.GetCurrentMethod().ToString)
            Finally
                If Not reg_key Is Nothing Then
                    reg_key.Close()
                End If
            End Try
        End Function
    End Class
End Namespace
