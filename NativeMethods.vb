

' ReSharper disable once ClassNeverInstantiated.Global
Friend NotInheritable Class NativeMethods
    Private Sub New()
    End Sub

    Declare Auto Function SendMessage Lib "user32.dll" (hWnd As IntPtr,
                                                       msg As Integer,
                                                       wParam As IntPtr,
                                                       lParam As IntPtr) As IntPtr
End Class
