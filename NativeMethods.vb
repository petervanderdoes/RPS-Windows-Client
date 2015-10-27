Imports System.Runtime.InteropServices

Friend NotInheritable Class NativeMethods

    Private Sub New()
    End Sub

    Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
End Class
