'------------------------------------------------
'Name: Module WindowHandleCounter.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 2/16/2012 12:00:00 AM.
'Notes: Read the window handle use. Show in a textbox to check it is not increasing, see GIS frmCcles for an example.
'Modifications:
Imports System
Imports System.Runtime.InteropServices
Public Class WindowHandleCounter
    <DllImport("kernel32.dll")> _
    Public Shared Function GetCurrentProcess() As IntPtr
    End Function


    <DllImport("user32.dll")> _
    Public Shared Function GetGuiResources(hProcess As IntPtr, uiFlags As Integer) As Integer

    End Function


    Enum ResourceType
        Gdi = 0
        User = 1
    End Enum

    Public Shared Function GetWindowHandlesForCurrentProcess() As Integer
        Dim processHandle As IntPtr = GetCurrentProcess()
        Dim gdiObjects As Integer = GetGuiResources(processHandle, ResourceType.Gdi)
        Dim userObjects As Integer = GetGuiResources(processHandle, ResourceType.User)

        Return Convert.ToInt32(gdiObjects + userObjects)
    End Function
End Class
