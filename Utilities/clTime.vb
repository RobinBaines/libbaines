'------------------------------------------------
'Name: Module clTime.vb
'Function: A class to bind to the drop down of the time of day combobox.
'Copyright Baines 2012. All rights reserved.
'Modifications:
'------------------------------------------------
Imports System.Configuration
'Imports Utilities
Public Class aTime

    Public Property Time As String
    Public Sub New(_Time As String)
        Time = _Time
    End Sub
End Class

Public Class clTime
    Public Times As List(Of aTime) = New List(Of aTime)
    Public Sub New(blnIncludeZero As Boolean, iFirst As Integer, iLast As Integer, iStep As Integer, blnInclude30mins As Boolean)
        'add 00 to indicate no order.
        If blnIncludeZero = True And iFirst > 0 Then
            Times.Add(New aTime("00" + ":00:00"))
            If blnInclude30mins Then
                Times.Add(New aTime("00" + ":30:00"))
            End If

        End If
        For iT As Integer = iFirst To iLast Step iStep
            Times.Add(New aTime(iT.ToString("00") + ":00:00"))
            If blnInclude30mins Then
                Times.Add(New aTime(iT.ToString("00") + ":30:00"))
            End If
        Next
    End Sub
End Class
