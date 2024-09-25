'------------------------------------------------
'Name: Module for clLogging.vb
'Function: Logging.
'Messages with a msec timer over the start and stop. Support nesting of msec messages to 1 deep.
'Filtering of log files on long delays; default 60 seconds.
'Copyright Robin Baines 2017. All rights reserved.
'Created Nov 2017.
'Notes: 
'Modifications:
'20200428 FilterFile. Added check that the file exists before trying to filter.
'20200430 ReNameLogFile modified to set the Logfile!!
'20240917 added try/catch with msgbox. Missing r/w permission may be the cause and it proves better to catch this than letting it go.
''------------------------------------------------
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Globalization
Public Class clLogging

    Private Shared Logfile As String = ReNameLogFile()
    Private Shared sw As New Stopwatch
    Private Shared blnStarted As Boolean = False
    Private Shared sw2 As New Stopwatch
    Private Shared blnStarted2 As Boolean = False
    Private Shared TotalTime As Long

    'length of a long delay in a log file.    'GISEventSettings.Instance.SecsMaxLogRecord
    Private Shared _SecsMaxLogRecord As Integer = 60
    Public Shared Property SecsMaxLogRecord() As Integer
        Get
            Return _SecsMaxLogRecord
        End Get
        Set(ByVal value As Integer)
            _SecsMaxLogRecord = value
        End Set
    End Property

    'If DebugLogging is true then any calls to LogStart and LogStop will support messages with a msec timer over the start and stop.
    Private Shared _DebugLogging As Boolean = False
    Public Shared Property DebugLogging() As Boolean
        Get
            Return _DebugLogging
        End Get
        Set(ByVal value As Boolean)
            _DebugLogging = value
        End Set
    End Property

    'If DebugLogging is true then any calls to LogStart and LogStop will support messages with a msec timer over the start and stop.
    Public Shared Sub LogStart()
        If DebugLogging Then
            If blnStarted = True Then
                If blnStarted2 = True Then
                    'Support nesting of msec messages to 1 deep.
                    LogStop("sw already started!")
                Else
                    blnStarted2 = True
                    sw2.Start()
                End If
            Else
                blnStarted = True
                sw.Start()
            End If
        End If

    End Sub

    Public Shared Sub LogStop(line As String)
        If DebugLogging Then
            If blnStarted2 = True Then
                sw2.Stop()
                TotalTime = 1000L * sw2.ElapsedTicks / Stopwatch.Frequency

                'show that this is nested with -----.
                Log(TotalTime.ToString() + " msec. ----- " + line)
                sw2.Reset()
                blnStarted2 = False
            Else
                sw.Stop()
                TotalTime = 1000L * sw.ElapsedTicks / Stopwatch.Frequency

                'moved in front of Log.
                sw.Reset()
                Log(TotalTime.ToString() + " msec. " + line)
                blnStarted = False
            End If
        End If
    End Sub

    'log messages to the log file and show in the console.
    Public Shared Sub Log(line As String)

        '20240924 added this check.
        If DebugLogging Then

            '20240917 added try/catch with msgbox. Missing r/w permission may be the cause and it proves better to catch this than letting it go.
            Try
                If Not Directory.Exists(Path.GetDirectoryName(Logfile)) Then
                    Directory.CreateDirectory(Path.GetDirectoryName(Logfile))
                End If
                Dim message As String = String.Format("{0} - {1}", DateTime.Now, line)
                Dim tw As TextWriter = New StreamWriter(Logfile, True)
                tw.WriteLine(message)
                tw.Close()

                'This call has no effect in GUI applications.
                Console.WriteLine(message)
            Catch ex As Exception
                MsgBox(line + " could not be written to the log. Check r/w permission. " + ex.Message)
            End Try
        End If
    End Sub

    'search a logfile looking for long delays.
    Public Shared Sub FilterFile(FileName As String)

        'Dim FileName As String = DateTime.Now.AddDays(-1).ToString("ddMMyyyy")
        Dim sFile As String = ReNameLogFile(FileName)
        Dim sLastline As String = ""
        Dim dLastDate As DateTime = DateTime.Now

        '20200428 FilterFile. Added check that the file exists before trying to filter.
        If File.Exists(sFile) Then
            Using sr As StreamReader = New StreamReader(sFile)
                While sr.Peek() >= 0
                    Dim Line As String = sr.ReadLine()
                    '30-5-2014 15:52:06
                    '30-5-2014 15:52:06
                    '12/9/2019 12:03:19 PM
                    '1/9/2019 01:03:19 AM
                    Try
                        Dim oDate As DateTime
                        Dim iL As Integer
                        iL = Line.IndexOf(" PM -")
                        If iL = -1 Then
                            iL = Line.IndexOf(" AM -")
                        End If
                        If iL > 0 Then
                            If DateTime.TryParse(Line.Substring(0, iL + 3), oDate) Then
                                If Not sLastline = "" Then

                                    Dim ts As TimeSpan = oDate.Subtract(dLastDate)
                                    If ts.TotalSeconds > SecsMaxLogRecord Then

                                        'open the file
                                        Dim tw As TextWriter = New StreamWriter(ReNameLogFile(FileName + "_filtered"), True)
                                        tw.WriteLine(sLastline)
                                        tw.WriteLine(Line)
                                        tw.WriteLine()
                                        tw.Close()

                                    End If
                                End If
                                sLastline = Line
                                dLastDate = oDate
                            End If
                        Else
                            '30-04-20 14:14:07
                            Dim culture As CultureInfo
                            culture = CultureInfo.CreateSpecificCulture("de-DE")
                            iL = Line.IndexOf(" -")
                            Dim strD As String = Line.Substring(0, iL)
                            strD = strD.Substring(0, 6) + "20" + strD.Substring(6)
                            If DateTime.TryParseExact(strD, "dd-MM-yyyy HH:mm:ss", culture, DateTimeStyles.None, oDate) Then
                                If Not sLastline = "" Then

                                    Dim ts As TimeSpan = oDate.Subtract(dLastDate)
                                    If ts.TotalSeconds > SecsMaxLogRecord Then

                                        'open the file
                                        Dim tw As TextWriter = New StreamWriter(ReNameLogFile(FileName + "_filtered"), True)
                                        tw.WriteLine(sLastline)
                                        tw.WriteLine(Line)
                                        tw.WriteLine()
                                        tw.Close()

                                    End If
                                End If
                                sLastline = Line
                                dLastDate = oDate
                            End If
                        End If
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                End While
            End Using
        End If
    End Sub

    Public Shared Function ReNameLogFile() As String

        '20130424 Create the log file name each time so that we get a new log file at the beginning of the day.
        '20200430 ReNameLogFile modified to set the Logfile!!
        Logfile = ReNameLogFile(DateTime.Now.ToString("ddMMyyyy"))
        Return Logfile
    End Function

    Public Shared Function ReNameLogFile(sFilename As String) As String

        '20240924 If running on Citrix use unique logfile names for each session.
        '20240924 The session name will defined and not equal to console if running on citrix. 

        'Dim TheSessionName As String
        'Returns The value of the environment variable specified by variable, or null if the environment variable is not found.
        'TheSessionName = Environment.GetEnvironmentVariable("SESSIONNAME")
        'If TheSessionName Is Nothing Then
        '    TheSessionName = ""
        'End If

        'If TheSessionName.ToLower = "console" Then
        '    TheSessionName = ""
        'End If

        'If TheSessionName.ToLower <> "" Then
        '    TheSessionName = TheSessionName + "_"
        'End If

        Dim sessionid As String = Process.GetCurrentProcess().SessionId.ToString()

        Return Path.Combine(Path.GetDirectoryName(New Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath), "Logfiles",
                                                "Logfile_" + sessionid + "_" +
                                                sFilename + ".txt")
    End Function

    Public Shared Sub RemoveOldLogFiles(days As Integer)
        Try
            Dim DirectoryInfo As Object = New DirectoryInfo(Path.GetDirectoryName(Logfile))
            For Each file As Object In DirectoryInfo.GetFiles("Logfile_*")
                If file.Name.StartsWith("Logfile_") Then

                    ' Remove when lastwritetime is older than 42 days
                    If file.LastWriteTime < DateTime.Now.AddDays(-days) Then

                        Try
                            file.Delete()

                        Catch ex As Exception
                            Log("Cannot delete logfile: " + file.FullName + ". Reason: " + ex.Message)
                        End Try
                    End If
                End If
            Next
        Catch ex As Exception
            Log(String.Format("An error occurred while removing the old logfiles : 0", ex.Message))
        End Try
    End Sub
End Class