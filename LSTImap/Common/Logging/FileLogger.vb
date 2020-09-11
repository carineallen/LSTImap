Imports System
Imports System.IO

Namespace Common.Logging
    ''' <summary>
	''' This logging object writes application error And debug output to a text file.
	''' </summary>

    Public Class FileLogger
        Implements ILog

        ''' <summary>
		''' Lock object to prevent thread interactions
		''' </summary>

        Private Shared ReadOnly LogLock As Object

        ''' <summary>
		''' Static constructor
		''' </summary>

        Shared Sub New()
            'Default log file is defined here
            LogFile = New FileInfo("LSTImap.log")
            Enabled = True
            Verbose = False
            LogLock = New Object()
        End Sub


        ''' <summary>
		''' Turns the logging on And off.
		''' </summary>
        Public Shared Property Enabled As Boolean

        ''' <summary>
		''' Enables Or disables the output of Debug level log messages
		''' </summary>
        Public Shared Property Verbose As Boolean

        ''' <summary>
		''' The file to which log messages will be written
		''' </summary>
		''' <remarks>This property defaults to LSTImap.log.</remarks>
        Public Shared Property LogFile As FileInfo

        ''' <summary>
		''' Write a message to the log file
		''' </summary>
		''' <param name="text">The error text to log</param>

        Private Shared Sub LogToFile(ByVal text As String)
            If text Is Nothing Then Throw New ArgumentNullException("text")

            'We want to open the file and append some text to it

            SyncLock LogLock

                Using sw As StreamWriter = LogFile.AppendText()
                    sw.WriteLine(DateTime.Now & " " & text)
                    sw.Flush()
                End Using
            End SyncLock
        End Sub

        ''' <summary>
		''' Logs an error message to the logs
		''' </summary>
		''' <param name="message">This Is the error message to log</param>

        Public Sub LogError(ByVal message As String)
            If Enabled Then LogToFile(message)
        End Sub

        ''' <summary>
		''' Logs a debug message to the logs
		''' </summary>
		''' <param name="message">This Is the debug message to log</param>

        Public Sub LogDebug(ByVal message As String)
            If Enabled AndAlso Verbose Then LogToFile("DEBUG: " & message)
        End Sub

        Private Sub ILog_LogError(message As String) Implements ILog.LogError
            Throw New NotImplementedException()
        End Sub

        Private Sub ILog_LogDebug(message As String) Implements ILog.LogDebug
            Throw New NotImplementedException()
        End Sub
    End Class
End Namespace

