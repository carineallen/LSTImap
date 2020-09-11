Imports System

Namespace Common.Logging
    ''' <summary>
	''' This logging object writes application error And debug output using the
	''' <see cref="System.Diagnostics.Trace"/> facilities.
	''' </summary>

    Public Class DiagnosticsLogger
        Implements ILog

        ''' <summary>
		''' Logs an error message to the System Trace facility
		''' </summary>
		''' <param name="message">This Is the error message to log</param>

        Public Sub LogError(ByVal message As String) Implements ILog.LogError
            If message Is Nothing Then Throw New ArgumentNullException("message")
            System.Diagnostics.Trace.WriteLine("LSTImap: " & message)
        End Sub

        ''' <summary>
		''' Logs a debug message to the system Trace Facility
		''' </summary>
		''' <param name="message">This Is the debug message to log</param>

        Public Sub LogDebug(ByVal message As String) Implements ILog.LogDebug
            If message Is Nothing Then Throw New ArgumentNullException("message")
            System.Diagnostics.Trace.WriteLine("LSTImap: (DEBUG) " & message)
        End Sub

    End Class
End Namespace
