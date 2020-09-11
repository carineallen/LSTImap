Imports System

Namespace Common.Logging
    '''<summary>
	'''This Is the log that all logging will go trough.
	''' </summary>

    Public Class DefaultLogger

        ''' <summary>
		''' This Is the logger used by all logging methods in the assembly.<br/>
		''' You can override this if you want, to move logging to one of your own
		''' logging implementations.<br/>
		''' <br/>
		''' By default a <see cref="DiagnosticsLogger"/> Is used.
		''' </summary>

        Public Shared Property Log As ILog

        Shared Sub New()
            Log = New DiagnosticsLogger()
        End Sub

		''' <summary>
		''' Changes the default logging to log to a New logger
		''' </summary>
		''' <param name="newLogger">The New logger to use to send log messages to</param>
		''' <exception cref="ArgumentNullException">
		''' Never set this to <see langword="null"/>.<br/>
		''' Instead you should implement a NullLogger which just does nothing.
		''' </exception>

		Sub SetLog(ByVal newLogger As ILog)
            If newLogger Is Nothing Then Throw New ArgumentNullException("newLogger")
            Log = newLogger
        End Sub
    End Class
End Namespace
