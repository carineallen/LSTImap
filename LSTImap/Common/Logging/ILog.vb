Namespace Common.Logging

    ''' <summary>
	''' Defines a logger for managing system logging output  
	''' </summary>

    Public Interface ILog

        ''' <summary>
        ''' Logs an error message to the logs
        ''' </summary>
        ''' <param name="message">This Is the error message to log</param>

        Sub LogError(ByVal message As String)

        ''' <summary>
		''' Logs a debug message to the logs
		''' </summary>
		''' <param name="message">This Is the debug message to log</param>

        Sub LogDebug(ByVal message As String)
    End Interface
End Namespace