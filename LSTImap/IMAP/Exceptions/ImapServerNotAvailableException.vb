Imports System
Imports System.Runtime.Serialization

Namespace IMAP.Exceptions

    ''' <summary>
	''' Thrown when the IMAP server sends an error "NO" during initial handshake "HELO".
	''' </summary>	
    <Serializable>
    Public Class ImapServerNotAvailableException
        Inherits ImapClientException

        '''<summary>
		''' Creates a ImapServerNotAvailableException with the given message And InnerException
		'''</summary>
		'''<param name="message">The message to include in the exception</param>
		'''<param name="innerException">The exception that Is the cause of this exception</param>
        Public Sub New(ByVal message As String, ByVal innerException As Exception)
            MyBase.New(message, innerException)
        End Sub

        Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class
End Namespace
