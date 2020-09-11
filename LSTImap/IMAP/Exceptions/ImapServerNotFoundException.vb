Imports System
Imports System.Runtime.Serialization

Namespace IMAP.Exceptions

    ''' <summary>
	''' Thrown when the specified IMAP server can Not be found Or connected to.
	''' </summary>
    <Serializable>
    Public Class ImapServerNotFoundException
        Inherits ImapClientException

        '''<summary>
		''' Creates a ImapServerNotFoundException with the given message And InnerException
		'''</summary>
		'''<param name="message">The message to include in the exception</param>
		'''<param name="innerException">The exception that Is the cause of this exception</param>
        Public Sub New(ByVal message As String, ByVal innerException As Exception)
            MyBase.New(message, innerException)
        End Sub

        ''' <summary>
		''' Creates a New instance of the ImapServerNotFoundException class with serialized data.
		''' </summary>
		''' <param name="info">holds the serialized object data about the exception being thrown</param>
		''' <param name="context">contains contextual information about the source Or destination</param>
        Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class
End Namespace