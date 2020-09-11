Imports System
Imports System.Runtime.Serialization

Namespace IMAP.Exceptions

    ''' <summary>
	''' Thrown when the server does Not return "OK" to a command.<br/>
	''' The server response Is then placed inside.
	''' </summary>
    <Serializable>
    Public Class ImapServerException
        Inherits ImapClientException

        '''<summary>
		''' Creates a ImapServerException with the given message
		'''</summary>
		'''<param name="message">The message to include in the exception</param>
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

        ''' <summary>
		''' Creates a New instance of the ImapServerException class with serialized data.
		''' </summary>
		''' <param name="info">holds the serialized object data about the exception being thrown</param>
		''' <param name="context">contains contextual information about the source Or destination</param>
        Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class
End Namespace
