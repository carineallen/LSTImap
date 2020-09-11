Imports System
Imports System.Runtime.Serialization

Namespace IMAP.Exceptions

    ''' <summary>
	''' Thrown when the <see cref="ImapClient"/> Is being used in an invalid way.<br/>
	''' This could for example happen if a someone tries to fetch a message without authenticating.
	''' </summary>
    <Serializable>
    Public Class InvalidUseException
        Inherits ImapClientException


        '''<summary>
		''' Creates a InvalidUseException with the given message
		'''</summary>
		'''<param name="message">The message to include in the exception</param>
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

        ''' <summary>
		''' Creates a New instance of the InvalidUseException class with serialized data.
		''' </summary>
		''' <param name="info">holds the serialized object data about the exception being thrown</param>
		''' <param name="context">contains contextual information about the source Or destination</param>
        Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class
End Namespace