Imports System
Imports System.Runtime.Serialization

Namespace IMAP.Exceptions

    ''' <summary>
	''' This Is the base exception for all <see cref="IMapClient"/> exceptions.
	''' </summary>

    <Serializable>
    Public MustInherit Class ImapClientException
        Inherits Exception

        ''' <summary>
		'''  Creates a New instance of the ImapClientException class
		''' </summary>

        Protected Sub New()
        End Sub

        '''<summary>
		''' Creates a ImapClientException with the given message And InnerException
		'''</summary>
		'''<param name="message">The message to include in the exception</param>
		'''<param name="innerException">The exception that Is the cause of this exception</param>

        Protected Sub New(ByVal message As String, ByVal innerException As Exception)
            MyBase.New(message, innerException)
            If message Is Nothing Then Throw New ArgumentNullException("message")
            If innerException Is Nothing Then Throw New ArgumentNullException("innerException")
        End Sub

        '''<summary>
		''' Creates a ImapClientException with the given message
		'''</summary>
		'''<param name="message">The message to include in the exception</param>
        Protected Sub New(ByVal message As String)
            MyBase.New(message)
            If message Is Nothing Then Throw New ArgumentNullException("message")
        End Sub


        ''' <summary>
		''' Creates a New instance of the ImapClientException class with serialized data.
		''' </summary>
		''' <param name="info">holds the serialized object data about the exception being thrown</param>
		''' <param name="context">contains contextual information about the source Or destination</param>
        Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class
End Namespace
