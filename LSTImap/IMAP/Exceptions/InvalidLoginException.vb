Imports System
Imports System.Runtime.Serialization

Namespace IMAP.Exceptions

    ''' <summary>
    ''' Thrown when the supplied username Or password Is Not accepted by the IMAP server.
    ''' </summary>

    <Serializable>
    Public Class InvalidLoginException
        Inherits ImapClientException

        '''<summary>
		''' Creates a InvalidLoginException with the given message And InnerException
		'''</summary>
		'''<param name="innerException"> The exception that Is the cause of this exception </param>

        Public Sub New(ByVal innerException As Exception)
            MyBase.New("Server did not accept user credentials", innerException)
        End Sub

        ''' <summary>
		''' Creates a New instance of the InvalidLoginException class with serialized data.
		''' </summary>
		''' <param name="info">holds the serialized object data about the exception being thrown</param>
		''' <param name="context">contains contextual information about the source Or destination</param>

        Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class
End Namespace