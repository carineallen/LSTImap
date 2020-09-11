Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Mail
Imports System.Text
Imports LSTImap.Mime.Header
Imports LSTImap.Mime.Traverse

Namespace Mime

    ''' <summary>
	''' This Is the root of the email tree structure.<br/>
	''' <see cref="Mime.MessagePart"/> for a description about the structure.<br/>
	''' <br/>
	''' A Message (this class) contains the headers of an email message such as:
	''' <code>
	'''  - To
	'''  - From
	'''  - Subject
	'''  - Content-Type
	'''  - Message-ID
	''' </code>
	''' which are located in the <see cref="Headers"/> property.<br/>
	''' <br/>
	''' Use the <see cref="Message.MessagePart"/> property to find the actual content of the email message.
	''' </summary>
	''' <example>
	''' Examples are available on the <a href="http://hpop.sourceforge.net/">project homepage</a>.
	''' </example>
    Public Class Message

        ''' <summary>
		''' Headers of the Message.
		''' </summary>
        Public Property Headers As MessageHeader

        ''' <summary>
		''' This Is the body of the email Message.<br/>
		''' <br/>
		''' If the body was parsed for this Message, this property will never be <see langword="null"/>.
		''' </summary>
        Public Property MessagePart As MessagePart

        ''' <summary>
		''' The raw content from which this message has been constructed.<br/>
		''' These bytes can be persisted And later used to recreate the Message.
		''' </summary>
        Public Property RawMessage As Byte()


        ''' <summary>
		''' Convenience constructor for <see cref="Mime.Message(Byte[], bool)"/>.<br/>
		''' <br/>
		''' Creates a message from a byte array. The full message including its body Is parsed.
		''' </summary>
		''' <param name="rawMessageContent">The byte array which Is the message contents to parse</param>
        Public Sub New(ByVal rawMessageContent As Byte())
            Me.New(rawMessageContent, True)
        End Sub


        ''' <summary>
		''' Constructs a message from a byte array.<br/>
		''' <br/>
		''' The headers are always parsed, but if <paramref name="parseBody"/> Is <see langword="false"/>, the body Is Not parsed.
		''' </summary>
		''' <param name="rawMessageContent">The byte array which Is the message contents to parse</param>
		''' <param name="parseBody">
		''' <see langword="true"/> if the body should be parsed,
		''' <see langword="false"/> if only headers should be parsed out of the <paramref name="rawMessageContent"/> byte array
		''' </param>
        Public Sub New(ByVal rawMessageContent As Byte(), ByVal parseBody As Boolean)
            RawMessage = rawMessageContent
            'Find the headers and the body parts of the byte array
            Dim headersTemp As MessageHeader
            Dim body As Byte()
            HeaderExtractor.ExtractHeadersAndBody(rawMessageContent, headersTemp, body)
            'Set the Headers property
            Headers = headersTemp

            'Should we also parse the body?
            If parseBody Then
                'Parse the body into a MessagePart
                MessagePart = New MessagePart(body, Headers)
            End If
        End Sub


        ''' <summary>
		''' This method will convert this <see cref="Message"/> into a <see cref="MailMessage"/> equivalent.<br/>
		''' The returned <see cref="MailMessage"/> can be used with <see cref="System.Net.Mail.SmtpClient"/> to forward the email.<br/>
		''' <br/>
		''' You should be aware of the following about this method:
		''' <list type="bullet">
		''' <item>
		'''    All sender And receiver mail addresses are set.
		'''  If you send this email using a <see cref="System.Net.Mail.SmtpClient"/> then all
		'''   receivers in To, From, Cc And Bcc will receive the email once again.
		''' </item>
		''' <item>
		'''    If you view the source code of this Message And looks at the source code of the forwarded
		'''    <see cref="MailMessage"/> returned by this method, you will notice that the source codes are Not the same.
		'''    The content that Is presented by a mail client reading the forwarded <see cref="MailMessage"/> should be the
		'''    same as the original, though.
		''' </item>
		''' <item>
		'''    Content-Disposition headers will Not be copied to the <see cref="MailMessage"/>.
		'''    It Is simply Not possible to set these on Attachments.
		''' </item>
		''' <item>
		'''    HTML content will be treated as the preferred view for the <see cref="MailMessage.Body"/>. Plain text content will be used for the
		'''  <see cref="MailMessage.Body"/> when HTML Is Not available.
		''' </item>
		''' </list>
		'''</summary>
		''' <returns>A <see cref="MailMessage"/> object that contains the same information that this Message does</returns>
        Public Function ToMailMessage() As MailMessage
            'Construct an empty MailMessage to which we will gradually build up to look like the current Message object (this)
            Dim message As MailMessage = New MailMessage()
            message.Subject = Headers.Subject
            'We here set the encoding to be UTF-8
            ' We cannot determine what the encoding of the subject was at this point.
            ' But since we know that strings in .NET Is stored in UTF, we can
            ' use UTF-8 to decode the subject into bytes
            message.SubjectEncoding = Encoding.UTF8
            'The HTML version should take precedent over the plain text if it is available
            Dim preferredVersion As MessagePart = FindFirstHtmlVersion()

            If preferredVersion IsNot Nothing Then
                'Make sure that the IsBodyHtml property is being set correctly for our content
                message.IsBodyHtml = True
            Else
                'otherwise use the first plain text version as the body, if it exists
                preferredVersion = FindFirstPlainTextVersion()
            End If

            If preferredVersion IsNot Nothing Then
                message.Body = preferredVersion.GetBodyAsText()
                message.BodyEncoding = preferredVersion.BodyEncoding
            End If

            'Add body and alternative views (html and such) to the message
            Dim textVersions As IEnumerable(Of MessagePart) = FindAllTextVersions()

            For Each textVersion As MessagePart In textVersions
                'The textVersions also contain the preferred version, therefore
                ' we should skip that one
                If textVersion Is preferredVersion Then Continue For
                Dim stream As MemoryStream = New MemoryStream(textVersion.Body)
                Dim alternative As AlternateView = New AlternateView(stream)
                alternative.ContentId = textVersion.ContentId
                alternative.ContentType = textVersion.ContentType
                message.AlternateViews.Add(alternative)
            Next

            'Add attachments to the message
            Dim attachments As IEnumerable(Of MessagePart) = FindAllAttachments()

            For Each attachmentMessagePart As MessagePart In attachments
                Dim stream As MemoryStream = New MemoryStream(attachmentMessagePart.Body)
                Dim attachment As Attachment = New Attachment(stream, attachmentMessagePart.ContentType)
                attachment.ContentId = attachmentMessagePart.ContentId
                message.Attachments.Add(attachment)
            Next

            If Headers.From IsNot Nothing AndAlso Headers.From.HasValidMailAddress Then message.From = Headers.From.MailAddress
            If Headers.ReplyTo IsNot Nothing AndAlso Headers.ReplyTo.HasValidMailAddress Then message.ReplyTo = Headers.ReplyTo.MailAddress
            If Headers.Sender IsNot Nothing AndAlso Headers.Sender.HasValidMailAddress Then message.Sender = Headers.Sender.MailAddress

            For Each [to] As RfcMailAddress In Headers.[To]
                If [to].HasValidMailAddress Then message.[To].Add([to].MailAddress)
            Next

            For Each cc As RfcMailAddress In Headers.Cc
                If cc.HasValidMailAddress Then message.CC.Add(cc.MailAddress)
            Next

            For Each bcc As RfcMailAddress In Headers.Bcc
                If bcc.HasValidMailAddress Then message.Bcc.Add(bcc.MailAddress)
            Next

            Return message
        End Function


        ''' <summary>
		''' Finds the first text/plain <see cref="MessagePart"/> in this message.<br/>
		''' This Is a convenience method - it simply propagates the call to <see cref="FindFirstMessagePartWithMediaType"/>.<br/>
		''' <br/>
		''' If no text/plain version Is found, <see langword="null"/> Is returned.
		''' </summary>
		''' <returns>
		''' <see cref="MessagePart"/> which has a MediaType of text/plain Or <see langword="null"/>
		''' if such <see cref="MessagePart"/> could Not be found.
		''' </returns>
        Public Function FindFirstPlainTextVersion() As MessagePart
            Return FindFirstMessagePartWithMediaType("text/plain")
        End Function


        ''' <summary>
		''' Finds the first text/html <see cref="MessagePart"/> in this message.<br/>
		''' This Is a convenience method - it simply propagates the call to <see cref="FindFirstMessagePartWithMediaType"/>.<br/>
		''' <br/>
		''' If no text/html version Is found, <see langword="null"/> Is returned.
		''' </summary>
		''' <returns>
		''' <see cref="MessagePart"/> which has a MediaType of text/html Or <see langword="null"/>
		''' if such <see cref="MessagePart"/> could Not be found.
		''' </returns>
        Public Function FindFirstHtmlVersion() As MessagePart
            Return FindFirstMessagePartWithMediaType("text/html")
        End Function


        ''' <summary>
		''' Finds all the <see cref="MessagePart"/>'s which contains a text version.<br/>
		''' <br/>
		''' <see cref="Mime.MessagePart.IsText"/> for MessageParts which are considered to be text versions.<br/>
		''' <br/>
		''' Examples of MessageParts media types are:
		''' <list type="bullet">
		'''    <item>text/plain</item>
		'''    <item>text/html</item>
		'''    <item>text/xml</item>
		''' </list>
		''' </summary>
		''' <returns>A List of MessageParts where each part Is a text version</returns>
        Public Function FindAllTextVersions() As List(Of MessagePart)
            Return New TextVersionFinder().VisitMessage(Me)
        End Function


        ''' <summary>
		''' Finds all the <see cref="MessagePart"/>'s which are attachments to this message.<br/>
		''' <br/>
		''' <see cref="Mime.MessagePart.IsAttachment"/> for MessageParts which are considered to be attachments.
		''' </summary>
		''' <returns>A List of MessageParts where each Is considered an attachment</returns>
        Public Function FindAllAttachments() As List(Of MessagePart)
            Return New AttachmentFinder().VisitMessage(Me)
        End Function


        ''' <summary>
		''' Finds the first <see cref="MessagePart"/> in the <see cref="Message"/> hierarchy with the given MediaType.<br/>
		''' <br/>
		''' The search in the hierarchy Is a depth-first traversal.
		''' </summary>
		''' <param name="mediaType">The MediaType to search for. Case Is ignored.</param>
		''' <returns>
		''' A <see cref="MessagePart"/> with the given MediaType Or <see langword="null"/> if no such <see cref="MessagePart"/> was found
		''' </returns>
        Public Function FindFirstMessagePartWithMediaType(ByVal mediaType As String) As MessagePart

            Return New FindFirstMessagePartWithMediaType().VisitMessage(Me, mediaType)

        End Function


        ''' <summary>
		''' Finds all the <see cref="MessagePart"/>s in the <see cref="Message"/> hierarchy with the given MediaType.
		''' </summary>
		''' <param name="mediaType">The MediaType to search for. Case Is ignored.</param>
		''' <returns>
		''' A List of <see cref="MessagePart"/>s with the given MediaType.<br/>
		''' The List might be empty if no such <see cref="MessagePart"/>s were found.<br/>
		'''/ The order of the elements in the list Is the order which they are found using
		''' a depth first traversal of the <see cref="Message"/> hierarchy.
		''' </returns>
        Public Function FindAllMessagePartsWithMediaType(ByVal mediaType As String) As List(Of MessagePart)

            Return New FindAllMessagePartsWithMediaType().VisitMessage(Me, mediaType)

        End Function


        ''' <summary>
		''' Save this <see cref="Message"/> to a file.<br/>
		''' <br/>
		''' Can be loaded at a later time using the <see cref="Load(FileInfo)"/> method.
		''' </summary>
		''' <param name="file">The File location to save the <see cref="Message"/> to. Existent files will be overwritten.</param>
		''' <exception cref="ArgumentNullException">If <paramref name="file"/> Is <see langword="null"/></exception>
		''' <exception>Other exceptions relevant to using a <see cref="FileStream"/> might be thrown as well</exception>
        Public Sub Save(ByVal file As FileInfo)
            If file Is Nothing Then Throw New ArgumentNullException("file")

            Using stream As FileStream = New FileStream(file.FullName, FileMode.Create)
                Save(stream)
            End Using
        End Sub


        ''' <summary>
		''' Save this <see cref="Message"/> to a stream.<br/>
		''' </summary>
		''' <param name="messageStream">The stream to write to</param>
		''' <exception cref="ArgumentNullException">If <paramref name="messageStream"/> Is <see langword="null"/></exception>
		''' <exception>Other exceptions relevant to <see cref="Stream.Write"/> might be thrown as well</exception>
        Public Sub Save(ByVal messageStream As Stream)
            If messageStream Is Nothing Then Throw New ArgumentNullException("messageStream")
            messageStream.Write(RawMessage, 0, RawMessage.Length)
        End Sub


        ''' <summary>
        ''' Loads a <see cref="Message"/> from a file containing a raw email.
        ''' </summary>
        ''' <param name="file">The File location to load the <see cref="Message"/> from. The file must exist.</param>
        ''' <exception cref="ArgumentNullException">If <paramref name="file"/> Is <see langword="null"/></exception>
        ''' <exception cref="FileNotFoundException">If <paramref name="file"/> does Not exist</exception>
        ''' <exception>Other exceptions relevant to a <see cref="FileStream"/> might be thrown as well</exception>
        ''' <returns>A <see cref="Message"/> with the content loaded from the <paramref name="file"/></returns>
        Public Shared Function Load(ByVal file As FileInfo) As Message
            If file Is Nothing Then Throw New ArgumentNullException("file")
            If Not file.Exists Then Throw New FileNotFoundException("Cannot load message from non-existent file", file.FullName)

            Using stream As FileStream = New FileStream(file.FullName, FileMode.Open)
                Return Load(stream)
            End Using
        End Function


        ''' <summary>
		''' Loads a <see cref="Message"/> from a <see cref="Stream"/> containing a raw email.
		'''/summary>
		''' <param name="messageStream">The <see cref="Stream"/> from which to load the raw <see cref="Message"/></param>
		''' <exception cref="ArgumentNullException">If <paramref name="messageStream"/> Is <see langword="null"/></exception>
		''' <exception>Other exceptions relevant to <see cref="Stream.Read"/> might be thrown as well</exception>
		''' <returns>A <see cref="Message"/> with the content loaded from the <paramref name="messageStream"/></returns>
        Public Shared Function Load(ByVal messageStream As Stream) As Message
            If messageStream Is Nothing Then Throw New ArgumentNullException("messageStream")

            Using outStream As MemoryStream = New MemoryStream()
                Dim bytesRead As Integer
                Dim buffer As Byte() = New Byte(4095) {}

                While (bytesRead = messageStream.Read(buffer, 0, 4096)) > 0
                    outStream.Write(buffer, 0, bytesRead)
                End While

                Dim content As Byte() = outStream.ToArray()
                Return New Message(content)
            End Using
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace
