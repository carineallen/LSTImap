Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Mime
Imports System.Text
Imports LSTImap.Mime.Decode
Imports LSTImap.Mime.Header
Imports LSTImap.Common
Imports System.Runtime.InteropServices

Namespace Mime

    ''' <summary>
	''' A MessagePart Is a part of an email message used to describe the whole email parse tree.<br/>
	''' <br/>
	''' <b>Email messages are tree structures</b>:<br/>
	''' Email messages may contain large tree structures, And the MessagePart are the nodes of the this structure.<br/>
	''' A MessagePart may either be a leaf in the structure Or a internal node with links to other MessageParts.<br/>
	''' The root of the message tree Is the <see cref="Message"/> class.<br/>
	''' <br/>
	''' <b>Leafs</b>:<br/>
	''' If a MessagePart Is a leaf, the part Is Not a <see cref="IsMultiPart">MultiPart</see> message.<br/>
	'''eafs are where the contents of an email are placed.<br/>
	''' This includes, but Is Not limited to: attachments, text Or images referenced from HTML.<br/>
	''' The content of an attachment can be fetched by using the <see cref="Body"/> property.<br/>
	''' If you want to have the text version of a MessagePart, use the <see cref="GetBodyAsText"/> method which will<br/>
	'''onvert the <see cref="Body"/> into a string using the encoding the message was sent with.<br/>
	''' <br/>
	''' <b>Internal nodes</b>:<br/>
	''' If a MessagePart Is an internal node in the email tree structure, then the part Is a <see cref="IsMultiPart">MultiPart</see> message.<br/>
	''' The <see cref="MessageParts"/> property will then contain links to the parts it contain.<br/>
	''' The <see cref="Body"/> property of the MessagePart will Not be set.<br/>
	''' <br/>
	''' See the example for a parsing example.<br/>
	''' This class cannot be instantiated from outside the library.
	''' </summary>
	''' <example>
	''' This example illustrates how the message parse tree looks Like given a specific message<br/>
	''' <br/>
	''' The message source in this example Is:<br/>
	''' <code>
	''' MIME-Version: 1.0
	'''	Content-Type: multipart/mixed; boundary="frontier"
	'''	
	'''	This Is a message with multiple parts in MIME format.
	'''	--frontier
	''' Content-Type: text/plain
	'''	
	'''	This Is the body of the message.
	'''	--frontier
	'''Content-Type: application/octet-stream
	'''	Content-Transfer-Encoding: base64
	'''	
	'''	PGh0bWw+CiAgPGHLYWQ+CiAgPC9oZWFkPgogIDxib2R5PgogICAgPHA+VGhpcyBpcyB0aGUg
	'''	Ym9keSBvZiB0aGUgbWVzc2FnZS48L3A+CiAgPC9ib2R5Pgo8L2h0bWw+Cg==
	'''	--frontier--
	''' </code>
	'''The tree will look as follows, where the content-type media type of the message Is listed<br/>
	''' <code>
	''' - Message root
	'''   - multipart/mixed MessagePart
	'''     - text/plain MessagePart
	'''     - application/octet-stream MessagePart
	''' </code>
	''' It Is possible to have more complex message trees Like the following:<br/>
	''' <code>
	''' - Message root
	'''   - multipart/mixed MessagePart
	'''     - text/plain MessagePart
	'''     - text/plain MessagePart
	'''     - multipart/parallel
	'''       - audio/basic
	'''       - image/tiff
	'''     - text/enriched
	'''     - message/rfc822
	''' </code>
	''' But it Is also possible to have very simple message trees Like:<br/>
	''' <code>
	''' - Message root
	'''   - text/plain
	''' </code>
	''' </example>
    Public Class MessagePart

        ''' <summary>
		''' The Content-Type header field.<br/>
		''' <br/>
		''' If Not set, the ContentType Is created by the default "text/plain; charset=us-ascii" which Is
		'''defined in <a href="http://tools.ietf.org/html/rfc2045#section-5.2">RFC 2045 section 5.2</a>.<br/>
		''' <br/>
		''' If set, the default Is overridden.
		''' </summary>
        Public Property ContentType As ContentType

        ''' <summary>
		''' A human readable description of the body<br/>
		''' <br/>
		''' <see langword="null"/> if no Content-Description header was present in the message.<br/>
		''' </summary>
        Public Property ContentDescription As String

        ''' <summary>
		''' This header describes the Content encoding during transfer.<br/>
		''' <br/>
		''' If no Content-Transfer-Encoding header was present in the message, it Is set
		''' to the default of <see cref="Header.ContentTransferEncoding.SevenBit">SevenBit</see> in accordance to the RFC.
		''' </summary>
		''' <remarks>See <a href="http://tools.ietf.org/html/rfc2045#section-6">RFC 2045 section 6</a> for details</remarks>
        Public Property ContentTransferEncoding As ContentTransferEncoding

        ''' <summary>
		''' ID of the content part (Like an attached image). Used with MultiPart messages.<br/>
		''' <br/>
		''' <see langword="null"/> if no Content-ID header field was present in the message.
		''' </summary>
        Public Property ContentId As String

        ''' <summary>
		''' Used to describe if a <see cref="MessagePart"/> Is to be displayed Or to be though of as an attachment.<br/>
		''' Also contains information about filename if such was sent.<br/>
		''' <br/>
		''' <see langword="null"/> if no Content-Disposition header field was present in the message
		''' </summary>
        Public Property ContentDisposition As ContentDisposition

        ''' <summary>
		''' This Is the encoding used to parse the message body if the <see cref="MessagePart"/><br/>
		''' Is Not a MultiPart message. It Is derived from the <see cref="ContentType"/> character set property.
		''' </summary>
        Public Property BodyEncoding As Encoding

        ''' <summary>
		''' This Is the parsed body of this <see cref="MessagePart"/>.<br/>
		''' It Is parsed in that way, if the body was ContentTransferEncoded, it has been decoded to the
		''' correct bytes.<br/>
		''' <br/>
		''' It will be <see langword="null"/> if this <see cref="MessagePart"/> Is a MultiPart message.<br/>
		''' Use <see cref="IsMultiPart"/> to check if this <see cref="MessagePart"/> Is a MultiPart message.
		''' </summary>
        Public Property Body As Byte()


        ''' <summary>
		''' Describes if this <see cref="MessagePart"/> Is a MultiPart message<br/>
		'''br/>
		''' The <see cref="MessagePart"/> Is a MultiPart message if the <see cref="ContentType"/> media type property starts with "multipart/"
		''' </summary>
        Public ReadOnly Property IsMultiPart As Boolean
            Get
                Return ContentType.MediaType.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property


        ''' <summary>
		''' A <see cref="MessagePart"/> Is considered to be holding text in it's body if the MediaType
		''' starts either "text/" Or Is equal to "message/rfc822"
		''' </summary>
        Public ReadOnly Property IsText As Boolean
            Get
                Dim mediaType As String = ContentType.MediaType
                Return mediaType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) OrElse mediaType.Equals("message/rfc822", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property


        ''' <summary>
		''' A <see cref="MessagePart"/> Is considered to be an attachment, if<br/>
		''' - it Is Not holding <see cref="IsText">text</see> And Is Not a <see cref="IsMultiPart">MultiPart</see> message<br/>
		''' Or<br/>
		''' - it has a Content-Disposition header that says it Is an attachment
		''' </summary>
        Public ReadOnly Property IsAttachment As Boolean
            Get
                'Inline is the opposite of attachment
                Return (Not IsText AndAlso Not IsMultiPart) OrElse (ContentDisposition IsNot Nothing AndAlso Not ContentDisposition.Inline)
            End Get
        End Property


        ''' <summary>
		''' This Is a convenient-property for figuring out a FileName for this <see cref="MessagePart"/>.<br/>
		''' If the <see cref="MessagePart"/> Is a MultiPart message, then it makes no sense to try to find a FileName.<br/>
		''' <br/>
		''' The FileName can be specified in the <see cref="ContentDisposition"/> Or in the <see cref="ContentType"/> properties.<br/>
		''' If none of these places two places tells about the FileName, a default "(no name)" Is returned.
		''' </summary>
        Public Property FileName As String

        ''' <summary>
		''' If this <see cref="MessagePart"/> Is a MultiPart message, then this property
		''' has a list of each of the Multiple parts that the message consists of.<br/>
		''' <br/>
		'''t Is <see langword="null"/> if it Is Not a MultiPart message.<br/>
		''' Use <see cref="IsMultiPart"/> to check if this <see cref="MessagePart"/> Is a MultiPart message.
		''' </summary>
        Public Property MessageParts As List(Of MessagePart)


        ''' <summary>
        ''' Used to construct the topmost message part
        ''' </summary>
        ''' <param name="rawBody">The body that needs to be parsed</param>
        ''' <param name="headers">The headers that should be used from the message</param>
        ''' <exception cref="ArgumentNullException">If <paramref name="rawBody"/> Or <paramref name="headers"/> Is <see langword="null"/></exception>
        Friend Sub New(ByVal rawBody As Byte(), ByVal headers As MessageHeader)
            If rawBody Is Nothing Then Throw New ArgumentNullException("rawBody")
            If headers Is Nothing Then Throw New ArgumentNullException("headers")
            ContentType = headers.ContentType
            ContentDescription = headers.ContentDescription
            ContentTransferEncoding = headers.ContentTransferEncoding
            ContentId = headers.ContentId
            ContentDisposition = headers.ContentDisposition
            FileName = FindFileName(ContentType, ContentDisposition, "(no name)")
            BodyEncoding = ParseBodyEncoding(ContentType.CharSet)
            ParseBody(rawBody)
        End Sub

        ''' <summary>
		''' Parses a character set into an encoding
		''' </summary>
		''' <param name="characterSet">The character set that needs to be parsed. <see langword="null"/> Is allowed.</param>
		''' <returns>The encoding specified by the <paramref name="characterSet"/> parameter, Or ASCII if the character set was <see langword="null"/> Or empty</returns>

        Private Shared Function ParseBodyEncoding(ByVal characterSet As String) As Encoding
            'Default encoding in Mime messages is US-ASCII
            Dim encoding As Encoding = Encoding.ASCII
            'If the character set was specified, find the encoding that the character
            ' set describes, And use that one instead
            If Not String.IsNullOrEmpty(characterSet) Then encoding = EncodingFinder.FindEncoding(characterSet)
            Return encoding
        End Function


        ''' <summary>
		''' Figures out the filename of this message part from some headers.
		''' <see cref="FileName"/> property.
		''' </summary>
		''' <param name="contentType">The Content-Type header</param>
		''' <param name="contentDisposition">The Content-Disposition header</param>
		''' <param name="defaultName">The default filename to use, if no other could be found</param>
		''' <returns>The filename found, Or the default one if Not such filename could be found in the headers</returns>
		''' <exception cref="ArgumentNullException">if <paramref name="contentType"/> Is <see langword="null"/></exception>
        Private Shared Function FindFileName(ByVal contentType As ContentType, ByVal contentDisposition As ContentDisposition, ByVal defaultName As String) As String
            If contentType Is Nothing Then Throw New ArgumentNullException("contentType")
            If contentDisposition IsNot Nothing AndAlso contentDisposition.FileName IsNot Nothing Then Return contentDisposition.FileName
            If contentType.Name IsNot Nothing Then Return contentType.Name
            Return defaultName
        End Function


        ''' <summary>
		''' Parses a byte array as a body of an email message.
		''' </summary>
		''' <param name="rawBody">The byte array to parse as body of an email message. This array may Not contain headers.</param>
        Private Sub ParseBody(ByVal rawBody As Byte())
            If IsMultiPart Then
                'Parses a MultiPart message
                ParseMultiPartBody(rawBody)
            Else
                'Parses a non MultiPart message
                ' Decode the body accodingly And set the Body property
                Body = DecodeBody(rawBody, ContentTransferEncoding)
            End If
        End Sub


        ''' <summary>
		''' Parses the <paramref name="rawBody"/> byte array as a MultiPart message.<br/>
		''' It Is Not valid to call this method if <see cref="IsMultiPart"/> returned <see langword="false"/>.<br/>
		''' Fills the <see cref="MessageParts"/> property of this <see cref="MessagePart"/>.
		''' </summary>
		''' <param name="rawBody">The byte array which Is to be parsed as a MultiPart message</param>
        Private Sub ParseMultiPartBody(ByVal rawBody As Byte())
            'Fetch out the boundary used to delimit the messages within the body
            Dim multipartBoundary As String = ContentType.Boundary
            'Fetch the individual MultiPart message parts using the MultiPart boundary
            Dim bodyParts As List(Of Byte()) = GetMultiPartParts(rawBody, multipartBoundary)
            'Initialize the MessageParts property, with room to as many bodies as we have found
            MessageParts = New List(Of MessagePart)(bodyParts.Count)

            'Now parse each byte array as a message body and add it the the MessageParts property
            For Each bodyPart As Byte() In bodyParts
                Dim messagePart As MessagePart = GetMessagePart(bodyPart)
                MessageParts.Add(messagePart)
            Next
        End Sub


        ''' <summary>
		''' Given a byte array describing a full message.<br/>
		''' Parses the byte array into a <see cref="MessagePart"/>.
		''' </summary>
		''' <param name="rawMessageContent">The byte array containing both headers And body of a message</param>
		''' <returns>A <see cref="MessagePart"/> which was described by the <paramref name="rawMessageContent"/> byte array</returns>
        Private Shared Function GetMessagePart(ByVal rawMessageContent As Byte()) As MessagePart
            'Find the headers and the body parts of the byte array
            Dim headers As MessageHeader
            Dim body As Byte()
            HeaderExtractor.ExtractHeadersAndBody(rawMessageContent, headers, body)
            'Create a new MessagePart from the headers and the body
            Return New MessagePart(body, headers)
        End Function


        ''' <summary>
        ''' Gets a list of byte arrays where each entry in the list Is a full message of a message part
        ''' </summary>
        ''' <param name="rawBody">The raw byte array describing the body of a message which Is a MultiPart message</param>
        ''' <param name="multipPartBoundary">The delimiter that splits the different MultiPart bodies from each other</param>
        ''' <returns>A list of byte arrays, each a full message of a <see cref="MessagePart"/></returns>
        ''' <exception cref="ArgumentNullException">If <paramref name="rawBody"/> Is <see langword="null"/></exception>
        Private Shared Function GetMultiPartParts(ByVal rawBody As Byte(), ByVal multipPartBoundary As String) As List(Of Byte())
            If rawBody Is Nothing Then Throw New ArgumentNullException("rawBody")
            'This is the list we want to return
            Dim messageBodies As List(Of Byte()) = New List(Of Byte())()

            Using stream As MemoryStream = New MemoryStream(rawBody)
                Dim lastMultipartBoundaryEncountered As Boolean
                'Find the start of the first message in this multipart
                ' Since the method returns the first character on a the line containing the MultiPart boundary, we
                ' need to add the MultiPart boundary with prepended "--" And appended CRLF pair to the position returned.
                Dim startLocation As Integer = FindPositionOfNextMultiPartBoundary(stream, multipPartBoundary, lastMultipartBoundaryEncountered) + ("--" & multipPartBoundary & vbCrLf).Length

                While True
                    'When we have just parsed the last multipart entry, stop parsing on
                    If lastMultipartBoundaryEncountered Then Exit While
                    'Find the end location of the current multipart
                    ' Since the method returns the first character on a the line containing the MultiPart boundary, we
                    ' need to go a CRLF pair back, so that we do Not get that into the body of the message part
                    Dim stopLocation As Integer = FindPositionOfNextMultiPartBoundary(stream, multipPartBoundary, lastMultipartBoundaryEncountered) - vbCrLf.Length

                    'If we could not find the next multipart boundary, but we had not yet discovered the last boundary, then
                    ' we will consider the rest of the bytes as contained in a last message part.
                    If stopLocation <= -1 Then
                        'Include everything except the last CRLF.
                        stopLocation = CInt(stream.Length) - vbCrLf.Length
                        'We consider this as the last part
                        lastMultipartBoundaryEncountered = True
                        'Special case: when the last multipart delimiter is not ending with "--", but is indeed the last
                        ' one, then the next multipart would contain nothing, And we should Not include such one.
                        If startLocation >= stopLocation Then Exit While
                    End If

                    'We have now found the start and end of a message part
                    ' Now we create a byte array with the correct length And put the message part's bytes into
                    ' it And add it to our list we want to return
                    Dim length As Integer = stopLocation - startLocation
                    Dim messageBody As Byte() = New Byte(length - 1) {}
                    Array.Copy(rawBody, startLocation, messageBody, 0, length)
                    messageBodies.Add(messageBody)
                    'We want to advance to the next message parts start.
                    ' We can find this by jumping forward the MultiPart boundary from the last
                    ' message parts end position
                    startLocation = stopLocation + (vbCrLf & "--" & multipPartBoundary & vbCrLf).Length
                End While
            End Using

            'We are done
            Return messageBodies
        End Function


        ''' <summary>
		''' Method that Is able to find a specific MultiPart boundary in a Stream.<br/>
		''' The Stream passed should Not be used for anything else then for looking for MultiPart boundaries
		''' <param name="stream">The stream to find the next MultiPart boundary in. Do Not use it for anything else then with this method.</param>
		''' <param name="multiPartBoundary">The MultiPart boundary to look for. This should be found in the <see cref="ContentType"/> header</param>
		''' <param name="lastMultipartBoundaryFound">Is set to <see langword="true"/> if the next MultiPart boundary was indicated to be the last one, by having -- appended to it. Otherwise set to <see langword="false"/></param>
		''' </summary>
		''' <returns>The position of the first character of the line that contained MultiPartBoundary Or -1 if no (more) MultiPart boundaries was found</returns>
        Private Shared Function FindPositionOfNextMultiPartBoundary(ByVal stream As Stream, ByVal multiPartBoundary As String, <Out> ByRef lastMultipartBoundaryFound As Boolean) As Integer
            lastMultipartBoundaryFound = False

            While True
                'Get the current position. This is the first position on the line - no characters of the line will
                ' have been read yet
                Dim currentPos As Integer = CInt(stream.Position)
                'Read the line
                Dim line As String = StreamUtility.ReadLineAsAscii(stream)
                'If we kept reading until there was no more lines, we did not meet
                ' the MultiPart boundary. -1 Is then returned to describe this.
                If line Is Nothing Then Return -1

                'The MultiPart boundary is the MultiPartBoundary with "--" in front of it
                ' which Is to be at the very start of a line
                If line.StartsWith("--" & multiPartBoundary, StringComparison.Ordinal) Then
                    'Check if the found boundary was also the last one
                    lastMultipartBoundaryFound = line.StartsWith("--" & multiPartBoundary & "--", StringComparison.OrdinalIgnoreCase)
                    Return currentPos
                End If
            End While
        End Function


        ''' <summary>
		''' Decodes a byte array into another byte array based upon the Content Transfer encoding
		''' </summary>
		''' <param name="messageBody">The byte array to decode into another byte array</param>
		''' <param name="contentTransferEncoding">The <see cref="ContentTransferEncoding"/> of the byte array</param>
		''' <returns>A byte array which comes from the <paramref name="contentTransferEncoding"/> being used on the <paramref name="messageBody"/></returns>
		''' <exception cref="ArgumentNullException">If <paramref name="messageBody"/> Is <see langword="null"/></exception>
		''' <exception cref="ArgumentOutOfRangeException">Thrown if the <paramref name="contentTransferEncoding"/> Is unsupported</exception>
        Private Shared Function DecodeBody(ByVal messageBody As Byte(), ByVal contentTransferEncoding As ContentTransferEncoding) As Byte()
            If messageBody Is Nothing Then Throw New ArgumentNullException("messageBody")

            Select Case contentTransferEncoding
                Case ContentTransferEncoding.QuotedPrintable
                    'If encoded in QuotedPrintable, everything in the body is in US-ASCII
                    Return QuotedPrintable.DecodeContentTransferEncoding(Encoding.ASCII.GetString(messageBody))
                Case ContentTransferEncoding.Base64
                    'If encoded in Base64, everything in the body is in US-ASCII
                    Return Base64.Decode(Encoding.ASCII.GetString(messageBody))
                Case ContentTransferEncoding.SevenBit, ContentTransferEncoding.Binary, ContentTransferEncoding.EightBit
                    'We do not have to do anything
                    Return messageBody
                Case Else
                    Throw New ArgumentOutOfRangeException("contentTransferEncoding")
            End Select
        End Function


        ''' <summary>
        ''' Gets this MessagePart's <see cref="Body"/> as text.<br/>
        ''' This Is simply the <see cref="BodyEncoding"/> being used on the raw bytes of the <see cref="Body"/> property.<br/>
        ''' This method Is only valid to call if it Is Not a MultiPart message And therefore contains a body.<br/>
        ''' </summary>
        ''' <returns>The <see cref="Body"/> property as a string</returns>
        Public Function GetBodyAsText() As String
            Return BodyEncoding.GetString(Body)
        End Function


        ''' <summary>
		''' Save this <see cref="MessagePart"/>'s contents to a file.<br/>
		''' There are no methods to reload the file.
		''' </summary>
		''' <param name="file">The File location to save the <see cref="MessagePart"/> to. Existent files will be overwritten.</param>
		''' <exception cref="ArgumentNullException">If <paramref name="file"/> Is <see langword="null"/></exception>
		''' <exception>Other exceptions relevant to using a <see cref="FileStream"/> might be thrown as well</exception>
        Public Sub Save(ByVal file As FileInfo)
            If file Is Nothing Then Throw New ArgumentNullException("file")

            Using stream As FileStream = New FileStream(file.FullName, FileMode.Create)
                Save(stream)
            End Using
        End Sub


        ''' <summary>
		''' Save this <see cref="MessagePart"/>'s contents to a stream.<br/>
		''' </summary>
		''' <param name="messageStream">The stream to write to</param>
		''' <exception cref="ArgumentNullException">If <paramref name="messageStream"/> Is <see langword="null"/></exception>
		''' <exception>Other exceptions relevant to <see cref="Stream.Write"/> might be thrown as well</exception>
        Public Sub Save(ByVal messageStream As Stream)
            If messageStream Is Nothing Then Throw New ArgumentNullException("messageStream")
            messageStream.Write(Body, 0, Body.Length)
        End Sub
    End Class
End Namespace
