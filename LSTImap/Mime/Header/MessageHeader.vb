Imports System
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Net.Mail
Imports System.Net.Mime
Imports LSTImap.Mime.Decode

Namespace Mime.Header

    ''' <summary>
	''' Class that holds all headers for a message<br/>
	''' Headers which are unknown the the parser will be held in the <see cref="UnknownHeaders"/> collection.<br/>
	''' <br/>
	''' This class cannot be instantiated from outside the library.
	''' </summary>
	''' <remarks>
	''' See <a href="http://tools.ietf.org/html/rfc4021">RFC 4021</a> for a large list of headers.<br/>
	''' </remarks>
    Public NotInheritable Class MessageHeader

        ''' <summary>
		''' All headers which were Not recognized And explicitly dealt with.<br/>
		''' This should mostly be custom headers, which are marked as X-[name].<br/>
		''' <br/>
		''' This list will be empty if all headers were recognized And parsed.
		''' </summary>
		''' <remarks>
		''' If you as a user, feels that a header in this collection should
		''' be parsed, feel free to notify the developers.
		''' </remarks>
        Public Property UnknownHeaders As NameValueCollection

        ''' <summary>
		''' A human readable description of the body<br/>
		''' <br/>
		''' <see langword="null"/> if no Content-Description header was present in the message.
		''' </summary>
        Public Property ContentDescription As String

        ''' <summary>
        ''' ID of the content part (Like an attached image). Used with MultiPart messages.<br/>
        ''' <br/>
        ''' <see langword="null"/> if no Content-ID header field was present in the message.
        ''' </summary>
        ''' <see cref="MessageId">For an ID of the message</see>
        Public Property ContentId As String

        ''' <summary>
		''' Message keywords<br/>
		''' <br/>
		''' The list will be empty if no Keywords header was present in the message
		''' </summary>
        Public Property Keywords As List(Of String)

        ''' <summary>
		''' A List of emails to people who wishes to be notified when some event happens.<br/>
		''' These events could be email:
		''' <list type="bullet">
		'''   <item>deletion</item>
		'''   <item>printing</item>
		'''   <item>received</item>
		'''   <item>...</item>
		''' </list>
		''' The list will be empty if no Disposition-Notification-To header was present in the message
		''' </summary>
		''' <remarks>See <a href="http://tools.ietf.org/html/rfc3798">RFC 3798</a> for details</remarks>
        Public Property DispositionNotificationTo As List(Of RfcMailAddress)

        ''' <summary>
		''' This Is the Received headers. This tells the path that the email went.<br/>
		''' <br/>
		''' The list will be empty if no Received header was present in the message
		''' </summary>
        Public Property Received As List(Of Received)

        ''' <summary>
		''' Importance of this email.<br/>
		''' <br/>
		''' The importance level Is set to normal, if no Importance header field was mentioned Or it contained
		''' unknown information. This Is the expected behavior according to the RFC.
		''' </summary>
        Public Property Importance As MailPriority

        ''' <summary>
		''' This header describes the Content encoding during transfer.<br/>
		''' <br/>
		''' If no Content-Transfer-Encoding header was present in the message, it Is set
		''' to the default of <see cref="Header.ContentTransferEncoding.SevenBit">SevenBit</see> in accordance to the RFC.
		''' </summary>
		''' <remarks>See <a href="http://tools.ietf.org/html/rfc2045#section-6">RFC 2045 section 6</a> for details</remarks>
        Public Property ContentTransferEncoding As ContentTransferEncoding

        ''' <summary>
		''' Carbon Copy. This specifies who got a copy of the message.<br/>
		''' <br/>
		''' The list will be empty if no Cc header was present in the message
		''' </summary>
        Public Property Cc As List(Of RfcMailAddress)

        ''' <summary>
		''' Blind Carbon Copy. This specifies who got a copy of the message, but others
		''' cannot see who these persons are.<br/>
		''' <br/>
		''' The list will be empty if no Received Bcc was present in the message
		''' </summary>
        Public Property Bcc As List(Of RfcMailAddress)

        ''' <summary>
        ''' Specifies who this mail was for<br/>
        ''' <br/>
        ''' The list will be empty if no To header was present in the message
        '''</summary>
        Public Property [To] As List(Of RfcMailAddress)

        ''' <summary>
		''' Specifies who sent the email<br/>
		''' <br/>
		''' <see langword="null"/> if no From header field was present in the message
		''' </summary>
        Public Property From As RfcMailAddress

        ''' <summary>
		''' Specifies who a reply to the message should be sent to<br/>
		''' <br/>
		'''<see langword="null"/> if no Reply-To header field was present in the message
		''' </summary>
        Public Property ReplyTo As RfcMailAddress


        ''' <summary>
		''' The message identifier(s) of the original message(s) to which the
		''' current message Is a reply.<br/>
		''' <br/>
		''' The list will be empty if no In-Reply-To header was present in the message
		''' </summary>
        Public Property InReplyTo As List(Of String)

        ''' <summary>
        ''' The message identifier(s) of other message(s) to which the current
        ''' message Is related to.<br/>
        ''' <br/>
        ''' The list will be empty if no References header was present in the message
        ''' </summary>
        Public Property References As List(Of String)

        ''' <summary>
        ''' This Is the sender of the email address.<br/>
        ''' <br/>
        ''' <see langword="null"/> if no Sender header field was present in the message
        '''</summary>
        '''<remarks>
        ''' The RFC states that this field can be used if a secretary
        ''' Is sending an email for someone she Is working for.
        ''' The email here will then be the secretary's email, and
        ''' the Reply-To field would hold the address of the person she works for.<br/>
        '''RFC states that if the Sender Is the same as the From field,
        ''' sender should Not be included in the message.
        ''' </remarks>
        Public Property Sender As RfcMailAddress

        ''' <summary>
		''' The Content-Type header field.<br/>
		''' <br/>
		''' If Not set, the ContentType Is created by the default "text/plain; charset=us-ascii" which Is
		''' defined in <a href="http://tools.ietf.org/html/rfc2045#section-5.2">RFC 2045 section 5.2</a>.<br/>
		''' If set, the default Is overridden.
		'''</summary>
        Public Property ContentType As ContentType

        ''' <summary>
		''' Used to describe if a <see cref="MessagePart"/> Is to be displayed Or to be though of as an attachment.<br/>
		''' Also contains information about filename if such was sent.<br/>
		''' <br/>
		''' <see langword="null"/> if no Content-Disposition header field was present in the message
		''' </summary>
        Public Property ContentDisposition As ContentDisposition

        ''' <summary>
        ''' The Date when the email was sent.<br/>
        ''' This Is the raw value. <see cref="DateSent"/> for a parsed up <see cref="DateTime"/> value of this field.<br/>
        ''' <br/>
        ''' <see langword="DateTime.MinValue"/> if no Date header field was present in the message Or if the date could Not be parsed.
        ''' </summary>
        ''' <remarks> See <a href="http://tools.ietf.org/html/rfc5322#section-3.6.1">RFC 5322 section 3.6.1</a> for more details</remarks>
        Public Property Date1 As String

        ''' <summary>
		''' The Date when the email was sent.<br/>
		''' This Is the parsed equivalent of <see cref="Date"/>.<br/>
		''' Notice that the <see cref="TimeZone"/> of the <see cref="DateTime"/> object Is in UTC And has Not been converted
		''' to local <see cref="TimeZone"/>.
		''' </summary>
		'''<remarks>See <a href="http://tools.ietf.org/html/rfc5322#section-3.6.1">RFC 5322 section 3.6.1</a> for more details</remarks>
        Public Property DateSent As DateTime

        ''' <summary>
		''' An ID of the message that Is SUPPOSED to be in every message according to the RFC.<br/>
		''' The ID Is unique.<br/>
		''' <br/>
		''' <see langword="null"/> if no Message-ID header field was present in the message
		''' </summary>
        Public Property MessageId As String

        ''' <summary>
        ''' The Mime Version.<br/>
        ''' This field will almost always show 1.0<br/>
        ''' <br/>
        ''' <see langword="null"/> if no Mime-Version header field was present in the message
        ''' </summary>
        Public Property MimeVersion As String

        ''' <summary>
		''' A single <see cref="RfcMailAddress"/> with no username inside.<br/>
		''' This Is a trace header field, that should be in all messages.<br/>
		''' Replies should be sent to this address.<br/>
		'''<br/>
		''' <see langword="null"/> if no Return-Path header field was present in the message
		'''</summary>
        Public Property ReturnPath As RfcMailAddress

        ''' <summary>
		''' The subject line of the message in decoded, one line state.<br/>
		''' This should be in all messages.<br/>
		''' <br/>
		''' <see langword="null"/> if no Subject header field was present in the message
		''' </summary>
        Public Property Subject As String


        ''' <summary>
        ''' Parses a <see cref="NameValueCollection"/> to a MessageHeader
        ''' </summary>
        ''' <param name="headers">The collection that should be traversed And parsed</param>
        ''' <returns>A valid MessageHeader object</returns>
        ''' <exception cref="ArgumentNullException">If <paramref name="headers"/> Is <see langword="null"/></exception>
        Friend Sub New(ByVal headers As NameValueCollection)
            If headers Is Nothing Then Throw New ArgumentNullException("headers")
            'Create empty lists as defaults. We do not like null values
            ' List with an initial capacity set to zero will be replaced
            ' when a corrosponding header Is found
            [To] = New List(Of RfcMailAddress)(0)
            Cc = New List(Of RfcMailAddress)(0)
            Bcc = New List(Of RfcMailAddress)(0)
            Received = New List(Of Received)()
            Keywords = New List(Of String)()
            InReplyTo = New List(Of String)(0)
            References = New List(Of String)(0)
            DispositionNotificationTo = New List(Of RfcMailAddress)()
            UnknownHeaders = New NameValueCollection()
            'Default importancetype is Normal (assumed if not set)
            Importance = MailPriority.Normal
            '7BIT is the default ContentTransferEncoding (assumed if not set)
            ContentTransferEncoding = ContentTransferEncoding.SevenBit
            'text/plain; charset=us-ascii is the default ContentType
            ContentType = New ContentType("text/plain; charset=us-ascii")
            'Now parse the actual headers
            ParseHeaders(headers)
        End Sub


        ''' <summary>
		''' Parses a <see cref="NameValueCollection"/> to a <see cref="MessageHeader"/>
		''' </summary>
		''' <param name="headers">The collection that should be traversed And parsed</param>
		''' <returns>A valid <see cref="MessageHeader"/> object</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="headers"/> Is <see langword="null"/></exception>
        Private Sub ParseHeaders(ByVal headers As NameValueCollection)
            If headers Is Nothing Then Throw New ArgumentNullException("headers")

            'Now begin to parse the header values
            For Each headerName As String In headers.Keys
                Dim headerValues As String() = headers.GetValues(headerName)

                If headerValues IsNot Nothing Then

                    For Each headerValue As String In headerValues
                        ParseHeader(headerName, headerValue)
                    Next
                End If
            Next
        End Sub


        ''' <summary>
		''' Parses a single header And sets member variables according to it.
		''' </summary>
		''' <param name="headerName">The name of the header</param>
		''' <param name="headerValue">The value of the header in unfolded state (only one line)</param>
		''' <exception cref="ArgumentNullException">If <paramref name="headerName"/> Or <paramref name="headerValue"/> Is <see langword="null"/></exception>
        Private Sub ParseHeader(ByVal headerName As String, ByVal headerValue As String)
            If headerName Is Nothing Then Throw New ArgumentNullException("headerName")
            If headerValue Is Nothing Then Throw New ArgumentNullException("headerValue")

            Select Case headerName.ToUpperInvariant()
                Case "TO"
                    [To] = RfcMailAddress.ParseMailAddresses(headerValue)
                Case "CC"
                    Cc = RfcMailAddress.ParseMailAddresses(headerValue)
                Case "BCC"
                    Bcc = RfcMailAddress.ParseMailAddresses(headerValue)
                Case "FROM"
                    'There is only one MailAddress in the from field
                    From = RfcMailAddress.ParseMailAddress(headerValue)
                Case "REPLY-TO"
                    'This field may actually be a list of addresses, but no
                    ' such case has been encountered
                    ReplyTo = RfcMailAddress.ParseMailAddress(headerValue)
                Case "SENDER"
                    Sender = RfcMailAddress.ParseMailAddress(headerValue)
                'See http: //tools.ietf.org/html/rfc5322#section-3.6.5
                ' RFC 5322:
                ' The "Keywords:" field contains a comma-separated list of one Or more
                ' words Or quoted-strings.
                ' The field are intended to have only human-readable content
                ' with information about the message
                Case "KEYWORDS"
                    Dim keywordsTemp As String() = headerValue.Split(","c)

                    For Each keyword As String In keywordsTemp
                        Keywords.Add(Utility.RemoveQuotesIfAny(keyword.Trim()))
                    Next

                Case "RECEIVED"
                    Received.Add(New Received(headerValue.Trim()))
                Case "IMPORTANCE"
                    Importance = HeaderFieldParser.ParseImportance(headerValue.Trim())
                Case "DISPOSITION-NOTIFICATION-TO"
                    DispositionNotificationTo = RfcMailAddress.ParseMailAddresses(headerValue)
                Case "MIME-VERSION"
                    MimeVersion = headerValue.Trim()
                Case "SUBJECT"
                    Subject = EncodedWord.Decode(headerValue)
                Case "RETURN-PATH"
                    ReturnPath = RfcMailAddress.ParseMailAddress(headerValue)
                Case "MESSAGE-ID"
                    MessageId = HeaderFieldParser.ParseId(headerValue)
                Case "IN-REPLY-TO"
                    InReplyTo = HeaderFieldParser.ParseMultipleIDs(headerValue)
                Case "REFERENCES"
                    References = HeaderFieldParser.ParseMultipleIDs(headerValue)
                Case "DATE"
                    Date1 = headerValue.Trim()
                    DateSent = Rfc2822DateTime.StringToDate(headerValue)
                Case "CONTENT-TRANSFER-ENCODING"
                    ContentTransferEncoding = HeaderFieldParser.ParseContentTransferEncoding(headerValue.Trim())
                Case "CONTENT-DESCRIPTION"
                    ContentDescription = EncodedWord.Decode(headerValue.Trim())
                Case "CONTENT-TYPE"
                    ContentType = HeaderFieldParser.ParseContentType(headerValue)
                Case "CONTENT-DISPOSITION"
                    ContentDisposition = HeaderFieldParser.ParseContentDisposition(headerValue)
                Case "CONTENT-ID"
                    ContentId = HeaderFieldParser.ParseId(headerValue)
                Case Else
                    'This is an unknown header

                    ' Custom headers are allowed. That means headers
                    ' that are Not mentionen in the RFC.
                    ' Such headers start with the letter "X"
                    ' We do Not have any special parsing of such

                    ' Add it to unknown headers
                    UnknownHeaders.Add(headerName, headerValue)
            End Select
        End Sub
    End Class
End Namespace
