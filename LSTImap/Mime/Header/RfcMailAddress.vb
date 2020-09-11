Imports System
Imports System.Collections.Generic
Imports System.Net.Mail
Imports LSTImap.Mime.Decode
Imports LSTImap.Common.Logging

Namespace Mime.Header

    ''' <summary>
	''' This class Is used for RFC compliant email addresses.<br/>
	''' <br/>
	''' The class cannot be instantiated from outside the library.
	''' </summary>
	''' <remarks>
	''' The <seealso cref="MailAddress"/> does Not cover all the possible formats 
	''' for <a href="http://tools.ietf.org/html/rfc5322#section-3.4">RFC 5322 section 3.4</a> compliant email addresses.
	''' This class Is used as an address wrapper to account for that deficiency.
	''' </remarks>
    Public Class RfcMailAddress

        '''<summary>
		'''The email address of this <see cref="RfcMailAddress"/><br/>
		''' It Is possibly string.Empty since RFC mail addresses does Not require an email address specified.
		'''</summary>
		'''<example>
		''' Example header with email address:<br/>
		''' To: <c> Test test@mail.com</c><br/>
		''' Address will be <c>test@mail.com</c><br/>
		'''</example>
		'''<example>
		''' Example header without email address:<br/>
		''' To: <c> Test</c><br/>
		''' Address will be <see cref="String.Empty"/>.
		'''</example>
        Public Property Address As String

        '''<summary>
		''' The display name of this <see cref="RfcMailAddress"/><br/>
		''' It Is possibly <see cref="String.Empty"/> since RFC mail addresses does Not require a display name to be specified.
		'''</summary>
		'''<example>
		''' Example header with display name:<br/>
		''' To: <c> Test test@mail.com</c><br/>
		''' DisplayName will be <c>Test</c>
		'''</example>
		'''<example>
		''' Example header without display name:<br/>
		''' To: <c> test@test.com</c><br/>
		''' DisplayName will be <see cref="String.Empty"/>
		'''</example>
        Public Property DisplayName As String

        ''' <summary>
		''' This Is the Raw string used to describe the <see cref="RfcMailAddress"/>.
		''' </summary>
        Public Property Raw As String

        ''' <summary>
		''' The <see cref="MailAddress"/> associated with the <see cref="RfcMailAddress"/>. 
		''' </summary>
		''' <remarks>
		''' The value of this property can be <see lanword="null"/> in instances where the <see cref="MailAddress"/> cannot represent the address properly.<br/>
		''' Use <see cref="HasValidMailAddress"/> property to see if this property Is valid.
		''' </remarks>
        Public Property MailAddress As MailAddress


        ''' <summary>
		''' Specifies if the object contains a valid <see cref="MailAddress"/> reference.
		''' </summary>
        Public ReadOnly Property HasValidMailAddress As Boolean
            Get
                Return MailAddress IsNot Nothing
            End Get
        End Property


        ''' <summary>
		''' Constructs an <see cref="RfcMailAddress"/> object from a <see cref="MailAddress"/> object.<br/>
		''' This constructor Is used when we were able to construct a <see cref="MailAddress"/> from a string.
		'''</summary>
		''' <param name="mailAddress">The address that <paramref name="raw"/> was parsed into</param>
		''' <param name="raw">The raw unparsed input which was parsed into the <paramref name="mailAddress"/></param>
		''' <exception cref="ArgumentNullException">If <paramref name="mailAddress"/> Or <paramref name="raw"/> Is <see langword="null"/></exception>
        Private Sub New(ByVal mailAddress As MailAddress, ByVal raw As String)
            If mailAddress Is Nothing Then Throw New ArgumentNullException("mailAddress")
            If raw Is Nothing Then Throw New ArgumentNullException("raw")
            mailAddress = mailAddress
            Address = mailAddress.Address
            DisplayName = mailAddress.DisplayName
            raw = raw
        End Sub

        ''' <summary>
		''' When we were unable to parse a string into a <see cref="MailAddress"/>, this constructor can be
		''' used. The Raw string Is then used as the <see cref="DisplayName"/>.
		''' </summary>
		''' <param name="raw">The raw unparsed input which could Not be parsed</param>
		''' <exception cref="ArgumentNullException">If <paramref name="raw"/> Is <see langword="null"/></exception>
        Private Sub New(ByVal raw As String)
            If raw Is Nothing Then Throw New ArgumentNullException("raw")
            MailAddress = Nothing
            Address = String.Empty
            DisplayName = raw
            raw = raw
        End Sub


        ''' <summary>
		''' A string representation of the <see cref="RfcMailAddress"/> object
		''' </summary>
		''' <returns>Returns the string representation for the object</returns>
        Public Overrides Function ToString() As String
            If HasValidMailAddress Then Return MailAddress.ToString()
            Return Raw
        End Function


		''' <summary>
		''' Parses an email address from a MIME header<br/>
		''' <br/>
		''' Examples of input:
		''' <c>Eksperten mailrobot &lt;noreply@mail.eksperten.dk&gt;</c><br/>
		''' <c>"Eksperten mailrobot" &lt;noreply@mail.eksperten.dk&gt;</c><br/>
		''' <c>&lt;noreply@mail.eksperten.dk&gt;</c><br/>
		''' <c>noreply@mail.eksperten.dk</c><br/>
		''' <br/>
		''' It might also contain encoded text, which will then be decoded.
		''' </summary>
		''' <param name="input">The value to parse out And email And/Or a username</param>
		''' <returns>A <see cref="RfcMailAddress"/></returns>
		''' <exception cref="ArgumentNullException">If <paramref name="input"/> Is <see langword="null"/></exception>
		''' <remarks>
		''' <see href="http://tools.ietf.org/html/rfc5322#section-3.4">RFC 5322 section 3.4</see> for more details on email syntax.<br/>
		''' <see cref="EncodedWord.Decode">For more information about encoded text</see>.
		''' </remarks>
		Friend Shared Function ParseMailAddress(ByVal input As String) As RfcMailAddress
			If input Is Nothing Then Throw New ArgumentNullException("input")
			'Decode the value, if it was encoded
			input = EncodedWord.Decode(input.Trim())
			'ind the location of the email address
			Dim indexStartEmail As Integer = input.LastIndexOf("<")
			Dim indexEndEmail As Integer = input.LastIndexOf(">")

			Try

				If indexStartEmail >= 0 AndAlso indexEndEmail >= 0 Then
					Dim username As String

					'Check if there is a username in front of the email address
					If indexStartEmail > 0 Then
						'Parse out the user
						username = input.Substring(0, indexStartEmail).Trim()
					Else
						'There was no user
						username = String.Empty
					End If

					'Parse out the email address without the "<"  and ">"
					indexStartEmail = indexStartEmail + 1
					Dim emailLength As Integer = indexEndEmail - indexStartEmail
					Dim emailAddress As String = input.Substring(indexStartEmail, emailLength).Trim()

					'There has been cases where there was no emailaddress between the < and >
					If Not String.IsNullOrEmpty(emailAddress) Then
						'If the username is quoted, MailAddress' constructor will remove them for us
						Return New RfcMailAddress(New MailAddress(emailAddress, username), input)
					End If
				End If

				'This might be on the form noreply@mail.eksperten.dk
				' Check if there Is an email, if notm there Is no need to try
				If input.Contains("@") Then Return New RfcMailAddress(New MailAddress(input), input)
			Catch __unusedFormatException1__ As FormatException
				'Sometimes invalid emails are sent, like sqlmap-user@sourceforge.net. (last period is illigal)
				DefaultLogger.Log.LogError("RfcMailAddress: Improper mail address: """ & input & """")
            End Try

			'It could be that the format used was simply a name
			' which Is indeed valid according to the RFC
			' Example:
			' Eksperten mailrobot
			Return New RfcMailAddress(input)
        End Function


		''' <summary>
		''' Parses input of the form<br/>
		''' <c>Eksperten mailrobot &lt;noreply@mail.eksperten.dk&gt;, ...</c><br/>
		''' to a list of RFCMailAddresses
		''' </summary>
		''' <param name="input">The input that Is a comma-separated list of EmailAddresses to parse</param>
		''' <returns>A List of <seealso cref="RfcMailAddress"/> objects extracted from the <paramref name="input"/> parameter.</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="input"/> Is <see langword="null"/></exception>
		Friend Shared Function ParseMailAddresses(ByVal input As String) As List(Of RfcMailAddress)
            If input Is Nothing Then Throw New ArgumentNullException("input")
			Dim returner As List(Of RfcMailAddress) = New List(Of RfcMailAddress)()
			'MailAddresses are split by commas
			Dim mailAddresses As IEnumerable(Of String) = Utility.SplitStringWithCharNotInsideQuotes(input, ","c)

			'Parse each of these
			For Each mailAddress As String In mailAddresses
                returner.Add(ParseMailAddress(mailAddress))
            Next

            Return returner
        End Function
    End Class
End Namespace
