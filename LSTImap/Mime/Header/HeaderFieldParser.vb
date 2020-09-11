Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Net.Mail
Imports System.Net.Mime
Imports LSTImap.Mime.Decode
Imports LSTImap.Common.Logging

Namespace Mime.Header

    ''' <summary>
	''' Class that can parse different fields in the header sections of a MIME message.
	''' </summary>
    Friend Module HeaderFieldParser

        ''' <summary>
		''' Parses the Content-Transfer-Encoding header.
		''' </summary>
		''' <param name="headerValue">The value for the header to be parsed</param>
		''' <returns>A <see cref="ContentTransferEncoding"/></returns>
		''' <exception cref="ArgumentNullException">If <paramref name="headerValue"/> Is <see langword="null"/></exception>
		''' <exception cref="ArgumentException">If the <paramref name="headerValue"/> could Not be parsed to a <see cref="ContentTransferEncoding"/></exception>
        Function ParseContentTransferEncoding(ByVal headerValue As String) As ContentTransferEncoding
            If headerValue Is Nothing Then Throw New ArgumentNullException("headerValue")

            Select Case headerValue.Trim().ToUpperInvariant()
                Case "7BIT"
                    Return ContentTransferEncoding.SevenBit
                Case "8BIT"
                    Return ContentTransferEncoding.EightBit
                Case "QUOTED-PRINTABLE"
                    Return ContentTransferEncoding.QuotedPrintable
                Case "BASE64"
                    Return ContentTransferEncoding.Base64
                Case "BINARY"
                    Return ContentTransferEncoding.Binary
                Case Else
                    'If a wrong argument Is passed To this parser method, Then we assume
                    ' default encoding, which Is SevenBit.
                    ' This Is to ensure that we do Not throw exceptions, even if the email Not MIME valid.
                    DefaultLogger.Log.LogDebug("Wrong ContentTransferEncoding was used. It was: " & headerValue)
                    Return ContentTransferEncoding.SevenBit
            End Select
        End Function


        ''' <summary>
        ''' Parses an ImportanceType from a given Importance header value.
        ''' </summary>
        ''' <param name="headerValue">The value to be parsed</param>
        ''' <returns>A <see cref="MailPriority"/>. If the <paramref name="headerValue"/> Is Not recognized, Normal Is returned.</returns>
        ''' <exception cref="ArgumentNullException">If <paramref name="headerValue"/> Is <see langword="null"/></exception>
        Function ParseImportance(ByVal headerValue As String) As MailPriority
            If headerValue Is Nothing Then Throw New ArgumentNullException("headerValue")

            Select Case headerValue.ToUpperInvariant()
                Case "5", "HIGH"
                    Return MailPriority.High
                Case "3", "NORMAL"
                    Return MailPriority.Normal
                Case "1", "LOW"
                    Return MailPriority.Low
                Case Else
                    DefaultLogger.Log.LogDebug("HeaderFieldParser: Unknown importance value: """ & headerValue & """. Using default of normal importance.")
                    Return MailPriority.Normal
            End Select
        End Function


        ''' <summary>
		''' Parses a the value for the header Content-Type to 
		''' a <see cref="ContentType"/> object.
		''' </summary>
		''' <param name="headerValue">The value to be parsed</param>
		''' <returns>A <see cref="ContentType"/> object</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="headerValue"/> Is <see langword="null"/></exception>
        Function ParseContentType(ByVal headerValue As String) As ContentType
            If headerValue Is Nothing Then Throw New ArgumentNullException("headerValue")
            'We create an empty Content-Type which we will fill in when we see the values
            Dim contentType As ContentType = New ContentType()
            'Now decode the parameters
            Dim parameters As List(Of KeyValuePair(Of String, String)) = Rfc2231Decoder.Decode(headerValue)

            For Each keyValuePair As KeyValuePair(Of String, String) In parameters
                Dim key As String = keyValuePair.Key.ToUpperInvariant().Trim()
                Dim value As String = Utility.RemoveQuotesIfAny(keyValuePair.Value.Trim())

                Select Case key
                    Case ""
                        'his is the MediaType - it has no key since it is the first one mentioned in the
                        ' headerValue And has no = in it.

                        ' Check for illegal content-type
                        If value.ToUpperInvariant().Equals("TEXT") Then value = "text/plain"
                        contentType.MediaType = value
                    Case "BOUNDARY"
                        contentType.Boundary = value
                    Case "CHARSET"
                        contentType.CharSet = value
                    Case "NAME"
                        contentType.Name = EncodedWord.Decode(value)
                    Case Else
                        'This is to shut up the code help that is saying that contentType.Parameters
                        ' can be null - which it cant!
                        If contentType.Parameters Is Nothing Then Throw New Exception("The ContentType parameters property is null. This will never be thrown.")
                        'We add the unknown value to our parameters list
                        ' "Known" unknown values are
                        ' - title
                        ' - report-type
                        contentType.Parameters.Add(key, value)
                End Select
            Next

            Return contentType
        End Function


        ''' <summary>
		''' Parses a the value for the header Content-Disposition to a <see cref="ContentDisposition"/> object.
		''' </summary>
		''' <param name="headerValue">The value to be parsed</param>
		''' <returns>A <see cref="ContentDisposition"/> object</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="headerValue"/> Is <see langword="null"/></exception>
        Function ParseContentDisposition(ByVal headerValue As String) As ContentDisposition
            If headerValue Is Nothing Then Throw New ArgumentNullException("headerValue")
            'See http://www.ietf.org/rfc/rfc2183.txt for RFC definition

            ' Create empty ContentDisposition - we will fill in details as we read them
            Dim contentDisposition As ContentDisposition = New ContentDisposition()
            'Now decode the parameters
            Dim parameters As List(Of KeyValuePair(Of String, String)) = Rfc2231Decoder.Decode(headerValue)

            For Each keyValuePair As KeyValuePair(Of String, String) In parameters
                Dim key As String = keyValuePair.Key.ToUpperInvariant().Trim()
                Dim value As String = Utility.RemoveQuotesIfAny(keyValuePair.Value.Trim())

                Select Case key
                    Case ""
                        'This is the DispisitionType - it has no key since it is the first one
                        ' And has no = in it.
                        contentDisposition.DispositionType = value
                    'The correct name of the parameter is filename, but some emails also contains the parameter
                    ' name, which also holds the name of the file. Therefore we use both names for the same field.
                    Case "NAME"

                    Case "FILENAME"
                        'The filename might be in qoutes, and it might be encoded-word encoded
                        contentDisposition.FileName = EncodedWord.Decode(value)
                    Case "CREATION-DATE"
                        'Notice that we need to create a new DateTime because of a failure in .NET 2.0.
                        ' The failure Is you cannot give contentDisposition a DateTime with a Kind of UTC
                        ' It will set the CreationDate correctly, but when trying to read it out it will throw an exception.
                        ' It Is the same with ModificationDate And ReadDate.
                        ' This Is fixed in 4.0 - maybe in 3.0 too.
                        ' Therefore we create a New DateTime which have a DateTimeKind set to unspecified
                        Dim creationDate As DateTime = New DateTime(Rfc2822DateTime.StringToDate(value).Ticks)
                        contentDisposition.CreationDate = creationDate
                    Case "MODIFICATION-DATE"
                        Dim midificationDate As DateTime = New DateTime(Rfc2822DateTime.StringToDate(value).Ticks)
                        contentDisposition.ModificationDate = midificationDate
                    Case "READ-DATE"
                        Dim readDate As DateTime = New DateTime(Rfc2822DateTime.StringToDate(value).Ticks)
                        contentDisposition.ReadDate = readDate
                    Case "SIZE"
                        contentDisposition.Size = SizeParser.Parse(value)
                    Case Else

                        If key.StartsWith("X-") Then
                            contentDisposition.Parameters.Add(key, value)
                            Exit Select
                        End If

                        Throw New ArgumentException("Unknown parameter in Content-Disposition. Ask developer to fix! Parameter: " & key)
                End Select
            Next

            Return contentDisposition
        End Function




        ''' <summary>
		''' Parses an ID Like Message-Id And Content-Id.<br/>
		''' Example:<br/>
		''' <c>&lt;test@test.com&gt;</c><br/>
		''' into<br/>
		''' <c>test@test.com</c>
		''' </summary>
		''' <param name="headerValue">The id to parse</param>
		''' <returns>A parsed ID</returns>
        Function ParseId(ByVal headerValue As String) As String
            'Remove whitespace in front and behind since
            ' whitespace Is allowed there
            ' Remove the last > And the first <
            Return headerValue.Trim().TrimEnd(">").TrimStart("<")
        End Function


        ''' <summary>
		''' Parses multiple IDs from a single string Like In-Reply-To.
		''' </summary>
		''' <param name="headerValue">The value to parse</param>
		''' <returns>A list of IDs</returns>
        Function ParseMultipleIDs(ByVal headerValue As String) As List(Of String)
            Dim returner As List(Of String) = New List(Of String)()
            'Split the string by >
            ' We cannot use ' ' (space) here since this is a possible value:
            ' <test@test.com><test2@test.com>
            Dim ids As String() = headerValue.Trim().Split({">"c}, StringSplitOptions.RemoveEmptyEntries)

            For Each id As String In ids
                returner.Add(ParseId(id))
            Next

            Return returner
        End Function
    End Module
End Namespace
