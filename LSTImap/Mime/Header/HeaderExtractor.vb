Imports System
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.IO
Imports System.Text
Imports LSTImap.Common
Imports System.Runtime.InteropServices

Namespace Mime.Header

    '''<summary>
	''' Utility class that divides a message into a body And a header.<br/>
	''' The header Is then parsed to a strongly typed <see cref="MessageHeader"/> object.
	'''</summary>
    Public Class HeaderExtractor

        ''' <summary>
		''' Find the end of the header section in a byte array.<br/>
		''' The headers have ended when a blank line Is found
		''' </summary>
		''' <param name="messageContent">The full message stored as a byte array</param>
		''' <returns>The position of the line just after the header end blank line</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="messageContent"/> Is <see langword="null"/></exception>
        Shared Function FindHeaderEndPosition(ByVal messageContent As Byte()) As Integer
            If messageContent Is Nothing Then Throw New ArgumentNullException("messageContent")

            'Convert the byte array into a stream
            Using stream As Stream = New MemoryStream(messageContent)

                While True
                    'Read a line from the stream. We know headers are in US-ASCII
                    ' therefore it Is Not problem to read them as such
                    Dim line As String = StreamUtility.ReadLineAsAscii(stream)
                    'The end of headers is signaled when a blank line is found
                    ' Or if the line Is null - in which case the email Is actually an email with
                    ' only headers but no body
                    If String.IsNullOrEmpty(line) Then Return CInt(stream.Position)
                End While
            End Using
        End Function


        ''' <summary>
        ''' Extract the header part And body part of a message.<br/>
        ''' The headers are then parsed to a strongly typed <see cref="MessageHeader"/> object.
        ''' </summary>
        ''' <param name="fullRawMessage">The full message in bytes where header And body needs to be extracted from</param>
        ''' <param name="headers">The extracted header parts of the message</param>
        ''' <param name="body">The body part of the message</param>
        ''' <exception cref="ArgumentNullException">If <paramref name="fullRawMessage"/> Is <see langword="null"/></exception>
        Shared Sub ExtractHeadersAndBody(ByVal fullRawMessage As Byte(), <Out> ByRef headers As MessageHeader, <Out> ByRef body As Byte())
            If fullRawMessage Is Nothing Then Throw New ArgumentNullException("fullRawMessage")
            'Find the end location of the headers
            Dim endOfHeaderLocation As Integer = FindHeaderEndPosition(fullRawMessage)
            'The headers are always in ASCII - therefore we can convert the header part into a string
            ' using US-ASCII encoding
            Dim headersString As String = Encoding.ASCII.GetString(fullRawMessage, 0, endOfHeaderLocation)
            'Now parse the headers to a NameValueCollection
            Dim headersUnparsedCollection As NameValueCollection = ExtractHeaders(headersString)
            'Use the NameValueCollection to parse it into a strongly-typed MessageHeader header
            headers = New MessageHeader(headersUnparsedCollection)
            'Since we know where the headers end, we also k
            'Copy the body part into the body parameter
            body = New Byte(fullRawMessage.Length - endOfHeaderLocation - 1) {}
            Array.Copy(fullRawMessage, endOfHeaderLocation, body, 0, body.Length)
        End Sub


        ''' <summary>
		''' Method that takes a full message And extract the headers from it.
		''' </summary>
		''' <param name="messageContent">The message to extract headers from. Does Not need the body part. Needs the empty headers end line.</param>
		''' <returns>A collection of Name And Value pairs of headers</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="messageContent"/> Is <see langword="null"/></exception>
        Shared Function ExtractHeaders(ByVal messageContent As String) As NameValueCollection
            If messageContent Is Nothing Then Throw New ArgumentNullException("messageContent")
            Dim headers As NameValueCollection = New NameValueCollection()

            Using messageReader As StringReader = New StringReader(messageContent)
                'Read until all headers have ended.
                ' The headers ends when an empty line Is encountered
                ' An empty message might actually Not have an empty line, in which
                ' case the headers end with null value.
                Dim line As String = messageReader.ReadLine()

                While Not String.IsNullOrEmpty(line)
                    'Split into name and value
                    Dim header As KeyValuePair(Of String, String) = SeparateHeaderNameAndValue(line)
                    'First index is header name
                    Dim headerName As String = header.Key
                    'Second index is the header value.
                    ' Use a StringBuilder since the header value may be continued on the next line
                    Dim headerValue As StringBuilder = New StringBuilder(header.Value)

                    'Keep reading until we would hit next header
                    ' This if for handling multi line headers
                    While IsMoreLinesInHeaderValue(messageReader)
                        'Unfolding is accomplished by simply removing any CRLF
                        ' that Is immediately followed by WSP
                        ' This was done using ReadLine (it discards CRLF)
                        ' See http://tools.ietf.org/html/rfc822#section-3.1.1 for more information
                        Dim moreHeaderValue As String = messageReader.ReadLine()
                        'If this exception is ever raised, there is an serious algorithm failure
                        'IsMoreLinesInHeaderValue does Not return true if the next line does Not exist
                        ' This check Is only included to stop the nagging "possibly null" code analysis hint
                        If moreHeaderValue Is Nothing Then Throw New ArgumentException("This will never happen")
                        'Simply append the line just read to the header value
                        headerValue.Append(moreHeaderValue)
                    End While

                    'Now we have the name and full value. Add it
                    headers.Add(headerName, headerValue.ToString())
                    line = messageReader.ReadLine()
                End While
            End Using

            Return headers
        End Function


        ''' <summary>
		''' Check if the next line Is part of the current header value we are parsing by
		''' peeking on the next character of the <see cref="TextReader"/>.<br/>
		''' This should only be called while parsing headers.
		''' </summary>
		''' <param name="reader">The reader from which the header Is read from</param>
		''' <returns><see langword="true"/> if multi-line header. <see langword="false"/> otherwise</returns>
        Shared Function IsMoreLinesInHeaderValue(ByVal reader As TextReader) As Boolean
            Dim peek As Integer = reader.Peek()
            If peek = -1 Then Return False
            Dim peekChar As Char = ChrW(peek)
            'A multi line header must have a whitespace character
            ' on the next line if it Is to be continued
            Return peekChar = " " OrElse peekChar = vbTab
        End Function


        ''' <summary>
		''' Separate a full header line into a header name And a header value.
		''' </summary>
		''' <param name="rawHeader">The raw header line to be separated</param>
		''' <exception cref="ArgumentNullException">If <paramref name="rawHeader"/> Is <see langword="null"/></exception>
        Shared Function SeparateHeaderNameAndValue(ByVal rawHeader As String) As KeyValuePair(Of String, String)
            If rawHeader Is Nothing Then Throw New ArgumentNullException("rawHeader")
            Dim key As String = String.Empty
            Dim value As String = String.Empty
            Dim indexOfColon As Integer = rawHeader.IndexOf(":")

            'Check if it is allowed to make substring calls
            If indexOfColon >= 0 AndAlso rawHeader.Length >= indexOfColon + 1 Then
                key = rawHeader.Substring(0, indexOfColon).Trim()
                value = rawHeader.Substring(indexOfColon + 1).Trim()
            End If

            Return New KeyValuePair(Of String, String)(key, value)
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
