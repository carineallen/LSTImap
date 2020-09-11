Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Mime.Decode

    ''' <summary>
	''' Used for decoding Quoted-Printable text.<br/>
	''' This Is a robust implementation of a Quoted-Printable decoder defined in <a href="http://tools.ietf.org/html/rfc2045">RFC 2045</a> And <a href="http://tools.ietf.org/html/rfc2047">RFC 2047</a>.<br/>
	''' Every measurement has been taken to conform to the RFC.
	''' </summary>
    Public Class QuotedPrintable

        ''' <summary>
		''' Decodes a Quoted-Printable string according to <a href="http://tools.ietf.org/html/rfc2047">RFC 2047</a>.<br/>
		''' RFC 2047 Is used for decoding Encoded-Word encoded strings.
		''' </summary>
		''' <param name="toDecode">Quoted-Printable encoded string</param>
		''' <param name="encoding">Specifies which encoding the returned string will be in</param>
		''' <returns>A decoded string in the correct encoding</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="toDecode"/> Or <paramref name="encoding"/> Is <see langword="null"/></exception>
        Shared Function DecodeEncodedWord(ByVal toDecode As String, ByVal encoding As Encoding) As String
            If toDecode Is Nothing Then Throw New ArgumentNullException("toDecode")
            If encoding Is Nothing Then Throw New ArgumentNullException("encoding")
            'Decode the QuotedPrintable string and return it
            Return encoding.GetString(Rfc2047QuotedPrintableDecode(toDecode, True))
        End Function

        ''' <summary>
        ''' Decodes a Quoted-Printable string according to <a href="http://tools.ietf.org/html/rfc2045">RFC 2045</a>.<br/>
        ''' RFC 2045 specifies the decoding of a body encoded with Content-Transfer-Encoding of quoted-printable.
        ''' </summary>
        ''' <param name="toDecode">Quoted-Printable encoded string</param>
        ''' <returns>A decoded byte array that the Quoted-Printable encoded string described</returns>
        ''' <exception cref="ArgumentNullException">If <paramref name="toDecode"/> Is <see langword="null"/></exception>
        Shared Function DecodeContentTransferEncoding(ByVal toDecode As String) As Byte()
            If toDecode Is Nothing Then Throw New ArgumentNullException("toDecode")
            ' Decode the QuotedPrintable string and return it
            Return Rfc2047QuotedPrintableDecode(toDecode, False)
        End Function


        ''' <summary>
		''' This Is the actual decoder.
		''' </summary>
		''' <param name="toDecode">The string to be decoded from Quoted-Printable</param>
		''' <param name="encodedWordVariant">
		''' If <see langword="true"/>, specifies that RFC 2047 quoted printable decoding Is used.<br/>
		''' This Is for quoted-printable encoded words<br/>
		''' <br/>
		''' If <see langword="false"/>, specifies that RFC 2045 quoted printable decoding Is used.<br/>
		''' This Is for quoted-printable Content-Transfer-Encoding
		''' </param>
		''' <returns>A decoded byte array that was described by <paramref name="toDecode"/></returns>
		''' <exception cref="ArgumentNullException">If <paramref name="toDecode"/> Is <see langword="null"/></exception>
		''' <remarks>See <a href="http://tools.ietf.org/html/rfc2047#section-4.2">RFC 2047 section 4.2</a> for RFC details</remarks>
        Private Shared Function Rfc2047QuotedPrintableDecode(ByVal toDecode As String, ByVal encodedWordVariant As Boolean) As Byte()

            If toDecode Is Nothing Then Throw New ArgumentNullException("toDecode")

            ' Create a byte array builder which is roughly equivalent to a StringBuilder
            Using byteArrayBuilder As MemoryStream = New MemoryStream()
                'Remove illegal control characters
                toDecode = RemoveIllegalControlCharacters(toDecode)

                'Run through the whole string that needs to be decoded
                For i As Integer = 0 To toDecode.Length - 1
                    Dim currentChar As Char = toDecode(i)

                    If currentChar = "=" Then
                        'Check that there is at least two characters behind the equal sign
                        If toDecode.Length - i < 3 Then
                            'We are at the end of the toDecode string, but something is missing. Handle it the way RFC 2045 states
                            WriteAllBytesToStream(byteArrayBuilder, DecodeEqualSignNotLongEnough(toDecode.Substring(i)))
                            'Since it was the last part, we should stop parsing anymore
                            Exit For
                        End If

                        'Decode the Quoted-Printable part
                        Dim quotedPrintablePart As String = toDecode.Substring(i, 3)
                        WriteAllBytesToStream(byteArrayBuilder, DecodeEqualSign(quotedPrintablePart))
                        'We now consumed two extra characters. Go forward two extra characters
                        i += 2
                    Else
                        'This character is not quoted printable hex encoded.

                        'Could it be the _ character, which represents space
                        ' And are we using the encoded word variant of QuotedPrintable
                        If currentChar = "_" AndAlso encodedWordVariant Then
                            'The RFC specifies that the "_" always represents hexadecimal 20 even if the
                            ' SPACE character occupies a different code position in the character set in use.
                            byteArrayBuilder.WriteByte(&H20)
                        Else
                            'This is not encoded at all. This is a literal which should just be included into the output.
                            byteArrayBuilder.WriteByte(Convert.ToByte(currentChar))
                        End If
                    End If
                Next

                Return byteArrayBuilder.ToArray()
            End Using
        End Function

        ''' <summary>
		''' Writes all bytes in a byte array to a stream
		''' </summary>
		''' <param name="stream">The stream to write to</param>
		''' <param name="toWrite">The bytes to write to the <paramref name="stream"/></param>
        Private Shared Sub WriteAllBytesToStream(ByVal stream As Stream, ByVal toWrite As Byte())
            stream.Write(toWrite, 0, toWrite.Length)
        End Sub

        ''' <summary>
		''' RFC 2045 states about robustness:<br/>
		''' <code>
		''' Control characters other than TAB, Or CR And LF as parts of CRLF pairs,
		''' must Not appear. The same Is true for octets with decimal values greater
		''' than 126.  If found in incoming quoted-printable data by a decoder, a
		''' robust implementation might exclude them from the decoded data And warn
		''' the user that illegal characters were discovered.
		''' </code>
		''' Control characters are defined in RFC 2396 as<br/>
		''' <c>control = US-ASCII coded characters 00-1F And 7F hexadecimal</c>
		''' </summary>
		''' <param name="input">String to be stripped from illegal control characters</param>
		''' <returns>A string with no illegal control characters</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="input"/> Is <see langword="null"/></exception>
        Private Shared Function RemoveIllegalControlCharacters(ByVal input As String) As String
            If input Is Nothing Then Throw New ArgumentNullException("input")
            'First we remove any \r or \n which is not part of a \r\n pair
            input = RemoveCarriageReturnAndNewLinewIfNotInPair(input)
            'Here only legal \r\n is left over
            ' We now simply keep them, And the \t which Is also allowed
            ' \x0A = \n
            ' \x0D = \r
            ' \x09 = \t)
            Return Regex.Replace(input, "[" & vbNullChar & "-" & vbBack & vbVerticalTab & vbFormFeed & ChrW(14) & "-" & ChrW(31) & ChrW(127) & "]", "")
        End Function

        ''' <summary>
		''' This method will remove any \r And \n which Is Not paired as \r\n
		''' </summary>
		''' <param name="input">String to remove lonely \r And \n's from</param>
		''' <returns>A string without lonely \r And \n's</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="input"/> Is <see langword="null"/></exception>
        Private Shared Function RemoveCarriageReturnAndNewLinewIfNotInPair(ByVal input As String) As String
            If input Is Nothing Then Throw New ArgumentNullException("input")
            'Use this for building up the new string. This is used for performance instead
            'of altering the input string each time a illegal token Is found
            Dim newString As StringBuilder = New StringBuilder(input.Length)

            For i As Integer = 0 To input.Length - 1

                'There is a character after it
                ' Check for lonely \r
                ' There Is a lonely \r if it Is the last character in the input Or if there
                ' Is no \n following it
                If input(i) = vbCr AndAlso (i + 1 >= input.Length OrElse input(i + 1) <> vbLf) Then
                    'Illegal token \r found. Do not add it to the new string

                    ' Check for lonely \n
                    ' There Is a lonely \n if \n Is the first character Or if there
                    ' Is no \r in front of it
                ElseIf input(i) = vbLf AndAlso (i - 1 < 0 OrElse input(i - 1) <> vbCr) Then
                    'Illegal token \n found. Do not add it to the new string
                Else
                    'No illegal tokens found. Simply insert the character we are at
                    ' in our New string
                    newString.Append(input(i))
                End If
            Next

            Return newString.ToString()
        End Function

        ''' <summary>
		''' RFC 2045 says that a robust implementation should handle:<br/>
		''' <code>
		''' An "=" cannot be the ultimate Or penultimate character in an encoded
		''' object. This could be handled as in case (2) above.
		''' </code>
		''' Case (2) Is:<br/>
		''' <code>
		''' An "=" followed by a character that Is neither a
		''' hexadecimal digit (including "abcdef") nor the CR character of a CRLF pair
		''' Is illegal.  This case can be the result of US-ASCII text having been
		''' included in a quoted-printable part of a message without itself having
		''' been subjected to quoted-printable encoding.  A reasonable approach by a
		''' robust implementation might be to include the "=" character And the
		''' following character in the decoded data without any transformation And, if
		''' possible, indicate to the user that proper decoding was Not possible at
		''' this point in the data.
		''' </code>
		''' </summary>
		''' <param name="decode">
		''' The string to decode which cannot have length above Or equal to 3
		''' And must start with an equal sign.
		''' </param>
		''' <returns>A decoded byte array</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="decode"/> Is <see langword="null"/></exception>
		''' <exception cref="ArgumentException">Thrown if a the <paramref name="decode"/> parameter has length above 2 Or does Not start with an equal sign.</exception>
        Private Shared Function DecodeEqualSignNotLongEnough(ByVal decode As String) As Byte()
            If decode Is Nothing Then Throw New ArgumentNullException("decode")
            'We can only decode wrong length equal signs
            If decode.Length >= 3 Then Throw New ArgumentException("decode must have length lower than 3", "decode")
            If decode.Length <= 0 Then Throw New ArgumentException("decode must have length lower at least 1", "decode")
            'First char must be =
            If decode(0) <> "=" Then Throw New ArgumentException("First part of decode must be an equal sign", "decode")
            'We will now believe that the string sent to us, was actually not encoded
            ' Therefore it must be in US-ASCII And we will return the bytes it corrosponds to
            Return Encoding.ASCII.GetBytes(decode)
        End Function

        ''' <summary>
		''' This helper method will decode a string of the form "=XX" where X Is any character.<br/>
		''' This method will never fail, unless an argument of length Not equal to three Is passed.
		''' </summary>
		''' <param name="decode">The length 3 character that needs to be decoded</param>
		''' <returns>A decoded byte array</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="decode"/> Is <see langword="null"/></exception>
		''' <exception cref="ArgumentException">Thrown if a the <paramref name="decode"/> parameter does Not have length 3 Or does Not start with an equal sign.</exception>
        Private Shared Function DecodeEqualSign(ByVal decode As String) As Byte()
            If decode Is Nothing Then Throw New ArgumentNullException("decode")
            'We can only decode the string if it has length 3 - other calls to this function is invalid
            If decode.Length <> 3 Then Throw New ArgumentException("decode must have length 3", "decode")
            'First char must be =
            If decode(0) <> "=" Then Throw New ArgumentException("decode must start with an equal sign", "decode")

            'There are two cases where an equal sign might appear
            ' It might be a
            '   - hex-string Like =3D, denoting the character with hex value 3D
            '   - it might be the last character on the line before a CRLF
            '     pair, denoting a soft linebreak, which simply
            '     splits the text up, because of the 76 chars per line restriction
            If decode.Contains(vbCrLf) Then
                'Soft break detected
                ' We want to return string.Empty which Is equivalent to a zero-length byte array
                Return New Byte(-1) {}
            End If

            'Hex string detected. Convertion needed.
            ' It might be that the string located after the equal sign Is Not hex characters
            ' An example:=JU
            ' In that case we would Like to catch the FormatException And do something else
            Try
                'The number part of the string is the last two digits. Here we simply remove the equal sign
                Dim numberString As String = decode.Substring(1)
                'Now we create a byte array with the converted number encoded in the string as a hex value (base 16)
                ' This will also handle illegal encodings Like =3d where the hex digits are Not uppercase,
                ' which Is a robustness requirement from RFC 2045.
                Dim oneByte As Byte() = {Convert.ToByte(numberString, 16)}
                'Simply return our one byte byte array
                Return oneByte
            Catch __unusedFormatException1__ As FormatException
                'RFC 2045 says about robust implementation:
                ' An "=" followed by a character that Is neither a
                ' hexadecimal digit (including "abcdef") nor the CR
                ' character of a CRLF pair Is illegal.  This case can be
                ' the result of US-ASCII text having been included in a
                ' quoted-printable part of a message without itself
                ' having been subjected to quoted-printable encoding.  A
                ' reasonable approach by a robust implementation might be
                ' to include the "=" character And the following
                ' character in the decoded data without any
                ' transformation And, if possible, indicate to the user
                ' that proper decoding was Not possible at this point in
                ' the data.

                ' So we choose to believe this Is actually an un-encoded string
                ' Therefore it must be in US-ASCII And we will return the bytes it corrosponds to
                Return Encoding.ASCII.GetBytes(decode)
            End Try
        End Function
    End Class
End Namespace

