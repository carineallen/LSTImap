Imports System
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Mime.Decode

    ''' <summary>
    ''' Utility class for dealing with encoded word strings<br/>
    ''' <br/>
    ''' EncodedWord encoded strings are only in ASCII, but can embed information
    ''' about characters in other character sets.<br/>
    ''' <br/>
    ''' It Is done by specifying the character set, an encoding that maps from ASCII to
    ''' the correct bytes And the actual encoded string.<br/>
    ''' <br/>
    ''' It Is specified in a format that Is best summarized by a BNF:<br/>
    ''' <c>"=?" character_set "?" encoding "?" encoded-text "?="</c><br/>
    ''' </summary>
    ''' <example>
    ''' <c>=?ISO-8859-1?Q?=2D?=</c>
    ''' Here <c>ISO-8859-1</c> Is the character set.<br/>
    ''' <c>Q</c> Is the encoding method (quoted-printable). <c>B</c> Is also supported (Base 64).<br/>
    ''' The encoded text Is the <c>=2D</c> part which Is decoded to a space.
    ''' </example>

    Public Class EncodedWord

        ''' <summary>
		''' Decode text that Is encoded with the <see cref="EncodedWord"/> encoding.<br/>
		'''<br/>
		''' This method will decode any encoded-word found in the string.<br/>
		''' All parts which Is Not encoded will Not be touched.<br/>
		''' <br/>
		''' From <a href="http://tools.ietf.org/html/rfc2047">RFC 2047</a>:<br/>
		''' <code>
		''' Generally, an "encoded-word" Is a sequence of printable ASCII
		''' characters that begins with "=?", ends with "?=", And has two "?"s in
		''' between.  It specifies a character set And an encoding method, And
		''' also includes the original text encoded as graphic ASCII characters,
		''' according to the rules for that encoding method.
		''' </code>
		''' Example:<br/>
		''' <c>=?ISO-8859-1?q?this=20is=20some=20text?= other text here</c>
		''' </summary>
		''' <remarks>See <a href="http://tools.ietf.org/html/rfc2047#section-2">RFC 2047 section 2</a> "Syntax of encoded-words" for more details</remarks>
		''' <param name="encodedWords">Source text. May be content which Is Not encoded.</param>
		''' <returns>Decoded text</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="encodedWords"/> Is <see langword="null"/></exception>

        Shared Function Decode(ByVal encodedWords As String) As String
            If encodedWords Is Nothing Then Throw New ArgumentNullException("encodedWords")

            'Notice that RFC2231 redefines the BNF to encoded-word := "=?" charset ["*" language] "?" encoded-text "?="
            'but no usage of this BNF have been spotted yet. It Is here to ease debugging if such a case Is discovered.

            'This Is the regex that should fit the BNF
            'RFC Says that NO WHITESPACE Is allowed in this encoding, but there are examples where whitespace Is there, And therefore this regex allows for such.

            Const encodedWordRegex As String = "\=\?(?<Charset>\S+?)\?(?<Encoding>\w)\?(?<Content>.+?)\?\="

            '\w	Matches any word character including underscore. Equivalent to "[A-Za-z0-9_]".
            '\S	Matches any nonwhite space character. Equivalent to "[^ \f\n\r\t\v]".
            '+?   non-greedy equivalent to +
            '(?<NAME>REGEX) Is a named group with name NAME And regular expression REGEX

            'Any amount of linear-space-white between 'encoded-word's, even if it includes a CRLF followed by one Or more SPACEs,
            'Is ignored for the purposes of display.
            'http://tools.ietf.org/html/rfc2047#page-12
            'Define a regular expression that captures two encoded words with some whitespace between them

            Const replaceRegex As String = "(?<first>" & encodedWordRegex & ")\s+(?<second>" & encodedWordRegex & ")"

            'Then, find an occurrence of such an expression, but remove the whitespace in between when found
            'Need to be done twice for encodings such as "=?UTF-8?Q?a?= =?UTF-8?Q?b?= =?UTF-8?Q?c?=" to be replaced correctly

            encodedWords = Regex.Replace(encodedWords, replaceRegex, "${first}${second}")
            encodedWords = Regex.Replace(encodedWords, replaceRegex, "${first}${second}")
            Dim decodedWords As String = encodedWords
            Dim matches As MatchCollection = Regex.Matches(encodedWords, encodedWordRegex)

            For Each match As Match In matches
                'If this match was not a success, we should not use it

                If Not match.Success Then Continue For
                Dim fullMatchValue As String = match.Value
                Dim encodedText As String = match.Groups("Content").Value
                Dim encoding As String = match.Groups("Encoding").Value
                Dim charset As String = match.Groups("Charset").Value

                'Get the encoding which corrosponds to the character set

                Dim charsetEncoding As Encoding = EncodingFinder.FindEncoding(charset)

                'Store decoded text here when done
                Dim decodedText As String

                Select Case encoding.ToUpperInvariant()

                        'RFC: The "B" encoding Is identical to the "BASE64" encoding defined by RFC 2045.
                        'http://tools.ietf.org/html/rfc2045#section-6.8
                    Case "B"
                        decodedText = Base64.Decode(encodedText, charsetEncoding)

RFC:
                        'The "Q" encoding Is similar to the "Quoted-Printable" content-transfer-encoding defined in RFC 2045.
                        'There are more details to this. Please check
                        'http://tools.ietf.org/html/rfc2047#section-4.2
                    Case "Q"
                        decodedText = QuotedPrintable.DecodeEncodedWord(encodedText, charsetEncoding)
                    Case Else
                        Throw New ArgumentException("The encoding " & encoding & " was not recognized")
                End Select

                'Repalce our encoded value with our decoded value
                decodedWords = decodedWords.Replace(fullMatchValue, decodedText)
            Next

            Return decodedWords
        End Function
    End Class
End Namespace
