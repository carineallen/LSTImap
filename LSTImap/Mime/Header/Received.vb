Imports System
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports LSTImap.Mime.Decode

Namespace Mime.Header

    ''' <summary>
	''' Class that hold information about one "Received:" header line.<br/>
	''' <br/>
	''' Visit these RFCs for more information:<br/>
	''' <see href="http://tools.ietf.org/html/rfc5321#section-4.4">RFC 5321 section 4.4</see><br/>
	''' <see href="http://tools.ietf.org/html/rfc4021#section-3.6.7">RFC 4021 section 3.6.7</see><br/>
	''' <see href="http://tools.ietf.org/html/rfc2822#section-3.6.7">RFC 2822 section 3.6.7</see><br/>
	''' <see href="http://tools.ietf.org/html/rfc2821#section-4.4">RFC 2821 section 4.4</see><br/>
	''' </summary>
    Public Class Received

        ''' <summary>
		''' The date of this received line.
		''' Is <see cref="DateTime.MinValue"/> if Not present in the received header line.
		''' </summary>
        Public Property Date1 As DateTime

        ''' <summary>
		''' A dictionary that contains the names And values of the
		''' received header line.<br/>
		''' If the received header Is invalid And contained one name
		''' multiple times, the first one Is used And the rest Is ignored.
		''' </summary>
		''' <example>
		''' If the header lines looks Like:
		''' <code>
		''' from sending.com (localMachine [127.0.0.1]) by test.net (Postfix)
		''' </code>
		''' then the dictionary will contain two keys: "from" And "by" with the values
		''' "sending.com (localMachine [127.0.0.1])" And "test.net (Postfix)".
		''' </example>
        Public Property Names As Dictionary(Of String, String)

        ''' <summary>
		''' The raw input string that was parsed into this class.
		''' </summary>
        Public Property Raw As String


        ''' <summary>
		''' Parses a Received header value.
		''' </summary>
		''' <param name="headerValue">The value for the header to be parsed</param>
		''' <exception cref="ArgumentNullException"><exception cref="ArgumentNullException">If <paramref name="headerValue"/> Is <see langword="null"/></exception></exception>
        Public Sub New(ByVal headerValue As String)
            If headerValue Is Nothing Then Throw New ArgumentNullException("headerValue")
            'Remember the raw input if someone whishes to use it
            Raw = headerValue
            'Default Date value
            Date1 = DateTime.MinValue

            'he date part is the last part of the string, and is preceeded by a semicolon
            ' Some emails forgets to specify the date, therefore we need to check if it Is there
            If headerValue.Contains(";") Then
                Dim datePart As String = headerValue.Substring(headerValue.LastIndexOf(";") + 1)
                Date1 = Rfc2822DateTime.StringToDate(datePart)
            End If

            Names = ParseDictionary(headerValue)
        End Sub


        ''' <summary>
		''' Parses the Received header name-value-list into a dictionary.
		''' </summary>
		''' <param name="headerValue">The full header value for the Received header</param>
		''' <returns>A dictionary where the name-value-list has been parsed into</returns>
        Private Shared Function ParseDictionary(ByVal headerValue As String) As Dictionary(Of String, String)
            Dim dictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)()
            'Remove the date part from the full headerValue if it is present
            Dim headerValueWithoutDate As String = headerValue

            If headerValue.Contains(";") Then
                headerValueWithoutDate = headerValue.Substring(0, headerValue.LastIndexOf(";"))
            End If

            'Reduce any whitespace character to one space only
            headerValueWithoutDate = Regex.Replace(headerValueWithoutDate, "\s+", " ")
            'The regex below should capture the following:
            ' The name consists of non-whitespace characters followed by a whitespace And then the value follows.
            ' There are multiple cases for the value part:
            '   1 Value Is just some characters Not including any whitespace
            '   2: Value Is some characters, a whitespace followed by an unlimited number of
            '      parenthesized values which can contain whitespaces, each delimited by whitespace
            '
            ' Cheat sheet for regex:
            ' \s means every whitespace character
            ' [^\s] means every character except whitespace characters
            ' +? Is a non-greedy equivalent of +
            Const pattern As String = "(?<name>[^\s]+)\s(?<value>[^\s]+(\s\(.+?\))*)"
            'Find each match in the string
            Dim matches As MatchCollection = Regex.Matches(headerValueWithoutDate, pattern)

            For Each match As Match In matches
                'Add the name and value part found in the matched result to the dictionary
                Dim name As String = match.Groups("name").Value
                Dim value As String = match.Groups("value").Value

                'Check if the name is really a comment.
                ' In this case, the first entry in the header value
                ' Is a comment
                If name.StartsWith("(") Then
                    Continue For
                End If

                'Only add the first name pair
                ' All subsequent pairs are ignored, as they are invalid anyway
                If Not dictionary.ContainsKey(name) Then dictionary.Add(name, value)
            Next

            Return dictionary
        End Function
    End Class
End Namespace
