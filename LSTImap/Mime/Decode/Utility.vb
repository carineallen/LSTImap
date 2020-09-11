Imports System
Imports System.Collections.Generic

Namespace Mime.Decode

    ''' <summary>
	''' Contains common operations needed while decoding.
	''' </summary>
    Public Class Utility

        ''' <summary>
		''' Remove quotes, if found, around the string.
		''' </summary>
		''' <param name="text">Text with quotes Or without quotes</param>
		''' <returns>Text without quotes</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="text"/> Is <see langword="null"/></exception>
        Shared Function RemoveQuotesIfAny(ByVal text As String) As String
            If text Is Nothing Then Throw New ArgumentNullException("text")
            'Check if there are quotes at both ends and have at least two characters
            If text.Length > 1 AndAlso text(0) = ChrW(34) AndAlso text(text.Length - 1) = ChrW(34) Then
                'Remove quotes at both ends
                Return text.Substring(1, text.Length - 2)
            End If
            'If no quotes were found, the text is just returned
            Return text
        End Function


        ''' <summary>
        ''' Split a string into a list of strings using a specified character.<br/>
        ''' Everything inside quotes are ignored.
        ''' </summary>
        ''' <param name="input">A string to split</param>
        ''' <param name="toSplitAt">The character to use to split with</param>
        ''' <returns>A List of strings that was delimited by the <paramref name="toSplitAt"/> character</returns>
        Shared Function SplitStringWithCharNotInsideQuotes(ByVal input As String, ByVal toSplitAt As Char) As List(Of String)
            Dim elements As List(Of String) = New List(Of String)()
            Dim lastSplitLocation As Integer = 0
            Dim insideQuote As Boolean = False
            Dim characters As Char() = input.ToCharArray()

            For i As Integer = 0 To characters.Length - 1
                Dim character As Char = characters(i)
                If character = """"c Then insideQuote = Not insideQuote

                'Only split if we are not inside quotes
                If character = toSplitAt AndAlso Not insideQuote Then
                    ' We need to split
                    Dim length As Integer = i - lastSplitLocation
                    elements.Add(input.Substring(lastSplitLocation, length))
                    'Update last split location
                    ' + 1 so that we do Not include the character used to split with next time
                    lastSplitLocation = i + 1
                End If
            Next

            'Add the last part
            elements.Add(input.Substring(lastSplitLocation, input.Length - lastSplitLocation))
            Return elements
        End Function
    End Class
End Namespace
