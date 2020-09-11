Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text

Namespace Mime.Decode

    ''' <summary>
	''' Utility class used by OpenPop for mapping from a characterSet to an <see cref="Encoding"/>.<br/>
	''' <br/>
	''' The functionality of the class can be altered by adding mappings
	''' using <see cref="AddMapping"/> And by adding a <see cref="FallbackDecoder"/>.<br/>
	''' <br/>
	''' Given a characterSet, it will try to find the Encoding as follows:
	''' <list type="number">
	'''     <item>
	'''         <description>If a mapping for the characterSet was added, use the specified Encoding from there. Mappings can be added using <see cref="AddMapping"/>.</description>
	'''     </item>
	'''     <item>
	'''         <description>Try to parse the characterSet And look it up using <see cref="Encoding.GetEncoding(int)"/> for codepages Or <see cref="Encoding.GetEncoding(String)"/> for named encodings.</description>
	'''    </item>
	'''     <item>
	'''         <description>If an encoding Is Not found yet, use the <see cref="FallbackDecoder"/> if defined. The <see cref="FallbackDecoder"/> Is user defined.</description>
	'''     </item>
	''' </list>
	''' </summary>


    Public Class EncodingFinder

        ''' <summary>
		''' Delegate that Is used when the EncodingFinder Is unable to find an encoding by
		''' using the <see cref="EncodingFinder.EncodingMap"/> Or general code.<br/>
		''' This Is used as a last resort And can be used for setting a default encoding Or
		''' for finding an encoding on runtime for some <paramref name="characterSet"/>.
		''' </summary>
		''' <param name="characterSet">The character set to find an encoding for.</param>
		''' <returns>An encoding for the <paramref name="characterSet"/> Or <see langword="null"/> if none could be found.</returns>
        Public Delegate Function FallbackDecoderDelegate(ByVal characterSet As String) As Encoding


        ''' <summary>
		''' Last resort decoder.
		''' </summary>
        Public Shared Property FallbackDecoder As FallbackDecoderDelegate

        ''' <summary>
		''' Mapping from charactersets to encodings.
		''' </summary>
        Private Shared Property EncodingMap As Dictionary(Of String, Encoding)

        ''' <summary>
		''' Initialize the EncodingFinder
		''' </summary>
        Shared Sub New()
            Reset()
        End Sub

        ''' <summary>
		''' Used to reset this static class to facilite isolated unit testing.
		''' </summary>
        Friend Shared Sub Reset()
            EncodingMap = New Dictionary(Of String, Encoding)()
            FallbackDecoder = Nothing
            'Some emails incorrectly specify the encoding as utf8, but it should have been utf-8.
            AddMapping("utf8", Encoding.UTF8)
        End Sub

        '''' <summary>
        ''' Parses a character set into an encoding.
        ''' </summary>
        ''' <param name="characterSet">The character set to parse</param>
        ''' <returns>An encoding which corresponds to the character set</returns>
        ''' <exception cref="ArgumentNullException">If <paramref name="characterSet"/> Is <see langword="null"/></exception>
        Shared Function FindEncoding(ByVal characterSet As String) As Encoding
            If characterSet Is Nothing Then Throw New ArgumentNullException("characterSet")
            Dim charSetUpper As String = characterSet.ToUpperInvariant()
            'Check if the characterSet is explicitly mapped to an encoding
            If EncodingMap.ContainsKey(charSetUpper) Then Return EncodingMap(charSetUpper)
            'Try to generally find the encoding
            Try

                If charSetUpper.Contains("WINDOWS") OrElse charSetUpper.Contains("CP") Then
                    'It seems the characterSet contains an codepage value, which we should use to parse the encoding
                    charSetUpper = charSetUpper.Replace("CP", "") 'Remove cp
                    charSetUpper = charSetUpper.Replace("WINDOWS", "") 'Remove windows
                    charSetUpper = charSetUpper.Replace("-", "") 'Remove - which could be used as cp-1554
                    'Now we hope the only thing left in the characterSet is numbers.
                    Dim codepageNumber As Integer = Integer.Parse(charSetUpper, CultureInfo.InvariantCulture)

                    Return Encoding.GetEncoding(codepageNumber)
                End If
                'It seems there is no codepage value in the characterSet. It must be a named encoding
                Return Encoding.GetEncoding(characterSet)
            Catch __unusedArgumentException1__ As ArgumentException
                'The encoding could not be found generally. 
                ' Try to use the FallbackDecoder if it Is defined.

                ' Check if it Is defined
                If FallbackDecoder Is Nothing Then Throw 'It was not defined - throw catched exception
                'Use the FallbackDecoder
                Dim fallbackDecoderResult As Encoding = FallbackDecoder(characterSet)
                'Check if the FallbackDecoder had a solution
                If fallbackDecoderResult IsNot Nothing Then Return fallbackDecoderResult
                'If no solution was found, throw catched exception
                Throw
            End Try
        End Function

        ''' <summary>
		''' Puts a mapping from <paramref name="characterSet"/> to <paramref name="encoding"/>
		''' into the <see cref="EncodingFinder"/>'s internal mapping Dictionary.
		''' </summary>
		''' <param name="characterSet">The string that maps to the <paramref name="encoding"/></param>
		''' <param name="encoding">The <see cref="Encoding"/> that should be mapped from <paramref name="characterSet"/></param>
		''' <exception cref="ArgumentNullException">If <paramref name="characterSet"/> Is <see langword="null"/></exception>
		''' <exception cref="ArgumentNullException">If <paramref name="encoding"/> Is <see langword="null"/></exception>
        Shared Sub AddMapping(ByVal characterSet As String, ByVal encoding As Encoding)
            If characterSet Is Nothing Then Throw New ArgumentNullException("characterSet")
            If encoding Is Nothing Then Throw New ArgumentNullException("encoding")
            'Add the mapping using uppercase
            EncodingMap.Add(characterSet.ToUpperInvariant(), encoding)
        End Sub
    End Class
End Namespace

