Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions
Imports LSTImap.Common.Logging
Imports System.Runtime.InteropServices

Namespace Mime.Decode

    ''' <summary>
	''' This class Is responsible for decoding parameters that has been encoded with:<br/>
	''' <list type="bullet">
	''' <item>
	'''    <b>Continuation</b><br/>
	'''    This Is where a single parameter has such a long value that it could
	'''    be wrapped while in transit. Instead multiple parameters Is used on each line.<br/>
	'''    <br/>
	'''    <b>Example</b><br/>
	'''    From: <c> Content-Type: text/html; boundary="someVeryLongStringHereWhichCouldBeWrappedInTransit"</c><br/>
	'''    To: <c> Content-Type: text/html; boundary*0="someVeryLongStringHere" boundary*1="WhichCouldBeWrappedInTransit"</c><br/>
	''' </item>
	''' <item>
	'''    <b>Encoding</b><br/>
	'''    Sometimes other characters then ASCII characters are needed in parameters.<br/>
	'''    The parameter Is then given a different name to specify that it Is encoded.<br/>
	'''    <br/>
	'''    <b>Example</b><br/>
	'''    From: <c> Content-Disposition attachment; filename="specialCharsÆØÅ"</c><br/>
	'''    To: <c> Content-Disposition attachment; filename*="ISO-8859-1'en-us'specialCharsC6D8C0"</c><br/>
	'''    This encoding Is almost the same as <see cref="EncodedWord"/> encoding, And Is used to decode the value.<br/>
	''' </item>
	''' <item>
	'''   <b>Continuation And Encoding</b><br/>
	'''    Both Continuation And Encoding can be used on the same time.<br/>
	'''    <br/>
	'''    <b>Example</b><br/>
	'''    From: <c> Content-Disposition attachment; filename="specialCharsÆØÅWhichIsSoLong"</c><br/>
	'''    To: <c> Content-Disposition attachment; filename*0*="ISO-8859-1'en-us'specialCharsC6D8C0"; filename*1*="WhichIsSoLong"</c><br/>
	'''    This could also be encoded as:<br/>
	'''    To: <c> Content-Disposition attachment; filename*0*="ISO-8859-1'en-us'specialCharsC6D8C0"; filename*1="WhichIsSoLong"</c><br/>
	'''    Notice that <c>filename*1</c> does Not have an <c>*</c> after it - denoting it Is Not encoded.<br/>
	'''    There are some rules about this:<br/>
	'''    <list type="number">
	'''      <item>The encoding must be mentioned in the first part (filename*0*), which has to be encoded.</item>
	'''      <item>No other part must specify an encoding, but if encoded it uses the encoding mentioned in the first part.</item>
	'''      <item>Parts may be encoded Or Not in any order.</item>
	'''    </list>
	'''    <br/>
	''' </item>
	''' </list>
	''' More information And the specification Is available in <see href="http://tools.ietf.org/html/rfc2231">RFC 2231</see>.
	''' </summary>
    Public Class Rfc2231Decoder

        ''' <summary>
		''' Decodes a string of the form:<br/>
		''' <c>value0; key1=value1; key2=value2; key3=value3</c><br/>
		''' The returned List of key value pairs will have the key as key And the decoded value as value.<br/>
		''' The first value0 will have a key of <see cref="String.Empty"/>.<br/>
		''' <br/>
		''' If continuation Is used, then multiple keys will be merged into one key with the different values
		''' decoded into on big value for that key.<br/>
		''' Example:<br/>
		''' <code>
		''' title*0=part1
		''' title*1=part2
		''' </code>
		''' will have key And value of:<br></br>
		''' <c>title=decode(part1)decode(part2)</c>
		''' </summary>
		''' <param name="toDecode">The string to decode.</param>
		''' <returns>A list of decoded key value pairs.</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="toDecode"/> Is <see langword="null"/></exception>
        Shared Function Decode(ByVal toDecode As String) As List(Of KeyValuePair(Of String, String))
            If toDecode Is Nothing Then Throw New ArgumentNullException("toDecode")
            'Normalize the input to take account for missing semicolons after parameters.
            ' Example
            'text/plain; charset=\"iso-8859-1\" name=\"somefile.txt\" Or
            ' text/plain;\tcharset=\"iso-8859-1\"\tname=\"somefile.txt\"
            ' Is normalized to
            ' text/plain; charset=\"iso-8859-1\"; name=\"somefile.txt\"
            ' Only works for parameters inside quotes
            ' \s = matches whitespace
            toDecode = Regex.Replace(toDecode, "=\s*""(?<value>[^""]*)""\s", "=""${value}""; ")
            'Normalize 
            ' Since the above only works for parameters inside quotes, we need to normalize
            ' the special case with the first parameter.
            ' Example:
            ' attachment filename="foo"
            ' Is normalized to
            ' attachment; filename="foo"
            ' ^= matches start of line (when Not inside square bracets [])
            toDecode = Regex.Replace(toDecode, "^(?<first>[^;\s]+)\s(?<second>[^;\s]+)", "${first}; ${second}")
            'Split by semicolon, but only if not inside quotes
            Dim splitted As List(Of String) = Utility.SplitStringWithCharNotInsideQuotes(toDecode.Trim(), ";"c)
            Dim collection As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))(splitted.Count)

            For Each part As String In splitted
                'Empty strings should not be processed
                If part.Trim().Length = 0 Then Continue For
                Dim keyValue As String() = part.Trim().Split({"="c}, 2)

                If keyValue.Length = 1 Then
                    collection.Add(New KeyValuePair(Of String, String)("", keyValue(0)))
                ElseIf keyValue.Length = 2 Then
                    collection.Add(New KeyValuePair(Of String, String)(keyValue(0), keyValue(1)))
                Else
                    Throw New ArgumentException("When splitting the part """ & part & """ by = there was " & keyValue.Length & " parts. Only 1 and 2 are supported")
                End If
            Next

            Return DecodePairs(collection)
        End Function

        ''' <summary>
		''' Decodes the list of key value pairs into a decoded list of key value pairs.<br/>
		''' There may be less keys in the decoded list, but then the values for the lost keys will have been appended
		''' to the New key.
		''' </summary>
		''' <param name="pairs">The pairs to decode</param>
		''' <returns>A decoded list of pairs</returns>
        Shared Function DecodePairs(ByVal pairs As List(Of KeyValuePair(Of String, String))) As List(Of KeyValuePair(Of String, String))
            If pairs Is Nothing Then Throw New ArgumentNullException("pairs")
            Dim resultPairs As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))(pairs.Count)
            Dim pairsCount As Integer = pairs.Count

            For i As Integer = 0 To pairsCount - 1
                Dim currentPair As KeyValuePair(Of String, String) = pairs(i)
                Dim key As String = currentPair.Key
                Dim value As String = Utility.RemoveQuotesIfAny(currentPair.Value)

                'Is it a continuation parameter? (encoded or not)
                If key.EndsWith("*0", StringComparison.OrdinalIgnoreCase) OrElse key.EndsWith("*0*", StringComparison.OrdinalIgnoreCase) Then
                    'This encoding will not be used if we get into the if which tells us
                    ' that the whole continuation Is Not encoded
                    Dim encoding As String = "notEncoded - Value here is never used"

                    'Now lets find out if it is encoded too.
                    If key.EndsWith("*0*", StringComparison.OrdinalIgnoreCase) Then
                        'It is encoded.

                        'Fetch out the encoding for later use And decode the value
                        ' If the value was Not encoded as the email specified
                        ' encoding will be set to null. This will be used later.
                        value = DecodeSingleValue(value, encoding)
                        'Find the right key to use to store the full value
                        ' Remove the start *0 which tells Is it Is a continuation, And the first one
                        ' And remove the * afterwards which tells us it Is encoded
                        key = key.Replace("*0*", "")
                    Else
                        'It is not encoded, and no parts of the continuation is encoded either

                        ' Find the right key to use to store the full value
                        ' Remove the start *0 which tells Is it Is a continuation, And the first one
                        key = key.Replace("*0", "")
                    End If

                    'The StringBuilder will hold the full decoded value from all continuation parts
                    Dim builder As StringBuilder = New StringBuilder()
                    'Append the decoded value
                    builder.Append(value)
                    'Now go trough the next keys to see if they are part of the continuation
                    Dim j As Integer = i + 1, continuationCount As Integer = 1

                    While j < pairsCount
                        Dim jKey As String = pairs(j).Key
                        Dim valueJKey As String = Utility.RemoveQuotesIfAny(pairs(j).Value)

                        If jKey.Equals(key & "*" & continuationCount) Then
                            'This value part of the continuation is not encoded
                            ' Therefore remove qoutes if any And add to our stringbuilder
                            builder.Append(valueJKey)
                            'Remember to increment i, as we have now treated one more KeyValuePair
                            i += 1
                        ElseIf jKey.Equals(key & "*" & continuationCount & "*") Then

                            'We will not get into this part if the first part was not encoded
                            ' Therefore the encoding will only be used if And only if the
                            ' first part was encoded, in which case we have remembered the encoding used

                            ' Sometimes an email creator says that a string was encoded, but it really
                            ' was Not. This Is to catch that problem.
                            If encoding IsNot Nothing Then
                                'This value part of the continuation is encoded
                                ' the encoding Is Not given in the current value,
                                ' but was given in the first continuation, which we remembered for use here
                                valueJKey = DecodeSingleValue2(valueJKey, encoding)
                            End If

                            builder.Append(valueJKey)
                            'Remember to increment i, as we have now treated one more KeyValuePair
                            i += 1
                        Else
                            'No more keys for this continuation
                            Exit For
                        End If

                        j += 1
                        continuationCount += 1
                    End While

                    'Add the key and the full value as a pair
                    value = builder.ToString()
                    resultPairs.Add(New KeyValuePair(Of String, String)(key, value))
                ElseIf key.EndsWith("*", StringComparison.OrdinalIgnoreCase) Then
                    'This parameter is only encoded - it is not part of a continuation
                    ' We need to change the key from "<key>*" to "<key>" And decode the value

                    ' To get the key we want, we remove the last * that denotes
                    ' that the value hold by the key was encoded
                    key = key.Replace("*", "")
                    'Decode the value
                    Dim throwAway As String
                    value = DecodeSingleValue(value, throwAway)
                    'Now input the new value with the new key
                    resultPairs.Add(New KeyValuePair(Of String, String)(key, value))
                Else
                    'Fully normal key - the value is not encoded
                    ' Therefore nothing to do, And we can simply pass the pair
                    ' as being decoded now
                    resultPairs.Add(currentPair)
                End If
            Next

            Return resultPairs
        End Function


        ''' <summary>
		''' This will decode a single value of the form: <c> ISO-8859-1'en-us'%3D%3DIamHere</c><br/>
		''' Which Is basically a <see cref="EncodedWord"/> form just using % instead of =<br/>
		''' Notice that 'en-us' part is not used for anything.<br/>
		''' <br/>
		''' If the single value given Is Not on the correct form, it will be returned without 
		''' being decoded And <paramref name="encodingUsed"/> will be set to <see langword="null"/>.
		''' </summary>
		''' <param name="encodingUsed">
		''' The encoding used to decode with - it Is given back for later use.<br/>
		''' <see langword="null"/> if input was Not in the correct form.
		''' </param>
		''' <param name="toDecode">The value to decode</param>
		''' <returns>
		''' The decoded value that corresponds to <paramref name="toDecode"/> Or if
		''' <paramref name="toDecode"/> Is Not on the correct form, it will be non-decoded.
		''' </returns>
		''' <exception cref="ArgumentNullException">If <paramref name="toDecode"/> Is <see langword="null"/></exception>
        Private Shared Function DecodeSingleValue(ByVal toDecode As String, <Out> ByRef encodingUsed As String) As String
            If toDecode Is Nothing Then Throw New ArgumentNullException("toDecode")
            'Check if input has a part describing the encoding
            If toDecode.IndexOf("'") = -1 Then
                'The input was not encoded (at least not valid) and it is returned as is
                DefaultLogger.Log.LogDebug("Rfc2231Decoder: Someone asked me to decode a string which was not encoded - returning raw string. Input: " & toDecode)
                encodingUsed = Nothing
                Return toDecode
            End If

            encodingUsed = toDecode.Substring(0, toDecode.IndexOf("'"c))
            toDecode = toDecode.Substring(toDecode.LastIndexOf("'"c) + 1)
            Return DecodeSingleValue2(toDecode, encodingUsed)
        End Function


        ''' <summary>
		''' This will decode a single value of the form: %3D%3DIamHere
		''' Which Is basically a <see cref="EncodedWord"/> form just using % instead of =
		''' </summary>
		''' <param name="valueToDecode">The value to decode</param>
		''' <param name="encoding">The encoding used to decode with</param>
		''' <returns>The decoded value that corresponds to <paramref name="valueToDecode"/></returns>
		''' <exception cref="ArgumentNullException">If <paramref name="valueToDecode"/> Is <see langword="null"/></exception>
		''' <exception cref="ArgumentNullException">If <paramref name="encoding"/> Is <see langword="null"/></exception>
        Shared Function DecodeSingleValue2(ByVal valueToDecode As String, ByVal encoding As String) As String
            If valueToDecode Is Nothing Then Throw New ArgumentNullException("valueToDecode")
            If encoding Is Nothing Then Throw New ArgumentNullException("encoding")
            'The encoding used is the same as QuotedPrintable, we only
            ' need to change % to =
            ' And otherwise make it look Like the correct EncodedWord encoding
            valueToDecode = "=?" & encoding & "?Q?" & valueToDecode.Replace("%", "=") & "?="
            Return EncodedWord.Decode(valueToDecode)
        End Function
    End Class
End Namespace

