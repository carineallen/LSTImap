Imports System
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports LSTImap.Common.Logging

Namespace Mime.Decode

    ''' <summary>
	''' Class used to decode RFC 2822 Date header fields.
	''' </summary>

    Public Class Rfc2822DateTime

        ''' <summary>
		''' Converts a string in RFC 2822 format into a <see cref="DateTime"/> object
		''' </summary>
		''' <param name="inputDate">The date to convert</param>
		''' <returns>
		''' A valid <see cref="DateTime"/> object, which represents the same time as the string that was converted. 
		''' If <paramref name="inputDate"/> Is Not a valid date representation, then <see cref="DateTime.MinValue"/> Is returned.
		''' </returns>
		''' <exception cref="ArgumentNullException">If <paramref name="inputDate"/> Is <see langword="null"/></exception>
		''' <exception cref="ArgumentException">If the <paramref name="inputDate"/> could Not be parsed into a <see cref="DateTime"/> object</exception>
        Shared Function StringToDate(ByVal inputDate As String) As DateTime
            If inputDate Is Nothing Then Throw New ArgumentNullException("inputDate")
            'Old date specification allows comments and a lot of whitespace
            inputDate = StripCommentsAndExcessWhitespace(inputDate)

            Try
                'Extract the DateTime
                Dim dateTime As DateTime = ExtractDateTime(inputDate)
                'Bail if we could not parse the date
                If dateTime = DateTime.MinValue Then Return dateTime
                'If a day-name is specified in the inputDate string, check if it fits with the date
                ValidateDayNameIfAny(dateTime, inputDate)
                'Convert the date into UTC
                dateTime = New DateTime(dateTime.Ticks, DateTimeKind.Utc)
                'Adjust according to the time zone
                dateTime = AdjustTimezone(dateTime, inputDate)
                'Return the parsed date
                Return dateTime
            Catch e As FormatException 'Convert.ToDateTime() Failure
                Throw New ArgumentException("Could not parse date: " & e.Message & ". Input was: """ & inputDate & """", e)
            Catch e As ArgumentException
                Throw New ArgumentException("Could not parse date: " & e.Message & ". Input was: """ & inputDate & """", e)
            End Try
        End Function

        ''' <summary>
		''' Adjust the <paramref name="dateTime"/> object given according to the timezone specified in the <paramref name="dateInput"/>.
		''' </summary>
		''' <param name="dateTime">The date to alter</param>
		''' <param name="dateInput">The input date, in which the timezone can be found</param>
		''' <returns>An date altered according to the timezone</returns>
		''' <exception cref="ArgumentException">If no timezone was found in <paramref name="dateInput"/></exception>
        Shared Function AdjustTimezone(ByVal dateTime As DateTime, ByVal dateInput As String) As DateTime
            'We know that the timezones are always in the last part of the date input
            Dim parts As String() = dateInput.Split(" ")
            Dim lastPart As String = parts(parts.Length - 1)
            'Convert timezones in older formats to [+-]dddd format.
            lastPart = Regex.Replace(lastPart, "UT|GMT|EST|EDT|CST|CDT|MST|MDT|PST|PDT|[A-I]|[K-Y]|Z", New MatchEvaluator(AddressOf MatchEvaluator))
            'Find the timezone specification
            ' Example Fri, 21 Nov 1997 09: 55:06 -0600
            ' finds -0600
            Dim match As Match = Regex.Match(lastPart, "[\+-](?<hours>\d\d)(?<minutes>\d\d)")

            If match.Success Then
                'We have found that the timezone is in +dddd or -dddd format
                ' Add the number of hours And minutes to our found date
                Dim hours As Integer = Integer.Parse(match.Groups("hours").Value)
                Dim minutes As Integer = Integer.Parse(match.Groups("minutes").Value)
                Dim factor As Integer = If(match.Value(0) = "+", -1, 1)
                dateTime = dateTime.AddHours(factor * hours)
                dateTime = dateTime.AddMinutes(factor * minutes)

                Return dateTime
            End If

            DefaultLogger.Log.LogDebug("No timezone found in date: " & dateInput & ". Using -0000 as default.")
            'A timezone of -0000 is the same as doing nothing
            Return dateTime
        End Function


        ''' <summary>
		''' Convert timezones in older formats to [+-]dddd format.
		''' </summary>
		''' <param name="match">The match that was found</param>
		''' <returns>The string to replace the matched string with</returns>
        Shared Function MatchEvaluator(ByVal match As Match) As String
            If Not match.Success Then
                Throw New ArgumentException("Match success are always true")
            End If

            Select Case match.Value
                '"A" through "I"
                'are equivalent to "+0100" through "+0900" respectively
                Case "A"
                    Return "+0100"
                Case "B"
                    Return "+0200"
                Case "C"
                    Return "+0300"
                Case "D"
                    Return "+0400"
                Case "E"
                    Return "+0500"
                Case "F"
                    Return "+0600"
                Case "G"
                    Return "+0700"
                Case "H"
                    Return "+0800"
                Case "I"
                    Return "+0900"
                '"K", "L", and "M"
                ' are equivalent to "+1000", "+1100", And "+1200" respectively
                Case "K"
                    Return "+1000"
                Case "L"
                    Return "+1100"
                Case "M"
                    Return "+1200"
                ' "N" through "Y"
                ' are equivalent to "-0100" through "-1200" respectively
                Case "N"
                    Return "-0100"
                Case "O"
                    Return "-0200"
                Case "P"
                    Return "-0300"
                Case "Q"
                    Return "-0400"
                Case "R"
                    Return "-0500"
                Case "S"
                    Return "-0600"
                Case "T"
                    Return "-0700"
                Case "U"
                    Return "-0800"
                Case "V"
                    Return "-0900"
                Case "W"
                    Return "-1000"
                Case "X"
                    Return "-1100"
                Case "Y"
                    Return "-1200"
                '"Z", "UT" and "GMT"
                ' Is equivalent to "+0000"
                Case "Z", "UT", "GMT"
                    Return "+0000"
                'US time zones
                Case "EDT"
                    Return "-0400"
                Case "EST"
                    Return "-0500"
                Case "CDT"
                    Return "-0500"
                Case "CST"
                    Return "-0600"
                Case "MDT"
                    Return "-0600"
                Case "MST"
                    Return "-0700"
                Case "PDT"
                    Return "-0700"
                Case "PST"
                    Return "-0800"
                Case Else
                    Throw New ArgumentException("Unexpected input")
            End Select
        End Function

        ''' <summary>
		''' Extracts the date And time parts from the <paramref name="dateInput"/>
		''' </summary>
		''' <param name="dateInput">The date input string, from which to extract the date And time parts</param>
		''' <returns>The extracted date part Or <see langword="DateTime.MinValue"/> if <paramref name="dateInput"/> Is Not recognized as a valid date.</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="dateInput"/> Is <see langword="null"/></exception>
        Shared Function ExtractDateTime(ByVal dateInput As String) As DateTime
            If dateInput Is Nothing Then Throw New ArgumentNullException("dateInput")
            'Matches the date and time part of a string
            ' Given string example Fri, 21 Nov 1997 09: 55:06 -0600
            ' Needs to find: 21 Nov 1997 09:55:06

            ' Seconds does Not need to be specified
            ' Even though it Is illigal, sometimes hours, minutes Or seconds are only specified with one digit

            ' Year with 2 Or 4 digits (1922 Or 22)
            Const year As String = "(\d\d\d\d|\d\d)"
            'Time with one or two digits for hour and minute and optinal seconds (06:04:06 or 6:4:6 or 06:04 or 6:4)
            Const time As String = "\d?\d:\d?\d(:\d?\d)?"
            'Correct format is 21 Nov 1997 09:55:06
            Const correctFormat As String = "\d\d? .+ " & year & " " & time
            'Some uses incorrect format: 2012-1-1 12:30
            Const incorrectFormat As String = year & "-\d?\d-\d?\d " & time
            'Some uses incorrect format: 08-May-2012 16:52:30 +0100
            Const correctFormatButWithDashes As String = "\d\d?-[A-Za-z]{3}-" & year & " " & time
            'We allow both correct and incorrect format
            Const joinedFormat As String = "(" & correctFormat & ")|(" & incorrectFormat & ")|(" & correctFormatButWithDashes & ")"
            Dim match As Match = Regex.Match(dateInput, joinedFormat)

            If match.Success Then
                Return Convert.ToDateTime(match.Value, CultureInfo.InvariantCulture)
            End If

            DefaultLogger.Log.LogError("The given date does not appear to be in a valid format: " & dateInput)
            Return DateTime.MinValue
        End Function

        ''' <summary>
		''' Validates that the given <paramref name="dateTime"/> agrees with a day-name specified
		''' in <paramref name="dateInput"/>.
		''' </summary>
		''' <param name="dateTime">The time to check</param>
		''' <param name="dateInput">The date input to extract the day-name from</param>
		''' <exception cref="ArgumentException">If <paramref name="dateTime"/> And <paramref name="dateInput"/> does Not agree on the day</exception>
        Shared Sub ValidateDayNameIfAny(ByVal dateTime As DateTime, ByVal dateInput As String)
            'Check if there is a day name in front of the date
            ' Example Fri, 21 Nov 1997 09: 55:06 -0600
            If dateInput.Length >= 4 AndAlso dateInput(3) = ","c Then
                Dim dayName As String = dateInput.Substring(0, 3)

                'If a dayName was specified. Check that the dateTime and the dayName
                ' agrees on which day it Is
                ' This Is just a failure-check And could be left out
                If (dateTime.DayOfWeek = DayOfWeek.Monday AndAlso Not dayName.Equals("Mon")) OrElse (dateTime.DayOfWeek = DayOfWeek.Tuesday AndAlso Not dayName.Equals("Tue")) OrElse (dateTime.DayOfWeek = DayOfWeek.Wednesday AndAlso Not dayName.Equals("Wed")) OrElse (dateTime.DayOfWeek = DayOfWeek.Thursday AndAlso Not dayName.Equals("Thu")) OrElse (dateTime.DayOfWeek = DayOfWeek.Friday AndAlso Not dayName.Equals("Fri")) OrElse (dateTime.DayOfWeek = DayOfWeek.Saturday AndAlso Not dayName.Equals("Sat")) OrElse (dateTime.DayOfWeek = DayOfWeek.Sunday AndAlso Not dayName.Equals("Sun")) Then
                    DefaultLogger.Log.LogDebug("Day-name does not correspond to the weekday of the date: " & dateInput)
                End If
            End If
            'If no day name was found no checks can be made
        End Sub

        ''' <summary>
		''' Strips And removes all comments And excessive whitespace from the string
		''' </summary>
		''' <param name="input">The input to strip from</param>
		''' <returns>The stripped string</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="input"/> Is <see langword="null"/></exception>
        Shared Function StripCommentsAndExcessWhitespace(ByVal input As String) As String
            If input Is Nothing Then Throw New ArgumentNullException("input")
            'Strip out comments
            ' Also strips out nested comments
            input = Regex.Replace(input, "(\((?>\((?<C>)|\)(?<-C>)|.?)*(?(C)(?!))\))", "")
            'Reduce any whitespace character to one space only
            input = Regex.Replace(input, "\s+", " ")
            'Remove all initial whitespace
            input = Regex.Replace(input, "^\s+", "")
            'Remove all ending whitespace
            input = Regex.Replace(input, "\s+$", "")
            'Remove spaces at colons
            ' Example 22: 33 : 44 => 22:33:44
            input = Regex.Replace(input, " ?: ?", ":")
            Return input
        End Function
    End Class
End Namespace
