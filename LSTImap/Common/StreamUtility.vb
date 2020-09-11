Imports System
Imports System.IO
Imports System.Text

Namespace Common
    ''' <summary>
	''' Utility to help reading bytes And strings of a <see cref="Stream"/>
	''' </summary>
    Public Class StreamUtility

        ''' <summary>
        ''' Read a line from the stream.
        ''' A line Is interpreted as all the bytes read until a CRLF Or LF Is encountered.<br/>
        ''' CRLF pair Or LF Is Not included in the string.
        ''' </summary>
        ''' <param name="stream">The stream from which the line Is to be read</param>
        ''' <returns>A line read from the stream returned as a byte array Or <see langword="null"/> if no bytes were readable from the stream</returns>
        ''' <exception cref="ArgumentNullException">If <paramref name="stream"/> Is <see langword="null"/></exception>
        Shared Function ReadLineAsBytes(ByVal stream As Stream) As Byte()
            If stream Is Nothing Then Throw New ArgumentNullException("stream")

            Using memoryStream As MemoryStream = New MemoryStream()

                While True
                    Dim justRead As Integer = stream.ReadByte()
                    If justRead = -1 AndAlso memoryStream.Length > 0 Then Exit While
                    If justRead = -1 AndAlso memoryStream.Length = 0 Then Return Nothing

                    'Check If we started at the end of the stream we read from
                    'And we have Not read anything from it yet

                    Dim readChar As Char = ChrW(justRead)
                    'Do not write \r or \n
                    If readChar <> vbCr AndAlso readChar <> vbLf Then memoryStream.WriteByte(CByte(justRead))
                    'Last point in CRLF pair
                    If readChar = vbLf Then Exit While
                End While

                Return memoryStream.ToArray()
            End Using
        End Function

        ''' <summary>
        ''' Read a line from the stream. <see cref="ReadLineAsBytes"/> for more documentation.
        ''' </summary>
        ''' <param name="stream">The stream to read from</param>
        ''' <returns>A line read from the stream Or <see langword="null"/> if nothing could be read from the stream</returns>
        ''' <exception cref="ArgumentNullException">If <paramref name="stream"/> Is <see langword="null"/></exception>
        Shared Function ReadLineAsAscii(ByVal stream As Stream) As String
            Dim readFromStream As Byte() = ReadLineAsBytes(stream)
            Return If(readFromStream IsNot Nothing, Encoding.ASCII.GetString(readFromStream), Nothing)
        End Function
    End Class
End Namespace

