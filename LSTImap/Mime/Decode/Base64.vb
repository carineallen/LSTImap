Imports System
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports LSTImap.Common.Logging

Namespace Mime.Decode

    ''' <summary>
	''' Utility class for dealing with Base64 encoded strings
	''' </summary>

    Public Class Base64

        ''' <summary>
		''' Decodes a base64 encoded string into the bytes it describes
		''' </summary>
		''' <param name="base64Encoded">The string to decode</param>
		''' <returns>A byte array that the base64 string described</returns>

        Shared Function Decode(ByVal base64Encoded As String) As Byte()

            'According to http://www.tribridge.com/blog/crm/blogs/brandon-kelly/2011-04-29/Solving-OutOfMemoryException-errors-when-attempting-to-attach-large-Base64-encoded-content-into-CRM-annotations.aspx
            'System.Convert.ToBase64String may leak a lot of memory
            'An OpenPop user reported that OutOfMemoryExceptions were thrown, And supplied the following
            'code for the fix. This should Not have memory leaks.
            'The code Is nearly identical to the example on MSDN:
            'http//msdn.microsoft.com/en-us/library/system.security.cryptography.frombase64transform.aspx#exampleToggle

            Try

                Using memoryStream As MemoryStream = New MemoryStream()
                    base64Encoded = base64Encoded.Replace(vbCrLf, "")
                    Dim inputBytes As Byte() = Encoding.ASCII.GetBytes(base64Encoded)

                    Using transform As FromBase64Transform = New FromBase64Transform(FromBase64TransformMode.DoNotIgnoreWhiteSpaces)
                        Dim outputBytes As Byte() = New Byte(transform.OutputBlockSize - 1) {}

                        'Transform the data in chunks the size of InputBlockSize.

                        Const inputBlockSize As Integer = 4
                        Dim currentOffset As Integer = 0

                        While inputBytes.Length - currentOffset > inputBlockSize
                            transform.TransformBlock(inputBytes, currentOffset, inputBlockSize, outputBytes, 0)
                            currentOffset += inputBlockSize
                            memoryStream.Write(outputBytes, 0, transform.OutputBlockSize)
                        End While

                        'Transform the final block of data.

                        outputBytes = transform.TransformFinalBlock(inputBytes, currentOffset, inputBytes.Length - currentOffset)
                        memoryStream.Write(outputBytes, 0, outputBytes.Length)
                    End Using

                    Return memoryStream.ToArray()
                End Using

            Catch e As FormatException
                DefaultLogger.Log.LogError("Base64: (FormatException) " & e.Message & vbCrLf & "On string: " & base64Encoded)

                Throw
            End Try
        End Function

        ''' <summary>
		''' Decodes a Base64 encoded string using a specified <see cref="System.Text.Encoding"/> 
		''' </summary>
		''' <param name="base64Encoded">Source string to decode</param>
		''' <param name="encoding">The encoding to use for the decoded byte array that <paramref name="base64Encoded"/> describes</param>
		''' <returns>A decoded string</returns>
		''' <exception cref="ArgumentNullException">If <paramref name="base64Encoded"/> Or <paramref name="encoding"/> Is <see langword="null"/></exception>
		''' <exception cref="FormatException">If <paramref name="base64Encoded"/> Is Not a valid base64 encoded string</exception>

        Shared Function Decode(ByVal base64Encoded As String, ByVal encoding As Encoding) As String
            If base64Encoded Is Nothing Then Throw New ArgumentNullException("base64Encoded")
            If encoding Is Nothing Then Throw New ArgumentNullException("encoding")
            Return encoding.GetString(Decode(base64Encoded))
        End Function
    End Class
End Namespace
