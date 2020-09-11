Imports System
Imports System.Collections.Generic
Imports System.Globalization

Namespace Mime.Decode

    ''' <summary>
	''' Thanks to http://stackoverflow.com/a/7333402/477854 for inspiration
	''' This class can convert from strings Like "104 kB" (104 kilobytes) to bytes.
	''' It does Not know about differences such as kilobits vs kilobytes.
	''' </summary>
    Module SizeParser
        Private ReadOnly UnitsToMultiplicator As Dictionary(Of String, Long) = InitializeSizes()

        Private Function InitializeSizes() As Dictionary(Of String, Long)
            Return New Dictionary(Of String, Long) From {
                {"", 1L}, 'No unit is the same as a byte
                {"B", 1L}, 'Byte
                {"KB", 1024L}, 'Kilobyte
                {"MB", 1024L * 1024L}, 'Megabyte
                {"GB", 1024L * 1024L * 1024L}, 'Gigabyte
                {"TB", 1024L * 1024L * 1024L * 1024L} 'Terabyte
            }
        End Function

        Function Parse(ByVal value As String) As Long
            value = value.Trim()
            Dim unit As String = ExtractUnit(value)
            Dim valueWithoutUnit As String = value.Substring(0, value.Length - unit.Length).Trim()
            Dim multiplicatorForUnit2 As Long = MultiplicatorForUnit(unit)
            Dim size As Double = Double.Parse(valueWithoutUnit, NumberStyles.Number, CultureInfo.InvariantCulture)
            Return CLng((multiplicatorForUnit2 * size))
        End Function

        Private Function ExtractUnit(ByVal sizeWithUnit As String) As String
            'start right, end at the first digit
            Dim lastChar As Integer = sizeWithUnit.Length - 1
            Dim unitLength As Integer = 0

            'stop when a space
            'or digit is found
            While unitLength <= lastChar AndAlso sizeWithUnit(lastChar - unitLength) <> " " AndAlso Not IsDigit(sizeWithUnit(lastChar - unitLength))
                unitLength += 1
            End While

            Return sizeWithUnit.Substring(sizeWithUnit.Length - unitLength).ToUpperInvariant()
        End Function

        Private Function IsDigit(ByVal value As Char) As Boolean
            'we don't want to use char.IsDigit since it would accept esoterical unicode digits
            Return value >= "0" AndAlso value <= "9"
        End Function

        Private Function MultiplicatorForUnit(ByVal unit As String) As Long
            unit = unit.ToUpperInvariant()
            If Not UnitsToMultiplicator.ContainsKey(unit) Then Throw New ArgumentException("illegal or unknown unit: """ & unit & """", "unit")
            Return UnitsToMultiplicator(unit)
        End Function
    End Module
End Namespace
