Imports System
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Net.Mail
Imports System.Net.Mime

Namespace IMAP

    ''' <summary>
	''' Class that holds all information of a folder.
	''' This class cannot be instantiated from outside the library.
	''' </summary>
	''' <remarks>
	''' See <a href="https://tools.ietf.org/html/rfc3501#page-32">RFC 3501</a> for reference.
	''' </remarks>
    Public NotInheritable Class EmailFolder

        ''' <summary>
		''' Defined flags in the mailbox.
		''' </summary>
        Public Property Flags As List(Of String)

        ''' <summary>
        ''' A list of message flags that the client can change permanently.
        ''' </summary>
        Public Property PermanentFlags As List(Of String)

        ''' <summary>
        ''' The number of messages in the mailbox.
        ''' </summary>
        Public Property Exists As Integer

        ''' <summary>
        ''' The number of messages with the \Recent flag set.
        ''' </summary>
        Public Property Recent As String

        ''' <summary>
        ''' The message sequence number of the first unseen message in the mailbox.
        ''' </summary>
        Public Property Unseen As String

        ''' <summary>
        ''' The next unique identifier value.
        ''' </summary>
        Public Property UIDNext As String

        ''' <summary>
        ''' The unique identifier validity value.
        ''' </summary>
        Public Property UIDValidity As String


        ''' <summary>
        ''' Create a new Info collection of a folder
        ''' </summary>       
        Friend Sub New(ByVal Infos As NameValueCollection)
            If Infos Is Nothing Then Throw New ArgumentNullException("Infos")

            Flags = New List(Of String)()
            PermanentFlags = New List(Of String)()
            ExtractInfos(Infos)
        End Sub

        ''' <summary>
        ''' Extract a <see cref="NameValueCollection"/> of a email folder.
        ''' </summary>
        ''' <param name="Infos">The collection that should be extracted</param>
        ''' <exception cref="ArgumentNullException">If <paramref name="Infos"/> Is <see langword="null"/></exception>
        Private Sub ExtractInfos(ByVal Infos As NameValueCollection)
            If Infos Is Nothing Then Throw New ArgumentNullException("Infos")

            For Each Key As String In Infos.Keys
                If Key IsNot Nothing Then
                    Dim InfosValues As String() = Infos.GetValues(Key)
                    For Each InfoValue As String In InfosValues
                        Select Case Key.ToUpperInvariant()
                            Case "FLAGS"
                                Flags.Add(InfoValue.Trim())
                            Case "PERMANENTFLAGS"
                                PermanentFlags.Add(InfoValue.Trim())
                            Case "EXISTS"
                                Exists = InfoValue.Trim()
                            Case "RECENT"
                                Recent = InfoValue.Trim()
                            Case "UNSEEN"
                                Unseen = InfoValue.Trim()
                            Case "UIDNEXT"
                                UIDNext = InfoValue.Trim()
                            Case "UIDVALIDITY"
                                UIDValidity = InfoValue.Trim()
                        End Select
                    Next
                End If
            Next
        End Sub

    End Class

End Namespace