Imports System
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Globalization
Imports System.Net
Imports System.Net.Security
Imports System.Net.Sockets
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic
Imports System.Text.RegularExpressions
Imports LSTImap.Mime
Imports LSTImap.Mime.Header
Imports LSTImap.IMAP.Exceptions
Imports LSTImap.Common
Imports LSTImap.Common.Logging

Namespace IMAP
    Public Class ImapClient
        Inherits Disposable

        ''' <summary>
        ''' Stream used to communicate with the server.
        ''' </summary>
        ''' <returns></returns>
        Private Property m_sslStream As SslStream

        ''' <summary>
        ''' <see cref="NetworkStream"/> used to send and recive data.
        ''' </summary>
        ''' <returns></returns>
        Private Property NetworkS_tream As NetworkStream

        ''' <summary>
        ''' <see cref="StreamReader"/> used to read the information.
        ''' </summary>
        ''' <returns></returns>
        Private Property Read_Stream As StreamReader

        ''' <summary>
		''' Describes what state the <see cref="ImapClient"/> Is in.
		''' </summary>
        Private Property State As ConnectionState

        ''' <summary>
		''' Tells whether the <see cref="ImapClient"/> Is connected to a IMAP server Or Not.
		''' </summary>
        Public Property Connected As Boolean

        ''' <summary>
        ''' <see cref="TcpClient"/> that connect to a specific port and a specific host.
        ''' </summary>
        ''' <returns></returns>
        Private Property Imap As TcpClient

        ''' <summary>
		''' This Is the last response the server sent back when a command was issued to it.
		''' </summary>
        Private Property LastServerResponse As String

        ''' <summary>
		''' The folder that was selected by the SELECT command.
		''' </summary>
        Public Property Selectedfolder As EmailFolder

        ''' <summary>
        ''' A list of emails from the SEARCH command result.
        ''' </summary>
        Public Property MatchedEmails As List(Of String)

        ''' <summary>
        ''' Creates a new <see cref="ImapClient"/>
        ''' </summary>
        Public Sub New()
            SetInitialValues()
        End Sub

        ''' <summary>
        ''' Set the initial values. 
        ''' </summary>
        Private Sub SetInitialValues()

            'ImapClient is not connected
            Connected = False
            State = ConnectionState.Disconnected
            MatchedEmails = New List(Of String)()

        End Sub

        ''' <summary>
        ''' Connect the <see cref="ImapClient"/> with the provided host.
        ''' </summary>
        ''' <param name="hostname">the name of the host adress to be connected with.</param>
        ''' <param name="port">the port to use when connecting to the host.</param>
        ''' <param name="useSsl">tells if the connection must use Sssl security.</param>
        Public Sub Connect(ByVal hostname As String, ByVal port As Integer, ByVal useSsl As Boolean)
            Const defaultTimeOut As Integer = 60000
            Connect(hostname, port, useSsl, defaultTimeOut, defaultTimeOut, Nothing)
        End Sub


        ''' <summary>
        ''' Function to connect to the server, with the establish  <see cref="ReceiveTimeout"/> and <see cref="SendTimeout"/>.
        ''' </summary>
        ''' <param name="hostname">the name of the host adress to be connected with.</param>
        ''' <param name="port">the port to use when connecting to the host.</param>
        ''' <param name="useSsl">tells if the connection must use Sssl security.</param>
        ''' <param name="receiveTimeout">the amount of time a System.Net.Sockets.TcpClient will wait to receive data once a read operation is initiated.</param>
        ''' <param name="sendTimeout">the amount of time a System.Net.Sockets.TcpClient will wait for a send operation to complete successfully.</param>
        ''' <param name="certificateValidator">Verifies the remote Secure Sockets Layer (SSL) certificate used for authentication.</param>
        Public Sub Connect(ByVal hostname As String, ByVal port As Integer, ByVal useSsl As Boolean, ByVal receiveTimeout As Integer, ByVal sendTimeout As Integer, ByVal certificateValidator As RemoteCertificateValidationCallback)
            AssertDisposed()
            If hostname Is Nothing Then Throw New ArgumentNullException("hostname")
            If hostname.Length = 0 Then Throw New ArgumentException("hostname cannot be empty", "hostname")
            If port > IPEndPoint.MaxPort OrElse port < IPEndPoint.MinPort Then Throw New ArgumentOutOfRangeException("port")
            If receiveTimeout < 0 Then Throw New ArgumentOutOfRangeException("receiveTimeout")
            If sendTimeout < 0 Then Throw New ArgumentOutOfRangeException("sendTimeout")
            If State <> ConnectionState.Disconnected Then Throw New InvalidUseException("You cannot ask to connect to a IMAP server, when we are already connected to one. Disconnect first.")
            Imap = New TcpClient()
            Imap.ReceiveTimeout = receiveTimeout
            Imap.SendTimeout = sendTimeout

            Try
                Imap.Connect(hostname, port)
            Catch e As SocketException
                Imap.Close()
                DefaultLogger.Log.LogError("Connect(): " & e.Message)
                Throw New ImapServerNotFoundException("Server not found", e)
            End Try

            Dim stream As Stream

            If useSsl Then
                NetworkS_tream = Imap.GetStream()

                If certificateValidator Is Nothing Then
                    m_sslStream = New SslStream(NetworkS_tream, False)
                Else
                    m_sslStream = New SslStream(NetworkS_tream, False, certificateValidator)
                End If

                m_sslStream.ReadTimeout = receiveTimeout
                m_sslStream.WriteTimeout = sendTimeout
                m_sslStream.AuthenticateAsClient(hostname)
                stream = m_sslStream
            Else
                stream = Imap.GetStream()
            End If

            Connect(stream)
        End Sub

        ''' <summary>
        ''' Connect to the Stream that will be used in the <see cref="ImapClient"/>.
        ''' </summary>
        ''' <param name="stream">Stream that will be used in the <see cref="ImapClient"/></param>.
        Public Sub Connect(ByVal stream As Stream)
            AssertDisposed()
            If State <> ConnectionState.Disconnected Then Throw New InvalidUseException("You cannot ask to connect to a IMAP server, when we are already connected to one. Disconnect first.")
            If stream Is Nothing Then Throw New ArgumentNullException("stream")
            m_sslStream = stream
            Dim response As String = StreamUtility.ReadLineAsAscii(stream)

            Try
                State = ConnectionState.NotAuthenticated
                DefaultLogger.Log.LogDebug(String.Format("Connect-Response: ""{0}""", response))
                IsOkResponse(response, "*")
                Connected = True
            Catch e As ImapServerException
                DisconnectStreams()
                DefaultLogger.Log.LogError("Connect(): " & "Error with connection, maybe IMAP server not exist")
                DefaultLogger.Log.LogDebug("Last response from server was: " & LastServerResponse)
                Throw New ImapServerNotAvailableException("Server is not available", e)
            End Try
        End Sub


        ''' <summary>
        ''' Check if the Response of the server was a OK response.
        ''' </summary>
        ''' <param name="response">The last response readed from the stream.</param>
        ''' <param name="command">The command sended to the server.</param>
        Private Shared Sub IsOkResponse(ByVal response As String, command As String)
            Dim rsp As String
            If response Is Nothing Then Throw New ImapServerException("The stream used to retrieve responses from was closed")
            Try
                Dim Split = command.Split(" ")
                rsp = Split(0).ToString
            Catch ex As Exception
                rsp = "*"
            End Try
            If response.StartsWith(rsp + " OK", StringComparison.OrdinalIgnoreCase) Then Return
            Throw New ImapServerException("The server did not respond with a OK response. The response was: """ & response & """")
        End Sub


        ''' <summary>
        ''' Disconect the stream used in <see cref="ImapClient"/> and set back the initial values.
        ''' </summary>
        Private Sub DisconnectStreams()
            Try
                m_sslStream.Close()
            Finally
                SetInitialValues()
            End Try
        End Sub

        ''' <summary>
		''' Disconnects from IMAP server.
		''' Sends the LOGOUT command before closing the connection.
		''' </summary>
        Public Sub Disconnect()
            AssertDisposed()
            If State = ConnectionState.Disconnected Then Throw New InvalidUseException("You cannot disconnect a connection which is already disconnected")

            If State = ConnectionState.Authenticated Then
                Try
                    SendCommand("A3 LOGOUT")
                Finally
                    DisconnectStreams()
                End Try
            ElseIf State = ConnectionState.NotAuthenticated Then

                DisconnectStreams()
            End If

        End Sub

        ''' <summary>
        ''' Authenticate the connection with the server through a login process.
        ''' </summary>
        ''' <param name="username">The username to be used in the Login process.</param>
        ''' <param name="password">The password to be used in the Login process.</param>
        Public Sub Login(ByVal username As String, ByVal password As String)
            AssertDisposed()
            If username Is Nothing Then Throw New ArgumentNullException("username")
            If password Is Nothing Then Throw New ArgumentNullException("password")
            If State <> ConnectionState.NotAuthenticated Then Throw New InvalidUseException("You have to be connected and not authenticated when trying to login")

            Try

                SendCommand("A1 LOGIN " + ChrW(34) + username + ChrW(34) + " " + ChrW(34) + password + ChrW(34))
                State = ConnectionState.Authenticated
                SendCommand("n namespace")

            Catch e As ImapServerException
                DefaultLogger.Log.LogError("Problem logging in using login method. Server response was: " + LastServerResponse)
                Throw New InvalidLoginException(e)
            End Try
        End Sub

        ''' <summary>
        ''' Set an Not authenticated state with the server with a logout process.
        ''' </summary>
        Public Sub Logout()
            AssertDisposed()
            If State <> ConnectionState.Selected Then
                If State <> ConnectionState.Authenticated Then
                    Throw New InvalidUseException("You have to be Authenticated when trying to logout")
                End If
            End If


                SendCommand("A3 LOGOUT")


                State = ConnectionState.NotAuthenticated
        End Sub

        ''' <summary>
        ''' Select a folder in the current <see cref="ImapClient"/> email folders.
        ''' </summary>
        ''' <param name="folder">The folder to be selected.</param>
        Public Sub SelectFolder(ByVal folder As String)
            AssertDisposed()
            If folder Is Nothing Then Throw New ArgumentNullException("folder")
            If State <> ConnectionState.Authenticated And State <> ConnectionState.Selected Then
                Throw New InvalidUseException("You cannot get select a folder without logging in first")
            End If



            SendCommandSelect("g21 SELECT " + ChrW(34) + folder + ChrW(34))

            State = ConnectionState.Selected

        End Sub


        ''' <summary>
        ''' Search emails after a certain date in the current email folder.
        ''' </summary>
        ''' <param name="EmailDate">The date to look for.</param>
        ''' <returns></returns>
        Public Function SeacrhSince(ByVal EmailDate As Date) As List(Of String)
            AssertDisposed()
            Try
                If EmailDate.Date.ToString = "" Then Throw New ArgumentNullException("EmailDate")
                If State <> ConnectionState.Selected Then Throw New InvalidUseException("You cannot do a search without selecting a folder first")
                MatchedEmails.Clear()
                SendCommandSearch("s1 SEARCH SINCE " + EmailDate.Day.ToString + "-" + MonthName(EmailDate.Month, True) + "-" + EmailDate.Year.ToString)

            Catch e As ImapClientException
                Throw New InvalidUseException(e.Message.ToString)
            End Try
            Return MatchedEmails
        End Function


        ''' <summary>
        ''' Search for emails that match a specific string in a Header field.
        ''' Set <see cref="ToBeEqual"/> as True if the expected result should match with the condition and set it to False if the result should not match with the condition.
        ''' </summary>
        ''' <param name="field_name">the header field to look for a string</param>
        ''' <param name="Value">The string to look for</param>
        ''' <param name="ToBeEqual"> If the email has to match or not with the string</param>
        ''' <returns></returns>
        Public Function SeacrhHeader(ByVal field_name As String, Value As String, Optional ByVal ToBeEqual As Boolean = True) As List(Of String)
            AssertDisposed()
            Dim equal As String
            Try
                If field_name = "" Then Throw New ArgumentNullException("field_name")
                If Value = "" Then Throw New ArgumentNullException("Value")
                If State <> ConnectionState.Selected Then Throw New InvalidUseException("You cannot do a search without selecting a folder first")

                If ToBeEqual = False Then
                    equal = "NOT "
                Else
                    equal = ""
                End If
                MatchedEmails.Clear()
                SendCommandSearch("s2 SEARCH " + equal + "HEADER " + field_name + " " + ChrW(34) + Value + ChrW(34))

            Catch e As ImapClientException
                Throw New InvalidUseException(e.Message.ToString)
            End Try
            Return MatchedEmails
        End Function

        ''' <summary>
        ''' Search for emails since a certain date and that match a specific string in a Header field.
        ''' Set <see cref="ToBeEqual"/> as True if the expected result should match with the condition and set it to False if the result should not match with the condition.
        ''' </summary>
        ''' <param name="EmailDate">The date to look for.</param>
        ''' <param name="field_name">the header field to look for a string</param>
        ''' <param name="Value">The string to look for</param>
        ''' <param name="ToBeEqual">If the email has to match or not with the string</param>
        ''' <returns></returns>
        Public Function SeacrhHeaderSince(ByVal EmailDate As Date, ByVal field_name As String, Value As String, Optional ByVal ToBeEqual As Boolean = True) As List(Of String)
            AssertDisposed()
            Dim equal As String
            Try
                If field_name = "" Then Throw New ArgumentNullException("field_name")
                If Value = "" Then Throw New ArgumentNullException("Value")
                If EmailDate.Date.ToString = "" Then Throw New ArgumentNullException("EmailDate")
                If State <> ConnectionState.Selected Then Throw New InvalidUseException("You cannot do a search without selecting a folder first")

                If ToBeEqual = False Then
                    equal = "NOT "
                Else
                    equal = ""
                End If
                MatchedEmails.Clear()
                SendCommandSearch("s3 SEARCH SINCE " + EmailDate.Day.ToString + "-" + MonthName(EmailDate.Month, True) + "-" + EmailDate.Year.ToString + " " + equal + "HEADER " + field_name + " " + ChrW(34) + Value + ChrW(34))

            Catch e As ImapClientException
                Throw New InvalidUseException(e.Message.ToString)
            End Try
            Return MatchedEmails
        End Function

        ''' <summary>
        ''' Fetch an email as bytes.
        ''' </summary>
        ''' <param name="EmailNr">The email Nr.</param>
        ''' <returns></returns>
        Public Function FetchAsBytes(ByVal EmailNr As String) As Byte()
            AssertDisposed()
            If EmailNr Is Nothing Then Throw New ArgumentNullException("EmailNr")
            If State <> ConnectionState.Selected Then Throw New InvalidUseException("You cannot get fetch a email without selecting a folder first")

            Dim receivedBytes As Byte() = SendCommandFetch("f1 fetch " + EmailNr.ToString + " RFC822 ")


            Return receivedBytes
        End Function

        ''' <summary>
        ''' Fetch an email as a <see cref="Message"/>.
        ''' </summary>
        ''' <param name="EmailNr">The email Nr.</param>
        ''' <returns></returns>
        Public Function FetchAsMessage(ByVal EmailNr As String) As Message
            AssertDisposed()
            If EmailNr Is Nothing Then Throw New ArgumentNullException("EmailNr")
            If State <> ConnectionState.Selected Then Throw New InvalidUseException("You cannot get fetch a email without selecting a folder first")

            Dim receivedBytes As Byte() = FetchAsBytes(EmailNr)


            Return New Message(receivedBytes)
        End Function

        ''' <summary>
        ''' Gets the english name of the month.
        ''' </summary>
        ''' <param name="month">The month number.</param>
        ''' <param name="Abbreviation">If the name should be abbreviated(abbreviation with 3 letters).</param>
        ''' <returns></returns>
        Private Function MonthName(month As Integer, Abbreviation As Boolean) As String
            Dim Name As String
            If Abbreviation = True Then
                If month = 1 Then
                    Name = "Jan"
                ElseIf month = 2 Then
                    Name = "Feb"
                ElseIf month = 3 Then
                    Name = "Mar"
                ElseIf month = 4 Then
                    Name = "Apr"
                ElseIf month = 5 Then
                    Name = "May"
                ElseIf month = 6 Then
                    Name = "Jun"
                ElseIf month = 7 Then
                    Name = "Jul"
                ElseIf month = 8 Then
                    Name = "Aug"
                ElseIf month = 9 Then
                    Name = "Sep"
                ElseIf month = 10 Then
                    Name = "Oct"
                ElseIf month = 11 Then
                    Name = "Nov"
                ElseIf month = 12 Then
                    Name = "Dec"
                End If
            ElseIf Abbreviation = False Then
                If month = 1 Then
                    Name = "January"
                ElseIf month = 2 Then
                    Name = "February"
                ElseIf month = 3 Then
                    Name = "March"
                ElseIf month = 4 Then
                    Name = "April"
                ElseIf month = 5 Then
                    Name = "May"
                ElseIf month = 6 Then
                    Name = "June"
                ElseIf month = 7 Then
                    Name = "July"
                ElseIf month = 8 Then
                    Name = "August"
                ElseIf month = 9 Then
                    Name = "September"
                ElseIf month = 10 Then
                    Name = "October"
                ElseIf month = 11 Then
                    Name = "November"
                ElseIf month = 12 Then
                    Name = "December"
                End If
            End If
            Return Name
        End Function

        Public Sub Store(ByVal EmailNr As String, ByVal Flag As String)
            AssertDisposed()
            If EmailNr Is Nothing Then Throw New ArgumentNullException("EmailNr")
            If Flag Is Nothing Then Throw New ArgumentNullException("Flag")
            If State <> ConnectionState.Selected Then Throw New InvalidUseException("You cannot Store a email without selecting a folder first")

            SendCommand("s4 STORE " + EmailNr.ToString + " +FLAGS (\" & Flag & ")")

        End Sub

        Public Sub Expunge()
            AssertDisposed()
            If State <> ConnectionState.Selected Then Throw New InvalidUseException("You cannot Expunge emails without selecting a folder first")

            SendCommand("E01 EXPUNGE")

        End Sub

        ''' <summary>
        ''' Sends a command to the IMAP server.<br/>
        ''' If this fails, an exception Is thrown.
        ''' </summary>
        ''' <param name="command">The command to send to server</param>
        ''' <exception cref="ImapServerException">If the server did Not send an OK message to the command</exception>
        Private Sub SendCommand(ByVal command As String)
            Dim commandBytes As Byte() = Encoding.ASCII.GetBytes(command & vbCrLf)
            DefaultLogger.Log.LogDebug(String.Format("SendCommand: ""{0}""", command))
            m_sslStream.Write(commandBytes, 0, commandBytes.Length)
            m_sslStream.Flush()
            Dim Split = command.Split(" ")
            LastServerResponse = StreamUtility.ReadLineAsAscii(m_sslStream)
            While LastServerResponse.StartsWith(Split(0).ToString + " ") = False
                LastServerResponse = StreamUtility.ReadLineAsAscii(m_sslStream)
            End While
            DefaultLogger.Log.LogDebug(String.Format("Server-Response: ""{0}""", LastServerResponse))
            IsOkResponse(LastServerResponse, command)
        End Sub

        ''' <summary>
		''' Sends a SEARCH command to the IMAP server.<br/>
		''' If this fails, an exception Is thrown.
		''' </summary>
		''' <param name="command">The command to send to server</param>
		''' <exception cref="ImapServerException">If the server did Not send an OK message to the command</exception>
        Private Sub SendCommandSearch(ByVal command As String)
            Dim commandBytes As Byte() = Encoding.ASCII.GetBytes(command & vbCrLf)
            DefaultLogger.Log.LogDebug(String.Format("SendCommand: ""{0}""", command))
            m_sslStream.Write(commandBytes, 0, commandBytes.Length)
            m_sslStream.Flush()
            Dim Split = command.Split(" ")
            LastServerResponse = ""
            While LastServerResponse.StartsWith(Split(0).ToString + " ") = False
                LastServerResponse = StreamUtility.ReadLineAsAscii(m_sslStream)
                If LastServerResponse.Contains("* SEARCH") = True Then
                    Dim Split1 = LastServerResponse.Split(" ")
                    For i = 2 To Split1.Count - 1
                        MatchedEmails.Add(Split1(i).ToString)
                    Next
                End If
            End While
            DefaultLogger.Log.LogDebug(String.Format("Server-Response: ""{0}""", LastServerResponse))
            IsOkResponse(LastServerResponse, command)
        End Sub

        ''' <summary>
        ''' Sends a FETCH command to the IMAP server.<br/>
        ''' If this fails, an exception Is thrown.
        ''' </summary>
        ''' <param name="command">The command to send to server</param>
        ''' <exception cref="ImapServerException">If the server did Not send an OK message to the command</exception>
        Private Function SendCommandFetch(ByVal command As String)
            Dim commandBytes As Byte() = Encoding.ASCII.GetBytes(command & vbCrLf)
            Dim byteArrayBuilder As New MemoryStream()
            DefaultLogger.Log.LogDebug(String.Format("SendCommand: ""{0}""", command))
            m_sslStream.Write(commandBytes, 0, commandBytes.Length)
            m_sslStream.Flush()
            Dim Split = command.Split(" ")
            LastServerResponse = StreamUtility.ReadLineAsAscii(m_sslStream)
            While LastServerResponse.StartsWith(Split(0).ToString + " ") = False
                LastServerResponse = StreamUtility.ReadLineAsAscii(m_sslStream)
                If LastServerResponse.Contains("OK FETCH") = False Then
                    Dim m_buffer2 As Byte() = System.Text.Encoding.ASCII.GetBytes((LastServerResponse + vbNewLine).ToCharArray)
                    byteArrayBuilder.Write(m_buffer2, 0, m_buffer2.Length)
                End If
            End While
            Dim receivedBytes As Byte() = byteArrayBuilder.ToArray
            DefaultLogger.Log.LogDebug(String.Format("Server-Response: ""{0}""", LastServerResponse))
            IsOkResponse(LastServerResponse, command)
            Return receivedBytes
        End Function

        ''' <summary>
		''' Sends a SELECT command to the IMAP server.<br/>
		''' If this fails, an exception Is thrown.
		''' </summary>
		''' <param name="command">The command to send to server</param>
		''' <exception cref="ImapServerException">If the server did Not send an OK message to the command</exception>
        Private Sub SendCommandSelect(ByVal command As String)
            Dim commandBytes As Byte() = Encoding.ASCII.GetBytes(command & vbCrLf)
            Dim Infos As New NameValueCollection
            DefaultLogger.Log.LogDebug(String.Format("SendCommand: ""{0}""", command))
            m_sslStream.Write(commandBytes, 0, commandBytes.Length)
            m_sslStream.Flush()
            Dim Split = command.Split(" ")
            LastServerResponse = ""
            While LastServerResponse.StartsWith(Split(0).ToString + " ") = False
                LastServerResponse = StreamUtility.ReadLineAsAscii(m_sslStream)
                If LastServerResponse.Contains(" FLAGS") = True Then
                    Dim Split1 = LastServerResponse.Substring(LastServerResponse.IndexOf("(") + 1, LastServerResponse.IndexOf(")") - LastServerResponse.IndexOf("(") - 1).Split(" ")
                    For i = 0 To Split1.Count - 1
                        Infos.Add("FLAGS", Split1(i).Substring(1))
                    Next
                ElseIf LastServerResponse.Contains("[PERMANENTFLAGS") = True Then
                    Dim Split1 = LastServerResponse.Substring(LastServerResponse.IndexOf("(") + 1, LastServerResponse.IndexOf(")") - LastServerResponse.IndexOf("(") - 1).Split(" ")
                    For i = 0 To Split1.Count - 1
                        If Split1(i).Contains("*") = False Then
                            Infos.Add("PERMANENTFLAGS", Split1(i).Substring(1))
                        End If
                    Next
                ElseIf LastServerResponse.Contains(" EXISTS") = True Then
                    Dim Split1 = LastServerResponse.Split(" ")
                    Infos.Add("EXISTS", Split1(1).ToString)

                ElseIf LastServerResponse.Contains(" RECENT") = True Then
                    Dim Split1 = LastServerResponse.Split(" ")
                    Infos.Add("RECENT", Split1(1).ToString)

                ElseIf LastServerResponse.Contains("[UNSEEN") = True Then
                    Dim Split1 = LastServerResponse.Substring(LastServerResponse.IndexOf("[") + 1, LastServerResponse.IndexOf("]") - LastServerResponse.IndexOf("[") - 1).Split(" ")
                    Infos.Add("UNSEEN", Split1(1).ToString)
                ElseIf LastServerResponse.Contains("[UIDVALIDITY") = True Then
                    Dim Split1 = LastServerResponse.Substring(LastServerResponse.IndexOf("[") + 1, LastServerResponse.IndexOf("]") - LastServerResponse.IndexOf("[") - 1).Split(" ")
                    Infos.Add("UIDVALIDITY", Split1(1).ToString)
                ElseIf LastServerResponse.Contains("[UIDNEXT") = True Then
                    Dim Split1 = LastServerResponse.Substring(LastServerResponse.IndexOf("[") + 1, LastServerResponse.IndexOf("]") - LastServerResponse.IndexOf("[") - 1).Split(" ")
                    Infos.Add("UIDNEXT", Split1(1).ToString)

                End If
            End While
            DefaultLogger.Log.LogDebug(String.Format("Server-Response: ""{0}""", LastServerResponse))
            IsOkResponse(LastServerResponse, command)
            Selectedfolder = New EmailFolder(Infos)
        End Sub

    End Class
End Namespace
