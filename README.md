### Introduction

This repository contains an easy-to-use .NET library written in Visual Basic(C# version will be coming soon) for sending
and receiving information from an email server using IMAP(Internet Message Access Protocol) according to [RFC 3501](https://tools.ietf.org/html/rfc3501).

### Motivation

I needed a Library to handle emails with IMAP, to read emails,create folders,delete messages,etc. there's already some libraries
for handle emails with a server, but I didn't find any that mached my needs(most I needed to be Free of charge and written in VB). 
But I found one that handles Post Office Protocol (POP3), [You can click here to see it](https://github.com/foens/hpop). It is written
in C# and dosen't handle IMAP but it help me to figure it out how to construct the decode part for my library.

### How to use

He is a small example on How to use the library:


		
	Imports System.IO
	Imports System.Net.Sockets
	Imports System.Net.Security
	Imports LSTImap.Mime
	Imports LSTImap.IMAP
	Imports System.ComponentModel

	Class Form1
	Dim ImapHost As String
	Dim Path As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
	Dim UserName As String
	Dim Password As String
	Dim PortNm As Integer
	Dim ImapClient As New ImapClient
	Dim LastDatum As Date
	Dim DataTable1 As New DataTable
	Dim EmailNr As Integer

	Public dateiname As String

	Private Sub Login()
		ImapHost = Host.Text
		PortNm = PortNr.Text
		If ImapClient.Connected = True Then
			ImapClient.Disconnect()
			Exit Sub
		Else
			UserName = TxtUsrNm.Text
			Password = TxtPass.Text
			Cursor = Cursors.WaitCursor

			Try
				ImapClient.Connect(ImapHost, PortNm, True)
				ListBox1.Items.Add("Connected")
				ImapClient.Login(UserName, Password)
				ListBox1.Items.Add("Login OK")
			Catch ex As Exception
				ImapClient.Disconnect()
				Cursor = Cursors.Default
				MsgBox(ex.Message.ToString)
				Exit Sub
			End Try

		End If

		Cursor = Cursors.Default
		CmdClose.Enabled = True
		CmdDownload.Enabled = False
		CmdDownload.Enabled = True
		TxtUsrNm.Enabled = False

		'Get Messages count

		ImapClient.SelectFolder("INBOX")
		MsgCount.Text = ImapClient.Selectedfolder.Exists.ToString
	End Sub

	Private Sub CmdDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDownload.Click
		getMesgs(MsgCount.Text)
	End Sub

	Sub getMesgs(ByVal Num_Emails As Integer)
		Dim Matched_Emails As List(Of String)
		LastDatum = Today.ToLongDateString
		Matched_Emails = ImapClient.SeacrhSince(LastDatum)
		For i = 0 To Matched_Emails.Count - 1
			Get_msg(Matched_Emails.Item(i).ToString)
		Next
		DataGridView1.DataSource = DataTable1
		Cursor = Cursors.Default
	End Sub

	Private Sub Get_msg(msg_num As String)
		DataGridView1.DataSource = Nothing
		Dim receivedBytes As Byte() = ImapClient.FetchAsBytes(msg_num)
		Dim Message1 As New Message(receivedBytes)
		Dim Email_To, Email_from, Email_Cc, Email_subject, Email_Date, Email_Body, Email_Attachements, Email_Attachements_name As String
		Dim TextBody As MessagePart = Message1.FindFirstHtmlVersion()
		Dim plainTextPart As MessagePart = Message1.FindFirstPlainTextVersion()
		Dim Datum As Date = Message1.Headers.DateSent.ToLocalTime()
		If Datum < LastDatum Then
			Exit Sub
		End If
		Dim Directory As String = ""
		Email_from = ""
		Email_To = ""
		Email_Cc = ""
		Email_subject = ""
		Email_Date = ""
		Email_Body = ""
		Email_Attachements = ""
		Email_Attachements_name = ""
		If (plainTextPart IsNot Nothing) Then

			For i = 0 To Message1.Headers.To.Count - 1
				If Email_To = "" Then
					Email_To = Message1.Headers.To(i).Address.ToString
				Else
					Email_To = Email_To + "," + Message1.Headers.To(i).Address.ToString
				End If
			Next

			For i = 0 To Message1.Headers.Cc.Count - 1
				If Email_Cc = "" Then
					Email_Cc = Message1.Headers.Cc(i).Address.ToString
				Else
					Email_Cc = Email_Cc + "," + Message1.Headers.Cc(i).Address.ToString
				End If
			Next
			Email_from = Message1.Headers.From.Address.ToString
			Email_subject = Message1.Headers.Subject.ToString
			Email_Date = Datum.Day.ToString + "." + Datum.Month.ToString + "." + Datum.Year.ToString + " " + Format(Datum.Hour, "00.#") + ":" + Format(Datum.Minute, "00.#") + ":" + Format(Datum.Second, "00.#")
			Email_Body = plainTextPart.GetBodyAsText()
		ElseIf TextBody IsNot Nothing Then
			For i = 0 To Message1.Headers.To.Count - 1
				If Email_To = "" Then
					Email_To = Message1.Headers.To(i).Address.ToString
				Else
					Email_To = Email_To + "," + Message1.Headers.To(i).Address.ToString
				End If
			Next

			For i = 0 To Message1.Headers.Cc.Count - 1
				If Email_Cc = "" Then
					Email_Cc = Message1.Headers.Cc(i).Address.ToString
				Else
					Email_Cc = Email_Cc + "," + Message1.Headers.Cc(i).Address.ToString
				End If
			Next
			Email_from = Message1.Headers.From.Address.ToString
			Email_subject = Message1.Headers.Subject.ToString
			Email_Date = Datum.Day.ToString + "." + Datum.Month.ToString + "." + Datum.Year.ToString + " " + Format(Datum.Hour, "00.#") + ":" + Format(Datum.Minute, "00.#") + ":" + Format(Datum.Second, "00.#")
			If TextBody IsNot Nothing Then
				Directory = "Emails\" + Message1.Headers.From.Address + "\" + Format(Datum.Year, "00.#") + Format(Datum.Month, "00.#") + "\" + Format(Datum.Day, "00.#") + "\" + Format(Datum.Hour, "00.#") + Format(Datum.Minute, "00.#") + Format(Datum.Second, "00.#")
				If My.Computer.FileSystem.DirectoryExists(Path + "\" + Directory) = False Then
					My.Computer.FileSystem.CreateDirectory(Path + "\" + Directory)
				End If
				TextBody.Save(New FileInfo(Path + "\" + Directory + "\" + "HtmlEmail.html"))
				If Email_Attachements = "" Then
					Email_Attachements = "HtmlEmail.html"
				Else
					Email_Attachements = Email_Attachements + ";" + "HtmlEmail.html"
				End If
			End If
			Email_Body = TextBody.GetBodyAsText()
		Else
			Dim textVersions As List(Of MessagePart) = Message1.FindAllTextVersions()
			If (textVersions.Count >= 1) Then

				For i = 0 To Message1.Headers.To.Count - 1
					If Email_To = "" Then
						Email_To = Message1.Headers.To(i).Address.ToString
					Else
						Email_To = Email_To + "," + Message1.Headers.To(i).Address.ToString
					End If
				Next

				For i = 0 To Message1.Headers.Cc.Count - 1
					If Email_To = "" Then
						Email_Cc = Message1.Headers.Cc(i).Address.ToString
					Else
						Email_Cc = Email_Cc + "," + Message1.Headers.Cc(i).Address.ToString
					End If
				Next
				Email_from = Message1.Headers.From.Address.ToString
				Email_subject = Message1.Headers.Subject.ToString
				Email_Date = Datum.Day.ToString + "." + Datum.Month.ToString + "." + Datum.Year.ToString + " " + Format(Datum.Hour, "00.#") + ":" + Format(Datum.Minute, "00.#") + ":" + Format(Datum.Second, "00.#")
				Email_Body = textVersions(0).GetBodyAsText()
			End If

		End If
		DataTable1.Rows.Add()
		EmailNr = EmailNr + 1
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(0) = EmailNr.ToString
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(1) = Email_from.ToString
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(2) = Email_To.ToString
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(3) = Email_Cc.ToString
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(4) = Email_Date.ToString
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(5) = Email_subject.ToString
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(6) = Email_Body.ToString

		Dim attachments As List(Of MessagePart) = Message1.FindAllAttachments()
		For Each attachment As MessagePart In attachments
			If attachment IsNot Nothing Then
				Try
					Directory = "Emails\" + Message1.Headers.From.Address + "\" + Format(Datum.Year, "00.#") + Format(Datum.Month, "00.#") + "\" + Format(Datum.Day, "00.#") + "\" + Format(Datum.Hour, "00.#") + Format(Datum.Minute, "00.#") + Format(Datum.Second, "00.#")
					If My.Computer.FileSystem.DirectoryExists(Path + "\" + Directory) = False Then
						My.Computer.FileSystem.CreateDirectory(Path + "\" + Directory)
					End If

					Dim sfd As New SaveFileDialog
					sfd.FileName = Path + "\" + Directory + "\" + attachment.FileName
					Try
						Dim File As New FileInfo(sfd.FileName)
						attachment.Save(File)
						Dim Split1 = sfd.FileName.Split("\")
						If Email_Attachements = "" Then
							Email_Attachements = Split1(Split1.Count - 1).ToString
						Else
							Email_Attachements = Email_Attachements + ";" + Split1(Split1.Count - 1).ToString
						End If
					Catch ex As Exception
					End Try


				Catch ex As Exception
				End Try
			Else
				Directory = ""
			End If
		Next
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(7) = Email_Attachements.ToString
		DataTable1.Rows(DataTable1.Rows.Count - 1).Item(8) = Directory.ToString
	End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		Login()
	End Sub

	Private Sub CmdClose_Click(sender As Object, e As EventArgs) Handles CmdClose.Click
		ImapClient.Logout()
		ImapClient.Disconnect()
	End Sub
	End Class
### Credits

This library is copyright © 2020 Carine Allen.


### License
This library is released under the [MIT license](https://github.com/carineallen/LSTImap/blob/add-license-1/LICENSE).
