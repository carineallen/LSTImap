Namespace IMAP

    ''' <summary>
	''' Some of these states are defined by <a href="https://tools.ietf.org/html/rfc3501">RFC 3501</a>.<br/>
	''' Which commands that are allowed in which state can be seen in the same RFC.<br/>
	''' <br/>
	''' Used to keep track of which state the <see cref="ImapClient"/> Is in.
	''' </summary>
    Friend Enum ConnectionState

        ''' <summary>
        ''' This is when the ImapClient is not connected to the server.
        ''' </summary>
        Disconnected

        ''' <summary>
        ''' This state is entered when a connection starts unless the connection has been pre-authenticated.
        ''' </summary>
        NotAuthenticated

        ''' <summary>
        ''' This state is entered when a pre-authenticated connection starts
        ''' </summary>
        Authenticated

        ''' <summary>
        ''' This state is entered when a mailbox has been successfully selected.
        ''' </summary>
        Selected

        ''' <summary>
        ''' In the logout state, the connection is being terminated.
        ''' </summary>
        Logout

    End Enum
End Namespace