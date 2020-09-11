Imports System

Namespace Mime.Traverse

    '''<summary>
    ''' Finds the first <see cref="MessagePart"/> which have a given MediaType in a depth first traversal.
    '''</summary>
    Friend Class FindFirstMessagePartWithMediaType
        Implements IQuestionAnswerMessageTraverser(Of String, MessagePart)


        ''' <summary>
        ''' Finds the first <see cref="MessagePart"/> with the given MediaType
        '''</summary>
        ''' <param name="message">The <see cref="Message"/> to start looking in</param>
        ''' <param name="question">The MediaType to look for. Case Is ignored.</param>
        ''' <returns>A <see cref="MessagePart"/> with the given MediaType Or <see langword="null"/> if no such <see cref="MessagePart"/> was found</returns>
        Public Property VisitMessage(message As Message, question As String) As MessagePart Implements IQuestionAnswerMessageTraverser(Of String, MessagePart).VisitMessage
            Get
                If message Is Nothing Then Throw New ArgumentNullException("message")
                Return VisitMessagePart(message.MessagePart, question)
            End Get
            Set(value As MessagePart)

            End Set
        End Property

        ''' <summary>
		''' Finds the first <see cref="MessagePart"/> with the given MediaType
		''' </summary>
		''' <param name="messagePart">The <see cref="MessagePart"/> to start looking in</param>
		''' <param name="question">The MediaType to look for. Case Is ignored.</param>
		''' <returns>A <see cref="MessagePart"/> with the given MediaType Or <see langword="null"/> if no such <see cref="MessagePart"/> was found</returns>
        Public Property VisitMessagePart(messagePart As MessagePart, question As String) As MessagePart Implements IQuestionAnswerMessageTraverser(Of String, MessagePart).VisitMessagePart
            Get
                If messagePart Is Nothing Then Throw New ArgumentNullException("messagePart")
                If messagePart.ContentType.MediaType.Equals(question, StringComparison.OrdinalIgnoreCase) Then Return messagePart

                If messagePart.IsMultiPart Then

                    For Each part As MessagePart In messagePart.MessageParts
                        Dim result As MessagePart = VisitMessagePart(part, question)
                        If result IsNot Nothing Then Return result
                    Next
                End If

                Return Nothing
            End Get
            Set(value As MessagePart)
            End Set
        End Property

    End Class
End Namespace
