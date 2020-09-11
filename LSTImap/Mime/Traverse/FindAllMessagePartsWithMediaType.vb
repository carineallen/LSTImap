Imports System
Imports System.Collections.Generic

Namespace Mime.Traverse

    '''<summary>
	''' Finds all the <see cref="MessagePart"/>s which have a given MediaType using a depth first traversal.
	'''</summary>
    Friend Class FindAllMessagePartsWithMediaType
		Implements IQuestionAnswerMessageTraverser(Of String, List(Of MessagePart))

        ''' <summary>
		'''inds all the <see cref="MessagePart"/>s with the given MediaType
		''' </summary>
		''' <param name="message">The <see cref="Message"/> to start looking in</param>
		''' <param name="question">The MediaType to look for. Case Is ignored.</param>
		''' <returns>
		''' A List of <see cref="MessagePart"/>s with the given MediaType.<br/>
		''' <br/>
		''' The List might be empty if no such <see cref="MessagePart"/>s were found.<br/>
		''' The order of the elements in the list Is the order which they are found using
		''' a depth first traversal of the <see cref="Message"/> hierarchy.
		''' </returns>
        Public Property VisitMessage(message As Message, question As String) As List(Of MessagePart) Implements IQuestionAnswerMessageTraverser(Of String, List(Of MessagePart)).VisitMessage
            Get
                If message Is Nothing Then Throw New ArgumentNullException("message")
                Return VisitMessagePart(message.MessagePart, question)
            End Get
            Set(value As List(Of MessagePart))

            End Set
        End Property

        ''' <summary>
        ''' Finds all the <see cref="MessagePart"/>s with the given MediaType
        ''' </summary>
        ''' <param name="messagePart">The <see cref="MessagePart"/> to start looking in</param>
        ''' <param name="question">The MediaType to look for. Case Is ignored.</param>
        ''' <returns>
        ''' A List of <see cref="MessagePart"/>s with the given MediaType.<br/>
        ''' <br/>
        ''' The List might be empty if no such <see cref="MessagePart"/>s were found.<br/>
        ''' The order of the elements in the list Is the order which they are found using
        ''' a depth first traversal of the <see cref="Message"/> hierarchy.
        ''' </returns>
        Private Property VisitMessagePart(messagePart As MessagePart, question As String) As List(Of MessagePart) Implements IQuestionAnswerMessageTraverser(Of String, List(Of MessagePart)).VisitMessagePart
            Get
                If messagePart Is Nothing Then Throw New ArgumentNullException("messagePart")
                Dim results As List(Of MessagePart) = New List(Of MessagePart)()
                If messagePart.ContentType.MediaType.Equals(question, StringComparison.OrdinalIgnoreCase) Then results.Add(messagePart)

                If messagePart.IsMultiPart Then

                    For Each part As MessagePart In messagePart.MessageParts
                        Dim result As List(Of MessagePart) = VisitMessagePart(part, question)
                        results.AddRange(result)
                    Next
                End If

                Return results
            End Get
            Set(value As List(Of MessagePart))

            End Set
        End Property


    End Class
End Namespace
