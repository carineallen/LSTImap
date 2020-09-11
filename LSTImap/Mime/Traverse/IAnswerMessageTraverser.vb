Namespace Mime.Traverse


	''' <summary>
	''' This interface describes a MessageTraverser which Is able to traverse a Message hierarchy structure
	''' And deliver some answer.
	''' </summary>
	''' <typeparam name="TAnswer">This Is the type of the answer you want to have delivered.</typeparam>
	Public Interface IAnswerMessageTraverser(Of TAnswer)
		''' <summary>
		''' Call this when you want to apply this traverser on a <see cref="Message"/>.
		''' </summary>
		''' <param name="message">The <see cref="Message"/> which you want to traverse. Must Not be <see langword="null"/>.</param>
		''' <returns>An answer</returns>
		Property VisitMessage(message As Message) As TAnswer

		''' <summary>
		''' Call this when you want to apply this traverser on a <see cref="MessagePart"/>.
		''' </summary>
		''' <param name="messagePart">The <see cref="MessagePart"/> which you want to traverse. Must Not be <see langword="null"/>.</param>
		''' <returns>An answer</returns>
		Property VisitMessage(messagePart As MessagePart) As TAnswer

	End Interface

End Namespace
