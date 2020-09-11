Namespace Mime.Traverse
	''' <summary>
	''' This interface describes a MessageTraverser which Is able to traverse a Message structure
	''' And deliver some answer given some question.
	''' </summary>
	''' <typeparam name="TAnswer">This Is the type of the answer you want to have delivered.</typeparam>
	''' <typeparam name="TQuestion">This Is the type of the question you want to have answered.</typeparam>
	Public Interface IQuestionAnswerMessageTraverser(Of TQuestion, TAnswer)

		''' <summary>
		''' Call this when you want to apply this traverser on a <see cref="Message"/>.
		''' </summary>
		''' <param name="message">The <see cref="Message"/> which you want to traverse. Must Not be <see langword="null"/>.</param>
		''' <param name="question">The question</param>
		''' <returns>An answer</returns>
		Property VisitMessage(message As Message, question As TQuestion) As TAnswer

		''' <summary>
		''' Call this when you want to apply this traverser on a <see cref="MessagePart"/>.
		''' </summary>
		''' <param name="messagePart">The <see cref="MessagePart"/> which you want to traverse. Must Not be <see langword="null"/>.</param>
		''' <param name="question">The question</param>
		''' <returns>An answer</returns>
		Property VisitMessagePart(messagePart As MessagePart, question As TQuestion) As TAnswer
	End Interface
End Namespace