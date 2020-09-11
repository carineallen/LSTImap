Imports System
Imports System.Collections.Generic

Namespace Mime.Traverse

    ''' <summary>
	''' This Is an abstract class which handles traversing of a <see cref="Message"/> tree structure.<br/>
	''' It runs through the message structure using a depth-first traversal.
	''' </summary>
	''' <typeparam name="TAnswer">The answer you want from traversing the message tree structure</typeparam>
    Public MustInherit Class AnswerMessageTraverser(Of TAnswer)
        Implements IAnswerMessageTraverser(Of TAnswer)

        ''' <summary>
        ''' Call this when you want an answer for a full message.
        ''' </summary>
        ''' <param name="message">The message you want to traverse</param>
        ''' <returns>An answer</returns>
        ''' <exception cref="ArgumentNullException">if <paramref name="message"/> Is <see langword="null"/></exception>
        Public Property VisitMessage(message As Message) As TAnswer Implements IAnswerMessageTraverser(Of TAnswer).VisitMessage
            Get
                If message Is Nothing Then Throw New ArgumentNullException("message")
                Return VisitMessagePart(message.MessagePart)
            End Get
            Set(value As TAnswer)

            End Set
        End Property

        ''' <summary>
        ''' Call this method when you want to find an answer for a <see cref="MessagePart"/>
        '''</summary>
        ''' <param name="messagePart">The <see cref="MessagePart"/> part you want an answer from.</param>
        ''' <returns>An answer</returns>
        ''' <exception cref="ArgumentNullException">if <paramref name="messagePart"/> Is <see langword="null"/></exception>
        Public Property VisitMessagePart(messagePart As MessagePart) As TAnswer Implements IAnswerMessageTraverser(Of TAnswer).VisitMessage
            Get
                If messagePart Is Nothing Then Throw New ArgumentNullException("messagePart")

                If messagePart.IsMultiPart Then
                    Dim leafAnswers As List(Of TAnswer) = New List(Of TAnswer)(messagePart.MessageParts.Count)

                    For Each part As MessagePart In messagePart.MessageParts
                        leafAnswers.Add(VisitMessagePart(part))
                    Next

                    Return MergeLeafAnswers(leafAnswers)
                End If

                Return CaseLeaf(messagePart)
            End Get
            Set(value As TAnswer)

            End Set
        End Property


        ''' <summary>
		''' For a concrete implementation an answer must be returned for a leaf <see cref="MessagePart"/>, which are
		''' MessageParts that are Not <see cref="MessagePart.IsMultiPart">MultiParts.</see>
		''' </summary>
		''' <param name="messagePart">The message part which Is a leaf And thereby Not a MultiPart</param>
		''' <returns>An answer</returns>
        Protected MustOverride Function CaseLeaf(ByVal messagePart As MessagePart) As TAnswer

        ''' <summary>
		''' For a concrete implementation, when a MultiPart <see cref="MessagePart"/> has fetched it's answers from it's children, these
		''' answers needs to be merged. This Is the responsibility of this method.
		''' </summary>
		''' <param name="leafAnswers">The answer that the leafs gave</param>
		''' <returns>A merged answer</returns>
        Protected MustOverride Function MergeLeafAnswers(ByVal leafAnswers As List(Of TAnswer)) As TAnswer
    End Class
End Namespace
