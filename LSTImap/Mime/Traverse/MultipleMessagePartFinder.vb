Imports System
Imports System.Collections.Generic

Namespace Mime.Traverse

    '''<summary>
    ''' An abstract class that implements the MergeLeafAnswers method.<br/>
    ''' The method simply returns the union of all answers from the leaves.
    '''</summary>
    Public MustInherit Class MultipleMessagePartFinder
        Inherits AnswerMessageTraverser(Of List(Of MessagePart))


        ''' <summary>
		''' Adds all the <paramref name="leafAnswers"/> in one big answer
		''' </summary>
		''' <param name="leafAnswers">The answers to merge</param>
		''' <returns>A list with has all the elements in the <paramref name="leafAnswers"/> lists</returns>
		'''<exception cref="ArgumentNullException">if <paramref name="leafAnswers"/> Is <see langword="null"/></exception>
        Protected Overrides Function MergeLeafAnswers(ByVal leafAnswers As List(Of List(Of MessagePart))) As List(Of MessagePart)
            If leafAnswers Is Nothing Then Throw New ArgumentNullException("leafAnswers")
            'We simply create a list with all the answer generated from the leaves
            Dim mergedResults As List(Of MessagePart) = New List(Of MessagePart)()

            For Each leafAnswer As List(Of MessagePart) In leafAnswers
                mergedResults.AddRange(leafAnswer)
            Next

            Return mergedResults
        End Function
    End Class
End Namespace
