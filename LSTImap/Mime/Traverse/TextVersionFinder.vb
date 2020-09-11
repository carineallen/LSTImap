Imports System
Imports System.Collections.Generic

Namespace Mime.Traverse

    ''' <summary>
	''' Finds all text/[something] versions in a Message hierarchy
	''' </summary>
    Friend Class TextVersionFinder
        Inherits MultipleMessagePartFinder

        Protected Overrides Function CaseLeaf(ByVal messagePart As MessagePart) As List(Of MessagePart)
            If messagePart Is Nothing Then Throw New ArgumentNullException("messagePart")
            'Maximum space needed is one
            Dim leafAnswer As List(Of MessagePart) = New List(Of MessagePart)(1)
            If messagePart.IsText Then leafAnswer.Add(messagePart)
            Return leafAnswer
        End Function
    End Class
End Namespace
