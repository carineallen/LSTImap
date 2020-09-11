Imports System

Namespace IMAP


    ''' <summary>
	''' Utility class that simplifies the usage of <see cref="IDisposable"/>
	''' </summary>
    Public MustInherit Class Disposable
        Implements IDisposable


        ''' <summary>
		''' Returns <see langword="true"/> if this instance has been disposed of, <see langword="false"/> otherwise
		''' </summary>
        Protected Property IsDisposed As Boolean

        ''' <summary>
		''' Releases unmanaged resources And performs other cleanup operations before the
		''' <see cref="Disposable"/> Is reclaimed by garbage collection.
		''' </summary>
        Protected Overrides Sub Finalize()
            Dispose(False)
        End Sub

        ''' <summary>
		''' Releases unmanaged And - optionally - managed resources
		''' </summary>
        Public Sub Dispose()
            If Not IsDisposed Then

                Try
                    Dispose(True)
                Finally
                    IsDisposed = True
                    GC.SuppressFinalize(Me)
                End Try
            End If
        End Sub

        ''' <summary>
		''' Releases unmanaged And - optionally - managed resources. Remember to call this method from your derived classes.
		''' </summary>
		''' <param name="disposing">
		''' Set to <c>true</c> to release both managed And unmanaged resources.<br/>
		''' Set to <c>false</c> to release only unmanaged resources.
		''' </param>
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        End Sub

        ''' <summary>
		''' Used to assert that the object has Not been disposed
		''' </summary>
		''' <exception cref="ObjectDisposedException">Thrown if the object Is in a disposed state.</exception>
		''' <remarks>
		''' The method Is to be used by the subclasses in order to provide a simple method for checking the 
		''' disposal state of the object.
		''' </remarks>
        Protected Sub AssertDisposed()
            If IsDisposed Then
                Dim typeName As String = [GetType]().FullName
                Throw New ObjectDisposedException(typeName, String.Format(System.Globalization.CultureInfo.InvariantCulture, "Cannot access a disposed {0}.", typeName))
            End If
        End Sub

        Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
            Throw New NotImplementedException()
        End Sub
    End Class
End Namespace
