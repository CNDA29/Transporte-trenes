Public Class Tren
    Public Property Matricula As String
    Public Property TipoTren As Long
    Public ReadOnly Property TDAO As TrenDAO

    Public Sub New()
        Me.TDAO = New TrenDAO
    End Sub

    Public Sub New(matricula As String)
        Me.TDAO = New TrenDAO
        Me.Matricula = matricula
    End Sub

    Public Sub LeerTodosTrenes(ruta As String)
        Me.TDAO.LeerTodas(ruta)
    End Sub

    Public Sub LeerTren()
        Me.TDAO.Leer(Me)
    End Sub

    Public Function InsertarTren() As Integer
        Return Me.TDAO.Insertar(Me)
    End Function

    Public Function ActualizarTren() As Integer
        Return Me.TDAO.Actualizar(Me)
    End Function

    Public Function BorrarTren() As Integer
        Return Me.TDAO.Borrar(Me)
    End Function
End Class
