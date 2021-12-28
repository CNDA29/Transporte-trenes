Public Class Producto
    Public Property IDProducto As Long
    Public Property DescripcionProducto As String
    Public ReadOnly Property PDAO As ProdDAO

    Public Sub New()
        Me.PDAO = New ProdDAO
    End Sub

    Public Sub New(idprod As Long)
        Me.PDAO = New ProdDAO
        Me.IDProducto = idprod
    End Sub

    Public Sub LeerTodosProductos(ruta As String)
        Me.PDAO.LeerTodas(ruta)
    End Sub

    Public Sub LeerProducto()
        Me.PDAO.Leer(Me)
    End Sub

    Public Function InsertarProducto() As Integer
        Return Me.PDAO.Insertar(Me)
    End Function

    Public Function ActualizarProducto() As Integer
        Return Me.PDAO.Actualizar(Me)
    End Function

    Public Function BorrarProducto() As Integer
        Return Me.PDAO.Borrar(Me)
    End Function
End Class
