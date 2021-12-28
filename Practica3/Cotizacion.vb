Public Class Cotizacion
    Public Property Producto As Long
    Public Property Fecha As Date
    Public Property EurosTonelada As Double
    Public ReadOnly Property CDAO As CotDAO

    Public Sub New()
        Me.CDAO = New CotDAO
    End Sub

    Public Sub New(prod As Long, f As Date)
        Me.CDAO = New CotDAO
        Me.Producto = prod
        Me.Fecha = f
    End Sub

    Public Sub LeerTodasCotizaciones(ruta As String)
        Me.CDAO.LeerTodas(ruta)
    End Sub

    Public Sub LeerCotizacion()
        Me.CDAO.Leer(Me)
    End Sub

    Public Function InsertarCotizacion() As Integer
        Return Me.CDAO.Insertar(Me)
    End Function

    Public Function ActualizarCotizacion() As Integer
        Return Me.CDAO.Actualizar(Me)
    End Function

    Public Function BorrarCotizacion() As Integer
        Return Me.CDAO.Borrar(Me)
    End Function
End Class
