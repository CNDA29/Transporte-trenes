Public Class Viaje
    Public Property FechaViaje As Date
    Public Property Tren As String
    Public Property Producto As Long
    Public Property ToneladasTransportadas As Long
    Public ReadOnly Property VDAO As ViajeDAO

    Public Sub New()
        Me.VDAO = New ViajeDAO
    End Sub

    Public Sub New(fechav As Date, t As String, p As Long)
        Me.VDAO = New ViajeDAO
        Me.FechaViaje = fechav
        Me.Tren = t
        Me.Producto = p
    End Sub

    Public Sub LeerTodosViajes(ruta As String)
        Me.VDAO.LeerTodas(ruta)
    End Sub

    Public Sub LeerViaje()
        Me.VDAO.Leer(Me)
    End Sub

    Public Function InsertarViaje() As Integer
        Return Me.VDAO.Insertar(Me)
    End Function

    Public Function ActualizarViaje() As Integer
        Return Me.VDAO.Actualizar(Me)
    End Function

    Public Function BorrarViaje() As Integer
        Return Me.VDAO.Borrar(Me)
    End Function
End Class
