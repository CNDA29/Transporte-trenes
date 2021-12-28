Public Class TipoTren
    Public Property IDTipoTren As Long
    Public Property Descripcion As String
    Public Property CapacidadMaxima As Long
    Public ReadOnly Property TTDAO As TipDAO

    Public Sub New()
        Me.TTDAO = New TipDAO
    End Sub

    Public Sub New(idtt As Long)
        Me.TTDAO = New TipDAO
        Me.IDTipoTren = idtt
    End Sub

    Public Sub LeerTodosTiposTren(ruta As String)
        Me.TTDAO.LeerTodas(ruta)
    End Sub

    Public Sub LeerTipo()
        Me.TTDAO.Leer(Me)
    End Sub

    Public Function InsertarTipo() As Integer
        Return Me.TTDAO.Insertar(Me)
    End Function

    Public Function ActualizarTipo() As Integer
        Return Me.TTDAO.Actualizar(Me)
    End Function

    Public Function BorrarTipo() As Integer
        Return Me.TTDAO.Borrar(Me)
    End Function

End Class