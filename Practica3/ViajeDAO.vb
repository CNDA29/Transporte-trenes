Public Class ViajeDAO
    Public ReadOnly Property Viajes As Collection
    Public Property ViajeMayorBeneficio As Collection

    Public Sub New()
        Me.Viajes = New Collection
        Me.ViajeMayorBeneficio = New Collection
    End Sub

    Public Sub LeerTodas(ruta As String)
        Dim v As Viaje
        Dim col, aux As Collection
        col = AgenteBD.ObtenerAgente(ruta).Leer("SELECT * FROM Viajes ORDER BY FechaViaje")
        For Each aux In col
            v = New Viaje(Convert.ToDateTime(aux(1)), aux(2).ToString, Convert.ToInt64(aux(3)))
            v.ToneladasTransportadas = Convert.ToInt64(aux(4))
            Me.Viajes.Add(v)
        Next
    End Sub

    Public Sub Leer(ByRef v As Viaje)
        Dim col As Collection : Dim aux As Collection
        col = AgenteBD.ObtenerAgente.Leer("SELECT * FROM Viajes WHERE FechaViaje=#" & Format(v.FechaViaje, "MM/dd/yyyy") & "# AND Tren='" & v.Tren & "' AND Producto=" & v.Producto & ";")
        For Each aux In col
            v.ToneladasTransportadas = Convert.ToInt64(aux(4))
        Next
    End Sub

    Public Function Insertar(ByVal v As Viaje) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("INSERT INTO Viajes VALUES (#" & Format(v.FechaViaje, "MM/dd/yyyy") & "#, '" & v.Tren & "', " & v.Producto & ", " & v.ToneladasTransportadas & ");")
    End Function

    Public Function Actualizar(ByVal v As Viaje) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("UPDATE Viajes SET ToneladasTransportadas=" & v.ToneladasTransportadas & " WHERE FechaViaje=#" & Format(v.FechaViaje, "MM/dd/yyyy") & "# AND Tren='" & v.Tren & "' AND Producto=" & v.Producto & ";")
    End Function

    Public Function Borrar(ByVal v As Viaje) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("DELETE FROM Viajes WHERE FechaViaje=#" & Format(v.FechaViaje, "MM/dd/yyyy") & "# AND Tren='" & v.Tren & "' AND Producto=" & v.Producto & ";")
    End Function

    'Mostrar  toda  la  información  disponible  (fecha,  tren,  tipo  de tren,  productos  transportados,  toneladas  transportadas  de cada  uno  de  ellos,  cotizaciones  de  cada  producto  en  esa fecha,  
    'beneficio  por  producto  y  beneficio  total)  acerca  del viaje que haya supuesto un mayor beneficio económico. 
    Public Sub ViajeMaxBeneficio()
        ViajeMayorBeneficio = AgenteBD.ObtenerAgente.Leer("SELECT Viajes.FechaViaje, Trenes.TipoTren, Tipos_Tren.DescTipoTren, Productos.DescripProducto, Viajes.ToneladasTransportadas, Cotizaciones.EurosPorTonelada,
        Cotizaciones.EurosPorTonelada*Viajes.ToneladasTransportadas
        FROM (((Cotizaciones INNER JOIN Viajes ON (Viajes.FechaViaje = Cotizaciones.Fecha AND Viajes.Producto = Cotizaciones.Producto))
        INNER JOIN Trenes ON (Viajes.Tren = Trenes.Matricula))
        INNER JOIN Productos ON (Viajes.Producto = Productos.IdProducto))
        INNER JOIN Tipos_Tren ON (Trenes.TipoTren = Tipos_Tren.IdTipoTren)
        GROUP BY Viajes.FechaViaje, Trenes.TipoTren, Tipos_Tren.DescTipoTren, Productos.DescripProducto, Viajes.ToneladasTransportadas, Cotizaciones.EurosPorTonelada
        HAVING Cotizaciones.EurosPorTonelada*Viajes.ToneladasTransportadas =  
        (SELECT MAX (Cotizaciones.EurosPorTonelada*Viajes.ToneladasTransportadas)
        FROM Viajes INNER JOIN Cotizaciones ON (Viajes.FechaViaje = Cotizaciones.Fecha AND Viajes.Producto = Cotizaciones.Producto));")
    End Sub
End Class
