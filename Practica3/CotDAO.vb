Public Class CotDAO
    Public ReadOnly Property Cotizaciones As Collection

    Public Sub New()
        Me.Cotizaciones = New Collection
    End Sub

    Public Sub LeerTodas(ruta As String)
        Dim c As Cotizacion
        Dim col, aux As Collection
        col = AgenteBD.ObtenerAgente(ruta).Leer("SELECT * FROM Cotizaciones ORDER BY Producto")
        For Each aux In col
            c = New Cotizacion(Convert.ToInt64(aux(1)), Convert.ToDateTime(aux(2)))
            c.EurosTonelada = Convert.ToDouble(aux(3))
            Me.Cotizaciones.Add(c)
        Next
    End Sub

    Public Sub Leer(ByRef c As Cotizacion)
        Dim col As Collection : Dim aux As Collection
        col = AgenteBD.ObtenerAgente.Leer("SELECT * FROM Cotizaciones WHERE Producto=" & c.Producto & " AND Fecha=#" & Format(c.Fecha, "MM/dd/yyyy") & "#;")
        For Each aux In col
            c.EurosTonelada = Convert.ToDouble(aux(3))
        Next
    End Sub

    Public Function Insertar(ByVal c As Cotizacion) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("INSERT INTO Cotizaciones VALUES (" & c.Producto & ", #" & Format(c.Fecha, "MM/dd/yyyy") & "#, " & c.EurosTonelada.ToString.Replace(",", ".") & ");")
    End Function

    Public Function Actualizar(ByVal c As Cotizacion) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("UPDATE Cotizaciones SET EurosPorTonelada=" & c.EurosTonelada.ToString.Replace(",", ".") & " WHERE Producto=" & c.Producto & " AND Fecha=#" & Format(c.Fecha, "MM/dd/yyyy") & "#;")
    End Function

    Public Function Borrar(ByVal c As Cotizacion) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("DELETE FROM Cotizaciones WHERE Producto=" & c.Producto & " AND Fecha=#" & Format(c.Fecha, "MM/dd/yyyy") & "#;")
    End Function

End Class
