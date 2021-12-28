Public Class ProdDAO
    Public ReadOnly Property Productos As Collection
    Public ReadOnly Property ProductosFechas As Collection
    Public ReadOnly Property ListaTrenyProd As Collection

    Public Sub New()
        Me.Productos = New Collection
        Me.ProductosFechas = New Collection
        Me.ListaTrenyProd = New Collection
    End Sub

    Public Sub LeerTodas(ruta As String)
        Dim p As Producto
        Dim col, aux As Collection
        col = AgenteBD.ObtenerAgente(ruta).Leer("SELECT * FROM Productos ORDER BY idProducto")
        For Each aux In col
            p = New Producto(Convert.ToInt64(aux(1)))
            p.DescripcionProducto = aux(2).ToString
            Me.Productos.Add(p)
        Next
    End Sub

    Public Sub Leer(ByRef p As Producto)
        Dim col As Collection : Dim aux As Collection
        col = AgenteBD.ObtenerAgente.Leer("SELECT * FROM Productos WHERE idProducto=" & p.IDProducto & ";")
        For Each aux In col
            p.DescripcionProducto = aux(2).ToString
        Next
    End Sub

    Public Sub LeerPorDescrip(ByRef p As Producto)
        Dim col As Collection : Dim aux As Collection
        col = AgenteBD.ObtenerAgente.Leer("SELECT * FROM Productos WHERE DescripProducto='" & p.DescripcionProducto & "';")
        For Each aux In col
            p.IDProducto = Convert.ToInt64(aux(1))
        Next
    End Sub

    Public Function Insertar(ByVal p As Producto) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("INSERT INTO Productos (DescripProducto) VALUES ('" & p.DescripcionProducto & "');")
    End Function

    Public Function Actualizar(ByVal p As Producto) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("UPDATE Productos SET DescripProducto='" & p.DescripcionProducto & "' WHERE idProducto=" & p.IDProducto & ";")
    End Function

    Public Function Borrar(ByVal p As Producto) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("DELETE FROM Productos WHERE idProducto=" & p.IDProducto & ";")
    End Function

    'Mostrar  el  número  de  viajes  y  una  lista  de  los  productos transportados por un  tren entre 2  fechas  (pudiendo elegir el tren y las fechas). 
    Public Sub TrenyFechas(d1 As Date, d2 As Date, tr As String)
        Dim col, aux As Collection
        Dim p As Producto
        col = AgenteBD.ObtenerAgente.Leer("SELECT IdProducto, DescripProducto
        FROM Viajes INNER JOIN Productos ON (Viajes.Producto = Productos.IdProducto)
        WHERE (FechaViaje BETWEEN #" & Format(d1, "MM/dd/yyyy") & "# AND #" & Format(d2, "MM/dd/yyyy") & "#)
        AND (Tren = '" & tr & "')
        GROUP BY IdProducto, DescripProducto;")
        For Each aux In col
            p = New Producto(Convert.ToInt64(aux(1)))
            p.DescripcionProducto = aux(2).ToString
            Me.ListaTrenyProd.Add(p)
        Next
    End Sub
    'Mostrar  un  listado  ordenado  (ranking)  de  los  productos  que más se han enviado entre 2 fechas que se podrán elegir. 
    Public Sub RankingProducto(d1 As Date, d2 As Date)
        Dim col, aux As Collection
        Dim p As Producto
        col = AgenteBD.ObtenerAgente.Leer("SELECT IdProducto, DescripProducto, COUNT(*)
        FROM Viajes INNER JOIN Productos ON (Viajes.Producto = Productos.IdProducto)
        WHERE FechaViaje BETWEEN #" & Format(d1, "MM/dd/yyyy") & "# AND #" & Format(d2, "MM/dd/yyyy") & "#
        GROUP BY IdProducto, DescripProducto
        ORDER BY COUNT(*) DESC;")
        For Each aux In col
            p = New Producto(Convert.ToInt64(aux(1)))
            p.DescripcionProducto = aux(2).ToString
            Me.ProductosFechas.Add(p)
        Next
    End Sub
End Class
