Public Class TipDAO

    Public ReadOnly Property Tipos_Tren As Collection
    Public ReadOnly Property TTFechas As Collection

    Public Sub New()
        Me.Tipos_Tren = New Collection
        Me.TTFechas = New Collection
    End Sub

    Public Sub LeerTodas(ruta As String)
        Dim tt As TipoTren
        Dim col, aux As Collection
        col = AgenteBD.ObtenerAgente(ruta).Leer("SELECT * FROM Tipos_Tren ORDER BY idTipoTren")
        For Each aux In col
            tt = New TipoTren(Convert.ToInt64(aux(1)))
            tt.Descripcion = aux(2).ToString
            tt.CapacidadMaxima = Convert.ToInt64(aux(3))
            Me.Tipos_Tren.Add(tt)
        Next
    End Sub

    Public Sub Leer(ByRef tt As TipoTren)
        Dim col As Collection : Dim aux As Collection
        col = AgenteBD.ObtenerAgente.Leer("SELECT * FROM Tipos_Tren WHERE idTipoTren=" & tt.IDTipoTren & ";")
        For Each aux In col
            tt.Descripcion = aux(2).ToString
            tt.CapacidadMaxima = Convert.ToInt64(aux(3))
        Next
    End Sub

    Public Sub LeerPorDescryCapMax(ByRef tt As TipoTren)
        Dim col As Collection : Dim aux As Collection
        col = AgenteBD.ObtenerAgente.Leer("SELECT * FROM Tipos_Tren WHERE DescTipoTren='" & tt.Descripcion & "' AND CapacidadMax=" & tt.CapacidadMaxima & ";")
        For Each aux In col
            tt.IDTipoTren = Convert.ToInt64(aux(1))
        Next
    End Sub

    Public Function Insertar(ByVal tt As TipoTren) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("INSERT INTO Tipos_Tren (DescTipoTren, CapacidadMax) VALUES ('" & tt.Descripcion & "', " & tt.CapacidadMaxima & ");")
    End Function

    Public Function Actualizar(ByVal tt As TipoTren) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("UPDATE Tipos_Tren SET DescTipoTren= '" & tt.Descripcion & "',CapacidadMax= " & tt.CapacidadMaxima & " WHERE idTipoTren=" & tt.IDTipoTren & ";")
    End Function

    Public Function Borrar(ByVal tt As TipoTren) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("DELETE FROM Tipos_Tren WHERE idTipoTren=" & tt.IDTipoTren & ";")
    End Function

    Public Sub RankingTipoTren()
        Dim col As Collection
        col = AgenteBD.ObtenerAgente.Leer("")
    End Sub
    'Mostrar un listado ordenado (ranking) del tipo de tren que ha realizado más viajes entre 2 fechas que se podrán elegir.
    Public Sub RankingTipoTren(d1 As Date, d2 As Date)
        Dim col, aux As Collection
        Dim tt As TipoTren
        col = AgenteBD.ObtenerAgente.Leer("SELECT IdTipoTren, DescTipoTren, CapacidadMax, COUNT(*)
        FROM ((Viajes INNER JOIN Trenes ON (Viajes.Tren = Trenes.Matricula))
        INNER JOIN Tipos_tren ON (Trenes.TipoTren = Tipos_Tren.IdTipoTren))
        WHERE FechaViaje BETWEEN #" & Format(d1, "MM/dd/yyyy") & "# AND #" & Format(d2, "MM/dd/yyyy") & "#
        GROUP BY IdTipoTren, DescTipoTren, CapacidadMax
        ORDER BY COUNT(*) DESC;")
        For Each aux In col
            tt = New TipoTren(Convert.ToInt64(aux(1)))
            tt.Descripcion = aux(2).ToString
            tt.CapacidadMaxima = Convert.ToInt64(aux(3))
            Me.TTFechas.Add(tt)
        Next
    End Sub
End Class

