Public Class TrenDAO

    Public ReadOnly Property Trenes As Collection

    Public Sub New()
        Me.Trenes = New Collection
    End Sub

    Public Sub LeerTodas(ruta As String)
        Dim t As Tren
        Dim col, aux As Collection
        col = AgenteBD.ObtenerAgente(ruta).Leer("SELECT * FROM Trenes ORDER BY Matricula")
        For Each aux In col
            t = New Tren(aux(1).ToString)
            t.TipoTren = Convert.ToInt64(aux(2))
            Me.Trenes.Add(t)
        Next
    End Sub

    Public Sub Leer(ByRef t As Tren)
        Dim col As Collection : Dim aux As Collection
        col = AgenteBD.ObtenerAgente.Leer("SELECT * FROM Trenes WHERE Matricula='" & t.Matricula & "';")
        For Each aux In col
            t.TipoTren = Convert.ToInt64(aux(2))
        Next
    End Sub

    Public Function Insertar(ByVal t As Tren) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("INSERT INTO Trenes VALUES ('" & t.Matricula & "', " & t.TipoTren & ");")
    End Function

    Public Function Actualizar(ByVal t As Tren) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("UPDATE Trenes SET TipoTren=" & t.TipoTren & " WHERE Matricula='" & t.Matricula & "';")
    End Function

    Public Function Borrar(ByVal t As Tren) As Integer
        Return AgenteBD.ObtenerAgente.Modificar("DELETE FROM Trenes WHERE Matricula='" & t.Matricula & "';")
    End Function

End Class
