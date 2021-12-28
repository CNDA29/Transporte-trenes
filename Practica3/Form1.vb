Public Class Form1
    Private t As Tren
    Private p As Producto
    Private c As Cotizacion
    Private v As Viaje
    Private tt As TipoTren

    Private Sub Seleccionar_Click(sender As Object, e As EventArgs) Handles ButtonSeleccionarBBDD.Click
        Me.Ruta.InitialDirectory = Application.StartupPath
        If (Me.Ruta.ShowDialog() = DialogResult.OK) Then
            Me.TextBoxBBDD.Text = Me.Ruta.FileName
            ButtonConectar.Enabled = True
        End If
    End Sub

    Private Sub Conectar_Click(sender As Object, e As EventArgs) Handles ButtonConectar.Click
        Dim taux As Tren
        Dim paux As Producto
        Dim caux As Cotizacion
        Dim vaux As Viaje
        Dim ttaux As TipoTren
        Me.t = New Tren
        Me.p = New Producto
        Me.c = New Cotizacion
        Me.v = New Viaje
        Me.tt = New TipoTren
        Try
            Me.t.LeerTodosTrenes(Ruta.FileName)
            Me.p.LeerTodosProductos(Ruta.FileName)
            Me.c.LeerTodasCotizaciones(Ruta.FileName)
            Me.v.LeerTodosViajes(Ruta.FileName)
            Me.tt.LeerTodosTiposTren(Ruta.FileName)
            For Each taux In Me.t.TDAO.Trenes
                Me.ListBoxMatriculaTren.Items.Add(taux.Matricula)
                Me.ListBoxMatTrenDispViajes.Items.Add(taux.Matricula)
                Me.ListBoxtTipoTren.Items.Add(taux.TipoTren)
                Me.ListBoxTrenesDispConsultas.Items.Add(taux.Matricula)
            Next
            For Each paux In Me.p.PDAO.Productos
                Me.ListBoxIDProdProd.Items.Add(paux.IDProducto)
                Me.ListBoxIDProdDisp.Items.Add(paux.IDProducto)
                Me.ListBoxProdDispCot.Items.Add(paux.IDProducto)
                Me.ListBoxDescrpProd.Items.Add(paux.DescripcionProducto)
            Next
            For Each caux In Me.c.CDAO.Cotizaciones
                Me.ListBoxProdCot.Items.Add(caux.Producto)
                Me.ListBoxFechCot.Items.Add(caux.Fecha)
                Me.ListBoxETCot.Items.Add(caux.EurosTonelada)
            Next
            For Each vaux In Me.v.VDAO.Viajes
                Me.ListBoxFechaViajes.Items.Add(vaux.FechaViaje)
                Me.ListBoxIDTren.Items.Add(vaux.Tren)
                Me.ListBoxIDProd.Items.Add(vaux.Producto)
                Me.ListBoxTonTrans.Items.Add(vaux.ToneladasTransportadas)
            Next
            For Each ttaux In Me.tt.TTDAO.Tipos_Tren
                Me.ListBoxIDTipoTT.Items.Add(ttaux.IDTipoTren)
                Me.ListBoxIDTTTren.Items.Add(ttaux.IDTipoTren)
                Me.ListBoxCapacMaxTT.Items.Add(ttaux.CapacidadMaxima)
                Me.ListBoxDescrpTT.Items.Add(ttaux.Descripcion)
            Next
            ButtonConectar.Enabled = False
            ButtonSeleccionarBBDD.Enabled = False
            ButtonRealizarConsultaMaxViaje.Enabled = True
            ButtonRealizarConsultaRankingProd.Enabled = True
            ButtonRealizarConsultaRankTr.Enabled = True
            ButtonAñadirTT.Enabled = True
            ButtonAñadirProd.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub
    ''GESTIONES''
    'GESTIONES TIPOS_TRENES'
    Private Sub ListBoxIDTipoTT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxIDTipoTT.SelectedIndexChanged
        If Not Me.ListBoxIDTipoTT.SelectedItem Is Nothing Then
            Me.TextBoxIDTipoTrenTT.ReadOnly = True
            Me.ButtonEliminarTT.Enabled = True
            Me.ButtonActualizarTT.Enabled = True
            Me.tt = New TipoTren(Convert.ToInt64(Me.ListBoxIDTipoTT.SelectedItem))
            Try
                Me.tt.LeerTipo()
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.TextBoxIDTipoTrenTT.Text = Me.tt.IDTipoTren.ToString
            Me.TextBoxDescripcionTT.Text = Me.tt.Descripcion.ToString
            Me.TextBoxCapacidadMaxTT.Text = Me.tt.CapacidadMaxima.ToString
        End If
    End Sub

    Private Sub ButtonAñadirTT_Click(sender As Object, e As EventArgs) Handles ButtonAñadirTT.Click
        If Me.TextBoxDescripcionTT.Text <> String.Empty And Me.TextBoxCapacidadMaxTT.Text <> String.Empty Then
            tt = New TipoTren()
            tt.Descripcion = TextBoxDescripcionTT.Text
            tt.CapacidadMaxima = TextBoxCapacidadMaxTT.Text
            Try
                If tt.InsertarTipo() <> 1 Then
                    MessageBox.Show("INSERT return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.tt.TTDAO.LeerPorDescryCapMax(tt)
            Me.ListBoxIDTipoTT.Items.Add(tt.IDTipoTren)
            Me.ListBoxIDTTTren.Items.Add(tt.IDTipoTren)
            Me.ListBoxCapacMaxTT.Items.Add(tt.CapacidadMaxima)
            Me.ListBoxDescrpTT.Items.Add(tt.Descripcion)
            Me.ButtonLimpiarTT.PerformClick()
        End If
    End Sub

    Private Sub ButtonActualizarTT_Click(sender As Object, e As EventArgs) Handles ButtonActualizarTT.Click
        If Not tt Is Nothing Then
            tt.Descripcion = TextBoxDescripcionTT.Text
            tt.CapacidadMaxima = TextBoxCapacidadMaxTT.Text
            Try
                If tt.ActualizarTipo() <> 1 Then
                    MessageBox.Show("UPDATE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            MessageBox.Show("Capacidad maxima= " & tt.CapacidadMaxima & " Descripcion= " & tt.Descripcion & " actualizado correctamente!!")
            Me.ButtonLimpiarTT.PerformClick()
        End If
    End Sub

    Private Sub ButtonEliminarTT_Click(sender As Object, e As EventArgs) Handles ButtonEliminarTT.Click
        If Not tt Is Nothing Then
            If MessageBox.Show("¿Estás seguro que quieres borrar " & Me.tt.IDTipoTren & "?", "Por favor, confirme", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    If tt.BorrarTipo() <> 1 Then
                        MessageBox.Show("DELETE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End Try
                Me.ListBoxIDTipoTT.Items.Remove(tt.IDTipoTren)
                Me.ListBoxIDTTTren.Items.Remove(tt.IDTipoTren)
                Me.ListBoxCapacMaxTT.Items.Remove(tt.CapacidadMaxima)
                Me.ListBoxDescrpTT.Items.Remove(tt.Descripcion)
            End If
            Me.ButtonLimpiarTT.PerformClick()
        End If
    End Sub

    Private Sub ButtonLimpiarTT_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarTT.Click
        Me.TextBoxIDTipoTrenTT.Text = String.Empty
        Me.TextBoxDescripcionTT.Text = String.Empty
        Me.TextBoxCapacidadMaxTT.Text = String.Empty
        Me.ButtonActualizarTT.Enabled = False
        Me.ButtonEliminarTT.Enabled = False
        Me.TextBoxIDTipoTrenTT.ReadOnly = True
        Me.ListBoxIDTipoTT.SelectedItems.Clear()
    End Sub
    'GESTIONES TRENES'
    Private Sub ListBoxMatriculaTren_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxMatriculaTren.SelectedIndexChanged
        If Not Me.ListBoxMatriculaTren.SelectedItem Is Nothing Then
            ButtonActualizarTren.Enabled = True
            ButtonEliminarTren.Enabled = True
            Me.TextBoxMatriculaTren.ReadOnly = True
            Me.t = New Tren(Me.ListBoxMatriculaTren.SelectedItem.ToString)
            Try
                Me.t.LeerTren()
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.TextBoxMatriculaTren.Text = Me.t.Matricula.ToString
            Me.TextBoxIDTipoTrenTren.Text = Me.t.TipoTren.ToString
        End If
    End Sub

    Private Sub ListBoxIDTTTren_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxIDTTTren.SelectedIndexChanged
        If Not Me.ListBoxIDTTTren.SelectedItem Is Nothing Then
            ButtonAñadirTren.Enabled = True
            Me.TextBoxIDTipoTrenTren.Text = Me.ListBoxIDTTTren.SelectedItem
        End If
    End Sub

    Private Sub ButtonAñadirTren_Click(sender As Object, e As EventArgs) Handles ButtonAñadirTren.Click
        If Me.TextBoxMatriculaTren.Text <> String.Empty And Me.TextBoxIDTipoTrenTren.Text <> String.Empty Then
            t = New Tren()
            t.Matricula = TextBoxMatriculaTren.Text
            t.TipoTren = TextBoxIDTipoTrenTren.Text
            Try
                If t.InsertarTren() <> 1 Then
                    MessageBox.Show("INSERT return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.ListBoxMatriculaTren.Items.Add(t.Matricula)
            Me.ListBoxMatTrenDispViajes.Items.Add(t.Matricula)
            Me.ListBoxtTipoTren.Items.Add(t.TipoTren)
            Me.ListBoxTrenesDispConsultas.Items.Add(t.Matricula)
            Me.ButtonLimpiarTren.PerformClick()
        End If
    End Sub

    Private Sub ButtonActualizarTren_Click(sender As Object, e As EventArgs) Handles ButtonActualizarTren.Click
        If Not t Is Nothing Then
            t.TipoTren = TextBoxIDTipoTrenTren.Text
            Try
                If t.ActualizarTren() <> 1 Then
                    MessageBox.Show("UPDATE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            MessageBox.Show("Tipo Tren = " & t.TipoTren & " actualizado correctamente!!")
            Me.ButtonLimpiarTren.PerformClick()
        End If
    End Sub

    Private Sub ButtonEliminarTren_Click(sender As Object, e As EventArgs) Handles ButtonEliminarTren.Click
        If Not t Is Nothing Then
            If MessageBox.Show("¿Estás seguro que quieres borrar " & Me.t.Matricula & "?", "Por favor, confirme", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    If t.BorrarTren() <> 1 Then
                        MessageBox.Show("DELETE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End Try
                Me.ListBoxMatriculaTren.Items.Remove(t.Matricula)
                Me.ListBoxMatTrenDispViajes.Items.Remove(t.Matricula)
                Me.ListBoxtTipoTren.Items.Remove(t.TipoTren)
                Me.ListBoxTrenesDispConsultas.Items.Remove(t.Matricula)
            End If
            Me.ButtonLimpiarTren.PerformClick()
        End If
    End Sub

    Private Sub ButtonLimpiarTren_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarTren.Click
        Me.TextBoxMatriculaTren.Text = String.Empty
        Me.TextBoxIDTipoTrenTren.Text = String.Empty
        Me.ButtonActualizarTren.Enabled = False
        Me.ButtonEliminarTren.Enabled = False
        Me.ButtonAñadirTren.Enabled = False
        Me.TextBoxMatriculaTren.ReadOnly = False
        Me.TextBoxIDTipoTrenTren.ReadOnly = True
        Me.ListBoxMatriculaTren.SelectedItems.Clear()
        Me.ListBoxIDTTTren.SelectedItems.Clear()
    End Sub
    'GESTIONES COTIZACIONES'
    Private Sub ListBoxProdCot_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxProdCot.SelectedIndexChanged
        If Not Me.ListBoxProdCot.SelectedItem Is Nothing Then
            Me.ListBoxFechCot.Enabled = True
            Me.ListBoxProdCot.Enabled = False
        End If
    End Sub

    Private Sub ListBoxFechCot_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxFechCot.SelectedIndexChanged
        If Not Me.ListBoxFechCot.SelectedItem Is Nothing Then
            Me.ListBoxFechCot.Enabled = False
            ButtonActualizarCot.Enabled = True
            ButtonEliminarCot.Enabled = True
            Me.DateTimePickerCot.Enabled = False
            Me.c = New Cotizacion(Convert.ToInt64(Me.ListBoxProdCot.SelectedItem), Convert.ToDateTime(Me.ListBoxFechCot.SelectedItem))
            Try
                Me.c.LeerCotizacion()
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.TextBoxIDProductoCot.Text = Me.c.Producto.ToString
            Me.DateTimePickerCot.Value = Me.c.Fecha.ToString
            Me.TextBoxEurosTonCot.Text = Me.c.EurosTonelada.ToString
        End If
    End Sub

    Private Sub ListBoxProdDispCot_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxProdDispCot.SelectedIndexChanged
        If Not Me.ListBoxProdDispCot.SelectedItem Is Nothing Then
            ButtonAñadirCot.Enabled = True
            Me.TextBoxIDProductoCot.Text = Me.ListBoxProdDispCot.SelectedItem
        End If
    End Sub

    Private Sub ButtonAñadirCot_Click(sender As Object, e As EventArgs) Handles ButtonAñadirCot.Click
        If Me.TextBoxIDProductoCot.Text <> String.Empty And Me.TextBoxEurosTonCot.Text <> String.Empty Then
            c = New Cotizacion()
            c.Producto = TextBoxIDProductoCot.Text
            c.Fecha = DateTimePickerCot.Value.Date.ToString("d")
            c.EurosTonelada = TextBoxEurosTonCot.Text
            Try
                If c.InsertarCotizacion() <> 1 Then
                    MessageBox.Show("INSERT return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.ListBoxProdCot.Items.Add(c.Producto)
            Me.ListBoxFechCot.Items.Add(c.Fecha)
            Me.ListBoxETCot.Items.Add(c.EurosTonelada)
            Me.ButtonLimpiarCot.PerformClick()
        End If
    End Sub

    Private Sub ButtonActualizarCot_Click(sender As Object, e As EventArgs) Handles ButtonActualizarCot.Click
        If Not c Is Nothing Then
            c.EurosTonelada = TextBoxEurosTonCot.Text
            Try
                If c.ActualizarCotizacion() <> 1 Then
                    MessageBox.Show("UPDATE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            MessageBox.Show("Euros/tonelada= " & c.EurosTonelada & " actualizado correctamente!!")
            Me.ButtonLimpiarCot.PerformClick()
        End If
    End Sub

    Private Sub ButtonEliminarCot_Click(sender As Object, e As EventArgs) Handles ButtonEliminarCot.Click
        If Not c Is Nothing Then
            If MessageBox.Show("¿Estás seguro que quieres borrar Producto=" & Me.c.Producto & " Fecha= " & Me.c.Fecha & "?", "Por favor, confirme", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    If c.BorrarCotizacion() <> 1 Then
                        MessageBox.Show("DELETE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End Try
                Me.ListBoxProdCot.Items.Remove(c.Producto)
                Me.ListBoxFechCot.Items.Remove(c.Fecha)
                Me.ListBoxETCot.Items.Remove(c.EurosTonelada)
            End If
            Me.ButtonLimpiarCot.PerformClick()
        End If
    End Sub

    Private Sub ButtonLimpiarCot_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarCot.Click
        Me.TextBoxIDProductoCot.Text = String.Empty
        Me.TextBoxEurosTonCot.Text = String.Empty
        Me.ButtonActualizarCot.Enabled = False
        Me.ButtonEliminarCot.Enabled = False
        Me.ButtonAñadirCot.Enabled = False
        Me.DateTimePickerCot.Enabled = True
        Me.ListBoxProdCot.Enabled = True
        Me.ListBoxFechCot.Enabled = False
        Me.ListBoxProdCot.SelectedItems.Clear()
        Me.ListBoxFechCot.SelectedItems.Clear()
        Me.ListBoxProdDispCot.SelectedItems.Clear()
    End Sub
    'GESTIONES PRODUCTOS'
    Private Sub ListBoxIDProdProd_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxIDProdProd.SelectedIndexChanged
        If Not Me.ListBoxIDProdProd.SelectedItem Is Nothing Then
            ButtonActualizarProd.Enabled = True
            ButtonEliminarProd.Enabled = True
            Me.TextBoxIDProductoProd.ReadOnly = True
            Me.p = New Producto(Convert.ToInt64(Me.ListBoxIDProdProd.SelectedItem.ToString))
            Try
                Me.p.LeerProducto()
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.TextBoxIDProductoProd.Text = Me.p.IDProducto.ToString
            Me.TextBoxDescripcionProd.Text = Me.p.DescripcionProducto.ToString
        End If
    End Sub

    Private Sub ButtonAñadirProd_Click(sender As Object, e As EventArgs) Handles ButtonAñadirProd.Click
        If Me.TextBoxDescripcionProd.Text <> String.Empty Then
            p = New Producto()
            p.DescripcionProducto = TextBoxDescripcionProd.Text
            Try
                If p.InsertarProducto() <> 1 Then
                    MessageBox.Show("INSERT return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.p.PDAO.LeerPorDescrip(p)
            Me.ListBoxIDProdProd.Items.Add(p.IDProducto)
            Me.ListBoxIDProdDisp.Items.Add(p.IDProducto)
            Me.ListBoxProdDispCot.Items.Add(p.IDProducto)
            Me.ListBoxDescrpProd.Items.Add(p.DescripcionProducto)
            Me.ButtonLimpiarProd.PerformClick()
        End If
    End Sub

    Private Sub ButtonActualizarProd_Click(sender As Object, e As EventArgs) Handles ButtonActualizarProd.Click
        If Not p Is Nothing Then
            p.DescripcionProducto = TextBoxDescripcionProd.Text
            Try
                If p.ActualizarProducto() <> 1 Then
                    MessageBox.Show("UPDATE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            MessageBox.Show("Descipcion= " & p.DescripcionProducto & " actualizado correctamente!!")
            Me.ButtonLimpiarProd.PerformClick()
        End If
    End Sub

    Private Sub ButtonEliminarProd_Click(sender As Object, e As EventArgs) Handles ButtonEliminarProd.Click
        If Not p Is Nothing Then
            If MessageBox.Show("¿Estás seguro que quieres borrar Producto= " & Me.p.IDProducto & "?", "Por favor, confirme", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    If p.BorrarProducto() <> 1 Then
                        MessageBox.Show("DELETE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End Try
                Me.ListBoxIDProdProd.Items.Remove(p.IDProducto)
                Me.ListBoxIDProdDisp.Items.Remove(p.IDProducto)
                Me.ListBoxProdDispCot.Items.Remove(p.IDProducto)
                Me.ListBoxDescrpProd.Items.Remove(p.DescripcionProducto)
            End If
            Me.ButtonLimpiarProd.PerformClick()
        End If
    End Sub

    Private Sub ButtonLimpiarProd_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarProd.Click
        Me.TextBoxIDProductoProd.Text = String.Empty
        Me.TextBoxDescripcionProd.Text = String.Empty
        Me.ButtonEliminarProd.Enabled = False
        Me.ButtonActualizarProd.Enabled = False
        Me.TextBoxIDProductoProd.ReadOnly = True
        Me.ListBoxIDProdProd.SelectedItems.Clear()
    End Sub
    'GESTIONES VIAJES'
    Private Sub ListBoxFechaViajes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxFechaViajes.SelectedIndexChanged
        If Not Me.ListBoxFechaViajes.SelectedItem Is Nothing Then
            Me.ListBoxFechaViajes.Enabled = False
            Me.ListBoxIDTren.Enabled = True
        End If
    End Sub

    Private Sub ListBoxIDTren_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxIDTren.SelectedIndexChanged
        If Not Me.ListBoxFechaViajes.SelectedItem Is Nothing Then
            Me.ListBoxIDTren.Enabled = False
            Me.ListBoxIDProd.Enabled = True
        End If
    End Sub

    Private Sub ListBoxIDProd_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxIDProd.SelectedIndexChanged
        If Not Me.ListBoxIDProd.SelectedItem Is Nothing Then
            Me.ListBoxIDProd.Enabled = False
            ButtonActualizarViaje.Enabled = True
            ButtonEliminarViaje.Enabled = True
            Me.DateTimePickerFechaViaje.Enabled = False
            Me.v = New Viaje(Convert.ToDateTime(Me.ListBoxFechaViajes.SelectedItem), Me.ListBoxIDTren.SelectedItem.ToString, Convert.ToInt64(Me.ListBoxIDProd.SelectedItem.ToString))
            Try
                Me.v.LeerViaje()
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.DateTimePickerFechaViaje.Text = Me.v.FechaViaje.ToString
            Me.TextBoxMatriculaTrenViaje.Text = Me.v.Tren.ToString
            Me.TextBoxIDProductoViaje.Text = Me.v.Producto.ToString
            Me.TextBoxToneladasTransViajes.Text = Me.v.ToneladasTransportadas.ToString
        End If
    End Sub

    Private Sub ListBoxMatTrenDispViajes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxMatTrenDispViajes.SelectedIndexChanged
        If Not Me.ListBoxMatTrenDispViajes.SelectedItem Is Nothing Then
            ButtonAñadirViaje.Enabled = True
            Me.TextBoxMatriculaTrenViaje.Text = Me.ListBoxMatTrenDispViajes.SelectedItem
        End If
    End Sub

    Private Sub ListBoxIDProdDisp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxIDProdDisp.SelectedIndexChanged
        If Not Me.ListBoxIDProdDisp.SelectedItem Is Nothing Then
            ButtonAñadirViaje.Enabled = True
            Me.TextBoxIDProductoViaje.Text = Me.ListBoxIDProdDisp.SelectedItem
        End If
    End Sub

    Private Sub ButtonAñadirViaje_Click(sender As Object, e As EventArgs) Handles ButtonAñadirViaje.Click
        If Me.TextBoxMatriculaTrenViaje.Text <> String.Empty And Me.TextBoxIDProductoViaje.Text <> String.Empty And Me.TextBoxToneladasTransViajes.Text <> String.Empty Then
            v = New Viaje()
            Dim vaux As Viaje
            Dim booleanaux As Boolean = True
            v.FechaViaje = DateTimePickerFechaViaje.Value.Date.ToString("d")
            v.Tren = TextBoxMatriculaTrenViaje.Text
            v.Producto = TextBoxIDProductoViaje.Text
            v.ToneladasTransportadas = TextBoxToneladasTransViajes.Text
            Me.v.LeerTodosViajes(Ruta.FileName)
            For Each vaux In Me.v.VDAO.Viajes
                If v.FechaViaje.Date = vaux.FechaViaje.Date And v.Tren = vaux.Tren Then
                    booleanaux = False
                End If
            Next
            t = New Tren(v.Tren)
            Me.t.TDAO.Leer(t)
            tt = New TipoTren(t.TipoTren)
            Me.tt.TTDAO.Leer(tt)
            If v.ToneladasTransportadas < tt.CapacidadMaxima Then
                If booleanaux Then
                    Try
                        If v.InsertarViaje() <> 1 Then
                            MessageBox.Show("INSERT return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Exit Sub
                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End Try
                    Me.ListBoxFechaViajes.Items.Add(v.FechaViaje)
                    Me.ListBoxIDTren.Items.Add(v.Tren)
                    Me.ListBoxIDProd.Items.Add(v.Producto)
                    Me.ListBoxTonTrans.Items.Add(v.ToneladasTransportadas)
                    Me.ButtonLimpiarViaje.PerformClick()
                Else
                    MessageBox.Show("El tren ya tiene programado un viaje ese dia", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            Else
                MessageBox.Show("Las toneladas transportadas exceden la capacidad maxima del tren", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        End If
    End Sub

    Private Sub ButtonActualizarViaje_Click(sender As Object, e As EventArgs) Handles ButtonActualizarViaje.Click
        If Not v Is Nothing Then
            v.ToneladasTransportadas = TextBoxToneladasTransViajes.Text
            Try
                If v.ActualizarViaje() <> 1 Then
                    MessageBox.Show("UPDATE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            MessageBox.Show("Toneladas transportadas= " & v.ToneladasTransportadas & " actualizado correctamente!!")
            Me.ButtonLimpiarViaje.PerformClick()
        End If
    End Sub

    Private Sub ButtonEliminarViaje_Click(sender As Object, e As EventArgs) Handles ButtonEliminarViaje.Click
        If Not v Is Nothing Then
            If MessageBox.Show("¿Estás seguro que quieres borrar Fecha= " & Me.v.FechaViaje & " Tren= " & Me.v.Tren & " Producto= " & Me.v.Producto & "?", "Por favor, confirme", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    If v.BorrarViaje() <> 1 Then
                        MessageBox.Show("DELETE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End Try
                Me.ListBoxFechaViajes.Items.Remove(v.FechaViaje)
                Me.ListBoxIDTren.Items.Remove(v.Tren)
                Me.ListBoxIDProd.Items.Remove(v.Producto)
                Me.ListBoxTonTrans.Items.Remove(v.ToneladasTransportadas)
            End If
            Me.ButtonLimpiarViaje.PerformClick()
        End If
    End Sub

    Private Sub ButtonLimpiarViaje_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarViaje.Click
        Me.TextBoxMatriculaTrenViaje.Text = String.Empty
        Me.TextBoxIDProductoViaje.Text = String.Empty
        Me.TextBoxToneladasTransViajes.Text = String.Empty
        Me.ButtonEliminarViaje.Enabled = False
        Me.ButtonActualizarViaje.Enabled = False
        Me.ButtonAñadirViaje.Enabled = False
        Me.DateTimePickerFechaViaje.Enabled = True
        Me.ListBoxFechaViajes.SelectedItems.Clear()
        Me.ListBoxIDTren.SelectedItems.Clear()
        Me.ListBoxIDProd.SelectedItems.Clear()
        Me.ListBoxMatTrenDispViajes.SelectedItems.Clear()
        Me.ListBoxIDProdDisp.SelectedItems.Clear()
        Me.ListBoxFechaViajes.Enabled = True
        Me.ListBoxIDTren.Enabled = False
        Me.ListBoxIDProd.Enabled = False
    End Sub
    ''CONSULTAS''
    'CONSULTA TREN'
    Private Sub ListBoxTrenesDispConsultas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxTrenesDispConsultas.SelectedIndexChanged
        If Not Me.ListBoxTrenesDispConsultas.SelectedItem Is Nothing Then
            Me.TextBoxTrenElegidoConsu.Text = Me.ListBoxTrenesDispConsultas.SelectedItem
            Me.ButtonRealizarConsultaTrFechas.Enabled = True
        End If
    End Sub

    Private Sub ButtonLimpiarConsultaTrFech_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarConsultaTrFech.Click
        Me.TextBoxTrenElegidoConsu.Text = String.Empty
        Me.ListBoxTrenesDispConsultas.SelectedItems.Clear()
        Me.ButtonRealizarConsultaTrFechas.Enabled = False
        Me.ListBoxProductosTransConsultas.Items.Clear()
        Me.ListBoxIDProdTrYFechCons.Items.Clear()
        Me.TextBoxNumeroDeViajes.Text = String.Empty
    End Sub

    Private Sub ButtonRealizarConsultaTrFechas_Click(sender As Object, e As EventArgs) Handles ButtonRealizarConsultaTrFechas.Click
        Dim paux As Producto
        Dim contador As Integer = 0
        Me.p = New Producto
        Me.ListBoxProductosTransConsultas.Items.Clear()
        Me.ListBoxIDProdTrYFechCons.Items.Clear()
        Try
            Me.p.PDAO.TrenyFechas(Me.DateTimePickerPrimeraFechaCons.Value.Date, Me.DateTimePickerSegundaFechaCons.Value.Date, Me.TextBoxTrenElegidoConsu.Text)
            For Each paux In Me.p.PDAO.ListaTrenyProd
                contador = contador + 1
                Me.ListBoxIDProdTrYFechCons.Items.Add(paux.IDProducto)
                Me.ListBoxProductosTransConsultas.Items.Add(paux.DescripcionProducto)
            Next
            Me.TextBoxNumeroDeViajes.Text = contador
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub
    'CONSULTA RANKING TRENES'
    Private Sub ButtonRealizarConsultaRankTr_Click(sender As Object, e As EventArgs) Handles ButtonRealizarConsultaRankTr.Click
        Dim ttaux As TipoTren
        Me.tt = New TipoTren
        Try
            Me.tt.TTDAO.RankingTipoTren(Me.DateTimePickerPrimFechaRankingTrenes.Value.Date, Me.DateTimePickerSegundaFechaRankingTrenes.Value.Date)
            For Each ttaux In Me.tt.TTDAO.TTFechas
                Me.ListBoxIDTTRanking.Items.Add(ttaux.IDTipoTren)
                Me.ListBoxDescripcionRanking.Items.Add(ttaux.Descripcion)
                Me.ListBoxCapacidadMaximaRanking.Items.Add(ttaux.CapacidadMaxima)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonLimpiarRankTr_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarRankTr.Click
        Me.ListBoxIDTTRanking.Items.Clear()
        Me.ListBoxDescripcionRanking.Items.Clear()
        Me.ListBoxCapacidadMaximaRanking.Items.Clear()
    End Sub
    'CONSULTA RANKING PRODUCTOS'
    Private Sub ButtonRealizarConsultaRankingProd_Click(sender As Object, e As EventArgs) Handles ButtonRealizarConsultaRankingProd.Click
        Dim paux As Producto
        Me.p = New Producto
        Try
            Me.p.PDAO.RankingProducto(Me.DateTimePickerPrimeraFechaRankingProd.Value.Date, Me.DateTimePickerSegundaFechaRankingProd.Value.Date)
            For Each paux In Me.p.PDAO.ProductosFechas
                Me.ListBoxIDProdRanking.Items.Add(paux.IDProducto)
                Me.ListBoxDescripcionProdRanking.Items.Add(paux.DescripcionProducto)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonLimpiarRankingProd_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarRankingProd.Click
        Me.ListBoxIDProdRanking.Items.Clear()
        Me.ListBoxDescripcionProdRanking.Items.Clear()
    End Sub
    'CONSULTA MAXIMO BENEFICIO'
    Private Sub ButtonRealizarConsultaMaxViaje_Click(sender As Object, e As EventArgs) Handles ButtonRealizarConsultaMaxViaje.Click
        Dim aux As Collection
        v = New Viaje
        Try
            Me.v.VDAO.ViajeMaxBeneficio()
            aux = CType(v.VDAO.ViajeMayorBeneficio(1), Collection)
            Me.TextBoxFechaViajeMaxBenef.Text = Convert.ToDateTime(aux(1)).Date.ToString("d")
            Me.TextBoxTipoTrenMaxBenef.Text = aux(2).ToString
            Me.TextBoxDescrTipoTrenMaxBenef.Text = aux(3).ToString
            Me.TextBoxDescProdMaxBenef.Text = aux(4).ToString
            Me.TextBoxTonelTranspMaxBenef.Text = aux(5).ToString
            Me.TextBoxEurosTonelMaxBenef.Text = aux(6).ToString
            Me.TextBoxBenefTotalMaxBenef.Text = aux(7).ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonLimpiarMaxViaje_Click(sender As Object, e As EventArgs) Handles ButtonLimpiarMaxViaje.Click
        Me.TextBoxFechaViajeMaxBenef.Text = String.Empty
        Me.TextBoxDescrTipoTrenMaxBenef.Text = String.Empty
        Me.TextBoxTipoTrenMaxBenef.Text = String.Empty
        Me.TextBoxDescProdMaxBenef.Text = String.Empty
        Me.TextBoxTonelTranspMaxBenef.Text = String.Empty
        Me.TextBoxEurosTonelMaxBenef.Text = String.Empty
        Me.TextBoxBenefTotalMaxBenef.Text = String.Empty
    End Sub

End Class

