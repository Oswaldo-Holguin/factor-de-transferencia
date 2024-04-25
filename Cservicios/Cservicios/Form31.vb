Public Class Frm31OrdenProduccion
    Public i As Integer
    Public _Folio, _Donante, _NombreClinica, _Grupo As String
    Public _FechaExtraccion, _FechaCaducidad As Date
    Public msg As String
    Public Style As MsgBoxStyle
    Public _Busca As DataRow
    Public _Material As String
    Public _Renglon As DataGridViewRow
    Public _Cantidad As Double
    Public _PrecioUnitario, _Importe, _Suma, _Utilidad As Double
    Public _Centro As String
    Public _OrdenProduccion, _Referencia, _Producto, _DescripcionProducto As String
    Public _Entrega, _REcibe, _Identificacion, _Lote As String
    Public Checacontrol As Control

    Private Sub Form31_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
       

        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet21.Productos' Puede moverla o quitarla según sea necesario.
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet5.Materiales' Puede moverla o quitarla según sea necesario.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet25.Produccion' Puede moverla o quitarla según sea necesario.
        ' Me.ProduccionTableAdapter.Fill(Me.CServiciosDataSet25.Produccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet25.DetalleProduccion' Puede moverla o quitarla según sea necesario.
        Me.DetalleProduccionTableAdapter.Fill(Me.CServiciosDataSet25.DetalleProduccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet25.Consumos' Puede moverla o quitarla según sea necesario.
        Me.ConsumosTableAdapter.Fill(Me.CServiciosDataSet25.Consumos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet26.Produccion' Puede moverla o quitarla según sea necesario.
        Me.ProduccionTableAdapter1.Fill(Me.CServiciosDataSet26.Produccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet8.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet27.Consumos' Puede moverla o quitarla según sea necesario.
        Me.ConsumosTableAdapter1.Fill(Me.CServiciosDataSet27.Consumos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet28.CierreProduccion' Puede moverla o quitarla según sea necesario.
        '  Me.CierreProduccionTableAdapter.Fill(Me.CServiciosDataSet28.CierreProduccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet29.CierreProduccion' Puede moverla o quitarla según sea necesario.
        Me.CierreProduccionTableAdapter1.Fill(Me.CServiciosDataSet29.CierreProduccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet24.PaquetesLeucocitarios' Puede moverla o quitarla según sea necesario.
        Me.PaquetesLeucocitariosTableAdapter.Fill(Me.CServiciosDataSet24.PaquetesLeucocitarios)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet23.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet23.VentaFactor)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.Clientes' Puede moverla o quitarla según sea necesario.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet49.MovMat' Puede moverla o quitarla según sea necesario.
        Me.MovMatTableAdapter.Fill(Me.CServiciosDataSet49.MovMat)

        Dim _FechaProduccion As Date

        Panel1.Enabled = False

        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False

        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                Checacontrol.Text = ""
            End If
            If TypeOf Checacontrol Is ComboBox Then
                Checacontrol.Text = ""
            End If
        Next

        Panel2.Visible = False
        DataGridView2.Rows.Clear()
        DataGridView1.Rows.Clear()
        DataGridView3.Rows.Clear()
        DataGridView4.Rows.Clear()
        ToolStripTextBox1.Text = ""

        TabControl2.SelectedIndex = 0

        For Each Checacontrol In Panel2.Controls
            If TypeOf Checacontrol Is TextBox Then
                Checacontrol.Text = ""
            End If
        Next

        ' Presenta las Ordenes de Producción
        For Me.i = 0 To CServiciosDataSet26.Produccion.Rows.Count - 1
            ' MsgBox("Valor de i " & CStr(i))
            _OrdenProduccion = CServiciosDataSet26.Produccion.Rows(i).Item("OrdenProduccion")
            _FechaProduccion = CServiciosDataSet26.Produccion.Rows(i).Item("FechaInicioProduccion")
            _Cantidad = CServiciosDataSet26.Produccion.Rows(i).Item("Cantidad")
            _Lote = CServiciosDataSet26.Produccion.Rows(i).Item("Lote")
            _Producto = CServiciosDataSet26.Produccion.Rows(i).Item("Producto")

            DataGridView1.Rows.Add(_OrdenProduccion, _FechaProduccion, _Cantidad, _Lote, _Producto)

        Next

        ' Muestra los Paquetes Leucocitarios utilizados en esta Orden de Producción
        ToolStripButton15_Click(sender, e)

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True
        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                Checacontrol.Text = ""
            End If
        Next
        TextBox1.Focus()
        TextBox9.Text = "ALTA"
        ToolStripButton2.Enabled = False
        ToolStripButton4.Enabled = True
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        DataGridView4.Rows.Clear()
        TextBox26.Text = "ACTIVA"
        TextBox26.Enabled = False

    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.Text = UCase(TextBox1.Text)

    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox2.Text = Today
        MonthCalendar1.Visible = True

    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox2.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False

    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox2.GotFocus
        ComboBox2.Items.Clear()
        For Me.i = 0 To CServiciosDataSet8.Centros.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet8.Centros.Rows(i).Item("Descripcion"))
        Next
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet8.Centros.Rows.Count - 1
            If CServiciosDataSet8.Centros.Rows(i).Item("Descripcion") = ComboBox2.Text Then
                TextBox4.Text = CServiciosDataSet8.Centros.Rows(i).Item("Centro")
            End If
        Next
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click


        Me.Close()

    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.GotFocus
        For Me.i = 0 To CServiciosDataSet21.Productos.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet21.Productos.Rows(i).Item("Descripcion"))
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet21.Productos.Rows.Count - 1
            If CServiciosDataSet21.Productos.Rows(i).Item("Descripcion") = ComboBox1.Text Then
                TextBox8.Text = CServiciosDataSet21.Productos.Rows(i).Item("Producto")
            End If
        Next
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click

        Dim _Find As Boolean
        Dim _Rows_Op() As DataRow
        Dim _Rows_Ce() As DataRow
        Dim _Rows_Pr() As DataRow

        _OrdenProduccion = TextBox1.Text
        _Find = False
        ' Guarda la Orden de Producción
        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                If Checacontrol.Text = "" Then
                    _Find = True
                End If
            End If
        Next

        If _Find Then
            msg = "Hay casillas en blanco. Debe proporcionar toda la información"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If


        ' Valida que no haya registrado esa Orden
        If TextBox9.Text = "ALTA" Then
            _OrdenProduccion = TextBox1.Text
            _Rows_Op = CServiciosDataSet25.Produccion.Select("OrdenProduccion = " & "'" & _OrdenProduccion & "'")

            _Find = False
            For Me.i = 0 To _Rows_Op.GetUpperBound(0)
                _Find = True
            Next

            If _Find Then
                msg = "Esta Orden de Producción ya está registrada. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If
        End If


        ' Valida El Centro
        _Centro = TextBox4.Text
        _Rows_Ce = CServiciosDataSet8.Centros.Select("Centro = " & "'" & _Centro & "'")

        _Find = False
        For Me.i = 0 To _Rows_Ce.GetUpperBound(0)
            _Find = True
        Next

        If _Find = False Then
            msg = "El Centro no existe. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If


        ' Valida el Producto
        _Producto = TextBox8.Text
        _Rows_Pr = CServiciosDataSet21.Productos.Select("Producto = " & "'" & _Producto & "'")

        _Find = False
        For Me.i = 0 To _Rows_Pr.GetUpperBound(0)
            _Find = True
        Next

        If _Find = False Then
            msg = "El Producto no existe. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        If TextBox5.Text = 0 Then
            msg = "Debe registrar un Número de Artículos a producir"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If


        ' Da de Alta la Orden de Producción
        If TextBox9.Text = "ALTA" Then
            _Busca = CServiciosDataSet26.Produccion.NewRow
            _Busca.Item("OrdenProduccion") = _OrdenProduccion
        End If
        If TextBox9.Text = "CAMBIO" Then

            _Busca = CServiciosDataSet26.Produccion.FindByOrdenProduccion(_OrdenProduccion)
        End If

        _Cantidad = CInt(TextBox5.Text)
        _Busca.Item("FechaInicioProduccion") = CDate(TextBox2.Text)
        _Busca.Item("FechaFinalProduccion") = CDate(TextBox2.Text)
        _Busca.Item("HoraInicioProduccion") = "08:00"
        _Busca.Item("HoraFinalProduccion") = "15:00"
        _Busca.Item("Referencia") = TextBox12.Text
        _Busca.Item("Comentarios") = TextBox26.Text
        _Busca.Item("Cantidad") = _Cantidad
        _Busca.Item("Usuario") = LoginForm1.TextBox1.Text
        _Busca.Item("Centro") = TextBox4.Text
        _Busca.Item("Formula") = TextBox31.Text
        _Busca.Item("FechaRegistro") = Today
        _Busca.Item("HoraREgistro") = Mid(TimeOfDay, 1, 5)
        _Busca.Item("NumeroInicial") = TextBox6.Text
        _Busca.Item("NumeroFinal") = TextBox7.Text
        _Busca.Item("Entrega") = TextBox10.Text
        _Busca.Item("Recibe") = TextBox11.Text
        _Busca.Item("Identificacion") = TextBox13.Text
        _Busca.Item("Lote") = TextBox3.Text
        _Busca.Item("Producto") = TextBox8.Text
        ' _Busca.Item("Documento") = "0"


        If TextBox9.Text = "ALTA" Then
            CServiciosDataSet26.Produccion.Rows.Add(_Busca)
        End If

        Me.Validate()
        ProduccionBindingSource1.EndEdit()
        Me.TableAdapterManager4.UpdateAll(CServiciosDataSet26)
        Me.ProduccionTableAdapter1.Fill(Me.CServiciosDataSet26.Produccion)

        msg = "Registro guardado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(msg, Style)
        Button1_Click(sender, e)


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Rows.Count = 0 Then
            Exit Sub
        End If


        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(0).Value
        _Busca = CServiciosDataSet26.Produccion.FindByOrdenProduccion(TextBox1.Text)
        TextBox2.Text = _Busca.Item("FechaInicioProduccion")
        TextBox4.Text = _Busca.Item("Centro")
        TextBox5.Text = _Busca.Item("Cantidad")
        TextBox6.Text = _Busca.Item("NumeroInicial")
        TextBox7.Text = _Busca.Item("NumeroFinal")
        TextBox8.Text = _Busca.Item("Producto")
        TextBox10.Text = _Busca.Item("Entrega")
        TextBox11.Text = _Busca.Item("Recibe")
        TextBox12.Text = _Busca.Item("Referencia")
        TextBox13.Text = _Busca.Item("Identificacion")
        TextBox3.Text = _Busca.Item("Lote")
        If DBNull.Value.Equals(_Busca.Item("Comentarios")) Then
            TextBox26.Text = ""
        Else
            TextBox26.Text = _Busca.Item("Comentarios")
        End If

        If TextBox26.Text = "CERRADA" Then
            TextBox26.BackColor = Color.LightSalmon
            ToolStripButton3.Enabled = False
            ToolStripButton9.Enabled = False
            ToolStripButton12.Enabled = False
        Else
            ToolStripButton3.Enabled = True
            ToolStripButton9.Enabled = True
            ToolStripButton12.Enabled = True
            TextBox26.BackColor = Color.White
        End If

        _PrecioUnitario = CInt(_Busca.Item("Formula"))
        TextBox31.Text = Format(_PrecioUnitario, "$###,###.##")

        _Busca = CServiciosDataSet8.Centros.FindByCentro(TextBox4.Text)
        ComboBox2.Text = _Busca.Item("Descripcion")

        ' _Busca = CServiciosDataSet21.Productos.FindByProducto(TextBox8.Text)
        ' ComboBox1.Text = _Busca.Item("Descripcion")
        'TextBox31.Text = _Busca.Item("PrecioUnitario")


        ' Busca la venta
        _Lote = TextBox3.Text
        Dim _Rows_Vf() As DataRow = CServiciosDataSet23.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
        Dim _VentaLocal, _Envios As Integer
        Dim _PVentaLocal, _PEnvios As Integer

        _VentaLocal = 0
        _Envios = 0
        _PVentaLocal = 0
        _PEnvios = 0
        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
            If _Rows_Vf(i).Item("Paciente") = "ENVIO FORANEO" Then
                _Envios = _Envios + _Rows_Vf(i).Item("NumeroFrascos")
                _Importe = _Rows_Vf(i).Item("PrecioNeto")
                _PEnvios = _PEnvios + _Importe
            Else
                _VentaLocal = _VentaLocal + _Rows_Vf(i).Item("NumeroFrascos")
                _Importe = _Rows_Vf(i).Item("PrecioNeto")
                _PVentaLocal = _PVentaLocal + _Importe
            End If
        Next

        TextBox28.Text = _Envios
        TextBox30.Text = Format(_PEnvios, "$###,###.##")
        TextBox27.Text = _VentaLocal
        TextBox29.Text = Format(_PVentaLocal, "$###,###.##")


        DataGridView2.Rows.Clear()
        DataGridView4.Rows.Clear()

        ToolStripButton4.Enabled = False

        ToolStripButton11_Click(sender, e)
        Panel2.Visible = False
        DataGridView5.Rows.Clear()

        ' Presenta datos de cierre
        If TextBox26.Text = "CERRADA" Or TextBox26.Text = "ENTREGADO" Then
            _OrdenProduccion = TextBox1.Text
            _Busca = CServiciosDataSet29.CierreProduccion.FindByOrdenProduccion(_OrdenProduccion)
            Panel2.Visible = True
            Button2.Enabled = False
            TextBox14.Text = _Busca.Item("OrdenProduccion")
            TextBox15.Text = _Busca.Item("FechaOrden")
            TextBox16.Text = TextBox8.Text
            TextBox17.Text = _Busca.Item("Cantidad")
            TextBox18.Text = _Busca.Item("PrecioUnitario")
            _Importe = _Busca.Item("ImporteProduccion")
            _Utilidad = _Importe
            TextBox19.Text = Format(_Importe, "$###,###,##0.00")
            _Importe = _Busca.Item("CostoMateriales")
            _Utilidad = _Utilidad - _Importe
            TextBox20.Text = Format(_Importe, "$###,###,##0.00")
            _Importe = _Busca.Item("CostoManodeObra")
            _Utilidad = _Utilidad - _Importe
            TextBox21.Text = Format(_Importe, "$###,###,##0.00")
            _Importe = _Busca.Item("CostoIndirectos")
            TextBox22.Text = Format(_Importe, "$###,###,##0.00")
            _Utilidad = _Utilidad - _Importe
            TextBox23.Text = Format(_Utilidad, "$###,###,##0.00")

        End If




    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick



    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click

        If TextBox26.Text = "ENTREGADO" Then
            msg = "Esta Orden de Producción ya se ha entregado a Ventas. No se puede modificar"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True
        Panel1.Enabled = True
        TextBox1.Enabled = False

        TextBox9.Text = "CAMBIO"



    End Sub

    Private Sub ToolStripButton12_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton12.Click
        Frm32AsignaMateriales.Show()

    End Sub

    Private Sub ToolStripButton11_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton11.Click
        ' Presenta los Materiales utilizados en la Producción
        Me.ConsumosTableAdapter1.Fill(Me.CServiciosDataSet27.Consumos)

        _OrdenProduccion = TextBox1.Text
        Dim _Rows_Co() As DataRow = CServiciosDataSet27.Consumos.Select("OrdenProduccion = " & "'" & _OrdenProduccion & "'")
        Dim _DescripcionMatreial As String
        Dim _UMedida, _Presentacion As String
        Dim _PrecioUnitario As Double
        Dim _CostosDirectos As Integer


        _CostosDirectos = 0
        DataGridView2.Rows.Clear()
        For Me.i = 0 To _Rows_Co.GetUpperBound(0)
            _Material = _Rows_Co(i).Item("Material")
            _DescripcionMatreial = _Rows_Co(i).Item("DescripcionMaterial")
            _Cantidad = _Rows_Co(i).Item("Cantidad")
            _UMedida = _Rows_Co(i).Item("UMedida")
            _Presentacion = _Rows_Co(i).Item("Referencia")
            _PrecioUnitario = _Rows_Co(i).Item("CostoUnitario")
            _Importe = _Cantidad * _PrecioUnitario

            _CostosDirectos = _CostosDirectos + (_Cantidad * _PrecioUnitario)

            DataGridView2.Rows.Add(_OrdenProduccion, _Material, _DescripcionMatreial, _Cantidad, _UMedida, _Presentacion, _PrecioUnitario, _Importe)

        Next

        ToolStripTextBox1.Text = Format(_CostosDirectos, "$###,##0.00")


    End Sub

    Private Sub ToolStripButton9_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton9.Click
        If TextBox1.Text = "" Then
            msg = "Debe seleccionar una Orden de Producción para el Cierre"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        _OrdenProduccion = TextBox1.Text
        Dim _Rows_Co() As DataRow
        Dim _Find As Boolean

        _Rows_Co = CServiciosDataSet26.Produccion.Select("OrdenProduccion = " & "'" & _OrdenProduccion & "'")
        _Find = False
        For Me.i = 0 To _Rows_Co.GetUpperBound(0)
            _Find = True
        Next

        If _Find = False Then
            MsgBox("Orden " & _OrdenProduccion)
            msg = "Esa Orden de Producción no está registrada. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        _Rows_Co = CServiciosDataSet27.Consumos.Select("OrdenProduccion = " & "'" & _OrdenProduccion & "'")
        Dim title As String
        Dim response As MsgBoxResult
        msg = "Está a punto de Cerrera esta Orden de Producción. Desea continuar?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cierre de Órden de Producción"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Cierre de la Órden de Producción
            Panel2.Visible = True
            Panel1.Enabled = False
            DataGridView1.Enabled = False
            DataGridView2.Enabled = False
            TabControl2.SelectedIndex = 2
            TextBox21.Focus()
            TextBox14.Text = TextBox1.Text
            TextBox14.Enabled = False
            TextBox15.Text = TextBox2.Text
            TextBox15.Enabled = False
            TextBox16.Text = TextBox8.Text
            TextBox16.Enabled = False
            TextBox17.Text = TextBox5.Text
            TextBox17.Enabled = False

            _Producto = TextBox16.Text

            _Busca = CServiciosDataSet21.Productos.FindByProducto(_Producto)

            _PrecioUnitario = _Busca.Item("PrecioUnitario")
            TextBox18.Text = _PrecioUnitario
            _Importe = CInt(TextBox17.Text) * _PrecioUnitario
            TextBox19.Text = Format(_Importe, "$###,###,##0.00")
            TextBox25.Text = _Importe

            _Suma = 0
            For Me.i = 0 To _Rows_Co.GetUpperBound(0)
                _Suma = _Suma + (_Rows_Co(i).Item("CostoUnitario") * _Rows_Co(i).Item("Cantidad"))
            Next
            TextBox24.Text = _Suma
            TextBox20.Text = Format(_Suma, "$###,###,##0.00")
            TextBox21.Text = 0
            TextBox22.Text = 0
            _Utilidad = _Importe - _Suma - CInt(TextBox21.Text) - CInt(TextBox22.Text)
            TextBox23.Text = Format(_Utilidad, "$###,###,##0.00")

           


        Else
            Exit Sub
        End If
    End Sub

    Private Sub TextBox21_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox21.LostFocus
        _Utilidad = _Importe - _Suma - CInt(TextBox21.Text) - CInt(TextBox22.Text)
        TextBox23.Text = Format(_Utilidad, "$###,###,##0.00")
    End Sub

    Private Sub TextBox21_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox21.TextChanged

    End Sub

    Private Sub TextBox22_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox22.LostFocus
        _Utilidad = _Importe - _Suma - CInt(TextBox21.Text) - CInt(TextBox22.Text)
        TextBox23.Text = Format(_Utilidad, "$###,###,##0.00")
    End Sub

    Private Sub TextBox22_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ' Guarda el Cierre
        Dim _FechaOrden As Date = CDate(TextBox2.Text)
        Dim _Find As Boolean
        Dim _CostoMateriales As Integer = CInt(TextBox24.Text)
        Dim _CostoManodeObra As Integer = CInt(TextBox21.Text)

        _Find = False
        For Each Me.Checacontrol In Panel2.Controls
            If TypeOf Checacontrol Is TextBox Then
                If Checacontrol.Text = "" Then
                    _Find = True
                End If
            End If
        Next

        If _Find Then
            msg = "Debe completar el procedimiento para el Cierre. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If


        _Importe = CInt(TextBox25.Text)
        _PrecioUnitario = CInt(TextBox18.Text)
        _OrdenProduccion = TextBox1.Text
        _Cantidad = CInt(TextBox5.Text)

        _Busca = CServiciosDataSet26.Produccion.FindByOrdenProduccion(_OrdenProduccion)
        _Busca.Item("Comentarios") = "ENTREGADO"

        Me.Validate()
        ProduccionBindingSource.EndEdit()
        Me.TableAdapterManager4.UpdateAll(CServiciosDataSet26)
        Me.ProduccionTableAdapter1.Fill(Me.CServiciosDataSet26.Produccion)

        ' Registra el Cierre 

        _Busca = CServiciosDataSet29.CierreProduccion.NewRow
        _Busca.Item("OrdenProduccion") = _OrdenProduccion
        _Busca.Item("FechaOrden") = _FechaOrden
        _Busca.Item("Cantidad") = _Cantidad
        _Busca.Item("PrecioUnitario") = _PrecioUnitario
        _Busca.Item("ImporteProduccion") = _Importe
        _Busca.Item("CostoMateriales") = _CostoMateriales
        _Busca.Item("CostoManodeObra") = _CostoManodeObra
        _Busca.Item("FechaCierre") = CDate(TextBox15.Text)
        _Busca.Item("Lote") = TextBox3.Text
        _Busca.Item("CostoIndirectos") = CInt(TextBox22.Text)
        _Busca.Item("Referencia") = CStr(Today)
        _Busca.Item("FechaRegistro") = Today
        _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)

        CServiciosDataSet29.CierreProduccion.Rows.Add(_Busca)

        Me.Validate()
        CierreProduccionBindingSource1.EndEdit()
        Me.TableAdapterManager7.UpdateAll(CServiciosDataSet29)

        Dim _BuscaMM As DataRow
        Dim _Umedida As String
        Dim _Usuario As String = LoginForm1.TextBox6.Text
        _Centro = LoginForm1.TextBox1.Text

        ' Registra la salida de material
        For Me.i = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(i).Cells(0).Value = "" Then
                Exit For
            End If


            _Material = DataGridView2.Rows(i).Cells(1).Value
            _Cantidad = DataGridView2.Rows(i).Cells(3).Value
            _PrecioUnitario = DataGridView2.Rows(i).Cells(6).Value
            _Lote = DataGridView2.Rows(i).Cells(0).Value
            _Umedida = DataGridView2.Rows(i).Cells(4).Value


            '  MsgBox("Valor de I " & CStr(i))
            '  MsgBox("Material " & _Material)
            '  MsgBox("Cantidad " & CStr(_Cantidad))


            _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_Material)
            _Busca.Item("Existencia") = _Busca.Item("Existencia") - _Cantidad

            _BuscaMM = CServiciosDataSet49.MovMat.NewRow
            _BuscaMM.Item("Material") = _Material
            _BuscaMM.Item("TipoMov") = "S"
            _BuscaMM.Item("FechaMov") = CStr(Today.Date)
            _BuscaMM.Item("Horamov") = Mid(TimeOfDay, 1, 10)
            _BuscaMM.Item("Cantidad") = _Cantidad
            _BuscaMM.Item("PrecioUnitario") = _PrecioUnitario
            _BuscaMM.Item("Proveedor") = ""
            _BuscaMM.Item("Documento") = _Lote
            _BuscaMM.Item("Factura") = _Lote
            _BuscaMM.Item("Referencia") = "PRODUCCION DE FACTOR DE TRANSFERENCIA"
            _BuscaMM.Item("Comentarios") = "MATERIALES UTILIZADOS EN LA PRODUCCION DE FACTOR DE TRANSFERENCIA DEL LOTE " & _Lote
            _BuscaMM.Item("Usuario") = _Usuario
            _BuscaMM.Item("Centro") = _Centro
            _BuscaMM.Item("UMedida") = _Umedida

            CServiciosDataSet49.MovMat.Rows.Add(_BuscaMM)

        Next

        Me.Validate()
        MaterialesBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet5)
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)

        Me.Validate()
        MovMatBindingSource.EndEdit()
        TableAdapterManager11.UpdateAll(CServiciosDataSet49)
        Me.MovMatTableAdapter.Fill(Me.CServiciosDataSet49.MovMat)


        msg = "Orden de Producción CERRADA correctamente"
        Style = MsgBoxStyle.Information

        MsgBox(msg, Style)
        Me.Close()



    End Sub

    Private Sub TextBox26_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox26.GotFocus

    End Sub

    Private Sub TextBox26_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox26.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = TextBox1.Text
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ToolStripButton10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton10.Click
        ' Genera el Recibo Entrega-Recepción
        If TextBox1.Text = "" Then
            msg = "Debe Seleccionar una Orden de Producción para generar este Documento"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If
        If TextBox26.Text <> "ENTREGADO" Then
            msg = "Esta orden no se ha cerrado. Debe de efectuar el Cierre antes de generar este Documento"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Recibo Entrega Recepcion.xlsx"
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de excel

        Dim _Folio As String = Mid(TextBox1.Text, 1, 2)
        Dim _FechaProduccion = CDate(TextBox2.Text)
        Dim _FechaEntrega As Date = Today
        _Lote = TextBox1.Text
        _Cantidad = CInt(TextBox5.Text)
        _Producto = ComboBox1.Text

        m_Excel.Worksheets("RECIBO").cells(10, 3).value = _Lote
        m_Excel.Worksheets("RECIBO").cells(10, 4).value = _FechaProduccion
        m_Excel.Worksheets("RECIBO").cells(10, 5).value = _Cantidad
        m_Excel.Worksheets("RECIBO").cells(10, 6).value = _Producto

        m_Excel.Worksheets("RECIBO").cells(10, 9).value = _FechaEntrega
        m_Excel.Worksheets("RECIBO").cells(18, 3).value = TextBox10.Text
        m_Excel.Worksheets("RECIBO").cells(18, 7).value = TextBox11.Text

        m_Excel.Visible = True

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton20_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton20.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton13_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton13.Click
        If TextBox1.Text = "" Then
            msg = "Debe Seleccionar una Orden de Producción para asignar Paquetes Leucocitarios"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If
        Frm36PLOP.Show()


    End Sub

    Private Sub ToolStripButton15_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton15.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet24.PaquetesLeucocitarios' Puede moverla o quitarla según sea necesario.
        Me.PaquetesLeucocitariosTableAdapter.Fill(Me.CServiciosDataSet24.PaquetesLeucocitarios)

        _OrdenProduccion = TextBox1.Text
        Dim _Rows_PQ() As DataRow = CServiciosDataSet24.PaquetesLeucocitarios.Select("Documento = " & "'" & _OrdenProduccion & "'")
        DataGridView6.Rows.Clear()

        For Me.i = 0 To _Rows_PQ.GetUpperBound(0)
            _Folio = _Rows_PQ(i).Item("Folio")
            _Donante = _Rows_PQ(i).Item("Donante")
            _FechaExtraccion = _Rows_PQ(i).Item("FechaExtraccion")
            _FechaCaducidad = _Rows_PQ(i).Item("FechaCaducidad")
            _Grupo = _Rows_PQ(i).Item("Grupo")
            _NombreClinica = _Rows_PQ(i).Item("NombreClinica")
            DataGridView6.Rows.Add(_Folio, _Donante, _FechaExtraccion, _FechaCaducidad, _Grupo, _NombreClinica)
        Next

    End Sub

    Private Sub ToolStripButton14_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton14.Click
        ' Genera la Estadística

        If TextBox1.Text = "" Then
            msg = "Debe seleccionar un Lote para poder generar la Estadística"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Estadistica.xlsx"
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de excel

        _Lote = TextBox1.Text
        Dim _Linea, _Columna As Integer
        Dim _Rows_VF() As DataRow = CServiciosDataSet23.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
        Dim _Cliente As String
        Dim _NumeroFrascos As Integer

        _Linea = 10
        _Columna = 4
        For Me.i = 0 To CServiciosDataSet30.Clientes.Rows.Count - 1
            m_Excel.Worksheets("CLTE").cells(_Linea, 2).value = CServiciosDataSet30.Clientes.Rows(i).Item("Cliente")
            m_Excel.Worksheets("CLTE").cells(_Linea, 3).value = CServiciosDataSet30.Clientes.Rows(i).Item("NombreCliente")
            _Cliente = CServiciosDataSet30.Clientes.Rows(i).Item("Cliente")

            m_Excel.Worksheets("CLTE").cells(2, _Columna).value = _Cliente

            _NumeroFrascos = 0
            For X = 0 To _Rows_VF.GetUpperBound(0)
                If _Rows_VF(X).Item("Caja") = _Cliente Then
                    _NumeroFrascos = _NumeroFrascos + _Rows_VF(i).Item("NumeroFrascos")
                End If
            Next
            m_Excel.Worksheets("CLTE").cells(_Linea, 4).value = _NumeroFrascos
            _Linea = _Linea + 1

            m_Excel.Worksheets("CLTE").CELLS(3, _Columna).VALUE = _NumeroFrascos
            _Columna = _Columna + 1
        Next

        m_Excel.Visible = True
        ' m_Excel.Workbooks.Close()
    End Sub
End Class