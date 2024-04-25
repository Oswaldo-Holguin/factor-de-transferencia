Public Class Frm56Inventarios
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Busca As DataRow
    Public _Checacontrol As Control
    Public _Material, _Centro, _Descripcion, _Presentacion, _UMedida, _Referencia As String
    Public _Maximo, _Minimo, _PuntoReorden, _Existencia, _PrecioUnitario, _Importe As Double

    Private Sub Frm56Inventarios_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Escape) Then
            _Renglon = DataGridView1.CurrentRow
            _Renglon.DefaultCellStyle.BackColor = Color.White
            Button4_Click(sender, e)
            ToolStripButton2_Click(sender, e)
            TabPage1.Parent = TabControl1
            TabPage4.Parent = Nothing
        End If


    End Sub


    Private Sub Frm56Inventarios_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load




        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet5.Materiales' Puede moverla o quitarla según sea necesario.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet47.Ajustes' Puede moverla o quitarla según sea necesario.
        Me.AjustesTableAdapter.Fill(Me.CServiciosDataSet47.Ajustes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet49.MovMat' Puede moverla o quitarla según sea necesario.
        Me.MovMatTableAdapter.Fill(Me.CServiciosDataSet49.MovMat)
        'TODO: This line of code loads data into the 'CServiciosDataSet56.Proveedores' table. You can move, or remove it, as needed.
        Me.ProveedoresTableAdapter.Fill(Me.CServiciosDataSet56.Proveedores)
        'TODO: This line of code loads data into the 'CServiciosDataSet56.DetalleRecepciones' table. You can move, or remove it, as needed.
        Me.DetalleRecepcionesTableAdapter.Fill(Me.CServiciosDataSet56.DetalleRecepciones)
        'TODO: This line of code loads data into the 'CServiciosDataSet56.Consumos' table. You can move, or remove it, as needed.
        Me.ConsumosTableAdapter.Fill(Me.CServiciosDataSet56.Consumos)
        'TODO: This line of code loads data into the 'CServiciosDataSet46.Recepciones' table. You can move, or remove it, as needed.
        Me.RecepcionesTableAdapter.Fill(Me.CServiciosDataSet46.Recepciones)

        For Me.I = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            ToolStripComboBox1.Items.Add(CServiciosDataSet.Centros.Rows(I).Item("Descripcion"))
        Next

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Visible = False
            End If
            If TypeOf _Checacontrol Is Label Then
                _Checacontrol.Visible = False
            End If
        Next

        For Each Me._Checacontrol In Me.TabPage2.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        ComboBox1.Items.Clear()
        TabControl1.SelectedIndex = 0
        ComboBox1.Text = ""
        RichTextBox1.Text = ""
        TabPage2.Parent = Nothing
        TabPage3.Parent = Nothing
        TabPage4.Parent = Nothing

        Button6.Enabled = False
        Button2.Enabled = True
        Button5.Enabled = True
        ' ToolStripButton2_Click(sender, e)



    End Sub

    Private Sub ToolStripComboBox1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox1.Click

        If ToolStripComboBox1.Items.Count = 0 Then
            Msg = "Presione el Boton de MOSTRAR INFORMACION"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        ToolStripButton2_Click(sender, e)
        
    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim _Renglon As DataGridViewRow
        _Renglon = DataGridView1.CurrentRow

        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Visible = True
            End If
        Next
        _Existencia = _Renglon.Cells(5).Value
        _PrecioUnitario = _Renglon.Cells(6).Value
        _UMedida = _Renglon.Cells(10).Value
        TextBox1.Text = _Renglon.Cells(0).Value
        TextBox2.Text = _Renglon.Cells(1).Value
        TextBox3.Text = _Renglon.Cells(5).Value

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TabPage1.Parent = TabControl1
        TabPage2.Parent = Nothing
        TabPage3.Parent = Nothing

        Button6.Enabled = False
        Button5.Enabled = True
        Button2.Enabled = True
        ToolStripButton2_Click(sender, e)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar un material para poder registrar un ajuste al inventario"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        ' Muestra la página para registrar el ajuste

        TabPage2.Parent = TabControl1
        TabPage1.Parent = Nothing
        TabPage3.Parent = Nothing

        Button2.Enabled = False
        Button5.Enabled = True
        Button6.Enabled = True

        TabControl1.SelectedIndex = 1

        TextBox4.Text = TextBox1.Text
        TextBox5.Text = TextBox2.Text
        TextBox4.Enabled = False
        TextBox5.Enabled = False

        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("AJUSTE POR ENTRADA DE MATERIAL")
        ComboBox1.Items.Add("AJUSTE POR SALIDA DE MATERIAL")



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim _Index As Integer = ComboBox1.SelectedIndex

        If _Index = 0 Then
            TextBox6.Text = "E"
        Else
            TextBox6.Text = "S"
        End If
        TextBox7.Focus()
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click


        Dim title As String
        Dim response As MsgBoxResult

        Msg = "Está seguro de ejecutar este Ajustea al Inventario?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Ajustes al Inventario de Materiales"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Verifica que no haya casillas en blanco
            For Each Me._Checacontrol In TabPage2.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    If _Checacontrol.Text = "" Then
                        Msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
                        Style = MsgBoxStyle.Information
                        MsgBox(Msg, MsgBoxStyle.AbortRetryIgnore)
                        Exit Sub
                    End If
                End If
            Next
            If CDbl(TextBox7.Text) = 0 Then
                Msg = "La cantidad a Ajustar no puede ser menor a 1 (uno)"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            ' Verifica que el ajuste de salida no sobrepase la existencia
            _Material = TextBox4.Text

            If TextBox6.Text = "S" Then
                '  MsgBox("Existencia, _Ajuste " & CStr(_Existencia) & " " & TextBox7.Text)

                If CDbl(TextBox7.Text) > _Existencia Then
                    Msg = "La cantidad a ajustar como Salida, no puede ser mayor a la exsitencia actual"
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                    Exit Sub
                End If

            End If

            ' Registra el movimiento

            Dim _Comentarios As String = RichTextBox1.Text
            If _Comentarios = "" Then
                _Comentarios = "SIN COMENTARIOS"
            End If
            Dim _Usuario As String = LoginForm1.TextBox6.Text
            _Centro = LoginForm1.TextBox1.Text

            _Busca = CServiciosDataSet49.MovMat.NewRow
            _Busca.Item("Material") = TextBox4.Text
            _Busca.Item("TipoMov") = TextBox6.Text
            _Busca.Item("FechaMov") = CStr(Today.Date)
            _Busca.Item("Horamov") = Mid(TimeOfDay, 1, 10)
            _Busca.Item("Cantidad") = CDbl(TextBox7.Text)
            _Busca.Item("PrecioUnitario") = _PrecioUnitario
            _Busca.Item("Proveedor") = ""
            _Busca.Item("Documento") = TextBox8.Text
            _Busca.Item("Factura") = ""
            _Busca.Item("Referencia") = TextBox9.Text
            _Busca.Item("Comentarios") = _Comentarios
            _Busca.Item("Usuario") = _Usuario
            _Busca.Item("Centro") = _Centro
            _Busca.Item("UMedida") = _UMedida

            CServiciosDataSet49.MovMat.Rows.Add(_Busca)

            Me.Validate()
            MovMatBindingSource.EndEdit()
            TableAdapterManager3.UpdateAll(CServiciosDataSet49)
            Me.MovMatTableAdapter.Fill(Me.CServiciosDataSet49.MovMat)

            ' Actualiza la Existencia
            _Material = TextBox4.Text

            _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_Material)
            Dim _Cantidad As Double = CDbl(TextBox7.Text)
            If TextBox6.Text = "E" Then
                _Busca.Item("Existencia") = _Busca.Item("Existencia") + _Cantidad
            Else
                _Busca.Item("Existencia") = _Busca.Item("Existencia") - _Cantidad
            End If

            Me.Validate()
            MaterialesBindingSource.EndEdit()
            TableAdapterManager2.UpdateAll(CServiciosDataSet5)
            Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)

            Msg = "Ajuste guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)

            Button1_Click(sender, e)
            ToolStripButton2_Click(sender, e)

        Else
            ' Perform some other action.
        End If

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click

        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        Dim _Index As Integer = ToolStripComboBox1.SelectedIndex
        Dim _CuantosArticulos As Integer
        Dim _ValorInventario As Double

        _CuantosArticulos = 0
        _ValorInventario = 0


        _Centro = CServiciosDataSet.Centros.Rows(_Index).Item("Centro")

        Dim _Rows_Ma() As DataRow = CServiciosDataSet5.Materiales.Select("Centro = " & "'" & _Centro & "'")

        DataGridView1.Rows.Clear()
        For Me.I = 0 To _Rows_Ma.GetUpperBound(0)
            _Material = _Rows_Ma(I).Item("Material")
            _Descripcion = _Rows_Ma(I).Item("Descripcion")
            _Minimo = _Rows_Ma(I).Item("Minimo")
            _Maximo = _Rows_Ma(I).Item("Maximo")
            _PuntoReorden = _Rows_Ma(I).Item("PuntoReorden")
            _Existencia = _Rows_Ma(I).Item("Existencia")
            _PrecioUnitario = _Rows_Ma(I).Item("PrecioUnitario")
            _Importe = _Existencia * _PrecioUnitario
            _Presentacion = _Rows_Ma(I).Item("Presentacion")
            _Referencia = _Rows_Ma(I).Item("Referencia")
            _UMedida = _Rows_Ma(I).Item("UMedida")

            DataGridView1.Rows.Add(_Material, _Descripcion, _Minimo, _Maximo, _PuntoReorden, _Existencia, _PrecioUnitario, _Importe, _Presentacion, _Referencia, _UMedida)
            _ValorInventario = _ValorInventario + _Importe
        Next

        Label8.Visible = True
        Label9.Visible = True
        TextBox10.Visible = True
        TextBox11.Visible = True
        _CuantosArticulos = DataGridView1.Rows.Count
        TextBox10.Text = _CuantosArticulos
        TextBox11.Text = Format(_ValorInventario, "$#,###,###.##")


    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        For Each Me._Checacontrol In TabPage2.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        ComboBox1.Text = ""
        RichTextBox1.Text = ""
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.Visible = False
            End If
        Next
        TabControl1.SelectedIndex = 0
        DataGridView2.Rows.Clear()
        DataGridView2.Visible = False
        Label7.Text = "."

    End Sub

    Private Sub DetalleDeMovimientosToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DetalleDeMovimientosToolStripMenuItem.Click
        ' Presenta el Detalle de movimientos de este material
        DataGridView2.Visible = True
        DataGridView2.Rows.Clear()

        _Material = TextBox1.Text

        TabPage1.Parent = Nothing
        TabPage4.Parent = TabControl1

        Dim _NombreProveedor As String
        Dim _Motivo As String
        Dim _Cantidad, _PrecioUnitario As Double
        Dim _FechaMov As Date
        Dim _TipoMov, _Umedida, _Proveedor, _Factura As String
        Dim _RowsMM() As DataRow = CServiciosDataSet49.MovMat.Select("Material = " & "'" & _Material & "'")
        Dim _Rows_Co() As DataRow = CServiciosDataSet56.Consumos.Select("Material = " & "'" & _Material & "'")
        Dim _Rows_Dr() As DataRow = CServiciosDataSet56.DetalleRecepciones.Select("MaterialRecibe = " & "'" & _Material & "'")
        Dim _Recepcion As String

        ' Presenta los Ajustes

        For Me.I = 0 To _RowsMM.GetUpperBound(0)
            _TipoMov = _RowsMM(I).Item("TipoMov")
            If _TipoMov = "E" Then
                _TipoMov = "ENTRADA"
            Else
                _TipoMov = "SALIDA"
            End If
            _Motivo = _TipoMov & " POR AJUSTE"
            _FechaMov = CDate(_RowsMM(I).Item("FechaMov"))
            _Cantidad = _RowsMM(I).Item("Cantidad")
            _PrecioUnitario = _RowsMM(I).Item("PrecioUnitario")
            _Umedida = _RowsMM(I).Item("UMedida")
            _Proveedor = _RowsMM(I).Item("Proveedor")
            _Factura = _RowsMM(I).Item("Factura")
            _Referencia = _RowsMM(I).Item("Referencia")

            DataGridView2.Rows.Add(_TipoMov, _FechaMov, _Cantidad, _PrecioUnitario, _Umedida, _Proveedor, _Factura, _Referencia, _Motivo)

            Label7.Text = "Precione ESC para regresar al listado de materiales"
            Label7.ForeColor = Color.Maroon
        Next

        ' Presenta las Salidas por Producción
        _Motivo = "SALIDA POR PRODUCCION"
        For Me.I = 0 To _Rows_Co.GetUpperBound(0)
            _TipoMov = "SALIDA"
            _FechaMov = _Rows_Co(I).Item("FechaOrden")
            _Cantidad = _Rows_Co(I).Item("Cantidad")
            _PrecioUnitario = _Rows_Co(I).Item("CostoUnitario")
            _Umedida = _Rows_Co(I).Item("Referencia")
            _Proveedor = ""
            _Factura = ""
            _Referencia = "O. PROD " & _Rows_Co(I).Item("OrdenProduccion")

            DataGridView2.Rows.Add(_TipoMov, _FechaMov, _Cantidad, _PrecioUnitario, _Umedida, _Proveedor, _Factura, _Referencia, _Motivo)

        Next


        ' Presenta las entradas por compra a proveedores
        _Motivo = "ENTRADA POR PROVEEDOR"

        For Me.I = 0 To _Rows_Dr.GetUpperBound(0)
            _TipoMov = "ENTRADA"
            _Recepcion = _Rows_Dr(I).Item("Recepcion")
            _Busca = CServiciosDataSet46.Recepciones.FindByRecepcion(_Recepcion)
            _FechaMov = _Busca.Item("Fecha")
            _Factura = _Busca.Item("Documento")
            _Cantidad = _Rows_Dr(I).Item("CantidadRecibe")
            _PrecioUnitario = _Rows_Dr(I).Item("PrecioUnitario")
            _Umedida = _Rows_Dr(I).Item("Umedida")
            _Proveedor = _Rows_Dr(I).Item("ProveedorRecibe")

            _NombreProveedor = "** SIN PROVEEDOR **"
            For X = 0 To CServiciosDataSet56.Proveedores.Rows.Count - 1
                If CServiciosDataSet56.Proveedores.Rows(X).Item("Proveedor") = _Proveedor Then
                    _NombreProveedor = CServiciosDataSet56.Proveedores.Rows(X).Item("RazonSocial")
                    Exit For
                End If
            Next

            _Referencia = _Rows_Dr(I).Item("MarcaRecibe")

            DataGridView2.Rows.Add(_TipoMov, _FechaMov, _Cantidad, _PrecioUnitario, _Umedida, _NombreProveedor, _Factura, _Referencia, _Motivo)

        Next

        DataGridView2.Sort(Fecha, System.ComponentModel.ListSortDirection.Ascending)

        For Me.I = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(I).Cells(0).Value = "ENTRADA" Then
                DataGridView2.Rows(I).DefaultCellStyle.BackColor = Color.LightGreen
            End If
        Next



    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub
End Class