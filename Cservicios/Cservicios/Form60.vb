Public Class Frm60Facturas
    Public I As Integer
    Public _Busca As DataRow
    Public _Renglon As DataGridViewRow
    Public _Remision As String
    Public _Fecha As Date
    Public _Importe As Double
    Public _ChecaControl As Control
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Documento, _Cliente, _Centro, _Paciente, _NombreCliente, _NombrePaciente As String
    Public _FechaDocumento As Date

    Private Sub BindingNavigatorMoveFirstItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Frm60Facturas_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet53.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter1.Fill(Me.CServiciosDataSet53.VentaFactor)

        Button2_Click(sender, e)

    End Sub

    Private Sub GeneraFacturasToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GeneraFacturasToolStripMenuItem.Click
        TabControl1.Visible = True
        TextBox1.Focus()
        Timer1.Enabled = True
        Timer1.Interval = 100

        TabPage2.Parent = Nothing
        TabPage3.Parent = Nothing
        TabPage7.Parent = Nothing
        TabPage1.Parent = TabControl1

        ComboBox1.Items.Clear()
        For Me.I = 0 To Me.CServiciosDataSet30.Clientes.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet30.Clientes.Rows(I).Item("NombreCliente"))
        Next


    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ' Muestra las ventas pendientes de facturar
        TabControl2.Visible = True
        TabControl2.SelectedIndex = 0

        If TextBox4.Text = "" Then
            Msg = "Debe seleccionar un Cliente para esta Factura"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        MsgBox("Hola")

        ' Nutriciones
        If RadioButton1.Checked = True Then
            _Centro = "CME"
            _Cliente = TextBox4.Text
            Dim _Rows_Re() As DataRow = Me.CServiciosDataSet50.Remisiones.Select("Cliente = " & "'" & _Cliente & "'")
            Dim _find As Boolean
            TabPage4.Parent = TabControl2
            TabPage5.Parent = Nothing


            _find = False
            For Me.I = 0 To _Rows_Re.GetUpperBound(0)
                If _Rows_Re(I).Item("Estatus") <> "CANCELADA" Then
                    If _Rows_Re(I).Item("Factura") = "SF" Then
                        _Remision = _Rows_Re(I).Item("Remision")
                        _Fecha = _Rows_Re(I).Item("Fecha")
                        _Cliente = _Rows_Re(I).Item("Cliente")
                        _Paciente = _Rows_Re(I).Item("Paciente")
                        _Importe = _Rows_Re(I).Item("Importe")


                        _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                        _NombreCliente = _Busca.Item("NombreCliente")

                        _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                        _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")


                        DataGridView1.Rows.Add(_Remision, _Fecha, _Cliente, _NombreCliente, _Paciente, _NombrePaciente, _Importe, False, True)
                        _find = True


                    End If
                End If
            Next

            If Not _find Then
                Msg = "No hay Remisiones de ese Cliente pendientes de Facturar. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If


        End If


        ' Factor de Transferencia
        If RadioButton2.Checked = True Then
            _Cliente = TextBox4.Text

            TabPage4.Parent = Nothing
            TabPage5.Parent = TabControl2

            Dim _Find As Boolean
            Dim _Rows_Vf() As DataRow = Me.CServiciosDataSet53.VentaFactor.Select("Caja = " & "'" & _Cliente & "'")
            Dim _VFactura As String
            Dim _Folio As String
            Dim _Id As Long
            Dim _Referencia As String

            _Find = False
            DataGridView2.Rows.Clear()
            For Me.I = 0 To _Rows_Vf.GetUpperBound(0)
                If DBNull.Value.Equals(_Rows_Vf(I).Item("Factura")) Then
                    _VFactura = ""
                Else
                    _VFactura = _Rows_Vf(I).Item("Factura")
                End If

                If _VFactura = "" Then
                    _Folio = _Rows_Vf(I).Item("Folio")
                    _FechaDocumento = _Rows_Vf(I).Item("FechaFolio")
                    _Cliente = _Rows_Vf(I).Item("Caja")
                    _Paciente = ""
                    _NombrePaciente = _Rows_Vf(I).Item("NombrePaciente")
                    _Importe = _Rows_Vf(I).Item("PrecioNeto")
                    _Id = _Rows_Vf(I).Item("Orden")

                    _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                    _NombreCliente = _Busca.Item("NombreCliente")
                    _Find = True
                    DataGridView2.Rows.Add(_Folio, _FechaDocumento, _Cliente, _NombreCliente, _Paciente, _NombrePaciente, _Importe, False, _Id)
                End If
            Next

            _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
            _Referencia = _Busca.Item("Referencia")

            _Rows_Vf = Me.CServiciosDataSet53.VentaFactor.Select("Caja = " & "'" & _Referencia & "'")
            _Find = False
            DataGridView2.Rows.Clear()
            For Me.I = 0 To _Rows_Vf.GetUpperBound(0)
                If DBNull.Value.Equals(_Rows_Vf(I).Item("Factura")) Then
                    _VFactura = ""
                Else
                    _VFactura = _Rows_Vf(I).Item("Factura")
                End If

                If _VFactura = "" Then
                    _Folio = _Rows_Vf(I).Item("Folio")
                    _FechaDocumento = _Rows_Vf(I).Item("FechaFolio")
                    _Cliente = TextBox4.Text
                    _Paciente = ""
                    _NombrePaciente = _Rows_Vf(I).Item("NombrePaciente")
                    _Importe = _Rows_Vf(I).Item("PrecioNeto")
                    _Id = _Rows_Vf(I).Item("Orden")

                    _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                    _NombreCliente = _Busca.Item("NombreCliente")
                    _Find = True
                    DataGridView2.Rows.Add(_Folio, _FechaDocumento, _Cliente, _NombreCliente, _Paciente, _NombrePaciente, _Importe, False, _Id)
                End If
            Next

            If Not _Find Then
                Msg = "No existen Folios pendientes de Facturar de ese Cliente. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If


        End If





    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet52.DetalleCxC' Puede moverla o quitarla según sea necesario.
        Me.DetalleCxCTableAdapter.Fill(Me.CServiciosDataSet52.DetalleCxC)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet52.CxC' Puede moverla o quitarla según sea necesario.
        Me.CxCTableAdapter.Fill(Me.CServiciosDataSet52.CxC)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet51.DetalleRemision' Puede moverla o quitarla según sea necesario.
        Me.DetalleRemisionTableAdapter.Fill(Me.CServiciosDataSet51.DetalleRemision)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet50.Remisiones' Puede moverla o quitarla según sea necesario.
        Me.RemisionesTableAdapter.Fill(Me.CServiciosDataSet50.Remisiones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.Clientes' Puede moverla o quitarla según sea necesario.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet21.Productos' Puede moverla o quitarla según sea necesario.
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet32.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet32.VentaFactor)
        TabControl1.Visible = False
        RadioButton1.Checked = True
        ToolStripButton2.Enabled = True
        ToolStripButton4.Enabled = True
        ToolStripButton5.Enabled = True
        ToolStripButton9.Enabled = False
        GuardarInformaciónToolStripMenuItem.Enabled = False

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        For Each Me._ChecaControl In TabPage1.Controls
            If TypeOf _ChecaControl Is TextBox Then
                If _ChecaControl.Focused Then
                    _ChecaControl.BackColor = Color.Gold
                End If
            End If
        Next
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox2.Text = CStr(Today.Date)
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = Mid(TimeOfDay, 1, 5)
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox4_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.BackColor = Color.White
        If TextBox4.Text = "" Then
            ComboBox1.Focus()
        Else
            ' Busca el Cliente
            _Cliente = TextBox4.Text

            Dim _Rows_Cl() As DataRow = Me.CServiciosDataSet30.Clientes.Select("Cliente = " & "'" & _Cliente & "'")
            Dim _Find As Boolean

            _Find = False

            For Me.I = 0 To _Rows_Cl.GetUpperBound(0)
                'MsgBox("hola")
                _Find = True
                _NombreCliente = _Rows_Cl(I).Item("NombreCliente")
                ComboBox1.Text = _NombreCliente
                TextBox5.Text = _Rows_Cl(I).Item("Calle")
                TextBox6.Text = _Rows_Cl(I).Item("Colonia")
                TextBox7.Text = _Rows_Cl(I).Item("Ciudad")
                TextBox8.Text = _Rows_Cl(I).Item("Estado")

                TextBox9.Focus()
            Next

            If Not _Find Then
                Msg = "Ese Cliente no está Registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                TextBox4.Focus()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Button1.Focus()
        End If
    End Sub

    Private Sub TextBox9_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox9.LostFocus
        TextBox10.Text = TextBox9.Text
        TextBox9.BackColor = Color.White
        _Importe = TextBox9.Text
        TextBox9.Text = Format(_Importe, "###,###,##0.00")
        DataGridView1.Focus()
        Timer1.Enabled = False

    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        If _Renglon.Cells(7).Value = False Then
            _Renglon.Cells(7).Value = True
            _Renglon.DefaultCellStyle.BackColor = Color.Gold
        Else
            _Renglon.Cells(7).Value = False
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        ' Revisa la Selección
        Dim _SumaSeleccion As Double

        _SumaSeleccion = 0

        For Me.I = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(I).Cells(0).Value = "" Then
                Exit For
            End If
            If DataGridView1.Rows(I).Cells(7).Value = True Then
                '  MsgBox("Remision " & DataGridView1.Rows(I).Cells(0).Value)
                _SumaSeleccion = _SumaSeleccion + CDbl(DataGridView1.Rows(I).Cells(6).Value)
            End If

        Next

        If _SumaSeleccion = 0 Then
            Msg = "No hay Remisiones seleccionadas para Facturar. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        Else

            If _SumaSeleccion <> CDbl(TextBox10.Text) Then
                Msg = "La Suma de las remisiones Seleccionadas " & CStr(_SumaSeleccion) & " No coinciden con la Importe de la Factura proporcionado " & CStr(TextBox10.Text)
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            Else
                TextBox11.Text = Format(_SumaSeleccion, "$###,###,##0.00")
                TextBox12.Text = _SumaSeleccion

            End If

        End If

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        GeneraFacturasToolStripMenuItem_Click(sender, e)
        GeneraFacturasToolStripMenuItem.Enabled = False
        ToolStripButton5.Enabled = False
        ToolStripButton9.Enabled = True

        GuardarInformaciónToolStripMenuItem.Enabled = True
        TextBox9.Text = 0
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim _Index As Integer = ComboBox1.SelectedIndex
        TextBox4.Text = Trim(CServiciosDataSet30.Clientes.Rows(_Index).Item("Cliente"))
        TextBox5.Text = CServiciosDataSet30.Clientes.Rows(_Index).Item("Calle")
        TextBox6.Text = CServiciosDataSet30.Clientes.Rows(_Index).Item("Colonia")
        TextBox7.Text = CServiciosDataSet30.Clientes.Rows(_Index).Item("Ciudad")
        TextBox8.Text = CServiciosDataSet30.Clientes.Rows(_Index).Item("Estado")

        TextBox9.Focus()
    End Sub

    Private Sub Button1_GotFocus(sender As Object, e As System.EventArgs) Handles Button1.GotFocus
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton9_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton9.Click
        ' Guarda Información


        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de Guardar esta Factura?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Generación de Facturas a Clientes"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Verifica que no exista

            Dim _Find As Boolean
            Dim _Rows_CxC() As DataRow = Me.CServiciosDataSet52.CxC.Select("Documento = " & "'" & TextBox1.Text & "'")

            _Find = False
            For Me.I = 0 To _Rows_CxC.GetUpperBound(0)
                _Find = True
                _Cliente = _Rows_CxC(I).Item("Cliente")
                _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                _NombreCliente = _Busca.Item("Nombrecliente")
            Next

            If _Find Then
                Msg = "Ese Número de Factura ya está registrada para el Cliente " & _Cliente & " " & _NombreCliente
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            ' Guarda la factura
            _Busca = Me.CServiciosDataSet52.CxC.NewRow
            _Busca.Item("cliente") = TextBox4.Text
            _Busca.Item("Documento") = TextBox1.Text
            _Busca.Item("FechaDocumento") = CDate(TextBox2.Text)
            _Busca.Item("Importe") = CDbl(TextBox10.Text)
            _Busca.Item("Iva") = 0
            _Busca.Item("Neto") = CDbl(TextBox10.Text)
            _Busca.Item("Pagada") = False
            _Busca.Item("DocumentoPago") = ""
            _Busca.Item("FechaPago") = DBNull.Value
            _Busca.Item("ImportePago") = 0
            _Busca.Item("Referencia") = ComboBox1.Text
            _Busca.Item("Comentarios") = ""
            _Busca.Item("FechaRegistro") = Today.Date
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
            _Busca.Item("FechaVence") = CDate(TextBox2.Text)
            If RadioButton1.Checked = True Then
                _Busca.Item("Plazo") = "CME"
            End If
            If RadioButton2.Checked = True Then
                _Busca.Item("Plazo") = "FTR"
            End If


            Me.CServiciosDataSet52.CxC.Rows.Add(_Busca)

            Me.Validate()
            CxCBindingSource.EndEdit()
            Me.TableAdapterManager6.UpdateAll(CServiciosDataSet52)
            Me.CxCTableAdapter.Fill(CServiciosDataSet52.CxC)

            _Documento = TextBox1.Text
            _Busca = Me.CServiciosDataSet52.CxC.FindByDocumento(_Documento)

            Dim _IdCxC As Long = _Busca.Item("Orden")
            Dim _Rows_Dr() As DataRow
            Dim X As Integer
            Dim _BuscaProducto As DataRow
            Dim _Producto As String

            ' Guarda el detalle
            If RadioButton1.Checked = True Then
                For Me.I = 0 To DataGridView1.Rows.Count - 1

                    If DataGridView1.Rows(I).Cells(0).Value = "" Then
                        Exit For
                    End If

                    If DataGridView1.Rows(I).Cells(7).Value = True Then

                        _Remision = DataGridView1.Rows(I).Cells(0).Value
                        _Rows_Dr = Me.CServiciosDataSet51.DetalleRemision.Select("Remision = " & "'" & _Remision & "'")

                        For X = 0 To _Rows_Dr.GetUpperBound(0)

                            _Busca = Me.CServiciosDataSet52.DetalleCxC.NewRow
                            _Busca.Item("Documento") = _Documento
                            _Busca.Item("Cliente") = TextBox4.Text
                            _Busca.Item("Producto") = _Rows_Dr(X).Item("Producto")
                            _Busca.Item("Cantidad") = _Rows_Dr(X).Item("Cantidad")
                            _Busca.Item("PrecioUnitario") = _Rows_Dr(X).Item("PrecioUnitario")
                            _Busca.Item("Iva") = 0
                            _Busca.Item("Neto") = _Rows_Dr(X).Item("ImporteNeto")
                            _Busca.Item("Referencia") = DataGridView1.Rows(I).Cells(5).Value
                            _Busca.Item("Comentarios") = ""
                            _Busca.Item("FechaRegistro") = Today.Date
                            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
                            _Busca.Item("IdCxc") = _IdCxC

                            _Producto = _Rows_Dr(X).Item("Producto")
                            _BuscaProducto = Me.CServiciosDataSet21.Productos.FindByProducto(_Producto)
                            _Busca.Item("DescripcionProducto") = _BuscaProducto.Item("Descripcion")

                            Me.CServiciosDataSet52.DetalleCxC.Rows.Add(_Busca)

                        Next

                        _Busca = Me.CServiciosDataSet50.Remisiones.FindByRemision(_Remision)
                        _Busca.Item("Factura") = TextBox1.Text
                        _Busca.Item("Estatus") = "FACTURADA"

                        Me.Validate()
                        DetalleCxCBindingSource.EndEdit()
                        Me.TableAdapterManager6.UpdateAll(CServiciosDataSet52)
                        Me.DetalleCxCTableAdapter.Fill(Me.CServiciosDataSet52.DetalleCxC)

                        Me.Validate()
                        RemisionesBindingSource.EndEdit()
                        Me.TableAdapterManager4.UpdateAll(CServiciosDataSet50)
                        Me.RemisionesTableAdapter.Fill(Me.CServiciosDataSet50.Remisiones)



                    End If

                Next

            End If


            ' Guarda el detalle cuando es Factor de Transferencia 
            If RadioButton2.Checked = True Then
                For Me.I = 0 To DataGridView2.Rows.Count - 1

                    If DataGridView2.Rows(I).Cells(0).Value = "" Then
                        Exit For
                    End If

                    If DataGridView2.Rows(I).Cells(7).Value = True Then
                        _Documento = DataGridView2.Rows(I).Cells(0).Value

                        _Busca = Me.CServiciosDataSet52.DetalleCxC.NewRow
                        _Busca.Item("Documento") = TextBox1.Text
                        _Busca.Item("Producto") = "FTR"
                        _Busca.Item("Cliente") = TextBox4.Text

                        _BuscaProducto = Me.CServiciosDataSet53.VentaFactor.FindByFolio(_Documento)

                        _Busca.Item("Cantidad") = _BuscaProducto.Item("NumeroFrascos")
                        _Busca.Item("PrecioUnitario") = _BuscaProducto.Item("PrecioNeto") / _BuscaProducto.Item("NumeroFrascos")
                        _Busca.Item("Iva") = 0
                        _Busca.Item("Neto") = _BuscaProducto.Item("PrecioNeto")
                        _Busca.Item("Referencia") = DataGridView2.Rows(I).Cells(5).Value
                        _Busca.Item("Comentarios") = _Documento
                        _Busca.Item("FechaRegistro") = Today.Date
                        _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
                        _Busca.Item("IdCxc") = _IdCxC
                        _Busca.Item("DescripcionProducto") = "DOSIS DE FACTOR DE TRANSFERENCIA"

                        Me.CServiciosDataSet52.DetalleCxC.Rows.Add(_Busca)

                        _BuscaProducto = Me.CServiciosDataSet53.VentaFactor.FindByFolio(_Documento)
                        _BuscaProducto.Item("Factura") = TextBox1.Text


                    End If
                Next

                Me.Validate()
                DetalleCxCBindingSource.EndEdit()
                Me.TableAdapterManager6.UpdateAll(CServiciosDataSet52)
                Me.DetalleCxCTableAdapter.Fill(Me.CServiciosDataSet52.DetalleCxC)

                Me.Validate()
                VentaFactorBindingSource.EndEdit()
                Me.TableAdapterManager8.UpdateAll(Me.CServiciosDataSet53)
                Me.VentaFactorTableAdapter1.Fill(Me.CServiciosDataSet53.VentaFactor)

            End If


            Msg = "Registro guardado correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)

            Button2_Click(sender, e)



        Else
            Exit Sub
        End If

    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        Dim _SumaSeleccion As Double

        _SumaSeleccion = 0

        For Me.I = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(I).Cells(0).Value = "" Then
                Exit For
            End If
            If DataGridView2.Rows(I).Cells(7).Value = True Then
                '  MsgBox("Remision " & DataGridView1.Rows(I).Cells(0).Value)
                _SumaSeleccion = _SumaSeleccion + CDbl(DataGridView2.Rows(I).Cells(6).Value)
            End If

        Next

        If _SumaSeleccion = 0 Then
            Msg = "No hay Remisiones seleccionadas para Facturar. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        Else

            If _SumaSeleccion <> CInt(TextBox10.Text) Then
                Msg = "La Suma de las remisiones Seleccionadas " & CStr(_SumaSeleccion) & " No coinciden con la Importe de la Factura proporcionado " & CStr(TextBox10.Text)
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            Else
                TextBox11.Text = Format(_SumaSeleccion, "$###,###,##0.00")
                TextBox12.Text = _SumaSeleccion

            End If

        End If
    End Sub

    Private Sub ToolStripButton11_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton11.Click
        Button2_Click(sender, e)
    End Sub

    Private Sub ToolStripButton10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton10.Click
        Button2_Click(sender, e)
    End Sub

    Private Sub ToolStripButton12_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton12.Click
        TabControl1.Visible = True
        TabPage1.Parent = Nothing
        TabPage2.Parent = TabControl1
        TabPage3.Parent = Nothing
        TabPage7.Parent = Nothing
        RadioButton12.Checked = True
        Button4_Click(sender, e)

    End Sub

    Private Sub TabPage2_Click(sender As System.Object, e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        If _Renglon.Cells(7).Value = False Then
            _Renglon.Cells(7).Value = True
            _Renglon.DefaultCellStyle.BackColor = Color.Gold
        Else
            _Renglon.Cells(7).Value = False
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If


    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ' Muestra las ventas pendientes de facturar
        TabControl2.Visible = True
        TabControl2.SelectedIndex = 0

        If TextBox4.Text = "" Then
            Msg = "Debe seleccionar un Cliente para esta Factura"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        ' Nutriciones
        If RadioButton1.Checked = True Then
            _Centro = "CME"
            _Cliente = TextBox4.Text
            Dim _Rows_Re() As DataRow = Me.CServiciosDataSet50.Remisiones.Select("Cliente = " & "'" & _Cliente & "'")
            Dim _find As Boolean
            TabPage4.Parent = TabControl2
            TabPage5.Parent = Nothing


            _find = False
            For Me.I = 0 To _Rows_Re.GetUpperBound(0)
                If _Rows_Re(I).Item("Estatus") <> "CANCELADA" Then
                    If _Rows_Re(I).Item("Factura") = "SF" Then
                        _Remision = _Rows_Re(I).Item("Remision")
                        _Fecha = _Rows_Re(I).Item("Fecha")
                        _Cliente = _Rows_Re(I).Item("Cliente")
                        _Paciente = _Rows_Re(I).Item("Paciente")
                        _Importe = _Rows_Re(I).Item("Importe")


                        _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                        _NombreCliente = _Busca.Item("NombreCliente")

                        _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                        _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")


                        DataGridView1.Rows.Add(_Remision, _Fecha, _Cliente, _NombreCliente, _Paciente, _NombrePaciente, _Importe, False, True)
                        _find = True


                    End If
                End If
            Next

            If Not _find Then
                Msg = "No hay Remisiones de ese Cliente pendientes de Facturar. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If


        End If


        ' Factor de Transferencia
        If RadioButton2.Checked = True Then
            _Cliente = TextBox4.Text

            TabPage4.Parent = Nothing
            TabPage5.Parent = TabControl2

            Dim _Find As Boolean
            Dim _Rows_Vf() As DataRow = Me.CServiciosDataSet53.VentaFactor.Select("Caja = " & "'" & _Cliente & "'")
            Dim _VFactura As String
            Dim _Folio As String
            Dim _Id As Long
            Dim _Referencia As String

            _Find = False
            DataGridView2.Rows.Clear()
            For Me.I = 0 To _Rows_Vf.GetUpperBound(0)
                If DBNull.Value.Equals(_Rows_Vf(I).Item("Factura")) Then
                    _VFactura = ""
                Else
                    _VFactura = _Rows_Vf(I).Item("Factura")
                End If

                If _VFactura = "" Then
                    _Folio = _Rows_Vf(I).Item("Folio")
                    _FechaDocumento = _Rows_Vf(I).Item("FechaFolio")
                    _Cliente = _Rows_Vf(I).Item("Caja")
                    _Paciente = ""
                    _NombrePaciente = _Rows_Vf(I).Item("NombrePaciente")
                    _Importe = _Rows_Vf(I).Item("PrecioNeto")
                    _Id = _Rows_Vf(I).Item("Orden")

                    _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                    _NombreCliente = _Busca.Item("NombreCliente")
                    _Find = True
                    DataGridView2.Rows.Add(_Folio, _FechaDocumento, _Cliente, _NombreCliente, _Paciente, _NombrePaciente, _Importe, False, _Id)
                End If
            Next

            _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
            _Referencia = _Busca.Item("Referencia")

            _Rows_Vf = Me.CServiciosDataSet53.VentaFactor.Select("Caja = " & "'" & _Referencia & "'")
            _Find = False
            DataGridView2.Rows.Clear()
            For Me.I = 0 To _Rows_Vf.GetUpperBound(0)
                If DBNull.Value.Equals(_Rows_Vf(I).Item("Factura")) Then
                    _VFactura = ""
                Else
                    _VFactura = _Rows_Vf(I).Item("Factura")
                End If

                If _VFactura = "" Then
                    _Folio = _Rows_Vf(I).Item("Folio")
                    _FechaDocumento = _Rows_Vf(I).Item("FechaFolio")
                    _Cliente = TextBox4.Text
                    _Paciente = ""
                    _NombrePaciente = _Rows_Vf(I).Item("NombrePaciente")
                    _Importe = _Rows_Vf(I).Item("PrecioNeto")
                    _Id = _Rows_Vf(I).Item("Orden")

                    _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                    _NombreCliente = _Busca.Item("NombreCliente")
                    _Find = True
                    DataGridView2.Rows.Add(_Folio, _FechaDocumento, _Cliente, _NombreCliente, _Paciente, _NombrePaciente, _Importe, False, _Id)
                End If
            Next

            If Not _Find Then
                Msg = "No existen Folios pendientes de Facturar de ese Cliente. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If


        End If
    End Sub

    Private Sub RadioButton12_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton12.CheckedChanged

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim _Rows_CxC() As DataRow
        Dim _Espacios As String = ""

        If RadioButton12.Checked = True Then
            _Rows_CxC = CServiciosDataSet52.CxC.Select("Documento <> " & "'" & _Espacios & "'")
        End If







        For Me.I = 0 To _Rows_CxC.GetUpperBound(0)
            _Documento = _Rows_CxC(I).Item("Documento")
            _FechaDocumento = _Rows_CxC(I).Item("Fechadocumento")
            _Cliente = _Rows_CxC(I).Item("Cliente")
            _Importe = _Rows_CxC(I).Item("Importe")
            _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
            _NombreCliente = _Busca.Item("NombreCliente")

            DataGridView3.Rows.Add(_Documento, _FechaDocumento, _NombreCliente, _Importe)

        Next

    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        _Renglon = DataGridView3.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox13.Text = _Renglon.Cells(0).Value
        TextBox14.Text = _Renglon.Cells(1).Value
        _Documento = TextBox13.Text

        _Busca = Me.CServiciosDataSet52.CxC.FindByDocumento(_Documento)
        TextBox15.Text = Format(_Busca.Item("Neto"), "$###,###,##.00")
        TextBox16.Text = _Busca.Item("Cliente")
        If DBNull.Value.Equals(_Busca.Item("FechaPago")) Then
            TextBox22.Text = ""
        Else
            TextBox22.Text = _Busca.Item("FechaPago")
        End If

        TextBox23.Text = Format(_Busca.Item("ImportePago"), "$###,###,##0.00")
        TextBox24.Text = _Busca.Item("DocumentoPago")
        _Cliente = TextBox16.Text

        _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
        TextBox17.Text = _Busca.Item("NombreCliente")
        TextBox18.Text = _Busca.Item("Calle")
        TextBox19.Text = _Busca.Item("Colonia")
        TextBox20.Text = _Busca.Item("Ciudad")
        TextBox21.Text = _Busca.Item("Estado")

        ' Presenta el Detalle

        Dim _Rows_Dcxc() As DataRow = Me.CServiciosDataSet52.DetalleCxC.Select("Documento = " & "'" & _Documento & "'")
        Dim _Producto, _NombreProducto As String
        Dim _Cantidad, _PrecioUnitario As Integer
        Dim _Id As Integer

        DataGridView4.Rows.Clear()
        For Me.I = 0 To _Rows_Dcxc.GetUpperBound(0)
            _Documento = _Rows_Dcxc(I).Item("Documento")
            _Producto = _Rows_Dcxc(I).Item("Producto")
            _NombreProducto = _Rows_Dcxc(I).Item("DescripcionProducto")
            _Cantidad = _Rows_Dcxc(I).Item("Cantidad")
            _PrecioUnitario = _Rows_Dcxc(I).Item("PrecioUnitario")
            _Importe = _Rows_Dcxc(I).Item("Neto")
            _NombrePaciente = _Rows_Dcxc(I).Item("Referencia")
            _Id = _Rows_Dcxc(I).Item("Orden")

            DataGridView4.Rows.Add(_Documento, _Producto, _NombreProducto, _Cantidad, _PrecioUnitario, _Importe, _NombrePaciente, _Id)

        Next






    End Sub

    Private Sub DataGridView3_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView3_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellLeave
        _Renglon = DataGridView3.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton7.CheckedChanged

    End Sub

    Private Sub HistorialDeFacturasToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HistorialDeFacturasToolStripMenuItem.Click
        ToolStripButton12_Click(sender, e)
    End Sub

    Private Sub RefrescarInformaciónToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RefrescarInformaciónToolStripMenuItem.Click
        Button2_Click(sender, e)
    End Sub

    Private Sub DeshacerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DeshacerToolStripMenuItem.Click
        Button2_Click(sender, e)
    End Sub

    Private Sub FacturasPorClienteToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FacturasPorClienteToolStripMenuItem.Click
        TabControl1.Visible = True
        TabPage1.Parent = Nothing
        TabPage2.Parent = Nothing
        TabPage3.Parent = Nothing
        TabPage7.Parent = TabControl1
        Timer1.Enabled = True



        ' Llena el Combo de los Clientes
        ComboBox2.Items.Clear()

        For Me.I = 0 To CServiciosDataSet30.Clientes.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet30.Clientes.Rows(I).Item("NombreCliente"))
        Next
        TextBox25.Focus()

    End Sub

    Private Sub TextBox25_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox25.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub TextBox25_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox25.LostFocus
        Dim _Find As Boolean

        If TextBox25.Text = "" Then
            ComboBox2.Focus()
        Else
            _Cliente = TextBox25.Text
            _Find = False

            For Me.I = 0 To Me.CServiciosDataSet30.Clientes.Rows.Count - 1
                If CServiciosDataSet30.Clientes.Rows(I).Item("Cliente") = _Cliente Then
                    ComboBox2.Text = CServiciosDataSet30.Clientes.Rows(I).Item("NombreCliente")
                    _Find = True
                End If
            Next

            If Not _Find Then
                Msg = "Ese cliente no está registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

        End If

    End Sub

    Private Sub TextBox25_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress

    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub
End Class