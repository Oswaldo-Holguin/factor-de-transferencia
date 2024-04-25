Public Class Frm35AltaFactura
    Public i As Integer
    Public Msg As String
    Public _Importe, _Iva, _Neto, _ImportePago As Integer
    Public Style As MsgBoxStyle
    Public _Busca As DataRow
    Public _Renglon As DataGridViewRow
    Public _Pagado As Boolean
    Public _Checacontrol As Control
    Public _Factura, _Referencia As String
    Public _Cliente, Documento, _Comentarios, _Plazo As String
    Public _FechaDocumento, _FechaVence, _FechaPago As Date


    Private Sub Frm35AltaFactura_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet35.CxC' Puede moverla o quitarla según sea necesario.
        Me.CxCTableAdapter1.Fill(Me.CServiciosDataSet35.CxC)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet31.Plazos' Puede moverla o quitarla según sea necesario.
        Me.PlazosTableAdapter.Fill(Me.CServiciosDataSet31.Plazos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.DetalleCxC' Puede moverla o quitarla según sea necesario.
        Me.DetalleCxCTableAdapter.Fill(Me.CServiciosDataSet30.DetalleCxC)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.CxC' Puede moverla o quitarla según sea necesario.
        '   Me.CxCTableAdapter.Fill(Me.CServiciosDataSet30.CxC)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet21.Productos' Puede moverla o quitarla según sea necesario.
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        ToolStripTextBox1.Text = Frm33Clientes.TextBox1.Text
        ToolStripTextBox2.Text = Frm33Clientes.TextBox2.Text
        Panel1.Enabled = False
        Panel2.Enabled = False
        Button3.Enabled = False

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True
        Panel2.Enabled = True
        TextBox1.Focus()
        TextBox13.Text = "ALTA"
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        For Each Me._Checacontrol In Panel2.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        ComboBox1.Text = ""
        ComboBox2.Text = ""
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
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

    Private Sub ComboBox1_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.GotFocus
        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet31.Plazos.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet31.Plazos.Rows(i).Item("Descripcion"))
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim _NumeroDias As Integer
        Dim _FechaVence As Date

        _NumeroDias = 0
        For Me.i = 0 To CServiciosDataSet31.Plazos.Rows.Count - 1
            If CServiciosDataSet31.Plazos.Rows(i).Item("Descripcion") = ComboBox1.Text Then
                TextBox3.Text = CServiciosDataSet31.Plazos.Rows(i).Item("Plazo")
                _NumeroDias = CServiciosDataSet31.Plazos.Rows(i).Item("NumeroDias")
            End If
        Next
        _FechaVence = CDate(TextBox2.Text).AddDays(_NumeroDias)
        TextBox4.Text = _FechaVence
        TextBox8.Focus()

    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox2.GotFocus
        ComboBox2.Items.Clear()
        For Me.i = 0 To CServiciosDataSet21.Productos.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet21.Productos.Rows(i).Item("Descripcion"))
        Next
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet21.Productos.Rows.Count - 1
            If CServiciosDataSet21.Productos.Rows(i).Item("Descripcion") = ComboBox2.Text Then
                TextBox9.Text = CServiciosDataSet21.Productos.Rows(i).Item("Producto")
                TextBox10.Text = CServiciosDataSet21.Productos.Rows(i).Item("PrecioUnitario")
            End If
        Next
        TextBox9.Enabled = False
        TextBox10.Enabled = True
        TextBox10.Focus()
    End Sub

    Private Sub TextBox8_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox8.LostFocus
        TextBox8.Text = UCase(TextBox8.Text)

    End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox11.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox11_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox11.LostFocus
        Dim _Importe As Integer = CInt(TextBox10.Text) * CInt(TextBox11.Text)
        TextBox12.Text = _Importe
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox9.Text = "" Then
            Exit Sub
        End If

        _Factura = TextBox1.Text
        Dim _Producto As String = TextBox9.Text
        Dim _PrecioUnitario As Integer = CInt(TextBox10.Text)
        Dim _Cantidad As Integer = CInt(TextBox11.Text)
        Dim _Importe As Integer = CInt(TextBox12.Text)
        Dim _Descripcion As String = ComboBox2.Text

        DataGridView1.Rows.Add(_Factura, _Producto, _Descripcion, _Cantidad, _PrecioUnitario, _Importe)

        ComboBox2.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""

        _Importe = 0
        For Me.i = 0 To DataGridView1.Rows.Count - 1
            _Importe = _Importe + CInt(DataGridView1.Rows(i).Cells(5).Value)
        Next
        TextBox5.Text = _Importe
        TextBox6.Text = 0
        TextBox7.Text = _Importe

        If _Producto = "FTR" Then
            Button3.Enabled = True
        End If



    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox10.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        ' Guardar Información.

        Dim _Producto As String
        Dim _Idcxc, _Cantidad, _PrecioUnitario As Integer
        Dim _Documento As String
        Dim _Find As Boolean
        Dim title As String
        Dim response As MsgBoxResult
        Dim _Rows_CxC() As DataRow

        Msg = "Está seguro de guardar esta Factura?"
        Style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Facturas por venta de Factor de Transferencia o  Nutricionales Parenterales"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Revisa los datos
            If DataGridView1.Rows.Count = 0 Then
                Msg = "Debe asignar un producto a esta Factura para poder guardar la información"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If
            _Find = False

            For Each Me._Checacontrol In Panel1.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    _Checacontrol.Text = UCase(_Checacontrol.Text)
                    If _Checacontrol.Text = "" Then
                        _Find = True
                    End If
                End If
            Next


            If _Find Then
                Msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            ' Verfifica que no se haya registrado esa factura
            _Documento = TextBox1.Text
            _Rows_CxC = CServiciosDataSet35.CxC.Select("Documento = " & "'" & _Documento & "'")

            _Find = False
            For Me.i = 0 To _Rows_CxC.GetUpperBound(0)
                _Find = True
                _Cliente = _Rows_CxC(i).Item("Cliente")
                _FechaDocumento = _Rows_CxC(i).Item("FechaDocumento")
            Next
            If _Find Then
                Msg = "Esta factura ya está registrada para el Cliente " & _Cliente & " Con Fecha " & CStr(_FechaDocumento)
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            ' Guarda la factura
            _Cliente = ToolStripTextBox1.Text
            _FechaDocumento = CDate(TextBox2.Text)
            _Importe = CInt(TextBox5.Text)
            _Iva = CInt(TextBox6.Text)
            _Neto = CInt(TextBox7.Text)
            _Pagado = False
            _Comentarios = "VENTA DE FACTOR DE TRANSFERENCIA"
            _FechaVence = CDate(TextBox4.Text)

            _Busca = CServiciosDataSet35.CxC.NewRow
            _Busca.Item("Cliente") = _Cliente
            _Busca.Item("Documento") = _Documento
            _Busca.Item("FechaDocumento") = _FechaDocumento
            _Busca.Item("Importe") = _Importe
            _Busca.Item("Iva") = _Iva
            _Busca.Item("Neto") = _Neto
            _Busca.Item("Pagada") = _Pagado
            _Busca.Item("DocumentoPago") = ""
            _Busca.Item("FechaPago") = DBNull.Value
            _Busca.Item("ImportePago") = 0
            _Busca.Item("Referencia") = _Referencia
            _Busca.Item("Comentarios") = _Comentarios
            _Busca.Item("FechaRegistro") = Today
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
            _Busca.Item("FechaVence") = _FechaVence
            _Busca.Item("Plazo") = TextBox3.Text

            CServiciosDataSet35.CxC.Rows.Add(_Busca)

            Me.Validate()
            CxCBindingSource1.EndEdit()
            Me.TableAdapterManager3.UpdateAll(CServiciosDataSet35)
            Me.CxCTableAdapter1.Fill(Me.CServiciosDataSet35.CxC)


            ' Registra el detalle
            _Busca = CServiciosDataSet35.CxC.FindByClienteDocumento(_Cliente, _Documento)
            _Idcxc = _Busca.Item("Orden")

            For Me.i = 0 To DataGridView1.Rows.Count - 1
                _Producto = DataGridView1.Rows(i).Cells(1).Value
                _Cantidad = CInt(DataGridView1.Rows(i).Cells(3).Value)
                _PrecioUnitario = CInt(DataGridView1.Rows(i).Cells(4).Value)
                _Iva = 0
                _Neto = _Cantidad * _PrecioUnitario

                _Busca = CServiciosDataSet30.DetalleCxC.NewRow
                _Busca.Item("Documento") = _Documento
                _Busca.Item("Cliente") = _Cliente
                _Busca.Item("Producto") = _Producto
                _Busca.Item("Cantidad") = _Cantidad
                _Busca.Item("PrecioUnitario") = _PrecioUnitario
                _Busca.Item("Iva") = 0
                _Busca.Item("Neto") = _Neto
                _Busca.Item("Referencia") = _Referencia
                _Busca.Item("Comentarios") = _Comentarios
                _Busca.Item("FechaRegistro") = Today
                _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
                _Busca.Item("IdCxc") = _Idcxc

                CServiciosDataSet30.DetalleCxC.Rows.Add(_Busca)
            Next

            Me.Validate()
            DetalleCxCBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet30)
            Me.DetalleCxCTableAdapter.Fill(Me.CServiciosDataSet30.DetalleCxC)

            Msg = "Registro guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)

            'Me.Close()

        Else
            Exit Sub
        End If

    End Sub
End Class