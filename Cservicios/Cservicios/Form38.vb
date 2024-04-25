Public Class Frm38Envios
    Public i As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Busca As DataRow
    Public _Checacontrol As Control
    Public _FechaEnvio, _FechaDeposito As Date
    Public _Comentarios, _Documento As String
    Public _Horaenvio, _CodigoRastreo, _BancoDeposito, _TelefonoCasa, _TelefonoOficina, _Celular As String
    Public _Cantidad, _PrecioUnitario, _CargoEnvio, _ImporteDeposito As Integer
    Public _PesoEnvio As Double
    Public _Envio, _Transporte, _Contacto, _Calle, _Colonia, _Ciudad, _Estado, _CP, _Referencia As String

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Frm38Envios_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
      


        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet34.Proveedores' Puede moverla o quitarla según sea necesario.
        Me.ProveedoresTableAdapter.Fill(Me.CServiciosDataSet34.Proveedores)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet34.Envios' Puede moverla o quitarla según sea necesario.
        Me.EnviosTableAdapter.Fill(Me.CServiciosDataSet34.Envios)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet32.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet32.VentaFactor)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet26.Produccion' Puede moverla o quitarla según sea necesario.
        Me.ProduccionTableAdapter.Fill(Me.CServiciosDataSet26.Produccion)

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.BackColor = Color.White
            End If
        Next
        ComboBox1.Text = ""
        RichTextBox1.Text = ""

        DataGridView1.Rows.Clear()
        For Me.i = 0 To CServiciosDataSet34.Envios.Rows.Count - 1
            _Envio = CServiciosDataSet34.Envios.Rows(i).Item("Envio")
            _FechaEnvio = CServiciosDataSet34.Envios.Rows(i).Item("FechaEnvio")
            _Contacto = CServiciosDataSet34.Envios.Rows(i).Item("Contacto")

            DataGridView1.Rows.Add(_Envio, _FechaEnvio, _Contacto)

        Next

        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        ToolStripButton2.Enabled = True

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True

        TextBox22.Text = "ALTA"

        TextBox1.Focus()

        ' Obtiene el Siguiente Consecutivo
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("ENV")
        Dim _Consecutivo As Integer = _Busca.Item("Consecutivo")
        Dim _XConsecutivo As String = CStr(_Consecutivo)
        _XConsecutivo = "000" & _XConsecutivo

        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "EN" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()
        TextBox21.Text = "SIN REFERENCIA"

    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        MonthCalendar1.Visible = True
        TextBox2.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox2.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False
    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = Mid(TimeOfDay, 1, 5)
        MonthCalendar1.Visible = False
        MonthCalendar2.Visible = False
        TextBox3.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.GotFocus
        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet34.Proveedores.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet34.Proveedores.Rows(i).Item("RazonSocial"))
        Next
        ComboBox1.BackColor = Color.LightGreen
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet34.Proveedores.Rows.Count - 1
            If CServiciosDataSet34.Proveedores.Rows(i).Item("RazonSocial") = ComboBox1.Text Then
                TextBox4.Text = CServiciosDataSet34.Proveedores.Rows(i).Item("Proveedor")
            End If
        Next
        TextBox4.Focus()
        ComboBox1.BackColor = Color.White
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox6.GotFocus
        MonthCalendar2.Visible = True
        TextBox5.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox6.LostFocus

    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub MonthCalendar2_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateChanged

    End Sub

    Private Sub MonthCalendar2_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateSelected
        TextBox6.Text = MonthCalendar2.SelectionRange.Start
        MonthCalendar2.Visible = False
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guarda la información.

        Dim _Find As Boolean
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de guardar la información del Envío ?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Envíos Foráneos "
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Guarda la información.
            _Find = False
            For Each Me._Checacontrol In Panel1.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    _Checacontrol.Text = UCase(_Checacontrol.Text)
                    If _Checacontrol.Text = "" Then
                        _Checacontrol.BackColor = Color.LightGray
                        _Find = True
                    End If
                End If
            Next
            RichTextBox1.Text = UCase(RichTextBox1.Text)

            If _Find Then
                Msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            If IsDate(TextBox2.Text) Then
            Else
                Msg = "Debe ingresar una fecha de Envío válida."
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                TextBox2.Focus()
                Exit Sub
            End If

            If IsDate(TextBox6.Text) Then
            Else
                Msg = "Debe ingresar una fecha de Depósito válida."
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                TextBox6.Focus()
                Exit Sub
            End If

            If TextBox22.Text = "ALTA" Then
                _Busca = CServiciosDataSet34.Envios.NewRow
                _Busca.Item("Envio") = TextBox1.Text
            End If

            _Envio = TextBox1.Text
            If TextBox22.Text = "CAMBIO" Then
                _Busca = CServiciosDataSet34.Envios.FindByEnvio(_Envio)
            End If
            _FechaEnvio = CDate(TextBox2.Text)
            _Horaenvio = TextBox3.Text
            _Transporte = TextBox4.Text
            _CodigoRastreo = TextBox5.Text
            _FechaDeposito = CDate(TextBox6.Text)
            _ImporteDeposito = CInt(TextBox7.Text)
            _BancoDeposito = TextBox8.Text
            _Contacto = TextBox9.Text
            _Cantidad = CInt(TextBox10.Text)
            _PrecioUnitario = CInt(TextBox11.Text)
            _CargoEnvio = CInt(TextBox12.Text)
            _Calle = TextBox13.Text
            _Colonia = TextBox14.Text
            _Ciudad = TextBox15.Text
            _Estado = TextBox16.Text
            _CP = TextBox23.Text
            _TelefonoCasa = TextBox17.Text
            _TelefonoOficina = TextBox18.Text
            _Celular = TextBox19.Text
            _PesoEnvio = CDbl(TextBox20.Text)
            _PesoEnvio = FormatNumber(_PesoEnvio, 3)
            _Referencia = ComboBox2.Text
            _Documento = ""
            _Comentarios = RichTextBox1.Text

            _Busca.Item("FechaEnvio") = _FechaEnvio
            _Busca.Item("HoraEnvio") = _Horaenvio
            _Busca.Item("Transporte") = _Transporte
            _Busca.Item("CodigoRastreo") = _CodigoRastreo
            _Busca.Item("FechaDeposito") = _FechaDeposito
            _Busca.Item("BancoDeposito") = _BancoDeposito
            _Busca.Item("ImporteDeposito") = _ImporteDeposito
            _Busca.Item("Contacto") = _Contacto
            _Busca.Item("Cantidad") = _Cantidad
            _Busca.Item("PrecioUnitario") = _PrecioUnitario
            _Busca.Item("CargoEnvio") = _CargoEnvio
            _Busca.Item("Calle") = _Calle
            _Busca.Item("Colonia") = _Colonia
            _Busca.Item("Ciudad") = _Ciudad
            _Busca.Item("Estado") = _Estado
            _Busca.Item("CodigoPostal") = _CP
            _Busca.Item("TelefonoCasa") = _TelefonoCasa
            _Busca.Item("TelefonoOficina") = _TelefonoOficina
            _Busca.Item("Celular") = _Celular
            _Busca.Item("PesoEnvio") = _PesoEnvio
            _Busca.Item("Referencia") = _Referencia
            _Busca.Item("Documento") = _Documento
            _Busca.Item("Comentarios") = _Comentarios


            If TextBox22.Text = "ALTA" Then
                CServiciosDataSet34.Envios.Rows.Add(_Busca)
            End If

            Me.Validate()
            EnviosBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet34)
            Me.EnviosTableAdapter.Fill(Me.CServiciosDataSet34.Envios)

            If TextBox22.Text = "ALTA" Then
                'Actualiza el consecutivo
                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("ENV")
                _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

                Me.Validate()
                DocumentosBindingSource.EndEdit()
                Me.TableAdapterManager1.UpdateAll(CServiciosDataSet3)

                ' Registra el envio como Venta
                _Busca = CServiciosDataSet32.VentaFactor.NewRow
                _Busca.Item("Folio") = TextBox1.Text
                _Busca.Item("FechaFolio") = _FechaEnvio
                _Busca.Item("Paciente") = "ENVIO FORANEO"
                _Busca.Item("NombrePaciente") = _Contacto
                _Busca.Item("NumeroFrascos") = _Cantidad
                _Busca.Item("PrecioUnitario") = _PrecioUnitario
                _Busca.Item("PorcentajeDescuento") = 0
                _Busca.Item("ImporteDescuento") = 0
                _Busca.Item("PrecioNeto") = _ImporteDeposito
                _Busca.Item("Referencia") = _Referencia
                _Busca.Item("FechaRegistro") = Today
                _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
                _Busca.Item("Comentarios") = "ENVIO FORANEO DE FACTOR DE TRANSFERENCIA"
                _Busca.Item("Caja") = "CL13"
                _Busca.Item("Consecutivo") = "0"
                _Busca.Item("Contado") = True
                _Busca.Item("Medico") = "N/A"
                _Busca.Item("Cedula") = "NA"
                _Busca.Item("Diagnostico") = "N/A"

                CServiciosDataSet32.VentaFactor.Rows.Add(_Busca)

                Me.Validate()
                VentaFactorBindingSource.EndEdit()
                Me.TableAdapterManager2.UpdateAll(CServiciosDataSet32)
                Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet32.VentaFactor)

            End If


            Button1_Click(sender, e)

            Msg = "Registro guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub

        Else
            Exit Sub
        End If
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        TextBox22.Text = "CAMBIO"
        ToolStripButton4.Enabled = False
        ToolStripButton3.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Focus()
    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub TextBox20_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox20.GotFocus
        TextBox20.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox20_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox20.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox20_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox20.LostFocus
        TextBox20.BackColor = Color.White
    End Sub

    Private Sub TextBox20_RegionChanged(sender As Object, e As System.EventArgs) Handles TextBox20.RegionChanged

    End Sub

    Private Sub TextBox20_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(0).Value
        _Envio = TextBox1.Text
        _Busca = CServiciosDataSet34.Envios.FindByEnvio(_Envio)
        TextBox2.Text = _Busca.Item("FechaEnvio")
        TextBox3.Text = _Busca.Item("HoraEnvio")
        TextBox4.Text = _Busca.Item("Transporte")
        TextBox5.Text = _Busca.Item("CodigoRastreo")
        TextBox6.Text = _Busca.Item("FechaDeposito")
        TextBox7.Text = _Busca.Item("ImporteDeposito")
        TextBox8.Text = _Busca.Item("BancoDeposito")
        TextBox9.Text = _Busca.Item("Contacto")
        TextBox10.Text = _Busca.Item("Cantidad")
        TextBox11.Text = _Busca.Item("PrecioUnitario")
        TextBox12.Text = _Busca.Item("CargoEnvio")
        TextBox13.Text = _Busca.Item("Calle")
        TextBox14.Text = _Busca.Item("Colonia")
        TextBox15.Text = _Busca.Item("Ciudad")
        TextBox16.Text = _Busca.Item("Estado")
        TextBox23.Text = _Busca.Item("CodigoPostal")
        TextBox17.Text = _Busca.Item("TelefonoCasa")
        TextBox18.Text = _Busca.Item("TelefonoOficina")
        TextBox19.Text = _Busca.Item("Celular")
        TextBox20.Text = _Busca.Item("PesoEnvio")
        'TextBox21.Text = _Busca.Item("Referencia")
        ComboBox2.Text = _Busca.Item("Referencia")
        RichTextBox1.Text = _Busca.Item("Comentarios")

        TextBox1.Enabled = False
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True

        _Transporte = TextBox4.Text
        _Busca = CServiciosDataSet34.Proveedores.FindByProveedor(_Transporte)
        ComboBox1.Text = _Busca.Item("RazonSocial")
        TextBox4.Enabled = False

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.BackColor = Color.LightGreen
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

    Private Sub TextBox7_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox10.GotFocus
        TextBox10.BackColor = Color.LightGreen
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

    Private Sub TextBox10_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox10.LostFocus
        TextBox10.BackColor = Color.White
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub BindingNavigator1_RefreshItems(sender As System.Object, e As System.EventArgs) Handles BindingNavigator1.RefreshItems

    End Sub

    Private Sub TextBox11_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox11.GotFocus
        TextBox11.BackColor = Color.LightGreen
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
        TextBox11.BackColor = Color.White
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox12_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox12.GotFocus
        TextBox12.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox12.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox12_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox12.LostFocus
        TextBox12.BackColor = Color.White
    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox23_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox23.GotFocus
        TextBox23.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox23_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox23.LostFocus
        TextBox23.Text = Mid(TextBox23.Text, 1, 5)
        TextBox23.BackColor = Color.White
    End Sub

    Private Sub TextBox23_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox23.TextChanged

    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox2.GotFocus
        ComboBox2.Items.Clear()
        Dim _OrdenProduccion As String

        For Me.i = 0 To CServiciosDataSet26.Produccion.Rows.Count - 1
            _OrdenProduccion = CServiciosDataSet26.Produccion.Rows(i).Item("OrdenProduccion")
            ComboBox2.Items.Add(_OrdenProduccion)
        Next
        ComboBox2.BackColor = Color.LightGreen
    End Sub

    Private Sub ComboBox2_LostFocus(sender As Object, e As System.EventArgs) Handles ComboBox2.LostFocus
        ComboBox2.BackColor = Color.White
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox5.GotFocus
        TextBox5.BackColor = Color.LightGreen

    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox8.GotFocus
        TextBox8.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox8_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox9.GotFocus
        TextBox9.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox9_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox9.LostFocus
        TextBox9.BackColor = Color.White
    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox13.GotFocus
        TextBox13.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox13_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox13.LostFocus
        TextBox13.BackColor = Color.White
    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox14_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox14.GotFocus
        TextBox14.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox14_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox14.LostFocus
        TextBox14.BackColor = Color.White
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox15_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox15.GotFocus
        TextBox15.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox15_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox15.LostFocus
        TextBox15.BackColor = Color.White
    End Sub

    Private Sub TextBox15_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox16_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox16.GotFocus
        TextBox16.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox16_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox16.LostFocus
        TextBox16.BackColor = Color.White
    End Sub

    Private Sub TextBox16_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox17_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox17.GotFocus
        TextBox17.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox17_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox17.LostFocus
        TextBox17.BackColor = Color.White
    End Sub

    Private Sub TextBox17_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub TextBox18_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox18.GotFocus
        TextBox18.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox18_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox18.LostFocus
        TextBox18.BackColor = Color.White
    End Sub

    Private Sub TextBox18_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox18.TextChanged

    End Sub

    Private Sub TextBox19_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox19.GotFocus
        TextBox19.BackColor = Color.LightGreen
    End Sub

    Private Sub TextBox19_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox19.LostFocus
        TextBox19.BackColor = Color.White
    End Sub

    Private Sub TextBox19_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub RichTextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles RichTextBox1.GotFocus
        RichTextBox1.BackColor = Color.LightGreen
    End Sub

    Private Sub RichTextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.BackColor = Color.White
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        'Imprime la portada

        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar un ENVIO para podre imprimir las portadas"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        CheckBox1.Visible = True
        CheckBox2.Visible = True
        Button2.Visible = True





    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Formatos Factor\Portada Horizontal.xlsx"

        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False

        Dim _Nombre As String = TextBox9.Text
        Dim _TelCasa As String = TextBox17.Text
        Dim _TelOfna As String = TextBox18.Text
        Dim _Celular As String = TextBox19.Text
        Dim _Telefono As String
        _Telefono = ""
        If Len(_TelCasa) > 6 Then
            _Telefono = "TEL. CASA " & _TelCasa
        End If

        If Len(_TelOfna) > 6 Then
            _Telefono = _Telefono & "  TEL. OFNA " & _TelOfna
        End If

        If Len(_Celular) > 6 Then
            _Telefono = _Telefono & "  CEL. " & _Celular
        End If


        Dim _Domicilio As String = "CALLE " & Trim(TextBox13.Text)
        _Colonia = " COL. " & Trim(TextBox14.Text)
        _Ciudad = Trim(TextBox15.Text) & ", " & Trim(TextBox16.Text)

        If CheckBox1.Checked = True Then
            m_Excel.Worksheets("VERTICAL").cells(12, 1).value = _Nombre
            m_Excel.Worksheets("VERTICAL").cells(13, 1).value = _Telefono
            m_Excel.Worksheets("VERTICAL").cells(14, 1).value = _Domicilio
            m_Excel.Worksheets("VERTICAL").cells(15, 1).value = _Colonia
            m_Excel.Worksheets("VERTICAL").cells(16, 1).value = _Ciudad

            m_Excel.Worksheets("VERTICAL").PrintOut(1, 1)
            ' m_Excel.Visible = True
        End If


        If CheckBox2.Checked = True Then
            m_Excel.Worksheets("HORIZONTAL").cells(15, 1).value = _Nombre
            m_Excel.Worksheets("HORIZONTAL").cells(16, 1).value = _Telefono
            m_Excel.Worksheets("HORIZONTAL").cells(17, 1).value = _Domicilio
            m_Excel.Worksheets("HORIZONTAL").cells(18, 1).value = _Colonia
            m_Excel.Worksheets("HORIZONTAL").cells(19, 1).value = _Ciudad

            m_Excel.Worksheets("HORIZONTAL").PrintOut(1, 1)
            'm_Excel.Visible = True
        End If

        CheckBox1.Visible = False
        CheckBox2.Visible = False
        Button2.Visible = False
        Msg = "Portada enviada a la Impresora"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

    End Sub
End Class