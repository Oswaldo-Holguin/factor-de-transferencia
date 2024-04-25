Public Class Frm61Cobranza
    Public i As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Busca As DataRow
    Public _Renglon As DataGridViewRow
    Public _Checacontrol As Control
    Public _Cobranza, _Cliente, _NombreCliente, _Referencia, _Documento As String
    Public _FechaDocumento As Date
    Public _Centro As String = LoginForm1.TextBox1.Text

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub Frm61Cobranza_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet52.DetalleCxC' Puede moverla o quitarla según sea necesario.
        Me.DetalleCxCTableAdapter.Fill(Me.CServiciosDataSet52.DetalleCxC)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet52.CxC' Puede moverla o quitarla según sea necesario.
        Me.CxCTableAdapter.Fill(Me.CServiciosDataSet52.CxC)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.Clientes' Puede moverla o quitarla según sea necesario.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet54.DetalleCobranza' Puede moverla o quitarla según sea necesario.
        Me.DetalleCobranzaTableAdapter.Fill(Me.CServiciosDataSet54.DetalleCobranza)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet54.Cobranza' Puede moverla o quitarla según sea necesario.
        Me.CobranzaTableAdapter.Fill(Me.CServiciosDataSet54.Cobranza)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)

        For Each Me._Checacontrol In TabPage1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.BackColor = Color.White
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
                _Checacontrol.BackColor = Color.White
            End If
        Next

        DataGridView1.Rows.Clear()

        ComboBox1.Items.Clear()

        For Me.i = 0 To Me.CServiciosDataSet30.Clientes.Rows.Count - 1
            ComboBox1.Items.Add(Me.CServiciosDataSet30.Clientes.Rows(i).Item("NombreCliente"))
        Next



    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        TabControl1.Visible = True
        TabPage1.Parent = TabControl1
        TabPage2.Parent = Nothing
        TabPage3.Parent = Nothing
        TextBox1.Focus()
        Timer1.Enabled = True
        Timer1.Interval = 100

        ToolStripComboBox1.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet.Centros.Rows.Count - 1
            ToolStripComboBox1.Items.Add(CServiciosDataSet.Centros.Rows(i).Item("Descripcion"))
        Next
        RadioButton1.Checked = True

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        For Each Me._Checacontrol In TabPage1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Focused Then
                    _Checacontrol.BackColor = Color.Gold
                End If
            End If
        Next
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus
        ' Obtiee el siguiente consecutivo
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("COB")
        Dim _Consecutivo As String = CStr(_Busca.Item("Consecutivo"))

        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "CO" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
        TextBox2.Text = Today.Date
        TextBox3.Text = Mid(TimeOfDay, 1, 5)
        TextBox4.Focus()


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
            _Cliente = TextBox4.Text

            Dim _Find As Boolean
            Dim _Rows_Cl() As DataRow = Me.CServiciosDataSet30.Clientes.Select("Cliente = " & "'" & _Cliente & "'")

            _Find = False
            For Me.i = 0 To _Rows_Cl.GetUpperBound(0)
                _Find = True
                ComboBox1.Text = _Rows_Cl(i).Item("NombreCliente")
                TextBox5.Text = _Rows_Cl(i).Item("Calle")
                TextBox6.Text = "COL. " & _Rows_Cl(i).Item("Colonia")
                TextBox7.Text = _Rows_Cl(i).Item("Ciudad")
                TextBox8.Text = _Rows_Cl(i).Item("Estado")
                TextBox9.Focus()
            Next

            If Not _Find Then
                Msg = "No exite ese cliente Registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If
        End If
        TextBox4.BackColor = Color.White

    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        ' Presenta los documentos peendientes de pagar

        If ToolStripTextBox1.Text = "" Then
            Msg = "Debe Seleccionar un Centro para registrar la Cobranza"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        _Centro = ToolStripTextBox1.Text
        _Cliente = TextBox4.Text
        Dim _Saldo As Double
        Dim _Importe, _ImportePagos, _Neto, _SaldoFra As Double

        _Saldo = 0
        DataGridView1.Rows.Clear()
        Dim _Rows_Cxc() As DataRow = Me.CServiciosDataSet52.CxC.Select("Cliente = " & "'" & _Cliente & "'")

        DataGridView1.Rows.Clear()
        For Me.i = 0 To _Rows_Cxc.GetUpperBound(0)
            If _Rows_Cxc(i).Item("Plazo") = _Centro Then
                _Documento = _Rows_Cxc(i).Item("Documento")
                _FechaDocumento = _Rows_Cxc(i).Item("FechaDocumento")
                _Cliente = _Rows_Cxc(i).Item("Cliente")
                _Importe = _Rows_Cxc(i).Item("Importe")
                _ImportePagos = _Rows_Cxc(i).Item("ImportePago")
                _Neto = _Importe - _ImportePagos
                _Referencia = _Rows_Cxc(i).Item("Referencia")
                _SaldoFra = _Neto
                _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                _NombreCliente = _Busca.Item("NombreCliente")

                DataGridView1.Rows.Add(_Documento, _FechaDocumento, _Cliente, _NombreCliente, _Importe, _ImportePagos, _Neto, _Referencia, False, 0, _SaldoFra)
                _Saldo = _Saldo + _Neto
            End If
        Next

        For Me.i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(0).Value = "" Then
                Exit For
            End If
            If DataGridView1.Rows(i).Cells(6).Value <= 0 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                DataGridView1.Rows(i).ReadOnly = True
            End If
        Next
        ToolStripTextBox2.Text = Format(_Saldo, "###,###,##0.00")

    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox9.GotFocus
        ToolStripButton7_Click(sender, e)
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox12.Focus()
        End If
    End Sub

    Private Sub TextBox9_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox9.LostFocus
        Dim _Cantidad As Double = TextBox9.Text
        TextBox9.BackColor = Color.White
    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ToolStripButton4_Click(sender, e)
        Timer1.Enabled = False

    End Sub

    Private Sub ToolStripComboBox1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox1.Click
      

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Dim _Index As Integer = ToolStripComboBox1.SelectedIndex
        ToolStripTextBox1.Text = Me.CServiciosDataSet.Centros.Rows(_Index).Item("Centro")
    End Sub

    Private Sub Label13_Click(sender As System.Object, e As System.EventArgs) Handles Label13.Click

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        ' Acumula las facturas seleccionadas
        Dim _Suma As Double
        Dim _Aplicado, _PorAplicar, _Neto As Double
        TextBox10.Text = 0

        If RadioButton1.Checked = True Then
            _Suma = 0
            For Me.i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value = "" Then
                    Exit For
                End If

                If DataGridView1.Rows(i).Cells(8).Value = True Then
                    _Suma = _Suma + DataGridView1.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(9).Value = DataGridView1.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(10).Value = 0
                End If

            Next

            If CDbl(TextBox9.Text) <> _Suma Then
                Msg = "La cantidad seleccionada " & CStr(_Suma) & " no coincide con el pago registrado " & CStr(TextBox9.Text)
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)

                For Me.i = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Rows(i).Cells(0).Value = "" Then
                        Exit For
                    End If
                    If DataGridView1.Rows(i).Cells(8).Value = True Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                        DataGridView1.Rows(i).Cells(9).Value = 0
                        DataGridView1.Rows(i).Cells(8).Value = False
                        DataGridView1.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(6).Value
                    End If
                Next
                TextBox10.Text = 0

                Exit Sub
            Else
                TextBox10.Text = _Suma
            End If

        End If

        ' **********************************************************************************
        ' Pagos parciales.
        If RadioButton2.Checked = True Then

            _Aplicado = 0
            _PorAplicar = CDbl(TextBox9.Text)

            For Me.i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value = "" Then
                    Exit For
                End If
                If DataGridView1.Rows(i).Cells(8).Value = True Then
                    _Neto = DataGridView1.Rows(i).Cells(6).Value
                    If _PorAplicar >= _Neto Then
                        _Aplicado = _Neto
                        TextBox10.Text = CDbl(TextBox10.Text) + _Aplicado
                        _PorAplicar = _PorAplicar - _Neto
                        DataGridView1.Rows(i).Cells(9).Value = _Aplicado
                        DataGridView1.Rows(i).Cells(10).Value = _Neto - _Aplicado
                    Else
                        _Aplicado = _PorAplicar
                        TextBox10.Text = CDbl(TextBox10.Text) + _PorAplicar
                        DataGridView1.Rows(i).Cells(9).Value = _Aplicado
                        DataGridView1.Rows(i).Cells(10).Value = _Neto - _Aplicado
                        _PorAplicar = _PorAplicar - _Aplicado
                        Exit For
                    End If

                End If

            Next

            If CDbl(TextBox9.Text) > CDbl(TextBox10.Text) Then
                Msg = "El importe seleccionado es menor al pago realizado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                ToolStripButton7_Click(sender, e)
                Exit Sub
            End If


        End If



    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox12.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox13.Focus()
        End If
    End Sub

    Private Sub TextBox12_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox12.LostFocus
        TextBox12.BackColor = Color.White
    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged
        
    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox13.GotFocus
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox13.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox13_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox13.LostFocus
      
        TextBox13.BackColor = Color.White

    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox13.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False
        TextBox11.Focus()

    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox11.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox14.Focus()
        End If
    End Sub

    Private Sub TextBox11_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox11.LostFocus
        TextBox11.BackColor = Color.White
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim _Index As Integer = ComboBox1.SelectedIndex
        TextBox4.Text = Me.CServiciosDataSet30.Clientes.Rows(_Index).Item("Cliente")
        TextBox5.Text = Me.CServiciosDataSet30.Clientes.Rows(_Index).Item("Calle")
        TextBox6.Text = "COL. " & Me.CServiciosDataSet30.Clientes.Rows(_Index).Item("Colonia")
        TextBox7.Text = Me.CServiciosDataSet30.Clientes.Rows(_Index).Item("Ciudad")
        TextBox8.Text = Me.CServiciosDataSet30.Clientes.Rows(_Index).Item("Estado")
        TextBox9.Focus()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        If _Renglon.Cells(8).Value = False Then
            _Renglon.DefaultCellStyle.BackColor = Color.Gold
            _Renglon.Cells(8).Value = True
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
            _Renglon.Cells(8).Value = False
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click

        TextBox10.Text = 0

        For Me.i = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(i).Cells(8).Value = False
            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
            DataGridView1.Rows(i).Cells(9).Value = 0
            DataGridView1.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(6).Value
        Next


    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        ' Guarda la información de la cobranza
 
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está a punto de guardar esta Cobranza. Desea continuar?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "MsgBox Demonstration"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            Dim _Find As Boolean
            Dim _ValidaFecha As String = TextBox13.Text


            ' Valida que no existan casillas en blanco
            _Find = False
            For Each Me._Checacontrol In TabPage1.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    If _Checacontrol.Text = "" Then
                        _Find = True
                        Exit For
                    End If
                End If
            Next

            If _Find Then
                Msg = "Hay casillas en blanco, debe teclear dátos válidos"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If


            ' valida la fecha 
            If Not IsDate(_ValidaFecha) Then
                Msg = "Debe teclear una fecha válida de este documento"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                TextBox13.Text = ""
                TextBox13.Focus()
                Exit Sub
            End If



            ' Guarda la cobranza
            _Busca = Me.CServiciosDataSet54.Cobranza.NewRow
            _Busca.Item("Cobranza") = TextBox1.Text
            _Busca.Item("Fecha") = CDate(TextBox2.Text)
            _Busca.Item("Hora") = TextBox3.Text
            _Busca.Item("Cliente") = TextBox4.Text
            _Busca.Item("Referencia") = TextBox14.Text
            _Busca.Item("Documento") = TextBox12.Text
            _Busca.Item("FechaDocumento") = CDate(TextBox13.Text)
            _Busca.Item("ImporteDocumento") = CDbl(TextBox9.Text)
            _Busca.Item("Autorizado") = False
            _Busca.Item("Autoriza") = ""
            _Busca.Item("FechaAutorizacion") = DBNull.Value
            _Busca.Item("Usuario") = LoginForm1.TextBox6.Text
            _Busca.Item("Centro") = LoginForm1.TextBox1.Text
            _Busca.Item("Depositado") = False
            _Busca.Item("FechaDeposito") = DBNull.Value
            _Busca.Item("Banco") = TextBox11.Text

            Me.CServiciosDataSet54.Cobranza.Rows.Add(_Busca)

            Me.Validate()
            CobranzaBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet54)
            Me.CobranzaTableAdapter.Fill(Me.CServiciosDataSet54.Cobranza)

            ' Guarda el Detalle de la cobranza
            Dim _ImportePago As Double
            Dim _Paciente As String
            Dim _ImporteDocumento As Double
            Dim _Banco As String
            Dim _IdCobranza As Integer
            _Busca = Me.CServiciosDataSet54.Cobranza.FindByCobranza(TextBox1.Text)
            _IdCobranza = _Busca.Item("Orden")

            For Me.i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value = "" Then
                    Exit For
                End If
                If DataGridView1.Rows(i).Cells(8).Value = True Then
                    _Cobranza = TextBox1.Text
                    _Documento = DataGridView1.Rows(i).Cells(0).Value
                    _FechaDocumento = CDate(DataGridView1.Rows(i).Cells(1).Value)
                    _Cliente = TextBox4.Text
                    _ImporteDocumento = CDbl(DataGridView1.Rows(i).Cells(6).Value)
                    _Banco = TextBox11.Text
                    _Paciente = ""
                    _Referencia = ComboBox1.Text
                    _ImportePago = DataGridView1.Rows(i).Cells(9).Value

                    _Busca = Me.CServiciosDataSet54.DetalleCobranza.NewRow()
                    _Busca.Item("Cobranza") = _Cobranza
                    _Busca.Item("Documento") = _Documento
                    _Busca.Item("FechaDocumento") = _FechaDocumento
                    _Busca.Item("Cliente") = _Cliente
                    _Busca.Item("ImporteDocumento") = _ImporteDocumento
                    _Busca.Item("Banco") = _Banco
                    _Busca.Item("IdCobranza") = _IdCobranza
                    _Busca.Item("Paciente") = _Paciente
                    _Busca.Item("Referencia") = _Referencia
                    _Busca.Item("ImportePago") = _ImportePago

                    Me.CServiciosDataSet54.DetalleCobranza.Rows.Add(_Busca)

                    ' Actualiza la Factura
                    _Busca = Me.CServiciosDataSet52.CxC.FindByDocumento(_Documento)

                    _Busca.Item("DocumentoPago") = TextBox12.Text
                    _Busca.Item("FechaPago") = TextBox2.Text
                    _Busca.Item("ImportePago") = _Busca.Item("ImportePago") + _ImportePago
                    If _Busca.Item("Neto") = _Busca.Item("ImportePago") Then
                        _Busca.Item("Pagada") = True
                    End If

                End If
            Next

            Me.Validate()
            DetalleCobranzaBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(Me.CServiciosDataSet54)
            Me.DetalleCobranzaTableAdapter.Fill(Me.CServiciosDataSet54.DetalleCobranza)

            Me.Validate()
            CxCBindingSource.EndEdit()
            Me.TableAdapterManager4.UpdateAll(Me.CServiciosDataSet52)
            Me.CxCTableAdapter.Fill(Me.CServiciosDataSet52.CxC)


            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("COB")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager5.UpdateAll(Me.CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)


            Msg = "Registro guardado correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Button1_Click(sender, e)

        Else
            ' Perform some other action.
        End If


    End Sub

    Private Sub TextBox14_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox14.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Button2.Focus()
        End If
    End Sub

    Private Sub TextBox14_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox14.LostFocus
        TextBox14.BackColor = Color.White
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TabPage1_Click(sender As System.Object, e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        Dim _ImporteCobranza As Double

        For Me.i = 0 To Me.CServiciosDataSet54.Cobranza.Rows.Count - 1
            _Cobranza = Me.CServiciosDataSet54.Cobranza.Rows(i).Item("Cobranza")
            _FechaDocumento = Me.CServiciosDataSet54.Cobranza.Rows(i).Item("Fecha")
            _ImporteCobranza = Me.CServiciosDataSet54.Cobranza.Rows(i).Item("ImporteDocumento")
            _Cliente = Me.CServiciosDataSet54.Cobranza.Rows(i).Item("Cliente")
            _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
            _NombreCliente = _Busca.Item("NombreCliente")

            DataGridView2.Rows.Add(_Cobranza, _FechaDocumento, _ImporteCobranza, _NombreCliente)

        Next

    End Sub

    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click
        TabControl1.Visible = True
        TabPage1.Parent = Nothing
        TabPage2.Parent = TabControl1
        TabPage3.Parent = Nothing

        Button3_Click(sender, e)

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If


        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _Cobranza = _Renglon.Cells(0).Value
        _Busca = Me.CServiciosDataSet54.Cobranza.FindByCobranza(_Cobranza)
        TextBox15.Text = _Cobranza
        TextBox16.Text = _Busca.Item("Fecha")
        TextBox17.Text = _Busca.Item("Hora")
        TextBox18.Text = _Busca.Item("Cliente")
        TextBox23.Text = _Busca.Item("ImporteDocumento")
        TextBox24.Text = _Busca.Item("Banco")
        TextBox25.Text = _Busca.Item("Documento")
        TextBox26.Text = _Busca.Item("FechaDocumento")
        
        _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
        TextBox19.Text = _Busca.Item("NombreCliente")
        TextBox20.Text = _Busca.Item("Calle")
        TextBox21.Text = "COL. " & _Busca.Item("Colonia")
        TextBox22.Text = _Busca.Item("Ciudad") & ", " & _Busca.Item("Estado")


        ' Presenta el detalle
        Dim _Rows_Dc() As DataRow = Me.CServiciosDataSet54.DetalleCobranza.Select("Cobranza = " & "'" & _Cobranza & "'")
        Dim _ImporteDocumento, _ImporteAplicado, _Saldo As Double
        Dim _Banco As String


        DataGridView3.Rows.Clear()
        For Me.i = 0 To _Rows_Dc.GetUpperBound(0)
            _Documento = _Rows_Dc(i).Item("Documento")
            _FechaDocumento = _Rows_Dc(i).Item("FechaDocumento")
            _ImporteDocumento = _Rows_Dc(i).Item("ImporteDocumento")
            _ImporteAplicado = _Rows_Dc(i).Item("ImportePago")
            _Saldo = _ImporteDocumento - _ImporteAplicado
            _Banco = _Rows_Dc(i).Item("Banco")
            _Referencia = _Rows_Dc(i).Item("Referencia")

            DataGridView3.Rows.Add(_Cobranza, _Documento, _FechaDocumento, _ImporteDocumento, _ImporteAplicado, _Saldo, _Banco, _Referencia)

        Next





    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellLeave
        _Renglon = DataGridView2.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        DataGridView2.Visible = False
        Panel1.Visible = True

        For Each Me._Checacontrol In TabPage2.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        DataGridView3.Rows.Clear()

    End Sub

    Private Sub Label26_Click(sender As System.Object, e As System.EventArgs) Handles Label26.Click

        ' Búsqueda
        Dim _Rows_Co() As DataRow

        ' Por Número de Cobranza
        If RadioButton4.Checked = True Then
            _Cobranza = UCase(InputBox("Teclee el número de Cobranza a buscar"))

            If _Cobranza = "" Then
                Msg = "Debe teclear un número válido de Cobranza"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            _Rows_Co = Me.CServiciosDataSet54.Cobranza.Select("Cobranza = " & "'" & _Cobranza & "'")

        End If

        ' Por Fecha
        If RadioButton5.Checked = True Then
            _FechaDocumento = InputBox("Proporcione la fecha a buscar")
            If Not IsDate(_FechaDocumento) Then
                Msg = "Debe teclear una fecha válida para buscar"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            _Rows_Co = Me.CServiciosDataSet54.Cobranza.Select("Fecha = " & "'" & _FechaDocumento & "'")
        End If



    End Sub

    Private Sub ToolStripButton11_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton11.Click

        TabControl1.Visible = True
        TabPage1.Parent = Nothing
        TabPage2.Parent = Nothing
        TabPage3.Parent = TabControl1

        ComboBox2.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet30.Clientes.Rows.Count - 1
            ComboBox2.Items.Add(Me.CServiciosDataSet30.Clientes.Rows(i).Item("NombreCliente"))
        Next

        ToolStripComboBox2.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet.Centros.Rows.Count - 1
            ToolStripComboBox2.Items.Add(Me.CServiciosDataSet.Centros.Rows(i).Item("Descripcion"))
        Next

    End Sub

    Private Sub TextBox27_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox27.GotFocus
        TextBox27.BackColor = Color.Gold
    End Sub

    Private Sub TextBox27_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox27.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Dim _Find As Boolean
            TextBox27.BackColor = Color.White
            _Find = False
            For Me.i = 0 To Me.CServiciosDataSet30.Clientes.Rows.Count - 1
                If Me.CServiciosDataSet30.Clientes.Rows(i).Item("Cliente") = TextBox27.Text Then
                    ComboBox2.Text = Me.CServiciosDataSet30.Clientes.Rows(i).Item("NombreCliente")
                    _Find = True
                End If
            Next

            If Not _Find Then
                Msg = "No existe ese cliente registrado. Verifique o busque en el Combo"
                Style = MsgBoxStyle.Information
                ComboBox2.Focus()
            End If
            TextBox28.Focus()
        End If
    End Sub

    Private Sub TextBox27_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox27.LostFocus
        TextBox27.BackColor = Color.White
    End Sub

    Private Sub TextBox27_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox27.TextChanged

    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox2.GotFocus
        ComboBox2.BackColor = Color.Gold
    End Sub

    Private Sub ComboBox2_LostFocus(sender As Object, e As System.EventArgs) Handles ComboBox2.LostFocus
        ComboBox2.BackColor = Color.White
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim _Indice As Integer = ComboBox2.SelectedIndex
        TextBox27.Text = Me.CServiciosDataSet30.Clientes.Rows(_Indice).Item("Cliente")

    End Sub

    Private Sub ToolStripButton9_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton9.Click
        ' Genera listado.
        If RadioButton14.Checked = True Then
            ' Kardex
            If TextBox27.Text = "" Then
                Msg = "Debe seleccionar un Cliente para generar un Kardex"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If


            _Cliente = TextBox27.Text
            Dim _Rows_CxC() As DataRow = Me.CServiciosDataSet52.CxC.Select("Cliente = " & "'" & _Cliente & "'")
            Dim _Descripcion As String
            Dim _Cargos, _ABonos, _Saldo As Double
            Dim _Rows_Dc() As DataRow
            Dim X As Integer

            For Me.i = 0 To _Rows_CxC.GetUpperBound(0)
                If _Rows_CxC(i).Item("Plazo") = ToolStripTextBox3.Text Then
                    _Documento = _Rows_CxC(i).Item("Documento")
                    _FechaDocumento = _Rows_CxC(i).Item("FechaDocumento")

                    _Descripcion = ""
                    If _Rows_CxC(i).Item("Plazo") = "FTR" Then
                        _Descripcion = "VENTA DE DOSIS DE FACTOR DE TRANSFERENCIA"
                    End If
                    If _Rows_CxC(i).Item("Plazo") = "CME" Then
                        _Descripcion = "VENTA DE NUTRICIONES PARENTERALES"
                    End If
                    _Cargos = _Rows_CxC(i).Item("Neto")
                    _ABonos = 0
                    _Saldo = 0

                    DataGridView4.Rows.Add(_FechaDocumento, _Documento, _Descripcion, _Cargos, _ABonos, _Saldo)


                    ' Busca los pagos
                    _Rows_Dc = Me.CServiciosDataSet54.DetalleCobranza.Select("Documento = " & "'" & _Documento & "'")

                    For X = 0 To _Rows_Dc.GetUpperBound(0)
                        _Documento = _Rows_Dc(X).Item("Cobranza")
                        _Cobranza = _Rows_Dc(X).Item("Cobranza")
                        _Busca = Me.CServiciosDataSet54.Cobranza.FindByCobranza(_Cobranza)
                        _FechaDocumento = _Busca.Item("Fecha")
                        _Descripcion = _Busca.Item("Documento") & " APLICADO A LA FRA " & _Rows_Dc(X).Item("Documento")
                        _Cargos = 0
                        _ABonos = _Rows_Dc(X).Item("ImportePago")
                        _Saldo = 0

                        DataGridView4.Rows.Add(_FechaDocumento, _Documento, _Descripcion, _Cargos, _ABonos, _Saldo)

                    Next

                End If

            Next

            Me.DataGridView4.Sort(Me.DataGridView4.Columns("Column24"), System.ComponentModel.ListSortDirection.Ascending)

            _Saldo = 0
            For Me.i = 0 To DataGridView4.Rows.Count - 1
                If DataGridView4.Rows(i).Cells(1).Value = "" Then
                    Exit For
                End If
                _Saldo = _Saldo + CDbl(DataGridView4.Rows(i).Cells(3).Value) - CDbl(DataGridView4.Rows(i).Cells(4).Value)
                DataGridView4.Rows(i).Cells(5).Value = _Saldo

            Next

        End If

        ' Genera el reporte
        Button5_Click(sender, e)

    End Sub

    Private Sub TextBox28_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox28.GotFocus
        TextBox28.BackColor = Color.Gold
        MonthCalendar2.Visible = True
        TextBox30.Text = "1"

    End Sub

    Private Sub TextBox28_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox28.LostFocus
        TextBox28.BackColor = Color.White

    End Sub

    Private Sub TextBox28_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox28.TextChanged

    End Sub

    Private Sub MonthCalendar2_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateChanged

    End Sub

    Private Sub MonthCalendar2_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateSelected
        If TextBox30.Text = "1" Then
            TextBox28.Text = MonthCalendar2.SelectionRange.Start
            MonthCalendar2.Visible = False
            TextBox29.Focus()
        Else
            TextBox29.Text = MonthCalendar2.SelectionRange.Start
            MonthCalendar2.Visible = False
            TextBox29.BackColor = Color.White
        End If


    End Sub

    Private Sub ToolStripComboBox2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox2.Click

    End Sub

    Private Sub ToolStripComboBox2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        Dim _Indice As Integer = ToolStripComboBox2.SelectedIndex
        ToolStripTextBox3.Text = Me.CServiciosDataSet.Centros.Rows(_Indice).Item("Centro")
    End Sub

    Private Sub TextBox29_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox29.GotFocus
        MonthCalendar2.Visible = True

        TextBox29.BackColor = Color.Gold
        TextBox30.Text = "2"
    End Sub

    Private Sub TextBox29_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox29.TextChanged

    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        ' Genera el reporte
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Formatos Mezclas\Formato Kardex.xlsx"
        Dim _Linea As Integer
        Dim _Fi As Date = DataGridView4.Rows(0).Cells(0).Value
        Dim _FechaInicial As Date = CDate(TextBox28.Text)
        Dim _FechaFinal As Date = CDate(TextBox29.Text)

        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False

        m_Excel.Worksheets("KARDEX").cells(4, 7).value = TextBox28.Text
        m_Excel.Worksheets("KARDEX").cells(5, 7).value = TextBox29.Text
        m_Excel.Worksheets("KARDEX").cells(9, 2).value = TextBox27.Text
        m_Excel.Worksheets("KARDEX").cells(9, 3).value = ComboBox2.Text

        _Linea = 12

        For Me.i = 0 To DataGridView4.Rows.Count - 1
            If DataGridView4.Rows(i).Cells(1).Value = "" Then
                Exit For
            End If

            m_Excel.Worksheets("KARDEX").cells(_Linea, 1).value = DataGridView4.Rows(i).Cells(0).Value
            m_Excel.Worksheets("KARDEX").cells(_Linea, 2).value = DataGridView4.Rows(i).Cells(1).Value
            m_Excel.Worksheets("KARDEX").cells(_Linea, 3).value = DataGridView4.Rows(i).Cells(2).Value
            m_Excel.Worksheets("KARDEX").cells(_Linea, 7).value = DataGridView4.Rows(i).Cells(3).Value
            m_Excel.Worksheets("KARDEX").cells(_Linea, 8).value = DataGridView4.Rows(i).Cells(4).Value
            m_Excel.Worksheets("KARDEX").cells(_Linea, 9).value = DataGridView4.Rows(i).Cells(5).Value
            _Linea = _Linea + 1

        Next

        m_Excel.Visible = True


    End Sub
End Class