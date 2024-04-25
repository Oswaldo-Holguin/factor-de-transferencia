Public Class Frm28VentaFT
    Public i As Integer
    Public msg, _Medico, _Cedula, _Recibe As String
    Public Style As MsgBoxStyle
    Public Checacontrol As Control
    Public _Renglon As DataGridViewRow
    Public _Folio, _Paciente, _NombrePaciente, _Referencia, _Comentarios, _Caja, _HoraRegistro As String
    Public _NumeroFrascos, _PrecioUnitario, _ImporteDescuento, _PrecioNeto As Integer
    Public _PorcentajeDescuento As Integer
    Public _FechaFolio As String
    Public strRutaExcel As String
    Public _Diagnostico As String
    Public _Consecutivo, _Mide As Integer
    Public _XConsecutivo As String
    Public _Busca As DataRow
    Public _MideNombre As Integer
    Public _NuevoDiagnostico As Boolean
    Public m_Excel As Microsoft.Office.Interop.Excel.Application

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Frm28VentaFT_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Escape) Then
            Label18.Text = "."
            Label18.ForeColor = Color.LightSteelBlue
            Button1_Click(sender, e)
        End If
    End Sub

    Private Sub Form28_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load



        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet22.VentaFactor' Puede moverla o quitarla según sea necesario.
        '    Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet22.VentaFactor)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet21.Productos' Puede moverla o quitarla según sea necesario.
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet23.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter1.Fill(Me.CServiciosDataSet23.VentaFactor)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet32.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter2.Fill(Me.CServiciosDataSet32.VentaFactor)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet26.Produccion' Puede moverla o quitarla según sea necesario.
        Me.ProduccionTableAdapter.Fill(Me.CServiciosDataSet26.Produccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.Clientes' Puede moverla o quitarla según sea necesario.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet40.Recibos' Puede moverla o quitarla según sea necesario.
        Me.RecibosTableAdapter.Fill(Me.CServiciosDataSet40.Recibos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Diagnosticos' Puede moverla o quitarla según sea necesario.
        Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet2.Diagnosticos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""

        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                Checacontrol.Text = ""
                Checacontrol.BackColor = Color.White
            End If
            If TypeOf Checacontrol Is ComboBox Then
                Checacontrol.Text = ""
                Checacontrol.BackColor = Color.White
            End If
        Next
        DataGridView1.Rows.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""

        Dim _Cantidad, X As Integer
        Dim _Rows_Vf() As DataRow
        Dim _Lote As String
        Dim _Xfecha As String
        Dim _dia As Integer = Int(Mid(CStr(Today), 1, 2))
        Dim _Fecha As Date
        _Fecha = Format(Today, "Short Date")
        _Cantidad = 0

        For Me.i = _dia To 1 Step -1
            _Xfecha = CStr(_Fecha)
            ToolStripComboBox1.Items.Add(_Xfecha)
            _Fecha = _Fecha.AddDays(-1)
        Next
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        Panel1.Enabled = False
        RadioButton1.Checked = True

        ToolStripComboBox2.Items.Clear()

        For Me.i = 0 To CServiciosDataSet26.Produccion.Rows.Count - 1
            _Lote = CServiciosDataSet26.Produccion.Rows(i).Item("Lote")
            _Cantidad = CServiciosDataSet26.Produccion.Rows(i).Item("Cantidad")

            _NumeroFrascos = 0
            _Rows_Vf = CServiciosDataSet23.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
            For X = 0 To _Rows_Vf.GetUpperBound(0)
                _NumeroFrascos = _NumeroFrascos + _Rows_Vf(X).Item("NumeroFrascos")
            Next

            If _NumeroFrascos < _Cantidad Then
                ToolStripComboBox2.Items.Add(_Lote)
            End If

        Next

        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet30.Clientes.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet30.Clientes.Rows(i).Item("NombreCliente"))
        Next

        ToolStripButton8.Enabled = False

        ComboBox3.Items.Clear()
        ComboBox3.Items.Add("N/A")
        For Me.i = 0 To 120
            ComboBox3.Items.Add(i)
        Next

        ComboBox4.Items.Clear()
        ComboBox4.Items.Add("")
        ComboBox4.Items.Add(" *** Agregar Nuevo Diagnóstico ***")
        For Me.i = 0 To CServiciosDataSet2.Diagnosticos.Rows.Count - 1
            ComboBox4.Items.Add(CServiciosDataSet2.Diagnosticos.Rows(i).Item("Descripcion"))
        Next
        ComboBox4.Sorted = True
        TextBox11.Text = "JR"


    End Sub

    Private Sub ToolStripComboBox1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_LostFocus(sender As Object, e As System.EventArgs) Handles ToolStripComboBox1.LostFocus
      
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

    Private Sub TextBox3_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.Text = UCase(TextBox3.Text)
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click

        If ToolStripComboBox1.Text = "Seleccione la fecha" Then
            msg = "Debe seleccionar una fecha de venta"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If
        If ToolStripTextBox1.Text = "Seleccione el Lote" Then
            msg = "Debe proporcionar un Número de Lote"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub

        End If

        If Not IsDate(ToolStripComboBox1.Text) Then
            msg = "Debe seleccionar una fecha válida"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If



        ToolStripTextBox1.Text = UCase(ToolStripTextBox1.Text)

        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                Checacontrol.Text = ""
            End If
        Next
        ComboBox1.Text = ""
        ComboBox2.Text = ""
      

        Panel1.Enabled = True
        TextBox1.Focus()
        ' Busca el precio unitario del Factor
        For Me.i = 0 To CServiciosDataSet21.Productos.Rows.Count - 1
            If CServiciosDataSet21.Productos.Rows(i).Item("Producto") = "FTR" Then
                _PrecioUnitario = CServiciosDataSet21.Productos.Rows(i).Item("PrecioUnitario")
            End If
        Next

        TextBox4.Text = _PrecioUnitario

        ' Valida los regsitros de esa fecha
        Dim _Rows_VF() As DataRow = CServiciosDataSet32.VentaFactor.Select("FechaFolio = " & "'" & CDate(ToolStripComboBox1.Text) & "'")
        Dim _id As Integer
        DataGridView1.Rows.Clear()
        For Me.i = 0 To _Rows_VF.GetUpperBound(0)
            _Folio = _Rows_VF(i).Item("folio")
            _FechaFolio = _Rows_VF(i).Item("FechaFolio")
            If DBNull.Value.Equals(_Rows_VF(i).Item("Paciente")) Then
                _Paciente = ""
            Else
                _Paciente = _Rows_VF(i).Item("Paciente")
            End If

            _NombrePaciente = _Rows_VF(i).Item("NombrePaciente")
            _NumeroFrascos = _Rows_VF(i).Item("NumeroFrascos")
            _PrecioUnitario = _Rows_VF(i).Item("PrecioUnitario")
            _PorcentajeDescuento = _Rows_VF(i).Item("PorcentajeDescuento")
            _ImporteDescuento = _Rows_VF(i).Item("ImporteDescuento")
            _PrecioNeto = _Rows_VF(i).Item("PrecioNeto")
            _id = _Rows_VF(i).Item("Orden")
            _Medico = _Rows_VF(i).Item("Medico")
            _Cedula = _Rows_VF(i).Item("Cedula")
            _Recibe = _Rows_VF(i).Item("Comentarios")

            DataGridView1.Rows.Add(_Folio, _FechaFolio, _Paciente, _NombrePaciente, _NumeroFrascos, _PrecioUnitario, _PorcentajeDescuento, _ImporteDescuento, _PrecioNeto, _id, _Medico, _Cedula, _Recibe)
        Next
        CheckBox1.Checked = True
        TextBox11.Text = "ALTA"

        Label18.Text = "Presione ESC para deshacer la Captura "
        Label18.ForeColor = Color.Maroon
        Timer1.Enabled = True
        Timer1.Interval = 100
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        If Asc(e.KeyChar) = 13 Then
            ComboBox4.Focus()
        End If

    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox6.LostFocus
        Dim _Valor As Integer = TextBox6.Text
        If Not IsNumeric(_Valor) Then
            msg = "Debe teclear un importe válido"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        Dim _Diferencia As Integer

        TextBox9.Text = Int(TextBox6.Text)
        TextBox6.Text = Format(_Valor, "$###,##0")

        _Diferencia = 0
        ' Determina si hay descuento en el precio
        _NumeroFrascos = CInt(TextBox5.Text)
        _PrecioNeto = _NumeroFrascos * _PrecioUnitario

        If _PrecioNeto = CInt(TextBox9.Text) Then
            TextBox7.Text = 0
            TextBox8.Text = 0

        Else

            _Diferencia = _PrecioNeto - CInt(TextBox9.Text)
            _ImporteDescuento = _Diferencia
            TextBox8.Text = _Diferencia
            _PorcentajeDescuento = (_Diferencia / _PrecioNeto) * 100
            TextBox7.Text = _PorcentajeDescuento
        End If
        ToolStripButton3.Enabled = True
        TextBox6.BackColor = Color.White
    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub ToolStripTextBox1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                If Checacontrol.Text = "" Then
                    msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
                    Style = MsgBoxStyle.Information
                    MsgBox(msg, Style)
                    Exit Sub
                End If
            End If
        Next

        ' Valida que no exista ese registro

        Dim _Rows_VF() As DataRow = CServiciosDataSet32.VentaFactor.Select("Folio = " & "'" & TextBox1.Text & "'")
        If TextBox11.Text = "ALTA" Then

            Dim _Find As Boolean

            _Find = False
            For Me.i = 0 To _Rows_VF.GetUpperBound(0)
                _Find = True
            Next

            If _Find Then
                msg = "Ese Folio ya está registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If


        End If

        ' Verficia que no exceda el número de frascos producidos
        Dim _FrascosProducidos As Integer
        Dim _FrascosVendidos As Integer
        Dim _Busca As DataRow
        Dim _Lote As String = ToolStripTextBox1.Text
        Dim _OrdenProduccion As String = ToolStripTextBox1.Text
        Dim _Rows_VentaFactor() As DataRow = CServiciosDataSet32.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
        Dim _FrascosPorVender As Integer

        _FrascosProducidos = 0
        _Busca = CServiciosDataSet26.Produccion.FindByOrdenProduccion(_OrdenProduccion)
        _FrascosProducidos = _Busca.Item("Cantidad")

        _FrascosVendidos = 0
        For Me.i = 0 To _Rows_VentaFactor.GetUpperBound(0)
            If _Rows_VentaFactor(i).Item("Referencia") = _Lote Then
                _FrascosVendidos = _FrascosVendidos + _Rows_VentaFactor(i).Item("NumeroFrascos")
            End If
        Next

        _FrascosPorVender = _FrascosProducidos - _FrascosVendidos

        If _FrascosVendidos >= _FrascosProducidos Then
            msg = "Ya se ha vendido la totalidad del Lote  " & _Lote & ". Verifique la información"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        Else
            If _FrascosPorVender < CInt(TextBox5.Text) Then
                msg = "El número de frascos a vender es mayor a la existencia de este Lote, que es " & CStr(_FrascosPorVender) & ". Verfifique la información"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If
        End If


        ' Da de alta el registro

        If TextBox11.Text = "ALTA" Then
            _Busca = CServiciosDataSet32.VentaFactor.NewRow
            _Busca.Item("Folio") = TextBox1.Text
        End If
        If TextBox11.Text = "CAMBIO" Then
            _Busca = CServiciosDataSet32.VentaFactor.FindByFolio(TextBox1.Text)
        End If

        _Busca.Item("FechaFolio") = ToolStripComboBox1.Text
        _Busca.Item("Paciente") = TextBox12.Text
        _Busca.Item("NombrePaciente") = TextBox3.Text
        _Busca.Item("NumeroFrascos") = CInt(TextBox5.Text)
        _Busca.Item("PrecioUnitario") = CInt(TextBox4.Text)
        _Busca.Item("PorcentajeDescuento") = CInt(TextBox7.Text)
        _Busca.Item("ImporteDescuento") = CInt(TextBox8.Text)
        _Busca.Item("PrecioNeto") = CInt(TextBox9.Text)
        _Busca.Item("Referencia") = ToolStripTextBox1.Text
        _Busca.Item("FechaRegistro") = Today
        _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
        _Busca.Item("Diagnostico") = UCase(TextBox10.Text)
        '  _Busca.Item("Caja") = ""
        _Busca.Item("Medico") = UCase(TextBox13.Text)
        _Busca.Item("cedula") = UCase(TextBox14.Text)
        _Busca.Item("Comentarios") = UCase(TextBox15.Text)
        _Busca.Item("Caja") = TextBox19.Text

        _Busca.Item("Consecutivo") = TextBox2.Text
        If CheckBox1.Checked = True Then
            _Busca.Item("Contado") = True
        Else
            _Busca.Item("Contado") = False
        End If

        If TextBox11.Text = "ALTA" Then
            CServiciosDataSet32.VentaFactor.Rows.Add(_Busca)
        End If


        Me.Validate()
        VentaFactorBindingSource2.EndEdit()
        Me.TableAdapterManager3.UpdateAll(CServiciosDataSet32)
        Me.VentaFactorTableAdapter2.Fill(Me.CServiciosDataSet32.VentaFactor)

        ' Actualiza el Consecutivo de los Diagnósticos así como la tabla de los mismos
        If _NuevoDiagnostico = True Then
            ' _Busca el siguiente Consecutivo
        
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("DGN")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager8.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)


            ' Registra el nuevo Diagn+ostico
            _Busca = CServiciosDataSet2.Diagnosticos.NewRow
            _Busca.Item("Diagnostico") = _XConsecutivo
            _Busca.Item("Descripcion") = _Diagnostico
            _Busca.Item("Centro") = "FTR"
            _Busca.Item("Referencia") = "VENTA DE FACTOR DE TRANSFERENCIA"
            _Busca.Item("Comentario") = "DIAGNOSTICO AUTOMATICO"
            _Busca.Item("FechaRegistro") = CStr(Today.Date)
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)

            Me.CServiciosDataSet2.Diagnosticos.Rows.Add(_Busca)
            Me.Validate()
            DiagnosticosBindingSource.EndEdit()
            Me.TableAdapterManager7.UpdateAll(CServiciosDataSet2)
            Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet2.Diagnosticos)


        End If



        msg = "Registro guardado correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(msg, Style)

        Button1_Click(sender, e)



    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox10.GotFocus
        TextBox10.Text = "N/A"
    End Sub

    Private Sub TextBox10_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox10.LostFocus
        TextBox10.Text = UCase(TextBox10.Text)
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim _Fecha As String
        Dim _Rows_VF() As DataRow
        Dim _Lote As String
        Dim _Consecutivo As String
        Dim _Parte, _Texto, _ParteNombre As String
        Dim _Mide As Integer
        Dim _Id As Integer

        If CheckBox2.Checked = False Then
            For Each Me.Checacontrol In Panel1.Controls
                If TypeOf Checacontrol Is TextBox Then
                    Checacontrol.Text = ""
                End If
            Next
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            CheckBox1.Checked = False
            Panel1.Enabled = False
        End If


        If CheckBox2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            RadioButton5.Checked = False
            RadioButton6.Checked = False
            RadioButton7.Checked = False
            RadioButton8.Checked = False
            RadioButton9.Checked = False
            RadioButton10.Checked = False
            _Parte = ComboBox2.Text
            _Parte = Trim(UCase(_Parte))
            _Mide = Len(_Parte)
            If _Mide = 0 Then
                Exit Sub
            End If

            '    Dim _MideNombre As Integer
            ComboBox2.Items.Clear()
            For Me.i = 0 To CServiciosDataSet32.VentaFactor.Rows.Count - 1
                _Texto = Trim(CServiciosDataSet32.VentaFactor.Rows(i).Item("NombrePaciente"))
                _Texto = UCase(_Texto)
                _FechaFolio = CServiciosDataSet32.VentaFactor.Rows(i).Item("FechaFolio")

                For x = 1 To Len(_Texto)
                    _ParteNombre = Mid(_Texto, x, _Mide)
                    If _ParteNombre = _Parte Then

                        _NombrePaciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("NombrePaciente")
                        _MideNombre = Len(_NombrePaciente)

                        ComboBox2.Items.Add(_NombrePaciente)

                    End If
                Next

            Next
            ComboBox2.DroppedDown = True
        End If


        DataGridView1.Rows.Clear()
        ' Busquedas por Fecha
        If RadioButton1.Checked = True Then
            _Fecha = InputBox("Teclee la Fecha a buscar")

            If _Fecha = "" Then
                msg = "Debe teclear una fecha Válida de búsqueda"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If
            If Not IsDate(_Fecha) Then
                msg = "Debe teclear una fecha válida de búsqueda"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If

            _Rows_VF = CServiciosDataSet32.VentaFactor.Select("FechaFolio  = " & "'" & _Fecha & "'")
        End If

        ' Busquedas por Lote
        If RadioButton2.Checked = True Then
            _Lote = InputBox("Teclee el LOTE a buscar")
            _Lote = UCase(_Lote)
            _Rows_VF = CServiciosDataSet32.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
        End If

        ' Busquedas por Folio
        If RadioButton3.Checked = True Then
            _Folio = InputBox("Teclee el FOLIO a buscar")
            _Folio = UCase(_Folio)
            _Rows_VF = CServiciosDataSet32.VentaFactor.Select("Folio = " & "'" & _Folio & "'")
        End If

        ' Busquedas por Cédula
        If RadioButton9.Checked = True Then
            _Cedula = InputBox("Teclee la CEDULA a buscar")
            _Cedula = UCase(_Cedula)
            _Rows_VF = CServiciosDataSet32.VentaFactor.Select("Cedula = " & "'" & _Cedula & "'")
        End If


        ' Busquedas por Nombre
        If RadioButton4.Checked = True Then
            _Parte = InputBox("Teclee el NOMBRE a buscar")
            _Parte = Trim(UCase(_Parte))
            _Mide = Len(_Parte)
            If _Mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet32.VentaFactor.Rows.Count - 1
                _Texto = Trim(CServiciosDataSet32.VentaFactor.Rows(i).Item("NombrePaciente"))
                _Texto = UCase(_Texto)

                For x = 1 To Len(_Texto)
                    _ParteNombre = Mid(_Texto, x, _Mide)
                    If _ParteNombre = _Parte Then
                        _Folio = CServiciosDataSet32.VentaFactor.Rows(i).Item("Folio")
                        _FechaFolio = CServiciosDataSet32.VentaFactor.Rows(i).Item("FechaFolio")
                        _Paciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("Paciente")
                        _NombrePaciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("NombrePaciente")
                        _NumeroFrascos = CServiciosDataSet32.VentaFactor.Rows(i).Item("NumeroFrascos")
                        _PrecioUnitario = CServiciosDataSet32.VentaFactor.Rows(i).Item("PrecioUnitario")
                        _PorcentajeDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("PorcentajeDescuento")
                        _ImporteDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("ImporteDescuento")
                        _Medico = CServiciosDataSet32.VentaFactor.Rows(i).Item("Medico")
                        _Cedula = CServiciosDataSet32.VentaFactor.Rows(i).Item("Cedula")
                        _Recibe = CServiciosDataSet32.VentaFactor.Rows(i).Item("Comentarios")
                        _Id = CServiciosDataSet32.VentaFactor.Rows(i).Item("Orden")

                        DataGridView1.Rows.Add(_Folio, _FechaFolio, _Paciente, _NombrePaciente, _NumeroFrascos, _PrecioUnitario, _PorcentajeDescuento, _ImporteDescuento, _PrecioNeto, _Id, _Medico, _Cedula, _Recibe)

                    End If
                Next

            Next


        End If

        ' Búesqueda por Nombre del Médico

        If RadioButton8.Checked = True Then
            _Parte = InputBox("Teclee el NOMBRE del Médico a buscar")
            _Parte = Trim(UCase(_Parte))
            _Mide = Len(_Parte)
            If _Mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet32.VentaFactor.Rows.Count - 1
                _Texto = UCase(Trim(CServiciosDataSet32.VentaFactor.Rows(i).Item("Medico")))
                _Texto = UCase(_Texto)

                For x = 1 To Len(_Texto)
                    _ParteNombre = Mid(_Texto, x, _Mide)
                    If _ParteNombre = _Parte Then
                        _Folio = CServiciosDataSet32.VentaFactor.Rows(i).Item("Folio")
                        _FechaFolio = CServiciosDataSet32.VentaFactor.Rows(i).Item("FechaFolio")
                        _Paciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("Paciente")
                        _NombrePaciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("NombrePaciente")
                        _NumeroFrascos = CServiciosDataSet32.VentaFactor.Rows(i).Item("NumeroFrascos")
                        _PrecioUnitario = CServiciosDataSet32.VentaFactor.Rows(i).Item("PrecioUnitario")
                        _PorcentajeDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("PorcentajeDescuento")
                        _ImporteDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("ImporteDescuento")
                        _Medico = CServiciosDataSet32.VentaFactor.Rows(i).Item("Medico")
                        _Cedula = CServiciosDataSet32.VentaFactor.Rows(i).Item("Cedula")
                        _Recibe = CServiciosDataSet32.VentaFactor.Rows(i).Item("Comentarios")
                        _Id = CServiciosDataSet32.VentaFactor.Rows(i).Item("Orden")

                        DataGridView1.Rows.Add(_Folio, _FechaFolio, _Paciente, _NombrePaciente, _NumeroFrascos, _PrecioUnitario, _PorcentajeDescuento, _ImporteDescuento, _PrecioNeto, _Id, _Medico, _Cedula, _Recibe)

                    End If
                Next

            Next

        End If

        ' Búsqueda por Nombre del Comprador

        If RadioButton10.Checked = True Then
            _Parte = InputBox("Teclee el NOMBRE del Comprador a buscar")
            _Parte = Trim(UCase(_Parte))
            _Mide = Len(_Parte)
            If _Mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet32.VentaFactor.Rows.Count - 1
                _Texto = Trim(CServiciosDataSet32.VentaFactor.Rows(i).Item("Comentarios"))
                _Texto = UCase(_Texto)
                For x = 1 To Len(_Texto)
                    _ParteNombre = Mid(_Texto, x, _Mide)
                    If _ParteNombre = _Parte Then
                        _Folio = CServiciosDataSet32.VentaFactor.Rows(i).Item("Folio")
                        _FechaFolio = CServiciosDataSet32.VentaFactor.Rows(i).Item("FechaFolio")
                        _Paciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("Paciente")
                        _NombrePaciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("NombrePaciente")
                        _NumeroFrascos = CServiciosDataSet32.VentaFactor.Rows(i).Item("NumeroFrascos")
                        _PrecioUnitario = CServiciosDataSet32.VentaFactor.Rows(i).Item("PrecioUnitario")
                        _PorcentajeDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("PorcentajeDescuento")
                        _ImporteDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("ImporteDescuento")
                        _Medico = CServiciosDataSet32.VentaFactor.Rows(i).Item("Medico")
                        _Cedula = CServiciosDataSet32.VentaFactor.Rows(i).Item("Cedula")
                        _Recibe = CServiciosDataSet32.VentaFactor.Rows(i).Item("Comentarios")
                        _Id = CServiciosDataSet32.VentaFactor.Rows(i).Item("Orden")

                        DataGridView1.Rows.Add(_Folio, _FechaFolio, _Paciente, _NombrePaciente, _NumeroFrascos, _PrecioUnitario, _PorcentajeDescuento, _ImporteDescuento, _PrecioNeto, _Id, _Medico, _Cedula, _Recibe)

                    End If
                Next

            Next


        End If



        ' Busquedas por Consecutivo
        If RadioButton5.Checked = True Then
            _Consecutivo = InputBox("Teclee el CONSECUTIVO a buscar")
            _Rows_VF = CServiciosDataSet32.VentaFactor.Select("Consecutivo = " & "'" & _Consecutivo & "'")
        End If

        If RadioButton7.Checked = True Then
            _Parte = InputBox("Teclee el DIAGNOSTICO a buscar")
            _Parte = Trim(UCase(_Parte))
            _Mide = Len(_Parte)
            If _Mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet32.VentaFactor.Rows.Count - 1
                If DBNull.Value.Equals(CServiciosDataSet32.VentaFactor.Rows(i).Item("Diagnostico")) Then
                    _Texto = ""
                Else
                    _Texto = Trim(CServiciosDataSet32.VentaFactor.Rows(i).Item("Diagnostico"))
                End If
                _Texto = UCase(_Texto)

                For x = 1 To Len(_Texto)
                    _ParteNombre = Mid(_Texto, x, _Mide)
                    If _ParteNombre = _Parte Then
                        _Folio = CServiciosDataSet32.VentaFactor.Rows(i).Item("Folio")
                        _FechaFolio = CServiciosDataSet32.VentaFactor.Rows(i).Item("FechaFolio")
                        _Paciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("Paciente")
                        _NombrePaciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("NombrePaciente")
                        _NumeroFrascos = CServiciosDataSet32.VentaFactor.Rows(i).Item("NumeroFrascos")
                        _PrecioUnitario = CServiciosDataSet32.VentaFactor.Rows(i).Item("PrecioUnitario")
                        _PorcentajeDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("PorcentajeDescuento")
                        _ImporteDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("ImporteDescuento")
                        _Id = CServiciosDataSet32.VentaFactor.Rows(i).Item("Orden")
                        _Medico = CServiciosDataSet32.VentaFactor.Rows(i).Item("Medico")
                        _Cedula = CServiciosDataSet32.VentaFactor.Rows(i).Item("Cedula")
                        _Recibe = CServiciosDataSet32.VentaFactor.Rows(i).Item("Comentarios")

                        DataGridView1.Rows.Add(_Folio, _FechaFolio, _Paciente, _NombrePaciente, _NumeroFrascos, _PrecioUnitario, _PorcentajeDescuento, _ImporteDescuento, _PrecioNeto, _Id, _Medico, _Cedula, _Recibe)

                    End If
                Next

            Next


        End If


        ' Todos
        If RadioButton6.Checked = True Then
            _Rows_VF = CServiciosDataSet32.VentaFactor.Select("Folio <> " & "'" & " " & "'")
        End If


        If RadioButton4.Checked = False And RadioButton7.Checked = False And RadioButton8.Checked = False And RadioButton10.Checked = False And CheckBox2.Checked = False Then
            DataGridView1.Rows.Clear()
            For Me.i = 0 To _Rows_VF.GetUpperBound(0)
                _Folio = _Rows_VF(i).Item("Folio")
                _FechaFolio = _Rows_VF(i).Item("FechaFolio")
                _Paciente = _Rows_VF(i).Item("Paciente")
                _NombrePaciente = _Rows_VF(i).Item("NombrePaciente")
                _NumeroFrascos = _Rows_VF(i).Item("NumeroFrascos")
                _PrecioUnitario = _Rows_VF(i).Item("PrecioUnitario")
                _PorcentajeDescuento = _Rows_VF(i).Item("PorcentajeDescuento")
                _ImporteDescuento = _Rows_VF(i).Item("ImporteDescuento")
                _PrecioNeto = _Rows_VF(i).Item("PrecioNeto")
                _Id = _Rows_VF(i).Item("Orden")
                If DBNull.Value.Equals(_Rows_VF(i).Item("Medico")) Then
                    _Medico = ""
                Else
                    _Medico = _Rows_VF(i).Item("Medico")
                End If

                If DBNull.Value.Equals(_Rows_VF(i).Item("Cedula")) Then
                    _Cedula = ""
                Else
                    _Cedula = _Rows_VF(i).Item("Cedula")
                End If

                _Recibe = _Rows_VF(i).Item("Comentarios")

                DataGridView1.Rows.Add(_Folio, _FechaFolio, _Paciente, _NombrePaciente, _NumeroFrascos, _PrecioUnitario, _PorcentajeDescuento, _ImporteDescuento, _PrecioNeto, _Id, _Medico, _Cedula, _Recibe)
            Next

        End If

        ' Muestra los totales
        _NumeroFrascos = 0
        _ImporteDescuento = 0
        _PrecioNeto = 0
        For Me.i = 0 To DataGridView1.Rows.Count - 1
            _NumeroFrascos = _NumeroFrascos + DataGridView1.Rows(i).Cells(4).Value
            _ImporteDescuento = _ImporteDescuento + DataGridView1.Rows(i).Cells(7).Value
            _PrecioNeto = _PrecioNeto + DataGridView1.Rows(i).Cells(8).Value
        Next
        TextBox16.Text = _NumeroFrascos
        TextBox17.Text = Format(_ImporteDescuento, "$###,##0.00")
        TextBox18.Text = Format(_PrecioNeto, "$###,##0.00")


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow

        Dim X As Integer
        Dim _Find As Boolean
        Dim _BuscaD() As DataRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(0).Value
        Dim _Busca As DataRow = CServiciosDataSet32.VentaFactor.FindByFolio(TextBox1.Text)
        TextBox2.Text = _Busca.Item("Consecutivo")
        TextBox3.Text = _Busca.Item("NombrePaciente")
        TextBox5.Text = _Busca.Item("NumeroFrascos")
        TextBox4.Text = _Busca.Item("PrecioUnitario")
        TextBox7.Text = _Busca.Item("PorcentajeDescuento")
        TextBox8.Text = _Busca.Item("ImporteDescuento")
        TextBox9.Text = _Busca.Item("PrecioNeto")
        _PrecioNeto = CInt(TextBox9.Text)
        TextBox6.Text = Format(_PrecioNeto, "$###,###")
        ToolStripTextBox1.Text = _Busca.Item("Referencia")
        _Diagnostico = ""
        If DBNull.Value.Equals(_Busca.Item("Diagnostico")) Then
            TextBox10.Text = ""
        Else
            TextBox10.Text = _Busca.Item("Diagnostico")
            _Diagnostico = _Busca.Item("Diagnostico")

            _BuscaD = CServiciosDataSet2.Diagnosticos.Select("Diagnostico = " & "'" & _Diagnostico & "'")
            _Find = False

            For X = 0 To _BuscaD.GetUpperBound(0)
                ComboBox4.Text = _BuscaD(X).Item("Descripcion")
                _Find = True
            Next

            If _Find = False Then
                ComboBox4.Text = "SIN DIAGNOSTICO"

            End If
        End If

        ToolStripComboBox1.Text = _Busca.Item("FechaFolio")
        CheckBox1.Checked = _Busca.Item("Contado")
        TextBox12.Text = _Busca.Item("Paciente")
        ComboBox3.Text = _Busca.Item("Paciente")
        TextBox13.Text = _Busca.Item("Medico")
        TextBox14.Text = _Busca.Item("Cedula")
        TextBox15.Text = _Busca.Item("Comentarios")
        If DBNull.Value.Equals(_Busca.Item("Caja")) Then
            TextBox19.Text = "CL13"
        Else
            TextBox19.Text = _Busca.Item("Caja")
        End If

        Dim _Cliente As String = TextBox19.Text
        Dim _Rows_Cl() As DataRow = CServiciosDataSet30.Clientes.Select("Cliente = " & "'" & _Cliente & "'")

        _Find = False
        For Me.i = 0 To _Rows_Cl.GetUpperBound(0)
            ComboBox1.Text = _Rows_Cl(i).Item("NombreCliente")
            _Find = True
        Next

        If Not _Find Then
            _Rows_Cl = CServiciosDataSet30.Clientes.Select("Referencia = " & "'" & _Cliente & "'")
            For Me.i = 0 To _Rows_Cl.GetUpperBound(0)
                ComboBox1.Text = _Rows_Cl(i).Item("NombreCliente")
            Next

        End If


        '_Busca = CServiciosDataSet30.Clientes.FindByCliente(TextBox19.Text)
        'ComboBox1.Text = _Busca.Item("NombreCliente")


        ToolStripButton5.Enabled = True
        ToolStripButton8.Enabled = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Panel1.Enabled = True
        ToolStripButton3.Enabled = True
        ToolStripButton5.Enabled = False
        TextBox1.Enabled = False
        TextBox11.Text = "CAMBIO"
    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub PorFechaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PorFechaToolStripMenuItem.Click
        strRutaExcel = "Z:\Reporte Venta Factor de Transferencia.xlsx"
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de excel

        Dim _FechaInicial As Date
        Dim _FechaFinal As Date
        Dim _Linea As Integer
        Dim _NumRegistros As Integer
        Dim _CuantosFrascos As Integer
        Dim _MontoVenta, _Descuentos, _VentaNeta As Integer

        _FechaInicial = InputBox("Teclee la Fecha Inicial (DD/MM/AAAA")
        _FechaFinal = InputBox("Teclee la Fecha Final (DD/MM/AAAA")
        _MontoVenta = 0
        _Descuentos = 0
        _VentaNeta = 0
        _Linea = 12
        _NumRegistros = 0
        _CuantosFrascos = 0

        For Me.i = 12 To 10000
            If m_Excel.Worksheets("VFT").cells(_Linea, 1).value = "" Then
                Exit For
            End If
            m_Excel.Worksheets("VFT").cells(_Linea, 1).value = ""
            m_Excel.Worksheets("VFT").cells(_Linea, 2).value = ""
            m_Excel.Worksheets("VFT").cells(_Linea, 3).value = ""
            m_Excel.Worksheets("VFT").cells(_Linea, 4).value = ""
            m_Excel.Worksheets("VFT").cells(_Linea, 5).value = ""
            m_Excel.Worksheets("VFT").cells(_Linea, 6).value = ""
            m_Excel.Worksheets("VFT").cells(_Linea, 7).value = ""
            m_Excel.Worksheets("VFT").cells(_Linea, 8).value = ""
            m_Excel.Worksheets("VFT").cells(_Linea, 9).value = ""
        Next



        _Linea = 12
        For Me.i = 0 To CServiciosDataSet32.VentaFactor.Rows.Count - 1
            If CServiciosDataSet32.VentaFactor.Rows(i).Item("FechaFolio") >= _FechaInicial And CServiciosDataSet32.VentaFactor.Rows(i).Item("FechaFolio") <= _FechaFinal Then
                _Folio = CServiciosDataSet32.VentaFactor.Rows(i).Item("Folio")
                _FechaFolio = CServiciosDataSet32.VentaFactor.Rows(i).Item("FechaFolio")
                _NombrePaciente = CServiciosDataSet32.VentaFactor.Rows(i).Item("NombrePaciente")
                _NumeroFrascos = CServiciosDataSet32.VentaFactor.Rows(i).Item("NumeroFrascos")
                _PrecioUnitario = CServiciosDataSet32.VentaFactor.Rows(i).Item("PrecioUnitario")
                _PorcentajeDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("PorcentajeDescuento")
                _ImporteDescuento = CServiciosDataSet32.VentaFactor.Rows(i).Item("ImporteDescuento")
                _PrecioNeto = CServiciosDataSet32.VentaFactor.Rows(i).Item("Precioneto")

                m_Excel.Worksheets("VFT").CELLS(_Linea, 1).VALUE = _Folio
                m_Excel.Worksheets("VFT").CELLS(_Linea, 2).VALUE = _FechaFolio
                m_Excel.Worksheets("VFT").CELLS(_Linea, 3).VALUE = _NombrePaciente
                m_Excel.Worksheets("VFT").CELLS(_Linea, 5).VALUE = _NumeroFrascos
                m_Excel.Worksheets("VFT").CELLS(_Linea, 6).VALUE = _PrecioUnitario
                m_Excel.Worksheets("VFT").CELLS(_Linea, 7).VALUE = _PorcentajeDescuento
                m_Excel.Worksheets("VFT").CELLS(_Linea, 8).VALUE = _ImporteDescuento
                m_Excel.Worksheets("VFT").CELLS(_Linea, 9).VALUE = _PrecioNeto

                _Linea = _Linea + 1
                _NumRegistros = _NumRegistros + 1
                _CuantosFrascos = _CuantosFrascos + _NumeroFrascos
                _MontoVenta = _MontoVenta + (_NumeroFrascos * _PrecioUnitario)
                _Descuentos = _Descuentos + _ImporteDescuento
                _VentaNeta = _VentaNeta + _PrecioNeto
            End If
        Next i

        m_Excel.Worksheets("VFT").cells(7, 3).value = "REPORTE DE FACTOR DE TRANSFERENCIA POR FECHA"
        m_Excel.Worksheets("VFT").CELLS(8, 3).VALUE = Today
        m_Excel.Worksheets("VFT").CELLS(8, 5).VALUE = _CuantosFrascos
        m_Excel.Worksheets("VFT").CELLS(8, 8).VALUE = _NumRegistros
        m_Excel.Worksheets("VFT").CELLS(9, 3).VALUE = _MontoVenta
        m_Excel.Worksheets("VFT").CELLS(9, 5).VALUE = _Descuentos
        m_Excel.Worksheets("VFT").CELLS(9, 8).VALUE = _VentaNeta

        m_Excel.Visible = True
    End Sub

    Private Sub ReportesDeVentasDeFactorToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ReportesDeVentasDeFactorToolStripMenuItem.Click

    End Sub

    Private Sub ContextMenuStrip1_Click(sender As Object, e As System.EventArgs) Handles ContextMenuStrip1.Click

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem1.Click

        Dim _Find As Boolean
        Dim title As String
        Dim response As MsgBoxResult
        Dim _Rows_Re() As DataRow

        msg = "Está a punto de eliminar este Registro. Desea continuar?"
        Style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Eliminar registro de Venta de Factor de Transferencia"
        response = MsgBox(msg, Style, title)
        If response = MsgBoxResult.Yes Then
            ' Elimina el Registro
            _Folio = TextBox1.Text


            ' Verfica que no tenga un recibo asociado
            _Rows_Re = CServiciosDataSet40.Recibos.Select("Folio = " & "'" & _Folio & "'")

            _Find = False
            For Me.i = 0 To _Rows_Re.GetUpperBound(0)
                _Find = True
            Next

            If _Find Then
                msg = "Este registro tiene asociado un Recibo Provisional. No puede eliminarlo"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If

            ' Verifica si fué envío foráneo


            _Busca = CServiciosDataSet32.VentaFactor.FindByFolio(_Folio)
            _Busca.Delete()

            Me.Validate()
            VentaFactorBindingSource2.EndEdit()
            Me.TableAdapterManager3.UpdateAll(CServiciosDataSet32)
            Me.VentaFactorTableAdapter2.Fill(Me.CServiciosDataSet32.VentaFactor)


            msg = "Registro eliminado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            ToolStripButton6_Click(sender, e)
        Else
            Exit Sub
        End If


    End Sub

    Private Sub ToolStripComboBox2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox2.Click
        ToolStripTextBox1.Text = ToolStripComboBox2.Text
    End Sub

    Private Sub ToolStripComboBox2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        ToolStripTextBox1.Text = ToolStripComboBox2.Text

        Button3_Click(sender, e)


    End Sub

    Private Sub TextBox15_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox15.GotFocus
        TextBox15.Text = TextBox3.Text
    End Sub

    Private Sub TextBox15_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox15.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox15_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox15.LostFocus
        TextBox15.BackColor = Color.White
    End Sub

    Private Sub TextBox15_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.GotFocus
        TextBox19.Text = "CL13"
        _Busca = CServiciosDataSet30.Clientes.FindByCliente(TextBox19.Text)
        ComboBox1.Text = _Busca.Item("NombreCliente")
    End Sub

    Private Sub ComboBox1_LostFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.LostFocus
        ComboBox1.BackColor = Color.White
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet30.Clientes.Rows.Count - 1
            If ComboBox1.Text = CServiciosDataSet30.Clientes.Rows(i).Item("NombreCliente") Then
                TextBox19.Text = CServiciosDataSet30.Clientes.Rows(i).Item("Cliente")
            End If
        Next

    End Sub

    Private Sub TextBox12_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox12.GotFocus
        TextBox12.Text = "N/A"
    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox13.GotFocus
        TextBox13.Text = "N/A"
    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox13.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox14.Focus()
        End If
    End Sub

    Private Sub TextBox13_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox13.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox14_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox14.GotFocus
        TextBox14.Text = "N/A"
    End Sub

    Private Sub TextBox14_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox14.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox15.Focus()
        End If
    End Sub

    Private Sub TextBox14_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox14.LostFocus
        TextBox14.BackColor = Color.White
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click
        ' Genera el recibo
        Dim _Find As Boolean

        _Find = False

        ToolStripButton8.Enabled = False
        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                If Checacontrol.Text = "" Then
                    _Find = True
                End If
            End If
        Next

        If _Find Then
            msg = "Debe llenar todas las casillas para generar el Recibo Provisional de Entrega"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If
        Frm48ReciboProvisional.Show()

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        If Asc(e.KeyChar) = 13 Then
            TextBox6.Focus()
        End If

    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub ComboBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        '    If CheckBox2.Checked = False Then
        ' ' ComboBox2.Text = UCase(ComboBox2.Text)
        ' If Asc(e.KeyChar) = 13 Then
        ' TextBox15.Text = UCase(ComboBox2.Text)
        ' TextBox3.Text = UCase(ComboBox2.Text)
        ' TextBox5.Focus()
        ' End If
        ' Else
        ' Button2.Focus()
        ' End If
        '
    End Sub

    Private Sub ComboBox2_LostFocus(sender As Object, e As System.EventArgs) Handles ComboBox2.LostFocus
        If CheckBox2.Checked = False Then
            TextBox15.Text = UCase((ComboBox2.Text))
            TextBox3.Text = UCase((ComboBox2.Text))
        End If
        ComboBox2.BackColor = Color.White
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox3.Text = (UCase(ComboBox2.Text))
        _NombrePaciente = (TextBox3.Text)

        Dim _Rows_Vf() As DataRow
        Dim _Rows_Di() As DataRow
        Dim _Find, _Encuentra As Boolean
        Dim X As Integer

        If CheckBox2.Checked = True Then
            If TextBox3.Text = "" Then
                Exit Sub
            End If

            _NombrePaciente = UCase(_NombrePaciente)
            _Rows_Vf = CServiciosDataSet32.VentaFactor.Select("NombrePaciente = " & "'" & _NombrePaciente & "'")

            _Encuentra = False

            For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
                _Encuentra = True
                TextBox10.Text = _Rows_Vf(i).Item("Diagnostico")
                _Diagnostico = TextBox10.Text

                CheckBox1.Checked = _Rows_Vf(i).Item("Contado")
                TextBox12.Text = _Rows_Vf(i).Item("Paciente")
                ComboBox3.Text = _Rows_Vf(i).Item("Paciente")
                TextBox13.Text = _Rows_Vf(i).Item("Medico")
                TextBox14.Text = _Rows_Vf(i).Item("Cedula")
                TextBox15.Text = _Rows_Vf(i).Item("Comentarios")
                If DBNull.Value.Equals(_Rows_Vf(i).Item("Caja")) Then
                    TextBox19.Text = "CL13"
                Else
                    TextBox19.Text = _Rows_Vf(i).Item("Caja")
                End If

                _Busca = CServiciosDataSet30.Clientes.FindByCliente(TextBox19.Text)
                ComboBox1.Text = _Busca.Item("NombreCliente")
            Next

            If _Encuentra Then
                If DBNull.Value.Equals(_Diagnostico) Then
                    TextBox10.Text = "DI0002"
                    ComboBox4.Text = "SIN DIAGNOSTICO"

                Else
                    _Rows_Di = Me.CServiciosDataSet2.Diagnosticos.Select("Diagnostico = " & "'" & _Diagnostico & "'")
                    _Find = False
                    For X = 0 To _Rows_Di.GetUpperBound(0)
                        ComboBox4.Text = _Rows_Di(X).Item("Descripcion")
                        _Find = True
                        Exit For
                    Next

                    If Not _Find Then
                        TextBox10.Text = "DI0002"
                        ComboBox4.Text = "SIN DIAGNOSTICO"
                    End If
                End If
            End If




                    TextBox5.Focus()

            End If

    End Sub

    Private Sub ComboBox3_LostFocus(sender As Object, e As System.EventArgs) Handles ComboBox3.LostFocus
        ComboBox3.BackColor = Color.White
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        TextBox12.Text = ComboBox3.Text
        TextBox13.Focus()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ' Verifica los frascos vendidos de este lote
        Dim _Lote As String = ToolStripComboBox2.Text
        Dim _Rows_Vf() As DataRow = CServiciosDataSet23.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
        Dim _Cantidad As Integer
        _NumeroFrascos = 0
        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
            _NumeroFrascos = _NumeroFrascos + _Rows_Vf(i).Item("NumeroFrascos")
        Next

        _Busca = CServiciosDataSet26.Produccion.FindByOrdenProduccion(_Lote)
        _Cantidad = _Busca.Item("Cantidad")
        TextBox20.Text = _Cantidad
        TextBox21.Text = _NumeroFrascos
        TextBox22.Text = _Cantidad - _NumeroFrascos


    End Sub

    Private Sub ComboBox4_LostFocus(sender As Object, e As System.EventArgs) Handles ComboBox4.LostFocus
        ComboBox4.BackColor = Color.White
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If TextBox11.Text <> "JR" Then

            Dim _indice As Integer = ComboBox4.SelectedIndex

            _NuevoDiagnostico = False
            _Consecutivo = 0
            _Mide = 0
            _XConsecutivo = ""
            If _indice = 1 Then
                _NuevoDiagnostico = True
                TextBox10.Text = ""
                _Diagnostico = InputBox("Teclee el Nuevo Diagnóstico")
                _Diagnostico = UCase(_Diagnostico)

                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("DGN")
                _Consecutivo = _Busca.Item("Consecutivo")
                _XConsecutivo = "0000" & Trim(CStr(_Consecutivo))
                _Mide = Len(_XConsecutivo)
                _XConsecutivo = "DI" & Mid(_XConsecutivo, (_Mide - 3), 4)
                TextBox10.Text = _XConsecutivo


                ComboBox4.Text = ""
                ComboBox4.Text = _Diagnostico

            Else
                ' Busca el Código
                For Me.i = 0 To CServiciosDataSet2.Diagnosticos.Rows.Count - 1
                    If ComboBox4.Text = CServiciosDataSet2.Diagnosticos.Rows(i).Item("Descripcion") Then
                        TextBox10.Text = CServiciosDataSet2.Diagnosticos.Rows(i).Item("Diagnostico")
                        Exit For
                    End If
                Next


            End If
            ComboBox3.Focus()
        End If

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                If Checacontrol.Focused Then
                    Checacontrol.BackColor = Color.Gold
                End If
            End If
            If TypeOf Checacontrol Is ComboBox Then
                If Checacontrol.Focused Then
                    Checacontrol.BackColor = Color.Gold
                End If

            End If
        Next
    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click

    End Sub
End Class