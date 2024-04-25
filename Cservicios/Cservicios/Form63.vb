Public Class Frm63VentaFactor

    Public i As Integer
    Public msg, _Medico, _Cedula, _Recibe, _Cliente As String
    Public Style As MsgBoxStyle
    Public _Checacontrol As Control
    Public _Renglon As DataGridViewRow
    Public _Folio, _Paciente, _NombrePaciente, _Referencia, _Comentarios, _Caja, _HoraRegistro As String
    Public _NumeroFrascos, _PrecioUnitario, _ImporteDescuento, _PrecioNeto As Integer
    Public _PorcentajeDescuento As Integer
    Public _FechaFolio As String
    Public strRutaExcel As String
    Public _Diagnostico As String
    Public _Consecutivo, _Mide As Integer
    Public _Rows_Vf() As DataRow

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Button6_Click(sender, e)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        For Each Me._Checacontrol In TabPage1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Focused Then
                    _Checacontrol.BackColor = Color.Gold
                Else
                    _Checacontrol.BackColor = Color.White
                End If
            End If
        Next
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView2.Rows.Clear()
        Dim _ParteNombre As String
        Dim _Mide As Integer
        Dim X As Integer
        Dim _Id As Integer

        _ParteNombre = UCase(InputBox("Teclee el Nombre del Paciente"))
        _ParteNombre = Trim(_ParteNombre)

        _Mide = Len(_ParteNombre)

        For Me.i = 0 To CServiciosDataSet55.VentaFactor.Rows.Count - 1
            _NombrePaciente = UCase(CServiciosDataSet55.VentaFactor.Rows(i).Item("NombrePaciente"))
            _FechaFolio = CServiciosDataSet55.VentaFactor.Rows(i).Item("FechaFolio")
            _Id = CServiciosDataSet55.VentaFactor.Rows(i).Item("Orden")

            For X = 1 To Len(_NombrePaciente)
                If _ParteNombre = Mid(_NombrePaciente, X, _Mide) Then
                    DataGridView2.Rows.Add(_NombrePaciente, _FechaFolio, _Id)
                End If

            Next x


        Next


    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Public _XConsecutivo, _Lote As String

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Public _Busca As DataRow

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox5.Focus()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim _DescripcionDiagnostico As String
        _DescripcionDiagnostico = ComboBox5.Text

        For Me.i = 0 To CServiciosDataSet55.Diagnosticos.Rows.Count - 1
            If _DescripcionDiagnostico = CServiciosDataSet55.Diagnosticos.Rows(i).Item("Descripcion") Then
                TextBox12.Text = CServiciosDataSet55.Diagnosticos.Rows(i).Item("Diagnostico")
                Exit For
            End If
        Next
        TextBox4.Focus()

    End Sub

    Public _MideNombre As Integer

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Public _NuevoDiagnostico As Boolean

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Public m_Excel As Microsoft.Office.Interop.Excel.Application

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        Dim _Find As Boolean

        _Find = False
        For Each Me._Checacontrol In TabPage1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Text = "" Then
                    _Find = True
                End If
            End If
        Next

        If _Find Then
            msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        ' Valida que no esté registrado ese folio

        If TextBox15.Text = "ALTA" Then
            _Folio = TextBox1.Text

            Dim _Rows_Vf() As DataRow = CServiciosDataSet55.VentaFactor.Select("Folio = " & "'" & _Folio & "'")

            _Find = False
            For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
                _Find = True
            Next

            If _Find Then
                msg = "Ese Folio ya está registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If
        End If


        ' Verifica que no se exceda del número de frascos
        ' Verficia que no exceda el número de frascos producidos
        Dim _FrascosProducidos As Integer
        Dim _FrascosVendidos As Integer

        Dim _FrascosPorVender As Integer


        _FrascosVendidos = CInt(TextBox17.Text)
        _FrascosProducidos = CInt(TextBox16.Text)
        _FrascosPorVender = CInt(TextBox18.Text)


        _FrascosPorVender = _FrascosProducidos - _FrascosVendidos


        If _FrascosVendidos >= _FrascosProducidos Then
            msg = "Ya se ha vendido la totalidad del Lote  " & _Lote & ". Verifique la información"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        Else
            If _FrascosPorVender < CInt(TextBox4.Text) Then
                msg = "El número de frascos a vender es mayor a la existencia de este Lote, que es " & CStr(_FrascosPorVender) & ". Verfifique la información"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If
        End If



        Dim title As String
        Dim response As MsgBoxResult
        msg = "Está a punto de registrar la Venta de Factor de Transferencia. Está Seguro?"
        Style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Registro de Venta de Factor de Transferencia"
        response = MsgBox(msg, Style)

        If response = MsgBoxResult.Yes Then


            ' Registra un Nuevo Diagnóstico
            If CheckBox1.Checked = True Then
                ' _Busca el siguiente Consecutivo

                _Busca = CServiciosDataSet55.Documentos.FindByDocumento("DGN")
                _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
                Me.Validate()
                DocumentosBindingSource.EndEdit()
                Me.TableAdapterManager.UpdateAll(CServiciosDataSet55)
                Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet55.Documentos)

                _Busca = CServiciosDataSet55.Documentos.FindByDocumento("DGN")
                _Consecutivo = _Busca.Item("Consecutivo")


                Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
                Dim _Mide As Integer = Len(_XConsecutivo)
                Dim _x As Integer = _Mide - 3
                _XConsecutivo = "DG" + Mid(_XConsecutivo, _x, 4)

                ' Registra el nuevo Diagn+ostico
                _Busca = CServiciosDataSet55.Diagnosticos.NewRow
                _Busca.Item("Diagnostico") = _XConsecutivo
                _Busca.Item("Descripcion") = ComboBox5.Text
                _Busca.Item("Centro") = "FTR"
                _Busca.Item("Referencia") = "VENTA DE FACTOR DE TRANSFERENCIA"
                _Busca.Item("Comentario") = "DIAGNOSTICO AUTOMATICO"
                _Busca.Item("FechaRegistro") = CStr(Today.Date)
                _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)

                Me.CServiciosDataSet55.Diagnosticos.Rows.Add(_Busca)
                Me.Validate()
                DiagnosticosBindingSource.EndEdit()
                Me.TableAdapterManager.UpdateAll(CServiciosDataSet55)
                Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet55.Diagnosticos)

            End If


            ' Registra la Venta
            If TextBox15.Text = "ALTA" Then
                _Busca = CServiciosDataSet55.VentaFactor.NewRow
                _Busca.Item("Folio") = TextBox1.Text
            End If
            If TextBox15.Text = "CAMBIO" Then
                _Busca = CServiciosDataSet55.VentaFactor.FindByFolio(TextBox1.Text)
            End If

            _Busca.Item("FechaFolio") = TextBox19.Text
            _Busca.Item("Paciente") = ComboBox3.Text
            _Busca.Item("NombrePaciente") = TextBox3.Text
            _Busca.Item("NumeroFrascos") = CInt(TextBox4.Text)
            _Busca.Item("PrecioUnitario") = CInt(TextBox14.Text)
            _Busca.Item("PorcentajeDescuento") = CInt(TextBox7.Text)
            _Busca.Item("ImporteDescuento") = CInt(TextBox6.Text)
            _Busca.Item("PrecioNeto") = CInt(TextBox5.Text)
            _Busca.Item("Referencia") = ComboBox1.Text
            _Busca.Item("FechaRegistro") = Today
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
            _Busca.Item("Diagnostico") = TextBox12.Text

            If CheckBox1.Checked = True Then
                _Busca.Item("Diagnostico") = _XConsecutivo
            End If


            _Busca.Item("Medico") = UCase(TextBox8.Text)
            _Busca.Item("cedula") = UCase(TextBox9.Text)
            _Busca.Item("Comentarios") = UCase(TextBox10.Text)
            _Busca.Item("Caja") = TextBox11.Text
            _Busca.Item("Contado") = True


            If TextBox15.Text = "ALTA" Then
                CServiciosDataSet55.VentaFactor.Rows.Add(_Busca)
            End If


            Me.Validate()
            VentaFactorBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet55)
            Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet55.VentaFactor)


            msg = "Registro guardado correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)

            Timer1.Stop()

            Button6_Click(sender, e)
            Button1.Focus()


        Else

        End If



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        ' Determina el número de Frascos Disponibles
        _Lote = ComboBox1.Text
        '   Dim _Rows_Vf() As DataRow
        Dim _Cantidad As Integer

        _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")

        _Cantidad = 0
        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
            _Cantidad = _Cantidad + _Rows_Vf(i).Item("NumeroFrascos")
        Next

        Dim _FrascosProducidos As Integer
        _FrascosProducidos = 0

        _Busca = CServiciosDataSet55.Produccion.FindByOrdenProduccion(_Lote)
        TextBox16.Text = _Busca.Item("Cantidad")

        TextBox17.Text = _Cantidad
        TextBox18.Text = _Busca.Item("Cantidad") - _Cantidad

        TextBox1.Focus()

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click

        Button6_Click(sender, e)
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True
        Button5.Enabled = True

        TextBox2.Text = "0"
        TextBox15.Text = "ALTA"
        Timer1.Start()
        Timer1.Interval = 100

        Dim _Cantidad As Integer
        '  Dim _Rows_Vf() As DataRow
        Dim _Edad As String
        Dim _Cuantos As Integer


        For Each Me._Checacontrol In TabPage1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = True
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Enabled = True
            End If
        Next


        ComboBox2.Text = Format(Today, "Long Date")
        TextBox19.Text = Format(Today, "Short Date")

        ' Llena el combo con las ordenes de Producción
        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet55.Produccion.Rows.Count - 1
            _Lote = CServiciosDataSet55.Produccion.Rows(i).Item("Lote")
            _Cantidad = CServiciosDataSet55.Produccion.Rows(i).Item("Cantidad")

            _NumeroFrascos = 0
            _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
            For X = 0 To _Rows_Vf.GetUpperBound(0)
                _NumeroFrascos = _NumeroFrascos + _Rows_Vf(X).Item("NumeroFrascos")
            Next

            If _NumeroFrascos < _Cantidad Then
                ComboBox1.Items.Add(_Lote)
            End If

        Next


        _Cuantos = ComboBox1.Items.Count

        '  ComboBox1.Text = ComboBox1.Items(_Cuantos - 1)

        ' Obtiene el número de Frascos disponibles
        '    _Lote = ComboBox1.Text
        '
        '        _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
        '
        '        _Cantidad = 0
        '        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
        '        _Cantidad = _Cantidad + _Rows_Vf(i).Item("NumeroFrascos")
        '        Next
        '
        '        Dim _FrascosProducidos As Integer
        '        _FrascosProducidos = 0
        '
        '        _Busca = CServiciosDataSet55.Produccion.FindByOrdenProduccion(_Lote)
        '        TextBox16.Text = _Busca.Item("Cantidad")
        '
        '        TextBox17.Text = _Cantidad
        '        TextBox18.Text = _Busca.Item("Cantidad") - _Cantidad



        ' Llena el combo de la Edad
        _Edad = ""
        ComboBox3.Items.Clear()
        For Me.i = 0 To 100
            _Edad = CStr(i)
            ComboBox3.Items.Add(_Edad)
        Next

        ' Llena el Combo de los Diagnósticos
        ComboBox5.Items.Clear()
        For Me.i = 0 To CServiciosDataSet55.Diagnosticos.Rows.Count - 1
            ComboBox5.Items.Add(CServiciosDataSet55.Diagnosticos.Rows(i).Item("Descripcion"))
        Next
        ComboBox5.Sorted = True

        ' Llena el Combo de los Clientes
        ComboBox4.Items.Clear()
        For Me.i = 0 To CServiciosDataSet55.Clientes.Rows.Count - 1
            ComboBox4.Items.Add(CServiciosDataSet55.Clientes.Rows(i).Item("NombreCliente"))
        Next

        ' Obtiene el precio unitario
        _Busca = CServiciosDataSet55.Productos.FindByProducto("FTR")
        TextBox13.Text = Format(_Busca.Item("PrecioUnitario"), "$#,###.00")
        TextBox14.Text = _Busca.Item("PrecioUnitario")
        TextBox1.Focus()

        ' Muestra las ventas de este dia
        Button7_Click(sender, e)



    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_Click(sender As Object, e As EventArgs) Handles ComboBox2.Click
        MonthCalendar1.Visible = True
        ComboBox2.DroppedDown = False
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        ComboBox2.Text = Format(MonthCalendar1.SelectionRange.Start, "Long Date")
        TextBox19.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False

        ' Muestra las ventas de esta fecha
        Button7_Click(sender, e)



    End Sub

    Private Sub Frm63VentaFactor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button6_Click(sender, e)
        TabPage1.Parent = TabControl1
        TabPage2.Parent = Nothing
        TabPage3.Parent = Nothing
        TabPage4.Parent = Nothing
        Button1.BackColor = Color.LightBlue
        Button2.BackColor = Color.LightGray
        Button3.BackColor = Color.LightGray
        Button4.BackColor = Color.LightGray




    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'TODO: This line of code loads data into the 'CServiciosDataSet55.VentaFactor' table. You can move, or remove it, as needed.
        Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet55.VentaFactor)
        'TODO: This line of code loads data into the 'CServiciosDataSet55.Recibos' table. You can move, or remove it, as needed.
        Me.RecibosTableAdapter.Fill(Me.CServiciosDataSet55.Recibos)
        'TODO: This line of code loads data into the 'CServiciosDataSet55.Productos' table. You can move, or remove it, as needed.
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet55.Productos)
        'TODO: This line of code loads data into the 'CServiciosDataSet55.Produccion' table. You can move, or remove it, as needed.
        Me.ProduccionTableAdapter.Fill(Me.CServiciosDataSet55.Produccion)
        'TODO: This line of code loads data into the 'CServiciosDataSet55.Documentos' table. You can move, or remove it, as needed.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet55.Documentos)
        'TODO: This line of code loads data into the 'CServiciosDataSet55.Diagnosticos' table. You can move, or remove it, as needed.
        Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet55.Diagnosticos)
        'TODO: This line of code loads data into the 'CServiciosDataSet55.Clientes' table. You can move, or remove it, as needed.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet55.Clientes)


        ToolStripButton3.Enabled = True
        ToolStripButton4.Enabled = False
        Button5.Enabled = False
        TextBox15.Text = ""

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
        DataGridView2.Rows.Clear()




    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        Dim _id As Double
        Dim _DescripcionDiagnostico As String
        Dim X As Integer

        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _id = _Renglon.Cells(2).Value

        _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("Orden = " & _id)
        Dim _Rows_D() As DataRow
        Dim _Find As Boolean

        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
            TextBox3.Text = _Rows_Vf(i).Item("NombrePaciente")
            _Diagnostico = _Rows_Vf(i).Item("Diagnostico")

            _Rows_D = CServiciosDataSet55.Diagnosticos.Select("Diagnostico =" & "'" & _Diagnostico & "'")

            _Find = False


            '  MsgBox("Valor de I" & CStr(i))

            If DBNull.Value.Equals(_Rows_Vf(i).Item("Medico")) Then
                TextBox8.Text = "N/A"
                TextBox9.Text = "N/A"
            Else
                TextBox8.Text = _Rows_Vf(i).Item("Medico")
                TextBox9.Text = _Rows_Vf(i).Item("Cedula")
            End If
            TextBox8.Text = _Rows_Vf(i).Item("Medico")
            TextBox9.Text = _Rows_Vf(i).Item("Cedula")
            _Cliente = _Rows_Vf(i).Item("Caja")
            _Busca = CServiciosDataSet55.Clientes.FindByCliente(_Cliente)
            ComboBox4.Text = _Busca.Item("NombreCliente")
            TextBox11.Text = _Cliente
            TextBox11.ReadOnly = True
            ComboBox3.Text = _Rows_Vf(i).Item("Paciente")
            TextBox10.Text = _Rows_Vf(i).Item("Comentarios")


            For X = 0 To _Rows_D.GetUpperBound(0)
                If _Diagnostico = _Rows_D(X).Item("Diagnostico") Then
                    ComboBox5.Text = _Rows_D(X).Item("Descripcion")
                    TextBox12.Text = _Diagnostico
                    TextBox12.ReadOnly = True
                    _Find = True
                    Exit For
                End If
            Next X

            If _Find = False Then
                ComboBox5.Text = "SIN DIAGNOSTICO"
                TextBox12.Text = "DI0002"
            End If
        Next

        TextBox1.Focus()


    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ' Llena el grid con la venta de el dia 

        Dim x As Integer
        '   Dim _Rows_Vf() As DataRow
        Dim _Cantidad As Integer
        Dim _Edad As String
        Dim _Fecha As Date
        Dim _Factura, _NombreCliente As String
        Dim _Busca As DataRow
        Dim _NombreDiagnostico As String
        Dim _Rows_D() As DataRow
        Dim _Find As Boolean

        _FechaFolio = ComboBox2.Text
        _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("FechaFolio = " & "'" & CDate(_FechaFolio) & "'")

        DataGridView1.Rows.Clear()

        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
            _Folio = _Rows_Vf(i).Item("Folio")
            _Lote = _Rows_Vf(i).Item("Referencia")
            _Cantidad = _Rows_Vf(i).Item("NumeroFrascos")
            _PrecioNeto = _Rows_Vf(i).Item("PrecioNeto")
            _ImporteDescuento = _Rows_Vf(i).Item("ImporteDescuento")
            _PorcentajeDescuento = _Rows_Vf(i).Item("PorcentajeDescuento")
            _NombrePaciente = _Rows_Vf(i).Item("NombrePaciente")
            _Edad = _Rows_Vf(i).Item("Paciente")
            _Medico = _Rows_Vf(i).Item("Medico")
            _Cedula = _Rows_Vf(i).Item("Cedula")
            _Fecha = _Rows_Vf(i).Item("FechaFolio")
            _Comentarios = _Rows_Vf(i).Item("Comentarios")
            If DBNull.Value.Equals(_Rows_Vf(i).Item("Factura")) Then
                _Factura = ""
            Else
                _Factura = _Rows_Vf(i).Item("Factura")
            End If
            _Cliente = _Rows_Vf(i).Item("Caja")

            _NombreCliente = "** CLIENTE NO REGISTRADO **"
            For x = 0 To CServiciosDataSet55.Clientes.Rows.Count - 1
                If CServiciosDataSet55.Clientes.Rows(x).Item("Cliente") = _Cliente Then
                    _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                End If
                If CServiciosDataSet55.Clientes.Rows(x).Item("Referencia") = _Cliente Then
                    _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                End If

            Next

            '_Busca = CServiciosDataSet55.Clientes.FindByCliente(_Cliente)

            _Diagnostico = _Rows_Vf(i).Item("Diagnostico")

            _Rows_D = CServiciosDataSet55.Diagnosticos.Select("Diagnostico = " & "'" & _Diagnostico & "'")

            _Find = False
            For X = 0 To _Rows_D.GetUpperBound(0)
                _NombreDiagnostico = _Rows_D(X).Item("Descripcion")
                _Find = True
            Next

            If Not _Find Then
                _NombreDiagnostico = "SIN DIAGNOSTICO"
                _Diagnostico = "DI0002"
            End If


            DataGridView1.Rows.Add(_Folio, _Lote, _Fecha, _Cantidad, _PrecioNeto, _ImporteDescuento, _PorcentajeDescuento, _NombrePaciente, _Edad, _Medico, _Cedula, _Comentarios, _Factura, _Cliente, _NombreCliente, _Diagnostico, _NombreDiagnostico)
        Next
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox3.Focus()
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click


        Dim _Factura = DataGridView1.CurrentRow.Cells(12).Value
        If _Factura <> "" Then
            msg = "Esta venta tiene asignada la Factura " & _Factura & " .No puede ser Cancelada"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If


        Dim title As String
        Dim response As MsgBoxResult

        msg = "Está a punto de Cancelar esta Venta. Está Seguro?"
        Style = MsgBoxStyle.DefaultButton2 Or
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cancelación Venta de Factor de Transferencia"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            _Folio = TextBox1.Text
            _Busca = CServiciosDataSet55.VentaFactor.FindByFolio(_Folio)
            _Busca.Delete()
            Me.Validate()
            VentaFactorBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet55)
            Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet55.VentaFactor)

            msg = "Registro eliminado correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)

            Button6_Click(sender, e)
            Exit Sub


        Else
            ' Perform some other action.
        End If


    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        For Each Me._Checacontrol In TabPage1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = True
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Enabled = True
            End If
        Next
        TextBox1.Enabled = False
        TextBox3.Focus()
        Timer1.Start()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabPage1.Parent = TabControl1
        TabPage2.Parent = Nothing
        TabPage3.Parent = Nothing
        TabPage4.Parent = Nothing
        DataGridView1.Visible = True
        DataGridView2.Visible = True
        Button1.BackColor = Color.LightBlue
        Button2.BackColor = Color.LightGray
        Button4.BackColor = Color.LightGray
        Button2.BackColor = Color.LightGray


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.BackColor = Color.LightBlue
        Button1.BackColor = Color.LightGray
        Button3.BackColor = Color.LightGray
        Button4.BackColor = Color.LightGray
        TabPage2.Parent = TabControl1
        TabPage1.Parent = Nothing
        TabPage3.Parent = Nothing
        TabPage4.Parent = Nothing

        DataGridView1.Visible = False
        DataGridView2.Visible = False
        MonthCalendar2.Visible = False
        MonthCalendar3.Visible = False
        ComboBox6.Text = Format(Today, "Long Date")
        ComboBox7.Text = Format(Today, "Long Date")
        TextBox20.Text = Format(Today, "Short Date")
        TextBox21.Text = Format(Today, "Short Date")

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

    End Sub

    Private Sub ComboBox5_GotFocus(sender As Object, e As EventArgs) Handles ComboBox5.GotFocus
        TextBox12.Text = ""
    End Sub

    Private Sub MonthCalendar2_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar2.DateChanged

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox8.Focus()
        End If
    End Sub

    Private Sub MonthCalendar3_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar3.DateChanged

    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        Dim _ImporteDescuento As Integer
        Dim _ImporteVenta As Integer = CInt(TextBox5.Text)
        Dim _Neto As Integer = CInt(TextBox4.Text) * CInt(TextBox14.Text)
        Dim _PorcentajeDescuento As Double


        ' MsgBox("Importe Venta " & CStr(_ImporteVenta))
        ' MsgBox("Neto " & CStr(_Neto))


        _ImporteDescuento = _Neto - _ImporteVenta
        TextBox6.Text = _ImporteDescuento
        _PorcentajeDescuento = (((_ImporteDescuento * 100) / _Neto) / 100)

        If _ImporteDescuento = 0 Then
            _PorcentajeDescuento = 0
        End If
        TextBox7.Text = _PorcentajeDescuento

        TextBox8.Focus()





    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As EventArgs) Handles ComboBox1.GotFocus

    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As EventArgs) Handles TextBox8.GotFocus
        If TextBox8.Text = "" Then
            TextBox8.Text = "N/A"
        End If

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim _FechaInicial As Date = CDate(TextBox20.Text)
        Dim _FechaFinal As Date = CDate(TextBox21.Text)
        Dim _NumeroDosis As Integer
        Dim _VentaTotal As Double
        Dim _VentaPublico, _VentaInstituciones As Double

        _NumeroDosis = 0
        _VentaTotal = 0
        _VentaPublico = 0
        _VentaInstituciones = 0
        _ImporteDescuento = 0


        _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("FechaFolio >= " & "'" & _FechaInicial & "'" & " and FechaFolio <= " & "'" & _FechaFinal & "'")

        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
            _NumeroDosis = _NumeroDosis + _Rows_Vf(i).Item("NumeroFrascos")
            _VentaTotal = _VentaTotal + _Rows_Vf(i).Item("PrecioNeto")
            _ImporteDescuento = _ImporteDescuento + _Rows_Vf(i).Item("ImporteDescuento")

            If _Rows_Vf(i).Item("caja") = "CL13" Then
                _VentaPublico = _VentaPublico + _Rows_Vf(i).Item("PrecioNeto")
            Else
                _VentaInstituciones = _VentaInstituciones + _Rows_Vf(i).Item("PrecioNeto")
            End If
        Next


        TextBox22.Text = Format(_NumeroDosis, "###,###")
        TextBox23.Text = Format(_VentaTotal, "$ ###,###,###.00")
        TextBox24.Text = Format(_ImporteDescuento, "$ ###,###,###.00")
        _PorcentajeDescuento = ((_ImporteDescuento * 100) / _VentaTotal)
        TextBox25.Text = " % " & CStr(_PorcentajeDescuento)
        TextBox26.Text = Format(_VentaPublico, "$ ###,###,###.00")
        TextBox27.Text = Format(_VentaInstituciones, "$ ###,###,###.00")

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ' Busquedas
        'Dim _Rows_Vf() As DataRow
        Dim _find As Boolean
        Dim _ParteNombre As String
        Dim _Cantidad As Integer
        Dim _Edad As String
        Dim _Fecha As Date
        Dim _Factura, _NombreCliente, _Nombrediagnostico As String
        Dim _Rows_D() As DataRow
        Dim _NumeroDosis As Integer
        Dim _VentaTotal As Double
        Dim _VentaPublico, _VentaInstituciones As Double
        Dim _Descuentos As Double
        Dim _FechaInicial As Date
        Dim _FechaFinal As Date
        Dim _Espacios As String = ""

        TextBox29.Text = ""
        TextBox30.Text = ""
        TextBox31.Text = ""
        TextBox32.Text = ""
        TextBox33.Text = ""
        TextBox34.Text = ""

        _NumeroDosis = 0
        _VentaTotal = 0
        _VentaPublico = 0
        _VentaInstituciones = 0
        _ImporteDescuento = 0
        _Descuentos = 0

        DataGridView1.Rows.Clear()


        _find = False
        _ParteNombre = ""


        ' Por Fecha
        If RadioButton1.Checked = True Then
            _FechaFolio = InputBox("Teclee la fecha a Consultar")
            TextBox28.Text = _FechaFolio
            If Not IsDate(TextBox28.Text) Then
                msg = "Debe proporcionar una fecha Válida. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If

            _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("FechaFolio = " & "'" & _FechaFolio & "'")
            _find = True

        End If

        ' Por Lote
        If RadioButton2.Checked = True Then
            _Lote = UCase(InputBox("Teclee el Lote a consultar"))
            TextBox28.Text = _Lote
            _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("Referencia = " & "'" & _Lote & "'")
            _find = True
        End If

        'Por Folio
        If RadioButton3.Checked = True Then
            _Folio = UCase(InputBox("Teclee el Folio a consultar"))
            TextBox28.Text = _Folio
            _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("Folio = " & "'" & _Folio & "'")
            _find = True
        End If

        ' Nombre del Paciente
        If RadioButton4.Checked = True Then
            _NombrePaciente = UCase(InputBox("Teclee el nombre del Paciente a Consultar"))
            _ParteNombre = _NombrePaciente
        End If



        ' Todos
        If RadioButton6.Checked = True Then
            _find = True
            If TextBox36.Text = "" Then
                _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("NombrePaciente <> " & "'" & _Espacios & "'")
            Else
                _FechaInicial = CDate(TextBox36.Text)
                _FechaFinal = CDate(TextBox37.Text)
                _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("FechaFolio >= " & "'" & _FechaInicial & "'" & " and FechaFolio <= " & "'" & _FechaFinal & "'")
            End If

        End If


        ' Por Diagnóstico
        If RadioButton7.Checked = True Then
            _Diagnostico = UCase(InputBox("Teclee el Diagnóstico a Consultar"))
            _ParteNombre = _Diagnostico
        End If

        ' Por Médico
        If RadioButton8.Checked = True Then
            _Medico = UCase(InputBox("Teclee el nombre del Médico a Consultar"))
            _ParteNombre = _Medico
        End If


        ' Por Cédula
        If RadioButton9.Checked = True Then
            _Cedula = UCase(InputBox("Teclee la Cédula a Consultar"))
            _ParteNombre = _Cedula
        End If


        ' Por Comprador
        If RadioButton11.Checked = True Then
            _Comentarios = UCase(InputBox("Teclee el nombre del Comprador a Consultar"))
            _ParteNombre = _Comentarios
        End If



        If _find = True Then

            For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
                _Folio = _Rows_Vf(i).Item("Folio")
                _Lote = _Rows_Vf(i).Item("Referencia")
                _Cantidad = _Rows_Vf(i).Item("NumeroFrascos")
                _PrecioNeto = _Rows_Vf(i).Item("PrecioNeto")
                _ImporteDescuento = _Rows_Vf(i).Item("ImporteDescuento")
                _PorcentajeDescuento = _Rows_Vf(i).Item("PorcentajeDescuento")
                _NombrePaciente = _Rows_Vf(i).Item("NombrePaciente")
                _Edad = _Rows_Vf(i).Item("Paciente")
                _Medico = _Rows_Vf(i).Item("Medico")
                _Cedula = _Rows_Vf(i).Item("Cedula")
                _Fecha = _Rows_Vf(i).Item("FechaFolio")
                _Comentarios = _Rows_Vf(i).Item("Comentarios")
                If DBNull.Value.Equals(_Rows_Vf(i).Item("Factura")) Then
                    _Factura = ""
                Else
                    _Factura = _Rows_Vf(i).Item("Factura")
                End If
                _Cliente = _Rows_Vf(i).Item("Caja")

                _NombreCliente = "** CLIENTE NO REGISTRADO **"
                For x = 0 To CServiciosDataSet55.Clientes.Rows.Count - 1
                    If _Cliente = CServiciosDataSet55.Clientes.Rows(x).Item("Cliente") Then
                        _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                        Exit For
                    End If
                    If _Cliente = CServiciosDataSet55.Clientes.Rows(x).Item("Referencia") Then
                        _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                        Exit For
                    End If

                Next


                _Diagnostico = _Rows_Vf(i).Item("Diagnostico")

                _Rows_D = CServiciosDataSet55.Diagnosticos.Select("Diagnostico = " & "'" & _Diagnostico & "'")

                _find = False
                For X = 0 To _Rows_D.GetUpperBound(0)
                    _Nombrediagnostico = _Rows_D(X).Item("Descripcion")
                    _find = True
                Next

                If Not _find Then
                    _NombreDiagnostico = "SIN DIAGNOSTICO"
                    _Diagnostico = "DI0002"
                End If


                DataGridView1.Rows.Add(_Folio, _Lote, _Fecha, _Cantidad, _PrecioNeto, _ImporteDescuento, _PorcentajeDescuento, _NombrePaciente, _Edad, _Medico, _Cedula, _Comentarios, _Factura, _Cliente, _NombreCliente, _Diagnostico, _Nombrediagnostico)

                _NumeroDosis = _NumeroDosis + _Rows_Vf(i).Item("NumeroFrascos")
                _VentaTotal = _VentaTotal + _Rows_Vf(i).Item("PrecioNeto")
                _Descuentos = _Descuentos + _Rows_Vf(i).Item("ImporteDescuento")

                If _Rows_Vf(i).Item("caja") = "CL13" Then
                    _VentaPublico = _VentaPublico + _Rows_Vf(i).Item("PrecioNeto")
                Else
                    _VentaInstituciones = _VentaInstituciones + _Rows_Vf(i).Item("PrecioNeto")
                End If

            Next







        End If



        If Not _find Then

            Dim _DatoRegistrado As String
            Dim _Incluye As Boolean



            If TextBox36.Text = "" Then
                _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("NombrePaciente <> " & "'" & _Espacios & "'")
            Else
                _FechaInicial = CDate(TextBox36.Text)
                _FechaFinal = CDate(TextBox37.Text)
                _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("FechaFolio >= " & "'" & _FechaInicial & "'" & " and FechaFolio <= " & "'" & _FechaFinal & "'")
            End If


            _Mide = Len(_ParteNombre)
            _DatoRegistrado = ""
            For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
                _Incluye = False

                If RadioButton4.Checked = True Then
                    _DatoRegistrado = _Rows_Vf(i).Item("NombrePaciente")
                End If
                If RadioButton8.Checked = True Then
                    _DatoRegistrado = _Rows_Vf(i).Item("Medico")
                End If
                If RadioButton9.Checked = True Then
                    _DatoRegistrado = _Rows_Vf(i).Item("Cedula")
                End If
                If RadioButton11.Checked = True Then
                    _DatoRegistrado = _Rows_Vf(i).Item("Comentarios")
                End If



                If RadioButton7.Checked = True Then
                    _Diagnostico = _Rows_Vf(i).Item("Diagnostico")

                    _DatoRegistrado = _Diagnostico
                    For x = 0 To CServiciosDataSet55.Diagnosticos.Rows.Count - 1
                        If _Diagnostico = CServiciosDataSet55.Diagnosticos.Rows(x).Item("Diagnostico") Then
                            _DatoRegistrado = CServiciosDataSet55.Diagnosticos.Rows(x).Item("Descripcion")
                            Exit For
                        End If
                    Next
                End If


                For x = 1 To Len(_DatoRegistrado)
                    If _ParteNombre = Mid(_DatoRegistrado, x, _Mide) Then
                        _Incluye = True
                    End If
                Next


                If _Incluye Then

                    _Folio = _Rows_Vf(i).Item("Folio")
                    _Lote = _Rows_Vf(i).Item("Referencia")
                    _Cantidad = _Rows_Vf(i).Item("NumeroFrascos")
                    _PrecioNeto = _Rows_Vf(i).Item("PrecioNeto")
                    _ImporteDescuento = _Rows_Vf(i).Item("ImporteDescuento")
                    _PorcentajeDescuento = _Rows_Vf(i).Item("PorcentajeDescuento")
                    _NombrePaciente = _Rows_Vf(i).Item("NombrePaciente")
                    _Edad = _Rows_Vf(i).Item("Paciente")
                    _Medico = _Rows_Vf(i).Item("Medico")
                    _Cedula = _Rows_Vf(i).Item("Cedula")
                    _Fecha = _Rows_Vf(i).Item("FechaFolio")
                    _Comentarios = _Rows_Vf(i).Item("Comentarios")
                    If DBNull.Value.Equals(_Rows_Vf(i).Item("Factura")) Then
                        _Factura = ""
                    Else
                        _Factura = _Rows_Vf(i).Item("Factura")
                    End If
                    _Cliente = Trim(_Rows_Vf(i).Item("Caja"))


                    _NombreCliente = "** CLIENTE NO REGISTRADO **"
                    For x = 0 To CServiciosDataSet55.Clientes.Rows.Count - 1
                        If _Cliente = Trim(CServiciosDataSet55.Clientes.Rows(x).Item("Cliente")) Then
                            _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                            Exit For
                        End If
                        If _Cliente = Trim(CServiciosDataSet55.Clientes.Rows(x).Item("Referencia")) Then
                            _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                            Exit For
                        End If

                    Next


                    _Diagnostico = _Rows_Vf(i).Item("Diagnostico")

                    _Rows_D = CServiciosDataSet55.Diagnosticos.Select("Diagnostico = " & "'" & _Diagnostico & "'")

                    _find = False
                    For X = 0 To _Rows_D.GetUpperBound(0)
                        _Nombrediagnostico = _Rows_D(X).Item("Descripcion")
                        _find = True
                    Next

                    If Not _find Then
                        _Nombrediagnostico = _Diagnostico
                        _Diagnostico = ""
                    End If


                    DataGridView1.Rows.Add(_Folio, _Lote, _Fecha, _Cantidad, _PrecioNeto, _ImporteDescuento, _PorcentajeDescuento, _NombrePaciente, _Edad, _Medico, _Cedula, _Comentarios, _Factura, _Cliente, _NombreCliente, _Diagnostico, _Nombrediagnostico)

                    _NumeroDosis = _NumeroDosis + _Rows_Vf(i).Item("NumeroFrascos")
                    _VentaTotal = _VentaTotal + _Rows_Vf(i).Item("PrecioNeto")
                    _Descuentos = _Descuentos + _Rows_Vf(i).Item("ImporteDescuento")

                    If _Rows_Vf(i).Item("caja") = "CL13" Then
                        _VentaPublico = _VentaPublico + _Rows_Vf(i).Item("PrecioNeto")
                    Else
                        _VentaInstituciones = _VentaInstituciones + _Rows_Vf(i).Item("PrecioNeto")
                    End If


                End If

            Next
        End If

        TextBox34.Text = Format(_NumeroDosis, "###,###")
        TextBox33.Text = Format(_VentaTotal, "$ ###,###,###.00")
        TextBox32.Text = Format(_Descuentos, "$ ###,###,###.00")
        If _VentaTotal > 0 Then
            _PorcentajeDescuento = ((_Descuentos * 100) / _VentaTotal)
        Else
            If _Descuentos > 0 Then
                _PorcentajeDescuento = 100
            End If
        End If
        TextBox31.Text = " % " & CStr(_PorcentajeDescuento)
        TextBox30.Text = Format(_VentaPublico, "$ ###,###,###.00")
        TextBox29.Text = Format(_VentaInstituciones, "$ ###,###,###.00")


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TabPage1.Parent = Nothing
        TabPage2.Parent = Nothing
        TabPage3.Parent = Nothing
        TabPage4.Parent = TabControl1
        DataGridView1.Visible = True
        DataGridView2.Visible = False
        Button1.BackColor = Color.LightGray
        Button2.BackColor = Color.LightGray
        Button4.BackColor = Color.LightBlue
        Button2.BackColor = Color.LightGray
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        TextBox35.Text = "D"
        Button17_Click(sender, e)


    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        ' Llena el grid con la venta de el dia 
        DataGridView1.Visible = True
        DataGridView1.Rows.Clear()

        Dim x As Integer
        Dim _Rows_Vf() As DataRow
        Dim _Cantidad As Integer
        Dim _Edad As String
        Dim _Fecha As Date
        Dim _Factura, _NombreCliente As String
        Dim _Busca As DataRow
        Dim _NombreDiagnostico As String
        Dim _Rows_D() As DataRow
        Dim _Find As Boolean
        Dim _FechaInicial As Date = CDate(TextBox20.Text)
        Dim _FechaFinal As Date = CDate(TextBox21.Text)


        If TextBox35.Text = "D" Then
            _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("FechaFolio >= " & "'" & _FechaInicial & "'" & " and FechaFolio <= " & "'" & _FechaFinal & "'" & " and ImporteDescuento > 0")
        End If

        If TextBox35.Text = "P" Then
            _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("FechaFolio >= " & "'" & _FechaInicial & "'" & " and FechaFolio <= " & "'" & _FechaFinal & "'" & " and ImporteDescuento = 0" & " and Caja = " & "'" & "CL13" & "'")
        End If

        If TextBox35.Text = "I" Then
            _Rows_Vf = CServiciosDataSet55.VentaFactor.Select("FechaFolio >= " & "'" & _FechaInicial & "'" & " and FechaFolio <= " & "'" & _FechaFinal & "'" & " and ImporteDescuento = 0" & " and Caja <> " & "'" & "CL13" & "'")
        End If

        DataGridView1.Rows.Clear()

        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)
            _Folio = _Rows_Vf(i).Item("Folio")
            _Lote = _Rows_Vf(i).Item("Referencia")
            _Cantidad = _Rows_Vf(i).Item("NumeroFrascos")
            _PrecioNeto = _Rows_Vf(i).Item("PrecioNeto")
            _ImporteDescuento = _Rows_Vf(i).Item("ImporteDescuento")
            _PorcentajeDescuento = _Rows_Vf(i).Item("PorcentajeDescuento")
            _NombrePaciente = _Rows_Vf(i).Item("NombrePaciente")
            _Edad = _Rows_Vf(i).Item("Paciente")
            _Medico = _Rows_Vf(i).Item("Medico")
            _Cedula = _Rows_Vf(i).Item("Cedula")
            _Fecha = _Rows_Vf(i).Item("FechaFolio")
            _Comentarios = _Rows_Vf(i).Item("Comentarios")
            If DBNull.Value.Equals(_Rows_Vf(i).Item("Factura")) Then
                _Factura = ""
            Else
                _Factura = _Rows_Vf(i).Item("Factura")
            End If
            _Cliente = _Rows_Vf(i).Item("Caja")

            _NombreCliente = "** CLIENTE NO REGISTRADO **"
            For x = 0 To CServiciosDataSet55.Clientes.Rows.Count - 1
                If CServiciosDataSet55.Clientes.Rows(x).Item("Cliente") = _Cliente Then
                    _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                End If
                If CServiciosDataSet55.Clientes.Rows(x).Item("Referencia") = _Cliente Then
                    _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                End If

            Next

            ' _Busca = CServiciosDataSet55.Clientes.FindByCliente(_Cliente)

            _Diagnostico = _Rows_Vf(i).Item("Diagnostico")

            _Rows_D = CServiciosDataSet55.Diagnosticos.Select("Diagnostico = " & "'" & _Diagnostico & "'")

            _Find = False
            For x = 0 To _Rows_D.GetUpperBound(0)
                _NombreDiagnostico = _Rows_D(x).Item("Descripcion")
                _Find = True
            Next

            If Not _Find Then
                _NombreDiagnostico = "SIN DIAGNOSTICO"
                _Diagnostico = "DI0002"
            End If


            DataGridView1.Rows.Add(_Folio, _Lote, _Fecha, _Cantidad, _PrecioNeto, _ImporteDescuento, _PorcentajeDescuento, _NombrePaciente, _Edad, _Medico, _Cedula, _Comentarios, _Factura, _Cliente, _NombreCliente, _Diagnostico, _NombreDiagnostico)
        Next

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox35.Text = "P"
        Button17_Click(sender, e)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        TextBox35.Text = "I"
        Button17_Click(sender, e)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        TextBox28.Text = "D"
        Button18_Click(sender, e)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        TextBox28.Text = "P"
        Button18_Click(sender, e)
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TextBox28.Text = "I"
        Button18_Click(sender, e)
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        ' Muestra el detalle
        Dim _Find As Boolean
        Dim _Cantidad As Integer
        Dim _Edad As String
        Dim _Fecha As Date
        Dim _Factura As String
        Dim _NombreCliente, _NombreDiagnostico As String
        Dim _Rows_D() As DataRow

        DataGridView1.Rows.Clear()


        _find = False
        For Me.i = 0 To _Rows_Vf.GetUpperBound(0)

            _Find = False
            If TextBox28.Text = "D" And _Rows_Vf(i).Item("ImporteDescuento") > 0 Then
                _Find = True
            End If
            If TextBox28.Text = "P" And _Rows_Vf(i).Item("Caja") = "CL13" Then
                _Find = True
            End If
            If TextBox28.Text = "I" And _Rows_Vf(i).Item("Caja") <> "CL13" Then
                _Find = True
            End If


            If _Find Then

                _Folio = _Rows_Vf(i).Item("Folio")
                _Lote = _Rows_Vf(i).Item("Referencia")
                _Cantidad = _Rows_Vf(i).Item("NumeroFrascos")
                _PrecioNeto = _Rows_Vf(i).Item("PrecioNeto")
                _ImporteDescuento = _Rows_Vf(i).Item("ImporteDescuento")
                _PorcentajeDescuento = _Rows_Vf(i).Item("PorcentajeDescuento")
                _NombrePaciente = _Rows_Vf(i).Item("NombrePaciente")
                _Edad = _Rows_Vf(i).Item("Paciente")
                _Medico = _Rows_Vf(i).Item("Medico")
                _Cedula = _Rows_Vf(i).Item("Cedula")
                _Fecha = _Rows_Vf(i).Item("FechaFolio")
                _Comentarios = _Rows_Vf(i).Item("Comentarios")
                If DBNull.Value.Equals(_Rows_Vf(i).Item("Factura")) Then
                    _Factura = ""
                Else
                    _Factura = _Rows_Vf(i).Item("Factura")
                End If
                _Cliente = _Rows_Vf(i).Item("Caja")

                _NombreCliente = "** CLIENTE NO REGISTRADO **"
                For x = 0 To CServiciosDataSet55.Clientes.Rows.Count - 1
                    If CServiciosDataSet55.Clientes.Rows(x).Item("Cliente") = _Cliente Then
                        _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                    End If
                    If CServiciosDataSet55.Clientes.Rows(x).Item("Referencia") = _Cliente Then
                        _NombreCliente = CServiciosDataSet55.Clientes.Rows(x).Item("NombreCliente")
                    End If
                Next

                ' _Busca = CServiciosDataSet55.Clientes.FindByCliente(_Cliente)

                _Diagnostico = _Rows_Vf(i).Item("Diagnostico")

                _Rows_D = CServiciosDataSet55.Diagnosticos.Select("Diagnostico = " & "'" & _Diagnostico & "'")

                _Find = False
                For x = 0 To _Rows_D.GetUpperBound(0)
                    _NombreDiagnostico = _Rows_D(x).Item("Descripcion")
                    _Find = True
                Next

                If Not _Find Then
                    _NombreDiagnostico = "SIN DIAGNOSTICO"
                    _Diagnostico = "DI0002"
                End If


                DataGridView1.Rows.Add(_Folio, _Lote, _Fecha, _Cantidad, _PrecioNeto, _ImporteDescuento, _PorcentajeDescuento, _NombrePaciente, _Edad, _Medico, _Cedula, _Comentarios, _Factura, _Cliente, _NombreCliente, _Diagnostico, _NombreDiagnostico)


            End If



        Next



    End Sub

    Private Sub MonthCalendar4_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar4.DateChanged

    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox9.Focus()
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        ComboBox8.Visible = True
        ComboBox9.Visible = True
        ComboBox8.Text = Format(Today, "Long Date")
        ComboBox9.Text = Format(Today, "Long Date")
        TextBox36.Text = Today
        TextBox37.Text = Today

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox10.Focus()
        End If
    End Sub

    Private Sub MonthCalendar5_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar5.DateChanged

    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As EventArgs) Handles TextBox10.GotFocus
        If TextBox10.Text = "" Then
            TextBox10.Text = TextBox3.Text
        End If

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged

    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox4.Focus()
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        ComboBox8.Text = ""
        ComboBox9.Text = ""
        TextBox36.Text = ""
        TextBox37.Text = ""
        ComboBox8.Visible = False
        ComboBox9.Visible = False
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click






        ' Exporta a Excel
        Dim X As Integer
        Dim _Columna As Integer
        Dim _Linea As Integer
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Formatos Factor\Reporte Consultas.xlsx"
        m_Excel = CreateObject("Excel.Application")

        On Error Resume Next
        m_Excel.Workbooks.Open(strRutaExcel)

        If Err.Number <> 0 Then
            msg = "Este libro de Excel ya está abierto. Debe cerrarlo primero para generar esta consulta"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If



        m_Excel.Visible = False
        '  m_Excel.Worksheets("FT").close(False)

        Label38.Visible = True
        ProgressBar1.Visible = True
        ProgressBar1.Maximum = DataGridView1.Rows.Count
        ProgressBar1.Minimum = 1

        _Linea = 7
        For Me.i = 7 To 60000
            'MsgBox("Valor de I " & CStr(i))
            If m_Excel.Worksheets("FT").cells(_Linea, 1).value = "" Then
                Exit For
            End If


            For X = 1 To 17
                m_Excel.Worksheets("FT").cells(_Linea, X).value = ""
            Next
            _Linea = _Linea + 1
        Next

        _Linea = 7

        For Me.i = 0 To DataGridView1.Rows.Count - 1
            ProgressBar1.Value = i + 1
            _Columna = 1

            For X = 0 To 16
                m_Excel.Worksheets("FT").cells(_Linea, _Columna).value = DataGridView1.Rows(i).Cells(X).Value
                _Columna = _Columna + 1

            Next
            _Linea = _Linea + 1

            If (_Linea Mod 2) = 0 Then
                Label38.Visible = True
            Else
                Label38.Visible = False
            End If
        Next

        Label38.Visible = False
        ProgressBar1.Visible = False
        m_Excel.Visible = True
        'm_Excel.Workbooks.Close()





    End Sub

    Private Sub ComboBox4_GotFocus(sender As Object, e As EventArgs) Handles ComboBox4.GotFocus
        If ComboBox4.Text = "" Then
            ComboBox4.Text = "PUBLICO EN GENERAL"
            TextBox11.Text = "CL13"
        End If

    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As EventArgs) Handles TextBox9.GotFocus
        If TextBox9.Text = "" Then
            TextBox9.Text = "N/A"
        End If

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(0).Value
        ComboBox1.Text = _Renglon.Cells(1).Value
        ComboBox2.Text = _Renglon.Cells(2).Value
        TextBox4.Text = _Renglon.Cells(3).Value
        TextBox5.Text = _Renglon.Cells(4).Value
        TextBox6.Text = _Renglon.Cells(5).Value
        TextBox7.Text = _Renglon.Cells(6).Value
        TextBox3.Text = _Renglon.Cells(7).Value
        ComboBox3.Text = _Renglon.Cells(8).Value
        TextBox8.Text = _Renglon.Cells(9).Value
        TextBox9.Text = _Renglon.Cells(10).Value
        TextBox10.Text = _Renglon.Cells(11).Value
        TextBox11.Text = _Renglon.Cells(13).Value
        ComboBox4.Text = _Renglon.Cells(14).Value
        TextBox5.Text = _Renglon.Cells(15).Value
        ComboBox5.Text = _Renglon.Cells(16).Value


        For Each Me._Checacontrol In TabPage1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = False
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Enabled = False
            End If
        Next


    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub DataGridView2_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellLeave
        _Renglon = DataGridView2.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ComboBox6_Click(sender As Object, e As EventArgs) Handles ComboBox6.Click
        MonthCalendar2.Visible = True
        MonthCalendar3.Visible = False
        ComboBox6.DroppedDown = False

    End Sub

    Private Sub MonthCalendar2_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar2.DateSelected
        TextBox20.Text = MonthCalendar2.SelectionRange.Start
        MonthCalendar2.Visible = False
        Dim _Today As Date = CDate(TextBox20.Text)
        ComboBox6.Text = Format(_Today, "Long Date")

    End Sub

    Private Sub ComboBox7_Click(sender As Object, e As EventArgs) Handles ComboBox7.Click
        MonthCalendar2.Visible = False
        MonthCalendar3.Visible = True
        ComboBox7.DroppedDown = False


    End Sub

    Private Sub MonthCalendar3_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar3.DateSelected
        TextBox21.Text = MonthCalendar3.SelectionRange.Start
        MonthCalendar3.Visible = False
        Dim _Today As Date = CDate(TextBox21.Text)
        ComboBox7.Text = Format(_Today, "Long Date")
    End Sub

    Private Sub MonthCalendar2_LostFocus(sender As Object, e As EventArgs) Handles MonthCalendar2.LostFocus
        MonthCalendar2.Visible = False

    End Sub

    Private Sub MonthCalendar3_LostFocus(sender As Object, e As EventArgs) Handles MonthCalendar3.LostFocus
        MonthCalendar3.Visible = False
    End Sub

    Private Sub MonthCalendar4_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar4.DateSelected
        Dim _FechaFinal As Date = MonthCalendar4.SelectionRange.Start
        TextBox37.Text = _FechaFinal
        ComboBox8.Text = Format(_FechaFinal, "Long Date")
        MonthCalendar4.Visible = False

    End Sub

    Private Sub ComboBox9_Click(sender As Object, e As EventArgs) Handles ComboBox9.Click
        ComboBox9.DroppedDown = False
        MonthCalendar5.Visible = True

    End Sub

    Private Sub MonthCalendar5_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar5.DateSelected
        Dim _FechaInicial As Date = MonthCalendar5.SelectionRange.Start
        TextBox36.Text = _FechaInicial
        ComboBox9.Text = Format(_FechaInicial, "Long Date")
        MonthCalendar5.Visible = False
    End Sub

    Private Sub ComboBox8_Click(sender As Object, e As EventArgs) Handles ComboBox8.Click
        MonthCalendar4.Visible = True
        ComboBox8.DroppedDown = False
    End Sub

    Private Sub Button5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Button5.KeyPress

    End Sub
End Class