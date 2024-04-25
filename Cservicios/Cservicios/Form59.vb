Public Class Frm59HistorialRemisiones
    Public i As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Busca As DataRow
    Public _Checacontrol As Control
    Public _Remision, _Hora As String
    Public _Fecha As Date
    Public _Producto As String
    Public _Centro As String = LoginForm1.TextBox1.Text
    Public _Cantidad As Integer
    Public _Rows_Re() As DataRow
    Public _PrecioUnitario As Double
    Public _Cliente, _Paciente, _NombreCliente, _NombrePaciente As String

    Private Sub Frm59HistorialRemisiones_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Escape) Then

            If TextBox21.Text = "P" Then
                For Each Me._Checacontrol In Panel3.Controls
                    If TypeOf _Checacontrol Is TextBox Then
                        _Checacontrol.Text = ""
                    End If
                    If TypeOf _Checacontrol Is ComboBox Then
                        _Checacontrol.Text = ""
                    End If
                Next
                Panel3.Visible = False
            End If

            If TextBox21.Text = "D" Then

                TextBox20.Text = ""
                Timer1.Enabled = False
                _Renglon = DataGridView2.CurrentRow
                _Renglon.DefaultCellStyle.BackColor = Color.White
            End If

            If TextBox21.Text = "C" Then
                TextBox20.Text = ""
                Timer1.Enabled = False
                RichTextBox1.BackColor = Color.White
                RichTextBox1.Text = ""
                Button1_Click(sender, e)
            End If


            TextBox21.Text = ""
        End If
    End Sub



    Private Sub Form59_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.Clientes' Puede moverla o quitarla según sea necesario.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet51.DetalleRemision' Puede moverla o quitarla según sea necesario.
        Me.DetalleRemisionTableAdapter.Fill(Me.CServiciosDataSet51.DetalleRemision)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet50.Remisiones' Puede moverla o quitarla según sea necesario.
        Me.RemisionesTableAdapter.Fill(Me.CServiciosDataSet50.Remisiones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet21.Productos' Puede moverla o quitarla según sea necesario.
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)

        RadioButton1.Checked = True
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        Panel3.Visible = False

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.BackColor = Color.White
            End If

        Next

        RichTextBox1.BackColor = Color.White
        For Each Me._Checacontrol In Panel3.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next
        ComboBox3.Items.Clear()

        For Me.i = 0 To CServiciosDataSet50.Remisiones.Rows.Count - 1
            _Remision = CServiciosDataSet50.Remisiones.Rows(i).Item("Remision")
            _Fecha = CServiciosDataSet50.Remisiones.Rows(i).Item("Fecha")
            _Paciente = CServiciosDataSet50.Remisiones.Rows(i).Item("Paciente")

            _Busca = CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
            _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")

            DataGridView1.Rows.Add(_Remision, _Fecha, _NombrePaciente)
        Next

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = False
            End If
        Next

        Dim _NombreProducto As String
        ' Llena el Combo con los productos
        ComboBox3.Items.Clear()
        For Me.i = 0 To CServiciosDataSet21.Productos.Rows.Count - 1
            _NombreProducto = CServiciosDataSet21.Productos.Rows(i).Item("Descripcion")
            ComboBox3.Items.Add(_NombreProducto)
        Next
        ComboBox3.Sorted = True
        RichTextBox1.Text = ""


    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        If _Renglon.Cells(3).Value = False Then
            _Renglon.Cells(3).Value = True
            _Renglon.DefaultCellStyle.BackColor = Color.Gold
        Else
            _Renglon.Cells(3).Value = False
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If


        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = True
            End If
        Next

        _Remision = _Renglon.Cells(0).Value
        _Busca = CServiciosDataSet50.Remisiones.FindByRemision(_Remision)
        TextBox1.Text = _Remision
        TextBox2.Text = _Busca.Item("Fecha")
        TextBox3.Text = _Busca.Item("Hora")
        TextBox4.Text = _Busca.Item("Cliente")
        TextBox10.Text = _Busca.Item("Paciente")
        TextBox12.Text = _Busca.Item("Referencia")
        TextBox19.Text = _Busca.Item("Estatus")
        If TextBox19.Text = "ACTIVA" Then
            TextBox19.BackColor = Color.LightGreen
        End If
        If TextBox19.Text = "MATERIAL ENTREGADO" Then
            TextBox19.BackColor = Color.Yellow
        End If
        If TextBox19.Text = "CANCELADA" Then
            TextBox19.BackColor = Color.LightCoral
        End If
        If TextBox19.Text = "FACTURADA" Then
            TextBox19.BackColor = Color.PeachPuff
        End If


        RichTextBox1.Text = _Busca.Item("Comentarios")

        If Not DBNull.Value.Equals(_Busca.Item("Factura")) Then
            TextBox22.Text = _Busca.Item("Factura")
        End If


        _Busca = CServiciosDataSet30.Clientes.FindByCliente(TextBox4.Text)
        TextBox5.Text = _Busca.Item("NombreCliente")
        TextBox6.Text = _Busca.Item("Calle")
        TextBox7.Text = _Busca.Item("Colonia")
        TextBox8.Text = _Busca.Item("Ciudad")
        TextBox9.Text = _Busca.Item("Estado")

        _Busca = CServiciosDataSet2.Pacientes.FindByPaciente(TextBox10.Text)
        _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")
        TextBox11.Text = _NombrePaciente

        ' Muestra los productos incluídos
        ToolStripButton4_Click(sender, e)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        '   _Renglon = DataGridView1.CurrentRow
        '   _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        _Remision = TextBox1.Text
        DataGridView2.Rows.Clear()
        Dim _Rows_dd() As DataRow = CServiciosDataSet51.DetalleRemision.Select("Remision = " & "'" & _Remision & "'")
        Dim _Producto, _NombreProducto As String
        Dim _PrecioUnitario, _Importe, _Descuento, _ImporteNeto As Double
        Dim _Cantidad As Integer
        Dim _Id As Long

        DataGridView2.Rows.Clear()
        For Me.i = 0 To _Rows_dd.GetUpperBound(0)
            _Producto = _Rows_dd(i).Item("Producto")
            _Busca = CServiciosDataSet21.Productos.FindByProducto(_Producto)
            _NombreProducto = _Busca.Item("Descripcion")
            _PrecioUnitario = _Rows_dd(i).Item("PrecioUnitario")
            _Cantidad = _Rows_dd(i).Item("Cantidad")
            _Importe = _Rows_dd(i).Item("Importe")
            _Descuento = _Rows_dd(i).Item("Descuento")
            _ImporteNeto = _Rows_dd(i).Item("ImporteNeto")
            _Id = _Rows_dd(i).Item("Orden")

            DataGridView2.Rows.Add(_Producto, _NombreProducto, _PrecioUnitario, _Cantidad, _Importe, _Descuento, _ImporteNeto, _Id)
        Next

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click

        If TextBox22.Text <> "" Then
            Msg = "Esta Remisión no puede incrementar los productos, ya está Facturada. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        Panel3.Visible = True
        For Each Me._Checacontrol In Panel3.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next
        TextBox21.Text = "P"
        TextBox13.Focus()

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

        For Me.i = 0 To CServiciosDataSet21.Productos.Rows.Count - 1
            If CServiciosDataSet21.Productos.Rows(i).Item("Descripcion") = ComboBox3.Text Then
                TextBox13.Text = CServiciosDataSet21.Productos.Rows(i).Item("Producto")
                TextBox14.Text = CServiciosDataSet21.Productos.Rows(i).Item("PrecioUnitario")
            End If
        Next
        TextBox13.Enabled = False
        TextBox15.Focus()


    End Sub

    Private Sub TextBox15_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox15.GotFocus
        TextBox15.Text = 1
    End Sub

    Private Sub TextBox15_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox15.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox15_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox15.LostFocus
        TextBox16.Text = CDbl(TextBox14.Text) * CInt(TextBox15.Text)
    End Sub

    Private Sub TextBox15_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox17_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox17.GotFocus
        TextBox17.Text = 0
    End Sub

    Private Sub TextBox17_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox17.LostFocus
        TextBox18.Text = CDbl(TextBox15.Text) * CDbl(TextBox16.Text)

    End Sub

    Private Sub TextBox17_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub TextBox18_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox18.GotFocus

    End Sub

    Private Sub TextBox18_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox18.TextChanged

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ' Agrega al Datagrid

        Dim _Completo As Boolean

        _Completo = True
        For Each Me._Checacontrol In Panel3.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Text = "" Then
                    _Completo = False
                End If
            End If
        Next
        If _Completo = False Then
            Msg = "Hay casillas en blanco. Debe llenar todos los datos"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        _Producto = TextBox13.Text
        Dim _DescripcionProducto As String = ComboBox3.Text
        _PrecioUnitario = CDbl(TextBox14.Text)
        _Cantidad = CInt(TextBox15.Text)
        Dim _Importe As Double = CDbl(TextBox16.Text)
        Dim _Descuento As Double = CDbl(TextBox17.Text)
        Dim _ImporteNeto As Double = CDbl(TextBox18.Text)

        ' Agrega el registro
        ' Registra el detalle de la remisión
        _Descuento = 0
        Dim _Neto As Double
        _Remision = TextBox1.Text
    
        _Producto = TextBox13.Text
        _PrecioUnitario = CDbl(TextBox14.Text)
        _Cantidad = CInt(TextBox15.Text)
        _Descuento = CDbl(TextBox17.Text)
        _Neto = CDbl(TextBox18.Text)
        _Importe = CDbl(TextBox16.Text)

            _Busca = CServiciosDataSet51.DetalleRemision.NewRow
            _Busca.Item("Remision") = _Remision
            _Busca.Item("Producto") = _Producto
            _Busca.Item("PrecioUnitario") = _PrecioUnitario
            _Busca.Item("Cantidad") = _Cantidad
            _Busca.Item("Importe") = _Importe
            _Busca.Item("Descuento") = _Descuento
            _Busca.Item("ImporteNeto") = _Neto

            CServiciosDataSet51.DetalleRemision.Rows.Add(_Busca)
 

        Me.Validate()
        DetalleRemisionBindingSource.EndEdit()
        Me.TableAdapterManager3.UpdateAll(CServiciosDataSet51)
        Me.DetalleRemisionTableAdapter.Fill(Me.CServiciosDataSet51.DetalleRemision)

        Msg = "Registro guardado correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        ToolStripButton4_Click(sender, e)

        For Each Me._Checacontrol In Panel3.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next
        Panel3.Visible = False
        '  DataGridView2.Rows.Add(_Producto, _DescripcionProducto, _PrecioUnitario, _Cantidad, _Importe, _Descuento, _ImporteNeto)
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        TextBox21.Text = "D"
        TextBox20.Text = "Para eliminar este Producto, mouse derecho"
        TextBox20.ForeColor = Color.Red
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        Timer1.Enabled = True
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick


    End Sub

    Private Sub EliminarEsteMaterialToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EliminarEsteMaterialToolStripMenuItem.Click
        ' Elimina el material

        If TextBox22.Text <> "" Then
            Msg = "Esta remisión ya está Facturada, no puede cancelar productos"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If



        Timer1.Interval = 500

        Dim title As String
        Dim _Id As Integer
        Dim response As MsgBoxResult
        Msg = "Está seguro de eliminar este Producto?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Eliminar Materiales de esta Remisión"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Eliina el registro
            _Id = _Renglon.Cells(7).Value

            _Busca = CServiciosDataSet51.DetalleRemision.FindByOrden(_Id)
            _Busca.Delete()

            Me.Validate()
            DetalleRemisionBindingSource.EndEdit()
            Me.TableAdapterManager3.UpdateAll(CServiciosDataSet51)
            Me.DetalleRemisionTableAdapter.Fill(Me.CServiciosDataSet51.DetalleRemision)

            Msg = "Producto eliminado correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)

            ToolStripButton4_Click(sender, e)
            For Each Me._Checacontrol In Panel3.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    _Checacontrol.Text = ""
                End If
                If TypeOf _Checacontrol Is ComboBox Then
                    _Checacontrol.Text = ""
                End If
            Next
            Panel3.Visible = False
          
        Else

        End If

        TextBox20.Text = ""
        Timer1.Enabled = False
    End Sub

    Private Sub DataGridView2_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellLeave
        _Renglon = DataGridView2.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        TextBox20.Text = " " & TextBox20.Text
    End Sub

    Private Sub TextBox21_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox21.TextChanged

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        ' Cancelar la remisión
        TextBox21.Text = "C"

        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar una Remisión para poder Cancelarla"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        _Remision = TextBox1.Text
        Dim _Rows_Re() As DataRow = CServiciosDataSet50.Remisiones.Select("Remision = " & "'" & _Remision & "'")
        Dim _Cancel As Boolean

        _Cancel = False

        For Me.i = 0 To _Rows_Re.GetUpperBound(0)
            If _Rows_Re(i).Item("Estatus") <> "ACTIVA" Then
                _Cancel = True
            End If
        Next

        If _Cancel Then
            Msg = "Esta Remisión no puede CANCELARLA. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de Cancelar la Remisión Totalmente ? "
        Style = MsgBoxStyle.DefaultButton2 Or _
         MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cancelar Esta Remisión"
        response = MsgBox(Msg, Style, title)

        If response = MsgBoxResult.Yes Then

            ToolStripButton2.Enabled = False
            ToolStripButton3.Enabled = True

            For Each Me._Checacontrol In Me.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    _Checacontrol.Enabled = False
                End If
            Next
            RichTextBox1.Enabled = True
            RichTextBox1.BackColor = Color.PeachPuff
            RichTextBox1.Focus()
            Timer1.Enabled = True
            TextBox20.Enabled = True
            TextBox20.Text = "Teclee el Motivo para Cancelar esta Remisión, luego, el ícono de Guardar Información"
            TextBox20.ForeColor = Color.Red

        Else


        End If



    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click

        ' Cancela la Remisión
        _Remision = TextBox1.Text
        _Busca = CServiciosDataSet50.Remisiones.FindByRemision(_Remision)
        _Busca.Item("Estatus") = "CANCELADA"
        _Busca.Item("Comentarios") = UCase(RichTextBox1.Text)
        RichTextBox1.BackColor = Color.White

        Me.Validate()
        RemisionesBindingSource.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet50)
        Me.RemisionesTableAdapter.Fill(Me.CServiciosDataSet50.Remisiones)

        Msg = "Remisión Cancelada Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        Timer1.Enabled = False
        TextBox20.Text = ""
        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click

        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de Imprimir esta Remisión?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Impresión de Notas de Remisión"

        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar una Remisión para poder Imprimir"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If TextBox19.Text <> "ACTIVA" Then
            Msg = "Esta remisión no puede imprimirla nuevamente. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Imprime
            Dim m_Excel As Microsoft.Office.Interop.Excel.Application
            Dim strRutaExcel As String = "Z:\Formatos Mezclas\Formato Remision.xlsx"
            m_Excel = CreateObject("Excel.Application")
            m_Excel.Workbooks.Open(strRutaExcel)
            m_Excel.Visible = False

            Dim _Meses(12) As String
            _Meses(0) = "ENE"
            _Meses(1) = "FEB"
            _Meses(2) = "MAR"
            _Meses(3) = "ABR"
            _Meses(4) = "MAY"
            _Meses(5) = "JUN"
            _Meses(6) = "JUL"
            _Meses(7) = "AGO"
            _Meses(8) = "SEP"
            _Meses(9) = "OCT"
            _Meses(10) = "NOV"
            _Meses(11) = "DIC"

            Dim _Mes As Integer = Month(_Fecha) - 1
            Dim _NombreFecha As String
            _NombreFecha = "Fecha :  " & Mid(CStr(_Fecha), 1, 3) & _Meses(_Mes) & "/" & Mid(CStr(_Fecha), 7, 4)

            If RadioButton1.Checked = True Then
                _Remision = TextBox1.Text
                _Cliente = TextBox1.Text
                _Paciente = TextBox10.Text
                _Fecha = TextBox2.Text

                m_Excel.Worksheets("HORIZONTAL").cells(2, 8).value = _Remision
                m_Excel.Worksheets("HORIZONTAL").cells(8, 3).value = _Cliente
                m_Excel.Worksheets("HORIZONTAL").cells(8, 5).value = _Fecha
                m_Excel.Worksheets("HORIZONTAL").cells(9, 2).value = TextBox5.Text
                m_Excel.Worksheets("HORIZONTAL").cells(10, 2).value = TextBox6.Text
                m_Excel.Worksheets("HORIZONTAL").cells(11, 2).value = TextBox7.Text
                m_Excel.Worksheets("HORIZONTAL").cells(12, 2).value = TextBox8.Text & ", " & TextBox9.Text
                m_Excel.Worksheets("HORIZONTAL").cells(8, 9).value = _Paciente
                m_Excel.Worksheets("HORIZONTAL").cells(9, 7).value = TextBox11.Text

                Dim _Cuantos As Integer
                Dim _Linea As Integer
                _Linea = 14
                _Cuantos = 0

                _PrecioUnitario = 0

                For Me.i = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView2.Rows(i).Cells(0).Value = "" Then
                        Exit For
                    End If


                    If CInt(DataGridView2.Rows(i).Cells(6).Value) = 0 Then
                        _PrecioUnitario = 0
                    Else
                        _PrecioUnitario = DataGridView2.Rows(i).Cells(6).Value / DataGridView2.Rows(i).Cells(3).Value
                    End If

                    m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 2).value = DataGridView2.Rows(i).Cells(0).Value
                    m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 3).value = DataGridView2.Rows(i).Cells(1).Value
                    m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 8).value = _PrecioUnitario
                    m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 9).value = DataGridView2.Rows(i).Cells(3).Value
                    m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 10).value = DataGridView2.Rows(i).Cells(6).Value
                    _Linea = _Linea + 1
                    _Cuantos = _Cuantos + 1
                Next

                m_Excel.Worksheets("HORIZONTAL").cells(32, 3).value = _Cuantos

            End If

            '************************************************************************************************************************
            If RadioButton1.Checked = True Then
                m_Excel.Worksheets("VERTICAL").cells(3, 6).value = _Remision
                m_Excel.Worksheets("VERTICAL").cells(7, 6).value = _NombreFecha
                m_Excel.Worksheets("VERTICAL").cells(10, 2).value = TextBox11.Text
                m_Excel.Worksheets("VERTICAL").cells(11, 2).value = TextBox5.Text
                m_Excel.Worksheets("VERTICAL").cells(12, 2).value = TextBox6.Text
                m_Excel.Worksheets("VERTICAL").cells(13, 2).value = "COL. " & TextBox7.Text
                m_Excel.Worksheets("VERTICAL").cells(14, 2).value = TextBox8.Text & "," & TextBox9.Text


                Dim _Linea As Integer
                _Linea = 17

                For Me.i = 0 To DataGridView2.Rows.Count - 1
                    If DataGridView2.Rows(i).Cells(0).Value = "" Then
                        Exit For
                    End If
                    _PrecioUnitario = DataGridView2.Rows(i).Cells(6).Value / DataGridView2.Rows(i).Cells(3).Value
                    m_Excel.Worksheets("VERTICAL").cells(_Linea, 1).value = DataGridView2.Rows(i).Cells(3).Value
                    m_Excel.Worksheets("VERTICAL").cells(_Linea, 2).value = DataGridView2.Rows(i).Cells(1).Value
                    m_Excel.Worksheets("VERTICAL").cells(_Linea, 7).value = _PrecioUnitario
                    m_Excel.Worksheets("VERTICAL").cells(_Linea, 9).value = DataGridView2.Rows(i).Cells(6).Value

                    _Linea = _Linea + 1
                Next

                m_Excel.Worksheets("VERTICAL").cells(45, 4).value = TextBox11.Text




            End If




            If CheckBox1.Checked = True Then
                m_Excel.Visible = True
            Else
                If RadioButton2.Checked = True Then
                    m_Excel.Worksheets("HORIZONTAL").Printout(1, 1)
                Else
                    m_Excel.Worksheets("VERTICAL").Printout(1, 1)
                End If

                m_Excel.ActiveWorkbook.Close(False)
            End If



            Msg = "Remisión enviada a la Impresora"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)


        Else
            ' Perform some other action.
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox22_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)

    End Sub

    Private Sub TextBox22_LostFocus(sender As Object, e As System.EventArgs)
        If TextBox13.Text = "" Then
            ' Busca por el combo
            ComboBox3.Focus()
        Else
            ' Busca el producto
            _Producto = TextBox13.Text
            Dim _Rows_Pr() As DataRow = CServiciosDataSet21.Productos.Select("Producto = " & "'" & _Producto & "'")
            Dim _Find As Boolean

            _Find = False
            For Me.i = 0 To _Rows_Pr.GetUpperBound(0)
                ComboBox3.Text = _Rows_Pr(i).Item("Descripcion")
                TextBox13.Text = TextBox13.Text
                _Find = True
                TextBox15.Focus()
                Exit For
            Next

            If _Find = False Then
                Msg = "No existe ese producto Registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

        End If
    End Sub

    Private Sub TextBox22_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub TextBox13_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox13.LostFocus
        Dim _Find As Boolean

        _Find = False
        _Producto = TextBox13.Text

        If TextBox13.Text = "" Then
            ' Busca el producto
            ComboBox3.Focus()
        Else

            Dim _Rows_pr() As DataRow = CServiciosDataSet21.Productos.Select("Producto = " & "'" & _Producto & "'")

            _Find = False
            For Me.i = 0 To _Rows_pr.GetUpperBound(0)
                _Find = True
                ComboBox3.Text = _Rows_pr(i).Item("Descripcion")
                TextBox14.Text = _Rows_pr(i).Item("PrecioUnitario")
                TextBox15.Focus()
                Exit For
            Next

            If _Find = False Then
                Msg = "No existe ese producto registrado en la tabla. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                TextBox13.Text = ""
                TextBox13.Focus()
                Exit Sub
            End If

        End If





    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ' Busqueda

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.BackColor = Color.White
            End If
        Next


        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()

        Dim _RevisaParte As String
        Dim _ParteNombre As String
        Dim _NumRem As String
        Dim _XNumRem As String
        Dim _Mide As Integer
        Dim X, R, Z As Integer
        Dim _Rows_Pa() As DataRow = Me.CServiciosDataSet2.Pacientes.Select("Centro = " & "'" & _Centro & "'")
        Dim _Estatus As String

        ' Por Remisión
        If RadioButton3.Checked = True Then
            _NumRem = UCase(InputBox("Teclee el número de Remisión a buscar"))
            _Rows_Re = Me.CServiciosDataSet50.Remisiones.Select("Remision = " & "'" & _NumRem & "'")
        End If

        ' Por Fecha
        If RadioButton4.Checked = True Then

            TextBox20.Text = "Teclle la fecha a buscar"
            Timer1.Enabled = True

            MonthCalendar1.Visible = True


            '  _Rows_Re = Me.CServiciosDataSet50.Remisiones.Select("Fecha = " & "'" & _Fecha & "'")

        End If

        ' Todas 

        If RadioButton8.Checked = True Then
            _RevisaParte = ""
            _Rows_Re = Me.CServiciosDataSet50.Remisiones.Select("Estatus <> " & "'" & _RevisaParte & "'")
        End If


        If RadioButton3.Checked = True Or RadioButton8.Checked = True Then
            For Me.i = 0 To _Rows_Re.GetUpperBound(0)
                _Remision = _Rows_Re(i).Item("Remision")
                _Fecha = _Rows_Re(i).Item("Fecha")
                _Paciente = _Rows_Re(i).Item("Paciente")
                _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")

                DataGridView1.Rows.Add(_Remision, _Fecha, _NombrePaciente, False)

            Next
        End If

        ' Por Paciente
        If RadioButton5.Checked = True Then
            _ParteNombre = InputBox("Teclee el Nombre del Paciente a buscar")
            _ParteNombre = UCase(_ParteNombre)
            _Mide = Len(_ParteNombre)



            For Me.i = 0 To Me.CServiciosDataSet50.Remisiones.Rows.Count - 1
                _Paciente = CServiciosDataSet50.Remisiones.Rows(i).Item("Paciente")
                _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")
                _Remision = CServiciosDataSet50.Remisiones.Rows(i).Item("Remision")
                _Fecha = CServiciosDataSet50.Remisiones.Rows(i).Item("Fecha")

                For R = 1 To Len(_NombrePaciente)
                    _RevisaParte = Mid(_NombrePaciente, R, _Mide)

                    If _RevisaParte = _ParteNombre Then
                        DataGridView1.Rows.Add(_Remision, _Fecha, _NombrePaciente, False)
                        Exit For
                    End If
                Next

            Next

        End If

        ' Por Cliente 
        If RadioButton6.Checked = True Then
            _ParteNombre = InputBox("Teclee el Nombre del Paciente a buscar")
            _ParteNombre = UCase(_ParteNombre)
            _Mide = Len(_ParteNombre)

            For Me.i = 0 To Me.CServiciosDataSet50.Remisiones.Rows.Count - 1
                _Cliente = CServiciosDataSet50.Remisiones.Rows(i).Item("Cliente")
                _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                _NombreCliente = _Busca.Item("NombreCliente")
                _Remision = CServiciosDataSet50.Remisiones.Rows(i).Item("Remision")
                _Fecha = CServiciosDataSet50.Remisiones.Rows(i).Item("Fecha")

                For R = 1 To Len(_NombreCliente)
                    _RevisaParte = Mid(_NombreCliente, R, _Mide)

                    If _RevisaParte = _ParteNombre Then
                        DataGridView1.Rows.Add(_Remision, _Fecha, _NombrePaciente, False)
                        Exit For
                    End If
                Next

            Next


        End If

        ' Por Estatus
        If RadioButton7.Checked = True Then
            _ParteNombre = InputBox("Teclee el Nombre del Estatus a buscar")
            _ParteNombre = UCase(_ParteNombre)
            _Mide = Len(_ParteNombre)

            For Me.i = 0 To Me.CServiciosDataSet50.Remisiones.Rows.Count - 1
                _Paciente = CServiciosDataSet50.Remisiones.Rows(i).Item("Paciente")
                _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                _NombreCliente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")
                _Remision = CServiciosDataSet50.Remisiones.Rows(i).Item("Remision")
                _Fecha = CServiciosDataSet50.Remisiones.Rows(i).Item("Fecha")
                _Estatus = CServiciosDataSet50.Remisiones.Rows(i).Item("Estatus")

                For R = 1 To Len(_Estatus)
                    _RevisaParte = Mid(_Estatus, R, _Mide)

                    If _RevisaParte = _ParteNombre Then
                        DataGridView1.Rows.Add(_Remision, _Fecha, _NombrePaciente, False)
                        Exit For
                    End If
                Next

            Next



        End If


    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox2.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False

        _Rows_Re = Me.CServiciosDataSet50.Remisiones.Select("Fecha = " & "'" & TextBox2.Text & "'")
        For Me.i = 0 To _Rows_Re.GetUpperBound(0)
            _Remision = _Rows_Re(i).Item("Remision")
            _Fecha = _Rows_Re(i).Item("Fecha")
            _Paciente = _Rows_Re(i).Item("Paciente")
            _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
            _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")

            DataGridView1.Rows.Add(_Remision, _Fecha, _NombrePaciente, False)
        Next


    End Sub

    Private Sub ACTIVAToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ACTIVAToolStripMenuItem.Click
        ' Regresa la Remisión al Estatus Activa
        '    If TextBox19.Text = "ACTIVA" Then
        ' Msg = "Esta remisión ya tiene ese Estatus de ACTIVA, verifique"
        ' Style = MsgBoxStyle.Information
        ' MsgBox(Msg, Style)
        ' Exit Sub
        ' End If

        Dim _Find, _Yasta As Boolean
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Esta a punto de asignar el estatus de ACTIVA a las remisiones seleccionadas.  Está  seguro?"
        Style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cambio de Estatus"
        _Find = False
        response = MsgBox(Msg, Style, title)
        If response = MsgBoxResult.Yes Then
            ' Cambia a Activa

            For Me.i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value = "" Then
                    Exit For
                End If

                If DataGridView1.Rows(i).Cells(3).Value = True Then
                    _Find = True
                    _Remision = DataGridView1.Rows(i).Cells(0).Value
                    _Busca = CServiciosDataSet50.Remisiones.FindByRemision(_Remision)
                    If _Busca.Item("Estatus") <> "ACTIVA" Then
                        _Busca.Item("Estatus") = "ACTIVA"
                        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") & vbCrLf & "REMISION ACTIVADA " & CStr(Today.Date)
                    Else
                        _Yasta = True
                        Msg = "La Remisión No. " & _Remision & " ya tiene el Estatus de ACTIVA. No será modificado este dato"
                        Style = MsgBoxStyle.Information
                        MsgBox(Msg, Style)

                    End If

                End If

            Next

            If _Find Then
                Me.Validate()
                RemisionesBindingSource.EndEdit()
                Me.TableAdapterManager2.UpdateAll(CServiciosDataSet50)
                Me.RemisionesTableAdapter.Fill(CServiciosDataSet50.Remisiones)
                Msg = "Registro modificado correctamente"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)

                Button1_Click(sender, e)
            End If

            Exit Sub
        Else
            Exit Sub
        End If



    End Sub

    Private Sub CambioDeEstatusToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CambioDeEstatusToolStripMenuItem.Click

    End Sub

    Private Sub MATERIALENTREGADOToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MATERIALENTREGADOToolStripMenuItem.Click
        ' If TextBox19.Text = "MATERIAL ENTREGADO" Then
        ' Msg = "Esta remisión ya tiene ese Estatus de MATERIAL ENTREGADO, verifique"
        ' Style = MsgBoxStyle.Information
        ' MsgBox(Msg, Style)
        ' Exit Sub
        ' End If

        Dim _Find, _Yasta As Boolean
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Esta a punto de asignar el estatus de MATERIAL ENTREGADO a las remisiones seleccionadas. Está seguro?"
        Style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cambio de Estatus"

        response = MsgBox(Msg, Style, title)
        If response = MsgBoxResult.Yes Then
            ' Cambia a MATERIAL ENTREGADO
            _Find = False
            _Yasta = False
            For Me.i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value = "" Then
                    Exit For
                End If

                _Remision = DataGridView1.Rows(i).Cells(0).Value
                _Busca = CServiciosDataSet50.Remisiones.FindByRemision(_Remision)

                If _Busca.Item("Estatus") <> "MATERIAL ENTREGADO" Then
                    _Busca.Item("Estatus") = "MATERIAL ENTREGADO"
                    _Busca.Item("Comentarios") = _Busca.Item("Comentarios") & vbCrLf & " MATERIALES ENTREGADOS " & CStr(Today.Date)
                    _Find = True
                Else
                    Msg = "La remision No. " & _Remision & " ya tiene asignado el estaus de MATERIAL ENTREGADO. Este dato no se modificará"
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                End If

            Next

            If _Find Then
                Me.Validate()
                RemisionesBindingSource.EndEdit()
                Me.TableAdapterManager2.UpdateAll(CServiciosDataSet50)
                Me.RemisionesTableAdapter.Fill(CServiciosDataSet50.Remisiones)

                Msg = "Registro modificado correctamente"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)

            End If

            Button1_Click(sender, e)
            Exit Sub
        Else
            Exit Sub
        End If




    End Sub

    Private Sub CONTRARECIBOToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CONTRARECIBOToolStripMenuItem.Click
        ' Cambio de Estatus a CANCELADA

        Dim _Find, _Yasta As Boolean
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está a punto de Cancelar las Remisiones seleccionadas. Está seguro de continuar?"   ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cancelación de Remisiones"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then

            _Find = False
            _Yasta = False
            Dim _Motivo As String = InputBox("Teclee el Motivo de la Cancelación")
            _Motivo = UCase(_Motivo)

            For Me.i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value = "" Then
                    Exit For
                End If
                _Remision = DataGridView1.Rows(i).Cells(0).Value

                _Busca = Me.CServiciosDataSet50.Remisiones.FindByRemision(_Remision)
                If _Busca.Item("Estatus") = "CANCELADA" Then
                    Msg = "La Remisiion No. " & _Remision & " ya está CANCELADA. No se modificará"
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                    _Yasta = True
                Else
                    _Busca.Item("Estatus") = "CANCELADA"
                    _Busca.Item("Comentarios") = _Busca.Item("Comentarios") & vbCrLf & "REMISION CANCELADA POR " & _Motivo & " " & CStr(Today.Date)
                    _Find = True
                End If

            Next

            If _Find Then
                Me.Validate()
                RemisionesBindingSource.EndEdit()
                Me.TableAdapterManager2.UpdateAll(CServiciosDataSet50)
                Me.RemisionesTableAdapter.Fill(CServiciosDataSet50.Remisiones)

                Msg = "Registros modificados correctamente"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
            End If
            Button1_Click(sender, e)


        Else
            ' Perform some other action.
        End If










    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub
End Class