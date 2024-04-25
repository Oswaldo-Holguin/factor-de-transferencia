Public Class Frm58Envios
    Public I As Integer
    Public _Busca As DataRow
    Public _Centro As String = LoginForm1.TextBox1.Text
    Public _Renglon As DataGridViewRow
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Remision As String
    Public _Fecha As String
    Public _Hora As String
    Public _Cliente As String
    Public _Paciente As String
    Public _Referencia As String
    Public _Producto As String
    Public _Cantidad As Double
    Public _Checacontrol As Control
    Public _PrecioUnitario As Double


    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub Frm58Envios_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Escape) Then
            If TextBox19.Text = "1" Then
                Button1_Click(sender, e)
            End If
            If TextBox19.Text = "2" Then
                Button3_Click(sender, e)
                TextBox7.Text = ""
            End If
            If TextBox19.Text = "3" Then
                For Each Me._Checacontrol In Panel3.Controls
                    If TypeOf _Checacontrol Is TextBox Then
                        _Checacontrol.Text = ""
                    End If
                    If TypeOf _Checacontrol Is ComboBox Then
                        _Checacontrol.Text = ""
                    End If
                    Panel3.Visible = False
                Next
            End If

        End If
    End Sub

    Private Sub Frm58Envios_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load


        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet21.Productos' Puede moverla o quitarla según sea necesario.
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet50.Remisiones' Puede moverla o quitarla según sea necesario.
        Me.RemisionesTableAdapter.Fill(Me.CServiciosDataSet50.Remisiones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet51.DetalleRemision' Puede moverla o quitarla según sea necesario.
        Me.DetalleRemisionTableAdapter.Fill(Me.CServiciosDataSet51.DetalleRemision)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.Clientes' Puede moverla o quitarla según sea necesario.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)


        CheckBox2.Checked = True
        CheckBox3.Checked = False
        RadioButton1.Checked = True
        TextBox11.Text = 0
        TextBox12.Text = 0

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
            '  _Checacontrol.Enabled = False
        Next
        RichTextBox1.Text = ""

        For Each Me._Checacontrol In Panel3.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next

        For Each Me._Checacontrol In Panel2.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next

        ToolStripButton1.Enabled = True
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False

        ComboBox1.Focus()
        Panel1.Enabled = False
        Panel2.Enabled = False
        Panel3.Visible = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False

        ' Llena el combo de los Clientes
        ComboBox1.Items.Clear()

        For Me.I = 0 To CServiciosDataSet30.Clientes.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet30.Clientes.Rows(I).Item("NombreCliente"))
        Next

        _Centro = LoginForm1.TextBox1.Text
        Dim _Rows_Pa() As DataRow = CServiciosDataSet2.Pacientes.Select("Centro = " & "'" & _Centro & "'")
        Dim _Nombre As String
        ComboBox2.Items.Clear()
        ComboBox2.Items.Add(" ** Agregar Nuevo Paciente ")
        Dim _Cuantos As Integer

        _Cuantos = 0
        For Me.I = 0 To _Rows_Pa.GetUpperBound(0)
            _Cuantos = _Cuantos + 1
        Next


        For Me.I = 0 To _Rows_Pa.GetUpperBound(0)
            _Nombre = _Rows_Pa(I).Item("PATERNO") & " " & _Rows_Pa(I).Item("MATERNO") & " " & _Rows_Pa(I).Item("NOMBRES")

            If CheckBox2.Checked = True Then
                If _Rows_Pa(I).Item("FechaAlta") = "SIN FECHA" Then
                    ComboBox2.Items.Add(_Nombre)
                End If
            End If
            If CheckBox3.Checked = True Then
                If _Rows_Pa(I).Item("FechaAlta") <> "SIN FECHA" Then
                    ComboBox2.Items.Add(_Nombre)
                End If
            End If


        Next
        ComboBox2.Sorted = True
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""

        DataGridView1.Rows.Clear()
        ComboBox4.Visible = False

        ToolStripTextBox1.Text = ""
        ToolStripTextBox1.Visible = False
        TextBox21.Text = "1"

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        ' Agrega una nueva remisión
        For Each Me._Checacontrol In Me.Controls
            _Checacontrol.Enabled = True
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""

        TextBox19.Text = "1"

        ' Obtiene el siguiente Documento
        ' Revisa si existe fecha predeterminada
        _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)

        Dim _FechaPredeterminada As String

        If DBNull.Value.Equals(_Busca.Item("Usuario")) Then
            _FechaPredeterminada = ""
        Else
            _FechaPredeterminada = _Busca.Item("Usuario")
        End If

        ' Dim _FechaPredeterminada As String = _Busca.Item("Usuario")

        If _FechaPredeterminada <> "" Then
            Msg = "Existe fecha predeterminadá. Esta se utilizará en todas las Remisiones que registre"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            TextBox9.Text = CDate(_FechaPredeterminada)
        Else
            TextBox9.Text = Date.Today
        End If

        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
        TextBox21.Text = "1"
        TextBox8.Focus()


    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox8.GotFocus

        Dim _Prefijo As String

        If Mid(Label8.Text, 1, 4) = "Nota" Then
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("RNP")
            _Prefijo = "RN"
        Else
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("COT")
            _Prefijo = "CO"
        End If


        Dim _Consecutivo As String = CStr(_Busca.Item("Consecutivo"))

        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = _Prefijo + Mid(_XConsecutivo, _x, 4)
        TextBox8.Text = _XConsecutivo


        TextBox10.Text = Mid(TimeOfDay, 1, 5)
        TextBox1.Focus()

    End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.GotFocus
        ComboBox1.BackColor = Color.PeachPuff
    End Sub

    Private Sub ComboBox1_LostFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.LostFocus
        ComboBox1.BackColor = Color.White
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim _Index As Integer = ComboBox1.SelectedIndex
        _Cliente = CServiciosDataSet30.Clientes.Rows(_Index).Item("Cliente")
        TextBox1.Text = _Cliente
        TextBox3.Text = CServiciosDataSet30.Clientes.Rows(_Index).Item("CALLE")
        TextBox4.Text = CServiciosDataSet30.Clientes.Rows(_Index).Item("COLONIA")
        TextBox5.Text = CServiciosDataSet30.Clientes.Rows(_Index).Item("CIUDAD")
        TextBox6.Text = CServiciosDataSet30.Clientes.Rows(_Index).Item("ESTADO")
        ComboBox2.Focus()
    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox2.GotFocus
        ComboBox2.BackColor = Color.PeachPuff
    End Sub

    Private Sub ComboBox2_Leave(sender As Object, e As System.EventArgs) Handles ComboBox2.Leave
        ComboBox2.BackColor = Color.White
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim _Index As Integer = ComboBox2.SelectedIndex
        Dim _Nuevo As Boolean

        _Nuevo = False
        If _Index = 0 Then
            _Nuevo = True
            '  Frm57NuevosPacientes.Show()
        Else
            Dim _Nombre As String
            _Centro = "CME"
            Dim _Rows_pa() As DataRow = CServiciosDataSet2.Pacientes.Select("Centro = " & "'" & _Centro & "'")
            For Me.I = 0 To _Rows_pa.GetUpperBound(0)
                _Nombre = _Rows_pa(I).Item("PATERNO") & " " & _Rows_pa(I).Item("MATERNO") & " " & _Rows_pa(I).Item("NOMBRES")
                If _Nombre = ComboBox2.Text Then
                    TextBox2.Text = _Rows_pa(I).Item("Paciente")
                    Exit For
                End If
            Next

        End If
        TextBox7.Focus()
        TextBox19.Text = "2"

        If _Nuevo Then
            Frm57NuevosPacientes.Show()
        End If

    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.BackColor = Color.PeachPuff
        TextBox7.Text = "SIN REFERENCIA"
    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        _Centro = "CME"
        Dim _Rows_Pa() As DataRow = CServiciosDataSet2.Pacientes.Select("Centro = " & "'" & _Centro & "'")
        Dim _Nombre As String
        ComboBox2.Text = ""
        ComboBox2.Items.Clear()
        ComboBox2.Items.Add(" ** Agregar Nuevo Paciente ")
        For Me.I = 0 To _Rows_Pa.GetUpperBound(0)
            _Nombre = _Rows_Pa(I).Item("PATERNO") & " " & _Rows_Pa(I).Item("MATERNO") & " " & _Rows_Pa(I).Item("NOMBRES")

            If CheckBox2.Checked = True Then
                If _Rows_Pa(I).Item("FechaAlta") = "SIN FECHA" Then
                    ComboBox2.Items.Add(_Nombre)
                End If
            End If

            If CheckBox3.Checked = True Then
                If _Rows_Pa(I).Item("FechaAlta") <> "SIN FECHA" Then
                    ComboBox2.Items.Add(_Nombre)
                End If
            End If



        Next
        ComboBox2.Sorted = True
    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        TextBox19.Text = "3"
        _Centro = "CME"
        Dim _Rows_Pr() As DataRow = CServiciosDataSet21.Productos.Select("Centro = " & "'" & _Centro & "'")

        Panel3.Visible = True
        Panel3.Enabled = True

        ' Llena el combo con los productos
        ComboBox3.Items.Clear()
        For Me.I = 0 To _Rows_Pr.GetUpperBound(0)
            ComboBox3.Items.Add(_Rows_Pr(I).Item("Descripcion"))
        Next
        ComboBox3.Sorted = True
        TextBox13.Focus()

    End Sub

    Private Sub ComboBox3_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox3.GotFocus
        ComboBox3.BackColor = Color.PeachPuff
    End Sub

    Private Sub ComboBox3_LostFocus(sender As Object, e As System.EventArgs) Handles ComboBox3.LostFocus
        ComboBox3.BackColor = Color.White
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        'Dim _Index As Integer = ComboBox3.SelectedIndex

        For Me.I = 0 To CServiciosDataSet21.Productos.Rows.Count - 1
            If CServiciosDataSet21.Productos.Rows(I).Item("Descripcion") = ComboBox3.Text Then
                TextBox13.Text = CServiciosDataSet21.Productos.Rows(I).Item("Producto")
                TextBox14.Text = CServiciosDataSet21.Productos.Rows(I).Item("PrecioUnitario")
            End If
        Next


        TextBox15.Focus()

    End Sub

    Private Sub TextBox15_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox15.GotFocus
        TextBox15.BackColor = Color.PeachPuff
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

        If Asc(e.KeyChar) = 13 Then
            TextBox17.Focus()
        End If
    End Sub

    Private Sub TextBox15_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox15.LostFocus
        TextBox15.BackColor = Color.White
        TextBox16.Text = CDbl(TextBox14.Text) * CDbl(TextBox15.Text)
        TextBox17.Focus()
    End Sub

    Private Sub TextBox15_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox17_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox17.GotFocus
        TextBox17.BackColor = Color.PeachPuff
        TextBox17.Text = 0
    End Sub

    Private Sub TextBox17_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox17.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Button2.Focus()
        End If
    End Sub

    Private Sub TextBox17_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox17.LostFocus
        TextBox17.BackColor = Color.White
        TextBox18.Text = CDbl(TextBox16.Text) - CDbl(TextBox17.Text)
    End Sub

    Private Sub TextBox17_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox17.TextChanged

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
        DataGridView1.Rows.Add(_Producto, _DescripcionProducto, _PrecioUnitario, _Cantidad, _Importe, _Descuento, _ImporteNeto)

        For Each Me._Checacontrol In Panel3.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If

        Next

        Dim _Suma As Double
        Dim _Cuenta As Integer

        _Suma = 0
        _Cuenta = 0
        For Me.I = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(I).Cells(0).Value = "" Then
                Exit For
            End If
            _Cuenta = _Cuenta + 1
            _Suma = _Suma + DataGridView1.Rows(I).Cells(6).Value

        Next
        TextBox11.Text = _Cuenta
        TextBox12.Text = Format(_Suma, "$#,###,###.00")
        TextBox20.Text = _Suma
        TextBox13.Focus()

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        ' Refresca información de los materiales
        For Each Me._Checacontrol In Panel3.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If

        Next
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        ' Llena el combo con los productos

        _Centro = "CME"
        Dim _Rows_Pr() As DataRow = CServiciosDataSet21.Productos.Select("Centro = " & "'" & _Centro & "'")
        ComboBox3.Items.Clear()
        For Me.I = 0 To _Rows_Pr.GetUpperBound(0)
            ComboBox3.Items.Add(_Rows_Pr(I).Item("Descripcion"))
        Next
    End Sub

    Private Sub EliminarEsteProductoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EliminarEsteProductoToolStripMenuItem.Click
        DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        Dim _Suma As Double
        Dim _Cuenta As Integer

        _Suma = 0
        _Cuenta = 0
        For Me.I = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(I).Cells(0).Value = "" Then
                Exit For
            End If
            _Cuenta = _Cuenta + 1
            _Suma = _Suma + DataGridView1.Rows(I).Cells(6).Value

        Next
        TextBox11.Text = _Cuenta
        TextBox12.Text = Format(_Suma, "$#,###,###.00")
        TextBox20.Text = _Suma
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click

        ' Registra la Remisión
        If Trim(TextBox1.Text) = "" Then
            Msg = "Debe seleccionar un Cliente para registrar este Envío"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        If Trim(TextBox2.Text) = "" Then
            Msg = "Debe seleccionar un Paciente para registrar este Envío"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        If CInt(TextBox11.Text) = 0 Then
            Msg = "Debe agregar al menos un producto para registrar este Envío"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If TextBox21.Text = "" Then
            TextBox21.Text = "1"
        End If


        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de guardar esta Remisión?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Envíos de Nutriciones"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Guarda el envío

            ToolStripTextBox1.Visible = True
            ToolStripTextBox1.Text = "Guardando Información e Imprimiendo Remisión NO. " & _Remision
            Timer1.Enabled = True
            Timer1.Interval = 100

            _Remision = TextBox8.Text
            _Fecha = TextBox9.Text
            _Hora = TextBox10.Text
            _Cliente = TextBox1.Text
            _Paciente = TextBox2.Text
            Dim _Importe As Double = CDbl(TextBox20.Text)
            Dim _Iva As Double = 0
            Dim _Neto As Double = _Importe
            Dim _Estatus As String = "ACTIVA"
            Dim _Factura As String = "SF"
            Dim _Documento As String = "SIN DOCUMENTO"
            _Referencia = ComboBox2.Text
            Dim _Comentarios As String = UCase(RichTextBox1.Text)
            Dim _FechaRegistro As Date = Today
            Dim _HoraRegistro As String = Mid(TimeOfDay, 1, 5)
            _Centro = "CME"
            Dim _Usuario As String = LoginForm1.TextBox6.Text


            _Busca = CServiciosDataSet50.Remisiones.NewRow
            _Busca.Item("Remision") = _Remision
            _Busca.Item("Fecha") = _Fecha
            _Busca.Item("Hora") = _Hora
            _Busca.Item("Cliente") = _Cliente
            _Busca.Item("Paciente") = _Paciente
            _Busca.Item("Importe") = _Importe
            _Busca.Item("Iva") = 0
            _Busca.Item("Neto") = _Neto
            _Busca.Item("Estatus") = _Estatus
            _Busca.Item("Factura") = _Factura
            _Busca.Item("Documento") = _Documento
            _Busca.Item("Referencia") = _Referencia
            _Busca.Item("Comentarios") = _Comentarios
            _Busca.Item("FechaRegistro") = _FechaRegistro
            _Busca.Item("HoraRegistro") = _HoraRegistro
            _Busca.Item("Atencion") = ""
            _Busca.Item("Usuario") = _Usuario
            _Busca.Item("Centro") = _Centro

            CServiciosDataSet50.Remisiones.Rows.Add(_Busca)

            Me.Validate()
            RemisionesBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet50)
            Me.RemisionesTableAdapter.Fill(Me.CServiciosDataSet50.Remisiones)


            ' Registra el detalle de la remisión
            Dim _Descuento As Double
            _Remision = TextBox8.Text
            For Me.I = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(I).Cells(0).Value = "" Then
                    Exit For
                End If
                _Producto = DataGridView1.Rows(I).Cells(0).Value
                _PrecioUnitario = DataGridView1.Rows(I).Cells(2).Value
                _Cantidad = DataGridView1.Rows(I).Cells(3).Value
                _Descuento = DataGridView1.Rows(I).Cells(5).Value
                _Neto = DataGridView1.Rows(I).Cells(6).Value
                _Importe = _Cantidad * _PrecioUnitario

                _Busca = CServiciosDataSet51.DetalleRemision.NewRow
                _Busca.Item("Remision") = _Remision
                _Busca.Item("Producto") = _Producto
                _Busca.Item("PrecioUnitario") = _PrecioUnitario
                _Busca.Item("Cantidad") = _Cantidad
                _Busca.Item("Importe") = _Importe
                _Busca.Item("Descuento") = _Descuento
                _Busca.Item("ImporteNeto") = _Neto

                CServiciosDataSet51.DetalleRemision.Rows.Add(_Busca)

            Next

            Me.Validate()
            DetalleRemisionBindingSource.EndEdit()
            Me.TableAdapterManager5.UpdateAll(CServiciosDataSet51)
            Me.DetalleRemisionTableAdapter.Fill(Me.CServiciosDataSet51.DetalleRemision)


            ' Incrementa el Consecutivo

            If Mid(Label8.Text, 1, 4) = "Nota" Then
                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("RNP")
            Else
                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("COT")
            End If

            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

            Msg = "Registro guardado correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)


            ' Imprime la remisión
            Button4_Click(sender, e)
            Button1_Click(sender, e)


        Else
            Exit Sub
        End If


    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click

        ' Imprime la remisión

        ' Valida que estén completos los datos
        If DataGridView1.Rows(0).Cells(0).Value = "" Then
            Msg = "Debe agregar al menos un Producto para poder imprimir esta Remisión"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        Dim _SinDatos As Boolean

        _SinDatos = False
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Text = "" Then
                    _SinDatos = True
                End If
            End If
        Next
        If _SinDatos Then
            Msg = "Debe Registrar un Envío para imprimir esta Remisión"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Text = "" Then
                    _SinDatos = True
                End If
            End If
        Next
        If _SinDatos Then
            Msg = "Debe Registrar a un CLIENTE para imprimir esta Remisión"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        For Each Me._Checacontrol In Panel2.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Text = "" Then
                    _SinDatos = True
                End If
            End If
        Next
        If _SinDatos Then
            Msg = "Debe Registrar a un PACIENTE para imprimir esta Remisión"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Formatos Mezclas\Formato Remision.xlsx"

        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False

        _Remision = TextBox8.Text
        _Cliente = TextBox1.Text
        _Paciente = TextBox2.Text
        _Fecha = TextBox9.Text


        If RadioButton2.Checked = True Then
            m_Excel.Worksheets("HORIZONTAL").cells(2, 8).value = _Remision
            m_Excel.Worksheets("HORIZONTAL").cells(8, 3).value = _Cliente
            m_Excel.Worksheets("HORIZONTAL").cells(8, 5).value = _Fecha
            m_Excel.Worksheets("HORIZONTAL").cells(9, 2).value = ComboBox1.Text
            m_Excel.Worksheets("HORIZONTAL").cells(10, 2).value = TextBox3.Text
            m_Excel.Worksheets("HORIZONTAL").cells(11, 2).value = TextBox4.Text
            m_Excel.Worksheets("HORIZONTAL").cells(12, 2).value = TextBox5.Text & ", " & TextBox6.Text
            m_Excel.Worksheets("HORIZONTAL").cells(8, 9).value = _Paciente
            m_Excel.Worksheets("HORIZONTAL").cells(9, 7).value = ComboBox2.Text


            Dim _Linea As Integer
            _Linea = 14

            For Me.I = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(I).Cells(0).Value = "" Then
                    Exit For
                End If
                _PrecioUnitario = DataGridView1.Rows(I).Cells(6).Value / DataGridView1.Rows(I).Cells(3).Value
                m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 2).value = DataGridView1.Rows(I).Cells(0).Value
                m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 3).value = DataGridView1.Rows(I).Cells(1).Value
                m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 7).value = _PrecioUnitario
                m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 9).value = DataGridView1.Rows(I).Cells(3).Value
                m_Excel.Worksheets("HORIZONTAL").cells(_Linea, 10).value = DataGridView1.Rows(I).Cells(6).Value
                _Linea = _Linea + 1
            Next

            m_Excel.Worksheets("HORIZONTAL").cells(32, 3).value = TextBox11.Text
        End If
        ' *************************************************************************************************************
        If RadioButton1.Checked = True Then


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


            ' Borra el formato 
            m_Excel.Worksheets("VERTICAL").cells(3, 6).value = ""
            m_Excel.Worksheets("VERTICAL").cells(7, 6).value = ""
            m_Excel.Worksheets("VERTICAL").cells(10, 2).value = ""
            m_Excel.Worksheets("VERTICAL").cells(11, 2).value = ""
            m_Excel.Worksheets("VERTICAL").cells(12, 2).value = ""
            m_Excel.Worksheets("VERTICAL").cells(13, 2).value = ""
            m_Excel.Worksheets("VERTICAL").cells(14, 2).value = ""

            ' Imprime datos

            If Mid(_Remision, 1, 2) = "RN" Then
                m_Excel.Worksheets("VERTICAL").CELLS(2, 6).value = "NOTA DE REMISIÓN"
            Else
                m_Excel.Worksheets("VERTICAL").CELLS(2, 6).value = "C O T I Z A C I Ó N"
            End If

            m_Excel.Worksheets("VERTICAL").cells(3, 6).value = _Remision
            m_Excel.Worksheets("VERTICAL").cells(7, 6).value = _NombreFecha
            m_Excel.Worksheets("VERTICAL").cells(10, 2).value = ComboBox2.Text
            m_Excel.Worksheets("VERTICAL").cells(11, 2).value = ComboBox1.Text
            m_Excel.Worksheets("VERTICAL").cells(12, 2).value = TextBox3.Text
            m_Excel.Worksheets("VERTICAL").cells(13, 2).value = "COL. " & TextBox4.Text
            m_Excel.Worksheets("VERTICAL").cells(14, 2).value = TextBox5.Text & "," & TextBox6.Text


            Dim _Linea As Integer
            _Linea = 17

            For Me.I = 17 To 44
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 1).value = ""
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 2).value = ""
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 7).value = ""
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 9).value = ""
            Next

            _Linea = 17
            For Me.I = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(I).Cells(0).Value = "" Then
                    Exit For
                End If
                _PrecioUnitario = DataGridView1.Rows(I).Cells(6).Value / DataGridView1.Rows(I).Cells(3).Value
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 1).value = DataGridView1.Rows(I).Cells(3).Value
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 2).value = DataGridView1.Rows(I).Cells(1).Value
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 7).value = _PrecioUnitario
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 9).value = DataGridView1.Rows(I).Cells(6).Value

                _Linea = _Linea + 1
            Next

            If Mid(_Remision, 1, 2) = "RN" Then
            Else
                m_Excel.Worksheets("VERTICAL").cells(_Linea, 2).value = "************************* FIN DE LA COTIZACIÓN *************************"
            End If

            m_Excel.Worksheets("VERTICAL").cells(45, 4).value = TextBox11.Text

        End If

        Dim _NumCopias As Integer = Int(TextBox21.Text)

        If CheckBox1.Checked = True Then
            m_Excel.Visible = True
        Else
            If RadioButton2.Checked = True Then
                m_Excel.Worksheets("HORIZONTAL").Printout(1, 1, _NumCopias, False)
            Else
                m_Excel.Worksheets("VERTICAL").Printout(1, 1, _NumCopias, False)
            End If

            'm_Excel.ActiveWorkbook.AutoUpdateSaveChanges = False
            m_Excel.ActiveWorkbook.Close(False)

        End If

    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox13.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub TextBox13_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox13.LostFocus
        If TextBox13.Text = "" Then
            ComboBox3.Focus()
        Else
            _Producto = TextBox13.Text
            Dim _Rows_Pr() As DataRow = CServiciosDataSet21.Productos.Select("Producto = " & "'" & _Producto & "'")
            Dim _Find As Boolean
            _Find = False

            For Me.I = 0 To _Rows_Pr.GetUpperBound(0)
                _Find = True
                ComboBox3.Text = _Rows_Pr(I).Item("Descripcion")
                Exit For
            Next

            If _Find = False Then
                Msg = "No existe ese Producto registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                TextBox13.Focus()
            End If

        End If
    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()

    End Sub

    Private Sub GenerarNuevaRemisiónToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GenerarNuevaRemisiónToolStripMenuItem.Click
        ToolStripButton2_Click(sender, e)
    End Sub

    Private Sub GuardarInformaciónToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GuardarInformaciónToolStripMenuItem.Click
        ToolStripButton3_Click(sender, e)
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox1.LostFocus
        If TextBox1.Text = "" Then
            ComboBox1.Focus()
        Else
            ' Busca ese Cliente
            _Cliente = UCase(TextBox1.Text)

            Dim _Rows_Cl() As DataRow
            _Rows_Cl = Me.CServiciosDataSet30.Clientes.Select("Cliente = " & "'" & _Cliente & "'")
            Dim _Find As Boolean

            _Find = False
            For Me.I = 0 To _Rows_Cl.GetUpperBound(0)
                ComboBox1.Text = _Rows_Cl(I).Item("NombreCliente")
                _Find = True
                ComboBox2.Focus()
                Exit For
            Next


            If _Find = False Then
                _Rows_Cl = Me.CServiciosDataSet30.Clientes.Select("Referencia = " & "'" & _Cliente & "'")

                For Me.I = 0 To _Rows_Cl.GetUpperBound(0)
                    ComboBox1.Text = _Rows_Cl(I).Item("NombreCliente")
                    _Find = True
                    ComboBox2.Focus()
                    Exit For
                Next
            End If


            If _Find = False Then
                Msg = "Ese Cliente no está registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                TextBox1.Focus()
                Exit Sub
            End If


        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox14_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox14.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox15.Focus()
        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub UtilizarFechaPredertinadaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UtilizarFechaPredertinadaToolStripMenuItem.Click
        MonthCalendar1.Visible = True
        MonthCalendar1.Enabled = True
        Panel1.Enabled = True
    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox9.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False
        _Fecha = CStr(TextBox9.Text)


        ' Guarda la fecha predeterminada

        _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
        _Busca.Item("Usuario") = _Fecha

        Me.Validate()
        CentrosBindingSource.EndEdit()
        Me.TableAdapterManager6.UpdateAll(CServiciosDataSet)
        Me.CentrosTableAdapter.Fill(CServiciosDataSet.Centros)

        Msg = "Fecha predeterminada guardada correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)
        Button1_Click(sender, e)


    End Sub

    Private Sub EliminarFechaPredeterminadaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EliminarFechaPredeterminadaToolStripMenuItem.Click
        ' Elimina la fecha predeterminada
        _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
        _Busca.Item("Usuario") = ""

        Me.Validate()
        CentrosBindingSource.EndEdit()
        Me.TableAdapterManager6.UpdateAll(CServiciosDataSet)
        Me.CentrosTableAdapter.Fill(CServiciosDataSet.Centros)

        Msg = "Fecha predeterminada eliminada correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        Button1_Click(sender, e)
    End Sub

    Private Sub RepetirDatosToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RepetirDatosToolStripMenuItem.Click
        ' Repetir Información
        If TextBox8.Text = "" Then
            Msg = "Debe seleccionar primero la opción de GENERAR UNA NUEVA REMISION"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        Dim _NombrePaciente As String
        ComboBox4.Visible = True
        ComboBox4.Items.Clear()
        ComboBox4.Text = "Selecciona la Remisión para Repetir Información en la Actual"

        For Me.I = 0 To CServiciosDataSet50.Remisiones.Rows.Count - 1
            _Remision = CServiciosDataSet50.Remisiones.Rows(I).Item("Remision")
            _Fecha = CServiciosDataSet50.Remisiones.Rows(I).Item("Fecha")
            _Paciente = CServiciosDataSet50.Remisiones.Rows(I).Item("Paciente")

            _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
            _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")

            _Referencia = _Remision & "    " & CStr(_Fecha) & "    " & _NombrePaciente

            ComboBox4.Items.Add(_Referencia)
        Next
        ComboBox4.Focus()


    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        _Remision = Mid(ComboBox4.Text, 1, 6)

        _Busca = CServiciosDataSet50.Remisiones.FindByRemision(_Remision)
        _Cliente = _Busca.Item("Cliente")
        _Paciente = _Busca.Item("Paciente")

        _Busca = Me.CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
        TextBox1.Text = _Cliente
        TextBox3.Text = _Busca.Item("Calle")
        TextBox4.Text = "COL. " & _Busca.Item("Colonia")
        TextBox5.Text = _Busca.Item("Ciudad") & "," & _Busca.Item("Estado")
        ComboBox1.Text = _Busca.Item("NombreCliente")

        _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
        ComboBox2.Text = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")
        TextBox2.Text = _Busca.Item("Paciente")

        ' Presenta el detalle de la remisión
        Dim _Rows_Dr() As DataRow = CServiciosDataSet51.DetalleRemision.Select("Remision = " & "'" & _Remision & "'")
        Dim _Importe, _Descuento, _ImporteNeto As Double
        Dim _Cuantos As Integer
        Dim _Suma As Double

        _Cuantos = 0
        _Suma = 0
        DataGridView1.Rows.Clear()
        For Me.I = 0 To _Rows_Dr.GetUpperBound(0)
            _Producto = _Rows_Dr(I).Item("Producto")
            _Busca = Me.CServiciosDataSet21.Productos.FindByProducto(_Producto)
            _Referencia = _Busca.Item("Descripcion")
            _PrecioUnitario = _Rows_Dr(I).Item("PrecioUnitario")
            _Cantidad = _Rows_Dr(I).Item("Cantidad")
            _Importe = _Rows_Dr(I).Item("Importe")
            _Descuento = _Rows_Dr(I).Item("Descuento")
            _ImporteNeto = _Rows_Dr(I).Item("ImporteNeto")

            _Cuantos = _Cuantos + 1
            _Suma = _Suma + _ImporteNeto

            DataGridView1.Rows.Add(_Producto, _Referencia, _PrecioUnitario, _Cantidad, _Importe, _Descuento, _ImporteNeto)
        Next

        TextBox11.Text = _Cuantos
        TextBox12.Text = Format(_Suma, "$###,###,##0.00")
        TextBox20.Text = _Suma

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        ToolStripTextBox1.Text = " " & ToolStripTextBox1.Text
    End Sub

    Private Sub TextBox21_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox21.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox21_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox21.TextChanged

    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        ' Agrega una nueva Cotización
        For Each Me._Checacontrol In Me.Controls
            _Checacontrol.Enabled = True
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""

        TextBox19.Text = "1"

        ' Obtiene el siguiente Documento
        ' Revisa si existe fecha predeterminada
        _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)

        Dim _FechaPredeterminada As String

        If DBNull.Value.Equals(_Busca.Item("Usuario")) Then
            _FechaPredeterminada = ""
        Else
            _FechaPredeterminada = _Busca.Item("Usuario")
        End If

        ' Dim _FechaPredeterminada As String = _Busca.Item("Usuario")

        If _FechaPredeterminada <> "" Then
            Msg = "Existe fecha predeterminadá. Esta se utilizará en todas las Remisiones que registre"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            TextBox9.Text = CDate(_FechaPredeterminada)
        Else
            TextBox9.Text = Date.Today
        End If

        Label8.Text = "Número de Cotización"
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
        TextBox21.Text = "1"
        TextBox8.Focus()
    End Sub
End Class