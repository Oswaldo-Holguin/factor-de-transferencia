Public Class Frm27TablaProductos
    Public i As Integer
    Public _Producto, _Descripcion, _Referencia, _Presentacion, _Contenido, _Formula, _Centro As String
    Public _Busca As DataRow
    Public msg As String
    Public style As MsgBoxStyle
    Public checacontrol As Control


    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus

        If CheckBox1.Checked = True Then
            ' Busca el consecutivo para los Productos
            Dim Busca As DataRow
            Dim _Consecutivo As Integer
            Dim _Xconsecutivo As String
            Busca = CServiciosDataSet3.Documentos.FindByDocumento("PRD")
            _Consecutivo = Busca.Item("Consecutivo")
            _Xconsecutivo = "000" & Trim(CStr(_Consecutivo))

            Dim _Mide As Integer = Len(_Xconsecutivo)
            Dim _x As Integer = _Mide - 3
            _Xconsecutivo = "PR" + Mid(_Xconsecutivo, _x, 4)
            TextBox1.Text = _Xconsecutivo
            TextBox1.Enabled = False
            TextBox2.Focus()
        Else

        End If

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Frm27TablaProductos_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Escape) Then
            Button1_Click(sender, e)
        End If
    End Sub

    Private Sub Form27_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    
        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet21.Productos' Puede moverla o quitarla según sea necesario.
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet21.Formulas' Puede moverla o quitarla según sea necesario.
        Me.FormulasTableAdapter.Fill(Me.CServiciosDataSet21.Formulas)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)


        _Centro = LoginForm1.TextBox1.Text
        DataGridView1.Rows.Clear()

        Dim _Rows_Pr() As DataRow = CServiciosDataSet21.Productos.Select("Centro = " & "'" & _Centro & "'")
        For Me.i = 0 To _Rows_Pr.GetUpperBound(0)
            _Producto = _Rows_Pr(i).Item("Producto")
            _Descripcion = _Rows_Pr(i).Item("Descripcion")

            DataGridView1.Rows.Add(_Producto, _Descripcion)
        Next
        Panel1.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        ToolStripButton2.Enabled = True

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
            checacontrol.Enabled = True
        Next
        RichTextBox1.Text = ""
        Timer1.Enabled = False


    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
        ToolStripButton4.Enabled = False

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""

        TextBox9.Text = "ALTA"
        TextBox1.Focus()
        Timer1.Enabled = True
        Timer1.Interval = 250

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.Text = UCase(TextBox2.Text)
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = "SIN REFERENCIA"
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.Text = UCase(TextBox3.Text)
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.Text = UCase(TextBox4.Text)
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.Text = UCase(TextBox5.Text)
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub RichTextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles RichTextBox1.GotFocus
        RichTextBox1.Text = "SIN COMENTARIOS"
    End Sub

    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ComboBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.GotFocus
        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet21.Formulas.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet21.Formulas.Rows(i).Item("Descripcion"))
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet21.Formulas.Rows.Count - 1
            If CServiciosDataSet21.Formulas.Rows(i).Item("Descripcion") = ComboBox1.Text Then
                TextBox7.Text = CServiciosDataSet21.Formulas.Rows(i).Item("Formula")
            End If
        Next
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub ComboBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.GotFocus
        ComboBox2.Items.Clear()
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet.Centros.Rows(i).Item("Descripcion"))
        Next
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            If CServiciosDataSet.Centros.Rows(i).Item("Descripcion") = ComboBox2.Text Then
                TextBox8.Text = CServiciosDataSet.Centros.Rows(i).Item("Centro")
            End If
        Next
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Panel1.Enabled = True
        TextBox9.Text = "CAMBIO"
        ToolStripButton4.Enabled = False
        ToolStripButton3.Enabled = True
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        ' Registra el producto
        Timer1.Enabled = False
        Dim _Blancos As Boolean
        _Blancos = False
        ' Verifica que no haya casillas en blanco
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                If checacontrol.Text = "" Then
                    _blancos = True
                End If
            End If
            If TypeOf checacontrol Is ComboBox Then
                If checacontrol.Text = "" Then
                    _Blancos = True
                End If
            End If
        Next

        If _Blancos Then
            msg = "Hay casillas en blanco. Debe llenar todos los datos"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If


        ' Verifica que no exista ese producto anteriormente
        _Producto = TextBox1.Text

        If TextBox9.Text <> "CAMBIO" Then
            Dim _Rows_Pr() As DataRow = CServiciosDataSet21.Productos.Select("Producto = " & "'" & _Producto & "'")
            Dim _Find As Boolean

            _Find = False

            For Me.i = 0 To _Rows_Pr.GetUpperBound(0)
                _Find = True
                _Descripcion = _Rows_Pr(i).Item("Descripcion")
            Next

            If _Find Then
                msg = "Ese producto ya existe con la descripcion " & _Descripcion
                style = MsgBoxStyle.Information
                MsgBox(msg, style)
                Exit Sub
            End If
        End If


        If TextBox9.Text = "ALTA" Then
            _Busca = CServiciosDataSet21.Productos.NewRow()
            _Busca.Item("Producto") = TextBox1.Text
        End If
        If TextBox9.Text = "CAMBIO" Then
            _Busca = CServiciosDataSet21.Productos.FindByProducto(TextBox1.Text)

        End If

        Dim _Usuario As String = LoginForm1.TextBox6.Text

        _Busca.Item("Descripcion") = TextBox2.Text
        _Busca.Item("Centro") = TextBox8.Text
        _Busca.Item("Referencia") = TextBox3.Text
        _Busca.Item("Comentarios") = UCase(RichTextBox1.Text)
        _Busca.Item("Usuario") = _Usuario
        _Busca.Item("Presentacion") = TextBox4.Text
        _Busca.Item("Contenido") = TextBox5.Text
        _Busca.Item("Existencia") = 0
        _Busca.Item("PrecioUnitario") = CDbl(TextBox6.Text)
        _Busca.Item("Formula") = TextBox7.Text

        If TextBox9.Text = "ALTA" Then
            CServiciosDataSet21.Productos.Rows.Add(_Busca)
        End If

        Me.Validate()
        ProductosBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet21)
        Me.ProductosTableAdapter.Fill(Me.CServiciosDataSet21.Productos)

        'Actualiza el Consecutivo

        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("PRD")
        _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet3)
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        msg = "Registro guardado Correctamente"
        style = MsgBoxStyle.Information
        MsgBox(msg, style)
        Button1_Click(sender, e)

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim _Renglon As DataGridViewRow = DataGridView1.CurrentRow
        Dim _BuscaF() As DataRow


        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(0).Value
        TextBox2.Text = _Renglon.Cells(1).Value
        _Busca = CServiciosDataSet21.Productos.FindByProducto(TextBox1.Text)
        If DBNull.Value.Equals(_Busca.Item("Referencia")) Then
            TextBox3.Text = ""
        Else
            TextBox3.Text = _Busca.Item("Referencia")
        End If
        If DBNull.Value.Equals(_Busca.Item("Comentarios")) Then
            RichTextBox1.Text = ""
        Else
            RichTextBox1.Text = _Busca.Item("Comentarios")
        End If
        If DBNull.Value.Equals(_Busca.Item("Presentacion")) Then
            TextBox4.Text = ""
        Else
            TextBox4.Text = _Busca.Item("Presentacion")
        End If
        If DBNull.Value.Equals(_Busca.Item("Contenido")) Then
            TextBox5.Text = ""
        Else
            TextBox5.Text = _Busca.Item("Contenido")
        End If
        If DBNull.Value.Equals(_Busca.Item("PrecioUnitario")) Then
            TextBox6.Text = 0
        Else
            TextBox10.Text = _Busca.Item("PrecioUnitario")
            TextBox6.Text = Format(_Busca.Item("PrecioUnitario"), "$##,###,##0.00")
        End If
        If DBNull.Value.Equals(_Busca.Item("Formula")) Then
            TextBox7.Text = ""
        Else
            TextBox7.Text = _Busca.Item("Formula")
            _Formula = TextBox7.Text
            _BuscaF = Me.CServiciosDataSet21.Formulas.Select("Formula = " & "'" & _Formula & "'")
            For x = 0 To _BuscaF.GetUpperBound(0)
                ComboBox1.Text = _BuscaF(x).Item("Descripcion")
            Next

        End If

        If DBNull.Value.Equals(_Busca.Item("Centro")) Then
            TextBox8.Text = ""
        Else
            TextBox8.Text = _Busca.Item("Centro")
            _Centro = TextBox8.Text
            _Busca = Me.CServiciosDataSet.Centros.FindByCentro(_Centro)
            ComboBox2.Text = _Busca.Item("Descripcion")
        End If
        ToolStripButton4.Enabled = True

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        Dim _Renglon As DataGridViewRow = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        '     If e.KeyChar.IsDigit(e.KeyChar) Then
        ' e.Handled = False
        ' ElseIf e.KeyChar.IsControl(e.KeyChar) Then
        ' e.Handled = False
        ' Else
        ' e.Handled = True
        ' End If
    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox6.LostFocus

        If Not IsNumeric(TextBox6.Text) Then
            msg = "Debe teclear un Precio Unitario Válido"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        TextBox10.Text = TextBox6.Text
        Dim _PrecioUnitario As Double = CDbl(TextBox6.Text)
        '  TextBox6.Text = Format(_PrecioUnitario, "$##,###,##0.00")

        _Centro = LoginForm1.TextBox1.Text

        If _Centro = "CME" Then
            TextBox7.Text = "NP"
            ComboBox1.Text = "NUTRICIONES PARENTERALES"
        End If

        TextBox8.Text = _Centro
        _Busca = Me.CServiciosDataSet.Centros.FindByCentro(_Centro)
        ComboBox2.Text = _Busca.Item("Descripcion")


    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                If checacontrol.Focused Then
                    checacontrol.BackColor = Color.PeachPuff
                End If
            End If
            If TypeOf checacontrol Is ComboBox Then
                If checacontrol.Focused Then
                    checacontrol.BackColor = Color.PeachPuff
                End If
            End If
            If TypeOf checacontrol Is RichTextBox Then
                checacontrol.BackColor = Color.PeachPuff
            End If
        Next

    End Sub
End Class