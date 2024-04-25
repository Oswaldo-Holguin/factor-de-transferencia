Public Class Frm6TablaMateriales
    Public _Material As String
    Public _Descripcion As String
    Public i As Integer
    Public checacontrol As Control
    Public _Busca As DataRow
    Public _Fila As DataRow
    Public _Referencia As String
    Public _Comentarios As String
    Public _Presentacion As String
    Public _renglon As DataGridViewRow
    Public _Contenido As String

    Private Sub Frm6TablaMateriales_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Escape) Then
            Button1_Click(sender, e)
        End If
    End Sub

    Private Sub Frm6TablaMateriales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
     

        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: This line of code loads data into the 'CServiciosDataSet5.TipoMaterial' table. You can move, or remove it, as needed.
        Me.TipoMaterialTableAdapter1.Fill(Me.CServiciosDataSet51.TipoMaterial)
        'TODO: This line of code loads data into the 'CServiciosDataSet5.Materiales' table. You can move, or remove it, as needed.
        Me.MaterialesTableAdapter1.Fill(Me.CServiciosDataSet51.Materiales)
        'TODO: This line of code loads data into the 'CServiciosDataSet.Centros' table. You can move, or remove it, as needed.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: This line of code loads data into the 'CServiciosDataSet3.Documentos' table. You can move, or remove it, as needed.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        Panel1.Enabled = False
        RadioButton5.Checked = True
        Dim _DescripcionTipo As String
        Dim _Tipo As String
        Dim _Existencia As Double
        DataGridView1.Rows.Clear()
        For Me.i = 0 To CServiciosDataSet51.Materiales.Rows.Count - 1

            _Material = CServiciosDataSet51.Materiales.Rows(i).Item("Material")
            _Descripcion = CServiciosDataSet51.Materiales.Rows(i).Item("Descripcion")
            _Tipo = CServiciosDataSet51.Materiales.Rows(i).Item("TipoMaterial")
            _Busca = CServiciosDataSet51.TipoMaterial.FindByTipoMaterial(_Tipo)
            _DescripcionTipo = _Busca.Item("Descripcion")
            _Existencia = CServiciosDataSet51.Materiales.Rows(i).Item("Existencia")
            DataGridView1.Rows.Add(_Material, _Descripcion, _DescripcionTipo, _Existencia)

        Next

        ' Llena el combo con los datos del Tipo de Material
        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet51.TipoMaterial.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet51.TipoMaterial.Rows(i).Item("Descripcion"))
        Next
        ComboBox2.Items.Clear()
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet.Centros.Rows(i).Item("Descripcion"))
        Next

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        ToolStripButton2.Enabled = True
        Timer1.Stop()


    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        DataGridView1.Rows.Clear()
        Panel1.Enabled = True
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
        Next
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
        ToolStripButton4.Enabled = False
        ToolStripButton5.Enabled = False
        TextBox1.Focus()
        TextBox13.Text = "ALTA"
        Timer1.Start()
        Timer1.Interval = 100
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        Dim _Consecutivo As Integer
        For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "MAT" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo")
            End If
        Next
        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "MA" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.Text = UCase(TextBox1.Text)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.Text = UCase(TextBox2.Text)
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet51.TipoMaterial.Rows.Count - 1
            If CServiciosDataSet51.TipoMaterial.Rows(i).Item("Descripcion") = ComboBox1.Text Then
                TextBox3.Text = CServiciosDataSet51.TipoMaterial.Rows(i).Item("TipoMaterial")
                TextBox4.Focus()
                Exit For
            End If
        Next
    End Sub

    Private Sub TextBox4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.Text = UCase(TextBox4.Text)
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.Text = UCase(TextBox5.Text)
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

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
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.LostFocus
        TextBox6.Text = UCase(TextBox6.Text)
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub TextBox12_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox12.LostFocus
        TextBox12.Text = UCase(TextBox12.Text)
        TextBox12.Text = Mid(TextBox12.Text, 1, 6)
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        ' Da de Alta un nuevo Material

        Dim _vacio As Boolean
        _vacio = False
        For Each checontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                If checacontrol.Text = "" Then
                    _vacio = True
                End If
            End If
        Next

        If _vacio Then
            MsgBox("Hay casillas en blanco. Debe llenar todos los datos")
            Exit Sub
        End If

        If TextBox13.Text = "ALTA" Then
            _Fila = CServiciosDataSet51.Materiales.NewRow
            _Fila.Item("Material") = TextBox1.Text
        End If

        If TextBox13.Text = "CAMBIO" Then
            _Fila = CServiciosDataSet51.Materiales.FindByMaterial(TextBox1.Text)
        End If

        _Fila.Item("Descripcion") = TextBox2.Text
        _Fila.Item("Referencia") = TextBox4.Text
        _Fila.Item("TipoMaterial") = TextBox3.Text
        _Fila.Item("Comentarios") = UCase(RichTextBox1.Text)
        _Fila.Item("FechaRegistro") = CStr(Today)
        _Fila.Item("HoraRegistro") = TimeString
        _Fila.Item("Maximo") = CDbl(TextBox7.Text)
        _Fila.Item("Minimo") = CDbl(TextBox8.Text)
        _Fila.Item("PuntoReorden") = TextBox9.Text
        _Fila.Item("PrecioUnitario") = TextBox10.Text
        _Fila.Item("UMedida") = TextBox12.Text
        _Fila.Item("Presentacion") = TextBox5.Text
        _Fila.Item("Contenido") = TextBox6.Text
        _Fila.Item("Centro") = TextBox14.Text
        _Fila.Item("Existencia") = TextBox11.Text

        If TextBox13.Text = "ALTA" Then
            CServiciosDataSet51.Materiales.Rows.Add(_Fila)
        End If

        Me.Validate()
        MaterialesBindingSource1.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet51)
        Me.MaterialesTableAdapter1.Fill(Me.CServiciosDataSet51.Materiales)

        ' Incremente al consecutivo
        If TextBox13.Text = "ALTA" Then
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("MAT")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager3.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(CServiciosDataSet3.Documentos)
        End If

        MsgBox("Registro guardado correctamente")

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""
        Panel1.Enabled = False
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        TextBox1.Enabled = True
        Button1_Click(sender, e)


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            If CServiciosDataSet.Centros.Rows(i).Item("Descripcion") = ComboBox2.Text Then
                TextBox14.Text = CServiciosDataSet.Centros.Rows(i).Item("Centro")
                TextBox2.Focus()
                Exit For
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _renglon = DataGridView1.CurrentRow
        If _renglon.Cells(0).Value <> "" Then

            _renglon.DefaultCellStyle.BackColor = Color.Gold
            TextBox1.Text = _renglon.Cells(0).Value
            TextBox2.Text = _renglon.Cells(1).Value
            _Material = TextBox1.Text
            _Busca = CServiciosDataSet51.Materiales.FindByMaterial(_Material)
            TextBox14.Text = _Busca.Item("Centro")
            TextBox3.Text = _Busca.Item("TipoMaterial")
            TextBox4.Text = _Busca.Item("Referencia")
            TextBox5.Text = _Busca.Item("Presentacion")
            TextBox6.Text = _Busca.Item("Contenido")
            TextBox7.Text = _Busca.Item("Maximo")
            TextBox8.Text = _Busca.Item("Minimo")
            TextBox9.Text = _Busca.Item("PuntoReorden")
            TextBox10.Text = _Busca.Item("PrecioUnitario")
            TextBox11.Text = _Busca.Item("Existencia")
            TextBox12.Text = _Busca.Item("Umedida")
            RichTextBox1.Text = _Busca.Item("Comentarios")

            _Busca = CServiciosDataSet.Centros.FindByCentro(TextBox14.Text)
            ComboBox2.Text = _Busca.Item("Descripcion")

            _Busca = CServiciosDataSet51.TipoMaterial.FindByTipoMaterial(TextBox3.Text)
            ComboBox1.Text = _Busca.Item("Descripcion")

            ToolStripButton4.Enabled = True
            ToolStripButton4.Enabled = True
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub DataGridView1_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _renglon = DataGridView1.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Panel1.Enabled = True
        TextBox1.Enabled = False
        ComboBox2.Focus()
        ToolStripButton4.Enabled = False
        ToolStripButton3.Enabled = True
        TextBox13.Text = "CAMBIO"
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ' Busqueda
        DataGridView1.Rows.Clear()

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
            RichTextBox1.Text = ""
        Next

        Dim _Rows_Ma() As DataRow
        Dim _Mide, X, R As Integer
        Dim _NombreMaterial, _ParteNombre, _DescripcionTipo, _Tipo As String
        Dim _Existencia As Double
        Dim _TipoMaterial As String

        If RadioButton1.Checked = True Then
            _Material = InputBox("Teclee el código del Material a buscar ")
            _Material = UCase(_Material)
            _Rows_Ma = CServiciosDataSet51.Materiales.Select("Material = " & "'" & _Material & "'")

            For Me.i = 0 To _Rows_Ma.GetUpperBound(0)
                _NombreMaterial = _Rows_Ma(i).Item("Descripcion")
                _Tipo = _Rows_Ma(i).Item("TipoMaterial")
                _Busca = CServiciosDataSet51.TipoMaterial.FindByTipoMaterial(_Tipo)
                _DescripcionTipo = _Busca.Item("Descripcion")
                _Existencia = _Rows_Ma(i).Item("Existencia")


                DataGridView1.Rows.Add(_Material, _NombreMaterial, _DescripcionTipo, _Existencia)
            Next
        End If

        If RadioButton2.Checked = True Then
            _Descripcion = InputBox("Teclee la Descripción a buscar ")
            _Descripcion = UCase(_Descripcion)
            _Mide = Len(_Descripcion)

            For Me.i = 0 To CServiciosDataSet51.Materiales.Rows.Count - 1
                _NombreMaterial = CServiciosDataSet51.Materiales.Rows(i).Item("Descripcion")
                _Tipo = CServiciosDataSet51.Materiales.Rows(i).Item("TipoMaterial")
                _Busca = CServiciosDataSet51.TipoMaterial.FindByTipoMaterial(_Tipo)
                _DescripcionTipo = _Busca.Item("Descripcion")
                _Existencia = CServiciosDataSet51.Materiales.Rows(i).Item("Existencia")
                _Material = CServiciosDataSet51.Materiales.Rows(i).Item("Material")

                For X = 1 To Len(_NombreMaterial)
                    _ParteNombre = Mid(_NombreMaterial, X, _Mide)
                    If _ParteNombre = _Descripcion Then
                        DataGridView1.Rows.Add(_Material, _NombreMaterial, _DescripcionTipo, _Existencia)
                    End If
                Next

            Next

        End If


        If RadioButton3.Checked = True Then
            _Tipo = InputBox("Teclee el Tipo de Material a buscar")
            _Tipo = UCase(_Tipo)
            _Mide = Len(_Tipo)

            For Me.i = 0 To CServiciosDataSet51.TipoMaterial.Rows.Count - 1
                _DescripcionTipo = CServiciosDataSet51.TipoMaterial.Rows(i).Item("Descripcion")

                For X = 1 To Len(_DescripcionTipo)
                    _ParteNombre = Mid(_DescripcionTipo, X, _Mide)

                    If _ParteNombre = _Tipo Then

                        _TipoMaterial = CServiciosDataSet51.TipoMaterial.Rows(i).Item("TipoMaterial")
                        _Rows_Ma = CServiciosDataSet51.Materiales.Select("TipoMaterial = " & "'" & _TipoMaterial & "'")

                        For R = 0 To _Rows_Ma.GetUpperBound(0)
                            _Material = _Rows_Ma(R).Item("Material")
                            _NombreMaterial = _Rows_Ma(R).Item("Descripcion")
                            _Existencia = _Rows_Ma(R).Item("Existencia")

                            DataGridView1.Rows.Add(_Material, _NombreMaterial, _DescripcionTipo, _Existencia)
                        Next
                    End If
                Next
            Next

        End If



    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                If checacontrol.Focused Then
                    checacontrol.BackColor = Color.Gold
                Else
                    checacontrol.BackColor = Color.White
                End If
            End If
        Next
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        If Asc(e.KeyChar) = 13 Then
            TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        If Asc(e.KeyChar) = 13 Then
            TextBox9.Focus()
        End If
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        If Asc(e.KeyChar) = 13 Then
            TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        If Asc(e.KeyChar) = 13 Then
            TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
        If Asc(e.KeyChar) = 13 Then
            TextBox12.Focus()
        End If
    End Sub
End Class