Public Class Frm5RecetaPaciente
    Public checacontrol As Control
    Public i As Integer
    Public _Busca As DataRow
    Public _renglon As DataGridViewRow
    Public _Fila As DataRow

    Private Sub Ftm5RecetaPaciente_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    
        ToolStripButton4.Enabled = False
        Button1_Click(sender, e)
        DataGridView1.Enabled = False



    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        Dim _Consecutivo As Integer
        For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "REC" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo")
            End If
        Next
        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "RE" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

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

    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True
        TextBox1.Focus()
        TextBox6.Text = "ALTA"
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
        DataGridView1.Enabled = True
    End Sub

    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox2.Text = CStr(Today)
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = TimeString
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guarda la Receta
        Dim _vacio As Boolean
        _vacio = False
        For Each Me.checacontrol In Panel1.Controls
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

        If DataGridView2.Rows(0).Cells(0).Value = "" Then
            MsgBox("Debe asignar Medicamentos a esta Receta")
            Exit Sub
        End If

        If TextBox6.Text = "ALTA" Then
            _Fila = CServiciosDataSet2.Recetas.NewRow
            _Fila.Item("NumReceta") = TextBox1.Text
        End If

        If TextBox6.Text = "CAMBIO" Then
            _Fila = CServiciosDataSet2.Recetas.FindByNumReceta(TextBox1.Text)
        End If

        _Fila.Item("Consulta") = ToolStripTextBox3.Text
        _Fila.Item("Paciente") = Frm3Pacientes.TextBox1.Text
        _Fila.Item("FechaReceta") = TextBox2.Text
        _Fila.Item("HoraReceta") = TextBox3.Text
        _Fila.Item("Centro") = Frm3Pacientes.TextBox5.Text
        _Fila.Item("Atiende") = TextBox5.Text
        _Fila.Item("Referencia") = TextBox4.Text
        _Fila.Item("Comentarios") = UCase(RichTextBox1.Text)
        _Fila.Item("FechaRegistro") = CStr(Today)
        _Fila.Item("HoraRegistro") = TimeString
        _Fila.Item("Usuario") = LoginForm1.TextBox6.Text

        If TextBox6.Text = "ALTA" Then
            CServiciosDataSet2.Recetas.Rows.Add(_Fila)
        End If

        Me.Validate()
        RecetasBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet2)
        Me.RecetasTableAdapter.Fill(CServiciosDataSet2.Recetas)

        If TextBox6.Text = "CAMBIO" Then
            Dim _Seleccion = "NumReceta = " & "'" & TextBox1.Text & "'"
            Dim _Rows_dr() As DataRow = CServiciosDataSet2.DetalleReceta.Select(_Seleccion)

            For Me.i = 0 To _Rows_dr.GetUpperBound(0)
                _Rows_dr(i).Delete()
            Next
            Me.Validate()
            DetalleRecetaBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet2)
            Me.DetalleRecetaTableAdapter.Fill(CServiciosDataSet2.DetalleReceta)
        End If

        ' Registra el detalle de la Receta
        For Me.i = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(i).Cells(0).Value <> "" Then
                _Fila = CServiciosDataSet2.DetalleReceta.NewRow
                _Fila.Item("NumReceta") = TextBox1.Text
                _Fila.Item("Material") = DataGridView2.Rows(i).Cells(0).Value
                _Fila.Item("Cantidad") = DataGridView2.Rows(i).Cells(3).Value
                _Fila.Item("Referencia") = DataGridView2.Rows(i).Cells(6).Value
                _Fila.Item("Comentarios") = ""
                _Fila.Item("Usuario") = LoginForm1.TextBox6.Text
                _Fila.Item("FechaRegistro") = CStr(Today)
                _Fila.Item("HoraRegistro") = TimeString
                _Fila.Item("Paciente") = Frm3Pacientes.TextBox1.Text
                _Fila.Item("Consulta") = ToolStripTextBox3.Text
                _Fila.Item("PrecioUnitario") = DataGridView2.Rows(i).Cells(2).Value

                CServiciosDataSet2.DetalleReceta.Rows.Add(_Fila)
            End If
        Next
        Me.Validate()
        DetalleRecetaBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet2)
        Me.DetalleRecetaTableAdapter.Fill(CServiciosDataSet2.DetalleReceta)

        ' Actualiza el registro de la Consulta con el número de Receta
        _Busca = CServiciosDataSet2.Consultas.FindByConsulta(ToolStripTextBox3.Text)
        _Busca.Item("NumReceta") = TextBox1.Text
        Me.Validate()
        ConsultasBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet2)
        Me.ConsultasTableAdapter.Fill(CServiciosDataSet2.Consultas)


        ' Incrementa el consecutivo de las Recetas
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("REC")
        _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet3)
        Me.DocumentosTableAdapter.Fill(CServiciosDataSet3.Documentos)


        MsgBox("Registro guardado Correctamente")
        Me.Close()

    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim _Cantidad As Integer
        Dim _Referencia As String
        Dim _Material, _Descripcion, _Contenido, _Presentacion As String
        Dim _PrecioUnitario As Double

        _renglon = DataGridView1.CurrentRow

        If _renglon.Cells(0).Value <> "" Then
            _renglon.DefaultCellStyle.BackColor = Color.Gold
            _Cantidad = InputBox("Teclee la Cantidad de Medicamento Recetado")
            _Referencia = InputBox("Teclee la Dosis del Medicamento ")
            _Referencia = UCase(_Referencia)
            _Material = _renglon.Cells(0).Value
            _Descripcion = _renglon.Cells(1).Value
            _PrecioUnitario = _renglon.Cells(2).Value
            _Presentacion = _renglon.Cells(3).Value
            _Contenido = _renglon.Cells(4).Value

            DataGridView2.Rows.Add(_Material, _Descripcion, _PrecioUnitario, _Cantidad, _Presentacion, _Contenido, _Referencia)

        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Consultas' table. You can move, or remove it, as needed.
        Me.ConsultasTableAdapter.Fill(Me.CServiciosDataSet2.Consultas)
        'TODO: This line of code loads data into the 'CServiciosDataSet5.Materiales' table. You can move, or remove it, as needed.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.DetalleReceta' table. You can move, or remove it, as needed.
        Me.DetalleRecetaTableAdapter.Fill(Me.CServiciosDataSet2.DetalleReceta)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Recetas' table. You can move, or remove it, as needed.
        Me.RecetasTableAdapter.Fill(Me.CServiciosDataSet2.Recetas)
        'TODO: This line of code loads data into the 'CServiciosDataSet3.Documentos' table. You can move, or remove it, as needed.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        Dim _Nombre As String = Frm3Pacientes.TextBox2.Text + " " + Frm3Pacientes.TextBox3.Text + " " + Frm3Pacientes.TextBox4.Text

        ToolStripTextBox1.Text = _Nombre
        ToolStripTextBox2.Text = Frm3Pacientes.ComboBox1.Text
        ToolStripTextBox3.Text = Frm3Pacientes.DataGridView3.CurrentRow.Cells(2).Value
        Panel1.Enabled = False
        Dim _Material, _Descripcion, _Presentacion, _Contenido As String
        Dim _PrecioUnitario As Double
        Dim _Seleccion As String = "TipoMaterial = 'MED'"
        Dim _Receta As String = Frm3Pacientes.DataGridView3.CurrentRow.Cells(8).Value
        Dim _Selecciona As String = "NumReceta = " & "'" & _Receta & "'"
        Dim _Rows_mat() As DataRow = CServiciosDataSet5.Materiales.Select(_Seleccion)
        Dim _Rows_dr() As DataRow = CServiciosDataSet2.DetalleReceta.Select(_Selecciona)
        Dim _Cantidad As Integer
        Dim _Referencia As String

        DataGridView1.Rows.Clear()
        For Me.i = 0 To _Rows_mat.GetUpperBound(0)
            _Material = _Rows_mat(i).Item("Material")
            _Descripcion = _Rows_mat(i).Item("Descripcion")
            _Presentacion = _Rows_mat(i).Item("Presentacion")
            _Contenido = _Rows_mat(i).Item("Contenido")
            _PrecioUnitario = _Rows_mat(i).Item("PrecioUnitario")

            DataGridView1.Rows.Add(_Material, _Descripcion, _PrecioUnitario, _Presentacion, _Contenido)
        Next
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        DataGridView1.Enabled = False

        ' Si la Consulta ya tiene receta, muestra los datos
        _renglon = Frm3Pacientes.DataGridView3.CurrentRow
        If _renglon.Cells(8).Value <> "" Then

            ToolStripButton2.Enabled = False
            ToolStripButton3.Enabled = False
            ToolStripButton4.Enabled = True
            _Busca = CServiciosDataSet2.Recetas.FindByNumReceta(_Receta)
            TextBox1.Text = _Receta
            TextBox2.Text = _Busca.Item("FechaReceta")
            TextBox3.Text = _Busca.Item("HoraReceta")
            TextBox4.Text = _Busca.Item("Referencia")
            TextBox5.Text = _Busca.Item("Atiende")
            RichTextBox1.Text = _Busca.Item("Comentarios")

            ' Muestra el Detalle de la Receta
            For Me.i = 0 To _Rows_dr.GetUpperBound(0)
                _Material = _Rows_dr(i).Item("Material")
                _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_Material)
                _Descripcion = _Busca.Item("Descripcion")
                _PrecioUnitario = _Rows_dr(i).Item("PrecioUnitario")
                _Cantidad = _Rows_dr(i).Item("Cantidad")
                _Presentacion = _Busca.Item("Presentacion")
                _Contenido = _Busca.Item("Contenido")
                _Referencia = _Rows_dr(i).Item("Referencia")

                DataGridView2.Rows.Add(_Material, _Descripcion, _PrecioUnitario, _Cantidad, _Presentacion, _Contenido, _Referencia)

            Next
            DataGridView2.Enabled = False

        End If



    End Sub

    Private Sub ContextMenuStrip1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ContextMenuStrip1.Click

    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub EliminarRenglónToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EliminarRenglónToolStripMenuItem.Click
        Dim _indice As Integer = DataGridView2.CurrentRow.Index
        DataGridView2.Rows.RemoveAt(_indice)
    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _renglon = DataGridView2.CurrentRow
        If _renglon.Cells(0).Value <> "" Then
            _renglon.DefaultCellStyle.BackColor = Color.Gold
        End If

    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellLeave
        _renglon = DataGridView2.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        ' Modifica Receta
        Panel1.Enabled = True
        DataGridView1.Enabled = True
        DataGridView2.Enabled = True
        ToolStripButton4.Enabled = False
        ToolStripButton3.Enabled = True
        TextBox1.Enabled = False
        TextBox6.Text = "CAMBIO"
    End Sub
End Class