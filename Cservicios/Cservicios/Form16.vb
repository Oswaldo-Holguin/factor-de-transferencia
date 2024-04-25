Public Class Frm16PersonalMedico
    Public i As Integer
    Public _Medico, _Nombre As String
    Public _Renglon As DataGridViewRow
    Public _checacontrol As Control
    Public _centro, _Descripcion As String
    Public msg As String
    Public Style As MsgBoxStyle
    Public _Busca As DataRow



    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Frm16PersonalMedico_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet14.Medicos' Puede moverla o quitarla según sea necesario.
        Me.MedicosTableAdapter1.Fill(Me.CServiciosDataSet14.Medicos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet13.Especialidades' Puede moverla o quitarla según sea necesario.
        Me.EspecialidadesTableAdapter.Fill(Me.CServiciosDataSet13.Especialidades)
        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet12.Medicos' Puede moverla o quitarla según sea necesario.
        Me.MedicosTableAdapter1.Fill(Me.CServiciosDataSet14.Medicos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        For Each Me._checacontrol In Panel1.Controls
            If TypeOf Me._checacontrol Is TextBox Then
                Me._checacontrol.Text = ""

            End If
            If TypeOf _checacontrol Is ComboBox Then
                _checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""

        Panel1.Enabled = False

        ' Llena el Grid del Personal Médico
        DataGridView1.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet14.Medicos.Rows.Count - 1
            _Medico = Me.CServiciosDataSet14.Medicos.Rows(i).Item("Medico")
            _Nombre = Me.CServiciosDataSet14.Medicos.Rows(i).Item("Nombre")

            DataGridView1.Rows.Add(_Medico, _Nombre)

        Next


        ' Llena el Combo con los datos de los Centros
        For Me.i = 0 To Me.CServiciosDataSet.Centros.Rows.Count - 1
            _Descripcion = Me.CServiciosDataSet.Centros.Rows(i).Item("Descripcion")

            ComboBox1.Items.Add(_Descripcion)
        Next
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = False

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True

        For Each Me._checacontrol In Panel1.Controls
            If TypeOf _checacontrol Is TextBox Then
                _checacontrol.Text = ""
            End If
        Next
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""


        ' Obtiene el siguiente prefijo para Personal Médico
        Dim _busca As DataRow = Me.CServiciosDataSet3.Documentos.FindByDocumento("MED")
        Dim _Siguiente As Integer = _busca.Item("Consecutivo")
  
        Dim _XConsecutivo As String = "000" + CStr(_Siguiente)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "PM" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()
        TextBox11.Text = "ACTIVO"
        TextBox11.Enabled = False
        TextBox12.Text = "ALTA"

        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True

    End Sub

    Private Sub ComboBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.GotFocus

        ' Llena el Combo con los datos de los Centros
        ComboBox1.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet.Centros.Rows.Count - 1
            _Descripcion = Me.CServiciosDataSet.Centros.Rows(i).Item("Descripcion")
            ComboBox1.Items.Add(_Descripcion)
        Next


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet.Centros.Rows.Count - 1
            If Me.CServiciosDataSet.Centros.Rows(i).Item("Descripcion") = ComboBox1.Text Then
                _centro = Me.CServiciosDataSet.Centros.Rows(i).Item("Centro")
            End If
        Next
        TextBox3.Text = _centro
        TextBox3.Enabled = False

    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.Text = UCase(TextBox2.Text)
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

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

    Private Sub TextBox6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.LostFocus
        TextBox6.Text = UCase(TextBox6.Text)
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox9_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox9.LostFocus
        TextBox9.Text = LCase(TextBox9.Text)
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox10.LostFocus
        TextBox10.Text = UCase(TextBox10.Text)
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub ComboBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.GotFocus
        Dim _HorarioInicial As String
        ComboBox2.Items.Clear()
        For Me.i = 7 To 23
            _HorarioInicial = CStr(i) & ":" & "00"
            ComboBox2.Items.Add(_HorarioInicial)
            _HorarioInicial = CStr(i) & ":" & "30"
            ComboBox2.Items.Add(_HorarioInicial)

        Next

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub ComboBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.GotFocus
        Dim _HorarioFinal As String
        ComboBox3.Items.Clear()
        For Me.i = 7 To 23
            _HorarioFinal = CStr(i) & ":" & "00"
            ComboBox3.Items.Add(_HorarioFinal)
            _HorarioFinal = CStr(i) & ":" & "30"
            ComboBox3.Items.Add(_HorarioFinal)
        Next
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        ToolStripButton5.Enabled = False
        ToolStripButton3.Enabled = True
        ToolStripButton2.Enabled = False
        TextBox12.Text = "CAMBIO"
        Panel1.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Focus()

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim _vacio As Boolean
        _vacio = False
        ' Verifica que no haya casillas en blanco
        For Each Me._checacontrol In Panel1.Controls
            If TypeOf _checacontrol Is TextBox Then
                If _checacontrol.Text = "" Then
                    _vacio = True
                End If
            End If
        Next
        If _vacio Then
            msg = "Hay casillas en blanco. Debe llenar todos los datos"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If


        ' Da de alta el nuevo registro o los cambios a uno existente
        Dim _Fila As DataRow
        Dim _Medico As String = TextBox1.Text
        If TextBox12.Text = "ALTA" Then
            _Fila = Me.CServiciosDataSet14.Medicos.NewRow
            _Fila.Item("Medico") = _Medico

        Else
            _Fila = Me.CServiciosDataSet14.Medicos.FindByMedico(_Medico)
        End If

        _Fila.Item("Nombre") = TextBox2.Text
        _Fila.Item("Centro") = TextBox3.Text
        _Fila.Item("Especialidad") = TextBox4.Text
        _Fila.Item("Referencia") = TextBox5.Text
        _Fila.Item("Prefijo") = TextBox6.Text
        _Fila.Item("TelefonoCasa") = Mid(TextBox7.Text, 1, 20)
        _Fila.Item("TelefonoCelular") = Mid(TextBox8.Text, 1, 20)
        _Fila.Item("CorreoElectronico") = TextBox9.Text
        _Fila.Item("ConsultorioExterno") = TextBox10.Text
        _Fila.Item("HorarioInicial") = ComboBox2.Text
        _Fila.Item("HorarioFinal") = ComboBox3.Text
        _Fila.Item("Estatus") = Mid(TextBox11.Text, 1, 3)
        _Fila.Item("Comentarios") = UCase(RichTextBox1.Text)

        If TextBox12.Text = "ALTA" Then
            Me.CServiciosDataSet14.Medicos.Rows.Add(_Fila)
        End If

        ' Actualiza los datos de la tabla de Médicos
        Me.Validate()
        MedicosBindingSource1.EndEdit()
        Me.TableAdapterManager4.UpdateAll(Me.CServiciosDataSet14)
        Me.MedicosTableAdapter1.Fill(Me.CServiciosDataSet14.Medicos)

        ' Actualiza el consecutivo
        If TextBox12.Text = "ALTA" Then
            _Busca = Me.CServiciosDataSet3.Documentos.FindByDocumento("MED")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager2.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        End If

        Button1_Click(sender, e)

        msg = "Registro guardado Correctamente"
        Style = MsgBoxStyle.Information

        MsgBox(msg, Style)

    End Sub

    Private Sub ComboBox4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.GotFocus
        For Me.i = 0 To Me.CServiciosDataSet13.Especialidades.Rows.Count - 1
            _Descripcion = Me.CServiciosDataSet13.Especialidades.Rows(i).Item("Descripcion")
            ComboBox4.Items.Add(_Descripcion)
        Next
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        TextBox4.Text = ComboBox4.Text
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim _Buscacentro As DataRow
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _Medico = _Renglon.Cells(0).Value
        _Busca = Me.CServiciosDataSet14.Medicos.FindByMedico(_Medico)
        Panel1.Enabled = True
        TextBox1.Text = _Medico
        TextBox2.Text = _Busca.Item("Nombre")
        TextBox3.Text = _Busca.Item("Centro")
        _Buscacentro = Me.CServiciosDataSet.Centros.FindByCentro(TextBox3.Text)
        ComboBox1.Text = _Buscacentro.Item("Descripcion")
        TextBox4.Text = _Busca.Item("Especialidad")
        ComboBox4.Text = _Busca.Item("Especialidad")
        TextBox5.Text = _Busca.Item("Referencia")
        TextBox6.Text = _Busca.Item("Prefijo")
        TextBox7.Text = _Busca.Item("TelefonoCasa")
        TextBox8.Text = _Busca.Item("TelefonoCelular")
        TextBox9.Text = _Busca.Item("CorreoElectronico")
        TextBox10.Text = _Busca.Item("ConsultorioExterno")
        If _Busca.Item("Estatus") = "ACT" Then
            TextBox11.Text = _Busca.Item("Estatus")
        End If

        RichTextBox1.Text = _Busca.Item("Comentarios")
        ComboBox2.Text = _Busca.Item("HorarioInicial")
        ComboBox3.Text = _Busca.Item("HorarioFinal")

        ToolStripButton5.Enabled = True










    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub
End Class