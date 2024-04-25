Public Class Frm3Pacientes
    Public i As Integer
    Public _renglon As DataGridViewRow
    Public checacontrol As Control
    Public _Busca As DataRow
    Public _Paciente As String
    Public _Referencia As String
    Public _Comentarios As String
    Public _Centro As String
    Public _Fila As DataRow
    Public msg As String
    Public Style As MsgBoxStyle

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Frm3Pacientes_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Escape) Then
            If TextBox8.Text = "ALTA" Or TextBox8.Text = "CAMBIO" Then
                Button1_Click(sender, e)
            End If
            If TextBox8.Text = "CITA" Then
                ToolStripButton16_Click(sender, e)
            End If
            If TextBox8.Text = "VALOR" Then
                Button5_Click(sender, e)
            End If

        End If
    End Sub

    Private Sub Frm3Pacientes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: This line of code loads data into the 'CServiciosDataSet5.Materiales' table. You can move, or remove it, as needed.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Recetas' table. You can move, or remove it, as needed.
        Me.RecetasTableAdapter.Fill(Me.CServiciosDataSet2.Recetas)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Pacientes' table. You can move, or remove it, as needed.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.EvetosPaciente' table. You can move, or remove it, as needed.
        Me.EvetosPacienteTableAdapter.Fill(Me.CServiciosDataSet2.EvetosPaciente)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Diagnosticos' table. You can move, or remove it, as needed.
        Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet2.Diagnosticos)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.DetalleReceta' table. You can move, or remove it, as needed.
        Me.DetalleRecetaTableAdapter.Fill(Me.CServiciosDataSet2.DetalleReceta)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Consultas' table. You can move, or remove it, as needed.
        '    Me.ConsultasTableAdapter.Fill(Me.CServiciosDataSet2.Consultas)
        'TODO: This line of code loads data into the 'CServiciosDataSet1.Usuarios' table. You can move, or remove it, as needed.
        Me.UsuariosTableAdapter.Fill(Me.CServiciosDataSet1.Usuarios)
        'TODO: This line of code loads data into the 'CServiciosDataSet.Centros' table. You can move, or remove it, as needed.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: This line of code loads data into the 'CServiciosDataSet3.Documentos' table. You can move, or remove it, as needed.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: This line of code loads data into the 'CServiciosDataSet4.Citas' table. You can move, or remove it, as needed.
        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)
        Me.ConsultasTableAdapter1.Fill(Me.CServiciosDataSet33.Consultas)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet17.Electros' Puede moverla o quitarla según sea necesario.
        Me.ElectrosTableAdapter.Fill(Me.CServiciosDataSet17.Electros)

        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet14.Medicos' Puede moverla o quitarla según sea necesario.
        Me.MedicosTableAdapter.Fill(Me.CServiciosDataSet14.Medicos)
        'TODO: This line of code loads data into the 'CServiciosDataSet6.Contactos' table. You can move, or remove it, as needed.
        Me.ContactosTableAdapter.Fill(Me.CServiciosDataSet6.Contactos)

        '   Dim checacontrol As Control
        Dim _usuario As String = Frm2Menu.ToolStripTextBox1.Text
        Dim _nivel As String = Trim(Frm2Menu.ToolStripTextBox3.Text)
        Dim _Rows_pa() As DataRow
        Dim _paterno, _materno, _nombres As String

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next
        Panel1.Enabled = False

        ' Valida el usuario
        If _nivel = "1" Then
            _Rows_pa = CServiciosDataSet2.Pacientes.Select("Centro = 'CME'")
        End If
        If _nivel = "2" Then
            _Rows_pa = CServiciosDataSet2.Pacientes.Select("Centro = 'FTR'")
        End If
        If _nivel = "3" Then
            _Rows_pa = CServiciosDataSet2.Pacientes.Select("Centro <> 'CME'")
        End If
        If _nivel = "5" Then
            _Rows_pa = CServiciosDataSet2.Pacientes.Select("Centro <> '   '")
        End If

        DataGridView1.Rows.Clear()
        For Me.i = 0 To _Rows_pa.GetUpperBound(0)
            _Paciente = _Rows_pa(i).Item("Paciente")
            _paterno = _Rows_pa(i).Item("Paterno")
            _materno = _Rows_pa(i).Item("Materno")
            _nombres = _Rows_pa(i).Item("Nombres")

            DataGridView1.Rows.Add(_Paciente, _paterno, _materno, _nombres)
        Next

        ' Llena el Combo con los Centros
        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet.Centros.Rows(i).Item("Descripcion"))
        Next


        ' Llena el Combo con los Diagósticos
        ComboBox2.Items.Clear()
        For Me.i = 0 To CServiciosDataSet2.Diagnosticos.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet2.Diagnosticos.Rows(i).Item("Descripcion"))
        Next

        ToolStripButton4.Enabled = False
        Panel3.Visible = False
        ToolStripButton10.Enabled = False
        ToolStripButton14.Enabled = False
        ToolStripButton16.Enabled = False
        Panel4.Visible = False
        ToolStripButton28.Enabled = False

        ComboBox4.Items.Clear()
        ComboBox4.Items.Add("MASCULINO")
        ComboBox4.Items.Add("FEMENINO")

        TextBox8.Text = ""
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        ' Dim checacontrol As Control

        Panel1.Enabled = True
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
        Next
        Panel1.Enabled = True
        For Each Me.checacontrol In Panel2.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
        Next
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True


        ToolStripButton4.Enabled = True
        TextBox1.Focus()
        TextBox8.Text = "ALTA"
        TextBox9.Text = "SIN FECHA"
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)

        '  MsgBox("Textbox 1 Got Focus")

        Dim _Consecutivo As Integer
        For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "PAC" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo")
            End If
        Next
        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "PA" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo

    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox1.Text = UCase(TextBox1.Text)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox2.Text = UCase(TextBox2.Text)
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox3.Text = UCase(TextBox3.Text)
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox6.Text = UCase(TextBox6.Text)
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox7.Text = CStr(Today)
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox9_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox4.Text = UCase(TextBox4.Text)
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        ' Guarda los datos del Paciente
        ' Dim checacontrol As Control
        Dim _Encuentra As Boolean

        ToolStripButton4.Enabled = False

        _Encuentra = False
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                If checacontrol.Text = "" Then
                    _Encuentra = True
                End If
            End If
        Next

        If _Encuentra Then
            MsgBox("Hay casillas en blanco. Debe proporcionar todos los datos")
            ToolStripButton4.Enabled = True
            Exit Sub
        End If

        If TextBox24.Text <> "MASCULINO" And TextBox24.Text <> "FEMENINO" And TextBox24.Text <> "M" And TextBox24.Text <> "F" Then
            MsgBox("En la casilla del Sexo, debe teclear los valores MASCULINO o FEMENINO o M  o F")
            TextBox24.Focus()
            Exit Sub
        End If


        ' Valida el Diagnóstico
        _Encuentra = False
        If TextBox11.Text = "" Then
            msg = "Debe seleccionar un DIAGNOSTICO para este Paciente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        Else
            For Me.i = 0 To CServiciosDataSet2.Diagnosticos.Rows.Count - 1
                If CServiciosDataSet2.Diagnosticos.Rows(i).Item("Diagnostico") = TextBox11.Text Then
                    _Encuentra = True
                End If
            Next

            If _Encuentra = False Then
                msg = "El Diagnóstico NO es correcto. Seleccione un dato del Combo respectivo"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If
        End If


        ' Valida el Centro
        If TextBox5.Text = "" Then
            msg = "Debe seleccionar un Centro para este Paciente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        Else
            _Encuentra = False
            For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
                If CServiciosDataSet.Centros.Rows(i).Item("Centro") = TextBox5.Text Then
                    _Encuentra = True
                End If
            Next
            If _Encuentra = False Then
                msg = "Debe seleccionar un CENTRO válido para este Paciente"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
                Exit Sub
            End If

        End If

        If Not IsDate(TextBox30.Text) Then
            msg = "Debe teclear una fecha de Nacimiento válida"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            TextBox30.Focus()
            Exit Sub
        End If



        Dim _Fila As DataRow
        Dim _Paciente As String
        Dim _Paterno As String
        Dim _Materno As String
        Dim _Nombres As String
        Dim _Centro As String
        Dim _Referencia As String
        Dim _FechaIngreso As String
        Dim _FechaAlta As String
        Dim _Siguiente As Integer
        Dim _Calle, _Colonia, _Ciudad, _Estado, _CodPost, _Telefono As String

        _Paciente = ""
        _Paterno = ""
        _Materno = ""
        _Nombres = ""
        _Centro = ""
        _Referencia = ""
        _FechaIngreso = ""
        _FechaAlta = ""
        _Calle = ""
        _Colonia = ""
        _Ciudad = ""
        _Estado = ""
        _CodPost = ""
        _Telefono = ""



        If TextBox8.Text = "ALTA" Then

            _Fila = CServiciosDataSet2.Pacientes.NewRow
            _Fila.Item("Paciente") = TextBox1.Text
        End If
        If TextBox8.Text = "CAMBIO" Then
            _Fila = CServiciosDataSet2.Pacientes.FindByPaciente(TextBox1.Text)
        End If

        _Fila.Item("Paterno") = TextBox2.Text
        _Fila.Item("Materno") = TextBox3.Text
        _Fila.Item("Nombres") = TextBox4.Text
        _Fila.Item("Centro") = TextBox5.Text
        _Fila.Item("Referencia") = TextBox6.Text
        _Fila.Item("FechaIngreso") = TextBox7.Text
        _Fila.Item("FechaAlta") = TextBox9.Text
        _Fila.Item("Diagnostico") = TextBox11.Text
        _Fila.Item("Calle") = TextBox10.Text
        _Fila.Item("Colonia") = TextBox12.Text
        _Fila.Item("Ciudad") = TextBox13.Text
        _Fila.Item("Estado") = TextBox14.Text
        _Fila.Item("CP") = TextBox15.Text
        _Fila.Item("Telefono") = TextBox16.Text
        _Fila.Item("FechaRegistro") = TextBox30.Text
        _Fila.Item("HoraRegistro") = TimeString
        _Fila.Item("Sexo") = TextBox24.Text
        _Fila.Item("Usuario") = LoginForm1.TextBox6.Text

        If TextBox8.Text = "ALTA" Then
            CServiciosDataSet2.Pacientes.Rows.Add(_Fila)
        End If

        Me.Validate()
        PacientesBindingSource.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet2)
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)


        If TextBox8.Text = "ALTA" Then
            For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
                If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "PAC" Then
                    _Siguiente = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo") + 1
                    CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo") = _Siguiente
                End If
            Next
        End If
        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager3.UpdateAll(CServiciosDataSet3)
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
        Next

        For Each Me.checacontrol In Panel2.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
        Next


        MsgBox("Registro guardado correctamente")
        Button1_Click(sender, e)
    End Sub

    Private Sub ComboBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim _Descripcion As String = ComboBox1.Text

        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            If CServiciosDataSet.Centros.Rows(i).Item("Descripcion") = _Descripcion Then
                TextBox5.Text = CServiciosDataSet.Centros.Rows(i).Item("Centro")
            End If
        Next


    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim _Descripcion As String = ComboBox2.Text
        TextBox11.Text = ""
        For Me.i = 0 To CServiciosDataSet2.Diagnosticos.Rows.Count - 1
            If CServiciosDataSet2.Diagnosticos.Rows(i).Item("Descripcion") = _Descripcion Then
                TextBox11.Text = CServiciosDataSet2.Diagnosticos.Rows(i).Item("Diagnostico")
            End If
        Next
        If TextBox11.Text = "" Then
            msg = "Debe seleccionar un Diagnóstico válido de la tabla que se le presenta"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
        End If
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.CurrentRow.Cells(0).Value = "" Then
            Exit Sub
        End If

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next
        Label32.Text = ""
        RichTextBox1.Text = ""



        Panel1.Enabled = True
        _renglon = DataGridView1.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.Gold


        TextBox1.Text = _renglon.Cells(0).Value
        TextBox2.Text = _renglon.Cells(1).Value
        TextBox3.Text = _renglon.Cells(2).Value
        TextBox4.Text = _renglon.Cells(3).Value

        ' Busca al Paciente para mostrar los datos
        _Busca = CServiciosDataSet2.Pacientes.FindByPaciente(TextBox1.Text)
        TextBox5.Text = _Busca.Item("Centro")
        TextBox6.Text = _Busca.Item("Referencia")
        TextBox7.Text = _Busca.Item("FechaIngreso")
        TextBox9.Text = _Busca.Item("FechaAlta")
        If TextBox9.Text = "SIN FECHA" Then
            TextBox9.BackColor = Color.LightGreen
        Else
            TextBox9.BackColor = Color.Crimson
        End If
        TextBox11.Text = _Busca.Item("Diagnostico")
        If DBNull.Value.Equals(_Busca.Item("Comentarios")) Then
            RichTextBox1.Text = ""
        Else
            RichTextBox1.Text = _Busca.Item("Comentarios")
        End If

        If DBNull.Value.Equals(_Busca.Item("Calle")) Then
            TextBox10.Text = ""
        Else
            TextBox10.Text = _Busca.Item("Calle")
        End If
        If DBNull.Value.Equals(_Busca.Item("Colonia")) Then
            TextBox12.Text = ""
        Else
            TextBox12.Text = _Busca.Item("Colonia")
        End If
        If DBNull.Value.Equals(_Busca.Item("Ciudad")) Then
            TextBox13.Text = ""
        Else
            TextBox13.Text = _Busca.Item("Ciudad")
        End If

        If DBNull.Value.Equals(_Busca.Item("Estado")) Then
            TextBox14.Text = ""
        Else
            TextBox14.Text = _Busca.Item("Estado")
        End If

        If DBNull.Value.Equals(_Busca.Item("CP")) Then
            TextBox15.Text = ""
        Else
            TextBox15.Text = _Busca.Item("CP")
        End If
        If DBNull.Value.Equals(_Busca.Item("Telefono")) Then
            TextBox16.Text = ""
        Else
            TextBox16.Text = _Busca.Item("Telefono")
        End If

        TextBox24.Text = _Busca.Item("Sexo")
        If _Busca.Item("Sexo") = "M" Then
            ComboBox4.Text = "MASCULINO"
        End If
        If _Busca.Item("Sexo") = "F" Then
            ComboBox4.Text = "FEMENINO"
        End If

        If DBNull.Value.Equals(_Busca.Item("FechaRegistro")) Then

        Else
            TextBox30.Text = _Busca.Item("FechaRegistro")
            ' Presenta la edad del Paciente
            If TextBox30.Text <> "" Then
                Dim _DiaHoy As Integer = Format$(Date.Today, "dd")
                Dim _FechaNace As Date = CDate(TextBox30.Text)
                Dim _DiaNace As Integer = Format$(_FechaNace, "dd")
                Dim _AñoHoy As Integer = Year(Date.Today)
                Dim _AñoNace As Integer = Year(_FechaNace)
                Dim _MesHoy As Integer = Month(Date.Today)
                Dim _MesNace As Integer = Month(_FechaNace)
                Dim _EdadDias = _DiaHoy - _DiaNace
                Dim _Meses, _Años As Integer
                If _EdadDias < 0 Then
                    _Meses = _MesHoy - _MesNace - 1
                Else
                    _Meses = _MesHoy - _MesNace
                End If
                If _Meses < 0 Then
                    _Meses = _Meses + 12
                    _Años = _AñoHoy - _AñoNace - 1
                Else
                    _Años = _AñoHoy - _AñoNace
                End If

                Label32.Text = "Edad : " & CStr(_Años) & " Años y " & CStr(_meses) & " Meses"
            End If

        End If

        ' Presenta la descripción del Centro
        _Busca = CServiciosDataSet.Centros.FindByCentro(TextBox5.Text)
        ComboBox1.Text = _Busca.Item("Descripcion")
        ComboBox1.Enabled = False
        ' Presenta la descripción del Diagnóstico
        _Busca = CServiciosDataSet2.Diagnosticos.FindByDiagnostico(TextBox11.Text)
        ComboBox2.Text = _Busca.Item("Descripcion")
        ComboBox2.Enabled = False

        'Presenta las Citas
        ToolStripButton15_Click(sender, e)

        ' Presenta las consultas
        ToolStripButton19_Click(sender, e)


        ' Presenta las Recetas
        ToolStripButton25_Click(sender, e)

        ' Presenta los Comentarios
        ToolStripButton26_Click(sender, e)

        ' Presenta los Contactos
        ToolStripButton26_Click(sender, e)

        ' Presenta los Electros
        ToolStripButton29_Click(sender, e)

        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = True

        ' Presenta la Foto
        If My.Computer.FileSystem.FileExists("z:\Fotos Paciente\" & TextBox1.Text & ".jpg") Then
            PictureBox1.Image = Image.FromFile("z:\Fotos Paciente\" & TextBox1.Text & ".jpg")
            With PictureBox1

                If .Image.Width < .Width And .Image.Height < .Height Then
                    .SizeMode = PictureBoxSizeMode.CenterImage
                ElseIf .Image.Width.ToString > .Width Or .Image.Height.ToString > .Height Then
                    .SizeMode = PictureBoxSizeMode.StretchImage
                End If
            End With
        Else
            msg = "No se ha proporcionado la foto de este Paciente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
 
    End Sub

    Private Sub DataGridView1_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _renglon = DataGridView1.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        ToolStripButton2.Enabled = False
        ToolStripButton4.Enabled = True
        ToolStripButton3.Enabled = False
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        TextBox8.Text = "CAMBIO"
        ComboBox4.Focus()
        TextBox24.Enabled = False
    End Sub

    Private Sub BindingNavigator1_RefreshItems(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigator1.RefreshItems

    End Sub

    Private Sub TextBox10_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox10.Text = UCase(TextBox10.Text)
    End Sub

    Private Sub TextBox10_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox12_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox12.Text = UCase(TextBox12.Text)
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox13_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox13.Text = UCase(TextBox13.Text)
    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox14_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox14.Text = UCase(TextBox14.Text)
    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
    
    End Sub

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox17.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False
    End Sub

    Private Sub TextBox17_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox17.GotFocus
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox17_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            If CServiciosDataSet.Centros.Rows(i).Item("Descripcion") = ComboBox3.Text Then
                TextBox19.Text = CServiciosDataSet.Centros.Rows(i).Item("Centro")
            End If
        Next
    End Sub

    Private Sub TextBox20_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox20.LostFocus
        TextBox20.Text = UCase(TextBox20.Text)
    End Sub

    Private Sub TextBox20_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub RichTextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox2.LostFocus
        RichTextBox2.Text = UCase(RichTextBox2.Text)
    End Sub

    Private Sub RichTextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox2.TextChanged

    End Sub

    Private Sub ToolStripButton11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton11.Click
        Panel3.Visible = False
        MonthCalendar1.Visible = False
    End Sub

    Private Sub ToolStripButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton10.Click
       
    End Sub

    Private Sub ToolStripButton12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton12.Click
        Panel3.Visible = True
        ToolStripButton14.Enabled = True
        ToolStripButton16.Enabled = True
        ToolStripButton12.Enabled = False

        MonthCalendar1.Visible = True
        TextBox17.Focus()
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            ComboBox3.Items.Add(CServiciosDataSet.Centros.Rows(i).Item("Descripcion"))
        Next
        ToolStripButton7.Enabled = False
        ToolStripButton9.Enabled = False

        Dim _Consecutivo As Integer

        ' Obtiene la siguiente Cita
        For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "CIT" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo")
            End If
        Next
        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "CI" + Mid(_XConsecutivo, _x, 4)
        TextBox21.Text = _XConsecutivo
        TextBox17.Focus()
        TextBox39.Text = "ALTA"

        TextBox8.Text = "CITA"


    End Sub

    Private Sub ToolStripButton14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton14.Click

        For Each Me.checacontrol In Panel3.Controls
            If TypeOf checacontrol Is TextBox Then
                If checacontrol.Text = "" Then
                    MsgBox("Hay Casillas en blanco. Debe llenar todos los campos")
                    Exit Sub
                End If
            End If
        Next
        ToolStripButton12.Enabled = True
        ToolStripButton14.Enabled = False
        Dim _Siguiente As Integer
        Dim _FechaCita As String = TextBox17.Text
        Dim _HoraCita As String = TextBox18.Text
        Dim _Cita As String = TextBox21.Text
        _Paciente = TextBox1.Text
        _Centro = TextBox19.Text
        _Referencia = TextBox20.Text
        _Comentarios = RichTextBox2.Text

        If TextBox39.Text = "ALTA" Then
            _Fila = CServiciosDataSet4.Citas.NewRow
            _Fila.Item("Paciente") = _Paciente
        End If
        If TextBox39.Text = "CAMBIO" Then
            _Fila = CServiciosDataSet4.Citas.FindByPacienteCita(_Paciente, _Cita)
        End If

        _Fila.Item("Cita") = _Cita
        _Fila.Item("FechaCita") = TextBox17.Text
        _Fila.Item("HoraCita") = TextBox18.Text
        _Fila.Item("Centro") = TextBox19.Text
        _Fila.Item("Referencia") = TextBox20.Text
        _Fila.Item("Comentarios") = UCase(RichTextBox2.Text)
        _Fila.Item("Atiende") = TextBox31.Text
        _Fila.Item("Estatus") = "PROGRAMADO"

        If TextBox39.Text = "ALTA" Then
            CServiciosDataSet4.Citas.Rows.Add(_Fila)
        End If


        Me.Validate()
        CitasBindingSource.EndEdit()
        Me.TableAdapterManager4.UpdateAll(CServiciosDataSet4)
        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)

        ' Incrementa el consecutivo
        For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "CIT" Then
                _Siguiente = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo") + 1
                CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo") = _Siguiente
            End If
        Next

        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager3.UpdateAll(CServiciosDataSet3)
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        ' Registra en el campo Comentarios del Paciente, las  indicaciones asignadas en la Cita Médica
        _Busca = CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + vbCrLf
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + " " + "CITA No. " + _Cita + " " + _FechaCita + " " + _HoraCita
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + vbCrLf
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + " " + UCase(RichTextBox2.Text)

        Me.Validate()
        PacientesBindingSource.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet2)
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        ToolStripButton6_Click(sender, e)

        For Each Me.checacontrol In Panel3.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next
        RichTextBox2.Text = ""

        MsgBox("Cita guardada correctamente")
        Panel3.Visible = False
        MonthCalendar1.Visible = False
        ToolStripButton15_Click(sender, e)
        ToolStripButton32.Enabled = True
    End Sub

    Private Sub ToolStripButton16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton16.Click
        Panel3.Visible = False
        MonthCalendar1.Visible = False
        ToolStripButton12.Enabled = True
        ToolStripButton14.Enabled = False
    End Sub

    Private Sub ToolStripButton15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton15.Click
        ' Presenta el Grid con las Citas
        Panel3.Visible = False
        MonthCalendar1.Visible = False

        Dim _Selecciona As String = "Paciente = " & "'" & TextBox1.Text & "'"
        Dim _Rows_ci() As DataRow = CServiciosDataSet4.Citas.Select(_Selecciona)
        Dim _FechaCita As String
        Dim _Horacita As String
        Dim _Atiende As String
        Dim _Cita As String
        Dim _NombreCentro, _Estatus As String
        Dim _Asiste As Boolean

        DataGridView2.Rows.Clear()
        For Me.i = 0 To _Rows_ci.GetUpperBound(0)
            _Paciente = _Rows_ci(i).Item("Paciente")
            _Cita = _Rows_ci(i).Item("Cita")
            _FechaCita = _Rows_ci(i).Item("FechaCita")
            _Horacita = _Rows_ci(i).Item("HoraCita")
            _Centro = _Rows_ci(i).Item("Centro")
            _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
            _NombreCentro = _Busca.Item("Descripcion")
            _Referencia = _Rows_ci(i).Item("Referencia")
            _Atiende = _Rows_ci(i).Item("Atiende")
            If DBNull.Value.Equals(_Rows_ci(i).Item("Comentarios")) Then
                _Comentarios = ""
            Else
                _Comentarios = _Rows_ci(i).Item("Comentarios")
            End If

            _Asiste = _Rows_ci(i).Item("Asiste")
            _Estatus = _Rows_ci(i).Item("Estatus")
            _Selecciona = False

            DataGridView2.Rows.Add(_Paciente, _Cita, _FechaCita, _Horacita, _Centro, _NombreCentro, _Atiende, _Referencia, _Comentarios, _Asiste, _Selecciona, _Estatus)

        Next

        For Me.i = 0 To DataGridView2.Rows.Count - 1

            If DataGridView2.Rows(i).Cells(9).Value = True Then
                DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                DataGridView2.Rows(i).ReadOnly = True

            End If
            If DataGridView2.Rows(i).Cells(11).Value = "CANCELADO" Then
                DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.LightCoral
            End If

        Next


    End Sub

    Private Sub TextBox22_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox22.LostFocus
        TextBox22.Text = UCase(TextBox22.Text)
    End Sub

    Private Sub TextBox22_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub ToolStripButton13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton13.Click
        'Imprime el Formato de Cita Médica
        Dim msg As String
        Dim style As MsgBoxStyle
        Dim _Cuantos As Integer
        _Cuantos = DataGridView2.Rows.Count
        ' MsgBox("Checa : " + CStr(_cuantos))

        msg = "No hay citas registradas para imprimir"
        style = MsgBoxStyle.Information

        If DataGridView2.Rows.Count <= 1 Then
            MsgBox(msg, style)
            Exit Sub
        Else
            msg = "Debe seleccionar una cita para poder Imprimir"
            If DataGridView2.CurrentRow.Cells(0).Value = "" Then
                MsgBox(msg, style)
                Exit Sub
            End If
        End If

        'c:\Archivos de programa (x86)\java\jre7
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Formato Cita Médica.xlsx"
        Dim _Nombre As String = "Paciente : " + "(" + Trim(TextBox1.Text) + ")" + " " + (TextBox2.Text) + " " + Trim(TextBox3.Text) + " " + Trim(TextBox4.Text)
        Dim _Domicilio As String = Trim(TextBox10.Text) + " " + Trim(TextBox12.Text) + " " + Trim(TextBox13.Text) + " " + Trim(TextBox14.Text)
        Dim _Cita As String = DataGridView2.CurrentRow.Cells(1).Value
        _Paciente = DataGridView2.CurrentRow.Cells(0).Value
        _Busca = CServiciosDataSet4.Citas.FindByPacienteCita(_Paciente, _Cita)
        If DBNull.Value.Equals(_Busca.Item("Comentarios")) Then
            _Comentarios = " "
        Else
            _Comentarios = DataGridView2.CurrentRow.Cells(8).Value
        End If
        _Comentarios = _Comentarios + ". " + "FAVOR DE PRESENTARSE A LA CITA EN FECHA Y HORA PUNTUALMENTE!!"
        TextBox23.Text = _Comentarios
        Dim _Mide As Integer = Len(_TextBox23.Text)
        'Dim _Cuantos As Integer = Int(_Mide / 75) + 1
        Dim _posicion As Integer
        Dim _Fraccion As String
        Dim _x As Integer

        _posicion = 17
        _Centro = DataGridView2.CurrentRow.Cells(4).Value
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de excel

        m_Excel.Worksheets("CITA").cells(15, 4).value = Today
        m_Excel.Worksheets("CITA").cells(15, 5).value = TimeString
        m_Excel.Worksheets("CITA").cells(3, 9).value = _Cita

        _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
        m_Excel.Worksheets("CITA").cells(7, 3).value = _Busca.Item("Descripcion") + " " + _Busca.Item("Referencia")
        m_Excel.Worksheets("CITA").cells(8, 3).value = _Nombre
        m_Excel.Worksheets("CITA").cells(10, 4).value = _Domicilio
        m_Excel.Worksheets("CITA").cells(11, 4).value = DataGridView2.CurrentRow.Cells(2).Value
        m_Excel.Worksheets("CITA").cells(12, 4).value = DataGridView2.CurrentRow.Cells(3).Value
        m_Excel.Worksheets("CITA").cells(13, 4).value = DataGridView2.CurrentRow.Cells(6).Value
        m_Excel.Worksheets("CITA").Cells(14, 4).value = DataGridView2.CurrentRow.Cells(7).Value

        ' Imprime las Indicaciones o Comentarios de la Cita
        _x = 1

        m_Excel.Worksheets("Hoja2").cells(1, 1).value = _Comentarios
        For Me.i = 0 To _Cuantos
            _Fraccion = (Mid(TextBox23.Text, _x, 75) + "-").ToString
            m_Excel.Worksheets("CITA").cells(_posicion, 2).value = _Fraccion
            _x = _x + 75
            _posicion = _posicion + 1
        Next


        m_Excel.Visible = True

    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick

        _renglon = DataGridView2.CurrentRow
        If _renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        If _renglon.Cells(9).Value = True Then
            msg = "Esta Cita ya tiene asignada una Consulta. No puede modificarla"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        _renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox23.Text = _renglon.Cells(8).Value
        _renglon.Cells(10).Value = True
        ToolStripButton33.Enabled = True
        ToolStripButton32.Enabled = True

    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellLeave
        _renglon = DataGridView2.CurrentRow
        If _renglon.Cells(9).Value = False Then
            _renglon.DefaultCellStyle.BackColor = Color.White
        Else
            _renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        End If
        _renglon.Cells(10).Value = False
    End Sub

    Private Sub ToolStripButton17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton17.Click
        ' Convertir una Cita en una Consulta Médica
        If DataGridView2.CurrentRow.Cells(0).Value = "" Then
            MsgBox("Debe seleccionar una Cita para Convertirla en Consulta Médica")
            Exit Sub
        End If
        If DataGridView2.CurrentRow.Cells(9).Value = True Then
            MsgBox("Esta Cita ya se atendió con una Consulta. No puede modificarla")
            Exit Sub
        End If

        DataGridView2.CurrentRow.Cells(9).Value = True

        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        Dim _Consecutivo As Integer
        Dim _XConsecutivo As String
        Dim _Siguiente As Integer
        Dim _Mide As Integer
        Dim _x As Integer
        Dim _FechaConsulta As String
        Dim _HoraConsulta As String
        Dim _NumReceta As String
        Dim _Atiende As String
        Dim _Cita As String = DataGridView2.CurrentRow.Cells(1).Value

        _XConsecutivo = ""
        _Mide = 0
        _x = 0
        _Siguiente = 0
        _FechaConsulta = ""
        _HoraConsulta = ""
        _Atiende = ""
        _NumReceta = ""
        msg = "Va a convertir una Cita en una Consulta Médica. Está Seguro?"   ' Define el mensaje.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Information Or MsgBoxStyle.YesNo
        title = "Registrar Asistencia a la Cita Médica"   ' Define el tìtulo.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.No Then   ' El usuario selecciona NO.
            ' Sale de el Procedimiento
            Exit Sub
        Else
            ' Convierte la Cita en Consulta Médica
            For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
                If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "CON" Then
                    _Consecutivo = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo")
                    Exit For
                End If
            Next
            _XConsecutivo = "000" + CStr(_Consecutivo)
            _Mide = Len(_XConsecutivo)
            _x = _Mide - 3
            _XConsecutivo = "CO" + Mid(_XConsecutivo, _x, 4)
            _Paciente = DataGridView2.CurrentRow.Cells(0).Value
            _FechaConsulta = DataGridView2.CurrentRow.Cells(2).Value
            _HoraConsulta = DataGridView2.CurrentRow.Cells(3).Value
            _Centro = DataGridView2.CurrentRow.Cells(4).Value
            _Atiende = DataGridView2.CurrentRow.Cells(6).Value
            _Referencia = "CITA No. " + DataGridView2.CurrentRow.Cells(1).Value

            _Fila = CServiciosDataSet33.Consultas.NewRow
            _Fila.Item("Consulta") = _XConsecutivo
            _Fila.Item("Paciente") = _Paciente
            _Fila.Item("FechaConsulta") = _FechaConsulta
            _Fila.Item("HoraConsulta") = _HoraConsulta
            _Fila.Item("NumReceta") = _NumReceta
            _Fila.Item("Centro") = _Centro
            _Fila.Item("Atiende") = _Atiende
            _Fila.Item("Referencia") = _Referencia
            _Fila.Item("Comentarios") = ""
            _Fila.Item("FechaRegistro") = CStr(Today)
            _Fila.Item("HoraRegistro") = TimeString
            _Fila.Item("Usuario") = LoginForm1.TextBox6.Text
            _Fila.Item("ImporteConsulta") = 0
            _Fila.Item("ImporteDescuento") = 0


            CServiciosDataSet33.Consultas.Rows.Add(_Fila)
            Me.Validate()
            ConsultasBindingSource1.EndEdit()
            Me.TableAdapterManager9.UpdateAll(CServiciosDataSet33)
            Me.ConsultasTableAdapter1.Fill(Me.CServiciosDataSet33.Consultas)

            ' Actualiza el Consecutivo 

            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("CON")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager3.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

            ToolStripButton19_Click(sender, e)


            ' Actualiza el Estatus de la Cita
            _Busca = CServiciosDataSet4.Citas.FindByPacienteCita(_Paciente, _Cita)
            _Busca.Item("Asiste") = True
            _Busca.Item("Estatus") = "ATENDIDO"
            Me.Validate()
            CitasBindingSource.EndEdit()
            Me.TableAdapterManager4.UpdateAll(CServiciosDataSet4)
            Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)


            MsgBox("Consulta registrada Correctamete")
            Exit Sub
        End If

    End Sub

    Private Sub ToolStripButton18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton18.Click



        Frm4CompletaConsulta.Show()

    End Sub

    Private Sub ToolStripButton19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton19.Click
        ' Refresca Información de Consultas
        Me.ConsultasTableAdapter.Fill(CServiciosDataSet2.Consultas)

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Debe seleccionar un Paciente para refrescar información de las Consultas")
            Exit Sub
        End If

        Dim _Pago As Integer
        Dim _Consulta As String
        Dim _FechaConsulta As String
        Dim _HoraConsulta As String
        Dim _NumReceta As String
        Dim _Atiende As String
        Dim _Cita As String
        Dim _NombreCentro As String
        Dim _Seleccion As String = "Paciente = " & "'" & TextBox1.Text & "'"
        Dim _Rows_co() As DataRow = CServiciosDataSet33.Consultas.Select(_Seleccion)
        Dim _Selecciona As Boolean

        DataGridView3.Rows.Clear()
        '  For i = 0 To CServiciosDataSet2.Consultas.Rows.Count - 1
        For Me.i = 0 To _Rows_co.GetUpperBound(0)
            _Paciente = _Rows_co(i).Item("Paciente")
            _Consulta = _Rows_co(i).Item("Consulta")
            _FechaConsulta = _Rows_co(i).Item("FechaConsulta")
            _HoraConsulta = _Rows_co(i).Item("HoraConsulta")
            _NumReceta = _Rows_co(i).Item("NumReceta")
            _Centro = _Rows_co(i).Item("Centro")
            _Atiende = _Rows_co(i).Item("Atiende")
            _Referencia = _Rows_co(i).Item("Referencia")
            _Comentarios = _Rows_co(i).Item("Comentarios")
            _Cita = Mid(_Referencia, 10, 10)
            _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
            _NombreCentro = _Busca.Item("Descripcion")
            _NumReceta = _Rows_co(i).Item("NumReceta")
            If DBNull.Value.Equals(_Rows_co(i).Item("Importeconsulta")) Then
                _Pago = 0
            Else
                _Pago = _Rows_co(i).Item("ImporteConsulta") - _Rows_co(i).Item("ImporteDescuento")
            End If

            _Selecciona = False

            DataGridView3.Rows.Add(_Paciente, _Cita, _Consulta, _FechaConsulta, _HoraConsulta, _Centro, _NombreCentro, _Atiende, _NumReceta, _Selecciona, _Pago)
        Next

        For Me.i = 0 To DataGridView3.Rows.Count - 1
            If DataGridView3.Rows(i).Cells(8).Value <> "" Then
                DataGridView3.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
            End If
            If DataGridView3.Rows(i).Cells(10).Value > 0 Then
                DataGridView3.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
            End If
        Next

    End Sub

    Private Sub DataGridView3_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        _renglon = DataGridView3.CurrentRow
        If _renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        If _renglon.Cells(9).Value = False Then
            _renglon.DefaultCellStyle.BackColor = Color.Gold
            _renglon.Cells(9).Value = True
            ToolStripButton18.Enabled = True
        Else
            _renglon.DefaultCellStyle.BackColor = Color.White
            _renglon.Cells(9).Value = False
            ToolStripButton18.Enabled = False
        End If

     
    End Sub

    Private Sub DataGridView3_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
     
    End Sub

    Private Sub DataGridView3_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellLeave
        _renglon = DataGridView3.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.White
        _renglon.Cells(9).Value = False


    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        ' Presenta los Comentarios del Paciente
        _Paciente = TextBox1.Text
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        _Busca = CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
        RichTextBox1.Text = _Busca.Item("Comentarios")
    End Sub

    Private Sub ToolStripButton23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton23.Click

    End Sub

    Private Sub ToolStripButton24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton24.Click
        'If DataGridView3.CurrentRow.Cells(8).Value <> "" Then
        ' MsgBox("Esa Consulta ya tiene asignada una Receta. No puede agregar otra")
        ' Exit Sub
        ' Else
        Frm5RecetaPaciente.Show()
        'End If


    End Sub

    Private Sub ToolStripButton25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton25.Click
        ' Muestra las Recetas de ese paciente
        If TextBox1.Text = "" Then
            MsgBox("Debe seleccionar un Paciente para visualizar las Recetas")
            Exit Sub
        End If
        Dim _Seleccion As String = "Paciente = " & "'" & TextBox1.Text & "'"
        Dim _Rows_r() As DataRow = CServiciosDataSet2.Recetas.Select(_Seleccion)
        Dim _Receta As String
        Dim _Consulta As String
        Dim _NombreCentro As String
        Dim _FechaReceta As String

        DataGridView4.Rows.Clear()
        For Me.i = 0 To _Rows_r.GetUpperBound(0)
            _Paciente = _Rows_r(i).Item("Paciente")
            _Receta = _Rows_r(i).Item("NumReceta")
            _Consulta = _Rows_r(i).Item("Consulta")
            _Centro = _Rows_r(i).Item("Centro")
            _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
            _NombreCentro = _Busca.Item("Descripcion")
            _FechaReceta = _Rows_r(i).Item("FechaReceta")
            _Referencia = _Rows_r(i).Item("Referencia")

            DataGridView4.Rows.Add(_Paciente, _Receta, _Consulta, _NombreCentro, _FechaReceta, _Referencia)
        Next


    End Sub

    Private Sub DataGridView4_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        _renglon = DataGridView4.CurrentRow
        If _renglon.Cells(0).Value <> "" Then
            Dim _Numreceta As String = _renglon.Cells(1).Value
            Dim _Seleccion As String = "NumReceta = " & "'" & _Numreceta & "'"
            Dim _Rows_dr() As DataRow = CServiciosDataSet2.DetalleReceta.Select(_Seleccion)
            Dim _Material, _Descripcion, _Contenido, _Presentacion As String
            Dim _Cantidad As Integer
            Dim _PrecioUnitario As Double

            _renglon.DefaultCellStyle.BackColor = Color.Gold

            ' Presenta los medicamentos asignados a esta receta

            DataGridView5.Rows.Clear()
            For Me.i = 0 To _Rows_dr.GetUpperBound(0)
                _Material = _Rows_dr(i).Item("Material")
                _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_Material)
                _Descripcion = _Busca.Item("Descripcion")
                _PrecioUnitario = _Rows_dr(i).Item("PrecioUnitario")
                _Cantidad = _Rows_dr(i).Item("Cantidad")
                _Presentacion = _Busca.Item("Presentacion")
                _Contenido = _Busca.Item("Contenido")
                _Referencia = _Rows_dr(i).Item("Referencia")

                DataGridView5.Rows.Add(_Material, _Descripcion, _PrecioUnitario, _Cantidad, _Presentacion, _Contenido, _Referencia)

            Next
        End If
    End Sub

    Private Sub DataGridView4_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView4_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellLeave
        _renglon = DataGridView4.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton27.Click
        If TextBox1.Text = "" Then
            MsgBox("Debe seleccionar un Paciente para poder registrar los Contactos")
            Exit Sub
        Else
            Panel4.Visible = True
            TextBox28.Focus()
            ToolStripButton27.Enabled = False
            ToolStripButton28.Enabled = True
        End If

    End Sub

    Private Sub ToolStripButton28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton28.Click
        ' Guarda el nuevo contacto

        If TextBox28.Text = "" Then
            MsgBox("Debe teclear un dato en el campo del Contacto")
            Exit Sub
        End If
        If TextBox28.Text = "" Then
            MsgBox("Debe teclear un dato en el campo del Domicilio del Contacto")
            Exit Sub
        End If
        If TextBox25.Text = "" Then
            MsgBox("Debe teclear un dato en el campo del Teléfono del Contacto")
            Exit Sub
        End If

        _Fila = CServiciosDataSet6.Contactos.NewRow
        _Fila.Item("Paciente") = TextBox1.Text
        _Fila.Item("Contacto") = TextBox28.Text
        _Fila.Item("Domicilio") = TextBox27.Text
        _Fila.Item("Telefono") = TextBox25.Text
        _Fila.Item("TipoContacto") = TextBox26.Text
        _Fila.Item("Email") = TextBox29.Text
        _Fila.Item("FechaRegistro") = CStr(Today)
        _Fila.Item("HoraRegistro") = TimeString
        _Fila.Item("Comentarios") = RichTextBox3.Text

        CServiciosDataSet6.Contactos.Rows.Add(_Fila)

        Me.Validate()
        ContactosBindingSource.EndEdit()
        Me.TableAdapterManager6.UpdateAll(CServiciosDataSet6)
        Me.ContactosTableAdapter.Fill(CServiciosDataSet6.Contactos)

        Panel4.Visible = False
        ToolStripButton27.Enabled = True
        ToolStripButton28.Enabled = False
        ToolStripButton26_Click(sender, e)

        MsgBox("Contacto guardado correctamente")

    End Sub

    Private Sub TextBox28_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox28.LostFocus
        TextBox28.Text = UCase(TextBox28.Text)
    End Sub

    Private Sub TextBox28_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox28.TextChanged

    End Sub

    Private Sub TextBox27_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox27.LostFocus
        TextBox27.Text = UCase(TextBox27.Text)
    End Sub

    Private Sub TextBox27_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox27.TextChanged

    End Sub

    Private Sub TextBox25_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox25.LostFocus
        TextBox25.Text = UCase(TextBox25.Text)
    End Sub

    Private Sub TextBox25_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub TextBox26_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox26.LostFocus
        TextBox26.Text = LCase(TextBox26.Text)
    End Sub

    Private Sub TextBox26_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox26.TextChanged

    End Sub

    Private Sub ToolStripButton26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton26.Click
        ' Muestra los Contactos
        If TextBox1.Text = "" Then
            MsgBox("Debe seleccionar un Paciente para mostrar los Contactos")
            Exit Sub
        End If
        _Paciente = TextBox1.Text
        Dim _Seleccion As String = "Paciente = " & "'" & _Paciente & "'"
        Dim _Rows_co() As DataRow = CServiciosDataSet6.Contactos.Select(_Seleccion)
        Dim _Contacto, _TipoContacto, _Domicilio, _Telefono, _Email As String
        Panel4.Visible = False
        DataGridView6.Rows.Clear()
        For Me.i = 0 To _Rows_co.GetUpperBound(0)
            _Contacto = _Rows_co(i).Item("Contacto")
            _TipoContacto = _Rows_co(i).Item("TipoContacto")
            _Domicilio = _Rows_co(i).Item("Domicilio")
            _Telefono = _Rows_co(i).Item("Telefono")
            _Email = _Rows_co(i).Item("Email")

            DataGridView6.Rows.Add(_Contacto, _Domicilio, _Telefono, _TipoContacto, _Email)

        Next
    End Sub

    Private Sub ToolStripButton21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton21.Click
        Frm7HojaPaciente.Show()

    End Sub

    Private Sub TextBox24_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox24_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox30_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox30.Text = "DD/MM/AAAA"
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox4_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        TextBox24.Text = Mid(ComboBox4.Text, 1, 1)
    End Sub

    Private Sub ComboBox4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If ComboBox4.Text = "MASCULINO" Then
            TextBox24.Text = "M"
        Else
            TextBox24.Text = "F"
        End If
    End Sub

    Private Sub TextBox18_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox18.GotFocus
        'Frm17HorarioAtencion.Show()

    End Sub

    Private Sub TextBox18_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox18.TextChanged

    End Sub

    Private Sub ComboBox5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.GotFocus
        ComboBox5.Items.Clear()
        Dim _Nombre As String
        For Me.i = 0 To Me.CServiciosDataSet14.Medicos.Rows.Count - 1
            _Nombre = Me.CServiciosDataSet14.Medicos.Rows(i).Item("Nombre")
            ComboBox5.Items.Add(_Nombre)
        Next
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet14.Medicos.Rows.Count - 1
            If Me.CServiciosDataSet14.Medicos.Rows(i).Item("Nombre") = ComboBox5.Text Then
                TextBox31.Text = Me.CServiciosDataSet14.Medicos.Rows(i).Item("Medico")
            End If
        Next

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Frm17HorarioAtencion.Show()
    End Sub

    Private Sub ToolStripButton30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Frm25NuevoElectro.Show()

    End Sub

    Private Sub ToolStripButton31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Frm25NuevoElectro.Show()

    End Sub

    Private Sub ToolStripButton29_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton29.Click

        ' Presenta los Electros
        If TextBox1.Text = "" Then
            MsgBox("Debe seleccionar un Paciente para mostrar los Electroencefalogramas")
            Exit Sub
        End If

        Me.ElectrosTableAdapter.Fill(Me.CServiciosDataSet17.Electros)

        DataGridView7.Rows.Clear()
        Dim _Rows_ee() As DataRow = CServiciosDataSet17.Electros.Select("Paciente = " & "'" & TextBox1.Text & "'")
        Dim _Electro, _Horaelectro, _Montaje, _Sensibilidad, _Frecuencia, _IdElectro, _Consecutivo, _Atiende As String
        Dim _FechaElectro As Date
        Dim _Exitoso As Boolean

        For Me.i = 0 To _Rows_ee.GetUpperBound(0)
            _Paciente = _Rows_ee(i).Item("Paciente")
            _Electro = _Rows_ee(i).Item("Electro")
            _FechaElectro = _Rows_ee(i).Item("FechaElectro")
            _Horaelectro = _Rows_ee(i).Item("Horaelectro")
            _IdElectro = _Rows_ee(i).Item("IdElecro")
            _Atiende = _Rows_ee(i).Item("Referencia")
            _Consecutivo = _Rows_ee(i).Item("Consecutivo")
            _Exitoso = _Rows_ee(i).Item("Exitoso")
            _Montaje = _Rows_ee(i).Item("Montaje")
            _Sensibilidad = _Rows_ee(i).Item("Sensibilidad")
            _Frecuencia = _Rows_ee(i).Item("Frecuencia")

            DataGridView7.Rows.Add(_Paciente, _Electro, _FechaElectro, _Horaelectro, _IdElectro, _Consecutivo, _Exitoso, _Atiende, _Montaje, _Sensibilidad, _Frecuencia)

        Next



    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim _Encuentra As Boolean
        Dim _Rows_pa() As DataRow
        Dim _Paterno, _Materno, _Nombres As String
        Dim _parte, _partenombre, _texto As String
        Dim _mide, x As Integer


        ' Borra los datagrid
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        DataGridView4.Rows.Clear()
        DataGridView5.Rows.Clear()
        DataGridView6.Rows.Clear()
        DataGridView7.Rows.Clear()

        ' Borra los datos de las cajas de texto
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next
        Label32.Text = ""
        RichTextBox1.Text = ""

        ' Borra los comentarios 
        RichTextBox1.Text = ""

        If RadioButton1.Checked = True Then
            ' **********************************************************************************************************************
            ' Busqueda por la Clave del Paciente
            _Paciente = UCase(InputBox("Teclee la Clave del Paciente a buscar"))
            _Rows_pa = CServiciosDataSet2.Pacientes.Select("Paciente = " & "'" & _Paciente & "'")

            _Encuentra = False
            For Me.i = 0 To _Rows_pa.GetUpperBound(0)
                _Encuentra = True
                _Paciente = _Rows_pa(i).Item("Paciente")
                _Paterno = _Rows_pa(i).Item("Paterno")
                _Materno = _Rows_pa(i).Item("Materno")
                _Nombres = _Rows_pa(i).Item("Nombres")

                DataGridView1.Rows.Add(_Paciente, _Paterno, _Materno, _Nombres)

            Next

            If _Encuentra = False Then
                msg = "No existe ese Paciente. Verififque"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
            End If
        End If


        If RadioButton2.Checked Then
            ' Busca por Nombre o por una parte
            ' **********************************************************************************************************************
            _parte = InputBox("Teclee el NOMBRE o parte deñ NOMBRE a buscar")
            _parte = Trim(UCase(_parte))
            _mide = Len(_parte)
            If _mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
                _texto = Trim(CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")) & " " & Trim(CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno")) & " " & Trim(CServiciosDataSet2.Pacientes.Rows(i).Item("Materno"))
                _Paciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
                _Paterno = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno")
                _Materno = CServiciosDataSet2.Pacientes.Rows(i).Item("Materno")
                _Nombres = CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")


                For x = 1 To Len(_texto)
                    _partenombre = Mid(_texto, x, _mide)

                    If _partenombre = _parte Then
                        DataGridView1.Rows.Add(_Paciente, _Paterno, _Materno, _Nombres)
                    End If
                Next
                


            Next

            If DataGridView1.Rows.Count = 0 Then
                msg = "No existen Pacientes que contengan ese Nombre"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
            End If

        End If
       
        If RadioButton3.Checked Then
            ' Busca por REFERENCIA
            ' **********************************************************************************************************************
            _parte = InputBox("Teclee la REFERENCIA a buscar")
            _parte = Trim(UCase(_parte))
            _mide = Len(_parte)
            If _mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
                _texto = Trim(UCase(CServiciosDataSet2.Pacientes.Rows(i).Item("Referencia")))
                _Paciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
                _Paterno = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno")
                _Materno = CServiciosDataSet2.Pacientes.Rows(i).Item("Materno")
                _Nombres = CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")


                For x = 1 To Len(_texto)
                    _partenombre = Mid(_texto, x, _mide)

                    If _partenombre = _parte Then
                        DataGridView1.Rows.Add(_Paciente, _Paterno, _Materno, _Nombres)
                    End If
                Next



            Next

            If DataGridView1.Rows.Count = 0 Then
                msg = "No existen Pacientes que contengan esa REFERENCIA"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
            End If

        End If


        If RadioButton4.Checked Then
            ' Busca por DOMICILIO
            ' **********************************************************************************************************************
            _parte = InputBox("Teclee el DOMICILIO a buscar")
            _parte = Trim(UCase(_parte))
            _mide = Len(_parte)
            If _mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
                _texto = Trim(UCase(CServiciosDataSet2.Pacientes.Rows(i).Item("Calle")))
                _texto = _texto & " " & Trim(UCase(CServiciosDataSet2.Pacientes.Rows(i).Item("Colonia")))
                _texto = _texto & " " & Trim(UCase(CServiciosDataSet2.Pacientes.Rows(i).Item("Ciudad")))
                _texto = _texto & " " & Trim(UCase(CServiciosDataSet2.Pacientes.Rows(i).Item("Estado")))

                _Paciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
                _Paterno = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno")
                _Materno = CServiciosDataSet2.Pacientes.Rows(i).Item("Materno")
                _Nombres = CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")

                For x = 1 To Len(_texto)
                    _partenombre = Mid(_texto, x, _mide)

                    If _partenombre = _parte Then
                        DataGridView1.Rows.Add(_Paciente, _Paterno, _Materno, _Nombres)
                    End If
                Next

            Next

            If DataGridView1.Rows.Count = 0 Then
                msg = "No existen Pacientes que contengan esa DOMICILIO"
                Style = MsgBoxStyle.Information
                MsgBox(msg, Style)
            End If

        End If


        If RadioButton5.Checked Then
            ' Busca TODOS
            ' **********************************************************************************************************************
            For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1

                _Paciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
                _Paterno = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno")
                _Materno = CServiciosDataSet2.Pacientes.Rows(i).Item("Materno")
                _Nombres = CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")

                DataGridView1.Rows.Add(_Paciente, _Paterno, _Materno, _Nombres)

            Next

        End If

    End Sub

    Private Sub ToolStripButton30_Click_1(sender As System.Object, e As System.EventArgs) Handles ToolStripButton30.Click
        ' Valorizar la Consulta
        ToolStripButton30.Enabled = False
        Dim _Find As Boolean
        Dim _Consulta As String

        _Find = False
        For Me.i = 0 To DataGridView3.Rows.Count - 1
            If DataGridView3.Rows(i).Cells(9).Value = True Then
                _Find = True
            End If
        Next

        If _Find = False Then
            msg = "No hay Consultas para Valorizar"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            ToolStripButton30.Enabled = True
            Exit Sub
        End If


        If _renglon.Cells(9).Value = False Then
            msg = "Debe seleccionar una Consulta para registrar el pago en efectivo"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            ToolStripButton30.Enabled = True
            Exit Sub
        End If

        _Consulta = _renglon.Cells(2).Value
        _Busca = CServiciosDataSet33.Consultas.FindByConsulta(_Consulta)



        If _Busca.Item("Importeconsulta") > 0 Then
            msg = "Esta Consulta ya tiene registrado un Pago"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        TextBox8.Text = "VALOR"




            Panel5.Visible = True
            TextBox32.Text = _renglon.Cells(2).Value



    End Sub

    Private Sub Label38_Click(sender As System.Object, e As System.EventArgs) Handles Label38.Click

    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Panel5.Visible = False
        MonthCalendar2.Visible = False
        TextBox33.Focus()
        ToolStripButton30.Enabled = True
    End Sub

    Private Sub TextBox35_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox35.GotFocus
        Dim _Neto = CInt(TextBox33.Text) - CInt(TextBox34.Text)
        TextBox35.Text = _Neto
    End Sub

    Private Sub TextBox35_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox35.TextChanged
 
    End Sub

    Private Sub TextBox37_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox37.GotFocus
        MonthCalendar2.Visible = True

    End Sub

    Private Sub TextBox37_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox37.TextChanged

    End Sub

    Private Sub MonthCalendar2_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateChanged

    End Sub

    Private Sub MonthCalendar2_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateSelected
        TextBox37.Text = MonthCalendar2.SelectionRange.Start
        MonthCalendar2.Visible = True
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        ' Guarda los datos del Pago
        Dim _Consulta As String = _renglon.Cells(2).Value
        Dim _Find As Boolean
        Dim title As String
        Dim response As MsgBoxResult
        msg = "Está a punto de guardar los datos del pago de la consulta. Desea continuar?"   ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Pago de Consultas"

        _Find = False
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            For Each Me.checacontrol In Panel5.Controls
                If TypeOf checacontrol Is TextBox Then
                    If checacontrol.Text = "" Then
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

            ' Verfica que no se haya registrado el pago
            _Busca = CServiciosDataSet33.Consultas.FindByConsulta(_Consulta)

            If DBNull.Value.Equals(_Busca.Item("Adicional")) Then
                ' Continua con el proceso de registro de pago
            Else
                If _Busca.Item("Adicional") <> "" Then
                    msg = "Esta consulta ya tiene registrado un pago. Verifique"
                    Style = MsgBoxStyle.Information
                    MsgBox(msg, Style)
                    Exit Sub
                End If
            End If

            _Busca.Item("ImporteConsulta") = CInt(TextBox33.Text)
            _Busca.Item("ImporteDescuento") = CInt(TextBox34.Text)
            _Busca.Item("Adicional") = TextBox38.Text & " " & TextBox37.Text

            Me.Validate()
            ConsultasBindingSource1.EndEdit()
            Me.TableAdapterManager9.UpdateAll(CServiciosDataSet33)
            Me.ConsultasTableAdapter1.Fill(Me.CServiciosDataSet33.Consultas)

            msg = "Registro guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Button5_Click(sender, e)
            Exit Sub

        Else
            Exit Sub
        End If


    End Sub

    Private Sub ToolStripButton32_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton32.Click
        ' Modificación de una Cita
        Dim _Find As Boolean
        Dim _Cita, _NombreMedico As String
        Dim _Rows_Ci() As DataRow
        Dim _Atiende, _HoraCita As String
        Dim _Asiste As Boolean

        _Find = False
        For Me.i = 0 To DataGridView2.Rows.Count - 1
            _Find = True
        Next
        If _Find = False Then
            msg = "No existen Citas para este Paciente. Verifique"
            Style = MsgBox(msg, Style)
            Exit Sub
        End If

        If _renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        _Cita = _renglon.Cells(1).Value

        _Rows_Ci = CServiciosDataSet4.Citas.Select("Cita = " & "'" & _Cita & "'")
        _Find = False

        _Asiste = False
        _HoraCita = ""
        _NombreMedico = ""
        _Atiende = ""
        For Me.i = 0 To _Rows_Ci.GetUpperBound(0)
            _Find = True
            TextBox21.Text = _Cita
            _Atiende = _Rows_Ci(i).Item("Atiende")
            _Busca = CServiciosDataSet14.Medicos.FindByMedico(_Atiende)
            _NombreMedico = _Busca.Item("Nombre")
            _HoraCita = _Rows_Ci(i).Item("HoraCita")
            _Centro = _Rows_Ci(i).Item("Centro")
            _Referencia = _Rows_Ci(i).Item("Referencia")
            _Asiste = _Rows_Ci(i).Item("Asiste")
        Next

        If _Find = False Then
            Exit Sub
        End If

        If _Asiste = True Then
            msg = "El Paciente asistió a esta Cita. No puede Modificarla"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If

        Panel3.Visible = True
        MonthCalendar1.Visible = True

        ToolStripButton14.Enabled = True
        ToolStripButton32.Enabled = False
        ToolStripButton16.Enabled = True

        Label21.Text = "Cambia Cita"
        TextBox31.Text = _Atiende
        ComboBox5.Text = _NombreMedico
        TextBox18.Text = _HoraCita
        TextBox17.Text = _renglon.Cells(2).Value
        TextBox19.Text = _Centro
        _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
        ComboBox3.Text = _Busca.Item("Descripcion")
        TextBox20.Text = _Referencia
        TextBox22.Text = _NombreMedico
        TextBox39.Text = "CAMBIO"


    End Sub

    Private Sub ToolStripButton33_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton33.Click
        ToolStripButton16_Click(sender, e)
        Dim _Find As Boolean
        Dim _Cita As String
        Dim title As String
        Dim response As MsgBoxResult

        _Find = False

        For Me.i = 0 To DataGridView3.Rows.Count - 1
            _Find = True
        Next

        If _Find = False Then
            msg = "Para Cancelar una Cita, debe seleccionar ese Renglón"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If


        msg = "Está a punto de Cancelar esta Cita. Desea continuar?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cancelación de Citas"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then

            _Paciente = _renglon.Cells(0).Value
            _Cita = _renglon.Cells(1).Value

            _Busca = CServiciosDataSet4.Citas.FindByPacienteCita(_Paciente, _Cita)
            _Busca.Item("Estatus") = "CANCELADO"

            Me.Validate()
            CitasBindingSource.EndEdit()
            Me.TableAdapterManager4.UpdateAll(CServiciosDataSet4)
            Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)

            msg = "Registro Eliminado correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub

        Else
            Exit Sub
        End If

    End Sub

    Private Sub TextBox1_GotFocus1(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus
        Dim _Consecutivo As Integer
        For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "PAC" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo")
            End If
        Next
        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "PA" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo

    End Sub

    Private Sub TextBox1_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox7_GotFocus1(sender As Object, e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.Text = Today
    End Sub

    Private Sub TextBox7_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox30_GotFocus1(sender As Object, e As System.EventArgs) Handles TextBox30.GotFocus
        TextBox30.Text = Today

    End Sub

    Private Sub TextBox30_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox30.LostFocus
      
    End Sub

    Private Sub TextBox30_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox30.TextChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim _Descripcion As String = ComboBox2.Text
        TextBox11.Text = ""
        For Me.i = 0 To CServiciosDataSet2.Diagnosticos.Rows.Count - 1
            If CServiciosDataSet2.Diagnosticos.Rows(i).Item("Descripcion") = _Descripcion Then
                TextBox11.Text = CServiciosDataSet2.Diagnosticos.Rows(i).Item("Diagnostico")
            End If
        Next
        If TextBox11.Text = "" Then
            msg = "Debe seleccionar un Diagnóstico válido de la tabla que se le presenta"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.Text = "MASCULINO" Then
            TextBox24.Text = "M"
        Else
            TextBox24.Text = "F"
        End If
    End Sub

    Private Sub TextBox2_LostFocus1(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.Text = UCase(TextBox2.Text)
    End Sub

    Private Sub TextBox2_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_LostFocus1(sender As Object, e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.Text = UCase(TextBox3.Text)
    End Sub

    Private Sub TextBox3_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox4_LostFocus1(sender As Object, e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.Text = UCase(TextBox4.Text)
    End Sub

    Private Sub TextBox4_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox6_LostFocus1(sender As Object, e As System.EventArgs) Handles TextBox6.LostFocus
        TextBox6.Text = UCase(TextBox6.Text)
    End Sub

    Private Sub TextBox6_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            If ComboBox1.Text = CServiciosDataSet.Centros.Rows(i).Item("Descripcion") Then
                TextBox5.Text = CServiciosDataSet.Centros.Rows(i).Item("Centro")
            End If
        Next
    End Sub

    Private Sub TextBox9_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox9.LostFocus
        If Not IsDate(TextBox9.Text) Then
            msg = "Debe teclear una fecha de Ingreso válida"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            TextBox9.Focus()
        End If
    End Sub

    Private Sub TextBox9_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub ToolStripButton34_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton34.Click
        ToolStripButton26_Click(sender, e)
    End Sub

    Private Sub ToolStripButton31_Click_1(sender As System.Object, e As System.EventArgs) Handles ToolStripButton31.Click
        Frm25NuevoElectro.Show()
    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Button1_Click(sender, e)
    End Sub
End Class