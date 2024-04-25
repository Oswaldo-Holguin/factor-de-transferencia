Public Class Frm8Empleados
    Public i, n As Integer
    Public _NumEmpleado As String
    Public _Busca As DataRow
    Public _Renglon As DataGridViewRow
    Public _Centro As String
    Public _Departamento, _categoria As String
    Public _Paterno As String
    Public _Materno As String
    Public _Nombre As String
    Public _Descripcion As String
    Public _fila As DataRow
    Public checacontrol As Control
    Public _Rows_em() As DataRow
    Public msg As String
    Public style As MsgBoxStyle

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Frm8Empleados_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: This line of code loads data into the 'CServiciosDataSet7.Permisos' table. You can move, or remove it, as needed.
        Me.PermisosTableAdapter.Fill(Me.CServiciosDataSet7.Permisos)
        'TODO: This line of code loads data into the 'CServiciosDataSet7.Incapacidades' table. You can move, or remove it, as needed.
        Me.IncapacidadesTableAdapter.Fill(Me.CServiciosDataSet7.Incapacidades)
        'TODO: This line of code loads data into the 'CServiciosDataSet7.EventosPersonal' table. You can move, or remove it, as needed.
        Me.EventosPersonalTableAdapter.Fill(Me.CServiciosDataSet7.EventosPersonal)
        'TODO: This line of code loads data into the 'CServiciosDataSet7.Estatus' table. You can move, or remove it, as needed.
        Me.EstatusTableAdapter.Fill(Me.CServiciosDataSet7.Estatus)
        'TODO: This line of code loads data into the 'CServiciosDataset19.Empleados' table. You can move, or remove it, as needed.
        Me.EmpleadosTableAdapter1.Fill(Me.CServiciosDataSet19.Empleados)
        'TODO: This line of code loads data into the 'CServiciosDataSet7.Departamentos' table. You can move, or remove it, as needed.
        Me.DepartamentosTableAdapter.Fill(Me.CServiciosDataSet7.Departamentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet20.Escolaridad' Puede moverla o quitarla según sea necesario.
        Me.EscolaridadTableAdapter.Fill(Me.CServiciosDataSet20.Escolaridad)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet19.Empleados' Puede moverla o quitarla según sea necesario.
        'Me.EmpleadosTableAdapter1.Fill(Me.CServiciosDataSet19.Empleados)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet19.Categorias' Puede moverla o quitarla según sea necesario.
        Me.CategoriasTableAdapter.Fill(Me.CServiciosDataSet19.Categorias)

        Panel2.Enabled = False
        Panel1.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = False
        ToolStripButton2.Enabled = True

        ' Llena el Combo con los Departamentos
        ComboBox3.Items.Clear()
        For Me.i = 0 To CServiciosDataSet7.Departamentos.Rows.Count - 1
            _Descripcion = CServiciosDataSet7.Departamentos.Rows(i).Item("Descripcion")
            ComboBox3.Items.Add(_Descripcion)
        Next


        ' Llena el Combo con la tabla de Estatus
        ComboBox2.Items.Clear()
        For Me.i = 0 To CServiciosDataSet7.Estatus.Rows.Count - 1
            _Descripcion = CServiciosDataSet7.Estatus.Rows(i).Item("Descripcion")
            ComboBox2.Items.Add(_Descripcion)
        Next

        ' Llena el Combo con la tabla de Sexo
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("MASCULINO")
        ComboBox1.Items.Add("FEMENINO")

        ' llena al Combo con el tipo de Empleado
        ComboBox4.Items.Clear()
        ComboBox4.Items.Add("SINDICALIZADO")
        ComboBox4.Items.Add("CONFIANZA")


        ' Muestra los datos
        DataGridView1.Rows.Clear()
        For Me.i = 0 To CServiciosDataset19.Empleados.Rows.Count - 1
            _NumEmpleado = CServiciosDataset19.Empleados.Rows(i).Item("NumEmpleado")
            _Paterno = CServiciosDataset19.Empleados.Rows(i).Item("Paterno")
            _Materno = CServiciosDataset19.Empleados.Rows(i).Item("Materno")
            _Nombre = CServiciosDataset19.Empleados.Rows(i).Item("Nombres")

            DataGridView1.Rows.Add(_NumEmpleado, _Paterno, _Materno, _Nombre)

        Next

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        TextBox22.Text = ""
        TextBox23.Text = ""

        TextBox1.Focus()
        ToolStripButton5.Enabled = True
        ToolStripButton2.Enabled = False
        TextBox10.Text = "ALTA"
        TextBox11.Enabled = False
        TextBox14.Enabled = False
        TextBox15.Enabled = False
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        ' Guarda el registro
        Dim _encuentra As Boolean
        _encuentra = False
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = UCase(checacontrol.Text)
                If checacontrol.Text = "" Then
                    _encuentra = True
                End If
            End If
        Next

        If _encuentra Then
            MsgBox("Hay casillas en blanco. Debe proporcionar todos los datos")
            Exit Sub
        End If

        _NumEmpleado = TextBox1.Text
        If TextBox10.Text = "ALTA" Then
            _encuentra = False

            Dim _Seleccion = "NumEmpleado = " & "'" & _NumEmpleado & "'"
            _Rows_em = CServiciosDataset19.Empleados.Select(_Seleccion)

            For Me.i = 0 To _Rows_em.GetUpperBound(0)
                _encuentra = True
            Next

            If _encuentra Then
                MsgBox("Ese número de Empleado ya existe. Verifique")
                Exit Sub
            End If


            ' Da de alta el registro
            _fila = CServiciosDataset19.Empleados.NewRow
            _fila.Item("NumEmpleado") = _NumEmpleado

        End If

        If TextBox10.Text = "CAMBIO" Then
            _fila = CServiciosDataset19.Empleados.FindByNumEmpleado(_NumEmpleado)
        End If

        _fila.Item("Paterno") = TextBox2.Text
        _fila.Item("Materno") = TextBox3.Text
        _fila.Item("Nombres") = TextBox4.Text
        _fila.Item("Departamento") = TextBox15.Text
        _fila.Item("Horario") = TextBox18.Text
        _fila.Item("Calle") = TextBox5.Text
        _fila.Item("Colonia") = TextBox6.Text
        _fila.Item("Ciudad") = TextBox7.Text
        _fila.Item("Estado") = TextBox8.Text
        _fila.Item("CP") = TextBox17.Text
        _fila.Item("Telefono") = TextBox9.Text
        _fila.Item("Celular") = ""
        _fila.Item("Email") = ""
        _fila.Item("Sexo") = TextBox11.Text
        _fila.Item("Referencia") = ComboBox4.Text
        _fila.Item("Comentarios") = ""
        _fila.Item("RFC") = TextBox12.Text
        _fila.Item("FechaIngreso") = TextBox13.Text
        _fila.Item("FechaBaja") = ""
        _fila.Item("MotivoBaja") = ""
        _fila.Item("FechaRegistro") = CStr(Today)
        _fila.Item("HoraRegistro") = TimeString
        _fila.Item("Estatus") = TextBox14.Text
        _fila.Item("NumAfiliacion") = TextBox16.Text
        _fila.Item("Celular") = TextBox23.Text
        _fila.Item("Email") = TextBox22.Text
        _fila.Item("FechaBaja") = TextBox20.Text
        _fila.Item("MotivoBaja") = TextBox21.Text
        _fila.Item("Categoria") = TextBox19.Text
        _fila.Item("Escolaridad") = TextBox24.Text

        If TextBox10.Text = "ALTA" Then
            CServiciosDataSet19.Empleados.Rows.Add(_fila)
        End If

        Me.Validate()
        EmpleadosBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet19)
        Me.EmpleadosTableAdapter1.Fill(CServiciosDataSet19.Empleados)

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next

        Button1_Click(sender, e)
        TextBox1.Enabled = True
        Panel1.Enabled = False
        MsgBox("Registro guardado correctamente")

    End Sub

    Private Sub ComboBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox11.Text = Mid(ComboBox1.Text, 1, 1)
    End Sub

    Private Sub ComboBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.Click

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet7.Estatus.Rows.Count - 1
            If CServiciosDataSet7.Estatus.Rows(i).Item("Descripcion") = ComboBox2.Text Then
                TextBox14.Text = CServiciosDataSet7.Estatus.Rows(i).Item("Estatus")
            End If
        Next
        If ComboBox2.Text = "ACTIVO" Then
            ComboBox2.BackColor = Color.LightGreen
            TextBox14.BackColor = Color.LightGreen
        Else
            If ComboBox2.Text = "BAJA" Then
                ComboBox2.BackColor = Color.Red
                TextBox14.BackColor = Color.Red
            Else
                ComboBox2.BackColor = Color.White
                TextBox14.BackColor = Color.White
            End If
        End If


    End Sub

    Private Sub ComboBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.Click

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet7.Departamentos.Rows.Count - 1
            If CServiciosDataSet7.Departamentos.Rows(i).Item("Descripcion") = ComboBox3.Text Then
                TextBox15.Text = CServiciosDataSet7.Departamentos.Rows(i).Item("Depatamento")
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value <> "" Then
            _Renglon.DefaultCellStyle.BackColor = Color.Gold

            ToolStripButton3.Enabled = True
            ToolStripButton2.Enabled = True
            ToolStripButton5.Enabled = False
            ToolStripButton11.Enabled = True

            ' Muestra los datos
            _NumEmpleado = _Renglon.Cells(0).Value
            _fila = CServiciosDataset19.Empleados.FindByNumEmpleado(_NumEmpleado)

            TextBox1.Text = _fila.Item("NumEmpleado")
            TextBox2.Text = _fila.Item("Paterno")
            TextBox3.Text = _fila.Item("Materno")
            TextBox4.Text = _fila.Item("Nombres")
            TextBox5.Text = _fila.Item("Calle")
            TextBox6.Text = _fila.Item("Colonia")
            TextBox7.Text = _fila.Item("Ciudad")
            TextBox8.Text = _fila.Item("Estado")
            TextBox17.Text = _fila.Item("CP")
            TextBox9.Text = _fila.Item("Telefono")
            TextBox11.Text = _fila.Item("Sexo")
            If TextBox11.Text = "M" Then
                ComboBox1.Text = "MASCULINO"
            End If
            If TextBox11.Text = "F" Then
                ComboBox1.Text = "FEMENINO"
            End If
            ComboBox4.Text = _fila.Item("Referencia")
            TextBox15.Text = _fila.Item("Departamento")
            _Busca = CServiciosDataSet7.Departamentos.FindByDepatamento(TextBox15.Text)
            ComboBox3.Text = _Busca.Item("Descripcion")
            TextBox18.Text = _fila.Item("Horario")
            RichTextBox1.Text = _fila.Item("Comentarios")
            TextBox12.Text = _fila.Item("RFC")
            TextBox13.Text = _fila.Item("FechaIngreso")
            TextBox20.Text = _fila.Item("FechaBaja")
            TextBox21.Text = _fila.Item("MotivoBaja")
            TextBox16.Text = _fila.Item("NumAfiliacion")
            TextBox14.Text = _fila.Item("Estatus")
            _Busca = CServiciosDataSet7.Estatus.FindByEstatus(TextBox14.Text)
            ComboBox2.Text = _Busca.Item("Descripcion")
            TextBox23.Text = _fila.Item("Celular")
            TextBox22.Text = _fila.Item("Email")
            TextBox19.Text = _fila.Item("Categoria")
            _Busca = CServiciosDataSet19.Categorias.FindByCategoria(TextBox19.Text)
            ComboBox5.Text = _Busca.Item("Descripcion")
            TextBox24.Text = _fila.Item("Escolaridad")
            _Busca = CServiciosDataSet20.Escolaridad.FindByEscolaridad(TextBox24.Text)
            ComboBox6.Text = _Busca.Item("Descripcion")

            DataGridView2.Rows.Clear()
            ToolStripButton7_Click(sender, e)
            ToolStripButton8_Click(sender, e)


        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Panel1.Enabled = True
        TextBox1.Enabled = False
        TextBox12.Focus()
        TextBox10.Text = "CAMBIO"
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = True
    End Sub

    Private Sub DataGridView1_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub ToolStripButton11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton11.Click
        ToolStripButton5.Enabled = True
        ToolStripButton11.Enabled = False
        TabPage7.Focus()
        Panel2.Enabled = True
        TextBox10.Text = "CAMBIO"
    End Sub

    Private Sub ToolStripButton12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton12.Click
        Frm9NuevoPermiso.Show()
        ToolStripButton7_Click(sender, e)
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Me.PermisosTableAdapter.Fill(CServiciosDataSet7.Permisos)

        Dim _Seleccion As String = "NumEmpleado = " & "'" & TextBox1.Text & "'"
        Dim _Rows_pe() As DataRow = CServiciosDataSet7.Permisos.Select(_seleccion)
        Dim _NumEmpleado, _Permiso, _Fecha, _HoraInicio, _Motivo, _Atiende, _Autoriza, _Comentarios As String
        Dim _NumDias, _NumHoras, _NumMinutos As Integer


        DataGridView2.Rows.Clear()
        For Me.i = 0 To _Rows_pe.GetUpperBound(0)
            _NumEmpleado = _Rows_pe(i).Item("NumEmpleado")
            _Permiso = _Rows_pe(i).Item("Permiso")
            _Fecha = _Rows_pe(i).Item("FechaPermiso")
            _HoraInicio = _Rows_pe(i).Item("HoraRegistro")
            _Motivo = _Rows_pe(i).Item("Motivo")
            _Atiende = _Rows_pe(i).Item("Atiende")
            _Autoriza = _Rows_pe(i).Item("Autoriza")
            _Comentarios = _Rows_pe(i).Item("Comentarios")
            _NumDias = _Rows_pe(i).Item("NumDias")
            _NumHoras = _Rows_pe(i).Item("NumHoras")
            _NumMinutos = _Rows_pe(i).Item("NumMinutos")

            DataGridView2.Rows.Add(_NumEmpleado, _Permiso, _Fecha, _HoraInicio, _NumDias, _NumHoras, _NumMinutos, _Motivo, _Atiende, _Autoriza)

        Next

    End Sub

    Private Sub ToolStripButton14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton14.Click
        Frm2AltaIncapacidades.Show()
        Button1_Click(sender, e)
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


        ' Borra los datos de las cajas de texto
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
            If TypeOf checacontrol Is ComboBox Then
                checacontrol.Text = ""
            End If
        Next

        ' Borra los comentarios 
        RichTextBox1.Text = ""

        If RadioButton1.Checked = True Then
            ' **********************************************************************************************************************
            ' Busqueda por la Clave del Empleado
            _NumEmpleado = UCase(InputBox("Teclee el Número de Empelado a buscar"))
            _Rows_pa = CServiciosDataset19.Empleados.Select("NumEmpleado = " & "'" & _NumEmpleado & "'")

            _Encuentra = False
            For Me.i = 0 To _Rows_pa.GetUpperBound(0)
                _Encuentra = True
                _NumEmpleado = _Rows_pa(i).Item("NumEmpleado")
                _Paterno = _Rows_pa(i).Item("Paterno")
                _Materno = _Rows_pa(i).Item("Materno")
                _Nombres = _Rows_pa(i).Item("Nombres")

                DataGridView1.Rows.Add(_NumEmpleado, _Paterno, _Materno, _Nombres)

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

            For Me.i = 0 To CServiciosDataset19.Empleados.Rows.Count - 1
                _texto = Trim(CServiciosDataset19.Empleados.Rows(i).Item("Nombres")) & " " & Trim(CServiciosDataset19.Empleados.Rows(i).Item("Paterno")) & " " & Trim(CServiciosDataset19.Empleados.Rows(i).Item("Materno"))
                _NumEmpleado = CServiciosDataset19.Empleados.Rows(i).Item("NumEmpleado")
                _Paterno = CServiciosDataset19.Empleados.Rows(i).Item("Paterno")
                _Materno = CServiciosDataset19.Empleados.Rows(i).Item("Materno")
                _Nombres = CServiciosDataset19.Empleados.Rows(i).Item("Nombres")

                For x = 1 To Len(_texto)
                    _partenombre = Mid(_texto, x, _mide)

                    If _partenombre = _parte Then
                        DataGridView1.Rows.Add(_NumEmpleado, _Paterno, _Materno, _Nombres)
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
            ' Busca por TIPO DE EMPLEADO
            ' **********************************************************************************************************************
            _parte = InputBox("Teclee la TIPO DE EMPLEADO a buscar (SINDICALIZADO/CONFIANZA)")
            _parte = Trim(UCase(_parte))
            _mide = Len(_parte)
            If _mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataset19.Empleados.Rows.Count - 1
                _texto = Trim(UCase(CServiciosDataset19.Empleados.Rows(i).Item("Referencia")))
                _NumEmpleado = CServiciosDataset19.Empleados.Rows(i).Item("NumEmpleado")
                _Paterno = CServiciosDataset19.Empleados.Rows(i).Item("Paterno")
                _Materno = CServiciosDataset19.Empleados.Rows(i).Item("Materno")
                _Nombres = CServiciosDataset19.Empleados.Rows(i).Item("Nombres")


                For x = 1 To Len(_texto)
                    _partenombre = Mid(_texto, x, _mide)

                    If _partenombre = _parte Then
                        DataGridView1.Rows.Add(_NumEmpleado, _Paterno, _Materno, _Nombres)
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
            ' Busca por DEPARTAMENTO
            ' **********************************************************************************************************************
            _parte = InputBox("Teclee el DEPARTAMENTO a buscar")
            _parte = Trim(UCase(_parte))
            _mide = Len(_parte)
            If _mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet7.Departamentos.Rows.Count - 1
                _texto = Trim(UCase(CServiciosDataSet7.Departamentos.Rows(i).Item("Descripcion")))
                _Departamento = CServiciosDataSet7.Departamentos.Rows(i).Item("Depatamento")

                For x = 1 To Len(_texto)
                    _partenombre = Mid(_texto, x, _mide)

                    If _partenombre = _parte Then
                        _Rows_em = CServiciosDataset19.Empleados.Select("Departamento = " & "'" & _Departamento & "'")
                        For Me.n = 0 To _Rows_em.GetUpperBound(0)
                            _NumEmpleado = _Rows_em(n).Item("NumEmpleado")
                            _Paterno = _Rows_em(n).Item("Paterno")
                            _Materno = _Rows_em(n).Item("Materno")
                            _Nombres = _Rows_em(n).Item("Nombres")

                            DataGridView1.Rows.Add(_NumEmpleado, _Paterno, _Materno, _Nombres)
                        Next

                    End If
                Next

            Next

        End If




        If RadioButton5.Checked Then
            ' Busca por SEXO
            ' **********************************************************************************************************************
            _parte = InputBox("Teclee F o M (Femenino o Masculino)  a buscar")
            _parte = UCase(_parte)
            If _parte <> "F" And _parte <> "M" Then
                msg = "Debe teclear F ó M solamente"
                style = MsgBoxStyle.Information
                MsgBox(msg, style)
                Exit Sub
            End If

            _Rows_em = CServiciosDataset19.Empleados.Select("Sexo = " & "'" & _parte & "'")
            For Me.n = 0 To _Rows_em.GetUpperBound(0)
                _NumEmpleado = _Rows_em(n).Item("NumEmpleado")
                _Paterno = _Rows_em(n).Item("Paterno")
                _Materno = _Rows_em(n).Item("Materno")
                _Nombres = _Rows_em(n).Item("Nombres")

                DataGridView1.Rows.Add(_NumEmpleado, _Paterno, _Materno, _Nombres)
            Next
        End If




        If RadioButton6.Checked Then
            ' Busca por CATEGORIA O PUESTO del empleado
            ' **********************************************************************************************************************
            _parte = InputBox("Teclee la CATEGORIA O PUESTO a buscar")
            _parte = Trim(UCase(_parte))
            _mide = Len(_parte)
            If _mide = 0 Then
                Exit Sub
            End If

            For Me.i = 0 To CServiciosDataSet19.Categorias.Rows.Count - 1
                _texto = Trim(UCase(CServiciosDataSet19.Categorias.Rows(i).Item("Descripcion")))
                _Categoria = CServiciosDataSet19.Categorias.Rows(i).Item("Categoria")

                For x = 1 To Len(_texto)
                    _partenombre = Mid(_texto, x, _mide)

                    If _partenombre = _parte Then
                        _Rows_em = CServiciosDataSet19.Empleados.Select("Categoria = " & "'" & _Categoria & "'")
                        For Me.n = 0 To _Rows_em.GetUpperBound(0)
                            _NumEmpleado = _Rows_em(n).Item("NumEmpleado")
                            _Paterno = _Rows_em(n).Item("Paterno")
                            _Materno = _Rows_em(n).Item("Materno")
                            _Nombres = _Rows_em(n).Item("Nombres")

                            DataGridView1.Rows.Add(_NumEmpleado, _Paterno, _Materno, _Nombres)
                        Next

                    End If
                Next

            Next

        End If







        If RadioButton7.Checked Then
            ' Busca TODOS
            ' **********************************************************************************************************************
            For Me.i = 0 To CServiciosDataSet19.Empleados.Rows.Count - 1


                _NumEmpleado = CServiciosDataSet19.Empleados.Rows(i).Item("NumEmpleado")
                _Paterno = CServiciosDataSet19.Empleados.Rows(i).Item("Paterno")
                _Materno = CServiciosDataSet19.Empleados.Rows(i).Item("Materno")
                _Nombres = CServiciosDataSet19.Empleados.Rows(i).Item("Nombres")

                DataGridView1.Rows.Add(_NumEmpleado, _Paterno, _Materno, _Nombres)

            Next

        End If
    End Sub

    Private Sub ComboBox5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.Click

    End Sub

    Private Sub ComboBox5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.GotFocus
        ComboBox5.Items.Clear()
        For Me.i = 0 To CServiciosDataSet19.Categorias.Rows.Count - 1
            ComboBox5.Items.Add(CServiciosDataSet19.Categorias.Rows(i).Item("Descripcion"))
        Next
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet19.Categorias.Rows.Count - 1
            If CServiciosDataSet19.Categorias.Rows(i).Item("Descripcion") = ComboBox5.Text Then
                TextBox19.Text = CServiciosDataSet19.Categorias.Rows(i).Item("Categoria")
            End If
        Next
    End Sub

    Private Sub ComboBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.GotFocus
        ComboBox6.Items.Clear()
        For Me.i = 0 To CServiciosDataSet20.Escolaridad.Rows.Count - 1
            ComboBox6.Items.Add(CServiciosDataSet20.Escolaridad.Rows(i).Item("Descripcion"))
        Next

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet20.Escolaridad.Rows.Count - 1
            If CServiciosDataSet20.Escolaridad.Rows(i).Item("Descripcion") = ComboBox6.Text Then
                TextBox24.Text = CServiciosDataSet20.Escolaridad.Rows(i).Item("Escolaridad")
            End If
        Next

    End Sub

    Private Sub ToolStripButton15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton15.Click
        If DataGridView2.Rows.Count = 0 Then
            msg = " Debe seleccionar un PERMISO para imprimir el Formato "
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If
        Dim _Permiso As String
        Dim _Fecha As Date
        Dim _TipoPermiso, _Motivo, _HoraInicio As String
        Dim _NumDias, _NumHoras, _NumMinutos As Integer
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Formato Permiso.xlsx"

        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de excel
        _Renglon = DataGridView2.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _NumEmpleado = _Renglon.Cells(0).Value

        _Busca = CServiciosDataSet19.Empleados.FindByNumEmpleado(_NumEmpleado)
        _Nombre = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")

        _Permiso = _Renglon.Cells(1).Value
        _Fecha = _Renglon.Cells(2).Value
        If Mid(_Permiso, 2, 1) = "E" Then
            _TipoPermiso = "ENTRADA"
        Else
            _TipoPermiso = "SALIDA"
        End If
        _HoraInicio = _Renglon.Cells(3).Value
        _NumDias = _Renglon.Cells(4).Value
        _NumHoras = _Renglon.Cells(5).Value
        _NumMinutos = _Renglon.Cells(6).Value
        _Motivo = _Renglon.Cells(7).Value

        ' Imprime el formato
        m_Excel.Worksheets("PERMISO").Cells(3, 9).value = _Permiso

        _Nombre = "(" & _NumEmpleado & ")" & " " & _Nombre
        m_Excel.Worksheets("PERMISO").Cells(8, 3).value = _Nombre
        m_Excel.Worksheets("PERMISO").Cells(9, 6).value = _Fecha
        m_Excel.Worksheets("PERMISO").Cells(10, 6).value = _TipoPermiso
        m_Excel.Worksheets("PERMISO").Cells(11, 6).value = _HoraInicio
        m_Excel.Worksheets("PERMISO").Cells(12, 6).value = _NumMinutos
        m_Excel.Worksheets("PERMISO").Cells(13, 6).value = _NumHoras
        m_Excel.Worksheets("PERMISO").Cells(14, 6).value = _NumDias
        m_Excel.Worksheets("PERMISO").Cells(20, 3).value = _Motivo

        m_Excel.Visible = True

    End Sub

    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click
        ' Presenta las Incapacidades Médicas
        Me.IncapacidadesTableAdapter.Fill(Me.CServiciosDataSet7.Incapacidades)
        Dim _Rows_Im() As DataRow = CServiciosDataSet7.Incapacidades.Select("NumEmpleado = " & "'" & TextBox1.Text & "'")
        Dim _FolioIncapacidad, _Clinica, _Medico, _Categoria, _Diagnostico, _Referencia As String
        Dim _FechaInicio, _FechaIncapacidad As Date
        Dim _NumDias As Integer
        Dim _NumAfiliacion As String


        DataGridView3.Rows.Clear()
        For Me.i = 0 To _Rows_Im.GetUpperBound(0)
            _FolioIncapacidad = _Rows_Im(i).Item("FolioIncapacidad")
            _FechaIncapacidad = _Rows_Im(i).Item("FechaIncapacidad")
            _NumDias = _Rows_Im(i).Item("NumDias")
            _FechaInicio = _Rows_Im(i).Item("FechaInicio")
            _NumAfiliacion = _Rows_Im(i).Item("NumAfiliacion")
            _Clinica = _Rows_Im(i).Item("clinica")
            _Medico = _Rows_Im(i).Item("Medico")
            _Categoria = _Rows_Im(i).Item("Categoria")
            _Diagnostico = _Rows_Im(i).Item("Diagnostico")
            _Referencia = _Rows_Im(i).Item("Referencia")

            DataGridView3.Rows.Add(_NumEmpleado, _FolioIncapacidad, _FechaIncapacidad, _NumDias, _FechaInicio, _NumAfiliacion, _Clinica, _Medico, _Categoria, _Diagnostico, _Referencia)

        Next

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellLeave
        _Renglon = DataGridView2.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub
End Class