Public Class Frm19ConsultadeCitas
    Public i, _Cuantos, _Atendidas, _Porcentaje As Integer
    Dim style As MsgBoxStyle
    Dim _Fecha As Date
    Public _Busca As DataRow
    Public _Renglon As DataGridViewRow
    Public _Cita As String
    Public _Checacontrol As Control
    Dim _FechaC, _Centro, _HoraCita As String
    Dim _BuscaPaciente, _BuscaMedico, _BuscaClinica As DataRow
    Public msg, _Medico, _Paciente, _NombreMedico, _NombrePaciente, _Referencia, _Clinica, _NombreClinica As String

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        MonthCalendar1.Visible = True
        DataGridView1.Rows.Clear()
        ToolStripTextBox1.Text = ""
        ToolStripTextBox7.Text = ""
        ToolStripTextBox11.Text = ""
        ToolStripLabel7.Text = ""
    End Sub

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        ToolStripTextBox1.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False
        DataGridView1.Enabled = False
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        TextBox12.Text = ToolStripButton2.Name
        DataGridView1.Enabled = True
        Dim _Asistencia As Boolean

        If ToolStripTextBox1.Text = "" Then
            msg = "Debe seleccionar una fecha para la Consulta"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If
        Dim _FechaCita As String = Trim(ToolStripTextBox1.Text)
        Dim _Find As Boolean

        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)


        _Find = False
        DataGridView1.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If Me.CServiciosDataSet4.Citas.Rows(i).Item("Estatus") <> "CANCELADO" Then
                If Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita") = _FechaCita Then
                    _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")

                    _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                    _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                    _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                    _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                    _NombreMedico = _BuscaMedico.Item("Nombre")
                    _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                    _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                    _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                    _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                    DataGridView1.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Referencia, _Asistencia, _HoraCita, _Cita)
                    _Find = True
                End If
            End If
        Next

        If _Find = False Then
            msg = "No existen Citas agendadas para esa fecha"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Atendidas = 0
        For Me.i = 0 To DataGridView1.Rows.Count - 1
            _Asistencia = DataGridView1.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next
        If DataGridView1.Rows.Count - 1 > 0 Then
            ToolStripTextBox7.Text = DataGridView1.Rows.Count - 1
            ToolStripTextBox11.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox7.Text) * 100
        Else
            _Porcentaje = 0
        End If


        ToolStripLabel7.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "
    End Sub

    Private Sub Frm19ConsultadeCintas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
   
        Button2_Click(sender, e)

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If ToolStripComboBox1.Text = "" Then
            msg = "Debe seleccionar un Personal Médico para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        TextBox12.Text = ToolStripButton4.Name
        DataGridView2.Enabled = True
        Dim _Asistencia As Boolean

        DataGridView2.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende") = ToolStripTextBox2.Text Then
                _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")

                _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                _NombreMedico = _BuscaMedico.Item("Nombre")
                _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                _Fecha = (Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"))
                _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                DataGridView2.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _Asistencia, _HoraCita, _Cita)
            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView2.Rows.Count - 1
            _Asistencia = DataGridView2.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next
        ToolStripTextBox8.Text = DataGridView2.Rows.Count - 1
        If CInt(ToolStripTextBox8.Text) > 0 Then
            ToolStripTextBox12.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox8.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel9.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "


    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            If Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre") = ToolStripComboBox1.Text Then
                ToolStripTextBox2.Text = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Medico")
            End If
        Next
        DataGridView2.Rows.Clear()
        ToolStripTextBox8.Text = ""
        ToolStripTextBox12.Text = ""
        ToolStripLabel9.Text = ""
    End Sub

    Private Sub ToolStripComboBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.Click

    End Sub

    Private Sub ToolStripComboBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.GotFocus
        ToolStripComboBox2.Items.Clear()

        For Me.i = 0 To Me.CServiciosDataSet8.Centros.Rows.Count - 1
            _Centro = Me.CServiciosDataSet8.Centros.Rows(i).Item("Descripcion")
            ToolStripComboBox2.Items.Add(_Centro)
        Next

    End Sub

    Private Sub ToolStripComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged

        For Me.i = 0 To Me.CServiciosDataSet8.Centros.Rows.Count - 1
            If Me.CServiciosDataSet8.Centros.Rows(i).Item("Descripcion") = ToolStripComboBox2.Text Then
                ToolStripTextBox3.Text = Me.CServiciosDataSet8.Centros.Rows(i).Item("Centro")
            End If
        Next
        DataGridView3.Rows.Clear()
        ToolStripTextBox9.Text = ""
        ToolStripTextBox13.Text = ""
        ToolStripLabel11.Text = ""


    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        If ToolStripComboBox2.Text = "" Then
            msg = "Debe seleccionar una Clínica para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If
        TextBox12.Text = ToolStripButton5.Name

        DataGridView3.Enabled = True
        Dim _Asistencia As Boolean
        DataGridView3.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If Me.CServiciosDataSet4.Citas.Rows(i).Item("Centro") = ToolStripTextBox3.Text Then
                _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                _NombreMedico = _BuscaMedico.Item("Nombre")
                _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                _Fecha = (Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"))
                _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                DataGridView3.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _Asistencia, _HoraCita, _Cita)
            End If

        Next
        _Atendidas = 0

        For Me.i = 0 To DataGridView3.Rows.Count - 1
            _Asistencia = DataGridView3.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView3.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox9.Text = DataGridView3.Rows.Count - 1
        If CInt(ToolStripTextBox9.Text) > 0 Then
            ToolStripTextBox13.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox9.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel11.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        ' Muestra la Información
        If ToolStripComboBox3.Text = "" Then
            msg = "Debe seleccionar un Paciente para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        TextBox12.Text = ToolStripButton6.Name

        DataGridView4.Enabled = True
        Dim _Asistencia As Boolean
        DataGridView4.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente") = ToolStripTextBox4.Text Then
                _Clinica = Me.CServiciosDataSet4.Citas.Rows(i).Item("Centro")
                _BuscaClinica = Me.CServiciosDataSet8.Centros.FindByCentro(_Clinica)
                _NombreClinica = _BuscaClinica.Item("Descripcion")
                _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                _NombreMedico = _BuscaMedico.Item("Nombre")
                _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                _Fecha = (Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"))
                _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                DataGridView4.Rows.Add(_Clinica, _NombreClinica, _Medico, _NombreMedico, _Fecha, _Asistencia, _HoraCita, _Cita)
            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView4.Rows.Count - 1
            _Asistencia = DataGridView4.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView4.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox10.Text = DataGridView4.Rows.Count - 1

        If CInt(ToolStripTextBox10.Text) > 0 Then
            ToolStripTextBox14.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox10.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel13.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "

    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()

    End Sub

    Private Sub ToolStripComboBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox3.Click

    End Sub

    Private Sub ToolStripComboBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox3.GotFocus
      

    End Sub

    Private Sub ToolStripComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox3.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet2.Pacientes.Rows.Count - 1
            If Trim(Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")) = Trim(Me.ToolStripComboBox3.Text) Then
                Me.ToolStripTextBox4.Text = Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
            End If
        Next
        DataGridView4.Rows.Clear()
        ToolStripTextBox10.Text = ""
        ToolStripTextBox14.Text = ""
        ToolStripLabel13.Text = ""
    End Sub

    Private Sub ToolStripButton7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        DataGridView4.Enabled = True
        Dim _Asistencia As Boolean

        TextBox12.Text = ToolStripButton7.Name

        DataGridView5.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1

            _Clinica = Me.CServiciosDataSet4.Citas.Rows(i).Item("Centro")
            _BuscaClinica = Me.CServiciosDataSet8.Centros.FindByCentro(_Clinica)
            _NombreClinica = _BuscaClinica.Item("Descripcion")
            _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
            _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
            _NombreMedico = _BuscaMedico.Item("Nombre")
            _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
            _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
            _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
            _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
            _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
            _Fecha = (Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"))
            _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
            _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

            DataGridView5.Rows.Add(_Clinica, _NombreClinica, _Medico, _NombreMedico, _Fecha, _Paciente, _NombrePaciente, _Asistencia, _HoraCita, _Cita)

        Next

        _Atendidas = 0
        For Me.i = 0 To DataGridView5.Rows.Count - 1
            _Asistencia = DataGridView5.Rows(i).Cells(7).Value
            If _Asistencia Then
                DataGridView5.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox15.Text = DataGridView5.Rows.Count - 1

        If CInt(ToolStripTextBox15.Text) > 0 Then
            ToolStripTextBox16.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox15.Text) * 100
        Else
            _Porcentaje = 0
        End If
        ToolStripLabel16.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "


    End Sub

    Private Sub BindingNavigatorMoveFirstItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripComboBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox4.Click

    End Sub

    Private Sub ToolStripComboBox4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox4.GotFocus
        ToolStripComboBox4.Items.Clear()
        ToolStripComboBox4.Items.Add("ENERO")
        ToolStripComboBox4.Items.Add("FEBRERO")
        ToolStripComboBox4.Items.Add("MARZO")
        ToolStripComboBox4.Items.Add("ABRIL")
        ToolStripComboBox4.Items.Add("MAYO")
        ToolStripComboBox4.Items.Add("JUNIO")
        ToolStripComboBox4.Items.Add("JULIO")
        ToolStripComboBox4.Items.Add("AGOSTO")
        ToolStripComboBox4.Items.Add("SEPTIEMBRE")
        ToolStripComboBox4.Items.Add("OCTUBRE")
        ToolStripComboBox4.Items.Add("NOVIEMBRE")
        ToolStripComboBox4.Items.Add("DICIEMBRE")

    End Sub

    Private Sub ToolStripComboBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripComboBox5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        ToolStripComboBox5.Items.Clear()
        For x = 2014 To 2030
            ToolStripComboBox5.Items.Add(x)
        Next
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        DataGridView6.Enabled = True
        Dim _Asistencia As Boolean
        Dim _FechaCita, _Mes, _Ejercicio, _Mescombo As String

        If ToolStripComboBox4.Text = "" Then
            msg = "Debe seleccionar un Mes para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        If ToolStripComboBox5.Text = "" Then
            msg = "Debe seleccionar un Ejercicio para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        DataGridView6.Rows.Clear()
        i = 0
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            _FechaCita = CStr(Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"))
            _Mes = Mid(_FechaCita, 4, 2)
            _Ejercicio = Mid(_FechaCita, 7, 4)
            _Mescombo = CStr(ToolStripComboBox4.SelectedIndex + 1)
            _Mescombo = "0" + _Mescombo
            If Len(_Mescombo) = 3 Then
                _Mescombo = Mid(_Mescombo, 2, 2)
            End If

            If _Mes = _Mescombo And _Ejercicio = ToolStripComboBox5.Text Then


                _Clinica = Me.CServiciosDataSet4.Citas.Rows(i).Item("Centro")
                _BuscaClinica = Me.CServiciosDataSet8.Centros.FindByCentro(_Clinica)
                _NombreClinica = _BuscaClinica.Item("Descripcion")
                _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                _NombreMedico = _BuscaMedico.Item("Nombre")
                _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                _Fecha = (Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"))
                _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                DataGridView6.Rows.Add(_Clinica, _NombreClinica, _Medico, _NombreMedico, _Fecha, _Paciente, _NombrePaciente, _Asistencia, _HoraCita, _Cita)

            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView6.Rows.Count - 1
            _Asistencia = DataGridView6.Rows(i).Cells(7).Value
            If _Asistencia Then
                DataGridView6.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox17.Text = DataGridView6.Rows.Count - 1

        If CInt(ToolStripTextBox17.Text) > 0 Then
            ToolStripTextBox18.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox17.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel19.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "



    End Sub

    Private Sub ToolStripComboBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox6.Click

    End Sub

    Private Sub ToolStripComboBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox6.GotFocus
        ToolStripComboBox6.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            _NombreMedico = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre")
            ToolStripComboBox6.Items.Add(_NombreMedico)
        Next
    End Sub

    Private Sub ToolStripComboBox6_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox6.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            If Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre") = ToolStripComboBox6.Text Then
                ToolStripTextBox6.Text = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Medico")
            End If
        Next

        DataGridView7.Rows.Clear()
        ToolStripTextBox5.Text = ""
        ToolStripTextBox19.Text = ""
        ToolStripTextBox20.Text = ""
        ToolStripLabel22.Text = ""

    End Sub

    Private Sub ToolStripButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton10.Click
        MonthCalendar2.Visible = True
    End Sub

    Private Sub MonthCalendar2_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateChanged

    End Sub

    Private Sub MonthCalendar2_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateSelected
        ToolStripTextBox5.Text = MonthCalendar2.SelectionRange.Start
        MonthCalendar2.Visible = False
    End Sub

    Private Sub ToolStripButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton9.Click
        DataGridView7.Enabled = True
        Dim _Asistencia As Boolean

        style = MsgBoxStyle.Information
        If ToolStripComboBox6.Text = "" Then
            msg = "Debe seleccionar un Personal Médico para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If

        If ToolStripTextBox5.Text = "" Then
            msg = "Debe seleccionar una Fecha para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If


        DataGridView7.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita") = ToolStripTextBox5.Text Then
                If Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende") = ToolStripTextBox6.Text Then
                    _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                    _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                    _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                    _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                    _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                    _NombreMedico = _BuscaMedico.Item("Nombre")
                    _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                    _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                    _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                    _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                    DataGridView7.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Referencia, _Asistencia, _HoraCita, _Cita)
                End If
            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView7.Rows.Count - 1
            _Asistencia = DataGridView7.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView7.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox19.Text = DataGridView7.Rows.Count - 1
        If CInt(ToolStripTextBox19.Text) > 0 Then
            ToolStripTextBox20.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox19.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel22.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If


        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya fué procesada"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = ToolStripTextBox1.Text
        ComboBox1.Text = ""
        Button1.Enabled = True
        Button1.Visible = True
        Button6.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripLabel9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel9.Click

    End Sub

    Private Sub ToolStripComboBox4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox4.SelectedIndexChanged
        ToolStripComboBox5.Text = ""
        ToolStripTextBox17.Text = ""
        ToolStripTextBox18.Text = ""
        ToolStripLabel19.Text = ""
    End Sub

    Private Sub BindingNavigator4_RefreshItems(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigator4.RefreshItems

    End Sub

    Private Sub ToolStripButton11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton11.Click
        style = MsgBoxStyle.Information
        If ToolStripComboBox7.Text = "" Then
            msg = "Debe seleccionar un Personal Médico para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If
        If ToolStripComboBox8.Text = "" Then
            msg = "Debe seleccionar un Mes para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If
        If ToolStripComboBox9.Text = "" Then
            msg = "Debe seleccionar un Ejercicio para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If
        Dim _Asistencia As Boolean
        Dim _MesCombo, _EjercicioCombo As String
        _MesCombo = CStr(ToolStripComboBox8.SelectedIndex + 1)
        _MesCombo = "0" + _MesCombo
        If Len(_MesCombo) = 3 Then
            _MesCombo = Mid(_MesCombo, 2, 2)
        End If
        _EjercicioCombo = CStr(ToolStripComboBox9.Text)

        Dim _Atiende As String = ToolStripTextBox23.Text
        Dim _Rows_Ci() As DataRow = CServiciosDataSet4.Citas.Select("Atiende = " & "'" & _Atiende & "'")

        DataGridView8.Rows.Clear()
        For Me.i = 0 To _Rows_Ci.GetUpperBound(0)
            If Mid(_Rows_Ci(i).Item("FechaCita"), 7, 4) = _EjercicioCombo Then
                '  MsgBox("Pasa el Ejercicio")
                '  MsgBox("Mes combo " & _MesCombo)
                '  MsgBox("Mes de la tabla " & _Rows_Ci(i).Item("FechaCita"))
                If CInt(Mid(_Rows_Ci(i).Item("FechaCita"), 4, 2)) = CInt(_MesCombo) Then
                    '     MsgBox("Pasa el mes")
                    If _Rows_Ci(i).Item("Atiende") = ToolStripTextBox23.Text Then
                        '        MsgBox("Pasa el Medico")
                        _Paciente = _Rows_Ci(i).Item("Paciente")
                        _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                        _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                        _Medico = Trim(_Rows_Ci(i).Item("Atiende"))
                        _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                        _NombreMedico = _BuscaMedico.Item("Nombre")
                        _Referencia = _Rows_Ci(i).Item("Referencia")
                        _Asistencia = _Rows_Ci(i).Item("Asiste")
                        _HoraCita = _Rows_Ci(i).Item("HoraCita")
                        _Fecha = _Rows_Ci(i).Item("FechaCita")
                        _Cita = _Rows_Ci(i).Item("Cita")

                        DataGridView8.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _Asistencia, _HoraCita, _Cita)
                    End If
                End If
            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView8.Rows.Count - 1
            _Asistencia = DataGridView8.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView8.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox21.Text = DataGridView8.Rows.Count - 1
        If CInt(ToolStripTextBox21.Text) > 0 Then
            ToolStripTextBox22.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox21.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel25.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "

    End Sub

    Private Sub ToolStripComboBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox7.Click

    End Sub

    Private Sub ToolStripComboBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox7.GotFocus
        ToolStripComboBox7.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            _NombreMedico = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre")
            ToolStripComboBox7.Items.Add(_NombreMedico)
        Next
    End Sub

    Private Sub ToolStripComboBox7_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox7.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            If Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre") = ToolStripComboBox7.Text Then
                ToolStripTextBox23.Text = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Medico")
                Exit For
            End If
        Next

        ToolStripComboBox8.Text = ""
        ToolStripComboBox9.Text = ""
        ToolStripTextBox21.Text = ""
        ToolStripTextBox22.Text = ""
        ToolStripLabel25.Text = ""
        DataGridView8.Rows.Clear()
    End Sub

    Private Sub ToolStripComboBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox8.Click

    End Sub

    Private Sub ToolStripComboBox8_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox8.GotFocus
        ToolStripComboBox8.Items.Clear()
        ToolStripComboBox8.Items.Add("ENERO")
        ToolStripComboBox8.Items.Add("FEBRERO")
        ToolStripComboBox8.Items.Add("MARZO")
        ToolStripComboBox8.Items.Add("ABRIL")
        ToolStripComboBox8.Items.Add("MAYO")
        ToolStripComboBox8.Items.Add("JUNIO")
        ToolStripComboBox8.Items.Add("JULIO")
        ToolStripComboBox8.Items.Add("AGOSTO")
        ToolStripComboBox8.Items.Add("SEPTIEMBRE")
        ToolStripComboBox8.Items.Add("OCTUBRE")
        ToolStripComboBox8.Items.Add("NOVIEMBRE")
        ToolStripComboBox8.Items.Add("DICIEMBRE")
    End Sub

    Private Sub ToolStripComboBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox9.Click

    End Sub

    Private Sub ToolStripComboBox9_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox9.GotFocus
        ToolStripComboBox9.Items.Clear()
        For x = 2014 To 2030
            ToolStripComboBox9.Items.Add(x)
        Next
    End Sub

    Private Sub ToolStripComboBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox10.Click

    End Sub

    Private Sub ToolStripComboBox10_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox10.GotFocus
        ToolStripComboBox10.Items.Clear()
        For x = 2014 To 2030
            ToolStripComboBox10.Items.Add(x)
        Next
    End Sub

    Private Sub ToolStripButton12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton12.Click

        style = MsgBoxStyle.Information
        If ToolStripComboBox10.Text = "" Then
            msg = "Debe seleccionar un Ejercicio para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If

        Dim _Asistencia As Boolean
        Dim _EjercicioCombo As String
        _EjercicioCombo = CStr(ToolStripComboBox10.Text)

        DataGridView9.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If Mid(Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"), 7, 4) = _EjercicioCombo Then

                _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                _NombreMedico = _BuscaMedico.Item("Nombre")
                _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                _Fecha = Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita")
                _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                DataGridView9.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _Asistencia, _HoraCita, _Cita)

            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView9.Rows.Count - 1
            _Asistencia = DataGridView9.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView9.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox24.Text = DataGridView9.Rows.Count - 1
        If CInt(ToolStripTextBox24.Text) > 0 Then
            ToolStripTextBox25.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox24.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel28.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "



    End Sub

    Private Sub ToolStripComboBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox11.Click

    End Sub

    Private Sub ToolStripComboBox11_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox11.GotFocus
        ToolStripComboBox11.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            _NombreMedico = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre")
            ToolStripComboBox11.Items.Add(_NombreMedico)
        Next
    End Sub

    Private Sub ToolStripComboBox12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox12.Click

    End Sub

    Private Sub ToolStripComboBox12_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox12.GotFocus
        ToolStripComboBox12.Items.Clear()

        For x = 2014 To 2030
            ToolStripComboBox12.Items.Add(x)
        Next

    End Sub

    Private Sub ToolStripComboBox11_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox11.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            If Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre") = ToolStripComboBox11.Text Then
                ToolStripTextBox28.Text = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Medico")
                Exit For
            End If
        Next
    End Sub

    Private Sub ToolStripButton13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton13.Click

        style = MsgBoxStyle.Information
        If ToolStripComboBox11.Text = "" Then
            msg = "Debe seleccionar un Personal Médico para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If
        If ToolStripComboBox12.Text = "" Then
            msg = "Debe seleccionar un Ejercicio para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If


        Dim _Asistencia As Boolean
        Dim _EjercicioCombo As String
        _EjercicioCombo = CStr(ToolStripComboBox12.Text)

        DataGridView10.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If Mid(Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"), 7, 4) = _EjercicioCombo Then
                If Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende") = ToolStripTextBox28.Text Then
                    _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                    _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                    _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                    _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                    _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                    _NombreMedico = _BuscaMedico.Item("Nombre")
                    _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                    _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                    _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                    _Fecha = Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita")
                    _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                    DataGridView10.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _Asistencia, _HoraCita, _Cita)
                End If

            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView10.Rows.Count - 1
            _Asistencia = DataGridView10.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView10.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox26.Text = DataGridView10.Rows.Count - 1
        If CInt(ToolStripTextBox26.Text) > 0 Then
            ToolStripTextBox27.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox26.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel31.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "



    End Sub

    Private Sub ToolStripComboBox13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox13.Click

    End Sub

    Private Sub ToolStripComboBox13_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox13.GotFocus
     
    End Sub

    Private Sub ToolStripComboBox13_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox13.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet2.Pacientes.Rows.Count - 1
            If ToolStripComboBox13.Text = Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres") Then
                ToolStripTextBox29.Text = Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
            End If
        Next
    End Sub

    Private Sub ToolStripComboBox14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox14.Click

    End Sub

    Private Sub ToolStripComboBox14_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox14.GotFocus
        ToolStripComboBox14.Items.Clear()
        For x = 2014 To 2030
            ToolStripComboBox14.Items.Add(x)
        Next
    End Sub

    Private Sub ToolStripButton14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton14.Click
        style = MsgBoxStyle.Information
        If ToolStripComboBox13.Text = "" Then
            msg = "Debe seleccionar un Paciente para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If
        If ToolStripComboBox14.Text = "" Then
            msg = "Debe seleccionar un Ejercicio para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If

        Dim _Asistencia As Boolean
        Dim _EjercicioCombo As String
        _EjercicioCombo = CStr(ToolStripComboBox14.Text)

        DataGridView11.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If Mid(Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"), 7, 4) = _EjercicioCombo Then
                If Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente") = ToolStripTextBox29.Text Then
                    _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                    _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                    _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                    _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                    _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                    _NombreMedico = _BuscaMedico.Item("Nombre")
                    _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                    _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                    _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                    _Fecha = Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita")
                    _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")

                    DataGridView11.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _Asistencia, _HoraCita, _Cita)
                End If

            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView11.Rows.Count - 1
            _Asistencia = DataGridView11.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView11.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox30.Text = DataGridView11.Rows.Count - 1
        If CInt(ToolStripTextBox30.Text) > 0 Then
            ToolStripTextBox31.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox30.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel34.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "

    End Sub

    Private Sub ToolStripButton15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton15.Click
        style = MsgBoxStyle.Information
        If ToolStripComboBox15.Text = "" Then
            msg = "Debe seleccionar un Paciente para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If
        If ToolStripComboBox16.Text = "" Then
            msg = "Debe seleccionar un Mes para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If
        If ToolStripComboBox17.Text = "" Then
            msg = "Debe seleccionar un Ejercicio para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If


        Dim _Asistencia As Boolean
        Dim _MesCombo, _EjercicioCombo As String
        _MesCombo = CStr(ToolStripComboBox16.SelectedIndex + 1)
        _EjercicioCombo = CStr(ToolStripComboBox17.Text)

        DataGridView12.Rows.Clear()
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If CInt(Mid(Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"), 7, 4)) = CInt(_EjercicioCombo) Then
                If CInt(Mid(Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita"), 4, 2)) = CInt(_MesCombo) Then
                    If Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente") = ToolStripTextBox32.Text Then
                        _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                        _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                        _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                        _Medico = Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende"))
                        _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                        _NombreMedico = _BuscaMedico.Item("Nombre")
                        _Referencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Referencia")
                        _Asistencia = Me.CServiciosDataSet4.Citas.Rows(i).Item("Asiste")
                        _HoraCita = Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")
                        _Fecha = Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita")
                        _Cita = Me.CServiciosDataSet4.Citas.Rows(i).Item("Cita")


                        DataGridView12.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _Asistencia, _HoraCita, _Cita)
                    End If
                End If

            End If

        Next
        _Atendidas = 0
        For Me.i = 0 To DataGridView12.Rows.Count - 1
            _Asistencia = DataGridView12.Rows(i).Cells(5).Value
            If _Asistencia Then
                DataGridView12.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                _Atendidas = _Atendidas + 1
            End If
        Next

        ToolStripTextBox33.Text = DataGridView12.Rows.Count - 1
        If CInt(ToolStripTextBox33.Text) > 0 Then
            ToolStripTextBox34.Text = _Atendidas
            _Porcentaje = _Atendidas / CInt(ToolStripTextBox33.Text) * 100
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel37.Text = "Porcentaje : " & CStr(_Porcentaje) & " % "

    End Sub

    Private Sub ToolStripComboBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox15.Click

    End Sub

    Private Sub ToolStripComboBox15_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox15.GotFocus

      

    End Sub

    Private Sub ToolStripComboBox15_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox15.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet2.Pacientes.Rows.Count - 1
            If ToolStripComboBox15.Text = Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres") Then
                ToolStripTextBox32.Text = Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
            End If
        Next

    End Sub

    Private Sub ToolStripComboBox16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox16.Click

    End Sub

    Private Sub ToolStripComboBox16_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox16.GotFocus
        ToolStripComboBox16.Items.Clear()

        ToolStripComboBox16.Items.Add("ENERO")
        ToolStripComboBox16.Items.Add("FEBRERO")
        ToolStripComboBox16.Items.Add("MARZO")
        ToolStripComboBox16.Items.Add("ABRIL")
        ToolStripComboBox16.Items.Add("MAYO")
        ToolStripComboBox16.Items.Add("JUNIO")
        ToolStripComboBox16.Items.Add("JULIO")
        ToolStripComboBox16.Items.Add("AGOSTO")
        ToolStripComboBox16.Items.Add("SEPTIEMBRE")
        ToolStripComboBox16.Items.Add("OCTUBRE")
        ToolStripComboBox16.Items.Add("NOVIEMBRE")
        ToolStripComboBox16.Items.Add("DICIEMBRE")


    End Sub

    Private Sub ToolStripComboBox17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox17.Click

    End Sub

    Private Sub ToolStripComboBox17_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox17.GotFocus
        ToolStripComboBox17.Items.Clear()
        For x = 2014 To 2030
            ToolStripComboBox17.Items.Add(x)
        Next

    End Sub

    Private Sub ToolStripComboBox5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox5.Click

    End Sub

    Private Sub ToolStripComboBox5_GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox5.GotFocus
        ToolStripComboBox5.Items.Clear()
        For x = 2014 To 2030
            ToolStripComboBox5.Items.Add(x)
        Next
    End Sub

    Private Sub ToolStripButton16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton16.Click
        Frm23CitasparaHoy.Show()

    End Sub

    Private Sub ToolStripButton17_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton17.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet8.Centros' Puede moverla o quitarla según sea necesario.
        '  Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)
        '  'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        '  Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        '  'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet12.Medicos' Puede moverla o quitarla según sea necesario.
        '  Me.MedicosTableAdapter.Fill(Me.CServiciosDataSet12.Medicos)
        '  'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet4.Citas' Puede moverla o quitarla según sea necesario.
        '  Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)
        '
        '        ToolStripComboBox1.Items.Clear()
        '        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
        ' _NombreMedico = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre")
        ' ToolStripComboBox1.Items.Add(_NombreMedico)
        '
        '        Next

        Button2_Click(sender, e)

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim _Espacios As String
        Dim title As String
        Dim response As MsgBoxResult
    
        Button6.Visible = False
        Button8.Visible = False

        _Espacios = ""
        msg = "Está seguro de la Asistencia de este Paciente?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Asistencia a la Cita"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Genera la Consulta

            _Espacios = TextBox1.Text & TextBox2.Text & TextBox3.Text & TextBox4.Text & TextBox5.Text
            If _Espacios = "" Then
                msg = "Debe seleccionar una Cita para poder dar la Asistencia"
                style = MsgBoxStyle.Information
                MsgBox(msg, style)
                Exit Sub
            End If

            If Not IsDate(TextBox6.Text) Then
                msg = "La Fecha que se muestra no es una fecha válida"
                style = MsgBoxStyle.Information
                TextBox6.BackColor = Color.PeachPuff
                MsgBox(msg, style)
                TextBox6.BackColor = Color.WhiteSmoke
                Exit Sub
            End If
            If ComboBox1.Text = "" Then
                msg = "Debe seleccionar un horario de esta Cita"
                style = MsgBoxStyle.Information
                MsgBox(msg, style)
                Exit Sub
            End If

            Panel1.Visible = True

            Button1.Enabled = False
            msg = "Propocione los datos de pago del Paciente"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            TextBox7.Focus()


        Else
            ' Perform some other action.
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet33.Consultas' Puede moverla o quitarla según sea necesario.
        Me.ConsultasTableAdapter.Fill(Me.CServiciosDataSet33.Consultas)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet8.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet12.Medicos' Puede moverla o quitarla según sea necesario.
        Me.MedicosTableAdapter.Fill(Me.CServiciosDataSet12.Medicos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet4.Citas' Puede moverla o quitarla según sea necesario.
        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)

        ToolStripComboBox1.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            _NombreMedico = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre")
            ToolStripComboBox1.Items.Add(_NombreMedico)
        Next

        Panel1.Visible = False
        Dim _Hora As String
        _Hora = "0"
        ComboBox1.Items.Clear()
        For Me.i = 0 To 23
            _Hora = Trim(CStr(i)) & ":00 HRS"
            ComboBox1.Items.Add(_Hora)
            _Hora = Trim(CStr(i)) & ":30 HRS"
            ComboBox1.Items.Add(_Hora)
        Next


        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        ComboBox1.Text = ""
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next


        Button6.Visible = False
        ToolStripComboBox3.Items.Clear()
        ToolStripComboBox13.Items.Clear()
        ToolStripComboBox15.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")
            ToolStripComboBox3.Items.Add(_NombrePaciente)
            ToolStripComboBox13.Items.Add(_NombrePaciente)
            ToolStripComboBox15.Items.Add(_NombrePaciente)
        Next
        ToolStripComboBox3.Sorted = True
        ToolStripComboBox13.Sorted = True
        ToolStripComboBox15.Sorted = True
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya fué registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

    
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = ToolStripTextBox2.Text
        TextBox5.Text = ToolStripComboBox1.Text
        TextBox6.Text = _Renglon.Cells(4).Value

        Button6.Visible = True
        Button1.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellLeave
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If

    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        _Renglon = DataGridView3.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya fué registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

       
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value
        Button1.Visible = True
        Button6.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
     
    End Sub

    Private Sub DataGridView3_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellLeave
        _Renglon = DataGridView3.CurrentRow
        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        ComboBox1.Text = ""
    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        _Renglon = DataGridView4.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya fué procesada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = ToolStripTextBox4.Text
        TextBox3.Text = ToolStripComboBox3.Text
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value

        Button1.Visible = True
        Button6.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub BindingNavigator1_RefreshItems(sender As System.Object, e As System.EventArgs) Handles BindingNavigator1.RefreshItems

    End Sub

    Private Sub DataGridView5_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView5.CellClick
        _Renglon = DataGridView5.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(7).Value = True Then
            msg = "Esta cita ya fué registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(8).Value
        TextBox2.Text = _Renglon.Cells(5).Value
        TextBox3.Text = _Renglon.Cells(6).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value

        Button1.Visible = True
        Button6.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex
    End Sub

    Private Sub DataGridView5_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick

    End Sub

    Private Sub DataGridView5_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView5.CellLeave
        _Renglon = DataGridView5.CurrentRow
        If _Renglon.Cells(7).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView6_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellClick
        _Renglon = DataGridView6.CurrentRow

        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(7).Value = True Then
            msg = "Esta Cita ya fué registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(8).Value
        TextBox2.Text = _Renglon.Cells(5).Value
        TextBox3.Text = _Renglon.Cells(6).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value

        Button1.Visible = True
        Button6.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex
    End Sub

    Private Sub DataGridView6_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub DataGridView6_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellLeave
        _Renglon = DataGridView6.CurrentRow

        If _Renglon.Cells(7).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView7_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellClick
        _Renglon = DataGridView7.CurrentRow

        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya ha sido registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = ToolStripTextBox5.Text

        Button1.Visible = True
        Button6.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex

    End Sub

    Private Sub DataGridView7_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellContentClick

    End Sub

    Private Sub DataGridView7_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellLeave
        _Renglon = DataGridView7.CurrentRow
        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView8_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView8.CellClick
        _Renglon = DataGridView8.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya ha sido registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value

        Button6.Visible = True
        Button1.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex

    End Sub

    Private Sub DataGridView8_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView8.CellContentClick

    End Sub

    Private Sub DataGridView8_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView8.CellLeave
        _Renglon = DataGridView8.CurrentRow
        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView9_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView9.CellClick
        _Renglon = DataGridView9.CurrentRow

        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya ha sido registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value

        Button1.Visible = True
        Button6.Visible = True
        Button8.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex
    End Sub

    Private Sub DataGridView9_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView9.CellContentClick

    End Sub

    Private Sub DataGridView9_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView9.CellLeave
        _Renglon = DataGridView9.CurrentRow
        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub DataGridView10_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView10.CellClick
        _Renglon = DataGridView10.CurrentRow

        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya ha sido registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value

        Button6.Visible = True
        Button8.Visible = True
        Button1.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex
    End Sub

    Private Sub DataGridView10_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView10.CellContentClick

    End Sub

    Private Sub DataGridView10_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView10.CellLeave
        _Renglon = DataGridView10.CurrentRow

        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If

    End Sub

    Private Sub DataGridView11_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView11.CellClick
        _Renglon = DataGridView11.CurrentRow

        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)

        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya ha sido registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value

        Button6.Visible = True
        Button8.Visible = True
        Button1.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex

    End Sub

    Private Sub DataGridView11_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView11.CellContentClick

    End Sub

    Private Sub DataGridView11_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView11.CellLeave
        _Renglon = DataGridView11.CurrentRow

        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If

    End Sub

    Private Sub DataGridView12_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView12.CellClick
        _Renglon = DataGridView12.CurrentRow

        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        Button3_Click(sender, e)


        If _Renglon.Cells(5).Value = True Then
            msg = "Esta Cita ya ha sido registrada como Asistencia"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _Renglon.Cells(7).Value
        TextBox2.Text = _Renglon.Cells(0).Value
        TextBox3.Text = _Renglon.Cells(1).Value
        TextBox4.Text = _Renglon.Cells(2).Value
        TextBox5.Text = _Renglon.Cells(3).Value
        TextBox6.Text = _Renglon.Cells(4).Value

        Button6.Visible = True
        Button8.Visible = True
        Button1.Visible = True

        TextBox11.Text = TabControl1.SelectedIndex

    End Sub

    Private Sub DataGridView12_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView12.CellContentClick

    End Sub

    Private Sub DataGridView12_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView12.CellLeave
        _Renglon = DataGridView12.CurrentRow

        If _Renglon.Cells(5).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click

        Dim _Consecutivo, _Mide As Integer
        Dim _Xconsecutivo As String
        Dim _Fila As DataRow
        Dim _FechaConsulta As String
        Dim _HoraConsulta As String = ComboBox1.Text
        Dim _NumReceta, _Atiende As String

        If Not IsNumeric(TextBox8.Text) Then
            msg = "El importe registrado en la casilla de IMPORTE PAGADO, no es válido"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Button5_Click(sender, e)
            Exit Sub
        End If


        ' Actualiza el Consecutivo 
        _Busca = Me.CServiciosDataSet3.Documentos.FindByDocumento("CON")
        _Consecutivo = _Busca.Item("Consecutivo")
        _Xconsecutivo = "0000" & Trim(CStr(_Consecutivo))
        _Mide = Len(_Xconsecutivo)
        _Xconsecutivo = "CO" & Mid(_Xconsecutivo, (_Mide - 3), 4)
        _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager4.UpdateAll(Me.CServiciosDataSet3)
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        ' Actualiza la Cita
        _Cita = TextBox1.Text
        _Paciente = TextBox2.Text
        _Busca = Me.CServiciosDataSet4.Citas.FindByPacienteCita(_Paciente, _Cita)
        _Busca.Item("Estatus") = "ATENDIDO"
        _Busca.Item("Asiste") = True
        _Centro = _Busca.Item("Centro")

        Me.Validate()
        CitasBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.CServiciosDataSet4)
        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)

        _FechaConsulta = (TextBox6.Text)
        _Atiende = TextBox4.Text
        _NumReceta = ""
        _Referencia = "Cita No. " & TextBox1.Text

        ' Registra la Consulta
        _Fila = CServiciosDataSet33.Consultas.NewRow
        _Fila.Item("Consulta") = _Xconsecutivo
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
        _Fila.Item("ImporteConsulta") = CInt(TextBox8.Text)
        _Fila.Item("ImporteDescuento") = 0
        _Fila.Item("Adicional") = TextBox7.Text

        Me.CServiciosDataSet33.Consultas.Rows.Add(_Fila)
        Me.Validate()
        ConsultasBindingSource.EndEdit()
        Me.TableAdapterManager5.UpdateAll(Me.CServiciosDataSet33)
        Me.ConsultasTableAdapter.Fill(Me.CServiciosDataSet33.Consultas)

        msg = "Proceso concluído correctamente"
        style = MsgBoxStyle.Information
        MsgBox(msg, style)

        Button2_Click(sender, e)
        ToolStripButton2_Click(sender, e)
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Button2_Click(sender, e)
        Button3_Click(sender, e)
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click

        ' Cancelar la Cita
        Button1.Visible = False
        Button8.Visible = False

        Dim _Find As Boolean
        Dim _Rows_Ci() As DataRow
        Dim title As String
        Dim response As MsgBoxResult

        If TextBox1.Text = "" Then
            msg = "Debe Seleccionar una Cita para poder Cancelar"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If

        ' Busca la Cita para verificar que exista
        _Cita = TextBox1.Text
        _Paciente = TextBox2.Text

        _Find = False

        _Rows_Ci = CServiciosDataSet4.Citas.Select("Cita = " & "'" & _Cita & "'")

        For Me.i = 0 To _Rows_Ci.GetUpperBound(0)
            _Find = True
        Next

        If _Find = False Then
            msg = "La Cita NO existe. Verifique"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)
            Exit Sub
        End If


        msg = "Está Seguro de Cancelar esta Cita"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cancelación de Citas"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            'Cancela
            _Busca = CServiciosDataSet4.Citas.FindByPacienteCita(_Paciente, _Cita)
            _Busca.Item("Estatus") = "CANCELADO"

            Me.Validate()
            CitasBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet4)
            Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)

            msg = "Cita Cancelada Correctamente"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)

            Button2_Click(sender, e)

        Else
            ' Perform some other action.
        End If


    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        TabControl1.SelectedTab = TabPage13
        TextBox9.Text = TextBox6.Text
        TextBox9.Enabled = False

        Button1.Visible = False
        Button6.Visible = False
    End Sub

    Private Sub MonthCalendar3_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar3.DateChanged

    End Sub

    Private Sub MonthCalendar3_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar3.DateSelected
        TextBox10.Text = MonthCalendar3.SelectionRange.Start

    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click

        Dim _Boton As New Button
        _Boton.Name = TextBox12.Text
        Dim _TabPage(12) As TabPage
        _TabPage(0) = TabPage1
        _TabPage(1) = TabPage2
        _TabPage(2) = TabPage3
        _TabPage(3) = TabPage4
        _TabPage(4) = TabPage5
        _TabPage(5) = TabPage6
        _TabPage(6) = TabPage7
        _TabPage(7) = TabPage8
        _TabPage(8) = TabPage9
        _TabPage(9) = TabPage10
        _TabPage(10) = TabPage11
        _TabPage(11) = TabPage12
        _TabPage(12) = TabPage13

        Dim title As String
        Dim response As MsgBoxResult
        Dim _Index As Integer = CInt(TextBox11.Text)

        msg = "Está seguro de Modificar la Fecha de Esta Cita?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Cambio de Fecha para las Citas"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Modifica Fecha
            _Paciente = TextBox2.Text
            _Cita = TextBox1.Text

            _Busca = CServiciosDataSet4.Citas.FindByPacienteCita(_Paciente, _Cita)
            _Busca.Item("FechaCita") = TextBox10.Text

            Me.Validate()
            CitasBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet4)
            Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)

            msg = "Cambio efectuado Correctamente"
            style = MsgBoxStyle.Information
            MsgBox(msg, style)

            TabControl1.SelectedTab = _TabPage(_Index)
            _Boton.PerformClick()
            Panel1.Visible = False
            Button1.Visible = False
            Button6.Visible = False
            Button8.Visible = False

        Else
            ' Perform some other action.
        End If
    End Sub

    Private Sub TabPage13_Click(sender As System.Object, e As System.EventArgs) Handles TabPage13.Click

    End Sub
End Class