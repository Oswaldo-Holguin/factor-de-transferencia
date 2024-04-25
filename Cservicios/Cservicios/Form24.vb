Public Class Frm24consultasMedicas
    Public _Paciente, _NombrePaciente, _NombreMedico, _Medico, _HoraCita, _Referencia, Msg, _Centro As String
    Public _BuscaPaciente, _BuscaMedico, _Busca As DataRow
    Public i, _ImporteConsulta, _Suma, _Porcentaje As Integer
    Public _Fecha As Date
    Public _NombreCentro As String
    Public _Rows_Co() As DataRow
    Public style As MsgBoxStyle



    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

        ' Por Fecha 
        DataGridView1.Enabled = True


        If ToolStripTextBox1.Text = "" Then
            msg = "Debe seleccionar una fecha para la Consulta"
            style = MsgBoxStyle.Information
            MsgBox(Msg, style)
            Exit Sub
        End If

        TextBox1.Text = "2"
        DataGridView1.Rows.Clear()
        Button1_Click(sender, e)


    End Sub

    Private Sub Form24_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet33.Consultas' Puede moverla o quitarla según sea necesario.
        Me.ConsultasTableAdapter1.Fill(Me.CServiciosDataSet33.Consultas)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet12.Medicos' Puede moverla o quitarla según sea necesario.
        Me.MedicosTableAdapter.Fill(Me.CServiciosDataSet12.Medicos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet8.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: esta línea de código carga datos en la tabla 'CserviciosDataset33.Consultas' Puede moverla o quitarla según sea necesario.
        '   Me.ConsultasTableAdapter.Fill(Me.CServiciosDataSet15.Consultas)

        For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno")
            _NombrePaciente = _NombrePaciente & " " & CServiciosDataSet2.Pacientes.Rows(i).Item("Materno")
            _NombrePaciente = _NombrePaciente & " " & CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")

            ToolStripComboBox3.Items.Add(_NombrePaciente)
            ToolStripComboBox13.Items.Add(_NombrePaciente)
            ToolStripComboBox15.Items.Add(_NombrePaciente)
        Next
        ToolStripComboBox3.Sorted = True
        ToolStripComboBox13.Sorted = True
        ToolStripComboBox15.Sorted = True

        Dim _Meses(12) As String
        _Meses(0) = "ENERO"
        _Meses(1) = "FEBRERO"
        _Meses(2) = "MARZO"
        _Meses(3) = "ABRIL"
        _Meses(4) = "MAYO"
        _Meses(5) = "JUNIO"
        _Meses(6) = "JULIO"
        _Meses(7) = "AGOSTO"
        _Meses(8) = "SEPTIEMBRE"
        _Meses(9) = "OCTUBRE"
        _Meses(10) = "NOVIEMBRE"
        _Meses(11) = "DICIEMBRE"
        ' Llena los Meses

        ToolStripComboBox4.Items.Clear()
        ToolStripComboBox8.Items.Clear()
        For i = 0 To 11
            ToolStripComboBox4.Items.Add(_Meses(i))
            ToolStripComboBox8.Items.Add(_Meses(i))
            ToolStripComboBox16.Items.Add(_Meses(i))
            ToolStripComboBox19.Items.Add(_Meses(i))
        Next

        

        ToolStripComboBox5.Items.Clear()
        For Me.i = 2014 To 2030
            ToolStripComboBox5.Items.Add(i)
            ToolStripComboBox9.Items.Add(i)
            ToolStripComboBox12.Items.Add(i)
            ToolStripComboBox14.Items.Add(i)
            ToolStripComboBox17.Items.Add(i)
            ToolStripComboBox20.Items.Add(i)
        Next

        For Me.i = 0 To CServiciosDataSet12.Medicos.Rows.Count - 1
            ToolStripComboBox6.Items.Add(CServiciosDataSet12.Medicos.Rows(i).Item("Nombre"))
            ToolStripComboBox7.Items.Add(CServiciosDataSet12.Medicos.Rows(i).Item("Nombre"))
            ToolStripComboBox11.Items.Add(CServiciosDataSet12.Medicos.Rows(i).Item("Nombre"))
        Next


        ToolStripComboBox10.Items.Clear()
        For Me.i = (Year(Date.Today) - 1) To (Year(Date.Today) + 15) Step 1
            ToolStripComboBox10.Items.Add(i)
        Next


        ToolStripComboBox18.Items.Clear()
        For Me.i = 0 To CServiciosDataSet8.Centros.Rows.Count - 1
            _NombreCentro = CServiciosDataSet8.Centros.Rows(i).Item("Descripcion")

            ToolStripComboBox18.Items.Add(_NombreCentro)

        Next



    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        MonthCalendar1.Visible = True
    End Sub

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged
   
    End Sub

    Private Sub MonthCalendar1_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        ToolStripTextBox1.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False
    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.GotFocus
        ToolStripComboBox1.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            _NombreMedico = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre")

            ToolStripComboBox1.Items.Add(_NombreMedico)
        Next
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

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If ToolStripComboBox1.Text = "" Then
            Msg = "Debe seleccionar un Personal Médico para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(Msg, style)
            Exit Sub
        End If

        DataGridView2.Enabled = True
        DataGridView2.Rows.Clear()

        TextBox1.Text = "4"
        Button1_Click(sender, e)

    
        ToolStripTextBox8.Text = DataGridView2.Rows.Count - 1
        If CInt(ToolStripTextBox8.Text) > 0 Then
            ToolStripTextBox12.Text = _Suma
            _Porcentaje = _Suma / CInt(ToolStripTextBox8.Text)
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel9.Text = "Promedio : " & CStr(_Porcentaje)

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
            Msg = "Debe seleccionar una Clínica para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(Msg, style)
            Exit Sub
        End If

        DataGridView3.Enabled = True
        DataGridView3.Rows.Clear()
        TextBox1.Text = "5"

        Button1_Click(sender, e)

        ToolStripTextBox9.Text = DataGridView3.Rows.Count - 1
        If CInt(ToolStripTextBox9.Text) > 0 Then
            ToolStripTextBox13.Text = _Suma
            _Porcentaje = _Suma / CInt(ToolStripTextBox9.Text)
        Else
            _Porcentaje = 0
        End If

        ToolStripLabel11.Text = "Promedio : " & CStr(_Porcentaje)
    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        ' Todas 
        TextBox1.Text = "7"
        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click

        TextBox1.Text = "6"
        DataGridView4.Rows.Clear()
        Button1_Click(sender, e)


    End Sub

    Private Sub ToolStripComboBox3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox3.Click

    End Sub

    Private Sub DataGridView6_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click

        TextBox1.Text = "8"
        Button1_Click(sender, e)



    End Sub

    Private Sub ToolStripButton10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton10.Click
        MonthCalendar2.Visible = True

    End Sub

    Private Sub ToolStripComboBox6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox6.Click

    End Sub

    Private Sub ToolStripComboBox6_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox6.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet12.Medicos.Rows.Count - 1
            If ToolStripComboBox6.Text = CServiciosDataSet12.Medicos.Rows(i).Item("Nombre") Then
                ToolStripTextBox6.Text = CServiciosDataSet12.Medicos.Rows(i).Item("Medico")
            End If
        Next

    End Sub

    Private Sub MonthCalendar2_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateChanged

    End Sub

    Private Sub MonthCalendar2_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateSelected
        ToolStripTextBox5.Text = MonthCalendar2.SelectionRange.Start
        MonthCalendar2.Visible = False

    End Sub

    Private Sub ToolStripButton9_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton9.Click
        ' Presenta Información por Fecha y Médico

        DataGridView7.Rows.Clear()
        _Suma = 0
        For Me.i = 0 To Me.CServiciosDataSet33.Consultas.Rows.Count - 1
            If CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta") = CDate(ToolStripTextBox5.Text) Then
                If CServiciosDataSet33.Consultas.Rows(i).Item("Atiende") = ToolStripTextBox6.Text Then

                    _Centro = CServiciosDataSet33.Consultas.Rows(i).Item("Centro")
                    _Busca = CServiciosDataSet8.Centros.FindByCentro(_Centro)
                    _NombreCentro = _Busca.Item("Descripcion")
                    _Paciente = Me.CServiciosDataSet33.Consultas.Rows(i).Item("Paciente")
                    _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                    _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
                    _Medico = Trim(Me.CServiciosDataSet33.Consultas.Rows(i).Item("Atiende"))
                    _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
                    _NombreMedico = _BuscaMedico.Item("Nombre")
                    _Referencia = Me.CServiciosDataSet33.Consultas.Rows(i).Item("Referencia")
                    _ImporteConsulta = Me.CServiciosDataSet33.Consultas.Rows(i).Item("ImporteConsulta") - Me.CServiciosDataSet33.Consultas.Rows(i).Item("ImporteDescuento")
                    _Fecha = (Me.CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta"))
                    _HoraCita = Me.CServiciosDataSet33.Consultas.Rows(i).Item("HoraConsulta")


                    _Suma = _Suma + _ImporteConsulta
                    ' DataGridView5.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _ImporteConsulta, _HoraCita)
                    DataGridView7.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Referencia, True, _HoraCita, _ImporteConsulta)

                End If
            End If

        Next

        ToolStripTextBox19.Text = DataGridView7.Rows.Count - 1
        ToolStripTextBox35.Text = Format(_Suma, "$###,###,##0.00")
    End Sub

    Private Sub ToolStripComboBox7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox7.Click

    End Sub

    Private Sub ToolStripComboBox7_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox7.SelectedIndexChanged
        ' 23
        For Me.i = 0 To CServiciosDataSet12.Medicos.Rows.Count - 1
            If CServiciosDataSet12.Medicos.Rows(i).Item("Nombre") = ToolStripComboBox7.Text Then
                ToolStripTextBox23.Text = CServiciosDataSet12.Medicos.Rows(i).Item("Medico")
            End If
        Next
    End Sub

    Private Sub ToolStripButton11_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton11.Click
        ' Mes y Médico

        TextBox1.Text = "11"
        Button1_Click(sender, e)

    End Sub

    Private Sub BindingNavigator9_RefreshItems(sender As System.Object, e As System.EventArgs) Handles BindingNavigator9.RefreshItems

    End Sub

    Private Sub ToolStripButton12_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton12.Click
        TextBox1.Text = "12"
        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click


        Dim _Mes, _Año As String

        If TextBox1.Text = "15" Then
            ' Consultas por Ejercicio
            DataGridView12.Rows.Clear()
        End If


        If TextBox1.Text = "12" Then
            ' Consultas por Ejercicio
            DataGridView9.Rows.Clear()
        End If
        If TextBox1.Text = "11" Then
            DataGridView8.Rows.Clear()
        End If

        If TextBox1.Text = "8" Then
            ' Por mes 
            DataGridView6.Enabled = True
            DataGridView6.Rows.Clear()
        End If

        If TextBox1.Text = "7" Then
            DataGridView5.Enabled = True
            DataGridView5.Rows.Clear()
        End If


        _Suma = 0
        For Me.i = 0 To Me.CServiciosDataSet33.Consultas.Rows.Count - 1

            _Centro = CServiciosDataSet33.Consultas.Rows(i).Item("Centro")
            _Busca = CServiciosDataSet8.Centros.FindByCentro(_Centro)
            _NombreCentro = _Busca.Item("Descripcion")
            _Paciente = Me.CServiciosDataSet33.Consultas.Rows(i).Item("Paciente")
            _BuscaPaciente = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
            _NombrePaciente = _BuscaPaciente.Item("Paterno") & " " & _BuscaPaciente.Item("Materno") & " " & _BuscaPaciente.Item("Nombres")
            _Medico = Trim(Me.CServiciosDataSet33.Consultas.Rows(i).Item("Atiende"))
            _BuscaMedico = Me.CServiciosDataSet12.Medicos.FindByMedico(_Medico)
            _NombreMedico = _BuscaMedico.Item("Nombre")
            _Referencia = Me.CServiciosDataSet33.Consultas.Rows(i).Item("Referencia")
            _ImporteConsulta = Me.CServiciosDataSet33.Consultas.Rows(i).Item("ImporteConsulta") - Me.CServiciosDataSet33.Consultas.Rows(i).Item("ImporteDescuento")
            _Fecha = (Me.CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta"))
            _HoraCita = Me.CServiciosDataSet33.Consultas.Rows(i).Item("HoraConsulta")
            _Mes = MonthName(Month(Me.CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")))
            _Año = CStr(Year(Me.CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")))


            If TextBox1.Text = "12" Then
                If Year(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox10.Text Then
                    DataGridView9.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, True, _HoraCita, _ImporteConsulta)
                    _Suma = _Suma + _ImporteConsulta
                End If
                ToolStripTextBox24.Text = DataGridView9.Rows.Count - 1
                ToolStripTextBox18.Text = Format(_Suma, "$###,###,##0.00")
            End If

            If TextBox1.Text = "11" Then
                If Month(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox8.SelectedIndex + 1 Then
                    If Year(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox9.Text Then
                        If CServiciosDataSet33.Consultas.Rows(i).Item("Atiende") = ToolStripTextBox23.Text Then
                            _Suma = _Suma + _ImporteConsulta
                            DataGridView8.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, True, _HoraCita, _ImporteConsulta)
                        End If
                    End If
                End If
                ToolStripTextBox21.Text = DataGridView8.Rows.Count - 1
                ToolStripTextBox20.Text = Format(_Suma, "$###,###,##0.00")
            End If

            ' Por Mes 
            If TextBox1.Text = "8" Then
         
                If UCase(_Mes) = UCase(ToolStripComboBox4.Text) Then
                    If CStr(_Año) = ToolStripComboBox5.Text Then
                        _Suma = _Suma + _ImporteConsulta
                        DataGridView6.Rows.Add(_Centro, _NombreCentro, _Medico, _NombreMedico, _Fecha, _Paciente, _NombrePaciente, _HoraCita, _ImporteConsulta)
                    End If
                End If
                ToolStripTextBox22.Text = Format(_Suma, "$###,###,##0.00")
                ToolStripTextBox17.Text = DataGridView6.Rows.Count - 1
            End If


            ' Todas 
            If TextBox1.Text = "7" Then
                _Suma = _Suma + _ImporteConsulta
                DataGridView5.Rows.Add(_Centro, _NombreCentro, _Medico, _NombreMedico, _Fecha, _Paciente, _NombrePaciente, _HoraCita, _ImporteConsulta)
                ToolStripTextBox15.Text = DataGridView5.Rows.Count - 1
                ToolStripTextBox16.Text = Format(_Suma, "$###,###,##0.00")
            End If

            ' Por Fecha 
            If TextBox1.Text = "2" Then
                If Me.CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta") = ToolStripTextBox1.Text Then
                    _HoraCita = Me.CServiciosDataSet33.Consultas.Rows(i).Item("HoraConsulta")
                    _Suma = _Suma + _ImporteConsulta

                    DataGridView1.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Referencia, _ImporteConsulta, _HoraCita, _ImporteConsulta)
                End If
                If DataGridView1.Rows.Count - 1 > 0 Then
                    ToolStripTextBox7.Text = DataGridView1.Rows.Count - 1
                    ToolStripTextBox11.Text = Format(_Suma, "$###,###,##0.00")
                    _Porcentaje = _Suma / CInt(ToolStripTextBox7.Text)
                Else
                    _Porcentaje = 0
                End If
            End If

            If TextBox1.Text = "4" Then
                If Me.CServiciosDataSet33.Consultas.Rows(i).Item("Atiende") = ToolStripTextBox2.Text Then

                    _Suma = _Suma + _ImporteConsulta
                    DataGridView2.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _ImporteConsulta, _HoraCita)
                End If
            End If

            If TextBox1.Text = "5" Then
                If Me.CServiciosDataSet33.Consultas.Rows(i).Item("Centro") = ToolStripTextBox3.Text Then
                    _Suma = _Suma + _ImporteConsulta
                    DataGridView3.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, _ImporteConsulta, _HoraCita)
                End If
            End If

            If TextBox1.Text = "13" Then
                If Year(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox12.Text Then
                    If CServiciosDataSet33.Consultas.Rows(i).Item("Atiende") = ToolStripTextBox28.Text Then
                        DataGridView10.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, True, _HoraCita, _ImporteConsulta)
                    End If

                    _Suma = _Suma + _ImporteConsulta
                End If
                ToolStripTextBox26.Text = DataGridView10.Rows.Count - 1
                ToolStripTextBox27.Text = Format(_Suma, "$###,###,##0.00")

            End If


            If TextBox1.Text = "14" Then
                If Year(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox14.Text Then
                    If CServiciosDataSet33.Consultas.Rows(i).Item("Paciente") = ToolStripTextBox29.Text Then
                        DataGridView11.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, True, _HoraCita, _ImporteConsulta)
                    End If

                    _Suma = _Suma + _ImporteConsulta
                End If
                ToolStripTextBox26.Text = DataGridView10.Rows.Count - 1
                ToolStripTextBox27.Text = Format(_Suma, "$###,###,##0.00")

            End If


            If TextBox1.Text = "6" Then
                If CServiciosDataSet33.Consultas.Rows(i).Item("Paciente") = ToolStripTextBox4.Text Then
                    DataGridView4.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, True, _HoraCita, _ImporteConsulta)
                End If

            End If


            If TextBox1.Text = "15" Then
                _Suma = 0
                If CServiciosDataSet33.Consultas.Rows(i).Item("Paciente") = ToolStripTextBox32.Text Then
                    If Year(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox17.Text Then
                        If Month(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox16.SelectedIndex + 1 Then
                            DataGridView12.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, True, _HoraCita, _ImporteConsulta)
                            _Suma = _Suma + CServiciosDataSet33.Consultas.Rows(i).Item("ImporteConsulta")
                        End If

                    End If
                End If


                ToolStripTextBox33.Text = DataGridView12.Rows.Count - 1
                ToolStripTextBox34.Text = Format(_Suma, "$###,###,##0.00")
            End If

            If TextBox1.Text = "16" Then

                _Suma = 0
                If CServiciosDataSet33.Consultas.Rows(i).Item("Centro") = ToolStripTextBox25.Text Then
                    If Year(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox20.Text Then
                        If Month(CServiciosDataSet33.Consultas.Rows(i).Item("FechaConsulta")) = ToolStripComboBox19.SelectedIndex + 1 Then
                            DataGridView13.Rows.Add(_Paciente, _NombrePaciente, _Medico, _NombreMedico, _Fecha, True, _HoraCita, _ImporteConsulta)
                            _Suma = _Suma + CServiciosDataSet33.Consultas.Rows(i).Item("ImporteConsulta")
                        End If

                    End If
                End If


                ToolStripTextBox36.Text = DataGridView13.Rows.Count - 1
                ToolStripTextBox37.Text = Format(_Suma, "$###,###,##0.00")
            End If




        Next


    End Sub

    Private Sub ToolStripComboBox11_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox11.Click
    
    End Sub

    Private Sub ToolStripComboBox11_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox11.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet12.Medicos.Rows.Count - 1
            If CServiciosDataSet12.Medicos.Rows(i).Item("Nombre") = ToolStripComboBox11.Text Then
                ToolStripTextBox28.Text = CServiciosDataSet12.Medicos.Rows(i).Item("Medico")
            End If
        Next
    End Sub

    Private Sub ToolStripButton13_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton13.Click
        TextBox1.Text = "13"

        If ToolStripTextBox28.Text = "" Then
            Msg = "Debe seleccionar un personal Médico para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(Msg, style)
            Exit Sub
        End If
        If ToolStripComboBox12.Text = "" Then
            Msg = "Debe seleccionar un Ejercicio para esta Consulta"
            style = MsgBoxStyle.Information
            MsgBox(Msg, style)
            Exit Sub
        End If


        DataGridView10.Enabled = True
        DataGridView10.Rows.Clear()
        Button1_Click(sender, e)


    End Sub

    Private Sub ToolStripLabel34_Click(sender As System.Object, e As System.EventArgs)
    End Sub

    Private Sub ToolStripButton14_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton14.Click
        TextBox1.Text = "14"
        DataGridView11.Enabled = True
        DataGridView11.Rows.Clear()

        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripComboBox13_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox13.Click

    End Sub

    Private Sub ToolStripComboBox13_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox13.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno")
            _NombrePaciente = _NombrePaciente & " " & CServiciosDataSet2.Pacientes.Rows(i).Item("Materno")
            _NombrePaciente = _NombrePaciente & " " & CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")

            If ToolStripComboBox13.Text = _NombrePaciente Then
                ToolStripTextBox29.Text = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
            End If
        Next
    End Sub

    Private Sub ToolStripComboBox3_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox3.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno")
            _NombrePaciente = _NombrePaciente & " " & CServiciosDataSet2.Pacientes.Rows(i).Item("Materno")
            _NombrePaciente = _NombrePaciente & " " & CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")

            If ToolStripComboBox3.Text = _NombrePaciente Then
                ToolStripTextBox4.Text = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
            End If
        Next
    End Sub

    Private Sub ToolStripLabel12_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripLabel12.Click

    End Sub

    Private Sub ToolStripTextBox29_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripTextBox29.Click

    End Sub

    Private Sub ToolStripComboBox15_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox15.Click

    End Sub

    Private Sub ToolStripComboBox15_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox15.SelectedIndexChanged
        ToolStripTextBox32.Text = ""

        For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " "
            _NombrePaciente = _NombrePaciente & CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " "
            _NombrePaciente = _NombrePaciente & CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")
            If _NombrePaciente = ToolStripComboBox15.Text Then
                ToolStripTextBox32.Text = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
            End If
        Next

    End Sub

    Private Sub ToolStripButton15_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton15.Click

        TextBox1.Text = "15"
        DataGridView12.Rows.Clear()
        Button1_Click(sender, e)


    End Sub

    Private Sub ToolStripButton16_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton16.Click
        TextBox1.Text = "16"
        DataGridView13.Rows.Clear()
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripLabel35_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripLabel35.Click

    End Sub

    Private Sub ToolStripComboBox18_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripComboBox18.Click

    End Sub

    Private Sub ToolStripComboBox18_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox18.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet8.Centros.Rows.Count - 1
            If CServiciosDataSet8.Centros.Rows(i).Item("Descripcion") = ToolStripComboBox18.Text Then
                ToolStripTextBox25.Text = CServiciosDataSet8.Centros.Rows(i).Item("Centro")
            End If
        Next
    End Sub
End Class