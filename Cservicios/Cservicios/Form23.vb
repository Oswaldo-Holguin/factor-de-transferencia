Public Class Frm23CitasparaHoy
    Public i As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Checacontrol As Control
    Public _NombreMedico, _NombrePaciente, _Centro As String

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Frm23ConsultasHoy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet8.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet4.Citas' Puede moverla o quitarla según sea necesario.
        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet12.Medicos' Puede moverla o quitarla según sea necesario.
        Me.MedicosTableAdapter.Fill(Me.CServiciosDataSet12.Medicos)

        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            _NombreMedico = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre")
            ToolStripComboBox1.Items.Add(_NombreMedico)
        Next

        ToolStripButton3.Enabled = False
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add(" ** Agregar Nuevo Paciente **")
        For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " "
            _NombrePaciente = _NombrePaciente & CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " "
            _NombrePaciente = _NombrePaciente & CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")
            ComboBox1.Items.Add(_NombrePaciente)
        Next
        ComboBox1.Sorted = True

        TextBox2.Text = "SIN REFERENCIA"

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.ContextMenuStrip = ContextMenuStrip1
            End If
        Next




    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Dim _Busca As DataRow
        For Me.i = 0 To Me.CServiciosDataSet12.Medicos.Rows.Count - 1
            If Me.CServiciosDataSet12.Medicos.Rows(i).Item("Nombre") = ToolStripComboBox1.Text Then
                ToolStripTextBox1.Text = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Medico")
                _Centro = Me.CServiciosDataSet12.Medicos.Rows(i).Item("Centro")
                TextBox1.Text = _Centro
                TextBox6.Text = ToolStripComboBox1.Text

                _Busca = CServiciosDataSet8.Centros.FindByCentro(TextBox1.Text)
                TextBox8.Text = _Busca.Item("Descripcion")

            End If
        Next



        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.BackColor = Color.White
            End If
        Next
        MonthCalendar1.Visible = True

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        TextBox6.Text = ToolStripComboBox1.Text
        TextBox7.Text = ToolStripTextBox2.Text
        Dim _Busca As DataRow
        _Busca = Me.CServiciosDataSet8.Centros.FindByCentro(_Centro)
        TextBox8.Text = _Busca.Item("Descripcion")


        Dim _CajaTexto, _Paciente As String

        Dim x As Integer
        Dim _Cuantos As Integer = Me.CServiciosDataSet4.Citas.Rows.Count

        _CajaTexto = ""
        _Paciente = ""
        ' Verifica las consultas para ese Personal Médico
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            ' MsgBox("Entró a recorrer el archivo")
            If Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende")) = Trim(ToolStripTextBox1.Text) Then
                '    MsgBox(" Si encuentra personal Médico ")
                If Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita") = TextBox7.Text Then
                    '  MsgBox("Si encuentra la fecha")

                    x = 1
                    For Each Me._Checacontrol In Me.Controls
                        If TypeOf _Checacontrol Is Label Then
                            If UCase(Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")) = UCase(_Checacontrol.Text) Then
                                '  MsgBox("Encuentra la hora " & _checacontrol.Text)
                                _CajaTexto = "TBHora" + Trim(Mid(_Checacontrol.Name, 6, 2))
                                ' MsgBox("CajaTexto = " & _CajaTexto)
                                _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                            End If
                            x = x + 1
                        End If
                    Next


                    For Each Me._Checacontrol In Me.Controls

                        If TypeOf _Checacontrol Is TextBox Then
                            If _Checacontrol.Name = _CajaTexto Then
                                _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                                _Checacontrol.Text = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")
                                _Checacontrol.Enabled = False
                            End If

                        End If
                    Next

                End If
            End If

        Next

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Text <> "" Then
                    _Checacontrol.BackColor = Color.PeachPuff
                End If

            End If
        Next







    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        ToolStripTextBox2.Text = MonthCalendar1.SelectionRange.Start
        TextBox7.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False
    End Sub

    Private Sub ToolStripTextBox2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripTextBox2.Click

    End Sub

    Private Sub ToolStripTextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles ToolStripTextBox2.GotFocus
        MonthCalendar1.Visible = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Mid(Trim(ComboBox1.Text), 1, 2) = "**" Then
            Frm57NuevosPacientes.Show()
        End If

        For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " "
            _NombrePaciente = _NombrePaciente & CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " "
            _NombrePaciente = _NombrePaciente & CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")

            If _NombrePaciente = ComboBox1.Text Then
                TextBox3.Text = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
                Exit For
            End If
        Next



    End Sub

    Private Sub Etiqu6_Click(sender As System.Object, e As System.EventArgs) Handles Etiqu6.Click

    End Sub

    Private Sub ToolStripButton2_Click_1(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        TextBox6.Text = ToolStripComboBox1.Text
        TextBox7.Text = ToolStripTextBox2.Text
        Dim _Busca As DataRow
        _Busca = Me.CServiciosDataSet8.Centros.FindByCentro(_Centro)
        TextBox8.Text = _Busca.Item("Descripcion")


        Dim _CajaTexto, _Paciente As String

        Dim x As Integer
        Dim _Cuantos As Integer = Me.CServiciosDataSet4.Citas.Rows.Count

        _CajaTexto = ""
        _Paciente = ""
        ' Verifica las consultas para ese Personal Médico
        For Me.i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If CServiciosDataSet4.Citas.Rows(i).Item("Estatus") <> "CANCELADO" Then
                If Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende")) = Trim(ToolStripTextBox1.Text) Then
                    '    MsgBox(" Si encuentra personal Médico ")
                    If Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita") = TextBox7.Text Then
                        '  MsgBox("Si encuentra la fecha")

                        x = 1
                        For Each Me._Checacontrol In Me.Controls
                            If TypeOf _Checacontrol Is Label Then
                                If UCase(Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")) = UCase(_Checacontrol.Text) Then
                                    '  MsgBox("Encuentra la hora " & _checacontrol.Text)
                                    _CajaTexto = "TBHora" + Trim(Mid(_Checacontrol.Name, 6, 2))
                                    ' MsgBox("CajaTexto = " & _CajaTexto)
                                    _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                                End If
                                x = x + 1
                            End If
                        Next


                        For Each Me._Checacontrol In Me.Controls

                            If TypeOf _Checacontrol Is TextBox Then
                                If _Checacontrol.Name = _CajaTexto Then
                                    _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                                    _Checacontrol.Text = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")
                                    _Checacontrol.Enabled = False
                                End If

                            End If
                        Next

                    End If
                End If
            End If
        Next

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Text <> "" Then
                    _Checacontrol.BackColor = Color.PeachPuff
                End If

            End If
        Next
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem1.Click
        ' Asigna a ese horario la Consulta del Paciente Seleccionado

        TextBox2.Visible = True
        Label6.Visible = True
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Focused = True Then
                    _Checacontrol.Text = ComboBox1.Text
                    _Checacontrol.BackColor = Color.PeachPuff
                    ToolStripButton3.Enabled = True
                End If
            End If
        Next

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guardar Informacón

        Dim _Xnombre, _Paciente As String
        Dim _Find As Boolean
        Dim _Registro, _Busca As DataRow
        Dim _Consecutivo, _Mide As Integer
        Dim _XConsecutivo, _HoraCita, _NombreControl As String
        Dim _HC As Control

        _Centro = TextBox1.Text
        _HoraCita = ""
        _NombreControl = ""
        _Paciente = ""
        _Xnombre = ""
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Enabled = True Then
                    If _Checacontrol.Text <> "" Then
                        If _Checacontrol.Name <> "TextBox2" And _Checacontrol.Name <> "TextBox3" Then
                            _NombrePaciente = _Checacontrol.Text

                            _Find = False
                            For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
                                _Xnombre = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " "
                                _Xnombre = _Xnombre & CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " "
                                _Xnombre = _Xnombre & CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")

                                If _Xnombre = _NombrePaciente Then
                                    _Paciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
                                    _Find = True
                                    Exit For
                                End If

                            Next

                            If _Find = False Then
                                Msg = "El Paciente " & _NombrePaciente & " no está registrado en la tabla de Pacientes. Verifique"
                                Style = MsgBoxStyle.Information
                                MsgBox(Msg, Style)
                                Exit Sub
                            Else



                                ' Obtiene el Siguiente Folio de la Cita
                                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("CIT")
                                _Consecutivo = _Busca.Item("Consecutivo")
                                _XConsecutivo = "0000" & Trim(CStr(_Consecutivo))
                                _Mide = Len(_XConsecutivo)
                                _XConsecutivo = "CI" & Mid(_XConsecutivo, (_Mide - 3), 4)

                                _Registro = CServiciosDataSet4.Citas.NewRow
                                _Registro.Item("Paciente") = TextBox3.Text
                                _Registro.Item("Cita") = _XConsecutivo
                                _Registro.Item("FechaCita") = TextBox7.Text
                                _NombreControl = _Checacontrol.Name
                                _NombreControl = Trim("Etiqu" & Mid(_NombreControl, 7, 2))
                                For Each _HC In Me.Controls
                                    If _HC.Name = _NombreControl Then
                                        _HoraCita = _HC.Text
                                        Exit For
                                    End If
                                Next

                            
                                _Registro.Item("HoraCita") = _HoraCita
                                _Registro.Item("Centro") = _Centro
                                _Registro.Item("Atiende") = ToolStripTextBox1.Text
                                If TextBox2.Text = "" Then
                                    _Registro.Item("Referencia") = "SIN REFERENCIA"
                                Else
                                    _Registro.Item("Referencia") = TextBox2.Text
                                End If

                                _Registro.Item("Comentarios") = "CITA AGENDADA EN CONSULTA DE HORARIO"
                                _Registro.Item("Asiste") = False
                                _Registro.Item("Estatus") = "PROGRAMADO"
                                CServiciosDataSet4.Citas.Rows.Add(_Registro)

                            End If




                        End If
                    End If
                End If

            End If

        Next

        Me.Validate()
        CitasBindingSource.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet4)
        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)

        ' Actualiza el Consecutivo
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("CIT")
        _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager4.UpdateAll(CServiciosDataSet3)
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)


        Msg = "Cita Agendada Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)
        Me.Close()


    End Sub

    Private Sub EliminarDeEsteHorarioLaCitaParaEstePacienteToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EliminarDeEsteHorarioLaCitaParaEstePacienteToolStripMenuItem.Click
        ' Elimina la cita para ese dia

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Focused = True Then
                    _Checacontrol.Text = ""
                    _Checacontrol.BackColor = Color.White

                End If
            End If
        Next

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)

        ComboBox1.Items.Clear()
        ComboBox1.Items.Add(" ** Agregar Nuevo Paciente **")
        For Me.i = 0 To CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " "
            _NombrePaciente = _NombrePaciente & CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " "
            _NombrePaciente = _NombrePaciente & CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")
            ComboBox1.Items.Add(_NombrePaciente)
        Next
        ComboBox1.Sorted = True
    End Sub
End Class