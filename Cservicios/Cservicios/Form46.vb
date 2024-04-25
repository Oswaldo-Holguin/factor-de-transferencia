Public Class Frm46ReservacionesConsecutivas
    Public I As Integer
    Public _HoraInicial, _HoraFinal As String
    Public _FechaInicial, _FechaFinal As Date
    Public _DiaSeleccionado As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Busca As DataRow
    Public _Reservacion, _Recurso, _Referencia, _Motivo, _Comentarios As String
    Public _Solicitante, _TelefonoSolicitante, _CelularSolicitante, _DepartamentoSolicitante As String
    Public _Autoriza, _Usuario As String
    Public _Asistencia, _Satisfactorio As Boolean
    Public _Asistentes As Integer
    Public _Fecha As Date
    Public _Checacontrol As Control

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Frm46ReservacionesConsecutivas_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet36.Recursos' Puede moverla o quitarla según sea necesario.
        Me.RecursosTableAdapter.Fill(Me.CServiciosDataSet36.Recursos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet38.Reservaciones' Puede moverla o quitarla según sea necesario.
        Me.ReservacionesTableAdapter.Fill(Me.CServiciosDataSet38.Reservaciones)


        For Me.I = 0 To CServiciosDataSet36.Recursos.Rows.Count - 1
            ComboBox3.Items.Add(CServiciosDataSet36.Recursos.Rows(I).Item("DescripcionLarga"))
        Next

        For Me.I = 0 To 24
            _HoraInicial = CStr(I) & ":00"
            _HoraFinal = CStr(I) & ":00"
            ComboBox1.Items.Add(_HoraInicial)
            ComboBox8.Items.Add(_HoraInicial)
            ComboBox12.Items.Add(_HoraInicial)
            ComboBox16.Items.Add(_HoraInicial)
            ComboBox20.Items.Add(_HoraInicial)
            ComboBox24.Items.Add(_HoraInicial)
            ComboBox28.Items.Add(_HoraInicial)
            ComboBox2.Items.Add(_HoraFinal)
            ComboBox7.Items.Add(_HoraFinal)
            ComboBox11.Items.Add(_HoraFinal)
            ComboBox15.Items.Add(_HoraFinal)
            ComboBox19.Items.Add(_HoraFinal)
            ComboBox23.Items.Add(_HoraFinal)
            ComboBox27.Items.Add(_HoraFinal)
        Next

        TabControl1.Enabled = False
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus
        MonthCalendar1.Visible = True
        TextBox15.Text = 1
    End Sub


    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected

        Dim _Caja As String = "TextBox" & TextBox15.Text

        For Each Me._Checacontrol In TabPage1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Name = _Caja Then
                    _Checacontrol.Text = MonthCalendar1.SelectionRange.Start
                    MonthCalendar1.Visible = False
                End If
            End If

        Next


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
       
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox15.Text = 2
        MonthCalendar1.Visible = True

    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox4.GotFocus
        TextBox15.Text = 4
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged
       

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox15.Text = 3
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged
        
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox15.Text = 6
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged
        
    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox5.GotFocus
        TextBox15.Text = 5
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox8.GotFocus
        TextBox15.Text = 8
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox15.Text = 7
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox10.GotFocus
        TextBox15.Text = 10
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox9.GotFocus
        TextBox15.Text = 9
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox12_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox12.GotFocus
        TextBox15.Text = 12
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox11_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox11.GotFocus
        TextBox15.Text = 11
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox14_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox14.GotFocus
        TextBox15.Text = 14
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox13.GotFocus
        TextBox15.Text = 13
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox24_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox24.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox24_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox24.TextChanged

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guarda Información
        Dim _Caja As String = "CheckBox"
        Dim _HI, _HF As Integer
        Dim title, _NombreDiaSemana, _NombreDiaReservacion As String
        Dim _Find, _Guarda As Boolean
        Dim _X, _Dias, _Consecutivo, _N As Integer
        Dim _Xconsecutivo As String
        Dim response As MsgBoxResult
        Dim _HoraInicioSolicitada, _HoraFinalSolicitada, _HoraInicioReservada, _HoraFinalReservada As Integer
        Dim _Rows_Re() As DataRow
        Dim _Colision As Boolean

        Msg = "Está a punto de guardar los datos de esta Reservación. Está seguro de Continuar?"   ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Reservaciones Consecutivas"
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Verifica los datos
            If ToolStripTextBox1.Text = "" Then
                Msg = "Debe seleccionar un Recurso para la Reservación"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If


            I = 15
            TextBox17.Text = LoginForm1.TextBox6.Text
            For Each Me._Checacontrol In TabPage1.Controls
                _Caja = "TextBox" & CStr(I)
                If _Checacontrol.Name = _Caja Then
                    _Checacontrol.Text = UCase(_Checacontrol.Text)
                    If _Checacontrol.Text = "" Then
                        _Find = True
                    End If
                End If
                I = I + 1
            Next


            If _Find Then
                Msg = "Hay casilas en blanco, debe llenar todos los Datos"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            '************************************************************************
            ' Registra las Reservaciones

            _Guarda = False
            _Find = False
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("RES")
            _Consecutivo = _Busca.Item("Consecutivo")
            _NombreDiaReservacion = ""
            _NombreDiaSemana = ""
            For Me.I = 1 To 7
                _Find = False
                If I = 1 Then
                    If CheckBox1.Checked = True Then
                        _Find = True
                        _NombreDiaReservacion = "DOMINGO"
                        _HoraInicial = ComboBox1.Text
                        _HoraFinal = ComboBox2.Text
                        _FechaInicial = CDate(TextBox1.Text)
                        _FechaFinal = CDate(TextBox2.Text)
                    End If
                End If
                If I = 2 Then
                    _Find = False
                    If CheckBox2.Checked = True Then
                        _Find = True
                        _NombreDiaReservacion = "LUNES"
                        _HoraInicial = ComboBox8.Text
                        _HoraFinal = ComboBox7.Text
                        _FechaInicial = CDate(TextBox4.Text)
                        _FechaFinal = CDate(TextBox3.Text)
                    End If
                End If
                If I = 3 Then
                    _Find = False
                    If CheckBox3.Checked = True Then
                        _Find = True
                        _NombreDiaReservacion = "MARTES"
                        _HoraInicial = ComboBox12.Text
                        _HoraFinal = ComboBox11.Text
                        _FechaInicial = CDate(TextBox6.Text)
                        _FechaFinal = CDate(TextBox5.Text)
                    End If
                End If
                If I = 4 Then
                    _Find = False
                    If CheckBox4.Checked = True Then
                        _Find = True
                        _NombreDiaReservacion = "MIERCOLES"
                        _HoraInicial = ComboBox16.Text
                        _HoraFinal = ComboBox15.Text
                        _FechaInicial = CDate(TextBox8.Text)
                        _FechaFinal = CDate(TextBox7.Text)
                    End If
                End If
                If I = 5 Then
                    _Find = False
                    If CheckBox5.Checked = True Then
                        _Find = True
                        _NombreDiaReservacion = "JUEVES"
                        _HoraInicial = ComboBox20.Text
                        _HoraFinal = ComboBox19.Text
                        _FechaInicial = CDate(TextBox10.Text)
                        _FechaFinal = CDate(TextBox9.Text)
                    End If
                End If
                If I = 6 Then
                    _Find = False
                    If CheckBox6.Checked = True Then
                        _Find = True
                        _NombreDiaReservacion = "VIERNES"
                        _HoraInicial = ComboBox24.Text
                        _HoraFinal = ComboBox23.Text
                        _FechaInicial = CDate(TextBox12.Text)
                        _FechaFinal = CDate(TextBox11.Text)
                    End If
                End If
                If I = 7 Then
                    _Find = False
                    If CheckBox7.Checked = True Then
                        _Find = True
                        _NombreDiaReservacion = "SABADO"
                        _HoraInicial = ComboBox28.Text
                        _HoraFinal = ComboBox27.Text
                        _FechaInicial = CDate(TextBox14.Text)
                        _FechaFinal = CDate(TextBox13.Text)
                    End If
                End If


                If _Find Then
                    '  MsgBox("Encuentra find es verdader")
                    If Len(_HoraInicial) = 5 Then
                        _HI = CInt(Mid(_HoraInicial, 1, 2))
                    Else
                        _HI = CInt(Mid(_HoraInicial, 1, 1))
                    End If
                    If Len(_HoraFinal) = 5 Then
                        _HF = CInt(Mid(_HoraFinal, 1, 2))
                    Else
                        _HF = CInt(Mid(_HoraFinal, 1, 1))
                    End If
                    If _HI >= _HF Then
                        Msg = "La hora Inicial no puede ser Mayor a la Hora final. Verifique"
                        Style = MsgBoxStyle.Information
                        MsgBox(Msg, Style)
                        MsgBox("Checa " & _HoraInicial & " " & _HoraFinal)
                        Exit Sub
                    End If
                    If _FechaInicial > _FechaFinal Then
                        Msg = "La Fecha Inicial no puede ser Mayor a la Fecha final. Verifique"
                        Style = MsgBoxStyle.Information
                        MsgBox(Msg, Style)
                        Exit Sub
                    End If

                    _Dias = (_FechaFinal - _FechaInicial).TotalDays + 1 

                    For x = 1 To _Dias
                        If _FechaInicial <= _FechaFinal Then
                            _NombreDiaSemana = WeekdayName(Weekday(_FechaInicial))
                            _NombreDiaSemana = UCase(_NombreDiaSemana)

                            '   MsgBox("NombreDiaSemana " & _NombreDiaSemana & " " & _NombreDiaReservacion)

                            If _NombreDiaSemana = _NombreDiaReservacion Then

                                ' Verifica que no haya colisión en horarios

                                _Fecha = _FechaInicial
                                _Recurso = TextBox10.Text
                                If Len(_HoraInicial) = 5 Then
                                    _HoraInicioSolicitada = CInt(Mid(_HoraInicial, 1, 2))
                                Else
                                    _HoraFinalSolicitada = CInt(Mid(_HoraInicial, 1, 1))
                                End If
                                If Len(_HoraFinal) = 5 Then
                                    _HoraFinalSolicitada = CInt(Mid(_HoraFinal, 1, 2))
                                Else
                                    _HoraFinalSolicitada = CInt(Mid(_HoraFinal, 1, 1))
                                End If

                                _Colision = False
                                _Rows_Re = CServiciosDataSet38.Reservaciones.Select("Fecha = " & "'" & _FechaInicial & "'")
                                For _N = 0 To _Rows_Re.GetUpperBound(0)
                                    If Len(_Rows_Re(_N).Item("HoraInicio")) = 5 Then
                                        _HoraInicioReservada = CInt(Mid(_Rows_Re(_N).Item("HoraInicio"), 1, 2))
                                    Else
                                        _HoraInicioReservada = CInt(Mid(_Rows_Re(_N).Item("HoraInicio"), 1, 1))
                                    End If
                                    If Len(_Rows_Re(_N).Item("HoraFinal")) = 5 Then
                                        _HoraFinalReservada = CInt(Mid(_Rows_Re(_N).Item("HoraFinal"), 1, 2))
                                    Else
                                        _HoraFinalReservada = CInt(Mid(_Rows_Re(_N).Item("HoraFinal"), 1, 1))
                                    End If

                                    If _HoraInicioSolicitada > _HoraInicioReservada And _HoraInicioSolicitada < _HoraFinalReservada Then
                                        _Colision = True
                                    End If

                                    If _HoraFinalSolicitada > _HoraInicioReservada And _HoraInicioSolicitada < _HoraFinalReservada Then
                                        _Colision = True
                                    End If

                                Next _N

                                If _Colision = True Then
                                    Msg = "La reservación del dia " & CStr(_FechaInicial) & " No se puede realizar por estar Ocupado el Recursos ese Dia a esa Hora"
                                    Style = MsgBoxStyle.Information
                                    MsgBox(Msg, Style)
                                Else

                                    _Xconsecutivo = "RS" & Trim(CStr(_Consecutivo))

                                    _Busca = CServiciosDataSet38.Reservaciones.NewRow
                                    _Busca.Item("Reservacion") = _Xconsecutivo
                                    _Busca.Item("Fecha") = _FechaInicial
                                    _Busca.Item("HoraInicio") = _HoraInicial
                                    _Busca.Item("HoraFinal") = _HoraFinal
                                    _Busca.Item("Recurso") = ToolStripTextBox1.Text
                                    _Busca.Item("Referencia") = TextBox23.Text
                                    _Busca.Item("Motivo") = TextBox22.Text
                                    _Busca.Item("Asistentes") = CInt(TextBox24.Text)
                                    _Busca.Item("RequiereAsistencia") = CheckBox8.Checked
                                    _Busca.Item("Comentarios") = ""
                                    _Busca.Item("FechaRegistro") = Today
                                    _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
                                    _Busca.Item("Solicitante") = TextBox21.Text
                                    _Busca.Item("TelefonoSolicitante") = TextBox20.Text
                                    _Busca.Item("CelularSolicitante") = TextBox18.Text
                                    _Busca.Item("Autoriza") = TextBox16.Text
                                    _Busca.Item("DepartamentoSolicitante") = TextBox19.Text
                                    _Busca.Item("Asistencia") = False
                                    _Busca.Item("Satisfactorio") = False
                                    _Busca.Item("Usuario") = TextBox17.Text

                                    CServiciosDataSet38.Reservaciones.Rows.Add(_Busca)

                                    Me.Validate()
                                    ReservacionesBindingSource.EndEdit()
                                    Me.TableAdapterManager.UpdateAll(CServiciosDataSet38)
                                    Me.ReservacionesTableAdapter.Fill(Me.CServiciosDataSet38.Reservaciones)

                                    _Guarda = True
                                    _Consecutivo = _Consecutivo + 1
                                    _FechaInicial = (_FechaInicial.AddDays(6))

                                    Msg = "Reservación " & _Xconsecutivo & " efectuada correctamente"
                                    Style = MsgBoxStyle.Information
                                    MsgBox(Msg, Style)
                                End If
                        End If

                        End If

                        _FechaInicial = (_FechaInicial.AddDays(1))
                    Next x

                End If

            Next I

          

            If _Guarda = False Then
                If _Find = False Then
                    Msg = "Debe seleccionar un dia de la semana para realizas las Reservaciones. No se registraron "
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                    Exit Sub
                End If
            End If

            If _Guarda Then
                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("RES")
                _Busca.Item("Consecutivo") = _Consecutivo

                Me.Validate()
                DocumentosBindingSource.EndEdit()
                Me.TableAdapterManager2.UpdateAll(CServiciosDataSet3)

                Msg = "Proceso Terminado Correctamente"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Me.Close()

            End If

        Else
            Exit Sub
        End If


    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        For Me.I = 0 To CServiciosDataSet36.Recursos.Rows.Count - 1
            If CServiciosDataSet36.Recursos.Rows(I).Item("DescripcionLarga") = ComboBox3.Text Then
                ToolStripTextBox1.Text = CServiciosDataSet36.Recursos.Rows(I).Item("ReccursoServicio")
            End If
        Next
        TabControl1.Enabled = True
    End Sub
End Class