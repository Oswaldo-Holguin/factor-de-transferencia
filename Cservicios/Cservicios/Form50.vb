Public Class Frm50Solicitudes
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Busca As DataRow
    Public _Solicitud, _Hora, _Solicitante, _DepartamentoSolicitante, _TelefonoSolicitante, _CelularSolicitante, _Descripcion As String
    Public _Referencia, _Motivo, _Prioridad, _Ubicacion, _Atiende, _FechaSolicitada, _HoraSolicitada, _MotivoCancelacion As String
    Public _Fecha, _FechaCancelacion As Date
    Public _Atendida, _Material As Boolean
    Public _ChecaControl As Control

    Private Sub Label19_Click(sender As System.Object, e As System.EventArgs) Handles Label19.Click

    End Sub

    Private Sub Frm50Solicitudes_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
 
        Button1_Click(sender, e)


    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet41.Solicitudes' Puede moverla o quitarla según sea necesario.
        '  Me.SolicitudesTableAdapter.Fill(Me.CServiciosDataSet41.Solicitudes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet43.Solicitudes' Puede moverla o quitarla según sea necesario.
        Me.SolicitudesTableAdapter1.Fill(Me.CServiciosDataSet43.Solicitudes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet42.Prioridades' Puede moverla o quitarla según sea necesario.
        Me.PrioridadesTableAdapter.Fill(Me.CServiciosDataSet42.Prioridades)

        Label18.Text = ""
        For Each Me._ChecaControl In Me.Controls
            _ChecaControl.BackColor = Color.White
            _ChecaControl.ForeColor = Color.Black
        Next
        MonthCalendar1.Visible = False
        For Each Me._ChecaControl In Me.Controls
            If TypeOf _ChecaControl Is TextBox Then
                _ChecaControl.Text = ""
                _ChecaControl.Enabled = False
            End If
        Next
        RichTextBox1.Text = ""
        RichTextBox1.Enabled = False

        ToolStripButton3.Enabled = True
        ToolStripButton4.Enabled = False
        ToolStripButton5.Enabled = False

        ToolStripButton7.Enabled = False
        ComboBox1.Text = ""
        ComboBox1.Enabled = False


        DataGridView1.Rows.Clear()
        For Me.I = 0 To CServiciosDataSet43.Solicitudes.Rows.Count - 1
            _Solicitud = CServiciosDataSet43.Solicitudes.Rows(I).Item("Solicitud")
            _Solicitante = CServiciosDataSet43.Solicitudes.Rows(I).Item("Solicitante")
            _Fecha = CServiciosDataSet43.Solicitudes.Rows(I).Item("Fecha")

            DataGridView1.Rows.Add(_Solicitud, _Solicitante, _Fecha)

        Next

        For Me.I = 0 To CServiciosDataSet42.Prioridades.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet42.Prioridades.Rows(I).Item("DescripcionLarga"))
        Next



    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus
        ' Obtiene el siguiente Consecutivo
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("SOL")
        Dim _Consecutivo As Integer = _Busca.Item("Consecutivo")
        Dim _XConsecutivo As String = "SS" & Trim(CStr(_Consecutivo))
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        ComboBox1.Focus()

        TextBox2.Text = Today
        TextBox3.Text = Mid(TimeOfDay, 1, 5)

        TextBox4.Focus()


    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True

        For Each Me._ChecaControl In Me.Controls
            If TypeOf _ChecaControl Is TextBox Then
                _ChecaControl.Enabled = True
                _ChecaControl.Text = ""
            End If
        Next
        RichTextBox1.Text = ""
        ComboBox1.Text = ""

        RichTextBox1.Enabled = True
        TextBox12.Enabled = False
        TextBox13.Enabled = False
        TextBox14.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        ComboBox1.Enabled = True
        TextBox17.Enabled = False
        TextBox18.Enabled = False
        '  Button1_Click(sender, e)

        TextBox1.Focus()
        TextBox19.Text = "ALTA"
        Timer1.Enabled = True
        Timer1.Interval = 100

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub RichTextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
        RichTextBox1.BackColor = Color.White
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click

        ToolStripButton4.Enabled = False
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de registrar esta Solicitud de Servicio?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Solicitudes de Servicio a Departamento de Mantenimiento"   
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' 'Revisa campos de fechas
            If TextBox19.Text = "ALTA" Or TextBox19.Text = "CAMBIO" Then
                If Not IsDate(TextBox15.Text) Then
                    Msg = "Debe teclear una fecha solicitada válida. Verifique"
                    TextBox15.Focus()
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                    Exit Sub
                End If
            End If

            If TextBox19.Text = "CANCELA" Then
                If Not IsDate(TextBox17.Text) Then
                    TextBox17.Focus()
                    Msg = "Debe teclear una fecha de Cancelación Válida"
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                    Exit Sub
                End If
            End If

            If TextBox19.Text = "ATIENDE" Then
                If Not IsDate(TextBox12.Text) Then
                    TextBox12.Focus()
                    Msg = "Debe teclear una Fecha de Inicio válida. Verifique"
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                    Exit Sub
                End If

                If Not IsDate(TextBox13.Text) Then
                    Msg = "Debe teclear una Fecha Final válida. Verifique"
                    TextBox13.Focus()
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                    Exit Sub
                End If
            End If



            ' Revisa que no haya casillas en blanco
            For Each Me._ChecaControl In Me.Controls
                If TypeOf _ChecaControl Is TextBox Then
                    If _ChecaControl.Enabled = True Then
                        If _ChecaControl.Text = "" Then
                            Msg = "Hay casillas en blanco. Debe proporcionar todos los Datos " & _ChecaControl.Name
                            Style = MsgBoxStyle.Information
                            MsgBox(Msg, Style)
                            _ChecaControl.Focus()
                            ToolStripButton4.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
            Next

            _Solicitud = TextBox1.Text
            If TextBox19.Text = "ALTA" Then
                _Busca = CServiciosDataSet43.Solicitudes.NewRow
                _Busca.Item("Solicitud") = _Solicitud
            End If

            If TextBox19.Text = "CAMBIO" Or TextBox19.Text = "ATIENDE" Or TextBox19.Text = "CANCELA" Then
                _Busca = CServiciosDataSet43.Solicitudes.FindBySolicitud(_Solicitud)
            End If

            If TextBox19.Text = "CANCELA" Then
                _Busca.Item("FechaCancelacion") = CDate(TextBox17.Text)
                _Busca.Item("MotivoCancelacion") = TextBox18.Text
            End If

            If TextBox19.Text = "CAMBIO" Or TextBox19.Text = "ATIENDE" Or TextBox19.Text = "ALTA" Then
                _Busca.Item("Fecha") = CDate(TextBox2.Text)
                _Busca.Item("Hora") = TextBox3.Text
                _Busca.Item("Solicitante") = TextBox4.Text
                _Busca.Item("DepartamentoSolicitante") = TextBox5.Text
                _Busca.Item("TelefonoSolicitante") = TextBox6.Text
                _Busca.Item("CelularSolicitante") = TextBox7.Text
                _Busca.Item("Descripción") = UCase(RichTextBox1.Text)
                _Busca.Item("Referencia") = TextBox8.Text
                _Busca.Item("Motivo") = TextBox9.Text
                _Busca.Item("Prioridad") = TextBox10.Text
                _Busca.Item("Ubicacion") = TextBox11.Text
                '   _Busca.Item("FechaInicio") = TextBox12.Text
                '   _Busca.Item("FechaFinal") = TextBox13.Text
                _Busca.Item("Atiende") = TextBox14.Text
                _Busca.Item("FechaSolicitada") = TextBox15.Text
                _Busca.Item("HoraSolicitada") = TextBox16.Text
                _Busca.Item("MotivoCancelacion") = TextBox18.Text
                _Busca.Item("Atendida") = CheckBox1.Checked
                _Busca.Item("Material") = CheckBox2.Checked

                If TextBox19.Text = "ATIENDE" Then
                    _Busca.Item("FechaInicio") = CDate(TextBox12.Text)
                    _Busca.Item("FechaFinal") = CDate(TextBox13.Text)
                End If
                If TextBox19.Text = "ALTA" Then
                    CServiciosDataSet43.Solicitudes.Rows.Add(_Busca)
                End If
            End If

            Me.Validate()
            SolicitudesBindingSource1.EndEdit()
            Me.TableAdapterManager3.UpdateAll(CServiciosDataSet43)
            Me.SolicitudesTableAdapter1.Fill(Me.CServiciosDataSet43.Solicitudes)

            If TextBox19.Text = "ALTA" Then
                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("SOL")
                _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

                Me.Validate()
                DocumentosBindingSource.EndEdit()
                Me.TableAdapterManager.UpdateAll(CServiciosDataSet3)
                Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
            End If

            Msg = "Registro guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            MonthCalendar1.Visible = False
            Timer1.Enabled = False

            Button1_Click(sender, e)



        Else
            Exit Sub
        End If



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.I = 0 To CServiciosDataSet42.Prioridades.Rows.Count - 1
            If CServiciosDataSet42.Prioridades.Rows(I).Item("DescripcionLarga") = ComboBox1.Text Then
                TextBox10.Text = CServiciosDataSet42.Prioridades.Rows(I).Item("Prioridad")
            End If
        Next
        TextBox10.Enabled = False
        TextBox11.Focus()
        ComboBox1.BackColor = Color.White

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If


        For Each Me._ChecaControl In Me.Controls
            _ChecaControl.BackColor = Color.White
            _ChecaControl.ForeColor = Color.Black
        Next



        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _Solicitud = _Renglon.Cells(0).Value

        _Busca = CServiciosDataSet43.Solicitudes.FindBySolicitud(_Solicitud)

        Label18.Text = "Solicitud VIGENTE"
        Label18.BackColor = Color.Green
        Label18.ForeColor = Color.White
        TextBox1.Text = _Solicitud
        TextBox2.Text = _Busca.Item("Fecha")
        TextBox3.Text = _Busca.Item("Hora")
        TextBox4.Text = _Busca.Item("Solicitante")
        TextBox5.Text = _Busca.Item("DepartamentoSolicitante")
        TextBox6.Text = _Busca.Item("TelefonoSolicitante")
        TextBox7.Text = _Busca.Item("CelularSolicitante")
        RichTextBox1.Text = _Busca.Item("Descripción")
        TextBox8.Text = _Busca.Item("Referencia")
        TextBox9.Text = _Busca.Item("Motivo")
        TextBox10.Text = _Busca.Item("Prioridad")
        _Prioridad = TextBox10.Text
        TextBox11.Text = _Busca.Item("Ubicacion")
        If DBNull.Value.Equals(_Busca.Item("FechaInicio")) Then
            TextBox12.Text = ""
        Else
            TextBox12.Text = _Busca.Item("FechaInicio")
        End If
        If DBNull.Value.Equals(_Busca.Item("FechaFinal")) Then
            TextBox13.Text = ""
        Else
            TextBox13.Text = _Busca.Item("FechaFinal")
        End If


        TextBox14.Text = _Busca.Item("Atiende")
        TextBox15.Text = _Busca.Item("FechaSolicitada")
        TextBox16.Text = _Busca.Item("HoraSolicitada")
        If DBNull.Value.Equals(_Busca.Item("FechaCancelacion")) Then
            TextBox17.Text = ""
        Else
            TextBox17.Text = _Busca.Item("FechaCancelacion")
            Label18.Text = "Solicitud CANCELADA"
            Label18.BackColor = Color.Red
            Label18.ForeColor = Color.White
        End If

        If DBNull.Value.Equals(_Busca.Item("MotivoCancelacion")) Then
            TextBox18.Text = ""
        Else
            TextBox18.Text = _Busca.Item("MotivoCancelacion")
        End If

        CheckBox1.Checked = _Busca.Item("Atendida")
        If CheckBox1.Checked = True Then
            CheckBox1.ForeColor = Color.PeachPuff
            Label18.Text = "Solicitud ATENDIDA"
            Label18.BackColor = Color.PeachPuff
            Label18.ForeColor = Color.Blue
            CheckBox1.BackColor = Color.PeachPuff
            CheckBox1.ForeColor = Color.Blue
            TextBox12.BackColor = Color.PeachPuff
            TextBox13.BackColor = Color.PeachPuff
            TextBox14.BackColor = Color.PeachPuff
        End If

        CheckBox2.Checked = _Busca.Item("Material")

        _Busca = CServiciosDataSet42.Prioridades.FindByPrioridad(_Prioridad)
        ComboBox1.Text = _Busca.Item("DescripcionLarga")

        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = True
        ToolStripButton7.Enabled = True
        RichTextBox1.Enabled = True
        For Each Me._ChecaControl In Me.Controls
            If TypeOf _ChecaControl Is TextBox Then
                If Label18.Text = "Solicitud VIGENTE" Then
                    _ChecaControl.Enabled = True

                    ComboBox1.Enabled = True
                    TextBox12.Enabled = False
                    TextBox13.Enabled = False
                    CheckBox1.Enabled = False
                    TextBox14.Enabled = False
                    TextBox17.Enabled = False
                    TextBox18.Enabled = False
                End If
                If Label18.Text = "Solicitud ATENDIDA" Then
                    _ChecaControl.Enabled = True
                    ComboBox1.Enabled = True
                    CheckBox1.Enabled = True
                    ToolStripButton5.Enabled = False
                    ToolStripButton7.Enabled = False
                    ToolStripButton10.Enabled = False
                End If
                If Label18.Text = "Solicitud CANCELADA" Then
                    _ChecaControl.Enabled = True
                    ComboBox1.Enabled = True
                    CheckBox1.Enabled = True
                    ToolStripButton5.Enabled = False
                    ToolStripButton7.Enabled = False
                    ToolStripButton10.Enabled = False
                    TextBox17.BackColor = Color.Red
                    TextBox17.ForeColor = Color.White
                    TextBox18.BackColor = Color.Red
                    TextBox18.ForeColor = Color.White
                End If
            End If
        Next


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        TextBox19.Text = "CANCELA"
        TextBox17.Enabled = True
        TextBox18.Enabled = True
        Msg = "Proporciona la Fecha y Motivo de la Cancelación"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)
        MonthCalendar1.Visible = True
        TextBox20.Text = "17"
        TextBox19.Text = "CANCELA"

        For Each Me._ChecaControl In Me.Controls
            If TypeOf _ChecaControl Is TextBox Then
                _ChecaControl.Enabled = False
            End If
        Next
        RichTextBox1.Enabled = False
        ComboBox1.Enabled = False
        CheckBox1.Enabled = False

        TextBox17.Enabled = True
        TextBox18.Enabled = True

        ToolStripButton4.Enabled = True
        ToolStripButton3.Enabled = False
        ToolStripButton8.Enabled = False
        ToolStripButton10.Enabled = False
        ToolStripButton5.Enabled = False
        Timer1.Enabled = True
        Timer1.Interval = 100
        TextBox17.Focus()


    End Sub

    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click

        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar una Solicitud de Servicio para darla como ATENDIDA"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        TextBox19.Text = "ATIENDE"
        TextBox12.Enabled = True
        TextBox13.Enabled = True
        TextBox14.Enabled = True
        CheckBox1.Checked = True
        TextBox12.Focus()
        ToolStripButton4.Enabled = True
        Timer1.Enabled = True
        Timer1.Interval = 100

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click

        For Each Me._ChecaControl In Me.Controls
            If TypeOf _ChecaControl Is TextBox Then
                _ChecaControl.Enabled = True
            End If
        Next
        TextBox1.Enabled = False
        ToolStripButton5.Enabled = False
        ToolStripButton4.Enabled = True
        RichTextBox1.Enabled = True
        TextBox19.Text = "CAMBIO"
        TextBox20.Text = "X"

        TextBox12.Enabled = False
        TextBox13.Enabled = False
        TextBox14.Enabled = False
        TextBox17.Enabled = False
        TextBox18.Enabled = False
        ComboBox1.Enabled = True

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Button1_Click(sender, e)

    End Sub

    Private Sub TextBox15_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox15.GotFocus
        TextBox20.Text = "15"
        MonthCalendar1.Visible = True

    End Sub

    Private Sub TextBox15_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox15.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox16.Focus()
        End If
    End Sub

    Private Sub TextBox15_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox15.LostFocus
        TextBox15.BackColor = Color.White
    End Sub

    Private Sub TextBox15_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox12_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox12.GotFocus
        TextBox20.Text = "12"
        MonthCalendar1.Visible = True

    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox13.GotFocus
        TextBox20.Text = "13"
        MonthCalendar1.Visible = True

    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        If TextBox20.Text = "12" Then
            TextBox12.Text = MonthCalendar1.SelectionRange.Start
            TextBox12.BackColor = Color.White
        End If
        If TextBox20.Text = "13" Then
            TextBox13.Text = MonthCalendar1.SelectionRange.Start
            If CDate(TextBox13.Text) < CDate(TextBox12.Text) Then
                Msg = "La Fecha Final no puede ser menor a la Fecha Inicial"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If
            TextBox13.BackColor = Color.White
            TextBox14.Focus()

        End If
        If TextBox20.Text = "15" Then
            TextBox15.Text = MonthCalendar1.SelectionRange.Start
            TextBox15.BackColor = Color.White
            TextBox16.Focus()
        End If
        If TextBox20.Text = "17" Then
            TextBox17.Text = MonthCalendar1.SelectionRange.Start
            TextBox17.BackColor = Color.White
            TextBox18.Focus()
        End If

        MonthCalendar1.Visible = False

    End Sub

    Private Sub TextBox17_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox17.GotFocus
        TextBox20.Text = "17"
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox17_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ' Búsqueda
        Label18.Text = ""
        DataGridView1.Rows.Clear()
        Dim _Rows_Ss() As DataRow
        For Each Me._ChecaControl In Me.Controls
            If TypeOf _ChecaControl Is TextBox Then
                _ChecaControl.Text = ""
            End If
        Next
        RichTextBox1.Text = ""

        ' Por Solicitud
        _Solicitud = ""
        If RadioButton1.Checked = True Then

            _Solicitud = InputBox("Teclee la Solicitud a Buscar")
            _Solicitud = Trim(UCase(_Solicitud))

            _Rows_Ss = CServiciosDataSet43.Solicitudes.Select("Solicitud = " & "'" & _Solicitud & "'")

        End If

        ' Por Fecha
        If RadioButton2.Checked = True Then
            _Fecha = InputBox("Teclee la Fecha a Buscar")
            '   If _Fecha = "" Then
            ' Exit Sub
            ' End If
            If Not IsDate(_Fecha) Then
                Msg = "Debe teclear una fecha válida "
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If
            _Rows_Ss = CServiciosDataSet43.Solicitudes.Select("Fecha = " & "'" & _Fecha & "'")
        End If


        ' Por Estatus 
        If RadioButton4.Checked = True Then
            Dim _Estatus As String
            Dim _Seleccion As String
            _Estatus = InputBox("1 = VIGENTES    2 = ATENDIDAS   3 = CANCELADAS                ")
            If _Estatus <> "1" And _Estatus <> "2" And _Estatus <> "3" Then
                Msg = "Debe seleccionar el Estatus Correcto"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            If _Estatus = "1" Then
                _Seleccion = "Atendida = False and MotivoCancelacion = ''"
                _Rows_Ss = CServiciosDataSet43.Solicitudes.Select(_Seleccion)
            End If
            If _Estatus = "2" Then
                _Seleccion = "Atendida = True"
                _Rows_Ss = CServiciosDataSet43.Solicitudes.Select(_Seleccion)
            End If
            If _Estatus = "3" Then
                _Seleccion = "MotivoCancelacion <> ''"
                _Rows_Ss = CServiciosDataSet43.Solicitudes.Select(_Seleccion)
            End If

        End If

        ' Todas
        If RadioButton5.Checked = True Then
            _Rows_Ss = CServiciosDataSet43.Solicitudes.Select("Solicitud <> ''")

        End If


        If RadioButton3.Checked = False Then

            DataGridView1.Rows.Clear()
            For Me.I = 0 To _Rows_Ss.GetUpperBound(0)
                _Solicitud = _Rows_Ss(I).Item("Solicitud")
                _Solicitante = _Rows_Ss(I).Item("Solicitante")
                _Fecha = _Rows_Ss(I).Item("Fecha")

                DataGridView1.Rows.Add(_Solicitud, _Solicitante, _Fecha)
            Next
        End If


        ' Por Solicitante
        If RadioButton3.Checked = True Then
            Dim _Mide As Integer
            Dim _Texto, _Partenombre, _Nombre As String
            _Solicitante = InputBox("Teclee el Nombre del Solicitante a Buscar")
            _Solicitante = Trim(UCase(_Solicitante))

            _Mide = Len(_Solicitante)
            If _Mide = 0 Then
                Exit Sub
            End If

            For Me.I = 0 To CServiciosDataSet43.Solicitudes.Rows.Count - 1
                _Texto = CServiciosDataSet43.Solicitudes.Rows(I).Item("Solicitante")
                _Solicitud = CServiciosDataSet43.Solicitudes.Rows(I).Item("Solicitud")
                _Nombre = CServiciosDataSet43.Solicitudes.Rows(I).Item("Solicitante")
                _Fecha = CServiciosDataSet43.Solicitudes.Rows(I).Item("Fecha")

                For x = 1 To Len(_Texto)
                    _Partenombre = Mid(_Texto, x, _Mide)
                    If _Partenombre = _Solicitante Then

                        DataGridView1.Rows.Add(_Solicitud, _Nombre, _Fecha)
                    End If
                Next


            Next

        End If







    End Sub

    Private Sub ToolStripButton10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton10.Click
        ' Imprime la solicitud

        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar una Solicitud para Imprimirla"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Solicitud de Servicio.xlsx"
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de exce

        _Fecha = TextBox2.Text
        _Solicitante = TextBox4.Text
        _DepartamentoSolicitante = TextBox5.Text
        _Motivo = TextBox9.Text
        _Ubicacion = TextBox11.Text
        _Prioridad = ComboBox1.Text
        Dim _Comentarios As String = RichTextBox1.Text

        m_Excel.Worksheets("SOL").Cells(7, 2).value = _Solicitud
        m_Excel.Worksheets("SOL").Cells(7, 5).value = _Fecha
        m_Excel.Worksheets("SOL").Cells(8, 5).value = _Solicitante
        m_Excel.Worksheets("SOL").Cells(9, 5).value = _DepartamentoSolicitante
        m_Excel.Worksheets("SOL").Cells(10, 5).value = _Motivo
        m_Excel.Worksheets("SOL").Cells(11, 5).value = _Ubicacion
        m_Excel.Worksheets("SOL").Cells(12, 5).value = _Prioridad
        m_Excel.Worksheets("SOL").Cells(15, 1).value = _Comentarios

        m_Excel.Visible = True

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        For Each Me._ChecaControl In Me.Controls
            If TypeOf _ChecaControl Is TextBox Or TypeOf _ChecaControl Is ComboBox Or TypeOf _ChecaControl Is RichTextBox Then
                If _ChecaControl.Focused Then
                    _ChecaControl.BackColor = Color.Gold
                End If
            End If
        Next
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox4_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox6.LostFocus
        TextBox6.BackColor = Color.White
    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox8.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox9.Focus()
        End If
    End Sub

    Private Sub TextBox8_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox9_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox9.LostFocus
        TextBox9.BackColor = Color.White
    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox11.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox15.Focus()
        End If
    End Sub

    Private Sub TextBox11_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox11.LostFocus
        TextBox11.BackColor = Color.White
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox16_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox16.KeyPress
        If Asc(e.KeyChar) = 13 Then
            RichTextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox16_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox16.LostFocus
        TextBox16.BackColor = Color.White
    End Sub

    Private Sub TextBox16_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox14_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox14.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox14.BackColor = Color.White
        End If
    End Sub

    Private Sub TextBox14_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox14.LostFocus
        TextBox14.BackColor = Color.White
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub
End Class