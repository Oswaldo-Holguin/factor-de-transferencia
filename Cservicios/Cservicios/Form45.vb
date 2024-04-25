Public Class Frm45Reservaciones
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Busca As DataRow
    Public _Reservacion, _HoraInicio, _HoraFinal, _Recurso, _Referencia, _Motivo, _Comentarios As String
    Public _Solicitante, _TelefonoSolicitante, _CelularSolicitante, _DepartamentoSolicitante As String
    Public _Autoriza, _Usuario As String
    Public _Asistencia, _Satisfactorio As Boolean
    Public _Asistentes As Integer
    Public _Fecha As Date
    Public _Checacontrol As Control


    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub Frm45Reservaciones_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet37.Equipos' Puede moverla o quitarla según sea necesario.
        Me.EquiposTableAdapter.Fill(Me.CServiciosDataSet37.Equipos)

        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet38.Reservaciones' Puede moverla o quitarla según sea necesario.
        Me.ReservacionesTableAdapter.Fill(Me.CServiciosDataSet38.Reservaciones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet36.Recursos' Puede moverla o quitarla según sea necesario.
        Me.RecursosTableAdapter.Fill(Me.CServiciosDataSet36.Recursos)

        ComboBox1.Items.Clear()
        For Me.I = 0 To CServiciosDataSet36.Recursos.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet36.Recursos.Rows(I).Item("Descripcionlarga"))
        Next
        Dim _DescripcionLarga As String
        Dim X As String
        DataGridView1.Rows.Clear()
        For Me.I = 0 To CServiciosDataSet38.Reservaciones.Rows.Count - 1
            _Reservacion = CServiciosDataSet38.Reservaciones.Rows(I).Item("Reservacion")

            _Recurso = CServiciosDataSet38.Reservaciones.Rows(I).Item("Recurso")

            _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
            _DescripcionLarga = _Busca.Item("DescripcionLarga")

            _Fecha = CServiciosDataSet38.Reservaciones.Rows(I).Item("Fecha")

            DataGridView1.Rows.Add(_Reservacion, _DescripcionLarga, _Fecha)
        Next

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If

        Next

        Label12.Text = ""
        For Me.I = 0 To 24
            X = CStr(I) & ":" & "00"
            ComboBox2.Items.Add(X)
        Next
        For Me.I = 0 To 24
            X = CStr(I) & ":" & "00"
            ComboBox3.Items.Add(X)
        Next

        ToolStripButton1.Enabled = True
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True
        ToolStripButton5.Enabled = False
        ToolStripButton6.Enabled = False
        ToolStripButton8.Enabled = False
        RichTextBox1.Text = ""
        RichTextBox1.Visible = False

        Panel1.Enabled = False
        Button3.Enabled = False

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        ' Registro de una nueva Reservación
        Panel1.Enabled = True
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
            End If
        Next

        TextBox14.Text = "ALTA"
        TextBox1.Focus()
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        TextBox13.Text = LoginForm1.TextBox6.Text

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus
        ' Obtiene el siguiente Consecutivo
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("RES")
        Dim _Consecutivo As Integer = _Busca.Item("Consecutivo")
        Dim _XConsecutivo As String = "RS" & Trim(CStr(_Consecutivo))
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        ComboBox1.Focus()



    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged


    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox2.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guardar Datos

        ToolStripButton3.Enabled = False
        Dim _Rows_Re() As DataRow
        Dim _HoraInicioSolicitada, _HoraFinalSolicitada, _HoraInicioReservada, _HoraFinalReservada As Integer
        Dim _Find As Boolean
        _Find = False

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = UCase(_Checacontrol.Text)
                If _Checacontrol.Text = "" Then
                    _Find = True
                End If

            End If
        Next

        If _Find Then
            Msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
            Style = MsgBoxStyle.Information
            ToolStripButton3.Enabled = True
            Exit Sub
        End If


        ' Verifica que no haya reservación ese dia en ese horario
        _Reservacion = TextBox1.Text
        _Fecha = CDate(TextBox2.Text)
        _HoraInicio = ComboBox2.Text
        _HoraFinal = ComboBox3.Text
        _Recurso = TextBox10.Text
        If Len(_HoraInicio) = 5 Then
            _HoraInicioSolicitada = CInt(Mid(_HoraInicio, 1, 2))
        Else
            _HoraFinalSolicitada = CInt(Mid(_HoraInicio, 1, 1))
        End If
        If Len(_HoraFinal) = 5 Then
            _HoraFinalSolicitada = CInt(Mid(_HoraFinal, 1, 2))
        Else
            _HoraFinalSolicitada = CInt(Mid(_HoraFinal, 1, 1))
        End If

        _Rows_Re = CServiciosDataSet38.Reservaciones.Select("Recurso = " & "'" & _Recurso & "'")

        Style = MsgBoxStyle.Information
        _Find = False
        For Me.I = 0 To _Rows_Re.GetUpperBound(0)
            If _Rows_Re(I).Item("Reservacion") <> _Reservacion Then
                If _Rows_Re(I).Item("Fecha") = _Fecha Then

                    If Len(_Rows_Re(I).Item("HoraInicio")) = 5 Then
                        _HoraInicioReservada = CInt(Mid(_Rows_Re(I).Item("HoraInicio"), 1, 2))
                    Else
                        _HoraInicioReservada = CInt(Mid(_Rows_Re(I).Item("HoraInicio"), 1, 1))
                    End If
                    If Len(_Rows_Re(I).Item("HoraFinal")) = 5 Then
                        _HoraFinalReservada = CInt(Mid(_Rows_Re(I).Item("HoraFinal"), 1, 2))
                    Else
                        _HoraFinalReservada = CInt(Mid(_Rows_Re(I).Item("HoraFinal"), 1, 1))
                    End If

                    If _HoraInicioSolicitada > _HoraInicioReservada And _HoraInicioSolicitada < _HoraFinalReservada Then
                        Msg = "La hora de Inicio Solicitada, se encuentra ya Reservada para " & _Rows_Re(I).Item("Solicitante") & " Reservación No. " & _Rows_Re(I).Item("Reservacion")
                        MsgBox(Msg, Style)
                        Exit Sub
                    End If

                    If _HoraFinalSolicitada > _HoraInicioReservada And _HoraInicioSolicitada < _HoraFinalReservada Then
                        Msg = "La hora Final Solicitada, se encuentra ya Reservada para " & _Rows_Re(I).Item("Solicitante") & " Reservación No. " & _Rows_Re(I).Item("Reservacion")
                        MsgBox(Msg, Style)
                        Exit Sub
                    End If

                End If
            End If

        Next

        _Reservacion = TextBox1.Text
        If TextBox14.Text = "ALTA" Then
            _Busca = CServiciosDataSet38.Reservaciones.NewRow
            _Busca.Item("Reservacion") = _Reservacion
        End If

        If TextBox14.Text = "CAMBIO" Or TextBox14.Text = "ASISTE" Then
            _Busca = CServiciosDataSet38.Reservaciones.FindByReservacion(_Reservacion)
        End If

        If TextBox14.Text <> "ASISTE" Then
            _Busca.Item("Fecha") = CDate(TextBox2.Text)
            _Busca.Item("HoraInicio") = ComboBox2.Text
            _Busca.Item("HoraFinal") = ComboBox3.Text
            _Busca.Item("Recurso") = TextBox10.Text
            _Busca.Item("Referencia") = TextBox4.Text
            _Busca.Item("Motivo") = TextBox5.Text
            _Busca.Item("Asistentes") = CInt(TextBox3.Text)
            _Busca.Item("RequiereAsistencia") = CheckBox1.Checked
            _Busca.Item("Comentarios") = ""
            _Busca.Item("FechaRegistro") = Today
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
            _Busca.Item("Solicitante") = TextBox6.Text
            _Busca.Item("TelefonoSolicitante") = TextBox7.Text
            _Busca.Item("CelularSolicitante") = TextBox9.Text
            _Busca.Item("Autoriza") = TextBox11.Text
            _Busca.Item("DepartamentoSolicitante") = TextBox8.Text
            _Busca.Item("Asistencia") = False
            _Busca.Item("Satisfactorio") = False
            _Busca.Item("Usuario") = TextBox13.Text
        End If


        If TextBox14.Text = "ALTA" Then
            CServiciosDataSet38.Reservaciones.Rows.Add(_Busca)
        End If

        If TextBox14.Text = "ASISTE" Then
            _Busca.Item("Asistencia") = CheckBox2.Checked
            _Busca.Item("Satisfactorio") = CheckBox3.Checked
            _Busca.Item("Comentarios") = UCase(RichTextBox1.Text)
        End If


        Me.Validate()
        ReservacionesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet38)
        Me.ReservacionesTableAdapter.Fill(Me.CServiciosDataSet38.Reservaciones)


        ' Actualiza el Consecutivo
        If TextBox14.Text = "ALTA" Then
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("RES")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager2.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        End If




        Msg = "Registro guardado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        Button1_Click(sender, e)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.I = 0 To CServiciosDataSet36.Recursos.Rows.Count - 1
            If ComboBox1.Text = CServiciosDataSet36.Recursos.Rows(I).Item("DescripcionLarga") Then
                TextBox10.Text = CServiciosDataSet36.Recursos.Rows(I).Item("ReccursoServicio")
                TextBox10.Enabled = False
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        _Reservacion = _Renglon.Cells(0).Value

        _Busca = CServiciosDataSet38.Reservaciones.FindByReservacion(_Reservacion)
        TextBox1.Text = _Busca.Item("Reservacion")
        TextBox2.Text = _Busca.Item("Fecha")
        _Fecha = _Busca.Item("Fecha")
        ComboBox2.Text = _Busca.Item("HoraInicio")
        ComboBox3.Text = _Busca.Item("HoraFinal")
        TextBox10.Text = _Busca.Item("Recurso")
        TextBox4.Text = _Busca.Item("Referencia")
        TextBox5.Text = _Busca.Item("Motivo")
        TextBox3.Text = _Busca.Item("Asistentes")
        CheckBox1.Checked = -_Busca.Item("RequiereAsistencia")
        TextBox6.Text = _Busca.Item("Solicitante")
        TextBox7.Text = _Busca.Item("TelefonoSolicitante")
        TextBox9.Text = _Busca.Item("CelularSolicitante")
        TextBox11.Text = _Busca.Item("Autoriza")
        TextBox8.Text = _Busca.Item("DepartamentoSolicitante")
        CheckBox2.Checked = _Busca.Item("Asistencia")
        If CheckBox2.Checked = True Then
            CheckBox2.BackColor = Color.LightGreen
        End If
        CheckBox3.Checked = _Busca.Item("Satisfactorio")

        If CheckBox3.Checked = True Then
            CheckBox3.BackColor = Color.LightGreen
        End If

        TextBox13.Text = _Busca.Item("Usuario")
        RichTextBox1.Text = _Busca.Item("Comentarios")

        ' Forma la fecha de Reservación
 
        Dim _NombreDia As String = WeekdayName(Weekday(_Fecha))
        _NombreDia = UCase(Mid(_NombreDia, 1, 1)) & Mid(_NombreDia, 2, Len(_NombreDia))
        Dim _NumeroDia As String = Mid(CStr(_Fecha), 1, 2)
        Dim _NombreMes As String = MonthName(Month(_Fecha))
        _NombreMes = UCase(Mid(_NombreMes, 1, 1)) & Mid(_NombreMes, 2, Len(_NombreMes))
        Dim _Año As Integer = Year(_Fecha)
        Dim _NombreFecha As String = _NombreDia & " " & _NumeroDia & " de " & _NombreMes & " de " & CStr(_Año) & " de " & ComboBox2.Text & " a " & ComboBox3.Text & " Hrs."
        Label12.Text = _NombreFecha
        Label12.BackColor = Color.Gold


        _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
        ComboBox1.Text = _Busca.Item("DescripcionLarga")

        ' Presenta el equipo

        ToolStripButton12_Click(sender, e)
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = True
        ToolStripButton8.Enabled = True

        Button3.Enabled = True




    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton12_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton12.Click
        _Recurso = TextBox10.Text

        Dim _Rows_Eq() As DataRow = CServiciosDataSet37.Equipos.Select("Recurso = " & "'" & _Recurso & "'")
        Dim _Equipo, _Modelo, _Descripcion, _Serie As String

        DataGridView2.Rows.Clear()
        For Me.I = 0 To _Rows_Eq.GetUpperBound(0)
            _Equipo = _Rows_Eq(I).Item("Equipo")
            _Descripcion = _Rows_Eq(I).Item("Descripcionlarga")
            _Serie = _Rows_Eq(I).Item("Serie")
            _Modelo = _Rows_Eq(I).Item("Modelo")
            _Referencia = _Rows_Eq(I).Item("Referencia")

            DataGridView2.Rows.Add(_Equipo, _Descripcion, _Modelo, _Serie, _Referencia)

        Next
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        Frm46ReservacionesConsecutivas.Show()

    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        ' Consultas
        Dim _Rows_Re() As DataRow
        Dim _DescripcionLarga As String
        Dim _Parte, _Texto, _ParteNombre As String
        Dim _Mide As Integer

        DataGridView1.Rows.Clear()

        ' Por Fecha
        If RadioButton1.Checked = True Then
            _Fecha = InputBox("Teclee la Fecha a  a buscar")
            _Rows_Re = CServiciosDataSet38.Reservaciones.Select("Fecha = " & "'" & _Fecha & "'")
        End If

        ' Por Recurso
        If RadioButton2.Checked = True Then
            _Recurso = UCase(InputBox("Teclee la RECURSO a  a buscar"))
            _Rows_Re = CServiciosDataSet38.Reservaciones.Select("Recurso = " & "'" & _Recurso & "'")
        End If

        ' Por Reservación
        If RadioButton7.Checked = True Then
            _Reservacion = UCase(InputBox("Teclee la RESERVACION a  a buscar"))
            _Rows_Re = CServiciosDataSet38.Reservaciones.Select("Reservacion = " & "'" & _Reservacion & "'")
        End If


        If RadioButton1.Checked = True Or RadioButton2.Checked = True Or RadioButton7.Checked = True Then
            For Me.I = 0 To _Rows_Re.GetUpperBound(0)
                _Reservacion = _Rows_Re(I).Item("Reservacion")
                _Recurso = _Rows_Re(I).Item("Recurso")
                _Fecha = _Rows_Re(I).Item("Fecha")

                _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
                _DescripcionLarga = _Busca.Item("DescripcionLarga")

                DataGridView1.Rows.Add(_Reservacion, _DescripcionLarga, _Fecha)

            Next
        End If


        ' Por Solicitante y/o Motivo
        If RadioButton3.Checked = True Then
            _Parte = InputBox("Teclee el SOLICITANTE o parte del NOMBRE DEL SOLICITANTE a buscar")
            _Parte = Trim(UCase(_Parte))
            _Mide = Len(_Parte)
            If _Mide = 0 Then
                Exit Sub
            End If
        End If
        If RadioButton4.Checked = True Then
            _Parte = InputBox("Teclee el MOTIVO o parte del NOMBRE DEL MOTIVO a buscar")
            _Parte = Trim(UCase(_Parte))
            _Mide = Len(_Parte)
            If _Mide = 0 Then
                Exit Sub
            End If
        End If
        If RadioButton6.Checked = True Then
            _Parte = InputBox("Teclee el DEPARTAMENTO SOLICITANTE a buscar")
            _Parte = Trim(UCase(_Parte))
            _Mide = Len(_Parte)
            If _Mide = 0 Then
                Exit Sub
            End If
        End If


        If RadioButton3.Checked = True Or RadioButton4.Checked = True Or RadioButton6.Checked = True Then
            For Me.I = 0 To CServiciosDataSet38.Reservaciones.Rows.Count - 1
                If RadioButton3.Checked = True Then
                    _Texto = Trim(CServiciosDataSet38.Reservaciones.Rows(I).Item("Solicitante"))
                End If
                If RadioButton4.Checked = True Then
                    _Texto = Trim(CServiciosDataSet38.Reservaciones.Rows(I).Item("Motivo"))
                End If
                If RadioButton6.Checked = True Then
                    _Texto = Trim(CServiciosDataSet38.Reservaciones.Rows(I).Item("DepartamentoSolicitante"))
                End If

                For x = 1 To Len(_Texto)
                    _ParteNombre = Mid(_Texto, x, _Mide)
                    _Reservacion = CServiciosDataSet38.Reservaciones.Rows(I).Item("Reservacion")
                    _Recurso = CServiciosDataSet38.Reservaciones.Rows(I).Item("Recurso")
                    _Fecha = CServiciosDataSet38.Reservaciones.Rows(I).Item("Fecha")

                    _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
                    _DescripcionLarga = _Busca.Item("DescripcionLarga")
                    If _ParteNombre = _Parte Then
                        DataGridView1.Rows.Add(_Reservacion, _DescripcionLarga, _Fecha)
                    End If
                Next

            Next

        End If


        ' Todos
        If RadioButton5.Checked = True Then
            For Me.I = 0 To CServiciosDataSet38.Reservaciones.Rows.Count - 1
               
                _Reservacion = CServiciosDataSet38.Reservaciones.Rows(I).Item("Reservacion")
                _Recurso = CServiciosDataSet38.Reservaciones.Rows(I).Item("Recurso")
                _Fecha = CServiciosDataSet38.Reservaciones.Rows(I).Item("Fecha")

                _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
                _DescripcionLarga = _Busca.Item("DescripcionLarga")

                DataGridView1.Rows.Add(_Reservacion, _DescripcionLarga, _Fecha)

            Next
        End If

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub EliminarRegistroToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EliminarRegistroToolStripMenuItem.Click
        ' Elimina el Registro
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        Dim _Equipo As String = _Renglon.Cells(0).Value

        _Busca = CServiciosDataSet37.Equipos.FindByEquipo(_Equipo)
        _Busca.Item("Recurso") = "SU"

        Me.Validate()
        EquiposBindingSource.EndEdit()
        Me.TableAdapterManager3.UpdateAll(CServiciosDataSet37)
        Me.EquiposTableAdapter.Fill(Me.CServiciosDataSet37.Equipos)

        Msg = "Registro Eliminado correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        ToolStripButton12_Click(sender, e)


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

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar una Reservación para Imprimir el Comprobante"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        Dim _Linea As Integer
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim _Descripcion As String
        Dim strRutaExcel As String = "Z:\Reservaciones.xlsx"
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de excel


        _Reservacion = TextBox1.Text
        _Fecha = CDate(TextBox2.Text)
        _Recurso = TextBox10.Text
        _Solicitante = TextBox6.Text
        _DepartamentoSolicitante = TextBox8.Text
        _Motivo = TextBox5.Text
        _HoraInicio = ComboBox2.Text
        _HoraFinal = ComboBox3.Text


        _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
        _Descripcion = _Busca.Item("DescripcionLarga")
        _Descripcion = "(" & _Recurso & ")" & " " & _Descripcion

        m_Excel.Worksheets("FOLIO").cells(7, 2).value = _Reservacion
        m_Excel.Worksheets("FOLIO").cells(7, 5).value = _Fecha
        m_Excel.Worksheets("FOLIO").cells(7, 7).value = _HoraInicio
        m_Excel.Worksheets("FOLIO").cells(7, 9).value = _HoraFinal
        m_Excel.Worksheets("FOLIO").cells(8, 5).value = _Solicitante
        m_Excel.Worksheets("FOLIO").cells(9, 5).value = _DepartamentoSolicitante
        m_Excel.Worksheets("FOLIO").cells(10, 5).value = _Motivo
        m_Excel.Worksheets("FOLIO").cells(11, 5).value = _Descripcion

        _Linea = 14
        For Me.I = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(I).Cells(0).Value = "" Then
                Exit For
            End If
            m_Excel.Worksheets("FOLIO").cells(_Linea, 1).value = DataGridView2.Rows(I).Cells(0).Value
            m_Excel.Worksheets("FOLIO").cells(_Linea, 2).value = DataGridView2.Rows(I).Cells(1).Value
            m_Excel.Worksheets("FOLIO").cells(_Linea, 6).value = DataGridView2.Rows(I).Cells(2).Value
            m_Excel.Worksheets("FOLIO").cells(_Linea, 8).value = DataGridView2.Rows(I).Cells(3).Value

            _Linea = _Linea + 1

        Next I

        _Linea = _Linea + 2


        m_Excel.Visible = True

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        ToolStripButton3.Enabled = True
        ToolStripButton5.Enabled = False
        TextBox14.Text = "CAMBIO"
        Panel1.Enabled = True

    End Sub

    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        RichTextBox1.Visible = True

        Panel1.Enabled = True

        Msg = "Llene las casillas de SE LLEVÓ A CABO EL EVENTO y RESULTADO SATISFACTORIO y puede teclear sus COMENTARIOS"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)
        ToolStripButton8.Enabled = False
        ToolStripButton3.Enabled = True

        TextBox14.Text = "ASISTE"
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = False
            End If
        Next
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        Button3.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If RichTextBox1.Visible = False Then
            RichTextBox1.Visible = True
        Else
            RichTextBox1.Visible = False
        End If
    End Sub
End Class