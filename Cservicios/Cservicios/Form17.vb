Public Class Frm17HorarioAtencion

    Private Sub Form17_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)

        Button22_Click(sender, e)

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click



        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Está seguro de Salir ?"   ' Define message.
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Asignar Horario de Ateción a Pacientes"   ' Define title.
        ' Display message.
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then   ' User chose Yes.
            Me.Close()
           
        Else
            Exit Sub
        End If




    End Sub

    Private Sub BindingNavigator1_RefreshItems(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigator1.RefreshItems

    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        ' Frm3Pacientes.Hide()
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet4.Citas' Puede moverla o quitarla según sea necesario.
        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)

        'MsgBox("Ejecuta el Button22")
        TextBox6.Text = Frm3Pacientes.ComboBox5.Text
        TextBox7.Text = Frm3Pacientes.TextBox17.Text
        TextBox8.Text = Frm3Pacientes.TextBox2.Text & " " & Frm3Pacientes.TextBox3.Text & " " & Frm3Pacientes.TextBox4.Text

        Dim _CajaTexto, _Paciente As String
        Dim _checacontrol As Control
        Dim x As Integer
        '  Dim _Rows_ci() As DataRow = Me.CServiciosDataSet4.Citas.Select("Atiende = " & Frm3Pacientes.TextBox31.Text)
        Dim _Busca As DataRow
        Dim _Cuantos As Integer = Me.CServiciosDataSet4.Citas.Rows.Count

        'MsgBox("Cuantos " & CStr(_Cuantos))

        _CajaTexto = ""
        _Paciente = ""
        ' Verifica las consultas para ese Personal Médico
        For i = 0 To Me.CServiciosDataSet4.Citas.Rows.Count - 1
            If CServiciosDataSet4.Citas.Rows(i).Item("Estatus") <> "CANCELADO" Then
                '    MsgBox("Entró a recorrer el archivo")
                If Trim(Me.CServiciosDataSet4.Citas.Rows(i).Item("Atiende")) = Trim(Frm3Pacientes.TextBox31.Text) Then
                    '     MsgBox(" Si encuentra personal Médico ")
                    If Me.CServiciosDataSet4.Citas.Rows(i).Item("FechaCita") = TextBox7.Text Then
                        x = 1

                        For Each _checacontrol In Me.Controls
                            If TypeOf _checacontrol Is Button Then
                                ' MsgBox("Hora cita" & " " & CServiciosDataSet4.Citas.Rows(i).Item("HoraCita"))
                                ' MsgBox("Nombre button " & _checacontrol.Text)
                                If UCase(Me.CServiciosDataSet4.Citas.Rows(i).Item("HoraCita")) = UCase(_checacontrol.Text) Then
                                    '    MsgBox("Pasa paso 1")
                                    _CajaTexto = "TBHora" + Trim(Mid(_checacontrol.Name, 7, 2))
                                    _checacontrol.Enabled = False
                                    _Paciente = Me.CServiciosDataSet4.Citas.Rows(i).Item("Paciente")
                                    Exit For
                                    '   MsgBox("Paso uno y medio")
                                End If
                                x = x + 1
                            End If
                        Next

                        '  MsgBox(_CajaTexto)
                        '  MsgBox("Paso 2")
                        For Each _checacontrol In Me.Controls

                            If TypeOf _checacontrol Is TextBox Then
                                If _checacontrol.Name = _CajaTexto Then
                                    _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                                    _checacontrol.Text = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")
                                End If
                            End If
                        Next

                    End If
                End If
            End If

        Next


        ' Pone color a las cajas de texto
        For Each _checacontrol In Me.Controls
            If TypeOf _checacontrol Is TextBox Then
                _checacontrol.Enabled = False

                If _checacontrol.Text = "" Then
                    _checacontrol.BackColor = Color.LightGray
                Else
                    _checacontrol.BackColor = Color.Gold
                End If
            End If
        Next
        TextBox24.BackColor = Color.White
        ComboBox1.Visible = False

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TBHora1.Text = TextBox8.Text
        TextBox24.Text = Button1.Text
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Frm3Pacientes.TextBox18.Text = Me.TextBox24.Text
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TBHora2.Text = TextBox8.Text
        TextBox24.Text = Button2.Text
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TBHora3.Text = TextBox8.Text
        TextBox24.Text = Button3.Text
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TBHora4.Text = TextBox8.Text
        TextBox24.Text = Button4.Text
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TBHora5.Text = TextBox8.Text
        TextBox24.Text = Button5.Text
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TBHora6.Text = TextBox8.Text
        TextBox24.Text = Button6.Text
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TBHora7.Text = TextBox8.Text
        TextBox24.Text = Button7.Text
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TBHora8.Text = TextBox8.Text
        TextBox24.Text = Button8.Text
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        TBHora9.Text = TextBox8.Text
        TextBox24.Text = Button9.Text
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        TBHora10.Text = TextBox8.Text
        TextBox24.Text = Button10.Text
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        TBHora11.Text = TextBox8.Text
        TextBox24.Text = Button11.Text
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        TBHora12.Text = TextBox8.Text
        TextBox24.Text = Button12.Text
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        TBHora13.Text = TextBox8.Text
        TextBox24.Text = Button13.Text
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        TBHora14.Text = TextBox8.Text
        TextBox24.Text = Button14.Text
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        TBHora15.Text = TextBox8.Text
        TextBox24.Text = Button15.Text
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        TBHora16.Text = TextBox8.Text
        TextBox24.Text = Button16.Text
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        TBHora17.Text = TextBox8.Text
        TextBox24.Text = Button17.Text
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        TBHora18.Text = TextBox8.Text
        TextBox24.Text = Button18.Text
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        TBHora19.Text = TextBox8.Text
        TextBox24.Text = Button19.Text
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        TBHora20.Text = TextBox8.Text
        TextBox24.Text = Button20.Text
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        ComboBox1.Visible = True
        Dim _Checacontrol As Control
        Dim _Hora As String

        For Each _Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = False
            End If
            If TypeOf _Checacontrol Is Button Then
                _Checacontrol.Enabled = False
            End If
        Next
        TextBox24.Enabled = True
        ComboBox1.Enabled = True

        ComboBox1.Items.Clear()
        For i = 0 To 23
            _Hora = CStr(i) & ":00"
            ComboBox1.Items.Add(_Hora)
            _Hora = CStr(i) & ":30"
            ComboBox1.Items.Add(_Hora)
        Next


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox24.Text = ComboBox1.Text + " Hrs"
    End Sub

    Private Sub TBHora7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TBHora7.TextChanged

    End Sub

    Private Sub TBHora1_GotFocus(sender As Object, e As System.EventArgs) Handles TBHora1.GotFocus

    End Sub

    Private Sub TBHora1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TBHora1.TextChanged

    End Sub
End Class