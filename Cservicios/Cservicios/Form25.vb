Public Class Frm25NuevoElectro
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Checacontrol As Control

    Private Sub Form25_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True

        ' Busca el consecutivo para los Electros
        Dim Busca As DataRow
        Dim _Consecutivo As Integer
        Dim _Xconsecutivo As String
        Busca = CServiciosDataSet3.Documentos.FindByDocumento("EEE")
        _Consecutivo = Busca.Item("Consecutivo")
        _Xconsecutivo = "000" & Trim(CStr(_Consecutivo))

        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _Xconsecutivo = "EE" + Mid(_Xconsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()
        Timer1.Enabled = True
        Timer1.Interval = 100
    End Sub

    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        MonthCalendar1.Visible = False
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        TextBox2.Text = MonthCalendar1.SelectionRange.Start
        MonthCalendar1.Visible = False
        TextBox2.BackColor = Color.White
        TextBox3.Focus()
    End Sub

    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.Text = Frm3Pacientes.DataGridView3.CurrentRow.Cells(2).Value
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.GotFocus
        TextBox5.Text = ToolStripTextBox1.Text
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
        RichTextBox1.BackColor = Color.White
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim Busca As DataRow
        Dim msg As String
        Dim Style As MsgBoxStyle
        Dim _respuesta As MsgBoxResult
        msg = "Está seguro de Guardar estos Datos?"
        Style = MsgBoxStyle.OkCancel

        _respuesta = MsgBox(msg, Style)

        Dim _Fecha As String = TextBox2.Text
        If Not IsDate(_Fecha) Then
            msg = "Debe teclear una fecha del Electro correcta"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            Exit Sub
        End If



        If _respuesta = MsgBoxResult.Ok Then
            ' Guarda la información del Electro
            Busca = CServiciosDataSet17.Electros.NewRow
            Busca.Item("Electro") = TextBox1.Text
            Busca.Item("Paciente") = Frm3Pacientes.TextBox1.Text
            Busca.Item("FechaElectro") = TextBox2.Text
            Busca.Item("HoraElectro") = TextBox3.Text
            Busca.Item("Consulta") = TextBox7.Text
            Busca.Item("Referencia") = UCase(TextBox6.Text)
            Busca.Item("Comentarios") = UCase(RichTextBox1.Text)
            Busca.Item("FechaRegistro") = Today
            Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
            Busca.Item("Centro") = Frm3Pacientes.TextBox5.Text
            Busca.Item("IdElecro") = TextBox5.Text
            If RadioButton1.Checked Then
                Busca.Item("Exitoso") = True
            Else
                Busca.Item("Exitoso") = False
            End If
            Busca.Item("Montaje") = TextBox9.Text
            Busca.Item("Sensibilidad") = TextBox10.Text
            Busca.Item("Frecuencia") = TextBox11.Text
            Busca.Item("Consecutivo") = TextBox4.Text

            Me.CServiciosDataSet17.Electros.Rows.Add(Busca)
            Me.Validate()
            ElectrosBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet17)
            Me.ElectrosTableAdapter.Fill(CServiciosDataSet17.Electros)


            ' Actualiza el consecutivo
            Busca = CServiciosDataSet3.Documentos.FindByDocumento("EEE")
            Busca.Item("Consecutivo") = Busca.Item("Consecutivo") + 1

            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet3)

            Style = MsgBoxStyle.Information
            msg = "Registro guardado correctamente"
            MsgBox(msg, Style)

            Button1_Click(sender, e)

        Else
            Exit Sub
        End If

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox9_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox9.LostFocus
        TextBox9.BackColor = Color.White
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet17.Electros' Puede moverla o quitarla según sea necesario.
        Me.ElectrosTableAdapter.Fill(Me.CServiciosDataSet17.Electros)

        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        ToolStripTextBox1.Text = Trim(Frm3Pacientes.TextBox2.Text) & " " & Trim(Frm3Pacientes.TextBox3.Text) & " " & Trim(Frm3Pacientes.TextBox4.Text)
        ToolStripButton3.Enabled = False
        Panel1.Enabled = False

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.BackColor = Color.White
            End If
        Next
        RichTextBox1.Text = ""
        RichTextBox1.BackColor = Color.White
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Focused Then
                    _Checacontrol.BackColor = Color.Gold
                End If
            End If
        Next
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = "15:00"
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

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

    Private Sub TextBox10_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox10.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox10_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox10.LostFocus
        TextBox10.BackColor = Color.White
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox11.KeyPress
        If Asc(e.KeyChar) = 13 Then
            RichTextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox11_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox11.LostFocus
        TextBox11.BackColor = Color.White
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub
End Class