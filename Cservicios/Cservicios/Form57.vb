Public Class Frm57NuevosPacientes
    Public I As Integer
    Public Msg As String
    Public _Consecutivo As Integer
    Public _XConsecutivo As String
    Public Style As MsgBoxStyle
    Public _Busca As DataRow
    Public _Registro As DataRow
    Public _Checacontrol As Control


    Private Sub Frm57NuevosPacientes_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Diagnosticos' Puede moverla o quitarla según sea necesario.
        Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet2.Diagnosticos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)


        ComboBox4.Items.Add("MASCULINO")
        ComboBox4.Items.Add("FEMENINO")

        For Me.I = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet.Centros.Rows(I).Item("Descripcion"))
        Next

        ' Obtiene el siguiente número de Paciente
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("PAC")
        _Consecutivo = _Busca.Item("Consecutivo")

        _XConsecutivo = "0000" & Trim(CStr(_Consecutivo))
        Dim _Mide As Integer = Len(_XConsecutivo)
        _XConsecutivo = "PA" & Mid(_XConsecutivo, (_Mide - 3), 4)

        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False

        For Me.I = 0 To CServiciosDataSet2.Diagnosticos.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet2.Diagnosticos.Rows(I).Item("Descripcion"))
        Next



    End Sub

    Private Sub ComboBox4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If ComboBox4.Text = "M" Or ComboBox4.Text = "m" Then
                ComboBox4.Text = "MASCULINO"
                TextBox24.Text = "M"
            Else
                If ComboBox4.Text = "M" Or ComboBox4.Text = "m" Then
                    ComboBox4.Text = "MASCULINO"
                    TextBox24.Text = "M"
                Else
                    Msg = "Debe teclear una F o una M en esta casilla"
                    Style = MsgBoxStyle.Information
                    MsgBox(Msg, Style)
                    Exit Sub
                End If
             

        End If
            TextBox2.Focus()
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.Text = "MASCULINO" Then
            TextBox24.Text = "M"
        Else
            TextBox24.Text = "F"
        End If
    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.GotFocus
        Dim _Centro As String = LoginForm1.TextBox1.Text
        TextBox5.Text = _Centro
        _Busca = Me.CServiciosDataSet.Centros.FindByCentro(_Centro)
        ComboBox1.Text = _Busca.Item("Descripcion")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim _Indice As Integer = ComboBox1.SelectedIndex
        TextBox5.Text = CServiciosDataSet.Centros.Rows(_Indice).Item("Centro")
    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click


        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está Seguro de registrar a este Paciente?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Registro de Pacientes"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Verifica que no haya casillas en blanco

            For Each Me._Checacontrol In Me.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    If _Checacontrol.Text = "" Then
                        Msg = "Hay casillas en blanco. Debe llenar todos los datos"
                        Style = MsgBoxStyle.Information
                        MsgBox(Msg, Style)
                        Exit Sub
                    End If
                End If
            Next

            ' Verifica que registre una fecha de Nacimiento válida
            If Not IsDate(TextBox30.Text) Then
                Msg = "Debe registrar una fecha de Nacimiento válida"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            _Registro = CServiciosDataSet2.Pacientes.NewRow
            _Registro.Item("Paciente") = TextBox1.Text
            _Registro.Item("Paterno") = TextBox2.Text
            _Registro.Item("Materno") = TextBox3.Text
            _Registro.Item("Nombres") = TextBox4.Text
            _Registro.Item("Centro") = TextBox5.Text
            _Registro.Item("Referencia") = TextBox6.Text
            _Registro.Item("FechaIngreso") = TextBox7.Text
            _Registro.Item("FechaAlta") = TextBox9.Text
            _Registro.Item("Diagnostico") = TextBox11.Text
            _Registro.Item("Calle") = ""
            _Registro.Item("Colonia") = ""
            _Registro.Item("Ciudad") = ""
            _Registro.Item("Estado") = ""
            _Registro.Item("CP") = ""
            _Registro.Item("Telefono") = ""
            _Registro.Item("FechaRegistro") = TextBox30.Text
            _Registro.Item("HoraRegistro") = TimeString
            _Registro.Item("Sexo") = TextBox24.Text
            _Registro.Item("Usuario") = LoginForm1.TextBox6.Text
            _Registro.Item("Comentarios") = ""

            CServiciosDataSet2.Pacientes.Rows.Add(_Registro)

            Me.Validate()
            PacientesBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet2)
            Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)

            ' Actualiza el Consecutivo de los Pacientes
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("PAC")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1


            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

            Msg = "Paciente registrado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Me.Close()


        Else
            ' Perform some other action.
        End If


    End Sub

    Private Sub TextBox9_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox9.GotFocus
        TextBox9.Text = "SIN FECHA"
    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.Text = Today.Date
    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim _Index As Integer = ComboBox2.SelectedIndex
        TextBox11.Text = CServiciosDataSet2.Diagnosticos.Rows(_Index).Item("Diagnostico")
    End Sub

    Private Sub TextBox30_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox30.GotFocus
        TextBox30.Text = Date.Today
    End Sub

    Private Sub TextBox30_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox30.TextChanged

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub
End Class