Public Class Frm22ConsultadePacientes
    Public i As Integer
    Public _NombrePaciente, _Paciente As String

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Form22_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Diagnosticos' Puede moverla o quitarla según sea necesario.
        Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet2.Diagnosticos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet8.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)

        For Me.i = 0 To Me.CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres")
            ToolStripComboBox1.Items.Add(_NombrePaciente)
        Next

        ToolStripComboBox1.Sorted = True
    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet2.Pacientes.Rows.Count - 1
            If Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paterno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Materno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Nombres") = ToolStripComboBox1.Text Then
                ToolStripTextBox1.Text = Me.CServiciosDataSet2.Pacientes.Rows(i).Item("Paciente")
            End If
        Next
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

        Dim msg, _Paciente, _Centro, _Diagnostico As String
        Dim Style As MsgBoxStyle
        Dim _Busca, _BuscaCentro, _BuscaDiagnostico As DataRow
        Dim _Rutaimagen As String

        Style = MsgBoxStyle.Information
        If ToolStripTextBox1.Text = "" Then
            msg = "Debe seleccionar un Paciente para esta consulta"
            MsgBox(msg, Style)
            Exit Sub
        End If

        _Paciente = ToolStripTextBox1.Text
        _Busca = Me.CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)

        TextBox1.Text = _Busca.Item("Paciente")

        _Rutaimagen = "Z:\Fotos Paciente\" & TextBox1.Text & ".jpg"
        TextBox2.Text = _Busca.Item("Paterno")
        TextBox3.Text = _Busca.Item("Materno")
        TextBox4.Text = _Busca.Item("Nombres")
        If DBNull.Value.Equals(_Busca.Item("Calle")) Then
            TextBox5.Text = ""
        Else
            TextBox5.Text = _Busca.Item("Calle")
        End If

        If DBNull.Value.Equals(_Busca.Item("Colonia")) Then
            TextBox6.Text = ""
        Else
            TextBox6.Text = _Busca.Item("Colonia")
        End If
        If DBNull.Value.Equals(_Busca.Item("Ciudad")) Then
            TextBox7.Text = ""
        Else
            TextBox7.Text = _Busca.Item("Ciudad")
        End If
        If DBNull.Value.Equals(_Busca.Item("Estado")) Then
            TextBox8.Text = ""
        Else
            TextBox8.Text = _Busca.Item("Estado")
        End If
        If DBNull.Value.Equals(_Busca.Item("cp")) Then
            TextBox9.Text = ""
        Else
            TextBox9.Text = _Busca.Item("CP")
        End If
        If DBNull.Value.Equals(_Busca.Item("Telefono")) Then
            TextBox10.Text = ""
        Else
            TextBox10.Text = _Busca.Item("Telefono")
        End If

        TextBox11.Text = _Busca.Item("FechaIngreso")
        TextBox12.Text = _Busca.Item("FechaAlta")
        _Centro = _Busca.Item("Centro")

        _BuscaCentro = Me.CServiciosDataSet8.Centros.FindByCentro(_Centro)
        TextBox13.Text = _BuscaCentro.Item("Descripcion")

        _Diagnostico = _Busca.Item("Diagnostico")
        _BuscaDiagnostico = Me.CServiciosDataSet2.Diagnosticos.FindByDiagnostico(_Diagnostico)
        TextBox14.Text = _BuscaDiagnostico.Item("Descripcion")
        '     Textbox15.text = _Busca.item("Escolaridad") 
        '     TextBox16.Text = _Busca.Item("ImporteDescuento")
        '     TextBox17.Text = _Busca.Item("PorcentajeDescuento")
        If DBNull.Value.Equals(_Busca.Item("FechaRegistro")) Then
            TextBox18.Text = ""
        Else
            TextBox18.Text = _Busca.Item("FechaRegistro")
        End If

        If DBNull.Value.Equals(_Busca.Item("Comentarios")) Then
            RichTextBox1.Text = ""
        Else
            RichTextBox1.Text = _Busca.Item("Comentarios")
        End If

        ' Busca la foto

        If My.Computer.FileSystem.FileExists(_Rutaimagen) Then
            PictureBox1.Image = System.Drawing.Image.FromFile(_Rutaimagen)

            With PictureBox1

                If .Image.Width < .Width And .Image.Height < .Height Then
                    .SizeMode = PictureBoxSizeMode.CenterImage
                ElseIf .Image.Width.ToString > .Width Or .Image.Height.ToString > .Height Then
                    .SizeMode = PictureBoxSizeMode.StretchImage
                End If
            End With
        Else
            msg = "No está registrada la Foto de este Paciente " & _Rutaimagen
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
            PictureBox1.Image = Nothing
        End If


    End Sub
End Class