Public Class Frm21ConsultasPersonalMedico
    Public i As Integer
    Public _Checacontrol As Control

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Frm21ConsultasPersonalMedico_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet13.Especialidades' Puede moverla o quitarla según sea necesario.
        Me.EspecialidadesTableAdapter.Fill(Me.CServiciosDataSet13.Especialidades)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet8.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet14.Medicos' Puede moverla o quitarla según sea necesario.
        Me.MedicosTableAdapter.Fill(Me.CServiciosDataSet14.Medicos)


        Dim _NombreMedico As String
        ToolStripComboBox1.Items.Clear()
        For Me.i = 0 To Me.CServiciosDataSet14.Medicos.Rows.Count - 1
            _NombreMedico = Me.CServiciosDataSet14.Medicos.Rows(i).Item("Nombre")
            ToolStripComboBox1.Items.Add(_NombreMedico)
        Next


    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        For Me.i = 0 To Me.CServiciosDataSet14.Medicos.Rows.Count - 1
            If Me.CServiciosDataSet14.Medicos.Rows(i).Item("Nombre") = ToolStripComboBox1.Text Then
                ToolStripTextBox1.Text = Me.CServiciosDataSet14.Medicos.Rows(i).Item("Medico")
            End If
        Next
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim msg, _Centro, _Especialidad As String
        Dim style As MsgBoxStyle
        Dim _Busca, _BuscaCentro, _BuscaEspecialidad As DataRow
        Dim _Medico As String = ToolStripTextBox1.Text
        Dim _RutaImagen As String

        If ToolStripTextBox1.Text = "" Then
            msg = "Debe seleccionar un Personal Médico para esta Consulta"
            MsgBox(msg, style)
            Exit Sub
        End If

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""



        _Busca = Me.CServiciosDataSet14.Medicos.FindByMedico(_Medico)
        TextBox1.Text = _Busca.Item("Medico")

        _RutaImagen = "Z:\Fotos Medicos\" & Trim(TextBox1.Text) & ".jpg"
        TextBox2.Text = _Busca.Item("Nombre")
        _Centro = _Busca.Item("Centro")
        _BuscaCentro = Me.CServiciosDataSet8.Centros.FindByCentro(_Centro)
        TextBox3.Text = _BuscaCentro.Item("Descripcion")
        _Especialidad = _Busca.Item("Especialidad")

        For Me.i = 0 To CServiciosDataSet13.Especialidades.Rows.Count - 1
            If CServiciosDataSet13.Especialidades.Rows(i).Item("Especialidad") = _Especialidad Then
                TextBox4.Text = CServiciosDataSet13.Especialidades.Rows(i).Item("Descripcion")
            Else
                msg = "La Especialidad de este personal médico no está registrada. Verifique"
                style = MsgBoxStyle.Information
                '  MsgBox(msg, style)
            End If
        Next



        ' _BuscaEspecialidad = Me.CServiciosDataSet13.Especialidades.FindByEspecialidad(_Especialidad)
        ' TextBox4.Text = _BuscaEspecialidad.Item("Descripcion")
        TextBox5.Text = _Busca.Item("Referencia")
        TextBox6.Text = _Busca.Item("Prefijo")
        TextBox7.Text = _Busca.Item("TelefonoCasa")
        TextBox8.Text = _Busca.Item("TelefonoCelular")
        TextBox9.Text = _Busca.Item("CorreoElectronico")
        TextBox10.Text = _Busca.Item("ConsultorioExterno")
        TextBox11.Text = _Busca.Item("HorarioInicial")
        TextBox12.Text = _Busca.Item("HorarioFinal")
        RichTextBox1.Text = _Busca.Item("Comentarios")

        ' Busca la foto

        If My.Computer.FileSystem.FileExists(_RutaImagen) Then
            PictureBox1.Image = System.Drawing.Image.FromFile(_RutaImagen)

            With PictureBox1

                If .Image.Width < .Width And .Image.Height < .Height Then
                    .SizeMode = PictureBoxSizeMode.CenterImage
                ElseIf .Image.Width.ToString > .Width Or .Image.Height.ToString > .Height Then
                    .SizeMode = PictureBoxSizeMode.StretchImage
                End If
            End With
        Else
            PictureBox1.Image = Nothing
        End If

        ' ☺☻♥♦♣♠•◘○◙♂♀♪♫☼►◄↕‼¶§▬↨↑↓→←∟↔▲▼ !"#$%&'()*+,-./0123456789:;<=>?@ABC
        '92 = \

    End Sub
End Class