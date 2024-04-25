Public Class Frm7HojaPaciente
    Public i As Integer
    Public msg As String
    Public Style As MsgBoxStyle

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Frm7HojaPaciente_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CServiciosDataSet2.DetalleReceta' table. You can move, or remove it, as needed.
        Me.DetalleRecetaTableAdapter.Fill(Me.CServiciosDataSet2.DetalleReceta)
        'TODO: This line of code loads data into the 'CServiciosDataSet5.Materiales' table. You can move, or remove it, as needed.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        Dim _Valor, _Contacto, _Telefono, _Domicilio, _Consulta, _Fecha, _Hora, _NombreClinica, _Receta As String
        ' Dim _UltimaConsulta, _UltimaReceta As String
        Dim _C, _UC As Date

        TextBox1.Text = Frm3Pacientes.TextBox1.Text
        TextBox2.Text = Frm3Pacientes.TextBox2.Text + " " + Frm3Pacientes.TextBox3.Text + " " + Frm3Pacientes.TextBox4.Text
        TextBox3.Text = Frm3Pacientes.TextBox5.Text
        TextBox4.Text = Frm3Pacientes.ComboBox1.Text
        TextBox5.Text = Frm3Pacientes.TextBox11.Text
        TextBox6.Text = Frm3Pacientes.ComboBox2.Text
        TextBox7.Text = Frm3Pacientes.TextBox24.Text
        TextBox8.Text = Frm3Pacientes.TextBox7.Text
        TextBox9.Text = Frm3Pacientes.TextBox9.Text
        TextBox10.Text = Frm3Pacientes.TextBox10.Text
        TextBox11.Text = Frm3Pacientes.TextBox12.Text
        TextBox12.Text = Frm3Pacientes.TextBox13.Text
        TextBox13.Text = Frm3Pacientes.TextBox14.Text
        TextBox14.Text = Frm3Pacientes.TextBox16.Text
        TextBox15.Text = Frm3Pacientes.TextBox30.Text
        Label14.Text = Frm3Pacientes.Label32.Text

        If TextBox9.Text = "SIN FECHA" Then
            TextBox9.BackColor = Color.LightGreen
        End If
 

        ' Muestra los Contactos
        DataGridView1.Rows.Clear()
        For Me.i = 0 To Frm3Pacientes.DataGridView6.Rows.Count - 1
            If Frm3Pacientes.DataGridView6.Rows(i).Cells(0).Value <> "" Then
                _Contacto = Frm3Pacientes.DataGridView6.Rows(i).Cells(0).Value
                _Domicilio = Frm3Pacientes.DataGridView6.Rows(i).Cells(1).Value
                _Telefono = Frm3Pacientes.DataGridView6.Rows(i).Cells(2).Value

                _Valor = _Contacto + _Telefono + _Domicilio
                DataGridView1.Rows.Add(_Contacto, _Telefono, _Domicilio)
            End If
        Next

        ' Muestra las Consultas
        DataGridView2.Rows.Clear()
        _UC = CDate(Frm3Pacientes.DataGridView3.Rows(0).Cells(3).Value)
        For Me.i = 0 To Frm3Pacientes.DataGridView3.Rows.Count - 1
            If Frm3Pacientes.DataGridView3.Rows(i).Cells(0).Value <> "" Then
                _Consulta = Frm3Pacientes.DataGridView3.Rows(i).Cells(2).Value
                _Fecha = Frm3Pacientes.DataGridView3.Rows(i).Cells(3).Value
                _C = CDate(_Fecha)
                If _C > _UC Then
                    _UC = _C
                End If
                _Hora = Frm3Pacientes.DataGridView3.Rows(i).Cells(4).Value
                _Receta = Frm3Pacientes.DataGridView3.Rows(i).Cells(8).Value
                _NombreClinica = Frm3Pacientes.DataGridView3.Rows(i).Cells(6).Value

                DataGridView2.Rows.Add(_Consulta, _Fecha, _Hora, _Receta, _NombreClinica)
            End If
        Next
        RichTextBox1.Text = Frm3Pacientes.RichTextBox1.Text

        ' Presenta la última Consulta
        TextBox16.Text = _UC

        ' Presenta la última Receta
        _UC = CDate(Frm3Pacientes.DataGridView4.Rows(0).Cells(4).Value)
        For Me.i = 0 To Frm3Pacientes.DataGridView4.Rows.Count - 1
            If Frm3Pacientes.DataGridView4.Rows(i).Cells(0).Value <> "" Then
                _Receta = Frm3Pacientes.DataGridView4.Rows(i).Cells(1).Value
                _Fecha = Frm3Pacientes.DataGridView4.Rows(i).Cells(4).Value
                _C = CDate(_Fecha)
                If _C > _UC Then
                    _UC = _C
                End If
            End If
        Next
        TextBox17.Text = _UC

        If TextBox17.Text <> "" Then
            DataGridView3.Visible = True

            ' Presemta los Medicamentos incluídos en la Receta
            Dim _Seleccion As String = "NumReceta = " & "'" & _Receta & "'"
            Dim _Rows_dr() As DataRow = CServiciosDataSet2.DetalleReceta.Select(_Seleccion)
            Dim _Material, _Descripcion, _Referencia As String
            Dim _Cantidad As Integer
            Dim _Busca As DataRow

            DataGridView3.Rows.Clear()
            For Me.i = 0 To _Rows_dr.GetUpperBound(0)
                _Material = _Rows_dr(i).Item("Material")
                _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_Material)
                _Descripcion = _Busca.Item("Descripcion")
                _Cantidad = _Rows_dr(i).Item("Cantidad")
                _Referencia = _Rows_dr(i).Item("Referencia")

                DataGridView3.Rows.Add(_Descripcion, _Cantidad, _Referencia)
            Next




        End If

        ' Muestra la Foto
        If My.Computer.FileSystem.FileExists("z:\Fotos Paciente\" & TextBox1.Text & ".jpg") Then
            PictureBox1.Image = Image.FromFile("z:\Fotos Paciente\" & TextBox1.Text & ".jpg")
            With PictureBox1

                If .Image.Width < .Width And .Image.Height < .Height Then
                    .SizeMode = PictureBoxSizeMode.CenterImage
                ElseIf .Image.Width.ToString > .Width Or .Image.Height.ToString > .Height Then
                    .SizeMode = PictureBoxSizeMode.StretchImage
                End If
            End With
        Else
            msg = "No se ha proporcionado la foto de este Paciente"
            Style = MsgBoxStyle.Information
            MsgBox(msg, Style)
        End If

    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub
End Class