Public Class Frm4CompletaConsulta
    Public _Busca As DataRow
    Private Sub Frm4CompletaConsulta_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet15.Consultas' Puede moverla o quitarla según sea necesario.
        Me.ConsultasTableAdapter1.Fill(Me.CServiciosDataSet15.Consultas)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Pacientes' table. You can move, or remove it, as needed.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Consultas' table. You can move, or remove it, as needed.
        Me.ConsultasTableAdapter.Fill(Me.CServiciosDataSet2.Consultas)

        Dim _Nombre As String = Frm3Pacientes.TextBox2.Text & " " & Frm3Pacientes.TextBox3.Text & " " & Frm3Pacientes.TextBox4.Text
        ToolStripTextBox1.Text = _Nombre
        ToolStripTextBox2.Text = Frm3Pacientes.ComboBox1.Text
        ToolStripTextBox3.Text = Frm3Pacientes.DataGridView3.CurrentRow.Cells(2).Value
        TextBox1.Text = Frm3Pacientes.DataGridView3.CurrentRow.Cells(3).Value
        TextBox2.Text = Frm3Pacientes.DataGridView3.CurrentRow.Cells(4).Value
        TextBox3.Text = Frm3Pacientes.DataGridView3.CurrentRow.Cells(7).Value
        TextBox4.Text = Frm3Pacientes.TextBox10.Text
        TextBox5.Text = Frm3Pacientes.TextBox12.Text
        TextBox6.Text = Frm3Pacientes.TextBox13.Text
        TextBox7.Text = Frm3Pacientes.TextBox16.Text

        _Busca = CServiciosDataSet15.Consultas.FindByConsulta(ToolStripTextBox3.Text)
        If DBNull.Value.Equals(_Busca.Item("PresionArterial")) Then
            TextBox8.Text = ""
        Else
            TextBox8.Text = _Busca.Item("PresionArterial")
        End If
        If DBNull.Value.Equals(_Busca.Item("Peso")) Then
            TextBox9.Text = ""
        Else
            TextBox9.Text = _Busca.Item("Peso")
        End If
        If DBNull.Value.Equals(_Busca.Item("Estatura")) Then
            TextBox10.Text = ""
        Else
            TextBox10.Text = _Busca.Item("Estatura")
        End If
        If DBNull.Value.Equals(_Busca.Item("Importeconsulta")) Then
            TextBox11.Text = 0
        Else
            TextBox11.Text = _Busca.Item("ImporteConsulta")
        End If

        If DBNull.Value.Equals(_Busca.Item("ImporteDescuento")) Then
            TextBox12.Text = 0
        Else
            TextBox12.Text = _Busca.Item("ImporteDescuento")
        End If

        TextBox13.Text = CInt(TextBox11.Text) - CInt(TextBox12.Text)


        If DBNull.Value.Equals(_Busca.Item("Comentarios")) Then
            RichTextBox1.Text = ""
        Else
            RichTextBox1.Text = _Busca.Item("Comentarios")
        End If
        ToolStripButton3.Enabled = False
        Panel2.Enabled = False
        Panel1.Enabled = False
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True
        RichTextBox1.Enabled = True
        TextBox1.Focus()
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guarda las modificaciones a la Consulta
        _Busca = CServiciosDataSet15.Consultas.FindByConsulta(ToolStripTextBox3.Text)
        _Busca.Item("FechaConsulta") = TextBox1.Text
        _Busca.Item("Horaconsulta") = TextBox2.Text
        _Busca.Item("Atiende") = TextBox3.Text
        _Busca.Item("Comentarios") = UCase(RichTextBox1.Text)
        _Busca.Item("ImporteConsulta") = CInt(TextBox11.Text)
        _Busca.Item("ImporteDescuento") = CInt(TextBox12.Text)
        _Busca.Item("PresionArterial") = TextBox8.Text
        _Busca.Item("Estatura") = TextBox10.Text
        _Busca.Item("Peso") = TextBox9.Text
        TextBox13.Text = CStr(CInt(TextBox11.Text) - CInt(TextBox12.Text))


        Me.Validate()
        ConsultasBindingSource1.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet15)
        Me.ConsultasTableAdapter1.Fill(Me.CServiciosDataSet15.Consultas)

        _Busca = CServiciosDataSet2.Pacientes.FindByPaciente(Frm3Pacientes.DataGridView3.CurrentRow.Cells(0).Value)
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + vbCrLf
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + vbCrLf
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + "CONSULTA " + ToolStripTextBox3.Text + " " + CStr(Today) + " " + TimeString
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + vbCrLf
        _Busca.Item("Comentarios") = _Busca.Item("Comentarios") + RichTextBox1.Text

        Me.Validate()
        PacientesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet2)
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)

        MsgBox("Registro guardado Correctamente")
        Me.Close()

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click

    End Sub
End Class