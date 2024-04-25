Public Class Frm26TablaTipoEventos
    Public i As Integer
    Private Sub Form26_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim _Consecutivo As Integer


        ToolStripButton3.Enabled = True
        ToolStripButton2.Enabled = False
        Panel1.Enabled = True
        TextBox1.Focus()

        ' Obtiene el consecutivo de la tabla de tipos de Eventos

        For Me.i = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(i).Item("Documento") = "EVN" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(i).Item("Consecutivo")
            End If
        Next
        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "EV" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False


    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.Text = UCase(TextBox2.Text)
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.Text = UCase(TextBox3.Text)
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If TextBox2.Text = "" Then
            MsgBox("Debe teclear un dato en el campo de DESCRIPCION ")
            Exit Sub
        End If

        Dim Busca As DataRow
        Dim msg As String
        Dim style As MsgBoxStyle


        Busca = CServiciosDataSet18.Eventos.NewRow
        Busca.Item("Evento") = TextBox1.Text
        Busca.Item("Descripcion") = TextBox2.Text
        Busca.Item("Referencia") = TextBox3.Text
        Busca.Item("Comentarios") = RichTextBox1.Text
        Busca.Item("FechaRegistro") = Today
        Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)

        CServiciosDataSet18.Eventos.Rows.Add(Busca)

        Me.Validate()
        EventosBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet18)
        Me.EventosTableAdapter.Fill(Me.CServiciosDataSet18.Eventos)


        ' Actuaqliza el Consecutivo de los Eventos
        Busca = CServiciosDataSet3.Documentos.FindByDocumento("EVN")
        Busca.Item("Consecutivo") = Busca.Item("Consecutivo") + 1

        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet3)
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        msg = "Registro guardado correctamente"
        style = MsgBoxStyle.Information
        MsgBox(msg, style)

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        RichTextBox1.Text = ""
        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet18.Eventos' Puede moverla o quitarla según sea necesario.
        Me.EventosTableAdapter.Fill(Me.CServiciosDataSet18.Eventos)
        Dim _Evento, _Descripcion, _Referencia, _Comentarios As String

        DataGridView1.Rows.Clear()
        For Me.i = 0 To CServiciosDataSet18.Eventos.Rows.Count - 1
            _Evento = CServiciosDataSet18.Eventos.Rows(i).Item("Evento")
            _Descripcion = CServiciosDataSet18.Eventos.Rows(i).Item("Descripcion")
            _Referencia = CServiciosDataSet18.Eventos.Rows(i).Item("Referencia")
            _Comentarios = CServiciosDataSet18.Eventos.Rows(i).Item("Comentarios")
            DataGridView1.Rows.Add(_Evento, _Descripcion, _Referencia, _Comentarios)
        Next

        ToolStripButton3.Enabled = False
        Panel1.Enabled = False
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim _renglon As DataGridViewRow = DataGridView1.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _renglon.Cells(0).Value
        TextBox2.Text = _renglon.Cells(1).Value
        TextBox3.Text = _renglon.Cells(2).Value
        RichTextBox1.Text = _renglon.Cells(3).Value

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        Dim _renglon As DataGridViewRow = DataGridView1.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.White

    End Sub
End Class