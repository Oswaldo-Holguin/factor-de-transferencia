Public Class Frm1Centros
    Public _renglon As DataGridViewRow
    Public i As Integer
    Public checacontrol As Control
    Public _Centro As String

    Private Sub Frm1Centros_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CServiciosDataSet8.Centros' table. You can move, or remove it, as needed.
        Me.CentrosTableAdapter1.Fill(Me.CServiciosDataSet8.Centros)
        'TODO: This line of code loads data into the 'CServiciosDataSet8.Centros' table. You can move, or remove it, as needed.
        Me.CentrosTableAdapter1.Fill(Me.CServiciosDataSet8.Centros)

        Button1_Click(sender, e)
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: This line of code loads data into the 'CServiciosDataSet.Centros' table. You can move, or remove it, as needed.
        Me.CentrosTableAdapter1.Fill(Me.CServiciosDataSet8.Centros)

        ' Dim _Centro As String
        Dim _Descripcion As String
        Dim _Titular As String
        Dim _Referencia As String
        Dim _Comentarios As String
        Dim _FechaRegistro As String
        Dim _HoraRegistro As String
        Dim _Funcion As String
        Dim i As Integer
        Dim _usuario As String
        Dim checacontrol As Control
        Dim _ImporteConsulta As Integer

        DataGridView1.Rows.Clear()
        For i = 0 To CServiciosDataSet8.Centros.Rows.Count - 1
            _Centro = CServiciosDataSet8.Centros.Rows(i).Item("Centro")
            _Descripcion = CServiciosDataSet8.Centros.Rows(i).Item("Descripcion")
            _Titular = CServiciosDataSet8.Centros.Rows(i).Item("Titular")
            _Funcion = CServiciosDataSet8.Centros.Rows(i).Item("Funcion")
            _Referencia = CServiciosDataSet8.Centros.Rows(i).Item("Referencia")
            If DBNull.Value.Equals(CServiciosDataSet8.Centros.Rows(i).Item("Comentarios")) Then
                _Comentarios = ""
            Else
                _Comentarios = CServiciosDataSet8.Centros.Rows(i).Item("Comentarios")
            End If

            _FechaRegistro = CServiciosDataSet8.Centros.Rows(i).Item("FechaRegistro")
            _HoraRegistro = CServiciosDataSet8.Centros.Rows(i).Item("HoraRegistro")
            If DBNull.Value.Equals(CServiciosDataSet8.Centros.Rows(i).Item("usuario")) Then
                _usuario = ""
            Else
                _usuario = CServiciosDataSet8.Centros.Rows(i).Item("Usuario")
            End If
            If DBNull.Value.Equals(CServiciosDataSet8.Centros.Rows(i).Item("ImporteConsulta")) Then
                _ImporteConsulta = 0
            Else
                _ImporteConsulta = CServiciosDataSet8.Centros.Rows(i).Item("ImporteConsulta")
            End If

            DataGridView1.Rows.Add(_Centro, _Descripcion, _Titular, _Funcion, _Referencia, _Comentarios, _usuario, _FechaRegistro, _HoraRegistro, _ImporteConsulta)

        Next

        For Each checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If

        Next
        Panel1.Enabled = False
        RichTextBox1.Enabled = False


    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _renglon = DataGridView1.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.Gold

        TextBox1.Text = _renglon.Cells(0).Value
        TextBox2.Text = _renglon.Cells(1).Value
        TextBox3.Text = _renglon.Cells(3).Value
        TextBox4.Text = _renglon.Cells(2).Value
        TextBox5.Text = _renglon.Cells(4).Value
        TextBox6.Text = _renglon.Cells(7).Value
        TextBox7.Text = _renglon.Cells(8).Value
        TextBox9.Text = _renglon.Cells(9).Value
        RichTextBox1.Text = _renglon.Cells(5).Value

        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
    
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.Text = UCase(TextBox1.Text)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

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

    Private Sub TextBox4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.Text = UCase(TextBox4.Text)
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.Text = UCase(TextBox5.Text)
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

        ToolStripButton2.Enabled = False
        Dim checacontrol As Control
        Panel1.Enabled = True
        TextBox1.Focus()
        RichTextBox1.Enabled = True
        ToolStripButton3.Enabled = True

        For Each checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""
        TextBox8.Text = "ALTA"
    End Sub

    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _renglon = DataGridView1.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guarda el registro
        Dim checacontrol As Control
        Dim _encuentra As Boolean
        Dim Busca As DataRow

        _encuentra = False
        For Each checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                If checacontrol.Text = "" Then
                    _encuentra = True
                End If
            End If
        Next
        If _encuentra Then
            MsgBox("Hay casillas en blanco. Debe llenar todos los datos.")
            Exit Sub
        End If

        ' ***************************************************************************************************
        If TextBox8.Text = "ALTA" Then
            Busca = CServiciosDataSet8.Centros.NewRow()
            Busca.Item("Centro") = TextBox1.Text
        End If

        If TextBox8.Text = "CAMBIO" Then
            Busca = CServiciosDataSet8.Centros.FindByCentro(TextBox1.Text)
        End If

        Busca.Item("Descripcion") = TextBox2.Text
        Busca.Item("Funcion") = TextBox3.Text
        Busca.Item("Titular") = TextBox4.Text
        Busca.Item("Referencia") = TextBox5.Text
        Busca.Item("FechaRegistro") = TextBox6.Text
        Busca.Item("HoraRegistro") = TextBox7.Text
        Busca.Item("Comentarios") = UCase(RichTextBox1.Text)
        Busca.Item("ImporteConsulta") = CInt(TextBox9.Text)

        If TextBox8.Text = "ALTA" Then
            CServiciosDataSet8.Centros.Rows.Add(Busca)
        End If


        Me.Validate()
        CentrosBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet8)
        Me.CentrosTableAdapter1.Fill(Me.CServiciosDataSet8.Centros)

        MsgBox("Registro guardado Correctamente")
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        RichTextBox1.Text = ""
        TextBox1.Enabled = True
        Button1_Click(sender, e)
    End Sub

    Private Sub TextBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox6.Text = CStr(Today)
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.Text = TimeString
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Panel1.Enabled = True
        ToolStripButton2.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Focus()
        ToolStripButton3.Enabled = True
        TextBox8.Text = "CAMBIO"
        ToolStripButton4.Enabled = False
        RichTextBox1.Enabled = True

    End Sub
End Class
