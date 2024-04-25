Public Class Frm33Clientes
    Public i As Integer
    Public _Checacontrol As Control
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Cliente, _NombreCliente, _Calle, _Colonia, _Ciudad, _Estado, _CP, _Contacto As String
    Public _Telefono, _TelefonoContacto, _Referencia As String
    Public _Busca As DataRow

    Private Sub Frm33Clientes_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.Clientes' Puede moverla o quitarla según sea necesario.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)


        Timer1.Enabled = False

        DataGridView1.Rows.Clear()
        For Me.i = 0 To CServiciosDataSet30.Clientes.Rows.Count - 1
            _Cliente = CServiciosDataSet30.Clientes.Rows(i).Item("Cliente")
            _NombreCliente = CServiciosDataSet30.Clientes.Rows(i).Item("NombreCliente")

            DataGridView1.Rows.Add(_Cliente, _NombreCliente)
        Next

        Panel1.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next

        ToolStripButton4.Enabled = True
        ToolStripButton2.Enabled = False
        TextBox13.Text = "ALTA"
        TextBox1.Focus()
        Timer1.Enabled = True
        Timer1.Interval = 100
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus

        If CheckBox1.Checked = True Then
            ' Obtiene el consecutivo para los Clientes
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("CLT")
            Dim _Consecutivo As Integer = _Busca.Item("Consecutivo")
            Dim _XConsecutivo As String = "CL" & Trim(CStr(_Consecutivo))
            TextBox1.Text = _XConsecutivo
            TextBox1.Enabled = False
            TextBox2.Focus()
        End If


    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

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

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White

    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox6.Text = "CHIHUAHUA"
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

    Private Sub TextBox7_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.Text = "CHIHUAHUA"
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.BackColor = Color.White
    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox11.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox12.Focus()
        End If
    End Sub

    Private Sub TextBox11_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox11.LostFocus
        TextBox11.BackColor = Color.White
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox12.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox12_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox12.LostFocus
        TextBox12.Text = LCase(TextBox12.Text)
        TextBox12.BackColor = Color.White
    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        ' Proceso de Alta o Modificación
        Dim _Find As Boolean
        Dim _Rows_Cl() As DataRow

        _Find = False
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Text = "" Then
                    _Find = True
                End If
            End If
        Next

        If _Find Then
            Msg = "Hay casillas en blanco, debe proporcionar todos los datos"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If TextBox13.Text = "ALTA" Then
            _Find = False
            _Cliente = TextBox1.Text
            _Rows_Cl = CServiciosDataSet30.Clientes.Select("Cliente = " & "'" & _Cliente & "'")

            For Me.i = 0 To _Rows_Cl.GetUpperBound(0)
                _Find = True
            Next

            If _Find Then
                Msg = "Ese Cliente ya está registrado"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If

            _Busca = CServiciosDataSet30.Clientes.NewRow
            _Busca.Item("Cliente") = _Cliente

        End If



        Timer1.Enabled = False

        If TextBox13.Text = "CAMBIO" Then
            _Cliente = TextBox1.Text
            _Busca = CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
        End If

        _Busca.Item("NombreCliente") = TextBox2.Text
        _Busca.Item("Contacto") = TextBox3.Text
        _Busca.Item("Calle") = TextBox4.Text
        _Busca.Item("Colonia") = TextBox5.Text
        _Busca.Item("Ciudad") = TextBox6.Text
        _Busca.Item("Estado") = TextBox7.Text
        _Busca.Item("CP") = TextBox8.Text
        _Busca.Item("Telefono") = TextBox9.Text
        _Busca.Item("TelefonoContacto") = TextBox10.Text
        _Busca.Item("Referencia") = TextBox11.Text
        _Busca.Item("Email") = TextBox12.Text
        _Busca.Item("FechaRegistro") = Today
        _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)

        If TextBox13.Text = "ALTA" Then
            CServiciosDataSet30.Clientes.Rows.Add(_Busca)
        End If

        Me.Validate()
        ClientesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet30)
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)


        ' Actualiza el Consecutivo
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("CLT")
        _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet3)
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        Msg = "Registro guardado correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)
        Button1_Click(sender, e)

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _Cliente = _Renglon.Cells(0).Value
        TextBox1.Text = _Cliente
        TextBox1.Enabled = False
        _Busca = CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
        TextBox2.Text = _Busca.Item("NombreCliente")
        TextBox3.Text = _Busca.Item("Contacto")
        TextBox4.Text = _Busca.Item("Calle")
        TextBox5.Text = _Busca.Item("Colonia")
        TextBox6.Text = _Busca.Item("Ciudad")
        TextBox7.Text = _Busca.Item("Estado")
        TextBox8.Text = _Busca.Item("CP")
        TextBox9.Text = _Busca.Item("Telefono")
        TextBox10.Text = _Busca.Item("TelefonoContacto")
        TextBox11.Text = _Busca.Item("Referencia")
        TextBox12.Text = _Busca.Item("Email")
        ToolStripButton3.Enabled = True
        Panel1.Enabled = True
        ToolStripButton2.Enabled = False
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
     

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        TextBox13.Text = "CAMBIO"
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True
        Timer1.Enabled = True
        Timer1.Interval = 100
    End Sub

    Private Sub ToolStripButton20_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton20.Click
        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton13_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton13.Click
        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar un Cliente para generar una nueva Factura"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        Else
            Frm35AltaFactura.Show()
        End If
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

    Private Sub TextBox9_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox9_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox9.LostFocus
        TextBox9.BackColor = Color.White
    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

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
End Class