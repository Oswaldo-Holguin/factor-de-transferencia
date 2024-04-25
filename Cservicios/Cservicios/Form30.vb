Public Class Frm30PaquetesLeucocitarios
    Public i As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Folio, _NombreClinica, _Calle, _Colonia, _Ciudad, _Estado, _CP, _Donante, _Grupo As String
    Public _FechaExtraccion, _FechaCaducidad As Date
    Public _HoraExtraccion, _Volumen, _SolucionAnticoagulante, _SolucionConservadora As String
    Public _Comentarios, _Referencia, _Documento As String
    Public _Busca As DataRow
    Public _Checacontrol As Control
    Public _Renglon As DataGridViewRow




    Private Sub Frm30PaquetesLeucocitarios_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet24.PaquetesLeucocitarios' Puede moverla o quitarla según sea necesario.
        Me.PaquetesLeucocitariosTableAdapter.Fill(Me.CServiciosDataSet24.PaquetesLeucocitarios)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet24.DetallePaquetes' Puede moverla o quitarla según sea necesario.
        Me.DetallePaquetesTableAdapter.Fill(Me.CServiciosDataSet24.DetallePaquetes)

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next

        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        ToolStripButton2.Enabled = True
        ToolStripButton5.Enabled = True
        ToolStripButton20.Enabled = True


        ' Presenta el Histórico

        Dim _FechaRegistro As Date

        DataGridView1.Rows.Clear()

        For Me.i = 0 To CServiciosDataSet24.PaquetesLeucocitarios.Rows.Count - 1
            _Folio = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("Folio")
            _FechaExtraccion = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("FechaExtraccion")
            _HoraExtraccion = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("HoraExtraccion")
            _Donante = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("Donante")
            _Grupo = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("Grupo")
            _FechaRegistro = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("FechaRegistro")
            _NombreClinica = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("NombreClinica")

            DataGridView1.Rows.Add(_Folio, _FechaExtraccion, _HoraExtraccion, _Donante, _Grupo, _FechaRegistro, _NombreClinica)
        Next

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click


        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        Panel1.Enabled = True
        ToolStripButton2.Enabled = False
        ToolStripButton4.Enabled = True
        TextBox1.Focus()
        TextBox17.Text = "ALTA"

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        Dim _Find As Boolean
        Dim _Rows_Pl() As DataRow

        ' Alta de Paquetes
        If _Find Then
            Msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If TextBox17.Text = "ALTA" Then
            _Find = False
            For Each Me._Checacontrol In Panel1.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    If _Checacontrol.Text = "" Then
                        _Find = True
                    End If
                End If
            Next

            ' Verifica que no exita el folio
            _Folio = TextBox1.Text

            _Find = False
            _Rows_Pl = CServiciosDataSet24.PaquetesLeucocitarios.Select("Folio = " & "'" & _Folio & "'")

            For Me.i = 0 To _Rows_Pl.GetUpperBound(0)
                _Find = True
            Next

            If _Find Then
                Msg = "Ese Folio ya está registrado. Verifique"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If
        End If

        If TextBox17.Text = "ALTA" Then
            _Busca = CServiciosDataSet24.PaquetesLeucocitarios.NewRow
            _Busca.Item("Folio") = TextBox1.Text
        End If
        If TextBox17.Text = "CAMBIO" Then
            _Busca = CServiciosDataSet24.PaquetesLeucocitarios.FindByFolio(_Folio)
        End If

        _Busca.Item("NombreClinica") = TextBox2.Text
        _Busca.Item("Calle") = TextBox3.Text
        _Busca.Item("Colonia") = TextBox5.Text
        _Busca.Item("Ciudad") = TextBox7.Text
        _Busca.Item("CodigoPostal") = Mid(TextBox8.Text, 1, 5)
        _Busca.Item("Donante") = TextBox10.Text
        _Busca.Item("Grupo") = TextBox4.Text
        _Busca.Item("FechaExtraccion") = TextBox12.Text
        _Busca.Item("HoraExtraccion") = TextBox9.Text
        _Busca.Item("FechaCaducidad") = TextBox16.Text
        _Busca.Item("Volumen") = TextBox11.Text
        _Busca.Item("SolucionAnticoagulante") = TextBox13.Text
        _Busca.Item("SolucionConservadora") = TextBox14.Text
        _Busca.Item("Referencia") = TextBox15.Text
        _Busca.Item("FechaRegistro") = Today
        _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)


        If TextBox17.Text = "ALTA" Then
            CServiciosDataSet24.PaquetesLeucocitarios.Rows.Add(_Busca)
        End If

        Me.Validate()
        PaquetesLeucocitariosBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet24)
        Me.PaquetesLeucocitariosTableAdapter.Fill(Me.CServiciosDataSet24.PaquetesLeucocitarios)

        Msg = "Registro guardado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        _NombreClinica = TextBox2.Text
        _Calle = TextBox3.Text
        _Colonia = TextBox5.Text
        _Ciudad = TextBox7.Text
        _CP = Mid(TextBox8.Text, 1, 5)
        _Volumen = TextBox11.Text
        _SolucionAnticoagulante = TextBox13.Text
        _SolucionConservadora = TextBox14.Text
        _Volumen = TextBox11.Text
        _Referencia = TextBox15.Text
        _HoraExtraccion = TextBox9.Text


        Button1_Click(sender, e)

    End Sub

    Private Sub TextBox16_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox16.GotFocus
        _FechaCaducidad = CDate(TextBox12.Text)
        _FechaCaducidad = _FechaCaducidad.AddDays(30)
        TextBox16.Text = _FechaCaducidad

    End Sub

    Private Sub TextBox16_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox7_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.Text = "CHIHUAHUA"
    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.Text = UCase(TextBox7.Text)

    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox2.Text = _NombreClinica

    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.Text = UCase(TextBox2.Text)


    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = _Calle
    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.Text = UCase(TextBox3.Text)
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox5_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox5.GotFocus
        TextBox5.Text = _Colonia

    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.Text = UCase(TextBox5.Text)

    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox6.Text = _Ciudad
    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox6.LostFocus
        TextBox6.Text = UCase(TextBox6.Text)

    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox10_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox10.LostFocus
        TextBox10.Text = UCase(TextBox10.Text)
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox4_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.Text = UCase(TextBox4.Text)

    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox13.GotFocus
        TextBox13.Text = _SolucionAnticoagulante
    End Sub

    Private Sub TextBox13_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox13.LostFocus
        TextBox13.Text = UCase(TextBox13.Text)

    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox14_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox14.GotFocus
        TextBox14.Text = _SolucionConservadora
    End Sub

    Private Sub TextBox14_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox14.LostFocus
        TextBox14.Text = UCase(TextBox14.Text)
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox15_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox15.GotFocus
        TextBox15.Text = _Referencia
    End Sub

    Private Sub TextBox15_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox15.LostFocus
        TextBox15.Text = UCase(TextBox15.Text)

    End Sub

    Private Sub TextBox15_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(0).Value
        _Folio = TextBox1.Text

        _Busca = CServiciosDataSet24.PaquetesLeucocitarios.FindByFolio(_Folio)
        TextBox2.Text = _Busca.Item("NombreClinica")
        TextBox3.Text = _Busca.Item("Calle")
        TextBox5.Text = _Busca.Item("Colonia")
        TextBox6.Text = _Busca.Item("Ciudad")
        TextBox7.Text = _Busca.Item("Estado")
        TextBox8.Text = _Busca.Item("CodigoPostal")
        TextBox10.Text = _Busca.Item("Donante")
        TextBox4.Text = _Busca.Item("Grupo")
        TextBox12.Text = _Busca.Item("FechaExtraccion")
        TextBox9.Text = _Busca.Item("HoraExtraccion")
        TextBox11.Text = _Busca.Item("Volumen")
        TextBox13.Text = _Busca.Item("SolucionAnticoagulante")
        TextBox14.Text = _Busca.Item("SolucionConservadora")
        TextBox15.Text = _Busca.Item("Referencia")
        TextBox16.Text = _Busca.Item("FechaCaducidad")

        ToolStripButton4.Enabled = False
        ToolStripButton3.Enabled = True
        ToolStripButton2.Enabled = False

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ToolStripButton3.Enabled = False
        ToolStripButton2.Enabled = False
        ToolStripButton4.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Focus()
        TextBox17.Text = "CAMBIO"
    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton20_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton20.Click
        Button1_Click(sender, e)

    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox8.GotFocus
        TextBox8.Text = _CP
    End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox11_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox11.GotFocus
        TextBox11.Text = _Volumen
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub
End Class