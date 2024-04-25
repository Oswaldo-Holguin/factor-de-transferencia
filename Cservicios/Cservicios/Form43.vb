Public Class Frm43Equipos
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Busca As DataRow
    Public _Checacontrol As Control
    Public _Equipo, _DescripcionLarga, _DescripcionCorta, _Uso, _Serie, _Modelo, _Recurso, _DescripcionRecurso As String
    Public _Referencia, _Comentarios, _Campus, _Marca, _Color, _FechaAsigna As String


    Private Sub Frm43Equipos_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
     
        Button1_Click(sender, e)


    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet37.Equipos' Puede moverla o quitarla según sea necesario.
        '  Me.EquiposTableAdapter.Fill(Me.CServiciosDataSet37.Equipos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet36.Recursos' Puede moverla o quitarla según sea necesario.
        Me.RecursosTableAdapter.Fill(Me.CServiciosDataSet36.Recursos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet48.Equipos' Puede moverla o quitarla según sea necesario.
        Me.EquiposTableAdapter1.Fill(Me.CServiciosDataSet48.Equipos)
        Panel1.Enabled = False

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        CheckBox1.Checked = False
        RichTextBox1.Text = ""

        ToolStripButton1.Enabled = True
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True
        ToolStripButton5.Enabled = False

        DataGridView1.Rows.Clear()


        '     Dim _Cuantos As Integer = CServiciosDataSet48.Equipos.Rows.Count
        '    MsgBox("Registros " & CStr(_Cuantos))
        'Me.Close()


        For Me.I = 0 To CServiciosDataSet48.Equipos.Rows.Count - 1


            _Equipo = CServiciosDataSet48.Equipos.Rows(I).Item("Equipo")
            _DescripcionCorta = CServiciosDataSet48.Equipos.Rows(I).Item("DescripcionCorta")

            If DBNull.Value.Equals(CServiciosDataSet48.Equipos.Rows(I).Item("Recurso")) Then
                _Recurso = "SU"
            Else
                _Recurso = CServiciosDataSet48.Equipos.Rows(I).Item("Recurso")
            End If
            _Recurso = CServiciosDataSet48.Equipos.Rows(I).Item("Recurso")
            _Referencia = CServiciosDataSet48.Equipos.Rows(I).Item("Referencia")

            If _Recurso <> "SU" Then
                _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
                _Campus = _Busca.Item("Campus")
            Else
                _Campus = "NO ASIGNADO"
            End If

            DataGridView1.Rows.Add(_Equipo, _DescripcionCorta, _Referencia)

        Next

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        CheckBox1.Checked = True

        TextBox13.Text = "ALTA"
        Panel1.Enabled = True

        Timer1.Enabled = True
        Timer1.Interval = 100


        ' Obtien el Consecutivo
        TextBox1.Focus()

        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
        ToolStripButton5.Enabled = False



    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs)
        ' Obtiene el consecutivo para los Recursos

        'MsgBox("Hola")

        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("EQU")
        Dim _Consecutivo As Integer = _Busca.Item("Consecutivo")
        Dim _XConsecutivo As String = "EQ" & Trim(CStr(_Consecutivo))
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()

        TextBox7.Text = "SU"
        TextBox9.Text = "SIN UBICACION"
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guarda la información
        Dim _Find As Boolean

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = UCase(_Checacontrol.Text)
                If _Checacontrol.Text = "" Then
                    _Find = True
                End If
            End If
        Next
        If _Find Then
            Msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        _Equipo = TextBox1.Text
        _DescripcionLarga = TextBox2.Text
        _DescripcionCorta = TextBox3.Text
        _Uso = TextBox4.Text
        _Modelo = TextBox5.Text
        _Serie = TextBox6.Text
        _Recurso = TextBox7.Text
        _Referencia = TextBox8.Text
        _Comentarios = RichTextBox1.Text
        _Marca = TextBox11.Text
        _FechaAsigna = CStr(TextBox12.Text)
        _Color = TextBox14.Text


        If TextBox13.Text = "ALTA" Then
            _Busca = CServiciosDataSet48.Equipos.NewRow
            _Busca.Item("Equipo") = _Equipo
        End If

        If TextBox13.Text = "CAMBIO" Then
            _Busca = CServiciosDataSet48.Equipos.FindByEquipo(_Equipo)
        End If

        _Busca.Item("DescripcionLarga") = _DescripcionLarga
        _Busca.Item("DescripcionCorta") = _DescripcionCorta
        _Busca.Item("Uso") = _Uso
        _Busca.Item("Modelo") = _Modelo
        _Busca.Item("Serie") = _Serie
        _Busca.Item("Recurso") = TextBox7.Text
        _Busca.Item("Referencia") = _Referencia
        _Busca.Item("Comentarios") = _Comentarios
        _Busca.Item("Disponible") = CheckBox1.Checked
        _Busca.Item("Marca") = _Marca
        _Busca.Item("Color") = _Color
        _Busca.Item("Documento") = _FechaAsigna


        If TextBox13.Text = "ALTA" Then
            CServiciosDataSet48.Equipos.Rows.Add(_Busca)
        End If

        Me.Validate()
        EquiposBindingSource.EndEdit()
        Me.TableAdapterManager3.UpdateAll(CServiciosDataSet48)
        Me.EquiposTableAdapter1.Fill(Me.CServiciosDataSet48.Equipos)

        If TextBox13.Text = "ALTA" Then
            ' Actualiza el Consecutivo
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("EQU")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        End If

        Msg = "Registro guardado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If


        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = True
            End If
        Next
        RichTextBox1.Enabled = True
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _Equipo = _Renglon.Cells(0).Value

        _Busca = CServiciosDataSet48.Equipos.FindByEquipo(_Equipo)
        TextBox1.Text = _Equipo
        If DBNull.Value.Equals(_Busca.Item("DescripcionLarga")) Then
            TextBox2.Text = ""
        Else
            TextBox2.Text = _Busca.Item("DescripcionLarga")
        End If
        If DBNull.Value.Equals(_Busca.Item("DescripcionCorta")) Then
            TextBox3.Text = ""
        Else
            TextBox3.Text = _Busca.Item("DescripcionCorta")
        End If

        If DBNull.Value.Equals(_Busca.Item("Uso")) Then
            TextBox4.Text = ""
        Else
            TextBox4.Text = _Busca.Item("Uso")
        End If

        If DBNull.Value.Equals(_Busca.Item("Serie")) Then
            TextBox6.Text = ""
        Else
            TextBox6.Text = _Busca.Item("Serie")
        End If
        If DBNull.Value.Equals(_Busca.Item("Modelo")) Then
            TextBox5.Text = ""
        Else
            TextBox5.Text = _Busca.Item("Modelo")
        End If

        If DBNull.Value.Equals(_Busca.Item("Recurso")) Then
            TextBox7.Text = ""
        Else
            TextBox7.Text = _Busca.Item("Recurso")
        End If

        If DBNull.Value.Equals(_Busca.Item("Referencia")) Then
            TextBox8.Text = ""
        Else
            TextBox8.Text = _Busca.Item("Referencia")
        End If

        CheckBox1.Checked = _Busca.Item("Disponible")

        TextBox11.Text = _Busca.Item("Marca")
        If DBNull.Value.Equals(_Busca.Item("Comentarios")) Then
            RichTextBox1.Text = ""
        Else
            RichTextBox1.Text = _Busca.Item("Comentarios")
        End If

        If DBNull.Value.Equals(_Busca.Item("Ubicacion")) Then
            TextBox10.Text = ""
        Else
            TextBox10.Text = _Busca.Item("Ubicacion")
        End If
        If DBNull.Value.Equals(_Busca.Item("Documento")) Then
            TextBox12.Text = ""
        Else
            TextBox12.Text = CDate(_Busca.Item("Documento"))
        End If

        If DBNull.Value.Equals(_Busca.Item("Color")) Then
            TextBox14.Text = ""
        Else
            TextBox14.Text = _Busca.Item("Color")
        End If


        _Recurso = TextBox7.Text
        If _Recurso = "SU" Then
            TextBox9.Text = "SIN UBICACION"
        Else
            _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
            TextBox9.Text = _Busca.Item("DescripcionLarga")
        End If


        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True
        ToolStripButton5.Enabled = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Panel1.Enabled = True
        TextBox1.Enabled = False
        TextBox13.Text = "CAMBIO"
        TextBox2.Focus()
        ToolStripButton5.Enabled = False
        ToolStripButton3.Enabled = True
        ToolStripButton4.Enabled = True
    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub TextBox1_GotFocus1(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus
        ' Obtiene el consecutivo para los Recursos

        '   MsgBox("Hola")

        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("EQU")
        Dim _Consecutivo As Integer = _Busca.Item("Consecutivo")
        Dim _XConsecutivo As String = "EQ" & Trim(CStr(_Consecutivo))
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()

        TextBox7.Text = "SU"
        TextBox9.Text = "SIN UBICACION"
    End Sub

  

    Private Sub BindingNavigator1_RefreshItems(sender As System.Object, e As System.EventArgs) Handles BindingNavigator1.RefreshItems

    End Sub

    Private Sub TextBox12_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox12.GotFocus
        TextBox12.Text = Today.Date
    End Sub

    Private Sub TextBox12_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox12.LostFocus
        Dim _Validafecha As DateTime

        If Not DateTime.TryParse(TextBox12.Text, _Validafecha) Then

            Msg = "la fecha ingresada no es valida"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            TextBox12.Focus()

        End If
        TextBox12.BackColor = Color.White

    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Focused Then
                    _Checacontrol.BackColor = Color.PeachPuff
                End If
            End If
            If TypeOf _Checacontrol Is RichTextBox Then
                If _Checacontrol.Focused Then
                    _Checacontrol.BackColor = Color.PeachPuff
                End If
            End If
        Next

    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.BackColor = Color.White
    End Sub

    Private Sub TextBox1_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.BackColor = Color.White
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox4_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.BackColor = Color.White
    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox6.LostFocus
        TextBox6.BackColor = Color.White
    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox8_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox8.LostFocus
        TextBox8.BackColor = Color.White
    End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox10_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox10.LostFocus
        TextBox10.BackColor = Color.White
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox11_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox11.LostFocus
        TextBox11.BackColor = Color.White
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox14_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox14.LostFocus
        TextBox4.BackColor = Color.White
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub RichTextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.BackColor = Color.White
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub
End Class