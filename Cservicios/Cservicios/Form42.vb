Public Class Frm42Recursos
    Public I As Integer
    Public _Checacontrol As Control
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Busca As DataRow
    Public _Recurso, _DescripcionLarga, _DescripcionCorta, _Campus, _Ubicacion, _Contacto As String
    Public _TelefonoContacto, _CelularContacto, _Referencia, _RE As String
    Public _Renglon As DataGridViewRow

    Private Sub Frm42Recursos_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet48.Equipos' Puede moverla o quitarla según sea necesario.
        Me.EquiposTableAdapter1.Fill(Me.CServiciosDataSet48.Equipos)


        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet36.Recursos' Puede moverla o quitarla según sea necesario.
        Me.RecursosTableAdapter.Fill(Me.CServiciosDataSet36.Recursos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet36.DetalleRecursos' Puede moverla o quitarla según sea necesario.
        Me.DetalleRecursosTableAdapter.Fill(Me.CServiciosDataSet36.DetalleRecursos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet37.Equipos' Puede moverla o quitarla según sea necesario.
        Me.EquiposTableAdapter.Fill(Me.CServiciosDataSet37.Equipos)

        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        For Me.I = 0 To CServiciosDataSet36.Recursos.Count - 1
            _Recurso = CServiciosDataSet36.Recursos.Rows(I).Item("ReccursoServicio")
            _DescripcionLarga = CServiciosDataSet36.Recursos.Rows(I).Item("DescripcionLarga")
            _Campus = CServiciosDataSet36.Recursos.Rows(I).Item("Campus")

            DataGridView1.Rows.Add(_Recurso, _DescripcionLarga, _Campus)

        Next

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If

        Next
        RadioButton1.Checked = True

        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = False
        ToolStripButton2.Enabled = True
        Panel1.Enabled = False

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Enabled = True

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.Enabled = True
            End If

        Next
        RadioButton1.Checked = True

        TextBox1.Focus()
        TextBox13.Text = "ALTA"
        ToolStripButton12.Enabled = False
        ToolStripButton3.Enabled = True

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus
        ' Obtiene el consecutivo para los Recursos
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("R/S")
        Dim _Consecutivo As Integer = _Busca.Item("Consecutivo")
        Dim _XConsecutivo As String = "RE" & Trim(CStr(_Consecutivo))
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()


    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        TextBox13.Text = "CAMBIO"
        TextBox2.Focus()
        ToolStripButton5.Enabled = False
        ToolStripButton3.Enabled = True
        Panel1.Enabled = True
        TextBox1.Enabled = False

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guardar datos

        Dim _Find As Boolean
        _Find = False

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

        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            Msg = "Debe seleccionar RECURSO o SERVICIO"
            RadioButton1.Focus()
            Exit Sub
        End If

        _Recurso = TextBox1.Text
        If TextBox13.Text = "ALTA" Then
            _Busca = CServiciosDataSet36.Recursos.NewRow
            _Busca.Item("ReccursoServicio") = _Recurso
        End If

        If TextBox13.Text = "CAMBIO" Then
            _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
        End If

        _DescripcionLarga = TextBox2.Text
        _DescripcionCorta = Mid(TextBox3.Text, 1, 20)
        _Ubicacion = TextBox4.Text
        _Campus = Mid(TextBox5.Text, 1, 20)
        _Contacto = TextBox6.Text
        _TelefonoContacto = Mid(TextBox7.Text, 1, 20)
        _CelularContacto = Mid(TextBox9.Text, 1, 20)
        _Referencia = TextBox8.Text
        If RadioButton1.Checked = True Then
            _RE = "R"
        End If
        If RadioButton2.Checked = True Then
            _RE = "S"
        End If

        _Busca.Item("DescripcionLarga") = _DescripcionLarga
        _Busca.Item("DescripcionCorta") = _DescripcionCorta
        _Busca.Item("Ubicacion") = _Ubicacion
        _Busca.Item("Campus") = _Campus
        _Busca.Item("Contacto") = _Contacto
        _Busca.Item("TelefonoContacto") = _TelefonoContacto
        _Busca.Item("CelularContacto") = _CelularContacto
        _Busca.Item("Referencia") = _Referencia
        _Busca.Item("RecSer") = _RE

        If TextBox13.Text = "ALTA" Then
            CServiciosDataSet36.Recursos.Rows.Add(_Busca)
        End If

        Me.Validate()
        RecursosBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet36)
        Me.RecursosTableAdapter.Fill(Me.CServiciosDataSet36.Recursos)

        ' Actualizar el consecutivo de los Documentos
        If TextBox13.Text = "ALTA" Then
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("R/S")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        End If

        Msg = "Registro guardado correctamente"
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
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _Recurso = _Renglon.Cells(0).Value
        TextBox1.Text = _Recurso

        _Busca = CServiciosDataSet36.Recursos.FindByReccursoServicio(_Recurso)
        TextBox2.Text = _Busca.Item("DescripcionLarga")
        TextBox3.Text = _Busca.Item("DescripcionCorta")
        TextBox4.Text = _Busca.Item("Ubicacion")
        TextBox5.Text = _Busca.Item("Campus")
        TextBox6.Text = _Busca.Item("Contacto")
        TextBox7.Text = _Busca.Item("TelefonoContacto")
        TextBox8.Text = _Busca.Item("Referencia")
        TextBox9.Text = _Busca.Item("CelularContacto")
        If _Busca.Item("RecSer") = "R" Then
            RadioButton1.Checked = True
        Else
            RadioButton2.Checked = True
        End If


        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = True

        TextBox1.Enabled = False

        ' Presenta el Equípo asignado
        DataGridView2.Rows.Clear()
        ToolStripButton12_Click(sender, e)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton13_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton13.Click
        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar un Recurso para asignarle Equipo"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        ToolStripButton12.Enabled = True
        Frm44AsignaEquipo.Show()

    End Sub

    Private Sub ToolStripButton12_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton12.Click
        ' Presenta el equipo asignado a este Recurso
        If TextBox1.Text = "" Then
            Exit Sub
        End If

        '   _Recurso = "SU"
        Me.EquiposTableAdapter1.Fill(Me.CServiciosDataSet48.Equipos)
        _Recurso = TextBox1.Text
        Dim _Rows_Eq() As DataRow = CServiciosDataSet48.Equipos.Select("Recurso = " & "'" & _Recurso & "'")
        Dim _Equipo, _Modelo, _Descripcion, _Serie As String

        DataGridView2.Rows.Clear()
        For Me.I = 0 To _Rows_Eq.GetUpperBound(0)
            If DBNull.Value.Equals(_Rows_Eq(I).Item("Equipo")) Then
                _Equipo = ""
            Else
                _Equipo = _Rows_Eq(I).Item("Equipo")
            End If

            If DBNull.Value.Equals(_Rows_Eq(I).Item("DescripcionLarga")) Then
                _Descripcion = ""
            Else
                _Descripcion = _Rows_Eq(I).Item("Descripcionlarga")
            End If
            If DBNull.Value.Equals(_Rows_Eq(I).Item("Serie")) Then
                _Serie = ""
            Else
                _Serie = _Rows_Eq(I).Item("Serie")
            End If

            If DBNull.Value.Equals(_Rows_Eq(I).Item("Modelo")) Then
                _Modelo = ""
            Else
                _Modelo = _Rows_Eq(I).Item("Modelo")
            End If

            If DBNull.Value.Equals(_Rows_Eq(I).Item("Referencia")) Then
                _Referencia = ""
            Else
                _Referencia = _Rows_Eq(I).Item("Referencia")
            End If
            If DBNull.Value.Equals(_Rows_Eq(I).Item("Ubicacion")) Then
                _Ubicacion = ""
            Else
                _Ubicacion = _Rows_Eq(I).Item("Ubicacion")
            End If


            DataGridView2.Rows.Add(_Equipo, _Descripcion, _Serie, _Modelo, _Referencia, _Ubicacion)

        Next


    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold


    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellLeave
        _Renglon = DataGridView2.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub EliminarRegistroToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EliminarRegistroToolStripMenuItem.Click
        Dim _Equipo As String = _Renglon.Cells(0).Value
        _Busca = CServiciosDataSet48.Equipos.FindByEquipo(_Equipo)
        _Busca.Item("Recurso") = "SU"

        Me.Validate()
        EquiposBindingSource1.EndEdit()
        Me.TableAdapterManager3.UpdateAll(CServiciosDataSet48)
        Me.EquiposTableAdapter1.Fill(Me.CServiciosDataSet48.Equipos)

        Msg = "Registro Eliminado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        ToolStripButton12_Click(sender, e)

    End Sub
End Class