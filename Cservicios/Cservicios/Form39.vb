Public Class Frm39Proveedores
    Public i As Integer
    Public _Checacontrol As Control
    Public _Busca As DataRow
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Proveedor, _RazonSocial, _Calle, _Colonia, _Ciudad, _Estado, _Contacto, _Telefono, _TelefonoContacto As String
    Public _Referencia As String

    Private Sub Form39_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click

        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True

        ' Borra los datos existentes
        Panel1.Enabled = True
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next


        ' Obtiene el Siguiente Consecutivo
        _Busca = CServiciosDataSet3.Documentos.FindByDocumento("PRO")
        Dim _Consecutivo As Integer = _Busca.Item("Consecutivo")
        Dim _XConsecutivo As String = CStr(_Consecutivo)
        _XConsecutivo = "000" & _XConsecutivo

        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "PR" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()

        TextBox12.Text = "ALTA"


    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet34.Proveedores' Puede moverla o quitarla según sea necesario.
        Me.ProveedoresTableAdapter.Fill(Me.CServiciosDataSet34.Proveedores)


        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        ToolStripButton2.Enabled = True

        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next


        DataGridView1.Rows.Clear()
        For Me.i = 0 To CServiciosDataSet34.Proveedores.Rows.Count - 1
            _Proveedor = CServiciosDataSet34.Proveedores.Rows(i).Item("Proveedor")
            _RazonSocial = CServiciosDataSet34.Proveedores.Rows(i).Item("RazonSocial")

            DataGridView1.Rows.Add(_Proveedor, _RazonSocial)
        Next
        Panel1.Enabled = False

    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Registro de Datos
        Dim _Find As Boolean
        Dim title As String
        Dim response As MsgBoxResult

        Msg = "Está seguro de guardar los datos de este Proveedor?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Mantenimiento a Tabla de Proveedores"
        response = MsgBox(Msg, Style, title)
        _Find = False

        If response = MsgBoxResult.Yes Then
            ' Revisa que no haya casillas vacías
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


            If TextBox12.Text = "ALTA" Then
                _Busca = CServiciosDataSet34.Proveedores.NewRow
                _Busca.Item("Proveedor") = TextBox1.Text
            End If

            _Busca.Item("RazonSocial") = TextBox2.Text
            _Busca.Item("Calle") = TextBox3.Text
            _Busca.Item("Colonia") = TextBox4.Text
            _Busca.Item("Ciudad") = TextBox5.Text
            _Busca.Item("Estado") = TextBox6.Text
            _Busca.Item("CP") = TextBox7.Text
            _Busca.Item("Contacto") = TextBox8.Text
            _Busca.Item("TelefonoContacto") = TextBox9.Text
            _Busca.Item("Email") = ""
            _Busca.Item("Referencia") = TextBox11.Text
            _Busca.Item("Email") = TextBox10.Text
            _Busca.Item("Comentarios") = ""
            _Busca.Item("FechaRegistro") = Mid(CStr(Today), 1, 10)
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)

            If TextBox12.Text = "ALTA" Then
                CServiciosDataSet34.Proveedores.Rows.Add(_Busca)
            End If

            Me.Validate()
            ProveedoresBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet34)
            Me.ProveedoresTableAdapter.Fill(Me.CServiciosDataSet34.Proveedores)

            If TextBox12.Text = "ALTA" Then
                ' Actualiza el Consecutivo
                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("PRO")
                _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

                Me.Validate()
                DocumentosBindingSource.EndEdit()
                Me.TableAdapterManager1.UpdateAll(CServiciosDataSet3)
                Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
            End If

            Msg = "Registro guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Button1_Click(sender, e)

        Else
            ' Perform some other action.
        End If



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

        TextBox1.Text = _Renglon.Cells(0).Value
        TextBox2.Text = _Renglon.Cells(1).Value
        _Proveedor = TextBox1.Text

        _Busca = CServiciosDataSet34.Proveedores.FindByProveedor(_Proveedor)

        TextBox3.Text = _Busca.Item("Calle")
        TextBox4.Text = _Busca.Item("Colonia")
        TextBox5.Text = _Busca.Item("Ciudad")
        TextBox6.Text = _Busca.Item("Estado")
        TextBox7.Text = _Busca.Item("CP")
        TextBox8.Text = _Busca.Item("Contacto")
        TextBox9.Text = _Busca.Item("TelefonoContacto")
        TextBox11.Text = _Busca.Item("Referencia")
        TextBox10.Text = _Busca.Item("Email")

        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        ToolStripButton3.Enabled = True
        ToolStripButton4.Enabled = False
        Panel1.Enabled = True
        TextBox1.Enabled = False
        TextBox12.Text = "CAMBIO"
    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub
End Class