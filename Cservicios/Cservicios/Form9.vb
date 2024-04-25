Public Class Frm9NuevoPermiso

    Private Sub Frm9NuevoPermiso_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CServiciosDataSet3.Documentos' table. You can move, or remove it, as needed.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: This line of code loads data into the 'CServiciosDataSet7.Permisos' table. You can move, or remove it, as needed.
        Me.PermisosTableAdapter.Fill(Me.CServiciosDataSet7.Permisos)
        Dim _NombreEmpleado As String = Trim(Frm8Empleados.TextBox2.Text) & " " & Trim(Frm8Empleados.TextBox3.Text) & " " & Trim(Frm8Empleados.TextBox4.Text)
        ToolStripTextBox1.Text = _NombreEmpleado
        Panel1.Enabled = False
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        Panel1.Enabled = True
        ToolStripButton1.Enabled = False
        ' TextBox1.Focus()
        RadioButton1.Focus()
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        Dim _Consecutivo As Integer


        For I = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If RadioButton1.Checked = True Then
                If CServiciosDataSet3.Documentos.Rows(I).Item("Documento") = "PEE" Then
                    _Consecutivo = CServiciosDataSet3.Documentos.Rows(I).Item("Consecutivo")
                End If
            End If
            If RadioButton2.Checked = True Then
                If CServiciosDataSet3.Documentos.Rows(I).Item("Documento") = "PES" Then
                    _Consecutivo = CServiciosDataSet3.Documentos.Rows(I).Item("Consecutivo")
                End If
            End If
        Next


        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        If RadioButton1.Checked = True Then
            _XConsecutivo = "PE" + Mid(_XConsecutivo, _x, 4)
        End If
        If RadioButton2.Checked = True Then
            _XConsecutivo = "PS" + Mid(_XConsecutivo, _x, 4)
        End If
        TextBox1.Text = _XConsecutivo
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox2.Text = CStr(Today)
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = TimeString
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        ' Agrega el registro
        Dim checacontrol As Control
        For Each checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = UCase(checacontrol.Text)
            End If
        Next
        RichTextBox1.Text = UCase(RichTextBox1.Text)


        Dim _fila As DataRow = CServiciosDataSet7.Permisos.NewRow
        _fila.Item("NumEmpleado") = Frm8Empleados.TextBox1.Text
        _fila.Item("Permiso") = TextBox1.Text
        _fila.Item("Referencia") = TextBox10.Text
        _fila.Item("Comentarios") = RichTextBox1.Text
        _fila.Item("Motivo") = TextBox4.Text
        _fila.Item("NumDias") = CInt(TextBox5.Text)
        _fila.Item("NumHoras") = CInt(TextBox6.Text)
        _fila.Item("NumMinutos") = CInt(TextBox7.Text)
        _fila.Item("Atiende") = TextBox8.Text
        _fila.Item("Autoriza") = TextBox9.Text
        _fila.Item("FechaPermiso") = TextBox2.Text
        _fila.Item("FechaRegistro") = CStr(Today)
        _fila.Item("HoraRegistro") = TextBox3.Text

        CServiciosDataSet7.Permisos.Rows.Add(_fila)
        Me.Validate()
        PermisosBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet7)
        Me.PermisosTableAdapter.Fill(CServiciosDataSet7.Permisos)

        ' Incrementa el consecutivo de los permisos

        Dim _Busca As DataRow
        If RadioButton1.Checked = True Then
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("PEE")
        End If
        If RadioButton2.Checked = True Then
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("PES")
        End If

        _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
        Me.Validate()
        DocumentosBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(CServiciosDataSet3)


        MsgBox("Registro guardado correctamente")
        Me.Close()

    End Sub

    Private Sub TextBox5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.GotFocus
        TextBox5.Text = 0
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox6.Text = 0
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.GotFocus
        TextBox7.Text = 0
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.Text = UCase(TextBox4.Text)
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox8_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.LostFocus
        TextBox5.Text = UCase(TextBox5.Text)
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox9_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox9.LostFocus
        TextBox9.Text = UCase(TextBox9.Text)
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox10.LostFocus
        TextBox10.Text = UCase(TextBox10.Text)
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub
End Class