Public Class Frm2AltaIncapacidades
    Public _NumEmpleado As String
    Public i As Integer
    Public _Busca As DataRow
    Public msg As String
    Public Style As MsgBoxStyle
    Public checacontrol As Control

    Private Sub Frm2AltaIncapacidades_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CServiciosDataSet7.Incapacidades' table. You can move, or remove it, as needed.
        Me.IncapacidadesTableAdapter.Fill(Me.CServiciosDataSet7.Incapacidades)
        Panel1.Enabled = False

        ToolStripTextBox1.Text = Frm8Empleados.TextBox2.Text & " " & Frm8Empleados.TextBox3.Text & " " & Frm8Empleados.TextBox4.Text
        ToolStripTextBox2.Text = Frm8Empleados.TextBox1.Text
        ToolStripTextBox3.Text = Frm8Empleados.TextBox16.Text

        ToolStripButton2.Enabled = False

    End Sub

    Private Sub ToolStripButton14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton14.Click
        Panel1.Enabled = True
        TextBox1.Focus()
        ToolStripButton14.Enabled = False
        ToolStripButton2.Enabled = True

        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                checacontrol.Text = ""
            End If
        Next

        TextBox16.Text = "ALTA"

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim _vacio As Boolean

        Dim _FolioIncapacidad As String = TextBox1.Text
        Dim _NumEmpleado As String = Frm8Empleados.TextBox1.Text
        For Each Me.checacontrol In Panel1.Controls
            If TypeOf checacontrol Is TextBox Then
                If checacontrol.Text = "" Then
                    _vacio = True
                End If
                checacontrol.Text = UCase(checacontrol.Text)
            End If
        Next

        If _vacio Then
            MsgBox("Hay casillas en blanco. Debe proporcionar todos los datos.")
            Exit Sub
        End If

        ' Verifica que no se haya registrado esa Incapacidad

        For Me.i = 0 To CServiciosDataSet7.Incapacidades.Rows.Count - 1
            If CServiciosDataSet7.Incapacidades.Rows(i).Item("NumEmpleado") = _NumEmpleado Then
                If CServiciosDataSet7.Incapacidades.Rows(i).Item("FolioIncapacidad") = _FolioIncapacidad Then
                    msg = "Esa Incapacidad ya está registrada. Verifique"
                    Style = MsgBoxStyle.Information
                    MsgBox(msg, Style)
                    Exit Sub
                End If

            End If
        Next

        If TextBox16.Text = "ALTA" Then
            _Busca = CServiciosDataSet7.Incapacidades.NewRow
            _Busca.Item("NumEmpleado") = _NumEmpleado
            _Busca.Item("FolioIncapacidad") = _FolioIncapacidad
        End If

        If TextBox16.Text = "CAMBIO" Then
            _Busca = CServiciosDataSet7.Incapacidades.FindByNumEmpleadoFolioIncapacidad(_NumEmpleado, _FolioIncapacidad)
        End If

        _Busca.Item("FechaIncapacidad") = TextBox2.Text
        _Busca.Item("NumDias") = CInt(TextBox4.Text)
        _Busca.Item("Medico") = TextBox10.Text
        _Busca.Item("Clinica") = TextBox17.Text
        _Busca.Item("Referencia") = TextBox18.Text
        _Busca.Item("Comentarios") = RichTextBox1.Text
        _Busca.Item("FechaRegistro") = CStr(Today)
        _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 10)
        _Busca.Item("Horario") = Mid(TextBox3.Text, 1, 10)
        _Busca.Item("ClaveMedico") = TextBox7.Text
        _Busca.Item("CedulaProfesional") = TextBox8.Text
        _Busca.Item("Especialidad") = TextBox11.Text
        _Busca.Item("DomicilioConsultorio") = TextBox12.Text
        _Busca.Item("TelefonoConsultorio") = ""
        _Busca.Item("NumConsultorio") = TextBox9.Text
        _Busca.Item("Categoria") = TextBox13.Text
        _Busca.Item("Diagnostico") = TextBox14.Text
        _Busca.Item("Frecuencia") = TextBox15.Text
        _Busca.Item("NumAfiliacion") = ToolStripTextBox3.Text
        _Busca.Item("FechaInicio") = TextBox5.Text

        If TextBox16.Text = "ALTA" Then
            CServiciosDataSet7.Incapacidades.Rows.Add(_Busca)
        End If

        Me.Validate()
        IncapacidadesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet7)

        msg = "Registro guardado correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(msg, Style)

        Me.Close()


    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub
End Class