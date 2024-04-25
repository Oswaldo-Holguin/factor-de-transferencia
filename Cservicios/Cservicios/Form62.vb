Public Class Frm62HistorialElectros
    Public I As Integer
    Public _Checacontrol As Control
    Public Msg As String
    Public _Renglon As DataGridViewRow
    Public _Fecha As Date
    Public _Busca As DataRow
    Public _Consecutivo As String
    Public _Paciente As String
    Public Style As MsgBoxStyle
    Public _Electro, _Hora, _Referencia, _NombrePaciente, _Comentarios As String
    Public _Exitoso As Boolean


    Private Sub Form62_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet17.Electros' Puede moverla o quitarla según sea necesario.
        Me.ElectrosTableAdapter.Fill(Me.CServiciosDataSet17.Electros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Pacientes' Puede moverla o quitarla según sea necesario.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)

        Dim _Rows_Pa() As DataRow
        Dim X As Integer

        DataGridView1.Rows.Clear()
        For Me.I = 0 To Me.CServiciosDataSet17.Electros.Rows.Count - 1
            _Electro = Me.CServiciosDataSet17.Electros.Rows(I).Item("Electro")
            _Paciente = Me.CServiciosDataSet17.Electros.Rows(I).Item("Paciente")
            _Hora = Me.CServiciosDataSet17.Electros.Rows(I).Item("HoraElectro")
            _Fecha = Me.CServiciosDataSet17.Electros.Rows(I).Item("FechaElectro")
            _Consecutivo = Me.CServiciosDataSet17.Electros.Rows(I).Item("Consecutivo")
            _NombrePaciente = ""

            _Rows_Pa = Me.CServiciosDataSet2.Pacientes.Select("Paciente = " & "'" & _Paciente & "'")

            For X = 0 To _Rows_Pa.GetUpperBound(0)
                _NombrePaciente = _Rows_Pa(X).Item("Paterno") & " " & _Rows_Pa(X).Item("Materno") & " " & _Rows_Pa(X).Item("Nombres")
            Next

            DataGridView1.Rows.Add(_Electro, _Fecha, _Consecutivo, _NombrePaciente)
        Next

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.BackColor = Color.White
            End If
        Next
        Timer1.Enabled = False
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = False
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

        _Electro = _Renglon.Cells(0).Value
        TextBox1.Text = _Electro
        TextBox1.Enabled = False
        _Busca = Me.CServiciosDataSet17.Electros.FindByElectro(_Electro)
        TextBox2.Text = _Busca.Item("FechaElectro")
        TextBox3.Text = _Busca.Item("HoraElectro")
        TextBox4.Text = _Busca.Item("Paciente")
        Dim _Rows_Pa() As DataRow = Me.CServiciosDataSet2.Pacientes.Select("Paciente = " & "'" & _Paciente & "'")

        _NombrePaciente = _Busca.Item("IdElecro")
        For Me.I = 0 To _Rows_Pa.GetUpperBound(0)
            _NombrePaciente = _Rows_Pa(I).Item("Paterno") & " " & _Rows_Pa(I).Item("Materno") & " " & _Rows_Pa(I).Item("Nombres")
        Next

        ComboBox1.Text = _NombrePaciente
        TextBox5.Text = _Busca.Item("Consecutivo")
        TextBox6.Text = _Busca.Item("Consulta")
        TextBox7.Text = _Busca.Item("Referencia")
        If _Busca.Item("Exitoso") = True Then
            RadioButton1.Checked = True
        Else
            RadioButton2.Checked = True
        End If
        TextBox8.Text = _Busca.Item("Montaje")
        TextBox9.Text = _Busca.Item("Sensibilidad")
        TextBox10.Text = _Busca.Item("Frecuencia")

        RichTextBox1.Text = _Busca.Item("Comentarios")
        ToolStripButton2.Enabled = True

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click

        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar una Electro para modificar los datos"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        Timer1.Enabled = True
        Timer1.Interval = 100
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True
        TextBox2.Focus()

        ComboBox1.Items.Clear()
        For Me.I = 0 To Me.CServiciosDataSet2.Pacientes.Rows.Count - 1
            ComboBox1.Items.Add(Me.CServiciosDataSet2.Pacientes.Rows(I).Item("Paterno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(I).Item("Materno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(I).Item("Nombres"))
        Next
        ComboBox1.Sorted = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.I = 0 To Me.CServiciosDataSet2.Pacientes.Rows.Count - 1
            _NombrePaciente = Me.CServiciosDataSet2.Pacientes.Rows(I).Item("Paterno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(I).Item("Materno") & " " & Me.CServiciosDataSet2.Pacientes.Rows(I).Item("Nombres")
            If _NombrePaciente = ComboBox1.Text Then
                TextBox4.Text = Me.CServiciosDataSet2.Pacientes.Rows(I).Item("Paciente")
                Exit For
            End If
        Next
    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Guarda información
        _Fecha = TextBox2.Text
        If Not IsDate(_Fecha) Then
            Msg = "Debe teclear una fecha válida de este Electro"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está a punto de guardar los datos modificados. Está seguro de continuar?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Modificación a registros de Electros"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Guarda los datos
            _Electro = TextBox1.Text






        Else
            ' Perform some other action.
        End If
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Focused Then
                    _Checacontrol.BackColor = Color.Gold
                End If
            End If
        Next
    End Sub
End Class