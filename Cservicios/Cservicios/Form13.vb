Public Class Frm13ElectrosParaHoy
    Public _cuantos As Integer
    Public i As Integer
    Public _Paciente, _Centro As String


    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        ' Muetra los Electros para Hoy.
        _Cuantos = CServiciosDataSet4.Citas.Rows.Count
        Dim _Fechahoy = Today

        Dim _Selecciona As String = "FechaCita = " & "'" & _Fechahoy & "'"
        Dim _Rows_ci() As DataRow = CServiciosDataSet4.Citas.Select(_Selecciona)
        Dim _FechaCita As String
        Dim _Horacita As String
        Dim _Atiende As String
        Dim _Cita As String
        Dim _NombrePaciente As String
        Dim _Asiste As Boolean
        Dim _Busca As DataRow
        Dim _Referencia, _comentarios As String

        DataGridView2.Rows.Clear()
        For Me.i = 0 To _Rows_ci.GetUpperBound(0)
            If _Rows_ci(i).Item("Centro") = "EEE" Then
                _Paciente = _Rows_ci(i).Item("Paciente")
                _Cita = _Rows_ci(i).Item("Cita")
                _FechaCita = _Rows_ci(i).Item("FechaCita")
                _Horacita = _Rows_ci(i).Item("HoraCita")
                _Centro = _Rows_ci(i).Item("Centro")
                _Busca = CServiciosDataSet2.Pacientes.FindByPaciente(_Paciente)
                _NombrePaciente = _Busca.Item("Paterno") & " " & _Busca.Item("Materno") & " " & _Busca.Item("Nombres")
                _Referencia = _Rows_ci(i).Item("Referencia")
                _Atiende = _Rows_ci(i).Item("Atiende")
                _comentarios = _Rows_ci(i).Item("Comentarios")
                _Asiste = _Rows_ci(i).Item("Asiste")
                DataGridView2.Rows.Add(_Paciente, _Cita, _FechaCita, _Horacita, _Centro, _NombrePaciente, _Atiende, _Referencia, _comentarios, _Asiste)
            End If
        Next

        For Me.i = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(i).Cells(9).Value = True Then
                DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
                DataGridView2.Rows(i).ReadOnly = True

            End If
        Next





    End Sub

    Private Sub Frm13ElectrosParaHoy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CServiciosDataSet8.Centros' table. You can move, or remove it, as needed.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)
        'TODO: This line of code loads data into the 'CServiciosDataSet2.Pacientes' table. You can move, or remove it, as needed.
        Me.PacientesTableAdapter.Fill(Me.CServiciosDataSet2.Pacientes)
        'TODO: This line of code loads data into the 'CServiciosDataSet4.Citas' table. You can move, or remove it, as needed.
        Me.CitasTableAdapter.Fill(Me.CServiciosDataSet4.Citas)
        ToolStripButton2_Click(sender, e)
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub
End Class