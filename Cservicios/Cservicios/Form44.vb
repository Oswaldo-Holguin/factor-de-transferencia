Public Class Frm44AsignaEquipo
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Renglon As DataGridViewRow
    Public _Equipo, _Modelo, _Serie, _Descripcion, _Recurso, _Referencia, _Ubicacion As String
    Public _Seleccion As Boolean


    Private Sub Form44_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet48.Equipos' Puede moverla o quitarla según sea necesario.
        Me.EquiposTableAdapter1.Fill(Me.CServiciosDataSet48.Equipos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet37.Equipos' Puede moverla o quitarla según sea necesario.
        '   Me.EquiposTableAdapter.Fill(Me.CServiciosDataSet37.Equipos)

        DataGridView2.Rows.Clear()
        _Recurso = "SU"
        Dim _Rows_Eq() As DataRow = CServiciosDataSet48.Equipos.Select("Recurso = " & "'" & _Recurso & "'")


        For Me.I = 0 To _Rows_Eq.GetUpperBound(0)
            _Equipo = _Rows_Eq(I).Item("Equipo")


            If DBNull.Value.Equals(_Rows_Eq(I).Item("DescripcionCorta")) Then
                _Descripcion = ""
            Else
                _Descripcion = _Rows_Eq(I).Item("DescripcionCorta")
            End If
            If DBNull.Value.Equals(_Rows_Eq(I).Item("Modelo")) Then
                _Modelo = ""
            Else
                _Modelo = _Rows_Eq(I).Item("Modelo")
            End If
            If DBNull.Value.Equals(_Rows_Eq(I).Item("Serie")) Then
                _Serie = ""
            Else
                _Serie = _Rows_Eq(I).Item("Serie")
            End If
            If DBNull.Value.Equals(_Rows_Eq(I).Item("Recurso")) Then
                _Recurso = ""
            Else
                _Recurso = _Rows_Eq(I).Item("Recurso")
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

            _Seleccion = False

            DataGridView2.Rows.Add(_Equipo, _Descripcion, _Modelo, _Serie, _Recurso, _Referencia, _Seleccion, _Ubicacion)

        Next

        For Me.I = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(I).Cells(0).Value = "" Then
                Exit For
            End If
            If DataGridView2.Rows(I).Cells(4).Value <> "SU" Then
                DataGridView2.Rows(I).DefaultCellStyle.BackColor = Color.PeachPuff
                DataGridView2.Rows(I).ReadOnly = True
            End If

        Next





        ToolStripTextBox1.Text = Frm42Recursos.TextBox1.Text
        ToolStripTextBox2.Text = Frm42Recursos.TextBox2.Text
        ToolStripTextBox3.Text = Frm42Recursos.TextBox4.Text
        ToolStripTextBox4.Text = Frm42Recursos.TextBox5.Text
    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(4).Value <> "SU" Then
            Exit Sub
        End If

        If _Renglon.Cells(6).Value = False Then
            _Renglon.DefaultCellStyle.BackColor = Color.Gold
            _Renglon.Cells(6).Value = True
        Else
            _Renglon.DefaultCellStyle.BackColor = Color.White
            _Renglon.Cells(6).Value = False
        End If

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Asgina el equipo a un recurso determinado

        Dim _Busca As DataRow
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de asignar ese Equipo al Recurso Seleccionado?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Asignar Equipo a un Recurso"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            For Me.I = 0 To DataGridView2.Rows.Count - 1
                If DataGridView2.Rows(I).Cells(6).Value = True Then
                    _Equipo = DataGridView2.Rows(I).Cells(0).Value

                    _Busca = CServiciosDataSet48.Equipos.FindByEquipo(_Equipo)
                    _Busca.Item("Recurso") = ToolStripTextBox1.Text

                End If
            Next
            Me.Validate()
            EquiposBindingSource1.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet48)
            Me.EquiposTableAdapter1.Fill(Me.CServiciosDataSet48.Equipos)

            Msg = "Registros guardados Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Me.Close()

        Else

        End If


    End Sub
End Class