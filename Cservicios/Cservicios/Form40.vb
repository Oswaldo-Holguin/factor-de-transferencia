Public Class Frm40diagnósticos
    Public i As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Busca As DataRow
    Public _Checacontrol As Control
    Public _Renglon As DataGridViewRow
    Public _Diagnostico, _Descripcion, _Referencia, _Centro As String


    Private Sub Frm40diagnósticos_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
 

        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet2.Diagnosticos' Puede moverla o quitarla según sea necesario.
        Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet2.Diagnosticos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)


        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""

        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = False
        ToolStripButton2.Enabled = True

        DataGridView1.Rows.Clear()
        For Me.i = 0 To CServiciosDataSet2.Diagnosticos.Rows.Count - 1
            _Diagnostico = CServiciosDataSet2.Diagnosticos.Rows(i).Item("Diagnostico")
            _Descripcion = CServiciosDataSet2.Diagnosticos.Rows(i).Item("Descripcion")

            DataGridView1.Rows.Add(_Diagnostico, _Descripcion)
        Next
        Panel1.Enabled = False


    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        TextBox5.Text = "ALTA"
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = True

        ' Obtiene el Siguiente consecutivo

        ' Obtiene el siguiente prefijo para Personal Médico
        Dim _busca As DataRow = Me.CServiciosDataSet3.Documentos.FindByDocumento("DGN")
        Dim _Siguiente As Integer = _busca.Item("Consecutivo")

        Dim _XConsecutivo As String = "000" + CStr(_Siguiente)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        _XConsecutivo = "DG" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()

        Panel1.Enabled = True
    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As System.EventArgs) Handles ComboBox1.GotFocus
        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            _Descripcion = CServiciosDataSet.Centros.Rows(i).Item("Descripcion")

            ComboBox1.Items.Add(_Descripcion)
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.i = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            If CServiciosDataSet.Centros.Rows(i).Item("Descripcion") = ComboBox1.Text Then
                TextBox4.Text = CServiciosDataSet.Centros.Rows(i).Item("Centro")
            End If
        Next


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(0).Value
        TextBox2.Text = _Renglon.Cells(1).Value

        _Diagnostico = TextBox1.Text
        _Busca = CServiciosDataSet2.Diagnosticos.FindByDiagnostico(_Diagnostico)
        _Referencia = _Busca.Item("Referencia")
        _Centro = _Busca.Item("Centro")
        TextBox3.Text = _Referencia
        TextBox4.Text = _Centro

        _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
        ComboBox1.Text = _Busca.Item("Descripcion")

        ToolStripButton2.Enabled = False
        ToolStripButton5.Enabled = True


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        ToolStripButton5.Enabled = False
        ToolStripButton3.Enabled = True
        Panel1.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Focus()
        TextBox5.Text = "CAMBIO"
    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click

        ' Registra ALTA o CAMBIO

        Dim _Find As Boolean
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de guardar los datos de este Diagnóstico?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Mantenimiento a Tabla de Diagnósticos"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
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

            If TextBox5.Text = "ALTA" Then
                _Busca = CServiciosDataSet2.Diagnosticos.NewRow
                _Busca.Item("Diagnostico") = TextBox1.Text
            End If

            If TextBox5.Text = "CAMBIO" Then
                _Busca = CServiciosDataSet2.Diagnosticos.FindByDiagnostico(TextBox1.Text)
            End If

            _Busca.Item("Descripcion") = TextBox2.Text
            _Busca.Item("Centro") = TextBox4.Text
            _Busca.Item("Referencia") = TextBox3.Text
            _Busca.Item("Comentario") = ""
            _Busca.Item("FechaRegistro") = Mid(CStr(Today), 1, 10)
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)


            If TextBox5.Text = "ALTA" Then
                CServiciosDataSet2.Diagnosticos.Rows.Add(_Busca)
            End If

            Me.Validate()
            DiagnosticosBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet2)
            Me.DiagnosticosTableAdapter.Fill(Me.CServiciosDataSet2.Diagnosticos)

            If TextBox5.Text = "ALTA" Then
                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("DGN")
                _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

                Me.Validate()
                DocumentosBindingSource.EndEdit()
                Me.TableAdapterManager2.UpdateAll(CServiciosDataSet3)
                Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
            End If

            Msg = "Registro guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)

            Button1_Click(sender, e)


        Else
            Exit Sub
        End If





    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        Button1_Click(sender, e)
    End Sub
End Class