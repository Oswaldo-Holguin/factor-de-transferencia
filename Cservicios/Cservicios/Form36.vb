Public Class Frm36PLOP
    Public i As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Seleccion As Boolean
    Public _Folio, _Donante, _NOmbreClinica, _Grupo As String
    Public _FechaExtraccion, _FechaCaducidad As Date


    Private Sub Frm36PLOP_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        ToolStripTextBox1.Text = Frm31OrdenProduccion.TextBox1.Text
        ToolStripTextBox2.Text = Frm31OrdenProduccion.TextBox2.Text

        Button3_Click(sender, e)

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet24.PaquetesLeucocitarios' Puede moverla o quitarla según sea necesario.
        Me.PaquetesLeucocitariosTableAdapter.Fill(Me.CServiciosDataSet24.PaquetesLeucocitarios)


        DataGridView1.Rows.Clear()

        For Me.i = 0 To CServiciosDataSet24.PaquetesLeucocitarios.Rows.Count - 1

            If CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("Documento") = "0" Then

                _Folio = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("Folio")
                _Donante = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("Donante")
                _FechaExtraccion = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("FechaExtraccion")
                _FechaCaducidad = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("FechaCaducidad")
                _Grupo = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("Grupo")
                _NOmbreClinica = CServiciosDataSet24.PaquetesLeucocitarios.Rows(i).Item("Nombreclinica")
                _Seleccion = False

                DataGridView1.Rows.Add(_Folio, _Donante, _FechaExtraccion, _FechaCaducidad, _Grupo, _NOmbreClinica, _Seleccion)
            End If

        Next
        DataGridView2.Rows.Clear()
        TextBox1.Text = DataGridView1.Rows.Count - 1
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim _Renglon As DataGridViewRow = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        If _Renglon.Cells(6).Value = True Then
            _Renglon.DefaultCellStyle.BackColor = Color.White
            _Renglon.Cells(6).Value = False
        Else
            _Renglon.Cells(6).Value = True
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        For Me.i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(0).Value = "" Then
                Exit For
            End If
            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Gold
            DataGridView1.Rows(i).Cells(6).Value = True
        Next
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        For Me.i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(0).Value = "" Then
                Exit For
            End If
            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
            DataGridView1.Rows(i).Cells(6).Value = False
        Next
    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        DataGridView2.Rows.Clear()
        For Me.i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(0).Value = "" Then
                Exit For
            End If
            If DataGridView1.Rows(i).Cells(6).Value = True Then
                _Folio = DataGridView1.Rows(i).Cells(0).Value
                _Donante = DataGridView1.Rows(i).Cells(1).Value
                _FechaExtraccion = DataGridView1.Rows(i).Cells(2).Value
                _FechaCaducidad = DataGridView1.Rows(i).Cells(3).Value
                _Grupo = DataGridView1.Rows(i).Cells(4).Value
                _NOmbreClinica = DataGridView1.Rows(i).Cells(5).Value

                DataGridView2.Rows.Add(_Folio, _Donante, _FechaExtraccion, _FechaCaducidad, _Grupo, _NOmbreClinica)
            End If
        Next
        DataGridView1.Rows.Clear()



    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        Dim _Busca As DataRow
        Dim title As String
        Dim response As MsgBoxResult
        Msg = "Está a punto de asignar el Lote de Paquetes Leucocitarios mostrados a la Orden de Produccion arriba señalada. Desea Continuar?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Paquetes Leucocitarios"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            For Me.i = 0 To DataGridView2.Rows.Count - 1
                If DataGridView2.Rows(i).Cells(0).Value = "" Then
                    Exit For
                End If
                _Folio = DataGridView2.Rows(i).Cells(0).Value
                _Busca = CServiciosDataSet24.PaquetesLeucocitarios.FindByFolio(_Folio)
                _Busca.Item("Documento") = ToolStripTextBox1.Text
            Next

            Me.Validate()
            PaquetesLeucocitariosBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet24)
            Me.PaquetesLeucocitariosTableAdapter.Fill(Me.CServiciosDataSet24.PaquetesLeucocitarios)

            Msg = "Proceso terminado Normalmente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Me.Close()
        Else
            Exit Sub
        End If
    End Sub
End Class