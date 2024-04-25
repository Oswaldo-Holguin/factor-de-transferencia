Public Class Frm32AsignaMateriales
    Public i As Integer
    Public _Material, _Descripcion, _UMedida, _Presentacion As String
    Public _Cantidad As Integer
    Public _Producto As String
    Public _OrdenProducion As String
    Public _CostoUnitario As Double
    Public _Fecha As Date
    Public Msg As String
    Public _Busca As DataRow
    Public Style As MsgBoxStyle
    Public Checacontrol As Control
    Public _Renglon As DataGridViewRow


    Private Sub Form32_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet27.Consumos' Puede moverla o quitarla según sea necesario.
        Me.ConsumosTableAdapter1.Fill(Me.CServiciosDataSet27.Consumos)


        ToolStripTextBox1.Text = Frm31OrdenProduccion.TextBox1.Text
        ToolStripTextBox2.Text = Frm31OrdenProduccion.TextBox2.Text
        ToolStripTextBox3.Text = Frm31OrdenProduccion.TextBox8.Text

        Button1_Click(sender, e)

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub ToolStripButton12_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton12.Click
        Panel1.Enabled = True

        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet5.Materiales.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet5.Materiales.Rows(i).Item("Descripcion"))
        Next

        ToolStripButton4.Enabled = True
        ToolStripButton12.Enabled = False
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet5.Materiales' Puede moverla o quitarla según sea necesario.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet26.Produccion' Puede moverla o quitarla según sea necesario.
        Me.ProduccionTableAdapter.Fill(Me.CServiciosDataSet26.Produccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet25.Consumos' Puede moverla o quitarla según sea necesario.
        Me.ConsumosTableAdapter.Fill(Me.CServiciosDataSet25.Consumos)

        Panel1.Enabled = False
        ToolStripButton4.Enabled = False

        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                Checacontrol.Text = ""
            End If
        Next


    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim _PrecioUnitario As Integer
        For Me.i = 0 To CServiciosDataSet5.Materiales.Rows.Count - 1
            If CServiciosDataSet5.Materiales.Rows(i).Item("Descripcion") = ComboBox1.Text Then
                TextBox1.Text = CServiciosDataSet5.Materiales.Rows(i).Item("Material")
                TextBox2.Text = CServiciosDataSet5.Materiales.Rows(i).Item("Umedida")
                TextBox3.Text = CServiciosDataSet5.Materiales.Rows(i).Item("Presentacion")
                _PrecioUnitario = CServiciosDataSet5.Materiales.Rows(i).Item("PrecioUnitario") / CInt(CServiciosDataSet5.Materiales.Rows(i).Item("Contenido"))
                TextBox5.Text = _PrecioUnitario
            End If
        Next
        TextBox4.Focus()
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        ' Guarda los Materiales
        Dim _PrecioUnitario As Double
        Dim title As String
        Dim response As MsgBoxResult
        Dim _Referencia As String
        Msg = "Está seguro de guardad esta Información?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Asignar Materiales a una Órden de Producción"
        response = MsgBox(Msg, Style, title)
        If response = MsgBoxResult.Yes Then
            For Me.i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value = "" Then
                    Exit For
                End If
                _Material = DataGridView1.Rows(i).Cells(0).Value
                _OrdenProducion = ToolStripTextBox1.Text
                _Fecha = CDate(ToolStripTextBox2.Text)
                _Descripcion = DataGridView1.Rows(i).Cells(1).Value
                _Producto = ToolStripTextBox3.Text
                _UMedida = DataGridView1.Rows(i).Cells(2).Value
                _Cantidad = DataGridView1.Rows(i).Cells(4).Value
                _Referencia = DataGridView1.Rows(i).Cells(3).Value
                _PrecioUnitario = DataGridView1.Rows(i).Cells(5).Value

                _Busca = CServiciosDataSet27.Consumos.NewRow
                _Busca.Item("OrdenProduccion") = _OrdenProducion
                _Busca.Item("Producto") = _Producto
                _Busca.Item("Material") = _Material
                _Busca.Item("Cantidad") = _Cantidad
                _Busca.Item("DescripcionMaterial") = _Descripcion
                _Busca.Item("Umedida") = _UMedida
                _Busca.Item("FechaOrden") = _Fecha
                _Busca.Item("FechaRegistro") = Today
                _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
                _Busca.Item("Referencia") = _Referencia
                _Busca.Item("CostoUnitario") = _PrecioUnitario

                CServiciosDataSet27.Consumos.Rows.Add(_Busca)

            Next

            Me.Validate()
            ConsumosBindingSource1.EndEdit()
            Me.TableAdapterManager3.UpdateAll(CServiciosDataSet27)
            Me.ConsumosTableAdapter1.Fill(Me.CServiciosDataSet27.Consumos)

            Msg = "Registros guardados Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Me.Close()

        Else
            Exit Sub
        End If



    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ' Agrega al Grid
        _Material = TextBox1.Text
        _Descripcion = ComboBox1.Text
        _UMedida = TextBox2.Text
        _Presentacion = TextBox3.Text
        _Cantidad = TextBox4.Text
        _CostoUnitario = TextBox5.Text

        DataGridView1.Rows.Add(_Material, _Descripcion, _UMedida, _Presentacion, _Cantidad, _CostoUnitario)

        For Each Me.Checacontrol In Panel1.Controls
            If TypeOf Checacontrol Is TextBox Then
                Checacontrol.Text = ""
            End If
        Next
        ComboBox1.Text = ""
    End Sub
End Class