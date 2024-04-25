Public Class Frm10Formulas
    Public i As Integer
    Public _Busca As DataRow
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Formula, Referencia, _Comentarios As String
    Public _material, _descripcion, _Centro, _Umedida As String
    Public _PrecioUnitario As Double
    Public _checacontrol As Control



    Private Sub Frm10Formulas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet11.Formulas' Puede moverla o quitarla según sea necesario.
        Me.FormulasTableAdapter.Fill(Me.CServiciosDataSet11.Formulas)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet11.DetalleFormulas' Puede moverla o quitarla según sea necesario.
        Me.DetalleFormulasTableAdapter.Fill(Me.CServiciosDataSet11.DetalleFormulas)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet5.Materiales' Puede moverla o quitarla según sea necesario.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet8.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet8.Centros)
        Panel1.Enabled = False
        TabControl2.Enabled = False


        ' Presenta las fórmulas existentes
        DataGridView1.Rows.Clear()
        For Me.i = 0 To CServiciosDataSet11.Formulas.Rows.Count - 1
            ' If CServiciosDataSet11.Formulas.Rows(i).Item("Centro") = "FTR" Then
            _Formula = CServiciosDataSet11.Formulas.Rows(i).Item("Formula")
            _descripcion = CServiciosDataSet11.Formulas(i).Item("Descripcion")
            DataGridView1.Rows.Add(_Formula, _descripcion)
            ' End If
        Next


        ' Llena el grid con los materiales 
        DataGridView3.Rows.Clear()
        For Me.i = 0 To CServiciosDataSet5.Materiales.Rows.Count - 1
            If CServiciosDataSet5.Materiales.Rows(i).Item("Centro") = "FTR" Then
                _material = CServiciosDataSet5.Materiales.Rows(i).Item("Material")
                _descripcion = CServiciosDataSet5.Materiales.Rows(i).Item("Descripcion")
                _Umedida = CServiciosDataSet5.Materiales.Rows(i).Item("UMedida")
                _PrecioUnitario = CServiciosDataSet5.Materiales.Rows(i).Item("PrecioUnitario")
                DataGridView3.Rows.Add(_material, _descripcion, _Umedida, _PrecioUnitario)

            End If

        Next

        ' Llena el combo con los Centros
        ComboBox1.Items.Clear()
        For Me.i = 0 To CServiciosDataSet8.Centros.Rows.Count - 1
            _descripcion = CServiciosDataSet8.Centros.Rows(i).Item("Descripcion")

            ComboBox1.Items.Add(_descripcion)
        Next

        ' Deshabilita el boton de Agregar Materiales.
        Button2.Enabled = False

        ' Deshabilita la tabla de materiales
        DataGridView3.Enabled = False
        ToolStripButton3.Enabled = False

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Panel1.Enabled = True
        TextBox1.Focus()
        ToolStripButton1.Enabled = False
        ToolStripButton3.Enabled = True

        For Each Me._checacontrol In Panel1.Controls
            If TypeOf _checacontrol Is TextBox Then
                _checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""
        TextBox8.Text = "ALTA"

    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.Text = UCase(TextBox1.Text)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.Text = UCase(TextBox2.Text)

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.LostFocus
        TextBox4.Text = UCase(TextBox4.Text)
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.LostFocus
        TextBox5.Text = UCase(TextBox5.Text)
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox7_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.LostFocus
        TextBox7.Text = UCase(TextBox7.Text)

    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
        Button2.Enabled = True
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        For Me.i = 0 To CServiciosDataSet8.Centros.Rows.Count - 1
            If CServiciosDataSet8.Centros.Rows(i).Item("Descripcion") = ComboBox1.Text Then
                TextBox3.Text = CServiciosDataSet8.Centros.Rows(i).Item("Centro")
            End If
        Next
        TextBox3.Enabled = False

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim msg As String
        'Dim title As String
        Dim style As MsgBoxStyle

        msg = "Hay casillas en blanco. Debe llenar todos los datos antes de asignar materiales a esta Fórmula"
        style = MsgBoxStyle.Information

        For Each checacontrol In Panel1.Controls
            If checacontrol.text = "" Then
                MsgBox(msg, style)
                Exit Sub
            End If
        Next



        DataGridView3.Enabled = True
        TabControl2.Enabled = True
        TabControl1.SelectedIndex = 1

        TabPage2.Show()

    End Sub

    Private Sub DataGridView3_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim msg As String
        'Dim title As String
        Dim style As MsgBoxStyle

        msg = "Debe registrar los datos de la fórmula para poder seleccionar los Materiales"
        style = MsgBoxStyle.Information

        If TextBox1.Text = "" Then
            MsgBox(msg, style)
            Exit Sub
        End If

        Dim _renglon As DataGridViewRow
        _renglon = DataGridView3.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.Gold
        Dim _cantidad As Double
        _cantidad = InputBox("Teclee la cantidad requerida ")
        Dim _Porcentaje As Double = 0
        Dim _Id As Integer = 0

        _material = _renglon.Cells(0).Value
        _descripcion = _renglon.Cells(1).Value
        _Umedida = _renglon.Cells(2).Value

        ' Obtiene el Precio Unitario
        _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_material)
        _PrecioUnitario = _Busca.Item("PrecioUnitario") / CInt(_Busca.Item("Contenido"))

        ' _PrecioUnitario = _renglon.Cells(3).Value

        DataGridView2.Rows.Add(TextBox1.Text, _material, _descripcion, _cantidad, _Umedida, _PrecioUnitario, _Id)


    End Sub

    Private Sub DataGridView3_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView3_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellLeave
        Dim _renglon As DataGridViewRow
        _renglon = DataGridView3.CurrentRow
        _renglon.DefaultCellStyle.BackColor = Color.White
    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        ' Registra la Fórmula

        ' Verifica que no exista esa fórula previamente
        _Formula = TextBox1.Text
        Dim _Rows_Fo() As DataRow = CServiciosDataSet11.Formulas.Select("Formula = " & "'" & _Formula & "'")
        Dim _Find As Boolean
        Dim _FechaRegistro As String = CStr(Today)

        _Find = False
        For Me.i = 0 To _Rows_Fo.GetUpperBound(0)
            _Find = True
        Next

        If _Find Then
            Msg = "Esa fórmula ya existe. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If TextBox8.Text = "ALTA" Then
            _Busca = CServiciosDataSet11.Formulas.NewRow
            _Busca.Item("Formula") = _Formula
        End If

        _Busca.Item("Descripcion") = TextBox2.Text
        _Busca.Item("Uso") = TextBox4.Text
        _Busca.Item("Centro") = TextBox3.Text
        _Busca.Item("Umedida") = TextBox5.Text
        _Busca.Item("Cantidad") = TextBox6.Text
        _Busca.Item("Referencia") = TextBox7.Text
        _Busca.Item("Comentarios") = RichTextBox1.Text
        _Busca.Item("Usuario") = ""
        _Busca.Item("FechaRegistro") = _FechaRegistro
        _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)

        If TextBox8.Text = "ALTA" Then
            CServiciosDataSet11.Formulas.Rows.Add(_Busca)
        End If

        Me.Validate()
        FormulasBindingSource.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet11)

        Msg = "Registro guardado correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)
        Me.Close()



    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress

        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False

        else if (Char.IsControl(e.KeyChar))

            e.Handled = False

       else if (Char.IsSeparator(e.KeyChar))

            e.Handled = False

       else

            e.Handled = True
        End If

    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim _Renglon As DataGridViewRow = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        Dim _Rows_Fo(), _Rows_Df() As DataRow
        Dim _Uso, _Referencia As String
        Dim _Cantidad, _Id As Integer

        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _Formula = _Renglon.Cells(0).Value
        _Rows_Fo = CServiciosDataSet11.Formulas.Select("Formula = " & "'" & _Formula & "'")
        _Rows_Df = CServiciosDataSet11.DetalleFormulas.Select("Formula = " & "'" & _Formula & "'")

        _Uso = ""
        _Referencia = ""
        For Me.i = 0 To _Rows_Fo.GetUpperBound(0)
            _descripcion = _Rows_Fo(i).Item("Descripcion")
            _Uso = _Rows_Fo(i).Item("Uso")
            _Umedida = _Rows_Fo(i).Item("Umedida")
            _Centro = _Rows_Fo(i).Item("Centro")
            _Cantidad = _Rows_Fo(i).Item("Cantidad")
            _Referencia = _Rows_Fo(i).Item("Referencia")
            _Comentarios = _Rows_Fo(i).Item("Comentarios")
        Next

        _Busca = CServiciosDataSet8.Centros.FindByCentro(_Centro)

        TextBox1.Text = _Formula
        TextBox2.Text = _descripcion
        TextBox3.Text = _Centro
        ComboBox1.Text = _Busca.Item("Descripcion")
        TextBox4.Text = _Uso
        TextBox5.Text = _Uso
        TextBox6.Text = _Cantidad
        TextBox7.Text = _Referencia
        RichTextBox1.Text = _Comentarios

        'Presenta los materiales 
        ToolStripButton2_Click(sender, e)



    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        Panel1.Enabled = True
        TextBox1.Enabled = False
        Button2.Enabled = True
        TextBox8.Text = "CAMBIO"
        ToolStripButton3.Enabled = True
        ToolStripButton5.Enabled = False
    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        If DataGridView2.Rows.Count = 0 Then
            Exit Sub
        End If
        Dim _Cantidad As Integer
        Dim _Referencia As String
        Dim _FechaRegistro As String = CStr(Today)
        Dim _Find As Boolean

        ' Primero elimina los registros anteriores.
        _Find = False
        For Me.i = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(i).Cells(0).Value = "" Then
                Exit For
            End If
            If DataGridView2.Rows(i).Cells(6).Value <> 0 Then
                _Formula = DataGridView2.Rows(i).Cells(0).Value
                _material = DataGridView2.Rows(i).Cells(1).Value

                _Busca = CServiciosDataSet11.DetalleFormulas.FindByFormulaMaterial(_Formula, _material)
                _Busca.Delete()
                _Find = True
            End If

        Next

        If _Find Then
            Me.Validate()
            FormulasBindingSource.EndEdit()
            Me.TableAdapterManager2.UpdateAll(CServiciosDataSet11)
        End If


        ' Agrega los Materiales

        For Me.i = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(i).Cells(0).Value = "" Then
                Exit For
            End If
            _Formula = DataGridView2.Rows(i).Cells(0).Value
            _material = DataGridView2.Rows(i).Cells(1).Value
            _descripcion = DataGridView2.Rows(i).Cells(2).Value
            _Cantidad = DataGridView2.Rows(i).Cells(3).Value
            _Umedida = DataGridView2.Rows(i).Cells(4).Value
            _PrecioUnitario = DataGridView2.Rows(i).Cells(5).Value
            _Referencia = DataGridView2.Rows(i).Cells(2).Value

            _Busca = CServiciosDataSet11.DetalleFormulas.NewRow
            _Busca.Item("Formula") = _Formula
            _Busca.Item("Material") = _material
            _Busca.Item("Cantidad") = _Cantidad
            _Busca.Item("Porcentaje") = 0
            _Busca.Item("Referencia") = _Referencia
            _Busca.Item("Comentarios") = RichTextBox1.Text
            _Busca.Item("FechaRegistro") = Mid(_FechaRegistro, 1, 10)
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)

            CServiciosDataSet11.DetalleFormulas.Rows.Add(_Busca)
        Next

        Me.Validate()
        DetalleFormulasBindingSource.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet11)

        Msg = "Registros guardados correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        ' Muestra los materiales de esta fórmula
        Me.DetalleFormulasTableAdapter.Fill(Me.CServiciosDataSet11.DetalleFormulas)
        Dim _Cantidad As Double
        _Formula = TextBox1.Text
        Dim _Rows_Df() As DataRow
        Dim _Id As Integer
        Dim _CostosDirectos As Double
        _Rows_Df = CServiciosDataSet11.DetalleFormulas.Select("Formula = " & "'" & _Formula & "'")

        ' Presenta los materiales que se utlizan en esta fórmula-
        _CostosDirectos = 0
        DataGridView2.Rows.Clear()
        For Me.i = 0 To _Rows_Df.GetUpperBound(0)
            _material = _Rows_Df(i).Item("Material")
            _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_material)
            _descripcion = _Busca.Item("Descripcion")
            _Cantidad = _Rows_Df(i).Item("Cantidad")
            _Umedida = _Busca.Item("Umedida")
            _PrecioUnitario = _Busca.Item("PrecioUnitario") / CInt(_Busca.Item("Contenido"))
            _Id = _Busca.Item("Orden")
            _CostosDirectos = _CostosDirectos + (_PrecioUnitario * _Cantidad)
            DataGridView2.Rows.Add(_Formula, _material, _descripcion, _Cantidad, _Umedida, _PrecioUnitario, _Id)

        Next

        TextBox9.Text = Format(_CostosDirectos, "$###,##0.00")

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub EliminarRegistroToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EliminarRegistroToolStripMenuItem.Click
        ' Elimina Registro

        Dim _Renglon As DataGridViewRow = DataGridView2.CurrentRow


        Dim _Id As Integer = _Renglon.Cells(6).Value
        Dim title As String


        Dim response As MsgBoxResult
        Msg = "Está a punto de eliminar este Registro. Desea continuar?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Eliminar Materiales de esta Fórmula"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Elimina el Material
            _material = _Renglon.Cells(1).Value
            _Formula = _Renglon.Cells(0).Value
            _Busca = CServiciosDataSet11.DetalleFormulas.FindByFormulaMaterial(_Formula, _material)
            _Busca.Delete()
            Me.Validate()
            DetalleFormulasBindingSource.EndEdit()
            Me.TableAdapterManager2.UpdateAll(CServiciosDataSet11)
            Msg = "Registro eliminado correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            ToolStripButton2_Click(sender, e)

            Exit Sub
        Else
            Exit Sub
        End If

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim _Renglon As DataGridViewRow = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold


    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub
End Class