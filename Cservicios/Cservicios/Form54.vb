Public Class Frm54Requisiciones
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Busca As DataRow
    Public _Renglon As DataGridViewRow
    Public _Checacontrol As Control
    Public _Requisicion, _Solicitante, _Centro, _Material, _Descripcion, _Comentarios As String
    Public _FechaRequisicion As Date
    Public _HoraRequisicion, _UMedida, _Color, _Marca, _Serie, _Modelo As String
    Public _Cantidad As Double

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        TextBox1.Text = _Renglon.Cells(0).Value

        _Busca = CServiciosDataSet44.Requisiciones.FindByRequisicion(TextBox1.Text)
        TextBox2.Text = _Busca.Item("FechaRequisicion")
        TextBox3.Text = _Busca.Item("HoraRequisicion")
        TextBox4.Text = _Busca.Item("Centro")
        TextBox5.Text = _Busca.Item("Solicitante")
        TextBox6.Text = _Busca.Item("Referencia")
        TextBox7.Text = _Busca.Item("Prioridad")
        TextBox8.Text = _Busca.Item("TipoRequisicion")
        RichTextBox1.Text = _Busca.Item("Comentarios")

        _Busca = CServiciosDataSet.Centros.FindByCentro(TextBox4.Text)
        ComboBox1.Text = _Busca.Item("Descripcion")
        _Busca = CServiciosDataSet42.Prioridades.FindByPrioridad(TextBox7.Text)
        ComboBox2.Text = _Busca.Item("DescripcionLarga")
        _Busca = CServiciosDataSet45.TipoRequisicion.FindByTipoRequisicion(TextBox8.Text)
        ComboBox3.Text = _Busca.Item("DescripcionLarga")

        TabControl2.Enabled = True

        ToolStripButton7.Enabled = False
        ToolStripButton8.Enabled = True
        RichTextBox1.Enabled = False
  
        ToolStripButton3.Enabled = True
        ToolStripButton2_Click(sender, e)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Frm54Requisiciones_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        ' Presione ESC
        If e.KeyChar = ChrW(Keys.Escape) Then
            If TextBox10.Text = "ALTA" Or TextBox10.Text = "CAMBIO" Then
                Button1_Click(sender, e)
            End If
            If TextBox10.Text = "MAT" Then
                ToolStripButton2_Click(sender, e)
            End If
            If TextBox10.Text = "MOV" Then
                DataGridView3.Rows.Clear()
                DataGridView3.Visible = False
                TabControl2.SelectedIndex = 0
                TabPage2.Text = ""
            End If
        End If
        TextBox10.Text = ""
    End Sub

    Private Sub Frm54Requisiciones_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet49.MovMat' Puede moverla o quitarla según sea necesario.
        Me.MovMatTableAdapter.Fill(Me.CServiciosDataSet49.MovMat)

        Button1_Click(sender, e)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet42.Prioridades' Puede moverla o quitarla según sea necesario.
        Me.PrioridadesTableAdapter.Fill(Me.CServiciosDataSet42.Prioridades)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet44.Requisiciones' Puede moverla o quitarla según sea necesario.
        Me.RequisicionesTableAdapter.Fill(Me.CServiciosDataSet44.Requisiciones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet44.DetalleRequisiciones' Puede moverla o quitarla según sea necesario.
        Me.DetalleRequisicionesTableAdapter.Fill(Me.CServiciosDataSet44.DetalleRequisiciones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet45.TipoRequisicion' Puede moverla o quitarla según sea necesario.
        Me.TipoRequisicionTableAdapter.Fill(Me.CServiciosDataSet45.TipoRequisicion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet5.Materiales' Puede moverla o quitarla según sea necesario.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)

        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        For Me.I = 0 To CServiciosDataSet44.Requisiciones.Rows.Count - 1
            _Requisicion = CServiciosDataSet44.Requisiciones.Rows(I).Item("Requisicion")
            _FechaRequisicion = CServiciosDataSet44.Requisiciones.Rows(I).Item("FechaRequisicion")
            _Solicitante = CServiciosDataSet44.Requisiciones.Rows(I).Item("Solicitante")
            _Centro = CServiciosDataSet44.Requisiciones.Rows(I).Item("Centro")
            _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
            _Centro = _Busca.Item("Descripcion")

            DataGridView1.Rows.Add(_Requisicion, _FechaRequisicion, _Solicitante, _Centro)

        Next

        ToolStripButton1.Enabled = True
        ToolStripButton2.Enabled = True
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        ToolStripButton5.Enabled = True

        TabControl2.Enabled = False


        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.Enabled = False
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
                _Checacontrol.Enabled = False
            End If
        Next
        RichTextBox1.Enabled = True
        ComboBox2.Items.Clear()
        For Me.I = 0 To CServiciosDataSet42.Prioridades.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet42.Prioridades.Rows(I).Item("DescripcionLarga"))
        Next

        ComboBox1.Items.Clear()
        For Me.I = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet.Centros.Rows(I).Item("Descripcion"))
        Next

        ComboBox3.Items.Clear()
        For Me.I = 0 To CServiciosDataSet45.TipoRequisicion.Rows.Count - 1
            ComboBox3.Items.Add(CServiciosDataSet45.TipoRequisicion.Rows(I).Item("DescripcionLarga"))
        Next

        ComboBox4.Items.Clear()
        For Me.I = 0 To CServiciosDataSet5.Materiales.Rows.Count - 1
            ComboBox4.Items.Add(CServiciosDataSet5.Materiales.Rows(I).Item("Descripcion"))
        Next
        ComboBox4.Sorted = True

        ToolStripButton7.Enabled = False
        ToolStripButton8.Enabled = False
        ToolStripButton6.Enabled = True
        ToolStripButton9.Enabled = True

        RichTextBox1.Text = ""
    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        'Nueva Requisición
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.Enabled = True
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Text = ""
                _Checacontrol.Enabled = True
            End If
        Next
        DataGridView2.Rows.Clear()
        TextBox10.Text = "ALTA"
        TextBox1.Focus()
        ToolStripButton6.Enabled = False
        ToolStripButton7.Enabled = True
        ToolStripButton8.Enabled = False
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus

        ' Obtiene el Consecutivo
        Dim _Find As Boolean
        Dim _Consecutivo As Integer
        Dim _XConsecutivo As String
        _Consecutivo = 0
        _XConsecutivo = ""
        _Find = False

        For Me.I = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(I).Item("Documento") = "REQ" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(I).Item("Consecutivo")
                _Find = True
            End If
        Next

        If _Find = False Then
            Msg = "No existe el registro de Consecutivo para las Requisiciones. Avise al Administrador del Sistema"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        _XConsecutivo = "RQ" & Trim(CStr(_Consecutivo))
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()

    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        If TextBox10.Text = "ALTA" Then
            TextBox2.Text = Today.Date
        End If

    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox2.LostFocus
        If Not IsDate(TextBox2.Text) Then
            Msg = "Debe teclear una fecha válida para esta Requisición"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            TextBox2.Focus()
        End If
        TextBox3.Focus()

    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = TimeOfDay
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub

    Private Sub RichTextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles RichTextBox1.GotFocus
        RichTextBox1.Text = "SIN COMENTARIOS"
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Me.I = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            If CServiciosDataSet.Centros.Rows(I).Item("Descripcion") = ComboBox1.Text Then
                TextBox4.Text = CServiciosDataSet.Centros.Rows(I).Item("Centro")
            End If
        Next
        TextBox4.Enabled = False
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        For Me.I = 0 To CServiciosDataSet42.Prioridades.Rows.Count - 1
            If CServiciosDataSet42.Prioridades.Rows(I).Item("DescripcionLarga") = ComboBox2.Text Then
                TextBox7.Text = CServiciosDataSet42.Prioridades.Rows(I).Item("Prioridad")
            End If
        Next
        TextBox7.Enabled = False
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        For Me.I = 0 To CServiciosDataSet45.TipoRequisicion.Rows.Count - 1
            If CServiciosDataSet45.TipoRequisicion.Rows(I).Item("DescripcionLarga") = ComboBox3.Text Then
                TextBox8.Text = CServiciosDataSet45.TipoRequisicion.Rows(I).Item("TipoRequisicion")
            End If
        Next
        TextBox8.Enabled = False
    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click

        ' Guarda Información

        'Valida los datos
        Dim _Find As Boolean

        _Find = False
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
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

        ' Valida el Centro

        _Find = False
        For Me.I = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            If CServiciosDataSet.Centros.Rows(I).Item("Centro") = TextBox4.Text Then
                _Find = True
            End If
        Next

        If _Find = False Then
            Msg = "El Centro al que se refiere esta Requisición no existe. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        ' Valida la Prioridad
        _Find = False

        For Me.I = 0 To CServiciosDataSet42.Prioridades.Rows.Count - 1
            If CServiciosDataSet42.Prioridades.Rows(I).Item("Prioridad") = TextBox7.Text Then
                _Find = True
            End If
        Next

        If _Find = False Then
            Msg = "La Prioridad de esta Requisición no está Registrada. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        ' Valida el Tipo de Requisición
        _Find = False

        For Me.I = 0 To CServiciosDataSet45.TipoRequisicion.Rows.Count - 1
            If CServiciosDataSet45.TipoRequisicion.Rows(I).Item("TipoRequisicion") = TextBox8.Text Then
                _Find = True
            End If
        Next

        If _Find = False Then
            Msg = "El tipo de Requisición no está Registrada. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        If TextBox10.Text = "ALTA" Then
            _Busca = CServiciosDataSet44.Requisiciones.NewRow
            _Busca.Item("Requisicion") = TextBox1.Text
        End If

        If TextBox10.Text = "CAMBIO" Then
            _Busca = CServiciosDataSet44.Requisiciones.FindByRequisicion(TextBox1.Text)
        End If

        _Busca.Item("FechaRequisicion") = CDate(TextBox2.Text)
        _Busca.Item("HoraRequisicion") = Mid(TextBox3.Text, 1, 5)
        _Busca.Item("Centro") = TextBox4.Text
        _Busca.Item("Referencia") = TextBox6.Text
        _Busca.Item("Solicitante") = TextBox5.Text
        _Busca.Item("Estatus") = "ACTIVA"
        _Busca.Item("TipoRequisicion") = TextBox8.Text
        _Busca.Item("Prioridad") = TextBox7.Text
        _Busca.Item("Comentarios") = UCase(RichTextBox1.Text)


        If TextBox10.Text = "ALTA" Then
            CServiciosDataSet44.Requisiciones.Rows.Add(_Busca)
        End If

        Me.Validate()
        RequisicionesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet44)


        If TextBox10.Text = "ALTA" Then
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("REQ")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager3.UpdateAll(CServiciosDataSet3)
        End If

        Msg = "Registro guardado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton9_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton9.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        Panel1.Visible = True
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = True
        TextBox10.Text = "MAT"
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        For Me.I = 0 To CServiciosDataSet5.Materiales.Rows.Count - 1
            If CServiciosDataSet5.Materiales.Rows(I).Item("Descripcion") = ComboBox4.Text Then
                TextBox9.Text = CServiciosDataSet5.Materiales.Rows(I).Item("Material")
                TextBox17.Text = CServiciosDataSet5.Materiales.Rows(I).Item("Umedida")
            End If
        Next
    End Sub

    Private Sub TextBox12_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox12.GotFocus
        TextBox12.Text = "N/A"
    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox13.GotFocus
        TextBox13.Text = "N/A"
    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox14_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox14.GotFocus
        TextBox14.Text = "N/A"
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox15_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox15.GotFocus
        TextBox15.Text = "N/A"
    End Sub

    Private Sub TextBox15_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox17_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox17.GotFocus
        TextBox17.Text = "PIEZAS"
    End Sub

    Private Sub TextBox17_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click

        ' Guarda los materiales
        If CInt(TextBox16.Text) = 0 Then
            Msg = "Debe teclear una cantidad Solicitada de este Material. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        If ComboBox4.Text = "" Then
            Msg = "Debe teclear un Material a Solicitar. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If TextBox12.Text = "" Then
            TextBox12.Text = "N/A"
        End If
        If TextBox13.Text = "" Then
            TextBox13.Text = "N/A"
        End If
        If TextBox14.Text = "" Then
            TextBox14.Text = "N/A"
        End If
        If TextBox15.Text = "" Then
            TextBox15.Text = "N/A"
        End If

        If TextBox17.Text = "" Then
            TextBox17.Text = "PIEZAS"
        End If

        _Busca = CServiciosDataSet44.DetalleRequisiciones.NewRow
        _Busca.Item("Requisicion") = TextBox1.Text
        _Busca.Item("Material") = TextBox9.Text
        _Busca.Item("Descripcion") = ComboBox4.Text
        _Busca.Item("CantidadSolicitada") = CInt(TextBox16.Text)
        _Busca.Item("Umedida") = TextBox17.Text
        _Busca.Item("Modelo") = TextBox13.Text
        _Busca.Item("Serie") = TextBox14.Text
        _Busca.Item("Color") = TextBox15.Text
        _Busca.Item("Referencia") = Date.Today
        _Busca.Item("ProveedorSugerido") = ""
        _Busca.Item("Marca") = TextBox12.Text
        _Busca.Item("CantidadSurte") = 0
        _Busca.Item("Documento") = ""
        _Busca.Item("ProveedorSurte") = ""
        _Busca.Item("FechaSurte") = DBNull.Value
        _Busca.Item("HoraSurte") = ""
        _Busca.Item("PrecioUnitario") = 0
        _Busca.Item("MarcaSurte") = ""
        _Busca.Item("ModeloSurte") = ""
        _Busca.Item("SerieSurte") = ""
        _Busca.Item("ColorSurte") = ""

        CServiciosDataSet44.DetalleRequisiciones.Rows.Add(_Busca)

        Me.Validate()
        DetalleRequisicionesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet44)

        Msg = "Registro guardado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)


        ToolStripButton2_Click(sender, e)


    End Sub

    Private Sub TextBox16_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox16.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox16_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox16.LostFocus
    
    End Sub

    Private Sub TextBox16_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Me.RequisicionesTableAdapter.Fill(Me.CServiciosDataSet44.Requisiciones)
        Panel1.Visible = False

        Dim _Rows_Dr() As DataRow = CServiciosDataSet44.DetalleRequisiciones.Select("Requisicion = " & "'" & TextBox1.Text & "'")
        Dim _CantidadSurte As Double
        Dim _Id As Integer

        DataGridView2.Rows.Clear()
        For Me.I = 0 To _Rows_Dr.GetUpperBound(0)
            _Material = _Rows_Dr(I).Item("Material")
            _Descripcion = _Rows_Dr(I).Item("Descripcion")
            _Marca = _Rows_Dr(I).Item("Marca")
            _Modelo = _Rows_Dr(I).Item("Modelo")
            _Serie = _Rows_Dr(I).Item("Serie")
            _Color = _Rows_Dr(I).Item("Color")
            _Cantidad = _Rows_Dr(I).Item("CantidadSolicitada")
            _UMedida = _Rows_Dr(I).Item("Umedida")
            _CantidadSurte = _Rows_Dr(I).Item("CantidadSurte")
            _Id = _Rows_Dr(I).Item("Orden")

            DataGridView2.Rows.Add(_Material, _Descripcion, _Marca, _Modelo, _Serie, _Color, _Cantidad, _UMedida, _CantidadSurte, _Id)
        Next
        ToolStripButton3.Enabled = True
        ToolStripButton4.Enabled = False

        For Me.I = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(I).Cells(0).Value = "" Then
                Exit For
            End If
            If DataGridView2.Rows(I).Cells(8).Value >= DataGridView2.Rows(I).Cells(6).Value Then
                DataGridView2.Rows(I).DefaultCellStyle.BackColor = Color.PeachPuff
            End If
        Next

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        ToolStripButton2_Click(sender, e)
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold



    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub EliminarEsteRegistroToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EliminarEsteRegistroToolStripMenuItem.Click
        ' Eliminar registro

        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Msg = "Debe seleccionar un Material para poder Eliminarlo"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        If _Renglon.Cells(8).Value <> 0 Then

        End If

        If _Renglon.Cells(8).Value <> 0 Then
            Msg = "Este material de esta Requisición ya se ha surtido. No. Puede Eliminarlo"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If






        Dim _Id As Integer = _Renglon.Cells(9).Value
        Dim title As String

        Dim response As MsgBoxResult
        Msg = "Está seguro de eliminar este Registro?"
        Style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Elininar Materiales de esta Requisición"

        response = MsgBox(Msg, Style, title)
        If response = MsgBoxResult.Yes Then
            ' Baja
            _Busca = CServiciosDataSet44.DetalleRequisiciones.FindByOrden(_Id)
            _Busca.Delete()
            Me.Validate()
            DetalleRequisicionesBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet44)

            Msg = "Registro Eliminado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)

            ToolStripButton2_Click(sender, e)
        Else
            ' Perform some other action.
        End If

    End Sub

    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click
        TextBox10.Text = "CAMBIO"
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = True
            End If
            If TypeOf _Checacontrol Is ComboBox Then
                _Checacontrol.Enabled = True
            End If
        Next
        RichTextBox1.Enabled = True
        ToolStripButton8.Enabled = False
        ToolStripButton7.Enabled = True
    End Sub

    Private Sub ToolStripButton10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton10.Click
        ' Imprime la Requisición

        If TextBox1.Text = "" Then
            Msg = "Debe seleccionar una Requisición para Impresión"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        Dim _Contador As Integer
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        ' Dim strRutaExcel As String = "Z:\Requisiciones de Material.xlsx"
        Dim strRutaExcel As String = "Z:\Formatos Factor\Requisicion Factor.xlsx"
        Dim _Linea As Integer
        Dim _Meses(12) As String
        _Meses(0) = "Enero"
        _Meses(1) = "Febrero"
        _Meses(2) = "Marzo"
        _Meses(3) = "Abril"
        _Meses(4) = "Mayo"
        _Meses(5) = "Junio"
        _Meses(6) = "Julio"
        _Meses(7) = "Agosto"
        _Meses(8) = "Septiembre"
        _Meses(9) = "Octubre"
        _Meses(10) = "Noviembre"
        _Meses(11) = "Diciembre"

        Dim _Fecha As String = Mid(TextBox2.Text, 1, 2)
        _Fecha = _Fecha & " de " & _Meses(Month(CDate(TextBox2.Text)) - 1)
        _Fecha = _Fecha & " de " & Mid(TextBox2.Text, 7, 4)

        _Linea = 14
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de exce
        '  Dim _Color As Color = m_Excel.Worksheets("REQ").Range("A18").interior.Color
        Dim _Hoja As Object = m_Excel.Worksheets("Formato")

        For Me.I = 14 To 35
            m_Excel.Worksheets("Formato").cells(I, 2) = ""
            m_Excel.Worksheets("Formato").cells(I, 3) = ""
            m_Excel.Worksheets("Formato").cells(I, 4) = ""
            m_Excel.Worksheets("Formato").cells(I, 5) = ""
        Next


        '    _Hoja.Range(18, 1).copy()
        m_Excel.Worksheets("Formato").cells(4, 2) = ""
        m_Excel.Worksheets("Formato").cells(5, 2) = ""
        '   m_Excel.Worksheets("Formato").cells(5, 10) = _Fecha
        m_Excel.Worksheets("Formato").cells(5, 10) = TextBox2.Text
        m_Excel.Worksheets("Formato").cells(3, 10) = "Folio " & " " & TextBox1.Text

        _Contador = 1
        For Me.I = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(I).Cells(0).Value = "" Then
                Exit For
            End If

            '   m_Excel.Worksheets("REQ").cells(_Linea, 1) = DataGridView2.Rows(I).Cells(0).Value
            _Material = DataGridView2.Rows(I).Cells(0).Value
            _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_Material)
            _Marca = DataGridView2.Rows(I).Cells(2).Value
            _UMedida = _Busca.Item("UMedida")


            If _Marca = "N/A" Then
                _Marca = ""
            Else
                _Marca = "MARCA " & _Marca
            End If

            ' _Descripcion = _Busca.Item("Presentacion") & " " & _Busca.Item("Descripcion") & " " & _Marca
            _Descripcion = _Busca.Item("Descripcion") & " " & _Marca

            ' m_Excel.Worksheets("Formato").Cells(_Linea, 2) = _Contador
            m_Excel.Worksheets("Formato").cells(_Linea, 5) = _Descripcion
            m_Excel.Worksheets("Formato").cells(_Linea, 3) = DataGridView2.Rows(I).Cells(6).Value
            m_Excel.Worksheets("Formato").cells(_Linea, 4) = _UMedida



            '      If (_Contador Mod 2) > 0 Then
            '      _Hoja.Range("B20").Select()
            '      '   _Hoja.Range("B20").PasteSpecial(Paste:=xlFormats)
            '      With (_Hoja.cells(_Linea, 1))
            '      .interior.colorindex = 24
            '      End With
            '      With _Hoja.cells(_Linea, 2)
            '      .interior.colorindex = 24
            '      End With
            '      With _Hoja.cells(_Linea, 3)
            '      .interior.colorindex = 24
            '      End With
            '      With _Hoja.cells(_Linea, 4)
            '      .interior.colorindex = 24
            '      End With
            '
            '            End If
            _Contador = _Contador + 1
            _Linea = _Linea + 1

        Next

        m_Excel.Worksheets("Formato").cells(43, 2) = TextBox5.Text
        m_Excel.Worksheets("Formato").cells(44, 2) = ComboBox1.Text

        ' Imprime los Comentarios

        _Comentarios = RichTextBox1.Text
        Dim _Mide As Integer = Len(_Comentarios)
        Dim X As Integer
        Dim _Texto As String

        _Linea = 35
        If _Mide > 0 Then
            For Me.I = 1 To _Mide Step 100
                _Texto = Mid(_Comentarios, I, 100)
                '              m_Excel.Worksheets("REQ").cells(_Linea, 2).value = _Texto
                _Linea = _Linea + 1
            Next


        End If

        m_Excel.Visible = True
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub RecepciónDeEsteMaterialToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RecepciónDeEsteMaterialToolStripMenuItem.Click
        ' Presenta los Movimientos de este Material
        TabControl2.SelectedIndex = 1
        TabPage2.Text = "Movimientos del Material " & _Renglon.Cells(1).Value
        DataGridView3.Visible = True
        DataGridView3.Rows.Clear()

        TextBox10.Text = "MOV"
        _Material = _Renglon.Cells(0).Value

        Dim _Cantidad, _PrecioUnitario As Double
        Dim _FechaMov As Date
        Dim _TipoMov, _Umedida, _Proveedor, _Factura, _Referencia As String
        Dim _RowsMM() As DataRow = CServiciosDataSet49.MovMat.Select("Material = " & "'" & _Material & "'")

        For Me.I = 0 To _RowsMM.GetUpperBound(0)
            _TipoMov = _RowsMM(I).Item("TipoMov")
            If _TipoMov = "E" Then
                _TipoMov = "ENTRADA"
            Else
                _TipoMov = "SALIDA"
            End If
            _FechaMov = CDate(_RowsMM(I).Item("FechaMov"))
            _Cantidad = _RowsMM(I).Item("Cantidad")
            _PrecioUnitario = _RowsMM(I).Item("PrecioUnitario")
            _Umedida = _RowsMM(I).Item("UMedida")
            _Proveedor = _RowsMM(I).Item("Proveedor")
            _Factura = _RowsMM(I).Item("Factura")
            _Referencia = _RowsMM(I).Item("Referencia")

            DataGridView3.Rows.Add(_TipoMov, _FechaMov, _Cantidad, _PrecioUnitario, _Umedida, _Proveedor, _Factura, _Referencia)

            '    Label7.Text = "Precione ESC para regresar al listado de materiales"
            '    Label7.ForeColor = Color.Maroon
        Next


    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub
End Class