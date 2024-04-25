Public Class Frm55Recepcion
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle
    Public _Busca As DataRow
    Public _Renglon As DataGridViewRow
    Public _Checacontrol As Control
    Public _Recepcion, _Descripcion, _Referencia, _Solicitante, _Recibe, _Documento, _Centro, _Requisicion As String
    Public _MarcaRecibe, SerieRecibe, _ModeloRecibe, _ColorRecibe, _Hora, _Comentarios As String
    Public _Fecha As Date
    Public _Material, _Marca, _Modelo, _Serie, _Color As String
    Public _CantidadRecibe, _CantidadSolicitada, _CantidadSurte As Double

    Private Sub Frm55Recepcion_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Escape) Then
            _Renglon = DataGridView2.CurrentRow
            _Renglon.DefaultCellStyle.BackColor = Color.White
            For Each Me._Checacontrol In Panel1.Controls
                If TypeOf _Checacontrol Is TextBox Then
                    _Checacontrol.Text = ""
                End If
            Next
            Label19.Text = "."

            Panel1.Visible = False
        End If
    End Sub


    Private Sub Form55_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
   
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Panel1.Visible = True
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet46.Recepciones' Puede moverla o quitarla según sea necesario.
        Me.RecepcionesTableAdapter.Fill(Me.CServiciosDataSet46.Recepciones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet46.DetalleRecepciones' Puede moverla o quitarla según sea necesario.
        Me.DetalleRecepcionesTableAdapter.Fill(Me.CServiciosDataSet46.DetalleRecepciones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet.Centros' Puede moverla o quitarla según sea necesario.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet44.Requisiciones' Puede moverla o quitarla según sea necesario.
        Me.RequisicionesTableAdapter.Fill(Me.CServiciosDataSet44.Requisiciones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet44.DetalleRequisiciones' Puede moverla o quitarla según sea necesario.
        Me.DetalleRequisicionesTableAdapter.Fill(Me.CServiciosDataSet44.DetalleRequisiciones)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet34.Proveedores' Puede moverla o quitarla según sea necesario.
        Me.ProveedoresTableAdapter.Fill(Me.CServiciosDataSet34.Proveedores)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet5.Materiales' Puede moverla o quitarla según sea necesario.
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet49.MovMat' Puede moverla o quitarla según sea necesario.
        Me.MovMatTableAdapter.Fill(Me.CServiciosDataSet49.MovMat)

        TabControl2.Enabled = False
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.Enabled = False
            End If
        Next
        For Each Me._Checacontrol In Panel1.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        ComboBox1.Text = ""
        RichTextBox1.Text = ""
        ComboBox1.Enabled = False
        RichTextBox1.Enabled = False

        ToolStripButton3.Enabled = False
        ToolStripButton5.Enabled = False
        ToolStripButton6.Enabled = False
        ToolStripButton4.Enabled = True

        ComboBox1.Items.Clear()
        For Me.I = 0 To CServiciosDataSet.Centros.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet.Centros.Rows(I).Item("Descripcion"))
        Next

        For Me.I = 0 To CServiciosDataSet46.Recepciones.Rows.Count - 1
            _Recepcion = CServiciosDataSet46.Recepciones.Rows(I).Item("Recepcion")
            _Fecha = CServiciosDataSet46.Recepciones.Rows(I).Item("Fecha")
            _Referencia = CServiciosDataSet46.Recepciones.Rows(I).Item("Referencia")

            DataGridView1.Rows.Add(_Recepcion, _Fecha, _Referencia)
        Next

        ComboBox2.Items.Clear()

        For Me.I = 0 To CServiciosDataSet34.Proveedores.Rows.Count - 1
            ComboBox2.Items.Add(CServiciosDataSet34.Proveedores.Rows(I).Item("RazonSocial"))
        Next
        Panel1.Visible = False

        TabControl2.SelectedIndex = 1
        Timer1.Stop()
        TabPage4.Parent = Nothing

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
                _Checacontrol.Enabled = True
            End If
        Next

        ToolStripButton4.Enabled = False
        ToolStripButton5.Enabled = True

        ComboBox1.Enabled = True
        RichTextBox1.Enabled = True
        ComboBox1.Text = ""
        RichTextBox1.Text = ""
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        TextBox14.Text = "ALTA"

        Dim _Consecutivo As Integer
        Dim _XConsecutivo As String
        Dim _Find As Boolean

        ' Obtiene el Consecutivo
        _Find = False
        For Me.I = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(I).Item("Documento") = "REM" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(I).Item("Consecutivo")
                _Find = True
            End If
        Next

        If Not _Find Then
            Msg = "No existe el Consecutivo para las recepciones de Material. Contacte al Administrador del Sistema"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        _XConsecutivo = "RM" & Trim(CStr(_Consecutivo))
        TextBox1.Text = _XConsecutivo
        TextBox1.Enabled = False
        TextBox2.Focus()

        Timer1.Start()



    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox2.Text = Date.Today
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = TimeOfDay
        TextBox3.Text = Mid(TextBox3.Text, 1, 5)

    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim _Registro As Integer = ComboBox1.SelectedIndex
        _Centro = CServiciosDataSet.Centros.Rows(_Registro).Item("Centro")
        TextBox8.Text = _Centro
        TextBox8.Enabled = False
        TextBox4.Focus()



    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        _Renglon = DataGridView3.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.LightGray
        _Requisicion = _Renglon.Cells(0).Value
        _Fecha = _Renglon.Cells(1).Value
        Dim _Rows_Dr() As DataRow = CServiciosDataSet44.DetalleRequisiciones.Select("Requisicion = " & "'" & _Requisicion & "'")
        Dim _Id As Integer
        Dim _Umedida As String

        TabControl2.SelectedIndex = 0

        DataGridView2.Rows.Clear()
        For Me.I = 0 To _Rows_Dr.GetUpperBound(0)
            If _Rows_Dr(I).Item("CantidadSurte") < _Rows_Dr(I).Item("CantidadSolicitada") Then
                _Material = _Rows_Dr(I).Item("Material")
                _Descripcion = _Rows_Dr(I).Item("Descripcion")
                _Marca = _Rows_Dr(I).Item("Marca")
                _Modelo = _Rows_Dr(I).Item("Modelo")
                _Serie = _Rows_Dr(I).Item("Serie")
                _Color = _Rows_Dr(I).Item("Color")

                _Descripcion = _Descripcion & " " & _Marca & " " & _Modelo & " " & _Serie & " " & _Color
                _CantidadSolicitada = _Rows_Dr(I).Item("CantidadSolicitada")
                _CantidadSurte = _Rows_Dr(I).Item("CantidadSurte")
                _Id = _Rows_Dr(I).Item("Orden")
                _Umedida = _Rows_Dr(I).Item("Umedida")

                DataGridView2.Rows.Add(_Requisicion, _Material, _Descripcion, _CantidadSolicitada, _CantidadSurte, _Marca, _Modelo, _Serie, _Color, False, _Id, _Umedida)
            End If

        Next
        If DataGridView2.Rows(0).Cells(0).Value = "" Then
            Msg = "Esta Requiisición Not tiene materiales pendientes de Surtir"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


    End Sub

    Private Sub DataGridView3_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        _Renglon = DataGridView2.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.PeachPuff
        Panel1.Visible = True
        ToolStripButton3.Enabled = True

        ' Asgina los valores solicitados
        Dim _id As Integer = _Renglon.Cells(10).Value
        _Busca = CServiciosDataSet44.DetalleRequisiciones.FindByOrden(_id)
        TextBox9.Text = _Busca.Item("CantidadSolicitada")
        TextBox10.Text = _Busca.Item("Marca")
        TextBox11.Text = _Busca.Item("Modelo")
        TextBox12.Text = _Busca.Item("Serie")
        TextBox13.Text = _Busca.Item("Color")
        TextBox9.Focus()

        Label19.Text = "Presione ESC para regresar al listado de materiales incluídos en esta Requisición"
        Label19.ForeColor = Color.Maroon


    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click

        Dim title As String
        Dim _Estatus As String
        Dim response As MsgBoxResult
        Msg = "Está seguro de Guardar esta Información?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Recepción de Materiales"

        _Recepcion = TextBox1.Text
        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Alta de Registro
            If TextBox14.Text = "ALTA" Then
                _Busca = CServiciosDataSet46.Recepciones.NewRow
                _Busca.Item("Recepcion") = _Recepcion
            End If
            If TextBox14.Text = "CAMBIO" Then
                _Busca = CServiciosDataSet44.Requisiciones.FindByRequisicion(_Requisicion)
            End If

            _Fecha = CDate(TextBox2.Text)
            _Hora = TextBox3.Text
            _Referencia = TextBox6.Text
            _Documento = TextBox5.Text
            _Recibe = TextBox7.Text
            _Centro = TextBox8.Text
            _Comentarios = UCase(RichTextBox1.Text)
            _Estatus = TextBox4.Text

            _Busca.Item("Fecha") = _Fecha
            _Busca.Item("Hora") = _Hora
            _Busca.Item("Referencia") = _Referencia
            _Busca.Item("Documento") = _Documento
            _Busca.Item("Recibe") = _Recibe
            _Busca.Item("Centro") = _Centro
            _Busca.Item("Comentarios") = _Comentarios
            _Busca.Item("Estatus") = _Estatus


            If TextBox14.Text = "ALTA" Then
                CServiciosDataSet46.Recepciones.Rows.Add(_Busca)

                _Busca = CServiciosDataSet3.Documentos.FindByDocumento("REM")
                _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1
            End If

            ' Registra la Recepción de Mateirales
            Me.Validate()
            RecepcionesBindingSource.EndEdit()
            Me.TableAdapterManager2.UpdateAll(CServiciosDataSet46)
            Me.RecepcionesTableAdapter.Fill(Me.CServiciosDataSet46.Recepciones)

            ' Registra el Consecutivo
            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager1.UpdateAll(CServiciosDataSet3)
            Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)


            Msg = "Registro guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)

            Timer1.Stop()
            Button1_Click(sender, e)


        Else
            ' Perform some other action.
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If

        _Renglon.DefaultCellStyle.BackColor = Color.Gold
        _Recepcion = _Renglon.Cells(0).Value

        TextBox1.Text = _Recepcion

        _Busca = CServiciosDataSet46.Recepciones.FindByRecepcion(_Recepcion)
        TextBox2.Text = _Busca.Item("Fecha")
        TextBox3.Text = _Busca.Item("Hora")
        TextBox8.Text = _Busca.Item("Centro")
        TextBox4.Text = _Busca.Item("Estatus")
        TextBox5.Text = _Busca.Item("Documento")
        TextBox6.Text = _Busca.Item("Referencia")
        TextBox7.Text = _Busca.Item("Recibe")
        RichTextBox1.Text = _Busca.Item("Comentarios")
        ToolStripButton6.Enabled = True

        _Centro = _Busca.Item("Centro")
        _Busca = CServiciosDataSet.Centros.FindByCentro(_Centro)
        ComboBox1.Text = _Busca.Item("Descripcion")

        ' Muestra las requisiciones de ese centro
        Dim _Rows_Re() As DataRow = CServiciosDataSet44.Requisiciones.Select("Centro = " & "'" & _Centro & "'")
        Dim _Umedida As String
        Dim _PrecioUnitario As Double

        DataGridView2.Rows.Clear()
        For Me.I = 0 To _Rows_Re.GetUpperBound(0)
            _Requisicion = _Rows_Re(I).Item("Requisicion")
            _Fecha = _Rows_Re(I).Item("FechaRequisicion")
            _Referencia = _Rows_Re(I).Item("Referencia")
            _Solicitante = _Rows_Re(I).Item("Solicitante")

            DataGridView3.Rows.Add(_Requisicion, _Fecha, _Referencia, _Solicitante)
        Next
        DataGridView3.Enabled = False


        ' Presenta el detalle de la Recepción
        DataGridView4.Rows.Clear()
        _Rows_Re = CServiciosDataSet46.DetalleRecepciones.Select("Recepcion = " & "'" & _Recepcion & "'")

        For Me.I = 0 To _Rows_Re.GetUpperBound(0)
            _Material = _Rows_Re(I).Item("MaterialRecibe")
            _CantidadRecibe = _Rows_Re(I).Item("CantidadRecibe")
            _Descripcion = _Rows_Re(I).Item("Descripcion")
            _MarcaRecibe = _Rows_Re(I).Item("MarcaRecibe")
            _ModeloRecibe = _Rows_Re(I).Item("ModeloRecibe")
            _Serie = _Rows_Re(I).Item("SerieRecibe")
            _ColorRecibe = _Rows_Re(I).Item("ColorRecibe")
            _Umedida = _Rows_Re(I).Item("Umedida")
            _Requisicion = _Rows_Re(I).Item("Requisicion")
            _PrecioUnitario = _Rows_Re(I).Item("PrecioUnitario")

            DataGridView4.Rows.Add(_Recepcion, _Material, _CantidadRecibe, _Descripcion, _MarcaRecibe, _ModeloRecibe, _Serie, _ColorRecibe, _Umedida, _Requisicion, _PrecioUnitario)

        Next




    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        TextBox14.Text = "CAMBIO"
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = True
            End If
        Next
        DataGridView3.Enabled = True
        ComboBox1.Enabled = True
        RichTextBox1.Enabled = True
        TabControl2.Enabled = True
        TextBox8.Enabled = False


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                If _Checacontrol.Focused Then
                    _Checacontrol.BackColor = Color.Gold
                Else
                    _Checacontrol.BackColor = Color.White
                End If
            End If
        Next
    End Sub

    Private Sub TextBox10_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox10.GotFocus

    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox11_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox11.GotFocus

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox12_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox12.GotFocus

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox12_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub TextBox13_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox13.GotFocus

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub PorProveedorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PorProveedorToolStripMenuItem.Click

        ComboBox3.Items.Clear()
        TabPage1.Parent = Nothing
        TabPage4.Parent = TabControl1
        TextBox18.Text = ""
        ComboBox3.Text = ""

        TextBox20.Text = "P"
        Label20.Text = "Seleccione el Proveedor"

        For Me.I = 0 To CServiciosDataSet34.Proveedores.Rows.Count - 1
            ComboBox3.Items.Add(CServiciosDataSet34.Proveedores.Rows(I).Item("RazonSocial"))
        Next

        Button3.Focus()



    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        I = ComboBox3.SelectedIndex
        Dim _Proveedor = CServiciosDataSet34.Proveedores.Rows(I).Item("Proveedor")
        TextBox18.Text = _Proveedor

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If TextBox18.Text = "" Then
            Msg = "Debe seleccionar un valor para generar el listado"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If


        Dim _Rows_Dr() As DataRow

        If TextBox20.Text = "P" Then
            Dim _Proveedor = TextBox18.Text
            _Rows_Dr = CServiciosDataSet46.DetalleRecepciones.Select("ProveedorRecibe = " & "'" & _Proveedor & "'")
        End If
        If TextBox20.Text = "M" Then
            _Material = TextBox18.Text
            _Rows_Dr = CServiciosDataSet46.DetalleRecepciones.Select("Material = " & "'" & _Material & "'")
        End If



    End Sub

    Private Sub PorMaterialToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PorMaterialToolStripMenuItem.Click
        TextBox20.Text = "M"
        ComboBox3.Text = ""
        TextBox18.Text = ""

        Label20.Text = "Seleccione al Material a Reportear"
        ComboBox3.Items.Clear()

        For Me.I = 0 To CServiciosDataSet5.Materiales.Rows.Count - 1
            ComboBox3.Items.Add(CServiciosDataSet5.Materiales.Rows(I).Item("Descripcion"))
        Next


    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        Dim _Id As Integer = _Renglon.Cells(10).Value

        If TextBox9.Text = "" Then
            Msg = "Debe teclear una Cantidad Recibida"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        '   _Busca = CServiciosDataSet46.DetalleRecepciones.FindByOrden(_Id)
        _Busca = CServiciosDataSet46.DetalleRecepciones.NewRow
        _Busca.Item("Recepcion") = TextBox1.Text
        _Busca.Item("MaterialRecibe") = _Renglon.Cells(1).Value
        _Busca.Item("Descripcion") = _Renglon.Cells(2).Value
        _Busca.Item("CantidadRecibe") = CInt(TextBox9.Text)
        _Busca.Item("MarcaRecibe") = TextBox10.Text
        _Busca.Item("ModeloRecibe") = TextBox11.Text
        _Busca.Item("SerieRecibe") = TextBox12.Text
        _Busca.Item("ColorRecibe") = TextBox13.Text
        _Busca.Item("ProveedorRecibe") = TextBox16.Text
        _Busca.Item("PrecioUnitario") = CInt(TextBox15.Text)
        _Busca.Item("Requisicion") = _Renglon.Cells(0).Value
        _Busca.Item("UMedida") = _Renglon.Cells(11).Value


        CServiciosDataSet46.DetalleRecepciones.Rows.Add(_Busca)

        Me.Validate()
        DetalleRecepcionesBindingSource.EndEdit()
        Me.TableAdapterManager2.UpdateAll(CServiciosDataSet46)
        Me.DetalleRecepcionesTableAdapter.Fill(Me.CServiciosDataSet46.DetalleRecepciones)

        ' Actualiza el Detalle de la Requisición


        _Requisicion = _Renglon.Cells(0).Value
        _Busca = CServiciosDataSet44.DetalleRequisiciones.FindByOrden(_Id)
        _Material = _Busca.Item("Material")
        _Busca.Item("CantidadSurte") = _Busca.Item("CantidadSurte") + CInt(TextBox9.Text)
        _Busca.Item("MarcaSurte") = TextBox10.Text
        _Busca.Item("ModeloSurte") = TextBox11.Text
        _Busca.Item("SerieSurte") = TextBox12.Text
        _Busca.Item("ColorSurte") = TextBox13.Text
        _Busca.Item("PrecioUnitario") = CInt(TextBox15.Text)
        _Busca.Item("FechaSurte") = CDate(TextBox2.Text)
        _Busca.Item("HoraSurte") = Mid(TimeOfDay, 1, 5)

        Me.Validate()
        DetalleRequisicionesBindingSource.EndEdit()
        Me.TableAdapterManager3.UpdateAll(CServiciosDataSet44)
        Me.DetalleRequisicionesTableAdapter.Fill(Me.CServiciosDataSet44.DetalleRequisiciones)

        Dim _Umedida As String
        ' Actualiza el Inventario
        _Busca = CServiciosDataSet5.Materiales.FindByMaterial(_Material)
        Dim _ImporteActual, _ImporteNuevo, _Suma, _Existencia, _PrecioUnitario As Integer
        _Existencia = _Busca.Item("Existencia") + CInt(TextBox9.Text)
        _ImporteActual = _Busca.Item("Existencia") * _Busca.Item("PrecioUnitario")
        _Busca.Item("Existencia") = _Busca.Item("Existencia") + CInt(TextBox9.Text)
        _ImporteNuevo = CInt(TextBox9.Text) * CInt(TextBox15.Text)
        _Suma = _ImporteActual + _ImporteNuevo
        _PrecioUnitario = _Suma / _Existencia
        _Busca.Item("PrecioUnitario") = _PrecioUnitario
        _Umedida = _Busca.Item("UMedida")

        Me.Validate()
        MaterialesBindingSource.EndEdit()
        Me.TableAdapterManager5.UpdateAll(CServiciosDataSet5)
        Me.MaterialesTableAdapter.Fill(Me.CServiciosDataSet5.Materiales)


        ' Registra los movimientos de los materiales
        Dim _Usuario As String = LoginForm1.TextBox6.Text
        _Centro = LoginForm1.TextBox1.Text

        _Busca = CServiciosDataSet49.MovMat.NewRow
        _Busca.Item("Material") = _Material
        _Busca.Item("TipoMov") = "E"
        _Busca.Item("FechaMov") = CStr(Today.Date)
        _Busca.Item("Horamov") = Mid(TimeOfDay, 1, 10)
        _Busca.Item("Cantidad") = CDbl(TextBox9.Text)
        _Busca.Item("PrecioUnitario") = CDbl(TextBox15.Text)
        _Busca.Item("Proveedor") = TextBox16.Text
        _Busca.Item("Documento") = TextBox1.Text
        _Busca.Item("Factura") = TextBox17.Text
        _Busca.Item("Referencia") = TextBox6.Text
        _Busca.Item("Comentarios") = "RECEPCION DE MATERIALES DE PROVEEDOR"
        _Busca.Item("Usuario") = _Usuario
        _Busca.Item("Centro") = _Centro
        _Busca.Item("UMedida") = _UMedida

        CServiciosDataSet49.MovMat.Rows.Add(_Busca)

        Me.Validate()
        MovMatBindingSource.EndEdit()
        Me.TableAdapterManager6.UpdateAll(CServiciosDataSet49)
        Me.MovMatTableAdapter.Fill(Me.CServiciosDataSet49.MovMat)

        Msg = "Registro Guardado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)
        Button1_Click(sender, e)


    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
     

        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If


    End Sub

    Private Sub TextBox9_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox15_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox15.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim _Indice As Integer = ComboBox2.SelectedIndex
        Dim _Proveedor As String = CServiciosDataSet34.Proveedores.Rows(_Indice).Item("Proveedor")
        TextBox16.Text = _Proveedor
        TextBox16.Enabled = False

    End Sub

    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub ToolStripButton10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton10.Click

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 13 Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If Asc(e.KeyChar) = 13 Then
            RichTextBox1.Focus()
        End If
    End Sub

    Private Sub RichTextBox1_LostFocus(sender As Object, e As EventArgs) Handles RichTextBox1.LostFocus
        RichTextBox1.Text = UCase(RichTextBox1.Text)
    End Sub
End Class