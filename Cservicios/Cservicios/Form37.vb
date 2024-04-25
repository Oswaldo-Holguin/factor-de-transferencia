Public Class Frm37InformeMensual
    Public i As Integer
    Public Msg As String
    Public Style As MsgBoxStyle

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub Frm37InformeMensual_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet29.CierreProduccion' Puede moverla o quitarla según sea necesario.
        Me.CierreProduccionTableAdapter.Fill(Me.CServiciosDataSet29.CierreProduccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet26.Produccion' Puede moverla o quitarla según sea necesario.
        Me.ProduccionTableAdapter.Fill(Me.CServiciosDataSet26.Produccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet32.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet32.VentaFactor)

        ToolStripComboBox1.Items.Add("ENERO")
        ToolStripComboBox1.Items.Add("FEBRERO")
        ToolStripComboBox1.Items.Add("MARZO")
        ToolStripComboBox1.Items.Add("ABRIL")
        ToolStripComboBox1.Items.Add("MAYO")
        ToolStripComboBox1.Items.Add("JUNIO")
        ToolStripComboBox1.Items.Add("JULIO")
        ToolStripComboBox1.Items.Add("AGOSTO")
        ToolStripComboBox1.Items.Add("SEPTIEMBRE")
        ToolStripComboBox1.Items.Add("OCTUBRE")
        ToolStripComboBox1.Items.Add("NOVIEMBRE")
        ToolStripComboBox1.Items.Add("DICIEMBRE")


        Dim _Año As Integer = CInt(Year(Today))

        For Me.i = _Año To (_Año - 10) Step -1
            ToolStripComboBox2.Items.Add(i)
        Next

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        If ToolStripComboBox1.Text = "Seleccione el Mes" Or ToolStripComboBox1.Text = "" Then
            Msg = "Debe seleccionar un Mes para generar este reporte"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        If ToolStripComboBox2.Text = "Seleccione el Año" Or ToolStripComboBox2.Text = "" Then
            Msg = "Debe seleccionar un Año para generar este reporte"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        Dim _Rows_VF() As DataRow
        Dim _Mes As Integer = ToolStripComboBox1.SelectedIndex + 1
        Dim _Año As Integer = ToolStripComboBox2.Text
        Dim _Busca As DataRow
        Dim _OrdenProduccion As String
        Dim _FechaInicioProduccion As Date
        Dim _Cantidad, _PrecioUnitario, _ImporteProduccion, _CostoMateriales, _CostoManodeObra, _CostosIndirectos As Integer
        Dim _Neto, _Venta, _Utilidad As Integer
        Dim X As Integer
        Dim _Linea As Integer
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Informe.xlsx"
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de excel

        For Me.i = 10 To 1000
            If m_Excel.Worksheets("MES").cells(2).value = "" Then
                Exit For
            End If
            m_Excel.Worksheets("MES").Cells(2).value = ""
            m_Excel.Worksheets("MES").Cells(3).value = ""
            m_Excel.Worksheets("MES").Cells(4).value = ""
            m_Excel.Worksheets("MES").Cells(5).value = ""
            m_Excel.Worksheets("MES").Cells(6).value = ""
            m_Excel.Worksheets("MES").Cells(7).value = ""
            m_Excel.Worksheets("MES").Cells(8).value = ""
            m_Excel.Worksheets("MES").Cells(9).value = ""
            m_Excel.Worksheets("MES").Cells(10).value = ""
            m_Excel.Worksheets("MES").Cells(11).value = ""
            m_Excel.Worksheets("MES").Cells(12).value = ""
            m_Excel.Worksheets("MES").Cells(13).value = ""
        Next

        _Linea = 10
        For Me.i = 0 To CServiciosDataSet26.Produccion.Rows.Count - 1
            If CServiciosDataSet26.Produccion.Rows(i).Item("Comentarios") = "ENTREGADO" Then
                If Year(CServiciosDataSet26.Produccion.Rows(i).Item("FechaInicioProduccion")) = _Año Then
                    If Month(CServiciosDataSet26.Produccion.Rows(i).Item("FechaInicioProduccion")) = _Mes Then
                        _OrdenProduccion = CServiciosDataSet26.Produccion.Rows(i).Item("OrdenProduccion")
                        _FechaInicioProduccion = CServiciosDataSet26.Produccion.Rows(i).Item("FechaInicioProduccion")
                        _Cantidad = CServiciosDataSet26.Produccion.Rows(i).Item("Cantidad")

                        _Busca = CServiciosDataSet29.CierreProduccion.FindByOrdenProduccion(_OrdenProduccion)
                        _PrecioUnitario = _Busca.Item("PrecioUnitario")
                        _ImporteProduccion = _Busca.Item("ImporteProduccion")
                        _CostoMateriales = _Busca.Item("CostoMateriales")
                        _CostoManodeObra = _Busca.Item("CostoManodeObra")
                        _CostosIndirectos = _Busca.Item("CostoIndirectos")
                        _Neto = _ImporteProduccion - _CostoMateriales - _CostoManodeObra - _CostosIndirectos

                        _Rows_VF = CServiciosDataSet32.VentaFactor.Select("Referencia = " & "'" & _OrdenProduccion & "'")

                        _Venta = 0
                        For X = 0 To _Rows_VF.GetUpperBound(0)
                            _Venta = _Venta + _Rows_VF(X).Item("PrecioNeto")
                        Next

                        m_Excel.Worksheets("MES").cells(_Linea, 2).value = _OrdenProduccion
                        m_Excel.Worksheets("MES").Cells(_Linea, 3).value = _FechaInicioProduccion
                        m_Excel.Worksheets("MES").Cells(_Linea, 4).value = _Cantidad
                        m_Excel.Worksheets("MES").Cells(_Linea, 5).value = _PrecioUnitario
                        m_Excel.Worksheets("MES").Cells(_Linea, 6).value = _ImporteProduccion
                        m_Excel.Worksheets("MES").Cells(_Linea, 7).value = _CostoMateriales
                        m_Excel.Worksheets("MES").Cells(_Linea, 8).value = _CostoManodeObra
                        m_Excel.Worksheets("MES").Cells(_Linea, 9).value = _CostosIndirectos
                        m_Excel.Worksheets("MES").Cells(_Linea, 10).value = _Neto
                        m_Excel.Worksheets("MES").Cells(_Linea, 11).value = _Venta
                        m_Excel.Worksheets("MES").Cells(_Linea, 12).value = _Venta / _ImporteProduccion

                        _Linea = _Linea + 1

                    End If
                End If
            End If

        Next

        m_Excel.Visible = True

    End Sub
End Class