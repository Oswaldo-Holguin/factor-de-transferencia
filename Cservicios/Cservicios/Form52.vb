Public Class Frm52Bitacoras
    Public I As Integer
    Public Msg As String
    Public Style As MsgBoxStyle

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus
        TextBox3.Text = "1"
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox2.GotFocus
        TextBox3.Text = "2"
        MonthCalendar1.Visible = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub RegistroDeProducciónToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RegistroDeProducciónToolStripMenuItem.Click

        Msg = "Proporcione las fehcas Inicial y Final"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        Label1.Visible = True
        Label2.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox4.Text = "PFT"

        
        TextBox1.Focus()
    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        If TextBox3.Text = "1" Then
            TextBox1.Text = MonthCalendar1.SelectionRange.Start
        End If
        If TextBox3.Text = "2" Then
            TextBox2.Text = MonthCalendar1.SelectionRange.Start
        End If
        MonthCalendar1.Visible = False

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Dim _Hoja, _Centro As String
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Formatos Factor\Bitacoras.xlsx"
        Dim _Rows_Pr() As DataRow
        Dim _Linea, _Suma As Integer

        _Suma = 0
        _Linea = 0
        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False   'No mostramos el libro de excel

        If TextBox4.Text = "" Then
            Msg = "Debe seleccionar una Bitácora a Reportear"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If Not IsDate(TextBox1.Text) Then
            Msg = "Debe seleccionar una Fecha de Inicio del reporte válida"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        If Not IsDate(TextBox2.Text) Then
            Msg = "Debe seleccionar una Fecha Final del reporte válida"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If CDate(TextBox1.Text) > CDate(TextBox2.Text) Then
            Msg = "La Fecha Inicial no puede ser Mayor a la Fecha Final. Verifique"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If TextBox4.Text = "PFT" Then
            ' Imprime reporte de Factor de Transferencia
            _Hoja = "PFT"

            m_Excel.Worksheets(_Hoja).cells(8, 4).value = TextBox1.Text
            m_Excel.Worksheets(_Hoja).cells(9, 4).value = TextBox2.Text

            _Linea = 12
            _Centro = "FTR"
            _Rows_Pr = CServiciosDataSet26.Produccion.Select("Centro = " & "'" & _Centro & "'")
            _Suma = 0
            For Me.I = 0 To _Rows_Pr.GetUpperBound(0)
                If _Rows_Pr(I).Item("FechaInicioProduccion") >= CDate(TextBox1.Text) And _Rows_Pr(I).Item("FechaFinalProduccion") <= CDate(TextBox2.Text) Then
                    m_Excel.Worksheets(_Hoja).cells(_Linea, 2).value = _Rows_Pr(I).Item("FechaInicioProduccion")
                    m_Excel.Worksheets(_Hoja).cells(_Linea, 3).value = _Rows_Pr(I).Item("OrdenProduccion")
                    m_Excel.Worksheets(_Hoja).cells(_Linea, 4).value = _Rows_Pr(I).Item("Cantidad")
                    m_Excel.Worksheets(_Hoja).cells(_Linea, 5).value = _Rows_Pr(I).Item("Identificacion")

                    _Suma = _Suma + _Rows_Pr(I).Item("Cantidad")
                    _Linea = _Linea + 1

                End If

            Next I

            m_Excel.Worksheets(_Hoja).cells(10, 4).value = _Suma
            m_Excel.Visible = True

        End If




    End Sub

    Private Sub Frm52Bitacoras_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet26.Produccion' Puede moverla o quitarla según sea necesario.
        Me.ProduccionTableAdapter.Fill(Me.CServiciosDataSet26.Produccion)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet32.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet32.VentaFactor)

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Frm31OrdenProduccion.Show()

    End Sub
End Class