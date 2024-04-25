<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm13ElectrosParaHoy
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm13ElectrosParaHoy))
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column9 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column12 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CServiciosDataSet4 = New Cservicios.CServiciosDataSet4
        Me.CitasBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CitasTableAdapter = New Cservicios.CServiciosDataSet4TableAdapters.CitasTableAdapter
        Me.TableAdapterManager = New Cservicios.CServiciosDataSet4TableAdapters.TableAdapterManager
        Me.CServiciosDataSet2 = New Cservicios.CServiciosDataSet2
        Me.PacientesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.PacientesTableAdapter = New Cservicios.CServiciosDataSet2TableAdapters.PacientesTableAdapter
        Me.TableAdapterManager1 = New Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager
        Me.CServiciosDataSet8 = New Cservicios.CServiciosDataSet8
        Me.CentrosBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CentrosTableAdapter = New Cservicios.CServiciosDataSet8TableAdapters.CentrosTableAdapter
        Me.TableAdapterManager2 = New Cservicios.CServiciosDataSet8TableAdapters.TableAdapterManager
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CitasBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PacientesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CentrosBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.BackColor = System.Drawing.Color.DodgerBlue
        Me.BindingNavigator1.CountItem = Nothing
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2})
        Me.BindingNavigator1.Location = New System.Drawing.Point(0, 0)
        Me.BindingNavigator1.MoveFirstItem = Nothing
        Me.BindingNavigator1.MoveLastItem = Nothing
        Me.BindingNavigator1.MoveNextItem = Nothing
        Me.BindingNavigator1.MovePreviousItem = Nothing
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Nothing
        Me.BindingNavigator1.Size = New System.Drawing.Size(1087, 25)
        Me.BindingNavigator1.TabIndex = 0
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "ToolStripButton1"
        Me.ToolStripButton1.ToolTipText = "Salir de esta Pantalla" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "ToolStripButton2"
        Me.ToolStripButton2.ToolTipText = "Refrescar Información" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column5, Me.Column6, Me.Column7, Me.Column8, Me.Column9, Me.Column10, Me.Column11, Me.Column13, Me.Column14, Me.Column12})
        Me.DataGridView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView2.Location = New System.Drawing.Point(0, 25)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataGridView2.Size = New System.Drawing.Size(1087, 254)
        Me.DataGridView2.TabIndex = 1
        '
        'Column5
        '
        Me.Column5.HeaderText = "Paciente"
        Me.Column5.Name = "Column5"
        Me.Column5.Width = 80
        '
        'Column6
        '
        Me.Column6.HeaderText = "Cita"
        Me.Column6.Name = "Column6"
        Me.Column6.Width = 80
        '
        'Column7
        '
        Me.Column7.HeaderText = "Fecha"
        Me.Column7.Name = "Column7"
        '
        'Column8
        '
        Me.Column8.HeaderText = "Hora"
        Me.Column8.Name = "Column8"
        '
        'Column9
        '
        Me.Column9.HeaderText = "Clínica"
        Me.Column9.Name = "Column9"
        Me.Column9.Width = 80
        '
        'Column10
        '
        Me.Column10.HeaderText = "Nombre del Paciente"
        Me.Column10.Name = "Column10"
        Me.Column10.Width = 200
        '
        'Column11
        '
        Me.Column11.HeaderText = "Atiende"
        Me.Column11.Name = "Column11"
        Me.Column11.Width = 150
        '
        'Column13
        '
        Me.Column13.HeaderText = "Referencia"
        Me.Column13.Name = "Column13"
        Me.Column13.Width = 150
        '
        'Column14
        '
        Me.Column14.HeaderText = "Comentarios"
        Me.Column14.Name = "Column14"
        Me.Column14.Visible = False
        '
        'Column12
        '
        Me.Column12.HeaderText = "Asistencia"
        Me.Column12.Name = "Column12"
        Me.Column12.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column12.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'CServiciosDataSet4
        '
        Me.CServiciosDataSet4.DataSetName = "CServiciosDataSet4"
        Me.CServiciosDataSet4.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CitasBindingSource
        '
        Me.CitasBindingSource.DataMember = "Citas"
        Me.CitasBindingSource.DataSource = Me.CServiciosDataSet4
        '
        'CitasTableAdapter
        '
        Me.CitasTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.CitasTableAdapter = Me.CitasTableAdapter
        Me.TableAdapterManager.UpdateOrder = Cservicios.CServiciosDataSet4TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'CServiciosDataSet2
        '
        Me.CServiciosDataSet2.DataSetName = "CServiciosDataSet2"
        Me.CServiciosDataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'PacientesBindingSource
        '
        Me.PacientesBindingSource.DataMember = "Pacientes"
        Me.PacientesBindingSource.DataSource = Me.CServiciosDataSet2
        '
        'PacientesTableAdapter
        '
        Me.PacientesTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager1
        '
        Me.TableAdapterManager1.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager1.ConsultasTableAdapter = Nothing
        Me.TableAdapterManager1.DetalleRecetaTableAdapter = Nothing
        Me.TableAdapterManager1.DiagnosticosTableAdapter = Nothing
        Me.TableAdapterManager1.EvetosPacienteTableAdapter = Nothing
        Me.TableAdapterManager1.PacientesTableAdapter = Me.PacientesTableAdapter
        Me.TableAdapterManager1.RecetasTableAdapter = Nothing
        Me.TableAdapterManager1.UpdateOrder = Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'CServiciosDataSet8
        '
        Me.CServiciosDataSet8.DataSetName = "CServiciosDataSet8"
        Me.CServiciosDataSet8.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CentrosBindingSource
        '
        Me.CentrosBindingSource.DataMember = "Centros"
        Me.CentrosBindingSource.DataSource = Me.CServiciosDataSet8
        '
        'CentrosTableAdapter
        '
        Me.CentrosTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager2
        '
        Me.TableAdapterManager2.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager2.CentrosTableAdapter = Me.CentrosTableAdapter
        Me.TableAdapterManager2.UpdateOrder = Cservicios.CServiciosDataSet8TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'Frm13ElectrosParaHoy
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1087, 279)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.BindingNavigator1)
        Me.Name = "Frm13ElectrosParaHoy"
        Me.Text = "Electros a realizarse el dia de hoy"
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CitasBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PacientesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CentrosBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents CServiciosDataSet4 As Cservicios.CServiciosDataSet4
    Friend WithEvents CitasBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CitasTableAdapter As Cservicios.CServiciosDataSet4TableAdapters.CitasTableAdapter
    Friend WithEvents TableAdapterManager As Cservicios.CServiciosDataSet4TableAdapters.TableAdapterManager
    Friend WithEvents CServiciosDataSet2 As Cservicios.CServiciosDataSet2
    Friend WithEvents PacientesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents PacientesTableAdapter As Cservicios.CServiciosDataSet2TableAdapters.PacientesTableAdapter
    Friend WithEvents TableAdapterManager1 As Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager
    Friend WithEvents CServiciosDataSet8 As Cservicios.CServiciosDataSet8
    Friend WithEvents CentrosBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CentrosTableAdapter As Cservicios.CServiciosDataSet8TableAdapters.CentrosTableAdapter
    Friend WithEvents TableAdapterManager2 As Cservicios.CServiciosDataSet8TableAdapters.TableAdapterManager
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column12 As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
