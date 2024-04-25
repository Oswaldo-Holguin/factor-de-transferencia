<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm37InformeMensual
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm37InformeMensual))
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripComboBox1 = New System.Windows.Forms.ToolStripComboBox()
        Me.ToolStripComboBox2 = New System.Windows.Forms.ToolStripComboBox()
        Me.CServiciosDataSet32 = New Cservicios.CServiciosDataSet32()
        Me.VentaFactorBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.VentaFactorTableAdapter = New Cservicios.CServiciosDataSet32TableAdapters.VentaFactorTableAdapter()
        Me.TableAdapterManager = New Cservicios.CServiciosDataSet32TableAdapters.TableAdapterManager()
        Me.CServiciosDataSet26 = New Cservicios.CServiciosDataSet26()
        Me.ProduccionBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ProduccionTableAdapter = New Cservicios.CServiciosDataSet26TableAdapters.ProduccionTableAdapter()
        Me.TableAdapterManager1 = New Cservicios.CServiciosDataSet26TableAdapters.TableAdapterManager()
        Me.CServiciosDataSet29 = New Cservicios.CServiciosDataSet29()
        Me.CierreProduccionBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CierreProduccionTableAdapter = New Cservicios.CServiciosDataSet29TableAdapters.CierreProduccionTableAdapter()
        Me.TableAdapterManager2 = New Cservicios.CServiciosDataSet29TableAdapters.TableAdapterManager()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        CType(Me.CServiciosDataSet32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VentaFactorBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet26, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ProduccionBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet29, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CierreProduccionBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.BackColor = System.Drawing.Color.DodgerBlue
        Me.BindingNavigator1.CountItem = Nothing
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton4, Me.ToolStripComboBox1, Me.ToolStripComboBox2})
        Me.BindingNavigator1.Location = New System.Drawing.Point(0, 0)
        Me.BindingNavigator1.MoveFirstItem = Nothing
        Me.BindingNavigator1.MoveLastItem = Nothing
        Me.BindingNavigator1.MoveNextItem = Nothing
        Me.BindingNavigator1.MovePreviousItem = Nothing
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Nothing
        Me.BindingNavigator1.Size = New System.Drawing.Size(670, 25)
        Me.BindingNavigator1.TabIndex = 3
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
        Me.ToolStripButton1.ToolTipText = "Salir"
        '
        'ToolStripButton4
        '
        Me.ToolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton4.Image = CType(resources.GetObject("ToolStripButton4.Image"), System.Drawing.Image)
        Me.ToolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton4.Text = "ToolStripButton4"
        Me.ToolStripButton4.ToolTipText = "Guardar Datos" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ToolStripComboBox1
        '
        Me.ToolStripComboBox1.Name = "ToolStripComboBox1"
        Me.ToolStripComboBox1.Size = New System.Drawing.Size(121, 25)
        Me.ToolStripComboBox1.Text = "Seleccione el Mes"
        Me.ToolStripComboBox1.ToolTipText = "Seleccione el Mes a Generar" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ToolStripComboBox2
        '
        Me.ToolStripComboBox2.Name = "ToolStripComboBox2"
        Me.ToolStripComboBox2.Size = New System.Drawing.Size(121, 25)
        Me.ToolStripComboBox2.Text = "Seleccione el Año"
        '
        'CServiciosDataSet32
        '
        Me.CServiciosDataSet32.DataSetName = "CServiciosDataSet32"
        Me.CServiciosDataSet32.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'VentaFactorBindingSource
        '
        Me.VentaFactorBindingSource.DataMember = "VentaFactor"
        Me.VentaFactorBindingSource.DataSource = Me.CServiciosDataSet32
        '
        'VentaFactorTableAdapter
        '
        Me.VentaFactorTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.UpdateOrder = Cservicios.CServiciosDataSet32TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        Me.TableAdapterManager.VentaFactorTableAdapter = Me.VentaFactorTableAdapter
        '
        'CServiciosDataSet26
        '
        Me.CServiciosDataSet26.DataSetName = "CServiciosDataSet26"
        Me.CServiciosDataSet26.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ProduccionBindingSource
        '
        Me.ProduccionBindingSource.DataMember = "Produccion"
        Me.ProduccionBindingSource.DataSource = Me.CServiciosDataSet26
        '
        'ProduccionTableAdapter
        '
        Me.ProduccionTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager1
        '
        Me.TableAdapterManager1.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager1.ProduccionTableAdapter = Me.ProduccionTableAdapter
        Me.TableAdapterManager1.UpdateOrder = Cservicios.CServiciosDataSet26TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'CServiciosDataSet29
        '
        Me.CServiciosDataSet29.DataSetName = "CServiciosDataSet29"
        Me.CServiciosDataSet29.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CierreProduccionBindingSource
        '
        Me.CierreProduccionBindingSource.DataMember = "CierreProduccion"
        Me.CierreProduccionBindingSource.DataSource = Me.CServiciosDataSet29
        '
        'CierreProduccionTableAdapter
        '
        Me.CierreProduccionTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager2
        '
        Me.TableAdapterManager2.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager2.CierreProduccionTableAdapter = Me.CierreProduccionTableAdapter
        Me.TableAdapterManager2.UpdateOrder = Cservicios.CServiciosDataSet29TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'Frm37InformeMensual
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(670, 40)
        Me.Controls.Add(Me.BindingNavigator1)
        Me.Name = "Frm37InformeMensual"
        Me.Text = "Informe Mensual"
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        CType(Me.CServiciosDataSet32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VentaFactorBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet26, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ProduccionBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet29, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CierreProduccionBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton4 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripComboBox1 As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents ToolStripComboBox2 As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents CServiciosDataSet32 As Cservicios.CServiciosDataSet32
    Friend WithEvents VentaFactorBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents VentaFactorTableAdapter As Cservicios.CServiciosDataSet32TableAdapters.VentaFactorTableAdapter
    Friend WithEvents TableAdapterManager As Cservicios.CServiciosDataSet32TableAdapters.TableAdapterManager
    Friend WithEvents CServiciosDataSet26 As Cservicios.CServiciosDataSet26
    Friend WithEvents ProduccionBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ProduccionTableAdapter As Cservicios.CServiciosDataSet26TableAdapters.ProduccionTableAdapter
    Friend WithEvents TableAdapterManager1 As Cservicios.CServiciosDataSet26TableAdapters.TableAdapterManager
    Friend WithEvents CServiciosDataSet29 As Cservicios.CServiciosDataSet29
    Friend WithEvents CierreProduccionBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CierreProduccionTableAdapter As Cservicios.CServiciosDataSet29TableAdapters.CierreProduccionTableAdapter
    Friend WithEvents TableAdapterManager2 As Cservicios.CServiciosDataSet29TableAdapters.TableAdapterManager
End Class
