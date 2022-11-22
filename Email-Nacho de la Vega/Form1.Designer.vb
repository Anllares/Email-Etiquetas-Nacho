<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
    	Dim dataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Dim dataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Dim dataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Dim dataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Dim dataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Dim dataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Dim dataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Dim dataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Dim dataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    	Me.DataGridViewDatos = New System.Windows.Forms.DataGridView()
    	Me.ButtonElegirFichero = New System.Windows.Forms.Button()
    	Me.ButtonSepararDatos = New System.Windows.Forms.Button()
    	Me.DataGridView1 = New System.Windows.Forms.DataGridView()
    	Me.AxwebBrowser1 = New System.Windows.Forms.WebBrowser()
    	Me.openFileDialog1 = New System.Windows.Forms.OpenFileDialog()
    	Me.groupBox1 = New System.Windows.Forms.GroupBox()
    	Me.textBoxNombrePrueba = New System.Windows.Forms.TextBox()
    	Me.TextBoxEmailPrueba = New System.Windows.Forms.TextBox()
    	Me.buttonEspera = New System.Windows.Forms.Button()
    	Me.buttonGenerarEtiquetas = New System.Windows.Forms.Button()
    	Me.buttonPrueba = New System.Windows.Forms.Button()
    	Me.labelSubjectEmail = New System.Windows.Forms.Label()
    	Me.textBoxSubjectEmail = New System.Windows.Forms.TextBox()
    	Me.dataGridViewEnvios = New System.Windows.Forms.DataGridView()
    	Me.buttonParar = New System.Windows.Forms.Button()
    	Me.buttonSalir = New System.Windows.Forms.Button()
    	Me.textBoxEnvio = New System.Windows.Forms.TextBox()
    	Me.textBoxMensajes = New System.Windows.Forms.TextBox()
    	Me.buttonEnviar = New System.Windows.Forms.Button()
    	Me.checkBoxCarta = New System.Windows.Forms.CheckBox()
    	Me.checkBoxEmail = New System.Windows.Forms.CheckBox()
    	Me.buttonCargarDatos = New System.Windows.Forms.Button()
    	Me.textBoxTotalRegistros = New System.Windows.Forms.TextBox()
    	Me.textBoxFiltrados = New System.Windows.Forms.TextBox()
    	Me.dataGridView2 = New System.Windows.Forms.DataGridView()
    	Me.textBoxElegidos = New System.Windows.Forms.TextBox()
    	Me.buttonMensajeElegirFiltro = New System.Windows.Forms.Button()
    	Me.groupBox2 = New System.Windows.Forms.GroupBox()
    	Me.buttonDatosProveedores = New System.Windows.Forms.Button()
    	CType(Me.DataGridViewDatos,System.ComponentModel.ISupportInitialize).BeginInit
    	CType(Me.DataGridView1,System.ComponentModel.ISupportInitialize).BeginInit
    	Me.groupBox1.SuspendLayout
    	CType(Me.dataGridViewEnvios,System.ComponentModel.ISupportInitialize).BeginInit
    	CType(Me.dataGridView2,System.ComponentModel.ISupportInitialize).BeginInit
    	Me.groupBox2.SuspendLayout
    	Me.SuspendLayout
    	'
    	'DataGridViewDatos
    	'
    	Me.DataGridViewDatos.AllowUserToAddRows = false
    	Me.DataGridViewDatos.AllowUserToDeleteRows = false
    	Me.DataGridViewDatos.BackgroundColor = System.Drawing.Color.White
    	dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
    	dataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
    	dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    	Me.DataGridViewDatos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1
    	Me.DataGridViewDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    	dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
    	dataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
    	dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
    	Me.DataGridViewDatos.DefaultCellStyle = dataGridViewCellStyle2
    	Me.DataGridViewDatos.Location = New System.Drawing.Point(13, 62)
    	Me.DataGridViewDatos.MultiSelect = false
    	Me.DataGridViewDatos.Name = "DataGridViewDatos"
    	Me.DataGridViewDatos.ReadOnly = true
    	dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
    	dataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
    	dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    	Me.DataGridViewDatos.RowHeadersDefaultCellStyle = dataGridViewCellStyle3
    	Me.DataGridViewDatos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
    	Me.DataGridViewDatos.Size = New System.Drawing.Size(531, 252)
    	Me.DataGridViewDatos.TabIndex = 0
    	AddHandler Me.DataGridViewDatos.CellContentClick, AddressOf Me.DataGridViewDatosCellContentClick
    	'
    	'ButtonElegirFichero
    	'
    	Me.ButtonElegirFichero.BackColor = System.Drawing.SystemColors.Control
    	Me.ButtonElegirFichero.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.ButtonElegirFichero.Font = New System.Drawing.Font("Microsoft Sans Serif", 9!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.ButtonElegirFichero.Location = New System.Drawing.Point(6, 23)
    	Me.ButtonElegirFichero.Name = "ButtonElegirFichero"
    	Me.ButtonElegirFichero.Size = New System.Drawing.Size(116, 31)
    	Me.ButtonElegirFichero.TabIndex = 1
    	Me.ButtonElegirFichero.Text = "Elegir Fichero"
    	Me.ButtonElegirFichero.UseVisualStyleBackColor = false
    	AddHandler Me.ButtonElegirFichero.Click, AddressOf Me.ButtonElegirFicheroClick
    	'
    	'ButtonSepararDatos
    	'
    	Me.ButtonSepararDatos.BackColor = System.Drawing.SystemColors.Control
    	Me.ButtonSepararDatos.Cursor = System.Windows.Forms.Cursors.Arrow
    	Me.ButtonSepararDatos.FlatAppearance.BorderSize = 2
    	Me.ButtonSepararDatos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.ButtonSepararDatos.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.ButtonSepararDatos.Location = New System.Drawing.Point(465, 374)
    	Me.ButtonSepararDatos.Name = "ButtonSepararDatos"
    	Me.ButtonSepararDatos.Size = New System.Drawing.Size(79, 111)
    	Me.ButtonSepararDatos.TabIndex = 2
    	Me.ButtonSepararDatos.Text = "Filtrar Datos"
    	Me.ButtonSepararDatos.UseVisualStyleBackColor = false
    	AddHandler Me.ButtonSepararDatos.Click, AddressOf Me.ButtonSepararDatosClick
    	'
    	'DataGridView1
    	'
    	Me.DataGridView1.AllowUserToAddRows = false
    	Me.DataGridView1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Left),System.Windows.Forms.AnchorStyles)
    	Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
    	dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Desktop
    	dataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
    	dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    	Me.DataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4
    	Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    	dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
    	dataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
    	dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
    	Me.DataGridView1.DefaultCellStyle = dataGridViewCellStyle5
    	Me.DataGridView1.Location = New System.Drawing.Point(13, 486)
    	Me.DataGridView1.Name = "DataGridView1"
    	dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
    	dataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
    	dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    	Me.DataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle6
    	Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
    	Me.DataGridView1.Size = New System.Drawing.Size(531, 161)
    	Me.DataGridView1.TabIndex = 3
    	AddHandler Me.DataGridView1.CellContentClick, AddressOf Me.DataGridView1CellContentClick
    	AddHandler Me.DataGridView1.RowsAdded, AddressOf Me.DataGridView1RowsAdded
    	AddHandler Me.DataGridView1.RowsRemoved, AddressOf Me.DataGridView1RowsRemoved
    	'
    	'AxwebBrowser1
    	'
    	Me.AxwebBrowser1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Left)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.AxwebBrowser1.Location = New System.Drawing.Point(6, 177)
    	Me.AxwebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
    	Me.AxwebBrowser1.Name = "AxwebBrowser1"
    	Me.AxwebBrowser1.ScriptErrorsSuppressed = true
    	Me.AxwebBrowser1.Size = New System.Drawing.Size(753, 476)
    	Me.AxwebBrowser1.TabIndex = 4
    	AddHandler Me.AxwebBrowser1.DocumentCompleted, AddressOf Me.AxwebBrowser1DocumentCompleted
    	'
    	'openFileDialog1
    	'
    	Me.openFileDialog1.FileName = "openFileDialog1"
    	'
    	'groupBox1
    	'
    	Me.groupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Left)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.groupBox1.BackColor = System.Drawing.Color.Transparent
    	Me.groupBox1.Controls.Add(Me.textBoxNombrePrueba)
    	Me.groupBox1.Controls.Add(Me.TextBoxEmailPrueba)
    	Me.groupBox1.Controls.Add(Me.buttonEspera)
    	Me.groupBox1.Controls.Add(Me.buttonGenerarEtiquetas)
    	Me.groupBox1.Controls.Add(Me.buttonPrueba)
    	Me.groupBox1.Controls.Add(Me.labelSubjectEmail)
    	Me.groupBox1.Controls.Add(Me.textBoxSubjectEmail)
    	Me.groupBox1.Controls.Add(Me.dataGridViewEnvios)
    	Me.groupBox1.Controls.Add(Me.buttonParar)
    	Me.groupBox1.Controls.Add(Me.buttonSalir)
    	Me.groupBox1.Controls.Add(Me.textBoxEnvio)
    	Me.groupBox1.Controls.Add(Me.textBoxMensajes)
    	Me.groupBox1.Controls.Add(Me.buttonEnviar)
    	Me.groupBox1.Controls.Add(Me.checkBoxCarta)
    	Me.groupBox1.Controls.Add(Me.checkBoxEmail)
    	Me.groupBox1.Controls.Add(Me.AxwebBrowser1)
    	Me.groupBox1.Controls.Add(Me.ButtonElegirFichero)
    	Me.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    	Me.groupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.groupBox1.ForeColor = System.Drawing.SystemColors.ControlText
    	Me.groupBox1.Location = New System.Drawing.Point(555, 10)
    	Me.groupBox1.Name = "groupBox1"
    	Me.groupBox1.Size = New System.Drawing.Size(765, 659)
    	Me.groupBox1.TabIndex = 5
    	Me.groupBox1.TabStop = false
    	AddHandler Me.groupBox1.Enter, AddressOf Me.GroupBox1Enter
    	'
    	'textBoxNombrePrueba
    	'
    	Me.textBoxNombrePrueba.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.textBoxNombrePrueba.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.textBoxNombrePrueba.Location = New System.Drawing.Point(418, 136)
    	Me.textBoxNombrePrueba.Name = "textBoxNombrePrueba"
    	Me.textBoxNombrePrueba.Size = New System.Drawing.Size(406, 26)
    	Me.textBoxNombrePrueba.TabIndex = 15
    	Me.textBoxNombrePrueba.Text = "PRUEBA ENVIO EMAILS"
    	'
    	'TextBoxEmailPrueba
    	'
    	Me.TextBoxEmailPrueba.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.TextBoxEmailPrueba.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.TextBoxEmailPrueba.Location = New System.Drawing.Point(6, 136)
    	Me.TextBoxEmailPrueba.Name = "TextBoxEmailPrueba"
    	Me.TextBoxEmailPrueba.Size = New System.Drawing.Size(406, 26)
    	Me.TextBoxEmailPrueba.TabIndex = 14
    	Me.TextBoxEmailPrueba.Text = "beni@nachodelavega.es"
    	'
    	'buttonEspera
    	'
    	Me.buttonEspera.Anchor = System.Windows.Forms.AnchorStyles.None
    	Me.buttonEspera.BackColor = System.Drawing.Color.White
    	Me.buttonEspera.ForeColor = System.Drawing.Color.Red
    	Me.buttonEspera.Location = New System.Drawing.Point(14, 133)
    	Me.buttonEspera.Name = "buttonEspera"
    	Me.buttonEspera.Size = New System.Drawing.Size(464, 38)
    	Me.buttonEspera.TabIndex = 12
    	Me.buttonEspera.Text = "ESPERA MIENTRAS CONSTRUYO LAS ETIQUETAS"
    	Me.buttonEspera.UseVisualStyleBackColor = false
    	Me.buttonEspera.Visible = false
    	'
    	'buttonGenerarEtiquetas
    	'
    	Me.buttonGenerarEtiquetas.BackColor = System.Drawing.SystemColors.Control
    	Me.buttonGenerarEtiquetas.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.buttonGenerarEtiquetas.ForeColor = System.Drawing.Color.Black
    	Me.buttonGenerarEtiquetas.Location = New System.Drawing.Point(196, 92)
    	Me.buttonGenerarEtiquetas.Name = "buttonGenerarEtiquetas"
    	Me.buttonGenerarEtiquetas.Size = New System.Drawing.Size(161, 38)
    	Me.buttonGenerarEtiquetas.TabIndex = 11
    	Me.buttonGenerarEtiquetas.Text = "Generar Etiquetas"
    	Me.buttonGenerarEtiquetas.UseVisualStyleBackColor = false
    	AddHandler Me.buttonGenerarEtiquetas.Click, AddressOf Me.ButtonGenerarEtiquetasClick
    	'
    	'buttonPrueba
    	'
    	Me.buttonPrueba.BackColor = System.Drawing.SystemColors.Control
    	Me.buttonPrueba.Enabled = false
    	Me.buttonPrueba.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.buttonPrueba.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.buttonPrueba.ForeColor = System.Drawing.Color.Black
    	Me.buttonPrueba.Location = New System.Drawing.Point(99, 92)
    	Me.buttonPrueba.Name = "buttonPrueba"
    	Me.buttonPrueba.Size = New System.Drawing.Size(96, 38)
    	Me.buttonPrueba.TabIndex = 10
    	Me.buttonPrueba.Text = "Prueba Email"
    	Me.buttonPrueba.UseVisualStyleBackColor = false
    	AddHandler Me.buttonPrueba.Click, AddressOf Me.ButtonPruebaClick
    	'
    	'labelSubjectEmail
    	'
    	Me.labelSubjectEmail.Location = New System.Drawing.Point(6, 57)
    	Me.labelSubjectEmail.Name = "labelSubjectEmail"
    	Me.labelSubjectEmail.Size = New System.Drawing.Size(114, 23)
    	Me.labelSubjectEmail.TabIndex = 9
    	Me.labelSubjectEmail.Text = "Título Email:"
    	Me.labelSubjectEmail.Visible = false
    	'
    	'textBoxSubjectEmail
    	'
    	Me.textBoxSubjectEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.textBoxSubjectEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.textBoxSubjectEmail.Location = New System.Drawing.Point(126, 55)
    	Me.textBoxSubjectEmail.Name = "textBoxSubjectEmail"
    	Me.textBoxSubjectEmail.Size = New System.Drawing.Size(634, 26)
    	Me.textBoxSubjectEmail.TabIndex = 8
    	Me.textBoxSubjectEmail.Text = "[Nacho de la Vega] Información."
    	Me.textBoxSubjectEmail.Visible = false
    	'
    	'dataGridViewEnvios
    	'
    	Me.dataGridViewEnvios.AllowUserToDeleteRows = false
    	Me.dataGridViewEnvios.AllowUserToResizeColumns = false
    	Me.dataGridViewEnvios.AllowUserToResizeRows = false
    	Me.dataGridViewEnvios.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    	Me.dataGridViewEnvios.Location = New System.Drawing.Point(475, 92)
    	Me.dataGridViewEnvios.MultiSelect = false
    	Me.dataGridViewEnvios.Name = "dataGridViewEnvios"
    	Me.dataGridViewEnvios.ReadOnly = true
    	Me.dataGridViewEnvios.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
    	Me.dataGridViewEnvios.Size = New System.Drawing.Size(16, 38)
    	Me.dataGridViewEnvios.TabIndex = 7
    	Me.dataGridViewEnvios.Visible = false
    	AddHandler Me.dataGridViewEnvios.RowsAdded, AddressOf Me.DataGridViewEnviosRowsAdded
    	'
    	'buttonParar
    	'
    	Me.buttonParar.BackColor = System.Drawing.SystemColors.Control
    	Me.buttonParar.Enabled = false
    	Me.buttonParar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.buttonParar.ForeColor = System.Drawing.Color.Black
    	Me.buttonParar.Location = New System.Drawing.Point(357, 92)
    	Me.buttonParar.Name = "buttonParar"
    	Me.buttonParar.Size = New System.Drawing.Size(114, 38)
    	Me.buttonParar.TabIndex = 6
    	Me.buttonParar.Text = "Parar Envío"
    	Me.buttonParar.UseVisualStyleBackColor = false
    	AddHandler Me.buttonParar.Click, AddressOf Me.ButtonPararClick
    	'
    	'buttonSalir
    	'
    	Me.buttonSalir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.buttonSalir.BackColor = System.Drawing.SystemColors.Control
    	Me.buttonSalir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.buttonSalir.ForeColor = System.Drawing.Color.Black
    	Me.buttonSalir.Location = New System.Drawing.Point(684, 92)
    	Me.buttonSalir.Name = "buttonSalir"
    	Me.buttonSalir.Size = New System.Drawing.Size(76, 38)
    	Me.buttonSalir.TabIndex = 5
    	Me.buttonSalir.Text = "Salir"
    	Me.buttonSalir.UseVisualStyleBackColor = false
    	AddHandler Me.buttonSalir.Click, AddressOf Me.ButtonSalirClick
    	'
    	'textBoxEnvio
    	'
    	Me.textBoxEnvio.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Left)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.textBoxEnvio.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.textBoxEnvio.Location = New System.Drawing.Point(605, 92)
    	Me.textBoxEnvio.Multiline = true
    	Me.textBoxEnvio.Name = "textBoxEnvio"
    	Me.textBoxEnvio.ReadOnly = true
    	Me.textBoxEnvio.ScrollBars = System.Windows.Forms.ScrollBars.Both
    	Me.textBoxEnvio.Size = New System.Drawing.Size(73, 34)
    	Me.textBoxEnvio.TabIndex = 4
    	Me.textBoxEnvio.Visible = false
    	'
    	'textBoxMensajes
    	'
    	Me.textBoxMensajes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.textBoxMensajes.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.textBoxMensajes.Location = New System.Drawing.Point(126, 23)
    	Me.textBoxMensajes.Name = "textBoxMensajes"
    	Me.textBoxMensajes.ReadOnly = true
    	Me.textBoxMensajes.Size = New System.Drawing.Size(634, 26)
    	Me.textBoxMensajes.TabIndex = 3
    	AddHandler Me.textBoxMensajes.TextChanged, AddressOf Me.TextBoxMensajesTextChanged
    	'
    	'buttonEnviar
    	'
    	Me.buttonEnviar.BackColor = System.Drawing.SystemColors.Control
    	Me.buttonEnviar.Enabled = false
    	Me.buttonEnviar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.buttonEnviar.ForeColor = System.Drawing.Color.Black
    	Me.buttonEnviar.Location = New System.Drawing.Point(6, 92)
    	Me.buttonEnviar.Name = "buttonEnviar"
    	Me.buttonEnviar.Size = New System.Drawing.Size(93, 38)
    	Me.buttonEnviar.TabIndex = 2
    	Me.buttonEnviar.Text = "Enviar"
    	Me.buttonEnviar.UseVisualStyleBackColor = false
    	AddHandler Me.buttonEnviar.Click, AddressOf Me.ButtonEnviarClick
    	'
    	'checkBoxCarta
    	'
    	Me.checkBoxCarta.BackColor = System.Drawing.Color.Transparent
    	Me.checkBoxCarta.Enabled = false
    	Me.checkBoxCarta.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.checkBoxCarta.Location = New System.Drawing.Point(497, 92)
    	Me.checkBoxCarta.Name = "checkBoxCarta"
    	Me.checkBoxCarta.Size = New System.Drawing.Size(30, 24)
    	Me.checkBoxCarta.TabIndex = 1
    	Me.checkBoxCarta.Text = "Carta"
    	Me.checkBoxCarta.UseVisualStyleBackColor = false
    	Me.checkBoxCarta.Visible = false
    	'
    	'checkBoxEmail
    	'
    	Me.checkBoxEmail.BackColor = System.Drawing.Color.Transparent
    	Me.checkBoxEmail.Enabled = false
    	Me.checkBoxEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.checkBoxEmail.Location = New System.Drawing.Point(497, 106)
    	Me.checkBoxEmail.Name = "checkBoxEmail"
    	Me.checkBoxEmail.Size = New System.Drawing.Size(30, 24)
    	Me.checkBoxEmail.TabIndex = 0
    	Me.checkBoxEmail.Text = "E-mail"
    	Me.checkBoxEmail.UseVisualStyleBackColor = false
    	Me.checkBoxEmail.Visible = false
    	'
    	'buttonCargarDatos
    	'
    	Me.buttonCargarDatos.BackColor = System.Drawing.SystemColors.Control
    	Me.buttonCargarDatos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.buttonCargarDatos.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.buttonCargarDatos.Location = New System.Drawing.Point(13, 33)
    	Me.buttonCargarDatos.Name = "buttonCargarDatos"
    	Me.buttonCargarDatos.Size = New System.Drawing.Size(286, 29)
    	Me.buttonCargarDatos.TabIndex = 6
    	Me.buttonCargarDatos.Text = "Cargar Datos ContaPlus"
    	Me.buttonCargarDatos.UseVisualStyleBackColor = false
    	AddHandler Me.buttonCargarDatos.Click, AddressOf Me.ButtonCargarDatosClick
    	'
    	'textBoxTotalRegistros
    	'
    	Me.textBoxTotalRegistros.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.textBoxTotalRegistros.Location = New System.Drawing.Point(13, 313)
    	Me.textBoxTotalRegistros.Name = "textBoxTotalRegistros"
    	Me.textBoxTotalRegistros.ReadOnly = true
    	Me.textBoxTotalRegistros.Size = New System.Drawing.Size(531, 22)
    	Me.textBoxTotalRegistros.TabIndex = 7
    	Me.textBoxTotalRegistros.Text = "...."
    	'
    	'textBoxFiltrados
    	'
    	Me.textBoxFiltrados.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left),System.Windows.Forms.AnchorStyles)
    	Me.textBoxFiltrados.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.textBoxFiltrados.Location = New System.Drawing.Point(13, 645)
    	Me.textBoxFiltrados.Name = "textBoxFiltrados"
    	Me.textBoxFiltrados.ReadOnly = true
    	Me.textBoxFiltrados.Size = New System.Drawing.Size(531, 22)
    	Me.textBoxFiltrados.TabIndex = 8
    	Me.textBoxFiltrados.Text = "...."
    	'
    	'dataGridView2
    	'
    	Me.dataGridView2.AllowUserToAddRows = false
    	Me.dataGridView2.AllowUserToDeleteRows = false
    	Me.dataGridView2.BackgroundColor = System.Drawing.Color.White
    	dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
    	dataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
    	dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    	Me.dataGridView2.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7
    	Me.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    	dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
    	dataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
    	dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
    	Me.dataGridView2.DefaultCellStyle = dataGridViewCellStyle8
    	Me.dataGridView2.Location = New System.Drawing.Point(13, 371)
    	Me.dataGridView2.Name = "dataGridView2"
    	Me.dataGridView2.ReadOnly = true
    	dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    	dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
    	dataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
    	dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
    	dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    	dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    	Me.dataGridView2.RowHeadersDefaultCellStyle = dataGridViewCellStyle9
    	Me.dataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
    	Me.dataGridView2.Size = New System.Drawing.Size(235, 116)
    	Me.dataGridView2.TabIndex = 9
    	AddHandler Me.dataGridView2.SelectionChanged, AddressOf Me.DataGridView2SelectionChanged
    	'
    	'textBoxElegidos
    	'
    	Me.textBoxElegidos.BackColor = System.Drawing.Color.White
    	Me.textBoxElegidos.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.textBoxElegidos.Location = New System.Drawing.Point(248, 371)
    	Me.textBoxElegidos.Multiline = true
    	Me.textBoxElegidos.Name = "textBoxElegidos"
    	Me.textBoxElegidos.ReadOnly = true
    	Me.textBoxElegidos.Size = New System.Drawing.Size(216, 116)
    	Me.textBoxElegidos.TabIndex = 10
    	'
    	'buttonMensajeElegirFiltro
    	'
    	Me.buttonMensajeElegirFiltro.BackColor = System.Drawing.SystemColors.Control
    	Me.buttonMensajeElegirFiltro.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.buttonMensajeElegirFiltro.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.buttonMensajeElegirFiltro.Location = New System.Drawing.Point(13, 343)
    	Me.buttonMensajeElegirFiltro.Name = "buttonMensajeElegirFiltro"
    	Me.buttonMensajeElegirFiltro.Size = New System.Drawing.Size(531, 29)
    	Me.buttonMensajeElegirFiltro.TabIndex = 11
    	Me.buttonMensajeElegirFiltro.Text = "Elige el tipo del clientes a filtrar"
    	Me.buttonMensajeElegirFiltro.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    	Me.buttonMensajeElegirFiltro.UseVisualStyleBackColor = false
    	'
    	'groupBox2
    	'
    	Me.groupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Left)  _
    	    	    	Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
    	Me.groupBox2.Controls.Add(Me.buttonDatosProveedores)
    	Me.groupBox2.Location = New System.Drawing.Point(9, 12)
    	Me.groupBox2.Name = "groupBox2"
    	Me.groupBox2.Size = New System.Drawing.Size(540, 657)
    	Me.groupBox2.TabIndex = 12
    	Me.groupBox2.TabStop = false
    	'
    	'buttonDatosProveedores
    	'
    	Me.buttonDatosProveedores.BackColor = System.Drawing.SystemColors.Control
    	Me.buttonDatosProveedores.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    	Me.buttonDatosProveedores.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
    	Me.buttonDatosProveedores.Location = New System.Drawing.Point(290, 21)
    	Me.buttonDatosProveedores.Name = "buttonDatosProveedores"
    	Me.buttonDatosProveedores.Size = New System.Drawing.Size(245, 29)
    	Me.buttonDatosProveedores.TabIndex = 13
    	Me.buttonDatosProveedores.Text = "Cargar Datos Profesionales"
    	Me.buttonDatosProveedores.UseVisualStyleBackColor = false
    	AddHandler Me.buttonDatosProveedores.Click, AddressOf Me.ButtonDatosProveedoresClick
    	'
    	'Form1
    	'
    	Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
    	Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    	Me.BackColor = System.Drawing.Color.SteelBlue
    	Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    	Me.ClientSize = New System.Drawing.Size(1322, 673)
    	Me.Controls.Add(Me.buttonMensajeElegirFiltro)
    	Me.Controls.Add(Me.textBoxElegidos)
    	Me.Controls.Add(Me.dataGridView2)
    	Me.Controls.Add(Me.textBoxFiltrados)
    	Me.Controls.Add(Me.textBoxTotalRegistros)
    	Me.Controls.Add(Me.buttonCargarDatos)
    	Me.Controls.Add(Me.groupBox1)
    	Me.Controls.Add(Me.DataGridView1)
    	Me.Controls.Add(Me.ButtonSepararDatos)
    	Me.Controls.Add(Me.DataGridViewDatos)
    	Me.Controls.Add(Me.groupBox2)
    	Me.Name = "Form1"
    	Me.Text = "Principal"
    	Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
    	AddHandler FormClosing, AddressOf Me.Form1FormClosing
    	AddHandler FormClosed, AddressOf Me.Form1FormClosed
    	AddHandler Load, AddressOf Me.Form1Load
    	CType(Me.DataGridViewDatos,System.ComponentModel.ISupportInitialize).EndInit
    	CType(Me.DataGridView1,System.ComponentModel.ISupportInitialize).EndInit
    	Me.groupBox1.ResumeLayout(false)
    	Me.groupBox1.PerformLayout
    	CType(Me.dataGridViewEnvios,System.ComponentModel.ISupportInitialize).EndInit
    	CType(Me.dataGridView2,System.ComponentModel.ISupportInitialize).EndInit
    	Me.groupBox2.ResumeLayout(false)
    	Me.ResumeLayout(false)
    	Me.PerformLayout
    End Sub
    Private TextBoxEmailPrueba As System.Windows.Forms.TextBox
    Private textBoxNombrePrueba As System.Windows.Forms.TextBox
    Private buttonEspera As System.Windows.Forms.Button
    Private buttonGenerarEtiquetas As System.Windows.Forms.Button
    Friend buttonDatosProveedores As System.Windows.Forms.Button
    Private buttonPrueba As System.Windows.Forms.Button
    Private labelSubjectEmail As System.Windows.Forms.Label
    Private textBoxSubjectEmail As System.Windows.Forms.TextBox
    Private groupBox2 As System.Windows.Forms.GroupBox
    Private dataGridViewEnvios As System.Windows.Forms.DataGridView
    Private buttonParar As System.Windows.Forms.Button
    Private buttonSalir As System.Windows.Forms.Button
    Friend buttonMensajeElegirFiltro As System.Windows.Forms.Button
    Private textBoxElegidos As System.Windows.Forms.TextBox
    Private dataGridView2 As System.Windows.Forms.DataGridView
    Private textBoxEnvio As System.Windows.Forms.TextBox
    Private textBoxFiltrados As System.Windows.Forms.TextBox
    Private textBoxTotalRegistros As System.Windows.Forms.TextBox
    Friend buttonCargarDatos As System.Windows.Forms.Button
    Private buttonEnviar As System.Windows.Forms.Button
    Private textBoxMensajes As System.Windows.Forms.TextBox
    Private checkBoxEmail As System.Windows.Forms.CheckBox
    Private checkBoxCarta As System.Windows.Forms.CheckBox
    Private groupBox1 As System.Windows.Forms.GroupBox
    Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
    Private AxwebBrowser1 As System.Windows.Forms.WebBrowser

    Friend WithEvents DataGridViewDatos As DataGridView
    Friend WithEvents ButtonElegirFichero As Button
    Friend WithEvents ButtonSepararDatos As Button
    Friend WithEvents DataGridView1 As DataGridView
End Class
