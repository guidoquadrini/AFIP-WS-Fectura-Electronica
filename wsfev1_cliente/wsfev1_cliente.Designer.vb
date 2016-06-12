<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class wsfev1_cliente
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
        Me.cmb_cuit = New System.Windows.Forms.ComboBox()
        Me.grp_FEAuthRequest = New System.Windows.Forms.GroupBox()
        Me.lbl_cuit = New System.Windows.Forms.Label()
        Me.lbl_sign = New System.Windows.Forms.Label()
        Me.lbl_token = New System.Windows.Forms.Label()
        Me.txt_sign = New System.Windows.Forms.TextBox()
        Me.txt_token = New System.Windows.Forms.TextBox()
        Me.grp_FEDummy = New System.Windows.Forms.GroupBox()
        Me.btn_FEDummy_DbServer = New System.Windows.Forms.Button()
        Me.bnt_FEDummy_AuthServer = New System.Windows.Forms.Button()
        Me.bnt_FEDummy_AppServer = New System.Windows.Forms.Button()
        Me.grp_MetodosDeConsultasGenericas = New System.Windows.Forms.GroupBox()
        Me.dgv_metodosDeConsultasGenericas = New System.Windows.Forms.DataGridView()
        Me.cbm_metodosDeconsultasGenericas = New System.Windows.Forms.ComboBox()
        Me.grp_FECompConsultar = New System.Windows.Forms.GroupBox()
        Me.grp_FECompUltimoAutorizado = New System.Windows.Forms.GroupBox()
        Me.ListBox_FECompUltimoAutorizado_PtoVta = New System.Windows.Forms.ListBox()
        Me.lbl_FECompUltimoAutorizado_PtoVta = New System.Windows.Forms.Label()
        Me.ListBox_FECompUltimoAutorizado_CbteTipo = New System.Windows.Forms.ListBox()
        Me.btn_FECompUltimoAutorizado = New System.Windows.Forms.Button()
        Me.lbl_FECompUltimoAutorizado_CbteTipo = New System.Windows.Forms.Label()
        Me.txt_FECompConsultar_CbteNro = New System.Windows.Forms.TextBox()
        Me.lbl_FECompConsultar_CbteNro = New System.Windows.Forms.Label()
        Me.btn_FECompConsultar = New System.Windows.Forms.Button()
        Me.grp_FECAESolicitar = New System.Windows.Forms.GroupBox()
        Me.ListBox_FECAESolicitar_CbteTipo = New System.Windows.Forms.ListBox()
        Me.lbl_FECAESolicitar_CbteTipo = New System.Windows.Forms.Label()
        Me.lbl_FECAEDetRequest = New System.Windows.Forms.Label()
        Me.btn_FECAEDetRequestFinEditar = New System.Windows.Forms.Button()
        Me.btn_FECAEDetRequestEditar = New System.Windows.Forms.Button()
        Me.btn_FECAEDetRequestBorrar = New System.Windows.Forms.Button()
        Me.btn_FECAEDetRequestNuevo = New System.Windows.Forms.Button()
        Me.dgv_FECAEDetRequest = New System.Windows.Forms.DataGridView()
        Me.Concepto = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DocTipo = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DocNro = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CbteDesde = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CbteHasta = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CbteFch = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ImpTotal = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ImpTotConc = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ImpNeto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ImpOpEx = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ImpTrib = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ImpIVA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FchServDesde = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FchServHasta = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FchVtoPago = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MonId = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.MonCotiz = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lbl_FECAESolicitar_CantReg = New System.Windows.Forms.Label()
        Me.txt_FECAESolicitar_CantReg = New System.Windows.Forms.TextBox()
        Me.btn_FECAESolicitar = New System.Windows.Forms.Button()
        Me.lbl_FECAESolicitar_PtoVta = New System.Windows.Forms.Label()
        Me.txt_FECAESolicitar_PtoVta = New System.Windows.Forms.TextBox()
        Me.txt_Periodo = New System.Windows.Forms.TextBox()
        Me.lbl_Periodo = New System.Windows.Forms.Label()
        Me.btn_FECAEASolicitar = New System.Windows.Forms.Button()
        Me.cmb_Orden = New System.Windows.Forms.ComboBox()
        Me.lbl_Orden = New System.Windows.Forms.Label()
        Me.gpr_FECAEAConsultar = New System.Windows.Forms.GroupBox()
        Me.grp_FECAEASolicitar = New System.Windows.Forms.GroupBox()
        Me.btn_FECAEAConsultar = New System.Windows.Forms.Button()
        Me.gpr_FECAEASinMovimientoConsultar = New System.Windows.Forms.GroupBox()
        Me.gpr_FECAEASinMovimientoInformar = New System.Windows.Forms.GroupBox()
        Me.txt_FECAEASinMovimientoConsultar_CAEA = New System.Windows.Forms.TextBox()
        Me.lbl_FECAEASinMovimientoConsultar_CAEA = New System.Windows.Forms.Label()
        Me.txt_FECAEASinMovimientoConsultar_PtoVta = New System.Windows.Forms.TextBox()
        Me.lbl_FECAEASinMovimientoConsultar_PtoVta = New System.Windows.Forms.Label()
        Me.bnt_FECAEASinMovimientoInformar = New System.Windows.Forms.Button()
        Me.bnt_FECAEASinMovimientoConsultar = New System.Windows.Forms.Button()
        Me.gpr_FECAEARegInformativo = New System.Windows.Forms.GroupBox()
        Me.ListBox_FECAEARegInformativo_CbteTipo = New System.Windows.Forms.ListBox()
        Me.lbl_FECAEARegInformativo_CbteTipo = New System.Windows.Forms.Label()
        Me.lbl_FECAEADetRequest = New System.Windows.Forms.Label()
        Me.btn_FECAEADetRequestFinEditar = New System.Windows.Forms.Button()
        Me.btn_FECAEADetRequestEditar = New System.Windows.Forms.Button()
        Me.btn_FECAEADetRequestBorrar = New System.Windows.Forms.Button()
        Me.btn_FECAEADetRequestNuevo = New System.Windows.Forms.Button()
        Me.dgv_FECAEADetRequest = New System.Windows.Forms.DataGridView()
        Me.DataGridViewComboBoxColumn1 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewComboBoxColumn2 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewComboBoxColumn3 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewTextBoxColumn14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CAEA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lbl_FECAEARegInformativo_CantReg = New System.Windows.Forms.Label()
        Me.txt_FECAEARegInformativo_CantReg = New System.Windows.Forms.TextBox()
        Me.btn_FECAEARegInformativo = New System.Windows.Forms.Button()
        Me.lbl_FECAEARegInformativo_PtoVta = New System.Windows.Forms.Label()
        Me.txt_FECAEARegInformativo_PtoVta = New System.Windows.Forms.TextBox()
        Me.grp_FEAuthRequest.SuspendLayout()
        Me.grp_FEDummy.SuspendLayout()
        Me.grp_MetodosDeConsultasGenericas.SuspendLayout()
        CType(Me.dgv_metodosDeConsultasGenericas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grp_FECompConsultar.SuspendLayout()
        Me.grp_FECompUltimoAutorizado.SuspendLayout()
        Me.grp_FECAESolicitar.SuspendLayout()
        CType(Me.dgv_FECAEDetRequest, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gpr_FECAEAConsultar.SuspendLayout()
        Me.grp_FECAEASolicitar.SuspendLayout()
        Me.gpr_FECAEASinMovimientoConsultar.SuspendLayout()
        Me.gpr_FECAEASinMovimientoInformar.SuspendLayout()
        Me.gpr_FECAEARegInformativo.SuspendLayout()
        CType(Me.dgv_FECAEADetRequest, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmb_cuit
        '
        Me.cmb_cuit.FormattingEnabled = True
        Me.cmb_cuit.Location = New System.Drawing.Point(41, 135)
        Me.cmb_cuit.Name = "cmb_cuit"
        Me.cmb_cuit.Size = New System.Drawing.Size(171, 21)
        Me.cmb_cuit.TabIndex = 14
        '
        'grp_FEAuthRequest
        '
        Me.grp_FEAuthRequest.Controls.Add(Me.cmb_cuit)
        Me.grp_FEAuthRequest.Controls.Add(Me.lbl_cuit)
        Me.grp_FEAuthRequest.Controls.Add(Me.lbl_sign)
        Me.grp_FEAuthRequest.Controls.Add(Me.lbl_token)
        Me.grp_FEAuthRequest.Controls.Add(Me.txt_sign)
        Me.grp_FEAuthRequest.Controls.Add(Me.txt_token)
        Me.grp_FEAuthRequest.Location = New System.Drawing.Point(585, 12)
        Me.grp_FEAuthRequest.Name = "grp_FEAuthRequest"
        Me.grp_FEAuthRequest.Size = New System.Drawing.Size(245, 168)
        Me.grp_FEAuthRequest.TabIndex = 104
        Me.grp_FEAuthRequest.TabStop = False
        Me.grp_FEAuthRequest.Text = "FEAuthRequest"
        '
        'lbl_cuit
        '
        Me.lbl_cuit.AutoSize = True
        Me.lbl_cuit.Location = New System.Drawing.Point(44, 119)
        Me.lbl_cuit.Name = "lbl_cuit"
        Me.lbl_cuit.Size = New System.Drawing.Size(24, 13)
        Me.lbl_cuit.TabIndex = 87
        Me.lbl_cuit.Text = "cuit"
        '
        'lbl_sign
        '
        Me.lbl_sign.AutoSize = True
        Me.lbl_sign.Location = New System.Drawing.Point(44, 70)
        Me.lbl_sign.Name = "lbl_sign"
        Me.lbl_sign.Size = New System.Drawing.Size(26, 13)
        Me.lbl_sign.TabIndex = 86
        Me.lbl_sign.Text = "sign"
        '
        'lbl_token
        '
        Me.lbl_token.AutoSize = True
        Me.lbl_token.Location = New System.Drawing.Point(44, 21)
        Me.lbl_token.Name = "lbl_token"
        Me.lbl_token.Size = New System.Drawing.Size(34, 13)
        Me.lbl_token.TabIndex = 85
        Me.lbl_token.Text = "token"
        '
        'txt_sign
        '
        Me.txt_sign.Location = New System.Drawing.Point(41, 86)
        Me.txt_sign.Name = "txt_sign"
        Me.txt_sign.Size = New System.Drawing.Size(171, 20)
        Me.txt_sign.TabIndex = 13
        '
        'txt_token
        '
        Me.txt_token.Location = New System.Drawing.Point(41, 37)
        Me.txt_token.Name = "txt_token"
        Me.txt_token.Size = New System.Drawing.Size(171, 20)
        Me.txt_token.TabIndex = 12
        '
        'grp_FEDummy
        '
        Me.grp_FEDummy.Controls.Add(Me.btn_FEDummy_DbServer)
        Me.grp_FEDummy.Controls.Add(Me.bnt_FEDummy_AuthServer)
        Me.grp_FEDummy.Controls.Add(Me.bnt_FEDummy_AppServer)
        Me.grp_FEDummy.Location = New System.Drawing.Point(7, 12)
        Me.grp_FEDummy.Name = "grp_FEDummy"
        Me.grp_FEDummy.Size = New System.Drawing.Size(190, 109)
        Me.grp_FEDummy.TabIndex = 105
        Me.grp_FEDummy.TabStop = False
        Me.grp_FEDummy.Text = "FEDummy"
        '
        'btn_FEDummy_DbServer
        '
        Me.btn_FEDummy_DbServer.Location = New System.Drawing.Point(9, 74)
        Me.btn_FEDummy_DbServer.Name = "btn_FEDummy_DbServer"
        Me.btn_FEDummy_DbServer.Size = New System.Drawing.Size(171, 23)
        Me.btn_FEDummy_DbServer.TabIndex = 3
        Me.btn_FEDummy_DbServer.Text = "FEDummy.DbServer"
        Me.btn_FEDummy_DbServer.UseVisualStyleBackColor = True
        '
        'bnt_FEDummy_AuthServer
        '
        Me.bnt_FEDummy_AuthServer.Location = New System.Drawing.Point(9, 45)
        Me.bnt_FEDummy_AuthServer.Name = "bnt_FEDummy_AuthServer"
        Me.bnt_FEDummy_AuthServer.Size = New System.Drawing.Size(171, 23)
        Me.bnt_FEDummy_AuthServer.TabIndex = 2
        Me.bnt_FEDummy_AuthServer.Text = "FEDummy.AuthServer"
        Me.bnt_FEDummy_AuthServer.UseVisualStyleBackColor = True
        '
        'bnt_FEDummy_AppServer
        '
        Me.bnt_FEDummy_AppServer.Location = New System.Drawing.Point(9, 16)
        Me.bnt_FEDummy_AppServer.Name = "bnt_FEDummy_AppServer"
        Me.bnt_FEDummy_AppServer.Size = New System.Drawing.Size(171, 23)
        Me.bnt_FEDummy_AppServer.TabIndex = 1
        Me.bnt_FEDummy_AppServer.Text = "FEDummy.AppServer"
        Me.bnt_FEDummy_AppServer.UseVisualStyleBackColor = True
        '
        'grp_MetodosDeConsultasGenericas
        '
        Me.grp_MetodosDeConsultasGenericas.Controls.Add(Me.dgv_metodosDeConsultasGenericas)
        Me.grp_MetodosDeConsultasGenericas.Controls.Add(Me.cbm_metodosDeconsultasGenericas)
        Me.grp_MetodosDeConsultasGenericas.Location = New System.Drawing.Point(7, 145)
        Me.grp_MetodosDeConsultasGenericas.Name = "grp_MetodosDeConsultasGenericas"
        Me.grp_MetodosDeConsultasGenericas.Size = New System.Drawing.Size(306, 265)
        Me.grp_MetodosDeConsultasGenericas.TabIndex = 107
        Me.grp_MetodosDeConsultasGenericas.TabStop = False
        Me.grp_MetodosDeConsultasGenericas.Text = "Métodos de consultas genéricas"
        '
        'dgv_metodosDeConsultasGenericas
        '
        Me.dgv_metodosDeConsultasGenericas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_metodosDeConsultasGenericas.Location = New System.Drawing.Point(7, 57)
        Me.dgv_metodosDeConsultasGenericas.Name = "dgv_metodosDeConsultasGenericas"
        Me.dgv_metodosDeConsultasGenericas.Size = New System.Drawing.Size(293, 195)
        Me.dgv_metodosDeConsultasGenericas.TabIndex = 94
        '
        'cbm_metodosDeconsultasGenericas
        '
        Me.cbm_metodosDeconsultasGenericas.FormattingEnabled = True
        Me.cbm_metodosDeconsultasGenericas.Items.AddRange(New Object() {"FECompTotXRequest", "FEParamGetPtosVenta", "FEParamGetTiposCbte", "FEParamGetTiposConcepto", "FEParamGetTiposDoc", "FEParamGetTiposIva", "FEParamGetTiposMonedas", "FEParamGetTiposOpcional", "FEParamGetTiposTributos"})
        Me.cbm_metodosDeconsultasGenericas.Location = New System.Drawing.Point(7, 19)
        Me.cbm_metodosDeconsultasGenericas.Name = "cbm_metodosDeconsultasGenericas"
        Me.cbm_metodosDeconsultasGenericas.Size = New System.Drawing.Size(260, 21)
        Me.cbm_metodosDeconsultasGenericas.TabIndex = 4
        '
        'grp_FECompConsultar
        '
        Me.grp_FECompConsultar.Controls.Add(Me.grp_FECompUltimoAutorizado)
        Me.grp_FECompConsultar.Controls.Add(Me.txt_FECompConsultar_CbteNro)
        Me.grp_FECompConsultar.Controls.Add(Me.lbl_FECompConsultar_CbteNro)
        Me.grp_FECompConsultar.Controls.Add(Me.btn_FECompConsultar)
        Me.grp_FECompConsultar.Location = New System.Drawing.Point(331, 12)
        Me.grp_FECompConsultar.Name = "grp_FECompConsultar"
        Me.grp_FECompConsultar.Size = New System.Drawing.Size(238, 398)
        Me.grp_FECompConsultar.TabIndex = 108
        Me.grp_FECompConsultar.TabStop = False
        Me.grp_FECompConsultar.Text = "FECompConsultar"
        '
        'grp_FECompUltimoAutorizado
        '
        Me.grp_FECompUltimoAutorizado.Controls.Add(Me.ListBox_FECompUltimoAutorizado_PtoVta)
        Me.grp_FECompUltimoAutorizado.Controls.Add(Me.lbl_FECompUltimoAutorizado_PtoVta)
        Me.grp_FECompUltimoAutorizado.Controls.Add(Me.ListBox_FECompUltimoAutorizado_CbteTipo)
        Me.grp_FECompUltimoAutorizado.Controls.Add(Me.btn_FECompUltimoAutorizado)
        Me.grp_FECompUltimoAutorizado.Controls.Add(Me.lbl_FECompUltimoAutorizado_CbteTipo)
        Me.grp_FECompUltimoAutorizado.Location = New System.Drawing.Point(13, 25)
        Me.grp_FECompUltimoAutorizado.Name = "grp_FECompUltimoAutorizado"
        Me.grp_FECompUltimoAutorizado.Size = New System.Drawing.Size(212, 254)
        Me.grp_FECompUltimoAutorizado.TabIndex = 92
        Me.grp_FECompUltimoAutorizado.TabStop = False
        Me.grp_FECompUltimoAutorizado.Text = "FECompUltimoAutorizado"
        '
        'ListBox_FECompUltimoAutorizado_PtoVta
        '
        Me.ListBox_FECompUltimoAutorizado_PtoVta.FormattingEnabled = True
        Me.ListBox_FECompUltimoAutorizado_PtoVta.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09"})
        Me.ListBox_FECompUltimoAutorizado_PtoVta.Location = New System.Drawing.Point(20, 112)
        Me.ListBox_FECompUltimoAutorizado_PtoVta.Name = "ListBox_FECompUltimoAutorizado_PtoVta"
        Me.ListBox_FECompUltimoAutorizado_PtoVta.Size = New System.Drawing.Size(171, 95)
        Me.ListBox_FECompUltimoAutorizado_PtoVta.TabIndex = 80
        '
        'lbl_FECompUltimoAutorizado_PtoVta
        '
        Me.lbl_FECompUltimoAutorizado_PtoVta.AutoSize = True
        Me.lbl_FECompUltimoAutorizado_PtoVta.Location = New System.Drawing.Point(23, 96)
        Me.lbl_FECompUltimoAutorizado_PtoVta.Name = "lbl_FECompUltimoAutorizado_PtoVta"
        Me.lbl_FECompUltimoAutorizado_PtoVta.Size = New System.Drawing.Size(39, 13)
        Me.lbl_FECompUltimoAutorizado_PtoVta.TabIndex = 79
        Me.lbl_FECompUltimoAutorizado_PtoVta.Text = "PtoVta"
        '
        'ListBox_FECompUltimoAutorizado_CbteTipo
        '
        Me.ListBox_FECompUltimoAutorizado_CbteTipo.FormattingEnabled = True
        Me.ListBox_FECompUltimoAutorizado_CbteTipo.Items.AddRange(New Object() {"01 Factura A", "02 Nota de Débito A", "03 Nota de Crédito A", "04 Recibos A", "05 Notas de Venta al contado A", "06 Factura B", "07 Nota de Débito B", "08 Nota de Crédito B", "09 Recibos B", "11 Factura C", "12 Nota de Débito C", "13 Nota de Crédito C", "15 Recibo C", "10 Notas de Venta al contado B", "34 Cbtes. A del Anexo I, Apartado A,inc.f),R.G.Nro. 1415", "35 Cbtes. B del Anexo I,Apartado A,inc. f),R.G. Nro. 1415", "39 Otros comprobantes A que cumplan con R.G.Nro. 1415", "40 Otros comprobantes B que cumplan con R.G.Nro. 1415", "60 Cta de Vta y Liquido prod. A", "61 Cta de Vta y Liquido prod. B", "63 Liquidacion A", "64 Liquidacion B"})
        Me.ListBox_FECompUltimoAutorizado_CbteTipo.Location = New System.Drawing.Point(20, 41)
        Me.ListBox_FECompUltimoAutorizado_CbteTipo.Name = "ListBox_FECompUltimoAutorizado_CbteTipo"
        Me.ListBox_FECompUltimoAutorizado_CbteTipo.Size = New System.Drawing.Size(171, 43)
        Me.ListBox_FECompUltimoAutorizado_CbteTipo.TabIndex = 78
        '
        'btn_FECompUltimoAutorizado
        '
        Me.btn_FECompUltimoAutorizado.Location = New System.Drawing.Point(20, 217)
        Me.btn_FECompUltimoAutorizado.Name = "btn_FECompUltimoAutorizado"
        Me.btn_FECompUltimoAutorizado.Size = New System.Drawing.Size(171, 23)
        Me.btn_FECompUltimoAutorizado.TabIndex = 82
        Me.btn_FECompUltimoAutorizado.Text = "FECompUltimoAutorizado"
        Me.btn_FECompUltimoAutorizado.UseVisualStyleBackColor = True
        '
        'lbl_FECompUltimoAutorizado_CbteTipo
        '
        Me.lbl_FECompUltimoAutorizado_CbteTipo.AutoSize = True
        Me.lbl_FECompUltimoAutorizado_CbteTipo.Location = New System.Drawing.Point(23, 25)
        Me.lbl_FECompUltimoAutorizado_CbteTipo.Name = "lbl_FECompUltimoAutorizado_CbteTipo"
        Me.lbl_FECompUltimoAutorizado_CbteTipo.Size = New System.Drawing.Size(50, 13)
        Me.lbl_FECompUltimoAutorizado_CbteTipo.TabIndex = 77
        Me.lbl_FECompUltimoAutorizado_CbteTipo.Text = "CbteTipo"
        '
        'txt_FECompConsultar_CbteNro
        '
        Me.txt_FECompConsultar_CbteNro.Location = New System.Drawing.Point(32, 308)
        Me.txt_FECompConsultar_CbteNro.Name = "txt_FECompConsultar_CbteNro"
        Me.txt_FECompConsultar_CbteNro.Size = New System.Drawing.Size(171, 20)
        Me.txt_FECompConsultar_CbteNro.TabIndex = 84
        '
        'lbl_FECompConsultar_CbteNro
        '
        Me.lbl_FECompConsultar_CbteNro.AutoSize = True
        Me.lbl_FECompConsultar_CbteNro.Location = New System.Drawing.Point(36, 292)
        Me.lbl_FECompConsultar_CbteNro.Name = "lbl_FECompConsultar_CbteNro"
        Me.lbl_FECompConsultar_CbteNro.Size = New System.Drawing.Size(46, 13)
        Me.lbl_FECompConsultar_CbteNro.TabIndex = 83
        Me.lbl_FECompConsultar_CbteNro.Text = "CbteNro"
        '
        'btn_FECompConsultar
        '
        Me.btn_FECompConsultar.Location = New System.Drawing.Point(33, 338)
        Me.btn_FECompConsultar.Name = "btn_FECompConsultar"
        Me.btn_FECompConsultar.Size = New System.Drawing.Size(171, 23)
        Me.btn_FECompConsultar.TabIndex = 85
        Me.btn_FECompConsultar.Text = "FECompConsultar"
        Me.btn_FECompConsultar.UseVisualStyleBackColor = True
        '
        'grp_FECAESolicitar
        '
        Me.grp_FECAESolicitar.Controls.Add(Me.ListBox_FECAESolicitar_CbteTipo)
        Me.grp_FECAESolicitar.Controls.Add(Me.lbl_FECAESolicitar_CbteTipo)
        Me.grp_FECAESolicitar.Controls.Add(Me.lbl_FECAEDetRequest)
        Me.grp_FECAESolicitar.Controls.Add(Me.btn_FECAEDetRequestFinEditar)
        Me.grp_FECAESolicitar.Controls.Add(Me.btn_FECAEDetRequestEditar)
        Me.grp_FECAESolicitar.Controls.Add(Me.btn_FECAEDetRequestBorrar)
        Me.grp_FECAESolicitar.Controls.Add(Me.btn_FECAEDetRequestNuevo)
        Me.grp_FECAESolicitar.Controls.Add(Me.dgv_FECAEDetRequest)
        Me.grp_FECAESolicitar.Controls.Add(Me.lbl_FECAESolicitar_CantReg)
        Me.grp_FECAESolicitar.Controls.Add(Me.txt_FECAESolicitar_CantReg)
        Me.grp_FECAESolicitar.Controls.Add(Me.btn_FECAESolicitar)
        Me.grp_FECAESolicitar.Controls.Add(Me.lbl_FECAESolicitar_PtoVta)
        Me.grp_FECAESolicitar.Controls.Add(Me.txt_FECAESolicitar_PtoVta)
        Me.grp_FECAESolicitar.Location = New System.Drawing.Point(7, 419)
        Me.grp_FECAESolicitar.Name = "grp_FECAESolicitar"
        Me.grp_FECAESolicitar.Size = New System.Drawing.Size(562, 277)
        Me.grp_FECAESolicitar.TabIndex = 159
        Me.grp_FECAESolicitar.TabStop = False
        Me.grp_FECAESolicitar.Text = "FECAESolicitar"
        '
        'ListBox_FECAESolicitar_CbteTipo
        '
        Me.ListBox_FECAESolicitar_CbteTipo.FormattingEnabled = True
        Me.ListBox_FECAESolicitar_CbteTipo.Items.AddRange(New Object() {"01 Factura A", "02 Nota de Débito A", "03 Nota de Crédito A", "04 Recibos A", "05 Notas de Venta al contado A", "06 Factura B", "07 Nota de Débito B", "08 Nota de Crédito B", "09 Recibos B", "11 Factura C", "12 Nota de Débito C", "13 Nota de Crédito C", "15 Recibo C", "10 Notas de Venta al contado B", "34 Cbtes. A del Anexo I, Apartado A,inc.f),R.G.Nro. 1415", "35 Cbtes. B del Anexo I,Apartado A,inc. f),R.G. Nro. 1415", "39 Otros comprobantes A que cumplan con R.G.Nro. 1415", "40 Otros comprobantes B que cumplan con R.G.Nro. 1415", "60 Cta de Vta y Liquido prod. A", "61 Cta de Vta y Liquido prod. B", "63 Liquidacion A", "64 Liquidacion B"})
        Me.ListBox_FECAESolicitar_CbteTipo.Location = New System.Drawing.Point(151, 36)
        Me.ListBox_FECAESolicitar_CbteTipo.Name = "ListBox_FECAESolicitar_CbteTipo"
        Me.ListBox_FECAESolicitar_CbteTipo.Size = New System.Drawing.Size(171, 43)
        Me.ListBox_FECAESolicitar_CbteTipo.TabIndex = 159
        '
        'lbl_FECAESolicitar_CbteTipo
        '
        Me.lbl_FECAESolicitar_CbteTipo.AutoSize = True
        Me.lbl_FECAESolicitar_CbteTipo.Location = New System.Drawing.Point(154, 20)
        Me.lbl_FECAESolicitar_CbteTipo.Name = "lbl_FECAESolicitar_CbteTipo"
        Me.lbl_FECAESolicitar_CbteTipo.Size = New System.Drawing.Size(50, 13)
        Me.lbl_FECAESolicitar_CbteTipo.TabIndex = 158
        Me.lbl_FECAESolicitar_CbteTipo.Text = "CbteTipo"
        '
        'lbl_FECAEDetRequest
        '
        Me.lbl_FECAEDetRequest.AutoSize = True
        Me.lbl_FECAEDetRequest.Location = New System.Drawing.Point(12, 73)
        Me.lbl_FECAEDetRequest.Name = "lbl_FECAEDetRequest"
        Me.lbl_FECAEDetRequest.Size = New System.Drawing.Size(98, 13)
        Me.lbl_FECAEDetRequest.TabIndex = 157
        Me.lbl_FECAEDetRequest.Text = "FECAEDetRequest"
        '
        'btn_FECAEDetRequestFinEditar
        '
        Me.btn_FECAEDetRequestFinEditar.Location = New System.Drawing.Point(285, 245)
        Me.btn_FECAEDetRequestFinEditar.Name = "btn_FECAEDetRequestFinEditar"
        Me.btn_FECAEDetRequestFinEditar.Size = New System.Drawing.Size(64, 23)
        Me.btn_FECAEDetRequestFinEditar.TabIndex = 57
        Me.btn_FECAEDetRequestFinEditar.Text = "Fin Editar"
        Me.btn_FECAEDetRequestFinEditar.UseVisualStyleBackColor = True
        '
        'btn_FECAEDetRequestEditar
        '
        Me.btn_FECAEDetRequestEditar.Location = New System.Drawing.Point(193, 245)
        Me.btn_FECAEDetRequestEditar.Name = "btn_FECAEDetRequestEditar"
        Me.btn_FECAEDetRequestEditar.Size = New System.Drawing.Size(64, 23)
        Me.btn_FECAEDetRequestEditar.TabIndex = 56
        Me.btn_FECAEDetRequestEditar.Text = "Editar"
        Me.btn_FECAEDetRequestEditar.UseVisualStyleBackColor = True
        '
        'btn_FECAEDetRequestBorrar
        '
        Me.btn_FECAEDetRequestBorrar.Location = New System.Drawing.Point(101, 245)
        Me.btn_FECAEDetRequestBorrar.Name = "btn_FECAEDetRequestBorrar"
        Me.btn_FECAEDetRequestBorrar.Size = New System.Drawing.Size(64, 23)
        Me.btn_FECAEDetRequestBorrar.TabIndex = 55
        Me.btn_FECAEDetRequestBorrar.Text = "Borrar"
        Me.btn_FECAEDetRequestBorrar.UseVisualStyleBackColor = True
        '
        'btn_FECAEDetRequestNuevo
        '
        Me.btn_FECAEDetRequestNuevo.Location = New System.Drawing.Point(9, 245)
        Me.btn_FECAEDetRequestNuevo.Name = "btn_FECAEDetRequestNuevo"
        Me.btn_FECAEDetRequestNuevo.Size = New System.Drawing.Size(64, 23)
        Me.btn_FECAEDetRequestNuevo.TabIndex = 54
        Me.btn_FECAEDetRequestNuevo.Text = "Nuevo"
        Me.btn_FECAEDetRequestNuevo.UseVisualStyleBackColor = True
        '
        'dgv_FECAEDetRequest
        '
        Me.dgv_FECAEDetRequest.AllowUserToAddRows = False
        Me.dgv_FECAEDetRequest.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_FECAEDetRequest.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Concepto, Me.DocTipo, Me.DocNro, Me.CbteDesde, Me.CbteHasta, Me.CbteFch, Me.ImpTotal, Me.ImpTotConc, Me.ImpNeto, Me.ImpOpEx, Me.ImpTrib, Me.ImpIVA, Me.FchServDesde, Me.FchServHasta, Me.FchVtoPago, Me.MonId, Me.MonCotiz})
        Me.dgv_FECAEDetRequest.Location = New System.Drawing.Point(9, 89)
        Me.dgv_FECAEDetRequest.Name = "dgv_FECAEDetRequest"
        Me.dgv_FECAEDetRequest.Size = New System.Drawing.Size(540, 150)
        Me.dgv_FECAEDetRequest.TabIndex = 53
        '
        'Concepto
        '
        Me.Concepto.HeaderText = "Concepto"
        Me.Concepto.Items.AddRange(New Object() {"01 Productos", "02 Servicios", "03 Productos y Servicios"})
        Me.Concepto.Name = "Concepto"
        Me.Concepto.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Concepto.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'DocTipo
        '
        Me.DocTipo.HeaderText = "DocTipo"
        Me.DocTipo.Items.AddRange(New Object() {"00 CI Policía Federal", "01 CI Buenos Aires", "02 CI Catamarca", "03 CI Córdoba", "04 CI Corrientes", "05 CI Entre Ríos", "06 CI Jujuy", "07 CI Mendoza", "08 CI La Rioja", "09 CI Salta", "10 CI San Juan", "11 CI San Luis", "12 CI Santa Fe", "13 CI Santiago del Estero", "14 CI Tucumán", "16 CI Chaco", "17 CI Chubut", "18 CI Formosa", "19 CI Misiones", "20 CI Neuquén", "21 CI La Pampa", "22 CI Río Negro", "23 CI Santa Cruz", "24 CI Tierra del Fuego", "80 CUIT", "86 CUIL", "87 CDI", "89 LE", "90 LC", "91 CI", "92 en trámite", "93 Acta Nacimiento", "95 CI Bs. As. RNP", "96 DNI", "94 Pasaporte", "99 Doc. (Otro)"})
        Me.DocTipo.Name = "DocTipo"
        '
        'DocNro
        '
        Me.DocNro.HeaderText = "DocNro"
        Me.DocNro.Name = "DocNro"
        '
        'CbteDesde
        '
        Me.CbteDesde.HeaderText = "CbteDesde"
        Me.CbteDesde.Name = "CbteDesde"
        '
        'CbteHasta
        '
        Me.CbteHasta.HeaderText = "CbteHasta"
        Me.CbteHasta.Name = "CbteHasta"
        Me.CbteHasta.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CbteHasta.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'CbteFch
        '
        Me.CbteFch.HeaderText = "CbteFch"
        Me.CbteFch.Name = "CbteFch"
        '
        'ImpTotal
        '
        Me.ImpTotal.HeaderText = "ImpTotal"
        Me.ImpTotal.Name = "ImpTotal"
        '
        'ImpTotConc
        '
        Me.ImpTotConc.HeaderText = "ImpTotConc"
        Me.ImpTotConc.Name = "ImpTotConc"
        '
        'ImpNeto
        '
        Me.ImpNeto.HeaderText = "ImpNeto"
        Me.ImpNeto.Name = "ImpNeto"
        '
        'ImpOpEx
        '
        Me.ImpOpEx.HeaderText = "ImpOpEx"
        Me.ImpOpEx.Name = "ImpOpEx"
        '
        'ImpTrib
        '
        Me.ImpTrib.HeaderText = "ImpTrib"
        Me.ImpTrib.Name = "ImpTrib"
        '
        'ImpIVA
        '
        Me.ImpIVA.HeaderText = "ImpIVA"
        Me.ImpIVA.Name = "ImpIVA"
        '
        'FchServDesde
        '
        Me.FchServDesde.HeaderText = "FchServDesde"
        Me.FchServDesde.Name = "FchServDesde"
        '
        'FchServHasta
        '
        Me.FchServHasta.HeaderText = "FchServHasta"
        Me.FchServHasta.Name = "FchServHasta"
        '
        'FchVtoPago
        '
        Me.FchVtoPago.HeaderText = "FchVtoPago"
        Me.FchVtoPago.Name = "FchVtoPago"
        '
        'MonId
        '
        Me.MonId.HeaderText = "MonId"
        Me.MonId.Items.AddRange(New Object() {"PES Pesos Argentinos"})
        Me.MonId.Name = "MonId"
        '
        'MonCotiz
        '
        Me.MonCotiz.HeaderText = "MonCotiz"
        Me.MonCotiz.Name = "MonCotiz"
        '
        'lbl_FECAESolicitar_CantReg
        '
        Me.lbl_FECAESolicitar_CantReg.AutoSize = True
        Me.lbl_FECAESolicitar_CantReg.Location = New System.Drawing.Point(12, 20)
        Me.lbl_FECAESolicitar_CantReg.Name = "lbl_FECAESolicitar_CantReg"
        Me.lbl_FECAESolicitar_CantReg.Size = New System.Drawing.Size(49, 13)
        Me.lbl_FECAESolicitar_CantReg.TabIndex = 131
        Me.lbl_FECAESolicitar_CantReg.Text = "CantReg"
        '
        'txt_FECAESolicitar_CantReg
        '
        Me.txt_FECAESolicitar_CantReg.Location = New System.Drawing.Point(9, 36)
        Me.txt_FECAESolicitar_CantReg.Name = "txt_FECAESolicitar_CantReg"
        Me.txt_FECAESolicitar_CantReg.Size = New System.Drawing.Size(60, 20)
        Me.txt_FECAESolicitar_CantReg.TabIndex = 41
        '
        'btn_FECAESolicitar
        '
        Me.btn_FECAESolicitar.Location = New System.Drawing.Point(377, 245)
        Me.btn_FECAESolicitar.Name = "btn_FECAESolicitar"
        Me.btn_FECAESolicitar.Size = New System.Drawing.Size(171, 23)
        Me.btn_FECAESolicitar.TabIndex = 58
        Me.btn_FECAESolicitar.Text = "FECAESolicitar"
        Me.btn_FECAESolicitar.UseVisualStyleBackColor = True
        '
        'lbl_FECAESolicitar_PtoVta
        '
        Me.lbl_FECAESolicitar_PtoVta.AutoSize = True
        Me.lbl_FECAESolicitar_PtoVta.Location = New System.Drawing.Point(88, 20)
        Me.lbl_FECAESolicitar_PtoVta.Name = "lbl_FECAESolicitar_PtoVta"
        Me.lbl_FECAESolicitar_PtoVta.Size = New System.Drawing.Size(39, 13)
        Me.lbl_FECAESolicitar_PtoVta.TabIndex = 110
        Me.lbl_FECAESolicitar_PtoVta.Text = "PtoVta"
        '
        'txt_FECAESolicitar_PtoVta
        '
        Me.txt_FECAESolicitar_PtoVta.Location = New System.Drawing.Point(85, 36)
        Me.txt_FECAESolicitar_PtoVta.Name = "txt_FECAESolicitar_PtoVta"
        Me.txt_FECAESolicitar_PtoVta.Size = New System.Drawing.Size(60, 20)
        Me.txt_FECAESolicitar_PtoVta.TabIndex = 32
        '
        'txt_Periodo
        '
        Me.txt_Periodo.Location = New System.Drawing.Point(23, 39)
        Me.txt_Periodo.Name = "txt_Periodo"
        Me.txt_Periodo.Size = New System.Drawing.Size(171, 20)
        Me.txt_Periodo.TabIndex = 161
        '
        'lbl_Periodo
        '
        Me.lbl_Periodo.AutoSize = True
        Me.lbl_Periodo.Location = New System.Drawing.Point(27, 23)
        Me.lbl_Periodo.Name = "lbl_Periodo"
        Me.lbl_Periodo.Size = New System.Drawing.Size(88, 13)
        Me.lbl_Periodo.TabIndex = 160
        Me.lbl_Periodo.Text = "Periodo (yyyymm)"
        '
        'btn_FECAEASolicitar
        '
        Me.btn_FECAEASolicitar.Location = New System.Drawing.Point(23, 113)
        Me.btn_FECAEASolicitar.Name = "btn_FECAEASolicitar"
        Me.btn_FECAEASolicitar.Size = New System.Drawing.Size(171, 23)
        Me.btn_FECAEASolicitar.TabIndex = 162
        Me.btn_FECAEASolicitar.Text = "FECAEASolicitar"
        Me.btn_FECAEASolicitar.UseVisualStyleBackColor = True
        '
        'cmb_Orden
        '
        Me.cmb_Orden.FormattingEnabled = True
        Me.cmb_Orden.Items.AddRange(New Object() {"1 Quincena", "2 Quincena"})
        Me.cmb_Orden.Location = New System.Drawing.Point(23, 78)
        Me.cmb_Orden.Name = "cmb_Orden"
        Me.cmb_Orden.Size = New System.Drawing.Size(171, 21)
        Me.cmb_Orden.TabIndex = 163
        '
        'lbl_Orden
        '
        Me.lbl_Orden.AutoSize = True
        Me.lbl_Orden.Location = New System.Drawing.Point(26, 62)
        Me.lbl_Orden.Name = "lbl_Orden"
        Me.lbl_Orden.Size = New System.Drawing.Size(91, 13)
        Me.lbl_Orden.TabIndex = 164
        Me.lbl_Orden.Text = "Orden (Quincena)"
        '
        'gpr_FECAEAConsultar
        '
        Me.gpr_FECAEAConsultar.Controls.Add(Me.grp_FECAEASolicitar)
        Me.gpr_FECAEAConsultar.Controls.Add(Me.btn_FECAEAConsultar)
        Me.gpr_FECAEAConsultar.Location = New System.Drawing.Point(589, 189)
        Me.gpr_FECAEAConsultar.Name = "gpr_FECAEAConsultar"
        Me.gpr_FECAEAConsultar.Size = New System.Drawing.Size(241, 221)
        Me.gpr_FECAEAConsultar.TabIndex = 165
        Me.gpr_FECAEAConsultar.TabStop = False
        Me.gpr_FECAEAConsultar.Text = "FECAEAConsultar"
        '
        'grp_FECAEASolicitar
        '
        Me.grp_FECAEASolicitar.Controls.Add(Me.cmb_Orden)
        Me.grp_FECAEASolicitar.Controls.Add(Me.lbl_Orden)
        Me.grp_FECAEASolicitar.Controls.Add(Me.txt_Periodo)
        Me.grp_FECAEASolicitar.Controls.Add(Me.lbl_Periodo)
        Me.grp_FECAEASolicitar.Controls.Add(Me.btn_FECAEASolicitar)
        Me.grp_FECAEASolicitar.Location = New System.Drawing.Point(15, 27)
        Me.grp_FECAEASolicitar.Name = "grp_FECAEASolicitar"
        Me.grp_FECAEASolicitar.Size = New System.Drawing.Size(213, 148)
        Me.grp_FECAEASolicitar.TabIndex = 166
        Me.grp_FECAEASolicitar.TabStop = False
        Me.grp_FECAEASolicitar.Text = "FECAEASolicitar"
        '
        'btn_FECAEAConsultar
        '
        Me.btn_FECAEAConsultar.Location = New System.Drawing.Point(37, 183)
        Me.btn_FECAEAConsultar.Name = "btn_FECAEAConsultar"
        Me.btn_FECAEAConsultar.Size = New System.Drawing.Size(171, 23)
        Me.btn_FECAEAConsultar.TabIndex = 165
        Me.btn_FECAEAConsultar.Text = "FECAEAConsultar"
        Me.btn_FECAEAConsultar.UseVisualStyleBackColor = True
        '
        'gpr_FECAEASinMovimientoConsultar
        '
        Me.gpr_FECAEASinMovimientoConsultar.Controls.Add(Me.gpr_FECAEASinMovimientoInformar)
        Me.gpr_FECAEASinMovimientoConsultar.Controls.Add(Me.bnt_FECAEASinMovimientoConsultar)
        Me.gpr_FECAEASinMovimientoConsultar.Location = New System.Drawing.Point(589, 418)
        Me.gpr_FECAEASinMovimientoConsultar.Name = "gpr_FECAEASinMovimientoConsultar"
        Me.gpr_FECAEASinMovimientoConsultar.Size = New System.Drawing.Size(241, 221)
        Me.gpr_FECAEASinMovimientoConsultar.TabIndex = 166
        Me.gpr_FECAEASinMovimientoConsultar.TabStop = False
        Me.gpr_FECAEASinMovimientoConsultar.Text = "FECAEASinMovimientoConsultar"
        '
        'gpr_FECAEASinMovimientoInformar
        '
        Me.gpr_FECAEASinMovimientoInformar.Controls.Add(Me.txt_FECAEASinMovimientoConsultar_CAEA)
        Me.gpr_FECAEASinMovimientoInformar.Controls.Add(Me.lbl_FECAEASinMovimientoConsultar_CAEA)
        Me.gpr_FECAEASinMovimientoInformar.Controls.Add(Me.txt_FECAEASinMovimientoConsultar_PtoVta)
        Me.gpr_FECAEASinMovimientoInformar.Controls.Add(Me.lbl_FECAEASinMovimientoConsultar_PtoVta)
        Me.gpr_FECAEASinMovimientoInformar.Controls.Add(Me.bnt_FECAEASinMovimientoInformar)
        Me.gpr_FECAEASinMovimientoInformar.Location = New System.Drawing.Point(15, 27)
        Me.gpr_FECAEASinMovimientoInformar.Name = "gpr_FECAEASinMovimientoInformar"
        Me.gpr_FECAEASinMovimientoInformar.Size = New System.Drawing.Size(213, 148)
        Me.gpr_FECAEASinMovimientoInformar.TabIndex = 166
        Me.gpr_FECAEASinMovimientoInformar.TabStop = False
        Me.gpr_FECAEASinMovimientoInformar.Text = "FECAEASinMovimientoInformar"
        '
        'txt_FECAEASinMovimientoConsultar_CAEA
        '
        Me.txt_FECAEASinMovimientoConsultar_CAEA.Location = New System.Drawing.Point(23, 78)
        Me.txt_FECAEASinMovimientoConsultar_CAEA.Name = "txt_FECAEASinMovimientoConsultar_CAEA"
        Me.txt_FECAEASinMovimientoConsultar_CAEA.Size = New System.Drawing.Size(171, 20)
        Me.txt_FECAEASinMovimientoConsultar_CAEA.TabIndex = 167
        '
        'lbl_FECAEASinMovimientoConsultar_CAEA
        '
        Me.lbl_FECAEASinMovimientoConsultar_CAEA.AutoSize = True
        Me.lbl_FECAEASinMovimientoConsultar_CAEA.Location = New System.Drawing.Point(26, 62)
        Me.lbl_FECAEASinMovimientoConsultar_CAEA.Name = "lbl_FECAEASinMovimientoConsultar_CAEA"
        Me.lbl_FECAEASinMovimientoConsultar_CAEA.Size = New System.Drawing.Size(35, 13)
        Me.lbl_FECAEASinMovimientoConsultar_CAEA.TabIndex = 164
        Me.lbl_FECAEASinMovimientoConsultar_CAEA.Text = "CAEA"
        '
        'txt_FECAEASinMovimientoConsultar_PtoVta
        '
        Me.txt_FECAEASinMovimientoConsultar_PtoVta.Location = New System.Drawing.Point(23, 39)
        Me.txt_FECAEASinMovimientoConsultar_PtoVta.Name = "txt_FECAEASinMovimientoConsultar_PtoVta"
        Me.txt_FECAEASinMovimientoConsultar_PtoVta.Size = New System.Drawing.Size(171, 20)
        Me.txt_FECAEASinMovimientoConsultar_PtoVta.TabIndex = 161
        '
        'lbl_FECAEASinMovimientoConsultar_PtoVta
        '
        Me.lbl_FECAEASinMovimientoConsultar_PtoVta.AutoSize = True
        Me.lbl_FECAEASinMovimientoConsultar_PtoVta.Location = New System.Drawing.Point(27, 23)
        Me.lbl_FECAEASinMovimientoConsultar_PtoVta.Name = "lbl_FECAEASinMovimientoConsultar_PtoVta"
        Me.lbl_FECAEASinMovimientoConsultar_PtoVta.Size = New System.Drawing.Size(39, 13)
        Me.lbl_FECAEASinMovimientoConsultar_PtoVta.TabIndex = 160
        Me.lbl_FECAEASinMovimientoConsultar_PtoVta.Text = "PtoVta"
        '
        'bnt_FECAEASinMovimientoInformar
        '
        Me.bnt_FECAEASinMovimientoInformar.Location = New System.Drawing.Point(23, 113)
        Me.bnt_FECAEASinMovimientoInformar.Name = "bnt_FECAEASinMovimientoInformar"
        Me.bnt_FECAEASinMovimientoInformar.Size = New System.Drawing.Size(171, 23)
        Me.bnt_FECAEASinMovimientoInformar.TabIndex = 162
        Me.bnt_FECAEASinMovimientoInformar.Text = "FECAEASinMovimientoInformar"
        Me.bnt_FECAEASinMovimientoInformar.UseVisualStyleBackColor = True
        '
        'bnt_FECAEASinMovimientoConsultar
        '
        Me.bnt_FECAEASinMovimientoConsultar.Location = New System.Drawing.Point(37, 183)
        Me.bnt_FECAEASinMovimientoConsultar.Name = "bnt_FECAEASinMovimientoConsultar"
        Me.bnt_FECAEASinMovimientoConsultar.Size = New System.Drawing.Size(171, 23)
        Me.bnt_FECAEASinMovimientoConsultar.TabIndex = 165
        Me.bnt_FECAEASinMovimientoConsultar.Text = "FECAEASinMovimientoConsultar"
        Me.bnt_FECAEASinMovimientoConsultar.UseVisualStyleBackColor = True
        '
        'gpr_FECAEARegInformativo
        '
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.ListBox_FECAEARegInformativo_CbteTipo)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.lbl_FECAEARegInformativo_CbteTipo)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.lbl_FECAEADetRequest)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.btn_FECAEADetRequestFinEditar)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.btn_FECAEADetRequestEditar)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.btn_FECAEADetRequestBorrar)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.btn_FECAEADetRequestNuevo)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.dgv_FECAEADetRequest)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.lbl_FECAEARegInformativo_CantReg)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.txt_FECAEARegInformativo_CantReg)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.btn_FECAEARegInformativo)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.lbl_FECAEARegInformativo_PtoVta)
        Me.gpr_FECAEARegInformativo.Controls.Add(Me.txt_FECAEARegInformativo_PtoVta)
        Me.gpr_FECAEARegInformativo.Location = New System.Drawing.Point(7, 701)
        Me.gpr_FECAEARegInformativo.Name = "gpr_FECAEARegInformativo"
        Me.gpr_FECAEARegInformativo.Size = New System.Drawing.Size(562, 280)
        Me.gpr_FECAEARegInformativo.TabIndex = 167
        Me.gpr_FECAEARegInformativo.TabStop = False
        Me.gpr_FECAEARegInformativo.Text = "FECAEARegInformativo"
        '
        'ListBox_FECAEARegInformativo_CbteTipo
        '
        Me.ListBox_FECAEARegInformativo_CbteTipo.FormattingEnabled = True
        Me.ListBox_FECAEARegInformativo_CbteTipo.Items.AddRange(New Object() {"01 Factura A", "02 Nota de Débito A", "03 Nota de Crédito A", "04 Recibos A", "05 Notas de Venta al contado A", "06 Factura B", "07 Nota de Débito B", "08 Nota de Crédito B", "09 Recibos B", "11 Factura C", "12 Nota de Débito C", "13 Nota de Crédito C", "15 Recibo C", "10 Notas de Venta al contado B", "34 Cbtes. A del Anexo I, Apartado A,inc.f),R.G.Nro. 1415", "35 Cbtes. B del Anexo I,Apartado A,inc. f),R.G. Nro. 1415", "39 Otros comprobantes A que cumplan con R.G.Nro. 1415", "40 Otros comprobantes B que cumplan con R.G.Nro. 1415", "60 Cta de Vta y Liquido prod. A", "61 Cta de Vta y Liquido prod. B", "63 Liquidacion A", "64 Liquidacion B"})
        Me.ListBox_FECAEARegInformativo_CbteTipo.Location = New System.Drawing.Point(151, 36)
        Me.ListBox_FECAEARegInformativo_CbteTipo.Name = "ListBox_FECAEARegInformativo_CbteTipo"
        Me.ListBox_FECAEARegInformativo_CbteTipo.Size = New System.Drawing.Size(171, 43)
        Me.ListBox_FECAEARegInformativo_CbteTipo.TabIndex = 159
        '
        'lbl_FECAEARegInformativo_CbteTipo
        '
        Me.lbl_FECAEARegInformativo_CbteTipo.AutoSize = True
        Me.lbl_FECAEARegInformativo_CbteTipo.Location = New System.Drawing.Point(154, 20)
        Me.lbl_FECAEARegInformativo_CbteTipo.Name = "lbl_FECAEARegInformativo_CbteTipo"
        Me.lbl_FECAEARegInformativo_CbteTipo.Size = New System.Drawing.Size(50, 13)
        Me.lbl_FECAEARegInformativo_CbteTipo.TabIndex = 158
        Me.lbl_FECAEARegInformativo_CbteTipo.Text = "CbteTipo"
        '
        'lbl_FECAEADetRequest
        '
        Me.lbl_FECAEADetRequest.AutoSize = True
        Me.lbl_FECAEADetRequest.Location = New System.Drawing.Point(12, 73)
        Me.lbl_FECAEADetRequest.Name = "lbl_FECAEADetRequest"
        Me.lbl_FECAEADetRequest.Size = New System.Drawing.Size(98, 13)
        Me.lbl_FECAEADetRequest.TabIndex = 157
        Me.lbl_FECAEADetRequest.Text = "FECAEDetRequest"
        '
        'btn_FECAEADetRequestFinEditar
        '
        Me.btn_FECAEADetRequestFinEditar.Location = New System.Drawing.Point(285, 245)
        Me.btn_FECAEADetRequestFinEditar.Name = "btn_FECAEADetRequestFinEditar"
        Me.btn_FECAEADetRequestFinEditar.Size = New System.Drawing.Size(64, 23)
        Me.btn_FECAEADetRequestFinEditar.TabIndex = 57
        Me.btn_FECAEADetRequestFinEditar.Text = "Fin Editar"
        Me.btn_FECAEADetRequestFinEditar.UseVisualStyleBackColor = True
        '
        'btn_FECAEADetRequestEditar
        '
        Me.btn_FECAEADetRequestEditar.Location = New System.Drawing.Point(193, 245)
        Me.btn_FECAEADetRequestEditar.Name = "btn_FECAEADetRequestEditar"
        Me.btn_FECAEADetRequestEditar.Size = New System.Drawing.Size(64, 23)
        Me.btn_FECAEADetRequestEditar.TabIndex = 56
        Me.btn_FECAEADetRequestEditar.Text = "Editar"
        Me.btn_FECAEADetRequestEditar.UseVisualStyleBackColor = True
        '
        'btn_FECAEADetRequestBorrar
        '
        Me.btn_FECAEADetRequestBorrar.Location = New System.Drawing.Point(101, 245)
        Me.btn_FECAEADetRequestBorrar.Name = "btn_FECAEADetRequestBorrar"
        Me.btn_FECAEADetRequestBorrar.Size = New System.Drawing.Size(64, 23)
        Me.btn_FECAEADetRequestBorrar.TabIndex = 55
        Me.btn_FECAEADetRequestBorrar.Text = "Borrar"
        Me.btn_FECAEADetRequestBorrar.UseVisualStyleBackColor = True
        '
        'btn_FECAEADetRequestNuevo
        '
        Me.btn_FECAEADetRequestNuevo.Location = New System.Drawing.Point(9, 245)
        Me.btn_FECAEADetRequestNuevo.Name = "btn_FECAEADetRequestNuevo"
        Me.btn_FECAEADetRequestNuevo.Size = New System.Drawing.Size(64, 23)
        Me.btn_FECAEADetRequestNuevo.TabIndex = 54
        Me.btn_FECAEADetRequestNuevo.Text = "Nuevo"
        Me.btn_FECAEADetRequestNuevo.UseVisualStyleBackColor = True
        '
        'dgv_FECAEADetRequest
        '
        Me.dgv_FECAEADetRequest.AllowUserToAddRows = False
        Me.dgv_FECAEADetRequest.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_FECAEADetRequest.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewComboBoxColumn1, Me.DataGridViewComboBoxColumn2, Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8, Me.DataGridViewTextBoxColumn9, Me.DataGridViewTextBoxColumn10, Me.DataGridViewTextBoxColumn11, Me.DataGridViewTextBoxColumn12, Me.DataGridViewTextBoxColumn13, Me.DataGridViewComboBoxColumn3, Me.DataGridViewTextBoxColumn14, Me.CAEA})
        Me.dgv_FECAEADetRequest.Location = New System.Drawing.Point(9, 89)
        Me.dgv_FECAEADetRequest.Name = "dgv_FECAEADetRequest"
        Me.dgv_FECAEADetRequest.Size = New System.Drawing.Size(540, 150)
        Me.dgv_FECAEADetRequest.TabIndex = 53
        '
        'DataGridViewComboBoxColumn1
        '
        Me.DataGridViewComboBoxColumn1.HeaderText = "Concepto"
        Me.DataGridViewComboBoxColumn1.Items.AddRange(New Object() {"01 Productos", "02 Servicios", "03 Productos y Servicios"})
        Me.DataGridViewComboBoxColumn1.Name = "DataGridViewComboBoxColumn1"
        Me.DataGridViewComboBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewComboBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'DataGridViewComboBoxColumn2
        '
        Me.DataGridViewComboBoxColumn2.HeaderText = "DocTipo"
        Me.DataGridViewComboBoxColumn2.Items.AddRange(New Object() {"00 CI Policía Federal", "01 CI Buenos Aires", "02 CI Catamarca", "03 CI Córdoba", "04 CI Corrientes", "05 CI Entre Ríos", "06 CI Jujuy", "07 CI Mendoza", "08 CI La Rioja", "09 CI Salta", "10 CI San Juan", "11 CI San Luis", "12 CI Santa Fe", "13 CI Santiago del Estero", "14 CI Tucumán", "16 CI Chaco", "17 CI Chubut", "18 CI Formosa", "19 CI Misiones", "20 CI Neuquén", "21 CI La Pampa", "22 CI Río Negro", "23 CI Santa Cruz", "24 CI Tierra del Fuego", "80 CUIT", "86 CUIL", "87 CDI", "89 LE", "90 LC", "91 CI", "92 en trámite", "93 Acta Nacimiento", "95 CI Bs. As. RNP", "96 DNI", "94 Pasaporte", "99 Doc. (Otro)"})
        Me.DataGridViewComboBoxColumn2.Name = "DataGridViewComboBoxColumn2"
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "DocNro"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "CbteDesde"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "CbteHasta"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "CbteFch"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "ImpTotal"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.HeaderText = "ImpTotConc"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.HeaderText = "ImpNeto"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.HeaderText = "ImpOpEx"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.HeaderText = "ImpTrib"
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.HeaderText = "ImpIVA"
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        '
        'DataGridViewTextBoxColumn11
        '
        Me.DataGridViewTextBoxColumn11.HeaderText = "FchServDesde"
        Me.DataGridViewTextBoxColumn11.Name = "DataGridViewTextBoxColumn11"
        '
        'DataGridViewTextBoxColumn12
        '
        Me.DataGridViewTextBoxColumn12.HeaderText = "FchServHasta"
        Me.DataGridViewTextBoxColumn12.Name = "DataGridViewTextBoxColumn12"
        '
        'DataGridViewTextBoxColumn13
        '
        Me.DataGridViewTextBoxColumn13.HeaderText = "FchVtoPago"
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        '
        'DataGridViewComboBoxColumn3
        '
        Me.DataGridViewComboBoxColumn3.HeaderText = "MonId"
        Me.DataGridViewComboBoxColumn3.Items.AddRange(New Object() {"PES Pesos Argentinos"})
        Me.DataGridViewComboBoxColumn3.Name = "DataGridViewComboBoxColumn3"
        '
        'DataGridViewTextBoxColumn14
        '
        Me.DataGridViewTextBoxColumn14.HeaderText = "MonCotiz"
        Me.DataGridViewTextBoxColumn14.Name = "DataGridViewTextBoxColumn14"
        '
        'CAEA
        '
        Me.CAEA.HeaderText = "CAEA"
        Me.CAEA.Name = "CAEA"
        '
        'lbl_FECAEARegInformativo_CantReg
        '
        Me.lbl_FECAEARegInformativo_CantReg.AutoSize = True
        Me.lbl_FECAEARegInformativo_CantReg.Location = New System.Drawing.Point(12, 20)
        Me.lbl_FECAEARegInformativo_CantReg.Name = "lbl_FECAEARegInformativo_CantReg"
        Me.lbl_FECAEARegInformativo_CantReg.Size = New System.Drawing.Size(49, 13)
        Me.lbl_FECAEARegInformativo_CantReg.TabIndex = 131
        Me.lbl_FECAEARegInformativo_CantReg.Text = "CantReg"
        '
        'txt_FECAEARegInformativo_CantReg
        '
        Me.txt_FECAEARegInformativo_CantReg.Location = New System.Drawing.Point(9, 36)
        Me.txt_FECAEARegInformativo_CantReg.Name = "txt_FECAEARegInformativo_CantReg"
        Me.txt_FECAEARegInformativo_CantReg.Size = New System.Drawing.Size(60, 20)
        Me.txt_FECAEARegInformativo_CantReg.TabIndex = 41
        '
        'btn_FECAEARegInformativo
        '
        Me.btn_FECAEARegInformativo.Location = New System.Drawing.Point(377, 245)
        Me.btn_FECAEARegInformativo.Name = "btn_FECAEARegInformativo"
        Me.btn_FECAEARegInformativo.Size = New System.Drawing.Size(171, 23)
        Me.btn_FECAEARegInformativo.TabIndex = 58
        Me.btn_FECAEARegInformativo.Text = "FECAEARegInformativo"
        Me.btn_FECAEARegInformativo.UseVisualStyleBackColor = True
        '
        'lbl_FECAEARegInformativo_PtoVta
        '
        Me.lbl_FECAEARegInformativo_PtoVta.AutoSize = True
        Me.lbl_FECAEARegInformativo_PtoVta.Location = New System.Drawing.Point(82, 20)
        Me.lbl_FECAEARegInformativo_PtoVta.Name = "lbl_FECAEARegInformativo_PtoVta"
        Me.lbl_FECAEARegInformativo_PtoVta.Size = New System.Drawing.Size(39, 13)
        Me.lbl_FECAEARegInformativo_PtoVta.TabIndex = 110
        Me.lbl_FECAEARegInformativo_PtoVta.Text = "PtoVta"
        '
        'txt_FECAEARegInformativo_PtoVta
        '
        Me.txt_FECAEARegInformativo_PtoVta.Location = New System.Drawing.Point(79, 36)
        Me.txt_FECAEARegInformativo_PtoVta.Name = "txt_FECAEARegInformativo_PtoVta"
        Me.txt_FECAEARegInformativo_PtoVta.Size = New System.Drawing.Size(60, 20)
        Me.txt_FECAEARegInformativo_PtoVta.TabIndex = 32
        '
        'wsfev1_cliente
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(842, 733)
        Me.Controls.Add(Me.gpr_FECAEARegInformativo)
        Me.Controls.Add(Me.gpr_FECAEASinMovimientoConsultar)
        Me.Controls.Add(Me.gpr_FECAEAConsultar)
        Me.Controls.Add(Me.grp_FECAESolicitar)
        Me.Controls.Add(Me.grp_FECompConsultar)
        Me.Controls.Add(Me.grp_MetodosDeConsultasGenericas)
        Me.Controls.Add(Me.grp_FEDummy)
        Me.Controls.Add(Me.grp_FEAuthRequest)
        Me.Name = "wsfev1_cliente"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cliente WS wsfev1 (Versión 13.04.04)"
        Me.grp_FEAuthRequest.ResumeLayout(False)
        Me.grp_FEAuthRequest.PerformLayout()
        Me.grp_FEDummy.ResumeLayout(False)
        Me.grp_MetodosDeConsultasGenericas.ResumeLayout(False)
        CType(Me.dgv_metodosDeConsultasGenericas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grp_FECompConsultar.ResumeLayout(False)
        Me.grp_FECompConsultar.PerformLayout()
        Me.grp_FECompUltimoAutorizado.ResumeLayout(False)
        Me.grp_FECompUltimoAutorizado.PerformLayout()
        Me.grp_FECAESolicitar.ResumeLayout(False)
        Me.grp_FECAESolicitar.PerformLayout()
        CType(Me.dgv_FECAEDetRequest, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gpr_FECAEAConsultar.ResumeLayout(False)
        Me.grp_FECAEASolicitar.ResumeLayout(False)
        Me.grp_FECAEASolicitar.PerformLayout()
        Me.gpr_FECAEASinMovimientoConsultar.ResumeLayout(False)
        Me.gpr_FECAEASinMovimientoInformar.ResumeLayout(False)
        Me.gpr_FECAEASinMovimientoInformar.PerformLayout()
        Me.gpr_FECAEARegInformativo.ResumeLayout(False)
        Me.gpr_FECAEARegInformativo.PerformLayout()
        CType(Me.dgv_FECAEADetRequest, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmb_cuit As System.Windows.Forms.ComboBox
    Friend WithEvents grp_FEAuthRequest As System.Windows.Forms.GroupBox
    Friend WithEvents lbl_cuit As System.Windows.Forms.Label
    Friend WithEvents lbl_sign As System.Windows.Forms.Label
    Friend WithEvents lbl_token As System.Windows.Forms.Label
    Friend WithEvents txt_sign As System.Windows.Forms.TextBox
    Friend WithEvents txt_token As System.Windows.Forms.TextBox
    Friend WithEvents grp_FEDummy As System.Windows.Forms.GroupBox
    Friend WithEvents btn_FEDummy_DbServer As System.Windows.Forms.Button
    Friend WithEvents bnt_FEDummy_AuthServer As System.Windows.Forms.Button
    Friend WithEvents bnt_FEDummy_AppServer As System.Windows.Forms.Button
    Friend WithEvents grp_MetodosDeConsultasGenericas As System.Windows.Forms.GroupBox
    Friend WithEvents dgv_metodosDeConsultasGenericas As System.Windows.Forms.DataGridView
    Friend WithEvents cbm_metodosDeconsultasGenericas As System.Windows.Forms.ComboBox
    Friend WithEvents grp_FECompConsultar As System.Windows.Forms.GroupBox
    Friend WithEvents grp_FECompUltimoAutorizado As System.Windows.Forms.GroupBox
    Friend WithEvents ListBox_FECompUltimoAutorizado_PtoVta As System.Windows.Forms.ListBox
    Friend WithEvents lbl_FECompUltimoAutorizado_PtoVta As System.Windows.Forms.Label
    Friend WithEvents ListBox_FECompUltimoAutorizado_CbteTipo As System.Windows.Forms.ListBox
    Friend WithEvents btn_FECompUltimoAutorizado As System.Windows.Forms.Button
    Friend WithEvents lbl_FECompUltimoAutorizado_CbteTipo As System.Windows.Forms.Label
    Friend WithEvents txt_FECompConsultar_CbteNro As System.Windows.Forms.TextBox
    Friend WithEvents lbl_FECompConsultar_CbteNro As System.Windows.Forms.Label
    Friend WithEvents btn_FECompConsultar As System.Windows.Forms.Button
    Friend WithEvents grp_FECAESolicitar As System.Windows.Forms.GroupBox
    Friend WithEvents lbl_FECAEDetRequest As System.Windows.Forms.Label
    Friend WithEvents btn_FECAEDetRequestFinEditar As System.Windows.Forms.Button
    Friend WithEvents btn_FECAEDetRequestEditar As System.Windows.Forms.Button
    Friend WithEvents btn_FECAEDetRequestBorrar As System.Windows.Forms.Button
    Friend WithEvents btn_FECAEDetRequestNuevo As System.Windows.Forms.Button
    Friend WithEvents dgv_FECAEDetRequest As System.Windows.Forms.DataGridView
    Friend WithEvents lbl_FECAESolicitar_CantReg As System.Windows.Forms.Label
    Friend WithEvents txt_FECAESolicitar_CantReg As System.Windows.Forms.TextBox
    Friend WithEvents btn_FECAESolicitar As System.Windows.Forms.Button
    Friend WithEvents lbl_FECAESolicitar_PtoVta As System.Windows.Forms.Label
    Friend WithEvents txt_FECAESolicitar_PtoVta As System.Windows.Forms.TextBox
    Friend WithEvents ListBox_FECAESolicitar_CbteTipo As System.Windows.Forms.ListBox
    Friend WithEvents lbl_FECAESolicitar_CbteTipo As System.Windows.Forms.Label
    Friend WithEvents Concepto As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DocTipo As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DocNro As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CbteDesde As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CbteHasta As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CbteFch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ImpTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ImpTotConc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ImpNeto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ImpOpEx As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ImpTrib As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ImpIVA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FchServDesde As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FchServHasta As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FchVtoPago As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MonId As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents MonCotiz As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txt_Periodo As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Periodo As System.Windows.Forms.Label
    Friend WithEvents btn_FECAEASolicitar As System.Windows.Forms.Button
    Friend WithEvents cmb_Orden As System.Windows.Forms.ComboBox
    Friend WithEvents lbl_Orden As System.Windows.Forms.Label
    Friend WithEvents gpr_FECAEAConsultar As System.Windows.Forms.GroupBox
    Friend WithEvents btn_FECAEAConsultar As System.Windows.Forms.Button
    Friend WithEvents grp_FECAEASolicitar As System.Windows.Forms.GroupBox
    Friend WithEvents gpr_FECAEASinMovimientoConsultar As System.Windows.Forms.GroupBox
    Friend WithEvents gpr_FECAEASinMovimientoInformar As System.Windows.Forms.GroupBox
    Friend WithEvents txt_FECAEASinMovimientoConsultar_CAEA As System.Windows.Forms.TextBox
    Friend WithEvents lbl_FECAEASinMovimientoConsultar_CAEA As System.Windows.Forms.Label
    Friend WithEvents txt_FECAEASinMovimientoConsultar_PtoVta As System.Windows.Forms.TextBox
    Friend WithEvents lbl_FECAEASinMovimientoConsultar_PtoVta As System.Windows.Forms.Label
    Friend WithEvents bnt_FECAEASinMovimientoInformar As System.Windows.Forms.Button
    Friend WithEvents bnt_FECAEASinMovimientoConsultar As System.Windows.Forms.Button
    Friend WithEvents gpr_FECAEARegInformativo As System.Windows.Forms.GroupBox
    Friend WithEvents ListBox_FECAEARegInformativo_CbteTipo As System.Windows.Forms.ListBox
    Friend WithEvents lbl_FECAEARegInformativo_CbteTipo As System.Windows.Forms.Label
    Friend WithEvents lbl_FECAEADetRequest As System.Windows.Forms.Label
    Friend WithEvents btn_FECAEADetRequestFinEditar As System.Windows.Forms.Button
    Friend WithEvents btn_FECAEADetRequestEditar As System.Windows.Forms.Button
    Friend WithEvents btn_FECAEADetRequestBorrar As System.Windows.Forms.Button
    Friend WithEvents btn_FECAEADetRequestNuevo As System.Windows.Forms.Button
    Friend WithEvents dgv_FECAEADetRequest As System.Windows.Forms.DataGridView
    Friend WithEvents lbl_FECAEARegInformativo_CantReg As System.Windows.Forms.Label
    Friend WithEvents txt_FECAEARegInformativo_CantReg As System.Windows.Forms.TextBox
    Friend WithEvents btn_FECAEARegInformativo As System.Windows.Forms.Button
    Friend WithEvents lbl_FECAEARegInformativo_PtoVta As System.Windows.Forms.Label
    Friend WithEvents txt_FECAEARegInformativo_PtoVta As System.Windows.Forms.TextBox
    Friend WithEvents DataGridViewComboBoxColumn1 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn2 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn3 As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CAEA As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
