Imports System.Xml.Serialization
Imports System.IO
Imports Helper
Imports LibGenerales

Public Class WSFE
    Property objWSFEV1 As New wsfev1.Service
    Property RutaTicketAcceso As String
    Property strUrlServicio As String
    Property strUrlWsfeWsdl As String
    Property RutaComprobantesGenerados As String
    Property RutaComprobantesConsultados As String
    Public Property FEAuthRequest As New wsfev1.FEAuthRequest
    Property Debug As Boolean = CBool(RegEdit.ObtenerRegistro(eCategorias.WSFE, "ModoDebug"))
    Property Eventos As List(Of KeyValuePair(Of Integer, String))
    Property Errores As List(Of KeyValuePair(Of Integer, String))
    Property Observaciones As New List(Of KeyValuePair(Of Integer, String))
    Dim _ModoProduccion As Boolean
    Property ModoProduccion As Boolean
        Get
            Return _ModoProduccion
        End Get
        Set(value As Boolean)
            _ModoProduccion = value
            If value Then
                strUrlServicio = RegEdit.ObtenerRegistro(eCategorias.WSFE, "URL_WSFE_Produccion")
            Else
                strUrlServicio = RegEdit.ObtenerRegistro(eCategorias.WSFE, "URL_WSFE_Testing")
            End If
            strUrlWsfeWsdl = strUrlServicio & "?WSDL"
            objWSFEV1.Url = strUrlServicio
        End Set
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal Token As String, ByVal Sign As String, ByVal CUIT As String)
        FEAuthRequest.Token = Token
        FEAuthRequest.Sign = Sign
        FEAuthRequest.Cuit = CLng(CUIT)
        RutaTicketAcceso = RegEdit.ObtenerRegistro(eCategorias.WSFE, "RUTATICKETACCESO") & "TA.xml"      ' Debe indicar la Ruta de su Ticket de acceso
        RutaComprobantesGenerados = RegEdit.ObtenerRegistro(eCategorias.WSFE, "RutaComprobantesGenerados")
        RutaComprobantesConsultados = RegEdit.ObtenerRegistro(eCategorias.WSFE, "RutaComprobantesConsultados")
        ModoProduccion = CBool(RegEdit.ObtenerRegistro(eCategorias.WSFE, "ModoProduccion"))
    End Sub

    Public Function FECompUltimoAutorizado(ByVal pPtoVta As Integer, ByVal TipoCbte As Integer) As wsfev1.FERecuperaLastCbteResponse
        Dim objFERecuperaLastCbteResponse As wsfev1.FERecuperaLastCbteResponse
        ' Invoco al método FECompUltimoAutorizado
        Try
            objFERecuperaLastCbteResponse = objWSFEV1.FECompUltimoAutorizado(FEAuthRequest, pPtoVta, TipoCbte)
            If objFERecuperaLastCbteResponse IsNot Nothing Then
                ' MsgBox("error" & Trim(objFERecuperaLastCbteResponse.Errors(0).Msg))
                Dim ultimoComprobante As New wsfev1.FECompConsultaResponse
                Dim consultaultimo As New wsfev1.FECompConsultaReq
                consultaultimo.CbteNro = objFERecuperaLastCbteResponse.CbteNro
                consultaultimo.CbteTipo = objFERecuperaLastCbteResponse.CbteTipo
                consultaultimo.PtoVta = objFERecuperaLastCbteResponse.PtoVta
                ultimoComprobante = objWSFEV1.FECompConsultar(FEAuthRequest, consultaultimo)

                If objFERecuperaLastCbteResponse.Events Is Nothing _
                    And objFERecuperaLastCbteResponse.Errors Is Nothing Then Return objFERecuperaLastCbteResponse
            End If
            RegistroErroresEventos(objFERecuperaLastCbteResponse)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
        End Try
        Return Nothing
    End Function

    Public Function FECompConsultar(ByVal PtoVta As String, ByVal TipoCbte As String, ByVal CbteNro As String) As wsfev1.FECompConsultaResponse
        Dim objFECompConsultaReq As New wsfev1.FECompConsultaReq
        Dim objFECompConsultaResponse As wsfev1.FECompConsultaResponse
        objFECompConsultaReq.PtoVta = PtoVta
        objFECompConsultaReq.CbteTipo = TipoCbte
        objFECompConsultaReq.CbteNro = CbteNro

        ' Invoco al método FECompConsultar
        Try
            objFECompConsultaResponse = objWSFEV1.FECompConsultar(FEAuthRequest, objFECompConsultaReq)
            If objFECompConsultaResponse IsNot Nothing Then
                'Serialize object to a text file.
                Dim ruta As String = RutaComprobantesConsultados & DateTime.Now.ToString & "Comprobante.xml"
                Dim objStreamWriter As New StreamWriter(ruta)
                Dim x As New XmlSerializer(objFECompConsultaResponse.GetType)
                x.Serialize(objStreamWriter, objFECompConsultaResponse)
                objStreamWriter.Close()
                If objFECompConsultaResponse.Events Is Nothing And objFECompConsultaResponse.Errors Is Nothing Then Return objFECompConsultaResponse
            End If
            RegistroErroresEventos(objFECompConsultaResponse)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
        End Try
        Return Nothing
    End Function

    Public Function FECAESolicitar(ByVal objFECAERequest As wsfev1.FECAERequest) As wsfev1.FECAEResponse
        Dim objFECAECabRequest As wsfev1.FECAECabRequest = objFECAERequest.FeCabReq
        Dim objFECAEResponse As New wsfev1.FECAEResponse
        Dim arrayFECAEDetRequest(objFECAERequest.FeCabReq.CantReg) As wsfev1.FECAEDetRequest
        arrayFECAEDetRequest = objFECAERequest.FeDetReq
        Try
            objFECAEResponse = objWSFEV1.FECAESolicitar(FEAuthRequest, objFECAERequest)
            If objFECAEResponse IsNot Nothing Then
                'Convertir Objeto de respuesta en archivo de texto.
                'Formato del Nombre del Archivo:
                'Tipo de Comprobante + Punto de Venta + Numero de Comprobante + -CAE + Numero de CAE + .xml
                'Ej. A00100000001CAE0000000001.xml

                Dim NombreArchivo As String
                NombreArchivo = "RES" & RutaComprobantesGenerados & objFECAEResponse.FeCabResp.CbteTipo _
                                      & objFECAEResponse.FeCabResp.PtoVta & objFECAEResponse.FeDetResp(0).CbteDesde _
                                      & "CAE" & objFECAEResponse.FeDetResp(0).CAE & ".xml"
                'TODO: Aca se debe controlar que los directorios existan, en el caso de no exister, estos deben ser creados segun la configuracion.
                Dim objStreamWriter As New StreamWriter(NombreArchivo)
                Dim Res As New XmlSerializer(objFECAEResponse.GetType)
                Res.Serialize(objStreamWriter, objFECAEResponse)
                objStreamWriter.Close()

                NombreArchivo = "REQ" & RutaComprobantesGenerados & objFECAEResponse.FeCabResp.CbteTipo _
                                      & objFECAEResponse.FeCabResp.PtoVta & objFECAEResponse.FeDetResp(0).CbteDesde _
                                      & "CAE" & objFECAEResponse.FeDetResp(0).CAE & ".xml"
                objStreamWriter = Nothing
                objStreamWriter = New StreamWriter(NombreArchivo)
                Dim Req As New XmlSerializer(objFECAEResponse.GetType)
                Req.Serialize(objStreamWriter, objFECAERequest)
                objStreamWriter.Close()

                For i = 0 To (objFECAEResponse.FeDetResp.Length - 1)
                    If objFECAEResponse.FeDetResp(i).Observaciones Is Nothing Then Exit For
                    If objFECAEResponse.FeDetResp(i).Observaciones.Length <> 0 Then
                        'Registrar Observaciones.
                        For Each oObservacion As wsfev1.Obs In objFECAEResponse.FeDetResp(i).Observaciones
                            Observaciones.Add(New KeyValuePair(Of Integer, String)(oObservacion.Code, oObservacion.Msg))
                        Next
                    End If
                Next
                If objFECAEResponse.Events Is Nothing And objFECAEResponse.Errors Is Nothing Then Return objFECAEResponse
            End If
            RegistroErroresEventos(objFECAEResponse)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
        End Try
        Return Nothing
    End Function

    Public Function FECAESolicitar(ByVal CbteTipo As Integer,
                                   ByVal PtoVenta As Integer, ByVal Concepto As Integer,
                                   ByVal DocTipo As Integer, ByVal DocNro As Long,
                                   ByVal NroCbte As Long, ByVal CbteFch As String,
                                   ByVal ImpTotal As Double, ByVal ImpTotConc As Double, ByVal ImpNeto As Double,
                                   ByVal ImpOpEx As Double, ByVal ImpTrib As Double, ByVal ImpIVA As Double,
                                   ByVal FchServDesde As String, ByVal FchServHasta As String,
                                   ByVal FchVtoPago As String, ByVal MonId As String, ByVal MonCotiz As Double
                                   ) As wsfev1.FECAEResponse
        Dim Detalle As New DataTable
        Detalle.Columns.Add("Concepto", GetType(Integer))
        Detalle.Columns.Add("DocTipo", GetType(Integer))
        Detalle.Columns.Add("DocNro", GetType(Long))
        Detalle.Columns.Add("CbteDesde", GetType(Long))
        Detalle.Columns.Add("CbteHasta", GetType(Long))
        Detalle.Columns.Add("CbteFch", GetType(String))
        Detalle.Columns.Add("ImpTotal", GetType(Double))
        Detalle.Columns.Add("ImpTotConc", GetType(Double))
        Detalle.Columns.Add("ImpNeto", GetType(Double))
        Detalle.Columns.Add("ImpOpEx", GetType(Double))
        Detalle.Columns.Add("ImpTrib", GetType(Double))
        Detalle.Columns.Add("ImpIVA", GetType(Double))
        Detalle.Columns.Add("FchServDesde", GetType(String))
        Detalle.Columns.Add("FchServHasta", GetType(String))
        Detalle.Columns.Add("FchVtoPago", GetType(String))
        Detalle.Columns.Add("MonId", GetType(String))
        Detalle.Columns.Add("MonCotiz", GetType(Double))
        Detalle.Rows.Add({Concepto, DocTipo, DocNro, NroCbte, NroCbte, CbteFch, ImpTotal,
                          ImpTotConc, ImpNeto, ImpOpEx, ImpTrib, ImpIVA, FchServDesde, FchServHasta,
                          FchVtoPago, MonId, MonCotiz})
        Return FECAESolicitar(1, CbteTipo, PtoVenta, Detalle)
    End Function

    Public Function FECAESolicitar(ByVal CantReg As Integer, ByVal CbteTipo As Integer, ByVal PtoVenta As Integer, ByVal Detalle As DataTable) As wsfev1.FECAEResponse
        Dim objFECAECabRequest As New wsfev1.FECAECabRequest
        Dim objFECAERequest As New wsfev1.FECAERequest
        Dim objFECAEResponse As New wsfev1.FECAEResponse

        Dim indicemax_arrayFECAEDetRequest As Integer = Detalle.Rows.Count - 1
        Dim d_arrayFECAEDetRequest As Integer = 0
        Dim arrayFECAEDetRequest(indicemax_arrayFECAEDetRequest) As wsfev1.FECAEDetRequest

        objFECAECabRequest.CantReg = CantReg
        objFECAECabRequest.CbteTipo = CbteTipo
        objFECAECabRequest.PtoVta = PtoVenta

        For d_arrayFECAEDetRequest = 0 To (indicemax_arrayFECAEDetRequest)
            Dim objFECAEDetRequest As New wsfev1.FECAEDetRequest
            With objFECAEDetRequest
                .Concepto = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("Concepto"), Integer)
                .DocTipo = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("DocTipo"), Integer)
                .DocNro = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("DocNro"), Long)
                .CbteDesde = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("CbteDesde"), Long)
                .CbteHasta = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("CbteHasta"), Long)
                .CbteFch = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("CbteFch"), String)
                .ImpTotal = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("ImpTotal"), Double)
                .ImpTotConc = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("ImpTotConc"), Double)
                .ImpNeto = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("ImpNeto"), Double)
                .ImpOpEx = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("ImpOpEx"), Double)
                .ImpTrib = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("ImpTrib"), Double)
                .ImpIVA = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("ImpIVA"), Double)
                .FchServDesde = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("FchServDesde"), String)
                .FchServHasta = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("FchServHasta"), String)
                .FchVtoPago = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("FchVtoPago"), String)
                .MonId = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("MonId"), String)
                .MonCotiz = CType(Detalle.Rows(d_arrayFECAEDetRequest).Item("MonCotiz"), Double)
                Dim AlicuotaIva As New DataTable
                For i = 0 To AlicuotaIva.Rows.Count
                    Dim AlicIva As New wsfev1.AlicIva()
                    AlicIva.Id = CType(AlicuotaIva.Rows(i).Item("Id"), Integer)
                    AlicIva.BaseImp = CType(AlicuotaIva.Rows(i).Item("BaseImp"), Double)
                    AlicIva.Importe = CType(AlicuotaIva.Rows(i).Item("Importe"), Double)
                    .Iva(i) = AlicIva
                Next
            End With
            arrayFECAEDetRequest(d_arrayFECAEDetRequest) = objFECAEDetRequest
        Next d_arrayFECAEDetRequest

        ' Invoco al método FECAESolicitar
        Try
            objFECAEResponse = FECAESolicitar(objFECAERequest)
            Return objFECAEResponse
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
        End Try
        Return Nothing
    End Function

    Private Sub RegistroErroresEventos(ByVal pObjeto As Object)
        LimpiarRegistros()
        'Dim wObjeto As Object
        'wObjeto = CType(pObjeto, IErrorEvento)
        If pObjeto.Errors IsNot Nothing Then
            For i = 0 To pObjeto.Errors.Length - 1
                Errores.Add(New KeyValuePair(Of Integer, String)(pObjeto.Errors(i).Code, pObjeto.Errors(i).Msg))
            Next
        End If
        If pObjeto.Events IsNot Nothing Then
            For i = 0 To pObjeto.Events.Length - 1
                Eventos.Add(New KeyValuePair(Of Integer, String)(pObjeto.Events(i).Code, pObjeto.Events(i).Msg))
            Next
        End If

    End Sub

    Public Interface IErrorEvento
        Property Errors As wsfev1.Err
        Property Events As wsfev1.Evt
    End Interface

#Region "Metodos Dummy"
    Private Sub Dummy_AppServer()
        Dim objDummy As New wsfev1.DummyResponse
        ' Invoco al método Dummy
        Try
            objDummy = objWSFEV1.FEDummy
            MessageBox.Show(objDummy.AppServer, "FEDummy.AppServer")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub FEDummy_AuthServer()
        Dim objDummy As New wsfev1.DummyResponse
        ' Invoco al método Dummy
        Try
            objDummy = objWSFEV1.FEDummy
            MessageBox.Show(objDummy.AuthServer, "FEDummy.AuthServer")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub FEDummy_DbServer()
        Dim objDummy As New wsfev1.DummyResponse
        ' Invoco al método Dummy
        Try
            objDummy = objWSFEV1.FEDummy
            MessageBox.Show(objDummy.DbServer, "FEDummy.DbServer")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region
#Region "Consultas Genericas"
    Function ConsultaGenerica(ByVal valor As String) As Object
        Dim vRet As Object = Nothing
        Select Case valor
            Case "FECompTotXRequest"
                ' Invoco al método FECompTotXRequest
                Dim objFERegXReqResponse As wsfev1.FERegXReqResponse
                Try
                    objFERegXReqResponse = objWSFEV1.FECompTotXRequest(FEAuthRequest)
                    If objFERegXReqResponse.RegXReq.ToString IsNot Nothing Then
                        MessageBox.Show("objFERegXReqResponse.RegXReq: " + objFERegXReqResponse.RegXReq.ToString + vbCrLf, "objFERegXReqResponse.RegXReq")
                    End If
                    If objFERegXReqResponse.Errors IsNot Nothing Then
                        For i = 0 To objFERegXReqResponse.Errors.Length - 1
                            MessageBox.Show("objFERegXReqResponse.Errors(i).Code: " + objFERegXReqResponse.Errors(i).Code.ToString + vbCrLf +
                            "objFERegXReqResponse.Errors(i).Msg: " + objFERegXReqResponse.Errors(i).Msg)
                        Next
                    End If
                    If objFERegXReqResponse.Events IsNot Nothing Then
                        For i = 0 To objFERegXReqResponse.Events.Length - 1
                            MessageBox.Show("objFERegXReqResponse.Events(i).Code: " + objFERegXReqResponse.Events(i).Code.ToString + vbCrLf +
                            "objFERegXReqResponse.Events(i).Msg: " + objFERegXReqResponse.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "FEParamGetPtosVenta"
                ' Invoco al método FEParamGetPtosVenta
                Dim objFEPtoVenta As wsfev1.FEPtoVentaResponse
                Try
                    objFEPtoVenta = objWSFEV1.FEParamGetPtosVenta(FEAuthRequest)
                    vRet = objFEPtoVenta.ResultGet
                    If objFEPtoVenta.Errors IsNot Nothing Then
                        For i = 0 To objFEPtoVenta.Errors.Length - 1
                            MessageBox.Show("objFEPtoVenta.Errors(i).Code: " + objFEPtoVenta.Errors(i).Code.ToString + vbCrLf +
                            "objFEPtoVenta.Errors(i).Msg: " + objFEPtoVenta.Errors(i).Msg)
                        Next
                    End If
                    If objFEPtoVenta.Events IsNot Nothing Then
                        For i = 0 To objFEPtoVenta.Events.Length - 1
                            MessageBox.Show("objFEPtoVenta.Events(i).Code: " + objFEPtoVenta.Events(i).Code.ToString + vbCrLf +
                            "objFEPtoVenta.Events(i).Msg: " + objFEPtoVenta.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "FEParamGetTiposCbte"
                ' Invoco al método FEParamGetTiposCbte
                Dim objCbteTipo As wsfev1.CbteTipoResponse
                Try
                    objCbteTipo = objWSFEV1.FEParamGetTiposCbte(FEAuthRequest)
                    vRet = objCbteTipo.ResultGet
                    If objCbteTipo.Errors IsNot Nothing Then
                        For i = 0 To objCbteTipo.Errors.Length - 1
                            MessageBox.Show("objCbteTipoResponse.Errors(i).Code: " + objCbteTipo.Errors(i).Code.ToString + vbCrLf +
                            "objCbteTipoResponse.Errors(i).Msg: " + objCbteTipo.Errors(i).Msg)
                        Next
                    End If
                    If objCbteTipo.Events IsNot Nothing Then
                        For i = 0 To objCbteTipo.Events.Length - 1
                            MessageBox.Show("objCbteTipoResponse.Events(i).Code: " + objCbteTipo.Events(i).Code.ToString + vbCrLf +
                            "objCbteTipoResponse.Events(i).Msg: " + objCbteTipo.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "FEParamGetTiposConcepto"
                ' Invoco al método FEParamGetTiposConcepto
                Dim objConceptoTipo As wsfev1.ConceptoTipoResponse
                Try
                    objConceptoTipo = objWSFEV1.FEParamGetTiposConcepto(FEAuthRequest)
                    vRet = objConceptoTipo.ResultGet
                    If objConceptoTipo.Errors IsNot Nothing Then
                        For i = 0 To objConceptoTipo.Errors.Length - 1
                            MessageBox.Show("objConceptoTipo.Errors(i).Code: " + objConceptoTipo.Errors(i).Code.ToString + vbCrLf +
                            "objConceptoTipo.Errors(i).Msg: " + objConceptoTipo.Errors(i).Msg)
                        Next
                    End If
                    If objConceptoTipo.Events IsNot Nothing Then
                        For i = 0 To objConceptoTipo.Events.Length - 1
                            MessageBox.Show("objConceptoTipo.Events(i).Code: " + objConceptoTipo.Events(i).Code.ToString + vbCrLf +
                            "objConceptoTipo.Events(i).Msg: " + objConceptoTipo.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "FEParamGetTiposDoc"
                ' Invoco al método FEParamGetTiposDoc
                Dim objDocTipo As wsfev1.DocTipoResponse
                Try
                    objDocTipo = objWSFEV1.FEParamGetTiposDoc(FEAuthRequest)
                    vRet = objDocTipo.ResultGet
                    If objDocTipo.Errors IsNot Nothing Then
                        For i = 0 To objDocTipo.Errors.Length - 1
                            MessageBox.Show("objDocTipo.Errors(i).Code: " + objDocTipo.Errors(i).Code.ToString + vbCrLf +
                            "objDocTipo.Errors(i).Msg: " + objDocTipo.Errors(i).Msg)
                        Next
                    End If
                    If objDocTipo.Events IsNot Nothing Then
                        For i = 0 To objDocTipo.Events.Length - 1
                            MessageBox.Show("objDocTipo.Events(i).Code: " + objDocTipo.Events(i).Code.ToString + vbCrLf +
                            "objDocTipo.Events(i).Msg: " + objDocTipo.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "FEParamGetTiposIva"
                ' Invoco al método FEParamGetTiposIva
                Dim objIvaTipo As wsfev1.IvaTipoResponse
                Try
                    objIvaTipo = objWSFEV1.FEParamGetTiposIva(FEAuthRequest)
                    vRet = objIvaTipo.ResultGet
                    If objIvaTipo.Errors IsNot Nothing Then
                        For i = 0 To objIvaTipo.Errors.Length - 1
                            MessageBox.Show("objIvaTipo.Errors(i).Code: " + objIvaTipo.Errors(i).Code.ToString + vbCrLf +
                            "objIvaTipo.Errors(i).Msg: " + objIvaTipo.Errors(i).Msg)
                        Next
                    End If
                    If objIvaTipo.Events IsNot Nothing Then
                        For i = 0 To objIvaTipo.Events.Length - 1
                            MessageBox.Show("objIvaTipo.Events(i).Code: " + objIvaTipo.Events(i).Code.ToString + vbCrLf +
                            "objIvaTipo.Events(i).Msg: " + objIvaTipo.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "FEParamGetTiposMonedas"
                ' Invoco al método FEParamGetTiposMonedas
                Dim objMoneda As wsfev1.MonedaResponse
                Try
                    objMoneda = objWSFEV1.FEParamGetTiposMonedas(FEAuthRequest)
                    vRet = objMoneda.ResultGet
                    If objMoneda.Errors IsNot Nothing Then
                        For i = 0 To objMoneda.Errors.Length - 1
                            MessageBox.Show("objMoneda.Errors(i).Code: " + objMoneda.Errors(i).Code.ToString + vbCrLf +
                            "objMoneda.Errors(i).Msg: " + objMoneda.Errors(i).Msg)
                        Next
                    End If
                    If objMoneda.Events IsNot Nothing Then
                        For i = 0 To objMoneda.Events.Length - 1
                            MessageBox.Show("objMoneda.Events(i).Code: " + objMoneda.Events(i).Code.ToString + vbCrLf +
                            "objMoneda.Events(i).Msg: " + objMoneda.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "FEParamGetTiposOpcional"
                ' Invoco al método FEParamGetTiposOpcional
                Dim objOpcionalTipo As wsfev1.OpcionalTipoResponse
                Try
                    objOpcionalTipo = objWSFEV1.FEParamGetTiposOpcional(FEAuthRequest)
                    vRet = objOpcionalTipo.ResultGet
                    If objOpcionalTipo.Errors IsNot Nothing Then
                        For i = 0 To objOpcionalTipo.Errors.Length - 1
                            MessageBox.Show("objOpcionalTipo.Errors(i).Code: " + objOpcionalTipo.Errors(i).Code.ToString + vbCrLf +
                            "objOpcionalTipo.Errors(i).Msg: " + objOpcionalTipo.Errors(i).Msg)
                        Next
                    End If
                    If objOpcionalTipo.Events IsNot Nothing Then
                        For i = 0 To objOpcionalTipo.Events.Length - 1
                            MessageBox.Show("objOpcionalTipo.Events(i).Code: " + objOpcionalTipo.Events(i).Code.ToString + vbCrLf +
                            "objOpcionalTipo.Events(i).Msg: " + objOpcionalTipo.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "FEParamGetTiposTributos"
                ' Invoco al método FEParamGetTiposTributos
                Dim objFETributoResponse As wsfev1.FETributoResponse
                Try
                    objFETributoResponse = objWSFEV1.FEParamGetTiposTributos(FEAuthRequest)
                    vRet = objFETributoResponse.ResultGet
                    If objFETributoResponse.Errors IsNot Nothing Then
                        For i = 0 To objFETributoResponse.Errors.Length - 1
                            MessageBox.Show("objFETributoResponse.Errors(i).Code: " + objFETributoResponse.Errors(i).Code.ToString + vbCrLf +
                            "objFETributoResponse.Errors(i).Msg: " + objFETributoResponse.Errors(i).Msg)
                        Next
                    End If
                    If objFETributoResponse.Events IsNot Nothing Then
                        For i = 0 To objFETributoResponse.Events.Length - 1
                            MessageBox.Show("objFETributoResponse.Events(i).Code: " + objFETributoResponse.Events(i).Code.ToString + vbCrLf +
                            "objFETributoResponse.Events(i).Msg: " + objFETributoResponse.Errors(i).Msg)
                        Next
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
        End Select
        Return vRet
    End Function
#End Region
#Region "CAE Anticipado"
    'Sub FECAEASolicitar()
    '    Dim Periodo As String = txt_Periodo.Text
    '    Dim Orden As Long = cmb_Orden.SelectedItem.ToString.Substring(0, 1)
    '    Dim objFECAEAGetResponse As wsfev1.FECAEAGetResponse

    '    ' Invoco al método FECAEASolicitar
    '    Try
    '        objFECAEAGetResponse = objWSFEV1.FECAEASolicitar(FEAuthRequest, Periodo, Orden)
    '        If objFECAEAGetResponse IsNot Nothing Then
    '            'Serialize object to a text file.
    '            Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEAGetResponse.xml")
    '            Dim x As New XmlSerializer(objFECAEAGetResponse.GetType)
    '            x.Serialize(objStreamWriter, objFECAEAGetResponse)
    '            objStreamWriter.Close()
    '            MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEAGetResponse.xml")
    '        End If
    '        If objFECAEAGetResponse.Errors IsNot Nothing Then
    '            For i = 0 To objFECAEAGetResponse.Errors.Length - 1
    '                MessageBox.Show("objFECAEAGetResponse.Errors(i).Code: " + objFECAEAGetResponse.Errors(i).Code.ToString + vbCrLf +
    '                "objFECAEAGetResponse.Errors(i).Msg: " + objFECAEAGetResponse.Errors(i).Msg)
    '            Next
    '        End If
    '        If objFECAEAGetResponse.Events IsNot Nothing Then
    '            For i = 0 To objFECAEAGetResponse.Events.Length - 1
    '                MessageBox.Show("objFECAEAGetResponse.Events(i).Code: " + objFECAEAGetResponse.Events(i).Code.ToString + vbCrLf +
    '                "objFECAEAGetResponse.Events(i).Msg: " + objFECAEAGetResponse.Events(i).Msg)
    '            Next
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub
    'Sub FECAEAConsultar()
    '    Dim Periodo As String = txt_Periodo.Text
    '    Dim Orden As Long = cmb_Orden.SelectedItem.ToString.Substring(0, 1)
    '    Dim objFECAEAGetResponse As wsfev1.FECAEAGetResponse

    '    ' Invoco al método FECAEAConsultar
    '    Try
    '        objFECAEAGetResponse = objWSFEV1.FECAEAConsultar(FEAuthRequest, Periodo, Orden)
    '        If objFECAEAGetResponse IsNot Nothing Then
    '            'Serialize object to a text file.
    '            Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEAGetResponse.xml")
    '            Dim x As New XmlSerializer(objFECAEAGetResponse.GetType)
    '            x.Serialize(objStreamWriter, objFECAEAGetResponse)
    '            objStreamWriter.Close()
    '            MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEAGetResponse.xml")
    '        End If
    '        If objFECAEAGetResponse.Errors IsNot Nothing Then
    '            For i = 0 To objFECAEAGetResponse.Errors.Length - 1
    '                MessageBox.Show("objFECAEAGetResponse.Errors(i).Code: " + objFECAEAGetResponse.Errors(i).Code.ToString + vbCrLf +
    '                "objFECAEAGetResponse.Errors(i).Msg: " + objFECAEAGetResponse.Errors(i).Msg)
    '            Next
    '        End If
    '        If objFECAEAGetResponse.Events IsNot Nothing Then
    '            For i = 0 To objFECAEAGetResponse.Events.Length - 1
    '                MessageBox.Show("objFECAEAGetResponse.Events(i).Code: " + objFECAEAGetResponse.Events(i).Code.ToString + vbCrLf +
    '                "objFECAEAGetResponse.Events(i).Msg: " + objFECAEAGetResponse.Events(i).Msg)
    '            Next
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub
    'Sub FECAEASinMovimientoInformar()
    '    Dim PtoVta As String = txt_FECAEASinMovimientoConsultar_PtoVta.Text
    '    Dim CAEA As Long = txt_FECAEASinMovimientoConsultar_CAEA.Text
    '    Dim objFECAEASinMovResponse As wsfev1.FECAEASinMovResponse

    '    ' Invoco al método FECAEASinMovimientoInformar
    '    Try
    '        objFECAEASinMovResponse = objWSFEV1.FECAEASinMovimientoInformar(FEAuthRequest, PtoVta, CAEA)
    '        If objFECAEASinMovResponse IsNot Nothing Then
    '            'Serialize object to a text file.
    '            Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEASinMovResponse.xml")
    '            Dim x As New XmlSerializer(objFECAEASinMovResponse.GetType)
    '            x.Serialize(objStreamWriter, objFECAEASinMovResponse)
    '            objStreamWriter.Close()
    '            MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEASinMovResponse.xml")
    '        End If
    '        If objFECAEASinMovResponse.Errors IsNot Nothing Then
    '            For i = 0 To objFECAEASinMovResponse.Errors.Length - 1
    '                MessageBox.Show("objFECAEASinMovResponse.Errors(i).Code: " + objFECAEASinMovResponse.Errors(i).Code.ToString + vbCrLf +
    '                "objFECAEASinMovResponse.Errors(i).Msg: " + objFECAEASinMovResponse.Errors(i).Msg)
    '            Next
    '        End If
    '        If objFECAEASinMovResponse.Events IsNot Nothing Then
    '            For i = 0 To objFECAEASinMovResponse.Events.Length - 1
    '                MessageBox.Show("objFECAEASinMovResponse.Events(i).Code: " + objFECAEASinMovResponse.Events(i).Code.ToString + vbCrLf +
    '                "objFECAEASinMovResponse.Events(i).Msg: " + objFECAEASinMovResponse.Events(i).Msg)
    '            Next
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub
    'Sub FECAEASinMovimientoConsultar()

    '    Dim PtoVta As String = txt_FECAEASinMovimientoConsultar_PtoVta.Text
    '    Dim CAEA As Long = txt_FECAEASinMovimientoConsultar_CAEA.Text
    '    Dim objFECAEASinMovConsResponse As wsfev1.FECAEASinMovConsResponse

    '    ' Invoco al método FECAEASinMovimientoConsultar
    '    Try
    '        objFECAEASinMovConsResponse = objWSFEV1.FECAEASinMovimientoConsultar(FEAuthRequest, CAEA, PtoVta)
    '        If objFECAEASinMovConsResponse IsNot Nothing Then
    '            'Serialize object to a text file.
    '            Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEASinMovConsResponse.xml")
    '            Dim x As New XmlSerializer(objFECAEASinMovConsResponse.GetType)
    '            x.Serialize(objStreamWriter, objFECAEASinMovConsResponse)
    '            objStreamWriter.Close()
    '            MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEASinMovConsResponse.xml")
    '        End If
    '        If objFECAEASinMovConsResponse.Errors IsNot Nothing Then
    '            For i = 0 To objFECAEASinMovConsResponse.Errors.Length - 1
    '                MessageBox.Show("objFECAEASinMovConsResponse.Errors(i).Code: " + objFECAEASinMovConsResponse.Errors(i).Code.ToString + vbCrLf +
    '                "objFECAEASinMovConsResponse.Errors(i).Msg: " + objFECAEASinMovConsResponse.Errors(i).Msg)
    '            Next
    '        End If
    '        If objFECAEASinMovConsResponse.Events IsNot Nothing Then
    '            For i = 0 To objFECAEASinMovConsResponse.Events.Length - 1
    '                MessageBox.Show("objFECAEASinMovConsResponse.Events(i).Code: " + objFECAEASinMovConsResponse.Events(i).Code.ToString + vbCrLf +
    '                "objFECAEASinMovConsResponse.Events(i).Msg: " + objFECAEASinMovConsResponse.Events(i).Msg)
    '            Next
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub
    'Sub FECAEARegInformativo()
    '    Dim objFECAEACabRequest As New wsfev1.FECAEACabRequest
    '    Dim objFECAEARequest As New wsfev1.FECAEARequest
    '    Dim objFECAEAResponse As New wsfev1.FECAEAResponse

    '    Dim indicemax_arrayFECAEADetRequest As Integer = dgv_FECAEDetRequest.RowCount - 1
    '    Dim d_arrayFECAEADetRequest As Integer = 0
    '    Dim arrayFECAEADetRequest(indicemax_arrayFECAEADetRequest) As wsfev1.FECAEADetRequest

    '    objFECAEACabRequest.CantReg = txt_FECAEARegInformativo_CantReg.Text
    '    objFECAEACabRequest.PtoVta = txt_FECAEARegInformativo_PtoVta.Text
    '    objFECAEACabRequest.CbteTipo = ListBox_FECAEARegInformativo_CbteTipo.SelectedItem.ToString.Substring(0, 2)


    '    For d_arrayFECAEADetRequest = 0 To (indicemax_arrayFECAEADetRequest)
    '        Dim objFECAEADetRequest As New wsfev1.FECAEADetRequest
    '        With objFECAEADetRequest
    '            .Concepto = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(0).Value.ToString.Substring(0, 2)
    '            .DocTipo = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(1).Value.ToString.Substring(0, 2)
    '            .DocNro = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(2).Value
    '            .CbteDesde = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(3).Value
    '            .CbteHasta = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(4).Value
    '            .CbteFch = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(5).Value
    '            .ImpTotal = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(6).Value
    '            .ImpTotConc = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(7).Value
    '            .ImpNeto = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(8).Value
    '            .ImpOpEx = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(9).Value
    '            .ImpTrib = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(10).Value
    '            .ImpIVA = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(11).Value
    '            .FchServDesde = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(12).Value
    '            .FchServHasta = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(13).Value
    '            .FchVtoPago = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(14).Value
    '            .MonId = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(15).Value.ToString.Substring(0, 3)
    '            .MonCotiz = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(16).Value
    '            .CAEA = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(17).Value
    '        End With
    '        arrayFECAEADetRequest(d_arrayFECAEADetRequest) = objFECAEADetRequest
    '    Next d_arrayFECAEADetRequest

    '    With objFECAEARequest
    '        .FeCabReq = objFECAEACabRequest
    '        .FeDetReq = arrayFECAEADetRequest
    '    End With

    '    ' Invoco al método FECAEARegInformativo
    '    Try
    '        objFECAEAResponse = objWSFEV1.FECAEARegInformativo(FEAuthRequest, objFECAEARequest)
    '        If objFECAEAResponse IsNot Nothing Then
    '            'Serialize object to a text file.
    '            Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEAResponse.xml")
    '            Dim x As New XmlSerializer(objFECAEAResponse.GetType)
    '            x.Serialize(objStreamWriter, objFECAEAResponse)
    '            objStreamWriter.Close()
    '            MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEAResponse.xml")
    '        End If
    '        If objFECAEAResponse.Errors IsNot Nothing Then
    '            For i = 0 To objFECAEAResponse.Errors.Length - 1
    '                MessageBox.Show("objFECAEAResponse.Errors(i).Code: " + objFECAEAResponse.Errors(i).Code.ToString + vbCrLf +
    '                "objFECAEAResponse.Errors(i).Msg: " + objFECAEAResponse.Errors(i).Msg)
    '            Next
    '        End If
    '        If objFECAEAResponse.Events IsNot Nothing Then
    '            For i = 0 To objFECAEAResponse.Events.Length - 1
    '                MessageBox.Show("objFECAEAResponse.Events(i).Code: " + objFECAEAResponse.Events(i).Code.ToString + vbCrLf +
    '                "objFECAEAResponse.Events(i).Msg: " + objFECAEAResponse.Events(i).Msg)
    '            Next
    '        End If
    '    Catch ex As Exception
    '    End Try
    'End Sub
#End Region

    Private Sub LimpiarRegistros()
        Eventos = Nothing
        Errores = Nothing
        Eventos = New List(Of KeyValuePair(Of Integer, String))
        Errores = New List(Of KeyValuePair(Of Integer, String))
    End Sub

    Public Shared Function ObtenerAppConfig() As DataTable
        Dim vRet As New DataTable
        vRet.Columns.Add("Clave")
        vRet.Columns.Add("Valor")
        vRet.Rows.Add({"ModoDebug", RegEdit.ObtenerRegistro(eCategorias.WSFE, "ModoDebug")})
        vRet.Rows.Add({"ModoProduccion", RegEdit.ObtenerRegistro(eCategorias.WSFE, "ModoProduccion")})
        vRet.Rows.Add({"RutaComprobantesConsultados", RegEdit.ObtenerRegistro(eCategorias.WSFE, "RutaComprobantesConsultados")})
        vRet.Rows.Add({"RutaComprobantesGenerados", RegEdit.ObtenerRegistro(eCategorias.WSFE, "RutaComprobantesGenerados")})
        vRet.Rows.Add({"RUTATICKETACCESO", RegEdit.ObtenerRegistro(eCategorias.WSFE, "RUTATICKETACCESO")})
        vRet.Rows.Add({"URL_WSFE_Produccion", RegEdit.ObtenerRegistro(eCategorias.WSFE, "URL_WSFE_Produccion")})
        vRet.Rows.Add({"URL_WSFE_Testing", RegEdit.ObtenerRegistro(eCategorias.WSFE, "URL_WSFE_Testing")})
        vRet.Rows.Add({"wsfev1_cliente_wsfev1_Service", RegEdit.ObtenerRegistro(eCategorias.WSFE, "wsfev1_cliente_wsfev1_Service")})
        Return vRet
    End Function
End Class
