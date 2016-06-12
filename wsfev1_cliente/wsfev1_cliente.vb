'<summary>
'EJEMPLO - Cliente consumidor del Web Service WSFEV1 (wsfev1)
'</summary>
'<description>
' Consume los métodos 'FECAEAConsultar', 'FECAEARegInformativo', 'FECAEASinMovimientoConsultar', 
' 'FECAEASinMovimientoInformar', 'FECAEASolicitar', 'FECAESolicitar', 'FECompConsultar', 'FECompTotXRequest', 
' 'FECompUltimoAutorizado', 'FEDummy',  'FEParamGetPtosVenta', 'FEParamGetTiposCbte', 
' 'FEParamGetTiposConcepto', 'FEParamGetTiposDoc', 'FEParamGetTiposIva', 'FEParamGetTiposMonedas', 
' 'FEParamGetTiposOpcional' y 'FEParamGetTiposTributos'.
'</description>
'<version>13.04.16/version>.
'<platform>.NET Framework 2.0</platform>
'<disclaimer>
' El Departamento de Arquitectura Informática de la AFIP (DeArIn/AFIP), pone a disposición
' el siguiente código para su utilización con el Web Service de Factura Electrónica V1 (WSFEV1)
' de la AFIP.
'
' El mismo no puede ser re-distribuido, publicado o descargado en forma total o parcial, ya sea
' en forma electrónica, mecanica u óptica, sin la autorización de DeArIn/AFIP. El uso no
' autorizado del mismo esta prohibido.
'
' DeArIn/AFIP no asume ninguna responsabilidad de los errores que pueda contener el código ni la
' obligación de subsanar dichos errores o informar de la existencia de los mismos.
'
' DeArIn/AFIP no asume ninguna responsabilidad que surja de la utilización del código, ya sea por
' utilización ilegal de patentes, pérdida de beneficios, pérdida de información o cualquier otro
' inconveniente.
'
' Bajo ninguna circunstancia DeArIn/AFIP podra ser indicada como responsable por consecuencias y/o
' incidentes ya sean directos o indirectos que puedan surgir de la utilización del código.
'
' DeArIn/AFIP no da ninguna garantía, expresa o implicita, de la utilidad del código, si el mismo es
' correcto, o si cumple con los requerimientos de algun proposito en particular.
'
' DeArIn/AFIP puede realizar cambios en cualquier momento en el código sin previo aviso.
'
' El código debera ser evaluado, verificado, corregido y/o adaptado por personal técnico calificado
' de las entidades que lo utilicen.
'
' EL SIGUIENTE CODIGO ES DISTRIBUIDO PARA EVALUACION, CON TODOS SUS ERRORES Y OMISIONES. LA
' RESPONSABILIDAD DEL CORRECTO FUNCIONAMIENTO DEL MISMO YA SEA POR SI SOLO O COMO PARTE DE
' OTRA APLICACION, QUEDA A CARGO DE LAS ENTIDADES QUE LO UTILICEN. LA UTILIZACION DEL CODIGO
' SIGNIFICA LA ACEPTACION DE TODOS LOS TERMINOS Y CONDICIONES MENCIONADAS ANTERIORMENTE.
'
' DeArIn-AFIP
'</disclaimer>
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO

Public Class wsfev1_cliente
    Dim objWSFEV1 As New wsfev1.Service
    Dim FEAuthRequest As New wsfev1.FEAuthRequest
    Private Sub wsfev1_cliente_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objWSFEV1.Url = "http://wswhomo.afip.gov.ar/wsfev1/service.asmx"
        Dim RUTATICKETACCESO = "C:\TA.xml"      ' Debe indicar la Ruta de su Ticket de acceso
        Dim tokenBytes As Byte()
        Dim tokenString As String
        dgv_FECAEDetRequest.RowCount = 1
        dgv_FECAEADetRequest.RowCount = 1

        ' ### A LOS CAMPOS DEL FORMULARIO LE ASIGNO UN VALOR POR DEFECTO ### '

        ' Obtengo del tikcet de acceso los campos token, sign y las CUIT del tag relations
        If FileExists(RUTATICKETACCESO) Then
            Dim xmldCredentials As XmlDocument
            Dim xmlnlCredentials As XmlNodeList
            Dim xmlnCredentials As XmlNode
            xmldCredentials = New XmlDocument()
            xmldCredentials.Load(RUTATICKETACCESO)
            xmlnlCredentials = xmldCredentials.SelectNodes("/loginTicketResponse/credentials")
            For Each xmlnCredentials In xmlnlCredentials
                txt_token.Text = xmlnCredentials.ChildNodes.Item(0).InnerText
                txt_sign.Text = xmlnCredentials.ChildNodes.Item(1).InnerText

                tokenBytes = System.Convert.FromBase64String(xmlnCredentials.ChildNodes.Item(0).InnerText)
                tokenString = System.Text.Encoding.UTF8.GetString(tokenBytes)
                Dim xmldToken As XmlDocument
                Dim xmlnlToken As XmlNodeList
                Dim xmlnToken As XmlNode
                xmldToken = New XmlDocument()
                xmldToken.LoadXml(tokenString)
                xmlnlToken = xmldToken.SelectNodes("/sso/operation/login/relations/relation")
                For Each xmlnToken In xmlnlToken
                    cmb_cuit.Items.Add(xmlnToken.Attributes("key").InnerText)
                Next
            Next
            FEAuthRequest.Token = txt_token.Text
            FEAuthRequest.Sign = txt_sign.Text
            FEAuthRequest.Cuit = CInt(cmb_cuit.Items(0))
        Else
            MessageBox.Show("No se encontró el ticket de acceso (Archivo: " + RUTATICKETACCESO + ").")
        End If

        ' Inicializo los controles utilizados en el método 'FECAESolicitar'
        txt_FECAESolicitar_CantReg.Text = "1"
        txt_FECAESolicitar_PtoVta.Text = "1"
        ListBox_FECAESolicitar_CbteTipo.SelectedItem = "06 Factura B"
        dgv_FECAEDetRequest.Rows(0).Cells(0).Value = "02 Servicios"             ' Concepto
        dgv_FECAEDetRequest.Rows(0).Cells(1).Value = "99 Doc. (Otro)"           ' DocTipo
        dgv_FECAEDetRequest.Rows(0).Cells(2).Value = "0"                        ' DocNro
        dgv_FECAEDetRequest.Rows(0).Cells(5).Value = Format(Now, "yyyyMMdd")    ' CbteFch
        dgv_FECAEDetRequest.Rows(0).Cells(6).Value = 121                        ' ImpTotal
        dgv_FECAEDetRequest.Rows(0).Cells(7).Value = 121                        ' ImpTotConc
        dgv_FECAEDetRequest.Rows(0).Cells(8).Value = 0                          ' ImpNeto
        dgv_FECAEDetRequest.Rows(0).Cells(9).Value = 0                          ' ImpOpEx
        dgv_FECAEDetRequest.Rows(0).Cells(10).Value = 0                         ' ImpTrib
        dgv_FECAEDetRequest.Rows(0).Cells(11).Value = 0                         ' ImpIVA
        dgv_FECAEDetRequest.Rows(0).Cells(12).Value = Format(Now, "yyyyMMdd")   ' FchServDesde
        dgv_FECAEDetRequest.Rows(0).Cells(13).Value = Format(Now, "yyyyMMdd")   ' FchServHasta
        dgv_FECAEDetRequest.Rows(0).Cells(14).Value = Format(Now, "yyyyMMdd")   ' FchVtoPago
        dgv_FECAEDetRequest.Rows(0).Cells(15).Value = "PES Pesos Argentinos"    ' MonId
        dgv_FECAEDetRequest.Rows(0).Cells(16).Value = 1                         ' MonCotiz

        ' Inicializo los controles utilizados en el método 'FECAEARegInformativo'
        txt_FECAEARegInformativo_CantReg.Text = "1"
        txt_FECAEARegInformativo_PtoVta.Text = "1"
        ListBox_FECAEARegInformativo_CbteTipo.SelectedItem = "06 Factura B"
        dgv_FECAEADetRequest.Rows(0).Cells(0).Value = "02 Servicios"             ' Concepto
        dgv_FECAEADetRequest.Rows(0).Cells(1).Value = "99 Doc. (Otro)"           ' DocTipo
        dgv_FECAEADetRequest.Rows(0).Cells(2).Value = "0"                        ' DocNro
        dgv_FECAEADetRequest.Rows(0).Cells(5).Value = Format(Now, "yyyyMMdd")    ' CbteFch
        dgv_FECAEADetRequest.Rows(0).Cells(6).Value = 121                        ' ImpTotal
        dgv_FECAEADetRequest.Rows(0).Cells(7).Value = 121                        ' ImpTotConc
        dgv_FECAEADetRequest.Rows(0).Cells(8).Value = 0                          ' ImpNeto
        dgv_FECAEADetRequest.Rows(0).Cells(9).Value = 0                          ' ImpOpEx
        dgv_FECAEADetRequest.Rows(0).Cells(10).Value = 0                         ' ImpTrib
        dgv_FECAEADetRequest.Rows(0).Cells(11).Value = 0                         ' ImpIVA
        dgv_FECAEADetRequest.Rows(0).Cells(12).Value = Format(Now, "yyyyMMdd")   ' FchServDesde
        dgv_FECAEADetRequest.Rows(0).Cells(13).Value = Format(Now, "yyyyMMdd")   ' FchServHasta
        dgv_FECAEADetRequest.Rows(0).Cells(14).Value = Format(Now, "yyyyMMdd")   ' FchVtoPago
        dgv_FECAEADetRequest.Rows(0).Cells(15).Value = "PES Pesos Argentinos"    ' MonId
        dgv_FECAEADetRequest.Rows(0).Cells(16).Value = 1                         ' MonCotiz
    End Sub
    Private Sub bnt_FEDummy_AppServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bnt_FEDummy_AppServer.Click
        Dim objDummy As New wsfev1.DummyResponse
        ' Invoco al método Dummy
        Try
            objDummy = objWSFEV1.FEDummy
            MessageBox.Show(objDummy.AppServer, "FEDummy.AppServer")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub bnt_FEDummy_AuthServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bnt_FEDummy_AuthServer.Click
        Dim objDummy As New wsfev1.DummyResponse
        ' Invoco al método Dummy
        Try
            objDummy = objWSFEV1.FEDummy
            MessageBox.Show(objDummy.AuthServer, "FEDummy.AuthServer")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub btn_FEDummy_DbServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FEDummy_DbServer.Click
        Dim objDummy As New wsfev1.DummyResponse
        ' Invoco al método Dummy
        Try
            objDummy = objWSFEV1.FEDummy
            MessageBox.Show(objDummy.DbServer, "FEDummy.DbServer")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub cbm_metodosDeconsultasGenericas_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbm_metodosDeconsultasGenericas.SelectedIndexChanged
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem

        Select Case cbm_metodosDeconsultasGenericas.SelectedItem
            Case "FECompTotXRequest"
                ' Invoco al método FECompTotXRequest
                Dim objFERegXReqResponse As wsfev1.FERegXReqResponse
                Try
                    objFERegXReqResponse = objWSFEV1.FECompTotXRequest(FEAuthRequest)
                    If objFERegXReqResponse.RegXReq.ToString IsNot Nothing Then
                        MessageBox.Show("objFERegXReqResponse.RegXReq: " + objFERegXReqResponse.RegXReq.ToString + vbCrLf _
    , "objFERegXReqResponse.RegXReq")
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
                    dgv_metodosDeConsultasGenericas.DataSource = objFEPtoVenta.ResultGet
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
                    dgv_metodosDeConsultasGenericas.DataSource = objCbteTipo.ResultGet
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
                    dgv_metodosDeConsultasGenericas.DataSource = objConceptoTipo.ResultGet
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
                    dgv_metodosDeConsultasGenericas.DataSource = objDocTipo.ResultGet
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
                    dgv_metodosDeConsultasGenericas.DataSource = objIvaTipo.ResultGet
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
                    dgv_metodosDeConsultasGenericas.DataSource = objMoneda.ResultGet
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
                    dgv_metodosDeConsultasGenericas.DataSource = objOpcionalTipo.ResultGet
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
                    dgv_metodosDeConsultasGenericas.DataSource = objFETributoResponse.ResultGet
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
    End Sub
    Private Sub btn_FECompUltimoAutorizado_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECompUltimoAutorizado.Click
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem
        Dim PtoVta As Long = ListBox_FECompUltimoAutorizado_PtoVta.SelectedItem.ToString.Substring(0, 2)
        Dim CbteTipo As Long = ListBox_FECompUltimoAutorizado_CbteTipo.SelectedItem.ToString.Substring(0, 2)
        Dim objFERecuperaLastCbteResponse As wsfev1.FERecuperaLastCbteResponse

        ' Invoco al método FECompUltimoAutorizado
        Try
            objFERecuperaLastCbteResponse = objWSFEV1.FECompUltimoAutorizado(FEAuthRequest, CInt(PtoVta), CInt(CbteTipo))
            If objFERecuperaLastCbteResponse IsNot Nothing Then
                MessageBox.Show("objFERecuperaLastCbte.CbteNro: " + objFERecuperaLastCbteResponse.CbteNro.ToString + vbCrLf _
                , "objFERecuperaLastCbte.CbteNro")
            End If
            If objFERecuperaLastCbteResponse.Errors IsNot Nothing Then
                For i = 0 To objFERecuperaLastCbteResponse.Errors.Length - 1
                    MessageBox.Show("objFERecuperaLastCbte.Errors(i).Code: " + objFERecuperaLastCbteResponse.Errors(i).Code.ToString + vbCrLf +
                    "objFERecuperaLastCbte.Errors(i).Msg: " + objFERecuperaLastCbteResponse.Errors(i).Msg)
                Next
            End If
            If objFERecuperaLastCbteResponse.Events IsNot Nothing Then
                For i = 0 To objFERecuperaLastCbteResponse.Events.Length - 1
                    MessageBox.Show("objFERecuperaLastCbte.Events(i).Code: " + objFERecuperaLastCbteResponse.Events(i).Code.ToString + vbCrLf +
                    "objFERecuperaLastCbte.Events(i).Msg: " + objFERecuperaLastCbteResponse.Events(i).Msg)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub btn_FECompConsultar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECompConsultar.Click
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem
        Dim objFECompConsultaReq As New wsfev1.FECompConsultaReq
        Dim objFECompConsultaResponse As wsfev1.FECompConsultaResponse
        objFECompConsultaReq.PtoVta = ListBox_FECompUltimoAutorizado_PtoVta.SelectedItem.ToString.Substring(0, 2)
        objFECompConsultaReq.CbteTipo = ListBox_FECompUltimoAutorizado_CbteTipo.SelectedItem.ToString.Substring(0, 2)
        objFECompConsultaReq.CbteNro = txt_FECompConsultar_CbteNro.Text

        ' Invoco al método FECompConsultar
        Try
            objFECompConsultaResponse = objWSFEV1.FECompConsultar(FEAuthRequest, objFECompConsultaReq)
            If objFECompConsultaResponse IsNot Nothing Then
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECompConsulta.xml")
                Dim x As New XmlSerializer(objFECompConsultaResponse.GetType)
                x.Serialize(objStreamWriter, objFECompConsultaResponse)
                objStreamWriter.Close()
                MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECompConsulta.xml")
            End If
            If objFECompConsultaResponse.Errors IsNot Nothing Then
                For i = 0 To objFECompConsultaResponse.Errors.Length - 1
                    MessageBox.Show("objFECompConsulta.Errors(i).Code: " + objFECompConsultaResponse.Errors(i).Code.ToString + vbCrLf +
                    "objFECompConsulta.Errors(i).Msg: " + objFECompConsultaResponse.Errors(i).Msg)
                Next
            End If
            If objFECompConsultaResponse.Events IsNot Nothing Then
                For i = 0 To objFECompConsultaResponse.Events.Length - 1
                    MessageBox.Show("objFECompConsulta.Events(i).Code: " + objFECompConsultaResponse.Events(i).Code.ToString + vbCrLf +
                    "objFECompConsulta.Events(i).Msg: " + objFECompConsultaResponse.Events(i).Msg)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub btn_FECAESolicitar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAESolicitar.Click
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem

        Dim objFECAECabRequest As New wsfev1.FECAECabRequest
        Dim objFECAERequest As New wsfev1.FECAERequest
        Dim objFECAEResponse As New wsfev1.FECAEResponse

        Dim indicemax_arrayFECAEDetRequest As Integer = dgv_FECAEDetRequest.RowCount - 1
        Dim d_arrayFECAEDetRequest As Integer = 0
        Dim arrayFECAEDetRequest(indicemax_arrayFECAEDetRequest) As wsfev1.FECAEDetRequest

        objFECAECabRequest.CantReg = txt_FECAESolicitar_CantReg.Text
        objFECAECabRequest.CbteTipo = ListBox_FECAESolicitar_CbteTipo.SelectedItem.ToString.Substring(0, 2)
        objFECAECabRequest.PtoVta = txt_FECAESolicitar_PtoVta.Text

        For d_arrayFECAEDetRequest = 0 To (indicemax_arrayFECAEDetRequest)
            Dim objFECAEDetRequest As New wsfev1.FECAEDetRequest
            With objFECAEDetRequest
                .Concepto = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(0).Value.ToString.Substring(0, 2)
                .DocTipo = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(1).Value.ToString.Substring(0, 2)
                .DocNro = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(2).Value
                .CbteDesde = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(3).Value
                .CbteHasta = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(4).Value
                .CbteFch = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(5).Value
                .ImpTotal = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(6).Value
                .ImpTotConc = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(7).Value
                .ImpNeto = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(8).Value
                .ImpOpEx = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(9).Value
                .ImpTrib = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(10).Value
                .ImpIVA = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(11).Value
                .FchServDesde = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(12).Value
                .FchServHasta = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(13).Value
                .FchVtoPago = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(14).Value
                .MonId = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(15).Value.ToString.Substring(0, 3)
                .MonCotiz = dgv_FECAEDetRequest.Rows(d_arrayFECAEDetRequest).Cells(16).Value
            End With
            arrayFECAEDetRequest(d_arrayFECAEDetRequest) = objFECAEDetRequest
        Next d_arrayFECAEDetRequest

        With objFECAERequest
            .FeCabReq = objFECAECabRequest
            .FeDetReq = arrayFECAEDetRequest
        End With

        ' Invoco al método FECAESolicitar
        Try
            objFECAEResponse = objWSFEV1.FECAESolicitar(FEAuthRequest, objFECAERequest)
            If objFECAEResponse IsNot Nothing Then
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAERespons.xml")
                Dim x As New XmlSerializer(objFECAEResponse.GetType)
                x.Serialize(objStreamWriter, objFECAEResponse)
                objStreamWriter.Close()
                MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEResponse.xml")
            End If
            If objFECAEResponse.Errors IsNot Nothing Then
                For i = 0 To objFECAEResponse.Errors.Length - 1
                    MessageBox.Show("objFECAEResponse.Errors(i).Code: " + objFECAEResponse.Errors(i).Code.ToString + vbCrLf +
                    "objFECAEResponse.Errors(i).Msg: " + objFECAEResponse.Errors(i).Msg)
                Next
            End If
            If objFECAEResponse.Events IsNot Nothing Then
                For i = 0 To objFECAEResponse.Events.Length - 1
                    MessageBox.Show("objFECAEResponse.Events(i).Code: " + objFECAEResponse.Events(i).Code.ToString + vbCrLf +
                    "objFECAEResponse.Events(i).Msg: " + objFECAEResponse.Events(i).Msg)
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btn_FECAEDetRequestNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEDetRequestNuevo.Click
        dgv_FECAEDetRequest.Rows.Add()
    End Sub
    Private Sub btn_FECAEDetRequestBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEDetRequestBorrar.Click
        dgv_FECAEDetRequest.Rows.Remove(dgv_FECAEDetRequest.CurrentRow)
    End Sub
    Private Sub btn_FECAEDetRequestEditar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEDetRequestEditar.Click
        dgv_FECAEDetRequest.BeginEdit(True)
    End Sub
    Private Sub btn_FECAEDetRequestFinEditar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEDetRequestFinEditar.Click
        dgv_FECAEDetRequest.EndEdit(True)
    End Sub
    Private Sub btn_FECAEADetRequestNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEADetRequestNuevo.Click
        dgv_FECAEADetRequest.Rows.Add()
    End Sub
    Private Sub btn_FECAEADetRequestBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEADetRequestBorrar.Click
        dgv_FECAEADetRequest.Rows.Remove(dgv_FECAEDetRequest.CurrentRow)
    End Sub
    Private Sub btn_FECAEADetRequestEditar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEADetRequestEditar.Click
        dgv_FECAEADetRequest.BeginEdit(True)
    End Sub
    Private Sub btn_FECAEADetRequestFinEditar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEADetRequestFinEditar.Click
        dgv_FECAEADetRequest.EndEdit(True)
    End Sub
    Private Sub btn_FECAEASolicitar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEASolicitar.Click
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem
        Dim Periodo As String = txt_Periodo.Text
        Dim Orden As Long = cmb_Orden.SelectedItem.ToString.Substring(0, 1)
        Dim objFECAEAGetResponse As wsfev1.FECAEAGetResponse

        ' Invoco al método FECAEASolicitar
        Try
            objFECAEAGetResponse = objWSFEV1.FECAEASolicitar(FEAuthRequest, CInt(Periodo), CShort(Orden))
            If objFECAEAGetResponse IsNot Nothing Then
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEAGetResponse.xml")
                Dim x As New XmlSerializer(objFECAEAGetResponse.GetType)
                x.Serialize(objStreamWriter, objFECAEAGetResponse)
                objStreamWriter.Close()
                MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEAGetResponse.xml")
            End If
            If objFECAEAGetResponse.Errors IsNot Nothing Then
                For i = 0 To objFECAEAGetResponse.Errors.Length - 1
                    MessageBox.Show("objFECAEAGetResponse.Errors(i).Code: " + objFECAEAGetResponse.Errors(i).Code.ToString + vbCrLf +
                    "objFECAEAGetResponse.Errors(i).Msg: " + objFECAEAGetResponse.Errors(i).Msg)
                Next
            End If
            If objFECAEAGetResponse.Events IsNot Nothing Then
                For i = 0 To objFECAEAGetResponse.Events.Length - 1
                    MessageBox.Show("objFECAEAGetResponse.Events(i).Code: " + objFECAEAGetResponse.Events(i).Code.ToString + vbCrLf +
                    "objFECAEAGetResponse.Events(i).Msg: " + objFECAEAGetResponse.Events(i).Msg)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub btn_FECAEAConsultar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEAConsultar.Click
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem
        Dim Periodo As String = txt_Periodo.Text
        Dim Orden As Long = cmb_Orden.SelectedItem.ToString.Substring(0, 1)
        Dim objFECAEAGetResponse As wsfev1.FECAEAGetResponse

        ' Invoco al método FECAEAConsultar
        Try
            objFECAEAGetResponse = objWSFEV1.FECAEAConsultar(FEAuthRequest, CInt(Periodo), CShort(Orden))
            If objFECAEAGetResponse IsNot Nothing Then
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEAGetResponse.xml")
                Dim x As New XmlSerializer(objFECAEAGetResponse.GetType)
                x.Serialize(objStreamWriter, objFECAEAGetResponse)
                objStreamWriter.Close()
                MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEAGetResponse.xml")
            End If
            If objFECAEAGetResponse.Errors IsNot Nothing Then
                For i = 0 To objFECAEAGetResponse.Errors.Length - 1
                    MessageBox.Show("objFECAEAGetResponse.Errors(i).Code: " + objFECAEAGetResponse.Errors(i).Code.ToString + vbCrLf +
                    "objFECAEAGetResponse.Errors(i).Msg: " + objFECAEAGetResponse.Errors(i).Msg)
                Next
            End If
            If objFECAEAGetResponse.Events IsNot Nothing Then
                For i = 0 To objFECAEAGetResponse.Events.Length - 1
                    MessageBox.Show("objFECAEAGetResponse.Events(i).Code: " + objFECAEAGetResponse.Events(i).Code.ToString + vbCrLf +
                    "objFECAEAGetResponse.Events(i).Msg: " + objFECAEAGetResponse.Events(i).Msg)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub bnt_FECAEASinMovimientoInformar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bnt_FECAEASinMovimientoInformar.Click
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem
        Dim PtoVta As String = txt_FECAEASinMovimientoConsultar_PtoVta.Text
        Dim CAEA As Long = txt_FECAEASinMovimientoConsultar_CAEA.Text
        Dim objFECAEASinMovResponse As wsfev1.FECAEASinMovResponse

        ' Invoco al método FECAEASinMovimientoInformar
        Try
            objFECAEASinMovResponse = objWSFEV1.FECAEASinMovimientoInformar(FEAuthRequest, CInt(PtoVta), CStr(CAEA))
            If objFECAEASinMovResponse IsNot Nothing Then
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEASinMovResponse.xml")
                Dim x As New XmlSerializer(objFECAEASinMovResponse.GetType)
                x.Serialize(objStreamWriter, objFECAEASinMovResponse)
                objStreamWriter.Close()
                MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEASinMovResponse.xml")
            End If
            If objFECAEASinMovResponse.Errors IsNot Nothing Then
                For i = 0 To objFECAEASinMovResponse.Errors.Length - 1
                    MessageBox.Show("objFECAEASinMovResponse.Errors(i).Code: " + objFECAEASinMovResponse.Errors(i).Code.ToString + vbCrLf +
                    "objFECAEASinMovResponse.Errors(i).Msg: " + objFECAEASinMovResponse.Errors(i).Msg)
                Next
            End If
            If objFECAEASinMovResponse.Events IsNot Nothing Then
                For i = 0 To objFECAEASinMovResponse.Events.Length - 1
                    MessageBox.Show("objFECAEASinMovResponse.Events(i).Code: " + objFECAEASinMovResponse.Events(i).Code.ToString + vbCrLf +
                    "objFECAEASinMovResponse.Events(i).Msg: " + objFECAEASinMovResponse.Events(i).Msg)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub bnt_FECAEASinMovimientoConsultar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bnt_FECAEASinMovimientoConsultar.Click
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem
        Dim PtoVta As String = txt_FECAEASinMovimientoConsultar_PtoVta.Text
        Dim CAEA As Long = txt_FECAEASinMovimientoConsultar_CAEA.Text
        Dim objFECAEASinMovConsResponse As wsfev1.FECAEASinMovConsResponse

        ' Invoco al método FECAEASinMovimientoConsultar
        Try
            objFECAEASinMovConsResponse = objWSFEV1.FECAEASinMovimientoConsultar(FEAuthRequest, CStr(CAEA), CInt(PtoVta))
            If objFECAEASinMovConsResponse IsNot Nothing Then
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEASinMovConsResponse.xml")
                Dim x As New XmlSerializer(objFECAEASinMovConsResponse.GetType)
                x.Serialize(objStreamWriter, objFECAEASinMovConsResponse)
                objStreamWriter.Close()
                MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEASinMovConsResponse.xml")
            End If
            If objFECAEASinMovConsResponse.Errors IsNot Nothing Then
                For i = 0 To objFECAEASinMovConsResponse.Errors.Length - 1
                    MessageBox.Show("objFECAEASinMovConsResponse.Errors(i).Code: " + objFECAEASinMovConsResponse.Errors(i).Code.ToString + vbCrLf +
                    "objFECAEASinMovConsResponse.Errors(i).Msg: " + objFECAEASinMovConsResponse.Errors(i).Msg)
                Next
            End If
            If objFECAEASinMovConsResponse.Events IsNot Nothing Then
                For i = 0 To objFECAEASinMovConsResponse.Events.Length - 1
                    MessageBox.Show("objFECAEASinMovConsResponse.Events(i).Code: " + objFECAEASinMovConsResponse.Events(i).Code.ToString + vbCrLf +
                    "objFECAEASinMovConsResponse.Events(i).Msg: " + objFECAEASinMovConsResponse.Events(i).Msg)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub btn_FECAEARegInformativo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FECAEARegInformativo.Click
        FEAuthRequest.Token = txt_token.Text
        FEAuthRequest.Sign = txt_sign.Text
        FEAuthRequest.Cuit = cmb_cuit.SelectedItem

        Dim objFECAEACabRequest As New wsfev1.FECAEACabRequest
        Dim objFECAEARequest As New wsfev1.FECAEARequest
        Dim objFECAEAResponse As New wsfev1.FECAEAResponse

        Dim indicemax_arrayFECAEADetRequest As Integer = dgv_FECAEDetRequest.RowCount - 1
        Dim d_arrayFECAEADetRequest As Integer = 0
        Dim arrayFECAEADetRequest(indicemax_arrayFECAEADetRequest) As wsfev1.FECAEADetRequest

        objFECAEACabRequest.CantReg = txt_FECAEARegInformativo_CantReg.Text
        objFECAEACabRequest.PtoVta = txt_FECAEARegInformativo_PtoVta.Text
        objFECAEACabRequest.CbteTipo = ListBox_FECAEARegInformativo_CbteTipo.SelectedItem.ToString.Substring(0, 2)


        For d_arrayFECAEADetRequest = 0 To (indicemax_arrayFECAEADetRequest)
            Dim objFECAEADetRequest As New wsfev1.FECAEADetRequest
            With objFECAEADetRequest
                .Concepto = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(0).Value.ToString.Substring(0, 2)
                .DocTipo = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(1).Value.ToString.Substring(0, 2)
                .DocNro = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(2).Value
                .CbteDesde = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(3).Value
                .CbteHasta = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(4).Value
                .CbteFch = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(5).Value
                .ImpTotal = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(6).Value
                .ImpTotConc = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(7).Value
                .ImpNeto = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(8).Value
                .ImpOpEx = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(9).Value
                .ImpTrib = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(10).Value
                .ImpIVA = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(11).Value
                .FchServDesde = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(12).Value
                .FchServHasta = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(13).Value
                .FchVtoPago = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(14).Value
                .MonId = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(15).Value.ToString.Substring(0, 3)
                .MonCotiz = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(16).Value
                .CAEA = dgv_FECAEADetRequest.Rows(d_arrayFECAEADetRequest).Cells(17).Value
            End With
            arrayFECAEADetRequest(d_arrayFECAEADetRequest) = objFECAEADetRequest
        Next d_arrayFECAEADetRequest

        With objFECAEARequest
            .FeCabReq = objFECAEACabRequest
            .FeDetReq = arrayFECAEADetRequest
        End With

        ' Invoco al método FECAEARegInformativo
        Try
            objFECAEAResponse = objWSFEV1.FECAEARegInformativo(FEAuthRequest, objFECAEARequest)
            If objFECAEAResponse IsNot Nothing Then
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter("C:\WSFEV1_objFECAEAResponse.xml")
                Dim x As New XmlSerializer(objFECAEAResponse.GetType)
                x.Serialize(objStreamWriter, objFECAEAResponse)
                objStreamWriter.Close()
                MessageBox.Show("Se generó el archivo C:\WSFEV1_objFECAEAResponse.xml")
            End If
            If objFECAEAResponse.Errors IsNot Nothing Then
                For i = 0 To objFECAEAResponse.Errors.Length - 1
                    MessageBox.Show("objFECAEAResponse.Errors(i).Code: " + objFECAEAResponse.Errors(i).Code.ToString + vbCrLf +
                    "objFECAEAResponse.Errors(i).Msg: " + objFECAEAResponse.Errors(i).Msg)
                Next
            End If
            If objFECAEAResponse.Events IsNot Nothing Then
                For i = 0 To objFECAEAResponse.Events.Length - 1
                    MessageBox.Show("objFECAEAResponse.Events(i).Code: " + objFECAEAResponse.Events(i).Code.ToString + vbCrLf +
                    "objFECAEAResponse.Events(i).Msg: " + objFECAEAResponse.Events(i).Msg)
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub grp_FECAESolicitar_Enter(sender As Object, e As EventArgs) Handles grp_FECAESolicitar.Enter

    End Sub

    Private Sub dgv_FECAEDetRequest_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_FECAEDetRequest.CellContentClick

    End Sub
End Class
