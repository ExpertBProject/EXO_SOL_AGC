Imports SAPbouiCOM
Imports EXO_SOL_AGC.Extensions
Imports System.Text

Public Class EXO_COMPRAS
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "variables globales"
    Dim _sDocNum As String = ""
    Dim _sSerie As String = ""
#End Region

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
            GenerarParametros()
        End If
    End Sub
    Private Sub cargaCampos()
        Dim bValidado As Boolean = False
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDF_EXO_OPCH.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDF_EXO_OPCH", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            If bValidado = True Then
                objGlobal.SBOApp.StatusBar.SetText("Validado: UDF_EXO_OPCH", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'Else
                '    objGlobal.SBOApp.StatusBar.SetText("No Validado: UDF_EXO_OPCH", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End If
    End Sub
    Private Sub GenerarParametros()
        If objGlobal.refDi.comunes.esAdministrador Then
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("AGC_Mail") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("AGC_Mail", "facturas@solariaenergia.com")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("AGC_Mail_US") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("AGC_Mail_US", "facturas@solariaenergia.com")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("AGC_Mail_Pass") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("AGC_Mail_Pass", "Cud06785")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("AGC_Mail_SMTP") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("AGC_Mail_SMTP", "smtp.office365.com")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("AGC_Mail_Port") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("AGC_Mail_Port", "587")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("AGC_Mail_Notif") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("AGC_Mail_Notif", "N")
            End If
        End If
    End Sub
    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "141"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    'If EventHandler_ItemPressed_After(infoEvento) = False Then
                                    '    GC.Collect()
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "141"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    'If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                    '    GC.Collect()
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "141"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "141"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sTable_Origen As String = ""
#End Region

        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "1"
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                        sTable_Origen = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).DataBind.TableName

                        _sDocNum = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("DocNum", 0).ToString
                        _sSerie = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("Series", 0).ToString
                    End If
            End Select

            EventHandler_ItemPressed_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sTable_Origen As String = ""
        Dim sDocEntry As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
#End Region

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "1"
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                        sTable_Origen = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).DataBind.TableName
                        sSQL = "SELECT ""DocEntry"" FROM """ & sTable_Origen & """ WHERE ""Series""=" & _sSerie & " and ""DocNum""=" & _sDocNum
                        oRs.DoQuery(sSQL)
                        If oRs.RecordCount > 0 Then
                            sDocEntry = oRs.Fields.Item("DocEntry").Value.ToString
                            If sDocEntry <> "" Then
                                Envio_de_Doc_Creada(objGlobal, objGlobal.compañia, objGlobal.compañia.CompanyDB, sTable_Origen, sDocEntry)
                            End If
                        End If
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim resultado As Boolean = True
        Dim sTable_Origen As String = ""
        Dim sDocEntry As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
#End Region

        Try
            If infoEvento.BeforeAction = False Then
                Select Case infoEvento.FormTypeEx
                    Case "141"
                        Select Case infoEvento.EventType
                            Case BoEventTypes.et_FORM_DATA_ADD, BoEventTypes.et_FORM_DATA_UPDATE
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                sTable_Origen = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).DataBind.TableName
                                _sDocNum = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("DocNum", 0).ToString
                                _sSerie = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("Series", 0).ToString
                                sDocEntry = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("DocEntry", 0).ToString
                                If sDocEntry = "" Then
                                    sSQL = "SELECT ""DocEntry"" FROM """ & sTable_Origen & """ WHERE ""Series""=" & _sSerie & " and ""DocNum""=" & _sDocNum
                                    oRs.DoQuery(sSQL)
                                    If oRs.RecordCount > 0 Then
                                        sDocEntry = oRs.Fields.Item("DocEntry").Value.ToString

                                    End If
                                End If
                                If sDocEntry <> "" Then
                                    Envio_de_Doc_Creada(objGlobal, objGlobal.compañia, objGlobal.compañia.CompanyDB, sTable_Origen, sDocEntry)
                                End If

                        End Select
                End Select
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oform)
        End Try
        Return resultado
    End Function
    Public Shared Sub Envio_de_Doc_Creada(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oCompany As SAPbobsCOM.Company, ByVal sBBDD As String, tipodoc As String, ByVal sDocEntry As String)
#Region "Variables"
        Dim sError As String = ""
        Dim sRutaFicheros As String = ""
        Dim SRutaLog As String = ""
        Dim dtDocumentos As Data.DataTable = New System.Data.DataTable()
        Dim sRuta As String = ""
        Dim sFicheroCrystal As String = ""
        Dim sTextoTipoDoc As String = ""
        Dim DocumentoPdf As String = ""
#End Region

        Try
            sRuta = oObjGlobal.pathHistorico & "\Report\"

            'Crear y validar carpetas, obtener rutas para los ficheros
            ValidarCarpetas(sRuta, sBBDD, tipodoc, sRutaFicheros, sFicheroCrystal)

            If tipodoc = "OPCH" Then
                sTextoTipoDoc = "Factura"
            Else
                sTextoTipoDoc = "Abono"
            End If

            'consulta para buscar documentos a realizar
            Dim sSQL As String = " SELECT T0.""DocEntry"",T0.""DocNum"",T0.""TaxDate"", T0.""CardCode"",T0.""CardName"" , t1.""CardFName"", ifnull(t1.""E_Mail"",'') as ""CorreoProv"",  " &
                    " ifnull(t2.""Email"",'') as ""CorreoComercial"", t3.""USER_CODE"" as ""CodigoCreador"", t3.""E_Mail"" as ""CorreoCreador"" " &
                    " FROM """ & sBBDD & """.""" & tipodoc & """ T0 left join """ & sBBDD & """.""OCRD"" t1 on t0.""CardCode""=t1.""CardCode"" " &
                    " left join """ & sBBDD & """.""OSLP"" t2 on T0.""SlpCode""=t2.""SlpCode"" " &
                    " left join """ & sBBDD & """.""OUSR"" t3 on T0.""UserSign""=t3.""USERID"" " &
                    " WHERE ifnull(T0.""U_EXO_ENVFACT"",'')='N' and ifnull(T0.""U_EXO_IPC"",'')='Y' and T0.""DocEntry""= " & sDocEntry
            ' dtDocumentos.ExecuteQuery(sSQL)

            dtDocumentos = oObjGlobal.refDi.SQL.executeQuery(sSQL)

            If dtDocumentos.Rows.Count > 0 Then
                For Each row As DataRow In dtDocumentos.Rows
                    GestionarDoc(oObjGlobal, oCompany, row, sFicheroCrystal, sRutaFicheros, sTextoTipoDoc, sBBDD, tipodoc)
                Next
            Else
                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - La documento no cumple los requisitos para ser enviado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            dtDocumentos = Nothing
        End Try
    End Sub
    Private Shared Sub ValidarCarpetas(sRuta As String, empresa As String, tipodoc As String, ByRef sRutaFicheros As String, ByRef sFicheroCrystal As String)
        Try
            Dim sArchivo As String = ""
            Dim sRutaFinal As String = sRuta

            If System.IO.Directory.Exists(sRutaFinal) = False Then
                System.IO.Directory.CreateDirectory(sRutaFinal)
            End If

            sRutaFinal = sRutaFinal & tipodoc & "\"
            If System.IO.Directory.Exists(sRutaFinal) = False Then
                System.IO.Directory.CreateDirectory(sRutaFinal)
            End If

            If tipodoc = "OPCH" Then
                sArchivo = "Factura"
            Else
                sArchivo = "Abono"
            End If

            sFicheroCrystal = sRutaFinal & sArchivo & ".rpt"

            Dim sRutaFinal2 As String = sRutaFinal & "Ficheros" & "\"
            If System.IO.Directory.Exists(sRutaFinal2) = False Then
                System.IO.Directory.CreateDirectory(sRutaFinal2)
            End If

            sRutaFicheros = sRutaFinal2


        Catch ex As Exception
        End Try
    End Sub
    Private Shared Sub GestionarDoc(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oCompany As SAPbobsCOM.Company, ByRef row As DataRow, sFicheroCrystal As String, sRutaFicheros As String, sTextoTipoDoc As String, sBBDD As String, ByVal tipodoc As String)

        Dim smensaje As String = ""
        Dim mensajeErrorCorreo As String = ""
        Dim DocumentoPdf As String = ""
        Dim dtDirMailsCli As Data.DataTable = New System.Data.DataTable()
        Dim oRs As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsFac As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Try
            Dim sSQL As String = ""
            Dim dFecha As Date = CDate(row.Item("TaxDate").ToString())
            'oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & dFecha.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim iNumero As Integer = 0
            'Buscamos la ultima factura de año y le cogemos los digitos del final para sumar el numero
            sSQL = "SELECT ""NumAtCard"", ""DocNum"" FROM """ & sBBDD & """.""OPCH"" where year(""TaxDate"")='" & dFecha.Year.ToString("0000") & "' and ""CANCELED""='N'"
            sSQL &= " And ""CardCode""='" & row.Item("CardCode").ToString() & "' order by ""DocEntry"" desc "
            oRsFac.DoQuery(sSQL)
            If oRsFac.RecordCount > 0 Then
                oRsFac.MoveFirst()
                oRsFac.MoveNext()
                Dim sTexto As String = "" : Dim sExtraeNum() As String
                sTexto = oRsFac.Fields.Item("NumAtCard").Value.ToString
                sExtraeNum = sTexto.Split("/")
                For i As Integer = 0 To sExtraeNum.Length - 1
                    If sExtraeNum(i) <> "" Then
                        sTexto = sExtraeNum(i)
                    End If
                Next
                If IsNumeric(sTexto) Then
                    iNumero = CInt(sTexto)
                Else
                    iNumero = 0
                End If
            Else
                iNumero = 0
            End If
            iNumero += 1
            Dim sNumAtCard As String = dFecha.Year.ToString("0000") & "-" & row.Item("CardFName").ToString() & "/" & iNumero.ToString
            'oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Nº Ref: " & sNumAtCard.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'todo ok, en la funcion ya hemos escrito en el log
            sSQL = "UPDATE """ & sBBDD & """.""" & tipodoc & """ Set "
            sSQL &= """NumAtCard""='" & sNumAtCard & "' "
            sSQL &= " WHERE ""DocEntry"" = " & row.Item("DocEntry").ToString()
            oRs.DoQuery(sSQL)
            sSQL = "UPDATE """ & sBBDD & """.""OJDT"" SET ""Ref2""='" & sNumAtCard & "' WHERE ""BaseRef""= " & row.Item("DocNum").ToString() & ";"
            oRs.DoQuery(sSQL)
            sSQL = "UPDATE A SET ""Ref2""='" & sNumAtCard & "' FROM """ & sBBDD & """.""OJDT"" O INNER JOIN """ & sBBDD & """.""JDT1"" A ON O.""TransId""=A.""TransId"" WHERE O.""BaseRef""= " & row.Item("DocNum").ToString() & ";"
            oRs.DoQuery(sSQL)
            oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Se actualiza el numero de referencia: " & sNumAtCard, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'buscamos los correos del cliente en las personas de contacto
            sSQL = "select ""CardCode"",""E_MailL"" from """ & sBBDD & """.""OCPR"" where ifnull(""EmlGrpCode"",'')='FACTURA_DS' and ""CardCode""='" & row.Item("CardCode").ToString() & "'"
            'dtDirMailsCli.ExecuteQuery(ssql)
            dtDirMailsCli = oObjGlobal.refDi.SQL.executeQuery(sSQL)
            If dtDirMailsCli.Rows.Count = 0 Then
                smensaje = "El Proveedor " & row.Item("CardCode").ToString() & " no tiene correo electronico. No se puede enviar la factura. Tiene que enviarla manualmente. "
                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & smensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oObjGlobal.SBOApp.MessageBox(smensaje)
            Else
                If GenerarPDFyEnviarCrystal(oObjGlobal, oCompany, DocumentoPdf, sFicheroCrystal, sRutaFicheros, row.Item("DocNum").ToString(), row.Item("DocEntry").ToString(), sBBDD, sTextoTipoDoc) = True Then
                    If EnviarEmail(oObjGlobal, "", dtDirMailsCli, row.Item("TaxDate").ToString(), row.Item("DocNum").ToString(), row.Item("DocEntry").ToString(), row.Item("DocEntry").ToString(), row.Item("CardCode").ToString(), sTextoTipoDoc, DocumentoPdf, mensajeErrorCorreo, sBBDD) = True Then
                        'todo ok, en la funcion ya hemos escrito en el log
                        sSQL = "UPDATE """ & sBBDD & """.""" & tipodoc & """ Set ""U_EXO_ENVFACT""='Y' "
                        sSQL &= " WHERE ""DocEntry"" = " & row.Item("DocEntry").ToString()
                        oRs.DoQuery(sSQL)
                        oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Se actualiza la factura como enviada", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Else
                        'alerta al creador del doc
                        smensaje = "Error al enviar correo electrónico: " + row.Item("CardCode").ToString() & " - " & mensajeErrorCorreo
                        oObjGlobal.SBOApp.MessageBox(smensaje)

                        ''correo al creador del doc y al comercial
                        'If row.Item("CorreoCreador").ToString <> "" Then
                        '    mensajeErrorCorreo = "noEnviado"
                        '    If Not EnviarEmail(oObjGlobal, "", dtDirMailsCli, row.Item("TaxDate").ToString(), row.Item("DocNum").ToString(), row.Item("DocEntry").ToString(), row.Item("CardCode").ToString(), row.Item("CardName").ToString(), sTextoTipoDoc, DocumentoPdf, mensajeErrorCorreo, sBBDD) Then
                        '        smensaje = "Error al enviar correo electrónico a creador: " + row.Item("CodigoCreador").ToString() + " " + row.Item("CardCode").ToString()
                        '        oObjGlobal.SBOApp.MessageBox(smensaje)
                        '    End If
                        'Else
                        '    smensaje = "El creador del documento no tiene correo electrónico configurado " & row.Item("CodigoCreador").ToString()
                        '    oObjGlobal.SBOApp.MessageBox(smensaje)
                        'End If

                    End If
                Else
                    'es un error en el catch y por tanto ya tenemos el log en la funcion de generar crystal.
                End If
            End If

        Catch ex As Exception
            smensaje = "Exception GestionaDoc: " + ex.Message
            oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & smensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oObjGlobal.SBOApp.MessageBox(smensaje)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsFac, Object))
        End Try
    End Sub
    Private Shared Function EnviarEmail(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, dirmail As String, dtdircli As Data.DataTable, FechaDoc As String, numpedido As String, docentry As String, cardcode As String, cardname As String, tipodoc As String, filename As String, ByRef mensajeErrorCorreo As String, empresa As String) As Boolean

        Dim correo As New System.Net.Mail.MailMessage()
        Dim adjunto As System.Net.Mail.Attachment

        Dim StrFirma As String = ""
        Dim htmbody As New System.Text.StringBuilder()
        Dim cuerpo As String = ""
        Dim sMail As String = ""
        Dim sMailUS As String = "" : Dim sMailPASS As String = ""
        Dim sMailSMTP As String = "" : Dim sMailPORT As String = ""
        Dim sMailNotif As String = ""
        Try
            sMail = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("AGC_Mail")
            sMailUS = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("AGC_Mail_US")
            sMailPASS = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("AGC_Mail_Pass")
            sMailSMTP = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("AGC_Mail_SMTP")
            sMailPORT = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("AGC_Mail_Port")
            sMailNotif = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("AGC_Mail_Notif")
            Select Case empresa
                'Case "SEMA_PROD" : correo.From = New System.Net.Mail.MailAddress("omartinez@expertone.es", "Prueba Solaria")
                Case Else
                    correo.From = New System.Net.Mail.MailAddress(sMail, "Solaria Energía y Medio Ambiente")
                    'correo.CC.Add(sMail)
            End Select


            If filename <> "" Then
                adjunto = New System.Net.Mail.Attachment(filename)
                correo.Attachments.Add(adjunto)
            End If

            If mensajeErrorCorreo = "noEnviado" Then
                'vuelvo a realizar la consulta y lo envio todo junto
                correo.To.Clear()
                'correo.To.Add("facturas@solariaenergia.com")
                If dirmail <> "" Then
                    correo.To.Add(dirmail)
                End If

                correo.Subject = "No se pudo enviar el correo al Interlocutor " & tipodoc & " " + numpedido.ToString()

                cuerpo = "No se pudo enviar el documento adjunto al Interlocutor " & cardcode & " " & cardname & "." & Chr(13)
                cuerpo = cuerpo & "Pongase en contacto con su departamento de IT o revise la información en la ficha del Interlocutor."
            Else
                'Dim FicheroCab As String = oObjGlobal.pathHistorico & "\Report\mail.htm"
                'Dim srCAB As StreamReader = New StreamReader(FicheroCab)
                'cuerpo = srCAB.ReadToEnd()
                'vuelvo a realizar la consulta y lo envio todo junto
                If dirmail = "" Then
                    For Each row As DataRow In dtdircli.Rows
                        correo.To.Add(row.Item("E_MailL").ToString)
                    Next
                Else
                    correo.To.Add(dirmail)
                End If


                Dim dDate As Date = CDate(FechaDoc)
                correo.Subject = tipodoc & " " & "DS " & MonthName(dDate.Month) & " " & dDate.Year.ToString("0000")
                Dim strHeader As String = "<table><tbody>"
                Dim strFooter As String = "</tbody></table>"
                Dim sbContent As New StringBuilder()
                sbContent.Append(String.Format("<td>{0}</td>", "Buenos días, "))
                sbContent.Append("</tr>")
                sbContent.Append("<tr>")
                sbContent.Append(String.Format("<td>{0}</td>", "Se adjunta la " + tipodoc + " de derechos de superficie correspondiente al mes de " & MonthName(dDate.Month) & " de " & dDate.Year.ToString("0000") & "."))
                sbContent.Append("</tr>")
                sbContent.Append("<tr>")
                sbContent.Append(String.Format("<td>{0}</td>", "Un saludo, "))
                sbContent.Append("</tr>")
                sbContent.Append("<tr>")
                sbContent.Append(String.Format("<td>{0}</td>", "Departamento Administración"))
                sbContent.Append("</tr>")
                sbContent.Append(String.Format("<td>{0}</td>", "Solaria Energía y Medio Ambiente"))
                sbContent.Append("</tr>")
                sbContent.Append("<tr>")
                sbContent.Append(String.Format("<td>{0}</td>", "Este email se ha generado automáticamente, no responda al mismo."))
                sbContent.Append("</tr>")
                sbContent.Append("<tr>")
                Dim emailTemplate As String = strHeader & sbContent.ToString() & strFooter
                cuerpo &= strHeader & sbContent.ToString() & strFooter
                correo.IsBodyHtml = True
                correo.Body = cuerpo
                correo.Priority = System.Net.Mail.MailPriority.Normal
                If sMailNotif = "Y" Then
                    correo.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.OnSuccess
                ElseIf sMailNotif = "F" Then
                    correo.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.OnFailure
                ElseIf sMailNotif = "N" Then
                    correo.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.None
                End If

            End If

            Dim smtp As New System.Net.Mail.SmtpClient
            Select Case empresa
                'Case "SEMA_PROD"
                '    smtp.Host = "smtp.office365.com"
                '    smtp.Port = 587
                '    smtp.UseDefaultCredentials = True
                '    smtp.Credentials = New System.Net.NetworkCredential("omartinez@expertone.es", "Osma@2021")
                '    smtp.EnableSsl = True
                Case Else
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - smtp: " & sMailSMTP & " - Port: " & sMailPORT & " - TargetName: STARTTLS/smtp.office365.com - Mail: " & sMail & " - Usuario: " & sMailUS & " - Pass: " & sMailPASS, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    smtp.UseDefaultCredentials = False
                    smtp.Credentials = New System.Net.NetworkCredential(sMailUS, sMailPASS)
                    smtp.Host = sMailSMTP
                    smtp.Port = sMailPORT
                    smtp.EnableSsl = True
                    smtp.TargetName = "STARTTLS/smtp.office365.com"
            End Select

            smtp.Send(correo)
            correo.Dispose()

            oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Correo enviado: " & dirmail & " " & tipodoc & " " & numpedido, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Return True
        Catch ex As Exception
            EnviarEmail = False
            mensajeErrorCorreo = ex.Message
        End Try
        Return False
    End Function
    Private Shared Function GenerarPDFyEnviarCrystal(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oCompany As SAPbobsCOM.Company, ByRef NombreDocumentoPdf As String, ByVal strRutaInforme As String, ByVal sRutaFicheros As String, numpedido As String, docentry As String, empresa As String, sTextoTipoDoc As String) As Boolean
#Region "Variables"
        Dim oCRReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Dim sFilePDF As String = sTextoTipoDoc & "_" & numpedido
        Dim strNombrePDF As String = sTextoTipoDoc & "_" & numpedido
        Dim Sql As String = ""
        Dim sMensaje As String = ""
        Dim sServer As String = ""
        Dim sDriver As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
#End Region
        GenerarPDFyEnviarCrystal = False
        Try
            'Establecemos las conexiones a la BBDD
            sServer = oObjGlobal.compañia.Server.ToString
            sBBDD = oObjGlobal.compañia.CompanyDB.ToString
            'sUser = "B1SQLUSER" 
            sUser = oObjGlobal.refDi.SQL.usuarioSQL()
            'sPwd = "12eXo$18" 
            sPwd = oObjGlobal.refDi.SQL.claveSQL()

            oCRReport.Load(strRutaInforme)

            If oObjGlobal.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                If Right(oObjGlobal.pathDLL, 6).ToUpper = "DLL_64" Then
                    sDriver = "{B1CRHPROXY}"
                Else
                    sDriver = "{B1CRHPROXY32}"
                End If
                oCRReport.ApplyNewServer(sDriver, sServer, sUser, sPwd, sBBDD)
            Else
                For Idx = 0 To oCRReport.DataSourceConnections.Count - 1
                    oCRReport.DataSourceConnections(Idx).SetConnection(sServer, sBBDD, False)
                    oCRReport.DataSourceConnections(Idx).SetLogon(sUser, sPwd)
                Next

                For Idx = 0 To oCRReport.Subreports.Count - 1
                    For Idx2 = 0 To oCRReport.Subreports.Item(Idx).DataSourceConnections.Count - 1
                        oCRReport.Subreports(Idx).DataSourceConnections(Idx2).SetConnection(sServer, sBBDD, False)
                        oCRReport.Subreports(Idx).DataSourceConnections(Idx2).SetLogon(sUser, sPwd)
                    Next
                Next
            End If

            oCRReport.SetParameterValue("DocKey@", docentry)
            oCRReport.SetParameterValue("Schema@", sBBDD)
            'Dim conrepor As CrystalDecisions.Shared.DataSourceConnections = oCRReport.DataSourceConnections
            ''conrepor(0).SetConnection(oCompany.Server.ToString, oCompany.CompanyDB.ToString, oCompany.DbUserName, sBDPassword)
            'conrepor(0).SetConnection(oCompany.Server.ToString, oCompany.CompanyDB.ToString, vbTrue)
            sFilePDF = sRutaFicheros & strNombrePDF & ".pdf"
            NombreDocumentoPdf = sFilePDF

            If IO.File.Exists(NombreDocumentoPdf) Then
                IO.File.Delete(NombreDocumentoPdf)
            End If
            oObjGlobal.SBOApp.StatusBar.SetText("Generando pdf para envio impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oCRReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, NombreDocumentoPdf)

            oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Pdf creado : " & strNombrePDF, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            GenerarPDFyEnviarCrystal = True

        Catch ex As Exception
            sMensaje = "Crear PDF: " + ex.Message
            oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oObjGlobal.SBOApp.MessageBox(sMensaje)
        Finally
            oCRReport.Close()
            oCRReport.Dispose()
            GC.Collect()
        End Try
    End Function
End Class
