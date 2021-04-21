Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_IPCG
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
        End If
        cargamenu()
    End Sub
    Private Sub cargaCampos()
        Dim bValidado As Boolean = False
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDF_EXO_ODRF.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDF_EXO_ODRF", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            bValidado = objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            If bValidado = True Then
                objGlobal.SBOApp.StatusBar.SetText("Validado: UDF_EXO_ODRF", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'Else
                '    objGlobal.SBOApp.StatusBar.SetText("No Validado: UDF_EXO_ODRF", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End If
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults

    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnIPCG"
                        'Cargamos pantalla de gestión.
                        If CargarForm() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarForm = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_IPCG.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            'CType(oForm.Items.Item("chkEmp").Specific, SAPbouiCOM.CheckBox).Checked = True

            CargarForm = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IPCG"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IPCG"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IPCG"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IPCG"
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
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim dFecha As Date = Now.Date
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                sSQL = "SELECT * FROM ""@EXO_IPCL"" "
                sSQL &= " WHERE ""U_EXO_FECHA""<='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' Order BY ""U_EXO_FECHA"" desc;"
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    'Ponemos la fecha 
                    dFecha = oRs.Fields.Item("U_EXO_FECHA").Value.ToString
                    oForm.DataSources.UserDataSources.Item("UD_FU").Value = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                    oForm.DataSources.UserDataSources.Item("UD_PU").Value = oRs.Fields.Item("U_EXO_VALOR").Value.ToString
                End If
            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sAnno As String = ""
        Dim dFecha As Date = Now.Date
        Dim dFechaAct As Date = Now.Date
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim bActualizaDocumento As Boolean = False
        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "1"
                    sAnno = Year(oForm.DataSources.UserDataSources.Item("UD_FN").Value).ToString("0000")
                    dFecha = oForm.DataSources.UserDataSources.Item("UD_FN").Value
                    'Buscamos si existe la fecha para no insertarla
                    sSQL = "SELECT * FROM ""@EXO_IPCL""  WHERE ""U_EXO_FECHA""='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "'"
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount = 0 Then
                        'No existe insertamos el dato
                        oDI_COM = objGlobal.refDi.dameEXO_UDOEntity("EXO_IPC")
                        oDI_COM.GetByKey(sAnno)
                        oDI_COM.GetNewChild("EXO_IPCL")
                        oDI_COM.SetValueChild("U_EXO_FECHA") = dFecha
                        oDI_COM.SetValueChild("U_EXO_VALOR") = oForm.DataSources.UserDataSources.Item("UD_PN").Value.Replace(".", "").Replace(",", ".")
                        oDI_COM.UDO_Update()
                        objGlobal.SBOApp.StatusBar.SetText("Dato Guardado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        If dFecha <= dFechaAct Then
                            bActualizaDocumento = True
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("La fecha es superior a la del día de hoy, por lo que no se actualiza el documento hoy.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            bActualizaDocumento = False
                        End If

                    Else
                        'Como existe mostramos mensaje
                        Dim sTexto As String = "La fecha que intenta actualizar ya existe - " & dFecha.Day.ToString("00") & "/" & dFecha.Month.ToString("00") & "/" & dFecha.Year.ToString("0000") & " - con un valor - " & oRs.Fields.Item("U_EXO_VALOR").Value.ToString
                        objGlobal.SBOApp.StatusBar.SetText(sTexto, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If objGlobal.SBOApp.MessageBox(sTexto & ". ¿Desea actualizar por el dato nuevo?", 1, "Sí", "No") = 1 Then
                            objGlobal.SBOApp.StatusBar.SetText("Se procede a actualizar el dato....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            sSQL = "UPDATE ""@EXO_IPCL"" SET ""U_EXO_VALOR""=" & oForm.DataSources.UserDataSources.Item("UD_PN").Value.Replace(".", "").Replace(",", ".")
                            sSQL &= " WHERE ""Code""='" & sAnno & "' And ""U_EXO_FECHA""='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "'"
                            oRs.DoQuery(sSQL)
                            objGlobal.SBOApp.StatusBar.SetText("Dato Guardado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            If dFecha <= dFechaAct Then
                                bActualizaDocumento = True
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("La fecha es superior a la del día de hoy, por lo que no se actualiza el documento hoy.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                bActualizaDocumento = False
                            End If
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("El dato no se ha actualizado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            bActualizaDocumento = False
                        End If
                    End If
                    'Buscamos el borrador para actualizarlo
                    If bActualizaDocumento = True Then
                        objGlobal.SBOApp.StatusBar.SetText("Se va a proceder a actualizar el IPC de documento(s) borrador...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        sSQL = "SELECT D.""DocNum"",D.""DocEntry"" FROM ""ODRF"" D INNER JOIN ""ORCP"" O ON D.""DocEntry""=O.""DraftEntry"" "
                        sSQL &= " WHERE D.""ObjType""='18' and D.""U_EXO_IPC""='Y'"
                        oRs.DoQuery(sSQL)
                        If oRs.RecordCount > 0 Then
                            For D = 0 To oRs.RecordCount - 1
                                objGlobal.SBOApp.StatusBar.SetText("Actualizando IPC """ & oForm.DataSources.UserDataSources.Item("UD_PN").Value.ToString & """ del documento:  " & oRs.Fields.Item("DocNum").Value.ToString & " DocEntry:" & oRs.Fields.Item("DocEntry").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                oDoc = objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts) 'Documento borrador
                                oDoc.GetByKey(oRs.Fields.Item("DocEntry").Value.ToString)
                                For I = 0 To oDoc.Lines.Count - 1
                                    oDoc.Lines.SetCurrentLine(I)
                                    Dim dPrecio As Double = oDoc.Lines.UnitPrice + ((oDoc.Lines.UnitPrice * oForm.DataSources.UserDataSources.Item("UD_PN").Value) / 100)
                                    objGlobal.SBOApp.StatusBar.SetText("Actualizando linea " & oDoc.Lines.LineNum.ToString & ". Precio anterior: " & oDoc.Lines.Price.ToString & " . Precio Actualizado:" & dPrecio.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    oDoc.Lines.UnitPrice = dPrecio
                                Next
                                If oDoc.Update() <> 0 Then 'Si ocurre un error en la grabación entra
                                    Dim sErrorDes As String = objGlobal.compañia.GetLastErrorCode & " " & objGlobal.compañia.GetLastErrorDescription
                                    objGlobal.SBOApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            Next

                            'Se actualiza la fecha indicando que si se ha realizado la actualización del IPC en documento(s)
                            sSQL = "UPDATE ""@EXO_IPCL"" SET ""U_EXO_ACT""='Y' "
                            sSQL &= " WHERE ""Code""='" & sAnno & "' And ""U_EXO_FECHA""='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "'"
                            oRs.DoQuery(sSQL)
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado ningún Borrador de facturas de compras para Actualizar el IPC.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("El documento borrador no se ha actualizado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox("El documento borrador no se ha actualizado.")
                    End If
                Case "btnHCO"
                    Dim sAnnoAct As String = Now.Year
                    objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_IPC", sAnnoAct)
            End Select

            EventHandler_ItemPressed_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class
