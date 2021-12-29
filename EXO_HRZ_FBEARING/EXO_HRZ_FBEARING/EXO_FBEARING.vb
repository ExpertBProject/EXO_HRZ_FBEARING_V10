Imports SAPbouiCOM
Imports System.Xml
Imports System.IO

Public Class EXO_FBEARING
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        cargamenu()
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        Path = objGlobal.refDi.OGEN.pathGeneral.ToString.Trim
        If objGlobal.SBOApp.Menus.Exists("EXO-MnGBEA") = True Then
            Path &= "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnGBEA.png") = True Then
                    objGlobal.SBOApp.Menus.Item("EXO-MnGBEA").Image = Path & "\MnGBEA.png"
                End If
            End If
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnGBEA"
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
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim Path As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarForm = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_FBEARING.srf")
            oFP.XmlData = oFP.XmlData.Replace("modality=""0""", "modality=""1""")

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

            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_FBEARING"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_FBEARING"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_FBEARING"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_FBEARING"
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
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sPath As String = "" : Dim sRutaFich As String = "" : Dim sArchivo As String = ""
        Dim sLinea As String = ""
        EventHandler_ItemPressed_After = False
        Dim sSQL As String = "" : Dim OdtStock As System.Data.DataTable = Nothing
        Dim sMensaje As String = ""
        Try

            Select Case pVal.ItemUID
                Case "btn_Fich"
#Region "Pedimos directorio para guardar donde quiere el usuario"
                    'pedimos directorio para guardar
                    If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
                        sPath = objGlobal.SBOApp.SendFileToBrowser(IO.Path.GetFileName(sRutaFich))
                    Else
                        sPath = objGlobal.funciones.SaveDialogFiles("Guardar archivo como", "Fichero TXT|*.txt", IO.Path.GetFileName(sRutaFich), Environment.SpecialFolder.Desktop)
                    End If
                    If sPath.Trim <> "" Then
                        oForm.DataSources.UserDataSources.Item("UDFICH").Value = sPath
                    Else
                        objGlobal.SBOApp.MessageBox("No ha indicado un directorio para guardar.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha indicado un directorio para guardar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
#End Region
                Case "btnGen"
                    If oForm.DataSources.UserDataSources.Item("UDFICH").Value.ToString.Trim <> "" Then
#Region "Generar fichero"
                        sPath = oForm.DataSources.UserDataSources.Item("UDFICH").Value.ToString.Trim
#Region "Se genera el fichero en el servidor"
                        sRutaFich = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\FBEARING\"

                        If IO.Directory.Exists(sRutaFich.Trim) = False Then
                            IO.Directory.CreateDirectory(sRutaFich.Trim)
                        End If
                        sArchivo = IO.Path.GetFileName(sPath.Trim)
                        sRutaFich &= sArchivo.Trim
                        If IO.File.Exists(sRutaFich) = True Then
                            IO.File.Delete(sRutaFich)
                        End If
                        FileOpen(1, sRutaFich, OpenMode.Output)
                        If Not My.Computer.FileSystem.FileExists(sRutaFich) Then
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Compruebe la ruta seleccionada - " & sRutaFich & " - para la creación del fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Function
                        End If
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Generando fichero: " & sRutaFich, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        sLinea = "HURYZA" & ChrW(9) & "T933-GM4B-F6NF-PRSD" & ChrW(9) & "1" & ChrW(9) & "paula@rodamientos-huryza.com"
                        PrintLine(1, sLinea)
#Region "Detalle"
                        sSQL = "SELECT I.""ItemName"", ifnull(I.""U_stec_marcas"",'') ""MARCA"",Cast(cast(S.""STOCK"" as integer) as varchar)""STOCK"",ifnull(B.""U_stec_tfe"",'') ""U_stec_tfe"" "
                        sSQL &= " From OITM I "
                        sSQL &= " INNER JOIN (SELECT ""ItemCode"", sum(""OnHand"") ""STOCK"" FROM OITW GROUP BY ""ItemCode"")S ON S.""ItemCode""= I.""ItemCode"" "
                        sSQL &= " Left JOIN OITB B ON I.""ItmsGrpCod""=B.""ItmsGrpCod"" "
                        sSQL &= " WHERE I.""QryGroup2""='Y' and I.""validFor""='Y' "
                        OdtStock = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                        If OdtStock.Rows.Count > 0 Then
                            Try
                                For Each dr In OdtStock.Rows
                                    sLinea = dr.Item("ItemName").ToString & ChrW(9) & dr.Item("MARCA").ToString & ChrW(9) & dr.Item("STOCK").ToString
                                    sLinea &= ChrW(9) & dr.Item("ItemName").ToString & ChrW(9) & dr.Item("U_stec_tfe").ToString
                                    PrintLine(1, sLinea)
                                Next
                                OdtStock = Nothing
                            Catch ex As Exception

                            End Try
                        Else
                            sMensaje = "(EXO) - No existen datos para insertar en el fichero."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        End If
#End Region
                        FileClose(1)
#End Region
#Region "Se copia al directorio que se haya pedido"
                            Copia_Seguridad(sRutaFich, sPath)
#End Region
#End Region
                        Else
                            objGlobal.SBOApp.MessageBox("No ha indicado un directorio para guardar.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha indicado un directorio para guardar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If

            End Select
            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub Copia_Seguridad(ByVal sArchivoOrigen As String, ByVal sArchivo As String)
        'Comprobamos el directorio de copia que exista
        Dim sPath As String = ""
        sPath = IO.Path.GetDirectoryName(sArchivo)
        If IO.Directory.Exists(sPath) = False Then
            IO.Directory.CreateDirectory(sPath)
        End If
        If IO.File.Exists(sArchivo) = True Then
            IO.File.Delete(sArchivo)
        End If
        'Subimos el archivo
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
            Dim fs As FileStream = New FileStream(sArchivoOrigen, FileMode.Open, FileAccess.Read)
            Dim b(CInt(fs.Length() - 1)) As Byte
            fs.Read(b, 0, b.Length)
            fs.Close()
            Dim fs2 As New System.IO.FileStream(sArchivo, IO.FileMode.Create, IO.FileAccess.Write)
            fs2.Write(b, 0, b.Length)
            fs2.Close()
        Else
            My.Computer.FileSystem.CopyFile(sArchivoOrigen, sArchivo)
        End If
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fichero guardado: " & sArchivo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        objGlobal.SBOApp.MessageBox("Fichero guardado." & sArchivo)
    End Sub

End Class
