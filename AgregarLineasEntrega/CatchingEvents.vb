Imports System.Windows.Forms

Friend Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF

    Public Sub New()
        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        addMenuItems()

        setFilters()

    End Sub

    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit()
            End
        End Try
    End Sub

    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
            'Finally
        End Try
    End Sub

    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
        Finally
            loRecSet = Nothing
        End Try
    End Sub


    Private Sub addMenuItems()
        Dim loForm As SAPbouiCOM.Form = Nothing
        Dim loMenus As SAPbouiCOM.Menus
        Dim loMenusRoot As SAPbouiCOM.Menus
        Dim loMenuItem As SAPbouiCOM.MenuItem

        Try
            '////// Obtiene referencia de la forma Principal de Modulos
            loForm = SBOApplication.Forms.GetForm(169, 1)

            loForm.Freeze(True)

            '////// Obtiene la referencia de los Menus de SBO
            loMenus = SBOApplication.Menus.Item(6).SubMenus

            '////// Adiciona un Nuevo Menu para la Aplicacion de VectorSBO
            If loMenus.Exists("DEL01") Then
                loMenus.RemoveEx("DEL01")
            End If

            loMenuItem = loMenus.Add("DEL01", "Entregas", SAPbouiCOM.BoMenuType.mt_POPUP, loMenus.Count)

            loMenusRoot = loMenuItem.SubMenus

            '////// Adiciona un menu Item
            If loMenusRoot.Exists("DEL11") Then
                loMenusRoot.RemoveEx("DEL11")
            End If
            If loMenusRoot.Exists("DEL12") Then
                loMenusRoot.RemoveEx("DEL12")
            End If
            loMenuItem = loMenusRoot.Add("DEL11", "Escanear Facturas", SAPbouiCOM.BoMenuType.mt_STRING, loMenusRoot.Count)
            loMenuItem = loMenusRoot.Add("DEL12", "Buscar Facturas", SAPbouiCOM.BoMenuType.mt_STRING, loMenusRoot.Count)
            loMenus = loMenuItem.SubMenus

            loForm.Freeze(False)
            loForm.Update()

        Catch ex As Exception
            If (Not loForm Is Nothing) Then
                loForm.Freeze(False)
                loForm.Update()
            End If
            SBOApplication.MessageBox("CatchingEvents. Error al agregar las opciones del menú. " & ex.Message)
            End
        Finally
            loMenus = Nothing
            loMenusRoot = Nothing
            loMenuItem = Nothing
        End Try
    End Sub


    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try

            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx("tekDelivery") '////// FORMA UDO DE ENTREGAS
            lofilter.AddEx("tekSearchDev") '////// FORMA UDO DE ENTREGAS
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
            lofilter.AddEx("tekDelivery") '////// FORMA UDO DE ENTREGAS
            lofilter.AddEx("tekSearchDev") '////// FORMA UDO DE ENTREGAS
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            lofilter.AddEx("tekDelivery") '////// FORMA UDO DE ENTREGAS
            lofilter.AddEx("tekSearchDev") '////// FORMA UDO DE ENTREGAS
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub


    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS MENU
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.MenuEvent
        Dim otekDel As FrmtekDel
        Dim otekSch As FrmtekSch

        Try
            '//ANTES DE PROCESAR SBO
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    '//////////////////////////////////SubMenu de Crear traslado inventario////////////////////////
                    Case "DEL11"

                        otekDel = New FrmtekDel
                        otekDel.openForm(csDirectory)

                    Case "DEL12"

                        otekSch = New FrmtekSch
                        otekSch.openForm(csDirectory)

                End Select
            End If

        Catch ex As Exception
            SBOApplication.MessageBox("clsCatchingEvents. MenuEvent " & ex.Message)
        Finally
            'oReservaPedido = Nothing
        End Try
    End Sub


    Private Sub SBOApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select
    End Sub


    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent
        Try
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then

                Select Case pVal.FormTypeEx
                    '////////////////FORMA PARA ACTIVAR LICENCIA
                    Case "tekDelivery"
                        FrmEntregaSBOControllerAfter(FormUID, pVal)

                    Case "tekSearchDev"
                        FrmSearchSBOControllerAfter(FormUID, pVal)

                End Select
            End If

        Catch ex As Exception
            SBOApplication.MessageBox("SBOApplication_ItemEvent. ItemEvent " & ex.Message)
        Finally
        End Try
    End Sub


    Private Sub FrmEntregaSBOControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim oDeliveries As Deliveries
        Dim oGrid As SAPbouiCOM.Grid
        Dim oForm As SAPbouiCOM.Form
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        oForm = SBOApplication.Forms.Item(FormUID)
        oGrid = oForm.Items.Item("11").Specific

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case "12"

                            oDeliveries = New Deliveries
                            oDeliveries.addDelivery(FormUID, csDirectory)

                    End Select

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                    Select Case pVal.ItemUID

                        Case "11"

                            Select Case pVal.CharPressed

                                Case "9"

                                    Select Case pVal.ColUID

                                        Case "Estatus"

                                            oDataTable = oGrid.DataTable

                                            Lineduplicadas(oDataTable, pVal.Row)
                                            SearchInvoices(oDataTable, pVal.Row)
                                            ExistInvoices(oDataTable, pVal.Row)

                                    End Select

                            End Select

                    End Select

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                    Select Case pVal.ItemUID

                        Case "1"

                            stQueryH = "Select ""U_Truck"" as ""Camion"" from ""@EP_EN2"" where ""Name""='" & oForm.DataSources.UserDataSources.Item("dsDriver").Value & "'"
                            oRecSetH.DoQuery(stQueryH)

                            oForm.DataSources.UserDataSources.Item("dsTruck").Value = oRecSetH.Fields.Item("Camion").Value

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("FrmEntregaSBOControllerAfter. Error en forma de Panel General. " & ex.Message)
        Finally

        End Try
    End Sub


    Private Sub FrmSearchSBOControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim oDeliveries As Deliveries
        Dim oGrid As SAPbouiCOM.Grid
        Dim oForm As SAPbouiCOM.Form
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        oForm = SBOApplication.Forms.Item(FormUID)
        oGrid = oForm.Items.Item("11").Specific

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case "12"

                            oDeliveries = New Deliveries
                            oDeliveries.SearchDeliveries(FormUID)

                        Case "13"

                            oDeliveries = New Deliveries
                            oDeliveries.updateDelivery(FormUID, csDirectory)

                        Case "14"

                            oDeliveries = New Deliveries
                            oDeliveries.BeforeAndAfter(FormUID, 1)

                        Case "16"

                            oDeliveries = New Deliveries
                            oDeliveries.BeforeAndAfter(FormUID, 2)

                    End Select

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                    Select Case pVal.ItemUID

                        Case "11"

                            Select Case pVal.CharPressed

                                Case "9"

                                    Select Case pVal.ColUID

                                        Case "Estatus"

                                            oDataTable = oGrid.DataTable

                                            Lineduplicadas(oDataTable, pVal.Row)
                                            SearchInvoices(oDataTable, pVal.Row)
                                            ExistInvoices(oDataTable, pVal.Row)

                                    End Select

                            End Select

                    End Select

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                    Select Case pVal.ItemUID

                        Case "1"

                            stQueryH = "Select ""U_Truck"" as ""Camion"" from ""@EP_EN2"" where ""Name""='" & oForm.DataSources.UserDataSources.Item("dsDriver").Value & "'"
                            oRecSetH.DoQuery(stQueryH)

                            oForm.DataSources.UserDataSources.Item("dsTruck").Value = oRecSetH.Fields.Item("Camion").Value

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("FrmEntregaSBOControllerAfter. Error en forma de Panel General. " & ex.Message)
        Finally

        End Try
    End Sub


    Public Function Lineduplicadas(ByVal oDataTable As SAPbouiCOM.DataTable, ByVal Limite As Integer)

        Dim Invoice As String
        Dim LineaG As Integer

        Try

            Invoice = Nothing

            If oDataTable.GetValue("Factura", Limite) Is Nothing Or oDataTable.GetValue("Factura", Limite) = "" Then

            Else

                Invoice = oDataTable.GetValue("Factura", Limite)

                For i = 0 To Limite

                    If oDataTable.GetValue("Factura", i) = Invoice And i <> Limite Then

                        LineaG = oDataTable.GetValue("#", i)
                        SBOApplication.MessageBox("La Factura " & Invoice & " esta siendo duplicada en la linea " & LineaG)

                    End If

                Next

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo la funcion Lineduplicadas: " & ex.Message)

        End Try

    End Function


    Public Function SearchInvoices(ByVal oDataTable As SAPbouiCOM.DataTable, ByVal Limite As Integer)

        Dim stQueryH, stQueryH2, stQueryH3 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim Factura, Entrega, Factura2 As String

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If oDataTable.GetValue("Factura", Limite) Is Nothing Or oDataTable.GetValue("Factura", Limite) = "" Then

            Else

                Factura = oDataTable.GetValue("Factura", Limite)
                stQueryH = "Select ""U_Delivery"" as ""DocEntry"" from ""@EP_EN1"" where ""U_DocNum""=" & Factura & " and ""U_Status""<>'Cambio'"
                oRecSetH.DoQuery(stQueryH)

                stQueryH2 = "Select ""DocNum"" as ""DocEntry"" from ""@EP_EN"" where ""U_DocNum""=" & Factura & " and ""Canceled""='N'"
                oRecSetH2.DoQuery(stQueryH2)

                stQueryH3 = "Select T7.""DocNum"" as ""Factura"", T8.""U_Delivery"" as ""DocEntry"" 
                            from ""OINV"" T0
                            Inner Join ""INV1"" T1 on T1.""DocEntry""=T0.""DocEntry""
                            Inner Join ""RDR1"" T2 on T2.""DocEntry""=T1.""BaseEntry"" and T2.""ObjType""=T1.""BaseType"" and T2.""LineNum""=T1.""BaseLine"" and T2.""ItemCode""=T1.""ItemCode""
                            Inner Join ""ORDR"" T3 on T3.""DocEntry""=T2.""DocEntry""
                            Inner Join ""ORIN"" T4 on T4.""U_ORDR""=T3.""DocNum""
                            Inner Join ""RIN1"" T5 on T5.""DocEntry""=T4.""DocEntry""
                            Inner Join ""INV1"" T6 on T6.""DocEntry""=T5.""BaseEntry"" and T6.""ObjType""=T5.""BaseType"" and T6.""LineNum""=T5.""BaseLine"" and T6.""ItemCode""=T5.""ItemCode""
                            Inner Join ""OINV"" T7 on T7.""DocEntry""=T6.""DocEntry""
                            Inner Join ""@EP_EN1"" T8 on T8.""U_DocNum""=T7.""DocNum""
                            where T0.""DocNum""=" & Factura
                oRecSetH3.DoQuery(stQueryH3)

                If oRecSetH.RecordCount > 0 Then

                    Entrega = oRecSetH.Fields.Item("DocEntry").Value
                    SBOApplication.MessageBox("La factura " & Factura & " ya fue registrada en la entrega " & Entrega & ".")

                End If

                If oRecSetH2.RecordCount > 0 Then

                    Entrega = oRecSetH2.Fields.Item("DocEntry").Value
                    SBOApplication.MessageBox("La factura " & Factura & " ya fue registrada en la entrega " & Entrega & " del complemento pasado.")

                End If

                If oRecSetH3.RecordCount > 0 Then

                    Entrega = oRecSetH3.Fields.Item("DocEntry").Value
                    Factura2 = oRecSetH3.Fields.Item("Factura").Value
                    SBOApplication.MessageBox("La factura " & Factura & " esta sustituyendo a la factura " & Factura2 & " la cual ya fue registrada y entregada en el numero de entrega " & Entrega & ".")

                End If

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo la funcion SearchInvoices: " & ex.Message)

        End Try

    End Function


    Public Function ExistInvoices(ByVal oDataTable As SAPbouiCOM.DataTable, ByVal Limite As Integer)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim Factura As String

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If oDataTable.GetValue("Factura", Limite) Is Nothing Or oDataTable.GetValue("Factura", Limite) = "" Then

            Else

                Factura = oDataTable.GetValue("Factura", Limite)
                stQueryH = "Select ""DocNum"",""U_CSM_SFAC"" from OINV where ""DocNum""=" & Factura & " and ""CANCELED""='N'"
                oRecSetH.DoQuery(stQueryH)

                If oRecSetH.RecordCount = 0 Then

                    SBOApplication.MessageBox("La factura " & Factura & " no esta registrada en el sistema.")

                ElseIf oRecSetH.Fields.Item("U_CSM_SFAC").Value = "Pendiente" Then

                    SBOApplication.MessageBox("La factura " & Factura & " aun esta registrada en el sistema con el status ""Pendiente"".")

                End If

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo la funcion ExistInvoices: " & ex.Message)

        End Try

    End Function


End Class
