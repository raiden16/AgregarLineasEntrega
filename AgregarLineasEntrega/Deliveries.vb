Public Class Deliveries

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim oInvoice As SAPbobsCOM.Documents
    Dim Duplicadas, Registradas, NExist As Integer

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub


    Public Function addDelivery(ByVal FormUID As String, ByVal csDirectory As String)

        Dim otekDel As FrmtekDel
        Dim coForm As SAPbouiCOM.Form
        Dim oGrid As SAPbouiCOM.Grid
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim stQueryH, stQueryH2, stQueryH3 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim Driver, User, Truck, DocNum, Invoice, Estatus, IFechas, DFechas, SFechas As String
        Dim DFecha, IFecha, SFecha As Date
        Dim Code, Linea As Integer

        Try

            coForm = SBOApplication.Forms.Item(FormUID)
            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Driver = coForm.DataSources.UserDataSources.Item("dsDriver").Value
            User = coForm.DataSources.UserDataSources.Item("dsUser").Value
            Truck = coForm.DataSources.UserDataSources.Item("dsTruck").Value
            DocNum = coForm.DataSources.UserDataSources.Item("dsDocN").Value
            DFecha = coForm.DataSources.UserDataSources.Item("dsDate").Value
            DFechas = DFecha.Year & "-" & DFecha.Month & "-" & DFecha.Day

            oGrid = coForm.Items.Item("11").Specific
            oDataTable = oGrid.DataTable

            Duplicadas = 0
            Registradas = 0
            NExist = 0
            Linea = 0

            Lineduplicadas(oDataTable)
            SearchInvoices(oDataTable)
            ExistInvoices(oDataTable)

            If Duplicadas = 0 And Registradas = 0 And NExist = 0 Then

                For i = 0 To oDataTable.Rows.Count - 1

                    If oDataTable.GetValue("Factura", i) Is Nothing Or oDataTable.GetValue("Factura", i) = "" Then

                    Else

                        Linea = Linea + 1

                        stQueryH = "Select count(""Code"")+1 as ""DocEntry"" from ""@EP_EN1"""
                        oRecSetH.DoQuery(stQueryH)
                        Code = oRecSetH.Fields.Item("DocEntry").Value

                        Invoice = oDataTable.GetValue("Factura", i)
                        IFecha = Convert.ToDateTime(oDataTable.GetValue("Fecha Factura", i))
                        SFecha = Convert.ToDateTime(oDataTable.GetValue("Fecha Escaneo", i))
                        Estatus = oDataTable.GetValue("Estatus", i)
                        IFechas = IFecha.Year & "-" & IFecha.Month & "-" & IFecha.Day
                        SFechas = SFecha.Year & "-" & SFecha.Month & "-" & SFecha.Day

                        stQueryH2 = "INSERT INTO ""@EP_EN1"" VALUES (" & Code & "," & Code & "," & DocNum & "," & Linea & "," & Invoice & ",'" & IFechas & "','" & SFechas & "','" & Estatus & "')"
                        oRecSetH2.DoQuery(stQueryH2)

                    End If

                Next

                If Linea > 0 Then

                    stQueryH3 = "INSERT INTO ""@EP_EN0"" VALUES (" & DocNum & "," & DocNum & ",'" & DFechas & "','" & Driver & "','" & User & "','" & Truck & "')"
                    oRecSetH3.DoQuery(stQueryH3)
                    otekDel = New FrmtekDel
                    otekDel.openForm(csDirectory)

                End If

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo creacion de UDO addDelivery: " & ex.Message)

        End Try

    End Function

    Public Function Lineduplicadas(ByVal oDataTable As SAPbouiCOM.DataTable)

        Dim Invoice As String
        Dim Linea, Lineal, LineaG As Integer

        Try

            Linea = 0
            Invoice = Nothing

            For i = 0 To oDataTable.Rows.Count - 1

                Linea = Linea + 1

                If oDataTable.GetValue("Factura", i) Is Nothing Or oDataTable.GetValue("Factura", i) = "" Then

                Else

                    Invoice = oDataTable.GetValue("Factura", i)

                    Lineal = 0

                    For l = 0 To oDataTable.Rows.Count - 1

                        Lineal = Lineal + 1

                        If oDataTable.GetValue("Factura", l) = Invoice And Linea <> Lineal Then

                            LineaG = oDataTable.GetValue("#", l)
                            SBOApplication.MessageBox("La Factura " & Invoice & " Esta siendo duplicada en la entrega, por favor quita la factura duplicada de la linea " & LineaG)
                            Duplicadas = Duplicadas + 1

                        End If

                    Next

                End If

            Next

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo la funcion Lineduplicadas: " & ex.Message)

        End Try

    End Function

    Public Function SearchInvoices(ByVal oDataTable As SAPbouiCOM.DataTable)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim Factura, Entrega As String

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i = 0 To oDataTable.Rows.Count - 1

                Factura = oDataTable.GetValue("Factura", i)

                If oDataTable.GetValue("Factura", i) Is Nothing Or oDataTable.GetValue("Factura", i) = "" Then

                Else

                    stQueryH = "Select ""U_Delivery"" as ""DocEntry"" from ""@EP_EN1"" where ""U_DocNum""=" & Factura & " and ""U_Status""<>'Cambio'"
                    oRecSetH.DoQuery(stQueryH)

                    If oRecSetH.RecordCount > 0 Then

                        Entrega = oRecSetH.Fields.Item("DocEntry").Value
                        SBOApplication.MessageBox("La factura " & Factura & " ya fue registrada en la entrega " & Entrega & ", por favor quita esta factura que se encuentra en la linea " & i + 1 & ".")
                        Registradas = Registradas + 1

                    End If

                End If

            Next

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo la funcion SearchInvoices: " & ex.Message)

        End Try

    End Function


    Public Function ExistInvoices(ByVal oDataTable As SAPbouiCOM.DataTable)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim Factura As String

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i = 0 To oDataTable.Rows.Count - 1

                Factura = oDataTable.GetValue("Factura", i)

                If oDataTable.GetValue("Factura", i) Is Nothing Or oDataTable.GetValue("Factura", i) = "" Then

                Else

                    stQueryH = "Select ""DocNum"",""U_CSM_SFAC"" from OINV where ""DocNum""=" & Factura & " and ""CANCELED""='N'"
                    oRecSetH.DoQuery(stQueryH)

                    If oRecSetH.RecordCount = 0 Then

                        SBOApplication.MessageBox("La factura " & Factura & " no esta registrada en el sistema, por favor quita esta factura que se encuentra en la linea " & i + 1 & ".")
                        NExist = NExist + 1

                    ElseIf oRecSetH.Fields.Item("U_CSM_SFAC").Value = "Pendiente" Then

                        SBOApplication.MessageBox("La factura " & Factura & " aun esta registrada en el sistema con el status ""Pendiente"", esta se encuentra en la linea " & i + 1 & ".")
                        NExist = NExist + 1

                    End If

                End If

            Next

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo la funcion ExistInvoices: " & ex.Message)

        End Try

    End Function


    Public Function SearchDeliveries(ByVal FormUID As String)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim oGrid As SAPbouiCOM.Grid
        Dim oForm As SAPbouiCOM.Form
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim Factura, Entrega As String

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oForm = SBOApplication.Forms.Item(FormUID)
            oGrid = oForm.Items.Item("11").Specific
            oDataTable = oGrid.DataTable

            For i = 0 To 0

                If oDataTable.GetValue("Factura", i) Is Nothing Or oDataTable.GetValue("Factura", i) = "" Then

                Else

                    Factura = oDataTable.GetValue("Factura", i)
                    stQueryH = "Select ""U_Delivery"" as ""Entrega"" from ""@EP_EN1"" where ""U_DocNum""=" & Factura
                    oRecSetH.DoQuery(stQueryH)

                    If oRecSetH.RecordCount > 0 Then

                        Entrega = oRecSetH.Fields.Item("Entrega").Value
                        findDelivery(FormUID, Entrega)

                    Else

                        SBOApplication.MessageBox("La factura " & Factura & " aun no ha sido registrada en ninguna entrega.")

                    End If

                End If

            Next

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo la funcion SearchDeliveries: " & ex.Message)

        End Try

    End Function


    Public Function findDelivery(ByVal FormUID As String, ByVal Entrega As String)
        Dim oGrid As SAPbouiCOM.Grid
        Dim coForm As SAPbouiCOM.Form
        Dim oCombo As SAPbouiCOM.ComboBoxColumn
        Dim stQuery, stQuery2 As String
        Dim oRecSet, oRecSet2 As SAPbobsCOM.Recordset
        Dim Linea As Integer

        Try

            coForm = SBOApplication.Forms.Item(FormUID)

            oRecSet = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "Select ""Name"" as ""Entrega"",""U_CreateDate"" as ""Fecha"",""U_Driver"" as ""Chofer"",""U_User"",""U_Truck"" as ""Camion"" from ""@EP_EN0"" where ""Code""=" & Entrega
            oRecSet.DoQuery(stQuery)

            coForm.DataSources.UserDataSources.Item("dsDocN").Value = oRecSet.Fields.Item("Entrega").Value
            coForm.DataSources.UserDataSources.Item("dsDate").Value = oRecSet.Fields.Item("Fecha").Value
            coForm.DataSources.UserDataSources.Item("dsDriver").Value = oRecSet.Fields.Item("Chofer").Value
            coForm.DataSources.UserDataSources.Item("dsUser").Value = oRecSet.Fields.Item("U_User").Value
            coForm.DataSources.UserDataSources.Item("dsTruck").Value = oRecSet.Fields.Item("Camion").Value

            oGrid = coForm.Items.Item("11").Specific

            oRecSet2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery2 = "Select 1 as ""#"",""U_DocNum"" as ""Factura"",""U_DocDate"" as ""Fecha Factura"",""U_ScanDate"" as ""Fecha Escaneo"",""U_Status"" as ""Estatus"" from ""@EP_EN1"" where ""U_Delivery""='" & Entrega & "' order by ""U_LineNum"""
            oGrid.DataTable.ExecuteQuery(stQuery2)
            oRecSet2.DoQuery(stQuery2)

            oGrid.Columns.Item(4).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oCombo = oGrid.Columns.Item(4)

            oCombo.ValidValues.Add("", "")
            oCombo.ValidValues.Add("Escaneado", "Escaneado")
            oCombo.ValidValues.Add("Cambio", "Cambio")
            oCombo.ValidValues.Add("Retenido", "Retenido")
            oCombo.ValidValues.Add("Cancelado", "Cancelado")

            Linea = 20 - oRecSet2.RecordCount
            oGrid.DataTable.Rows.Add(Linea)

            For i = 1 To 20 - 1
                oGrid.DataTable.SetValue("#", i, i + 1)
            Next

            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False

        Catch ex As Exception

            MsgBox("findDelivery. fallo la carga de los valores de la factura: " & ex.Message)

        End Try

    End Function


    Public Function updateDelivery(ByVal FormUID As String, ByVal csDirectory As String)

        Dim otekDel As FrmtekDel
        Dim coForm As SAPbouiCOM.Form
        Dim oGrid As SAPbouiCOM.Grid
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim stQueryH2, stQueryH3 As String
        Dim oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim Driver, Truck, DocNum, Invoice, Estatus, Orden As String
        Dim Linea As Integer

        Try

            coForm = SBOApplication.Forms.Item(FormUID)
            oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Driver = coForm.DataSources.UserDataSources.Item("dsDriver").Value
            Truck = coForm.DataSources.UserDataSources.Item("dsTruck").Value
            DocNum = coForm.DataSources.UserDataSources.Item("dsDocN").Value

            oGrid = coForm.Items.Item("11").Specific
            oDataTable = oGrid.DataTable

            Linea = 0

            For i = 0 To oDataTable.Rows.Count - 1

                If oDataTable.GetValue("Factura", i) Is Nothing Or oDataTable.GetValue("Factura", i) = "" Then

                Else

                    Linea = Linea + 1

                    Invoice = oDataTable.GetValue("Factura", i)
                    Estatus = oDataTable.GetValue("Estatus", i)
                    Orden = oDataTable.GetValue("#", i)

                    stQueryH2 = "UPDATE ""@EP_EN1"" SET ""U_Status""='" & Estatus & "' where ""U_Delivery""=" & DocNum & " and ""U_LineNum""=" & Orden & " and ""U_DocNum""=" & Invoice
                    oRecSetH2.DoQuery(stQueryH2)

                End If

            Next

            If Linea > 0 Then

                stQueryH3 = "UPDATE ""@EP_EN0"" SET ""U_Driver""='" & Driver & "' where ""Code""=" & DocNum
                oRecSetH3.DoQuery(stQueryH3)
                stQueryH3 = "UPDATE ""@EP_EN0"" SET ""U_Truck""='" & Truck & "' where ""Code""=" & DocNum
                oRecSetH3.DoQuery(stQueryH3)
                otekDel = New FrmtekDel
                otekDel.openForm(csDirectory)

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo creacion de UDO updateDelivery: " & ex.Message)

        End Try

    End Function


    Public Function BeforeAndAfter(ByVal FormUID As String, ByVal Action As Integer)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim oForm As SAPbouiCOM.Form
        Dim Entrega As String

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oForm = SBOApplication.Forms.Item(FormUID)

            If Action = 1 Then

                Entrega = oForm.DataSources.UserDataSources.Item("dsDocN").Value - 1

            ElseIf Action = 2 Then

                Entrega = oForm.DataSources.UserDataSources.Item("dsDocN").Value + 1

            End If

            stQueryH = "Select * from ""@EP_EN0"" where ""Code""=" & Entrega
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                findDelivery(FormUID, Entrega)

            End If


        Catch ex As Exception

            SBOApplication.MessageBox("Deliveries, fallo la funcion BeforeAndAfter: " & ex.Message)

        End Try

    End Function


End Class
