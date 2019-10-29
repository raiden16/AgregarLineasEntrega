Imports System.Drawing

Public Class FrmtekDel

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    'Private Property stRuta As String

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function openForm(ByVal psDirectory As String)
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        'Dim Monto As Integer

        Try

            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            csFormUID = "tekDelivery"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")
            End If

            '--- Referencia de Forma
            setForm(csFormUID)

            cargarComboChofi()

            AgregarLineas()

            '---- refresca forma
            coForm.Refresh()
            coForm.Visible = True

            coForm = cSBOApplication.Forms.Item("tekDelivery")
            coForm.DataSources.UserDataSources.Item("dsUser").Value = cSBOCompany.UserName
            coForm.DataSources.UserDataSources.Item("dsDate").Value = Now.Date
            coForm.DataSources.UserDataSources.Item("dsTruck").Value = ""

            stQueryH = "Select count(""Code"")+1 as ""DocEntry"" from ""@EP_EN0"""
            oRecSetH.DoQuery(stQueryH)

            coForm.DataSources.UserDataSources.Item("dsDocN").Value = oRecSetH.Fields.Item("DocEntry").Value

            coForm.Items.Item("2").Enabled = False
            coForm.Items.Item("4").Enabled = False
            coForm.Items.Item("5").Enabled = False

            Return Monto

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmtekDel. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- CIERRA LA VENTANA
    Public Function close() As Integer
        close = 0
        coForm.Close()
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function setForm(ByVal psFormUID As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources()
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function


    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    Private Function getUserDataSources() As Integer
        'Dim llIndice As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA
                getUserDataSources = bindUserDataSources()
            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function


    '//----- ASOCIA LOS USERDATA A ITEMS
    Private Function bindUserDataSources() As Integer
        Dim loText As SAPbouiCOM.EditText
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid
        Dim oCombo As SAPbouiCOM.ComboBox

        Try
            bindUserDataSources = 0

            loDS = coForm.DataSources.UserDataSources.Add("dsDriver", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            oCombo = coForm.Items.Item("1").Specific  'identifico mi combobox
            oCombo.DataBind.SetBound(True, "", "dsDriver")   ' uno mi userdatasources a mi combobox

            loDS = coForm.DataSources.UserDataSources.Add("dsUser", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("2").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsUser")   ' uno mi userdatasources a mi caja de texto

            loDS = coForm.DataSources.UserDataSources.Add("dsTruck", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("3").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsTruck")   ' uno mi userdatasources a mi caja de texto

            loDS = coForm.DataSources.UserDataSources.Add("dsDocN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("4").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsDocN")   ' uno mi userdatasources a mi caja de texto

            loDS = coForm.DataSources.UserDataSources.Add("dsDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("5").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsDate")   ' uno mi userdatasources a mi caja de texto

            oGrid = coForm.Items.Item("11").Specific
            oDataTable = coForm.DataSources.DataTables.Add("Invoices")
            oGrid.DataTable = oDataTable

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function


    '---- Carga de Porcentajes
    Public Function cargarComboChofi()

        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecSet As SAPbobsCOM.Recordset

        Try
            cargarComboChofi = 0
            '--- referencia de combo 
            oCombo = coForm.Items.Item("1").Specific
            coForm.Freeze(True)
            '---- SI YA SE TIENEN VALORES, SE ELIMMINAN DEL COMBO
            If oCombo.ValidValues.Count > 0 Then
                Do
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Loop While oCombo.ValidValues.Count > 0
            End If
            '--- realizar consulta
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSet.DoQuery(" Select null as ""Name"",null as ""Code"" from dummy union all Select ""Name"",""Code"" from ""@EP_EN2""")
            '---- cargamos resultado
            oRecSet.MoveFirst()
            Do While oRecSet.EoF = False
                oCombo.ValidValues.Add(oRecSet.Fields.Item(0).Value, oRecSet.Fields.Item(1).Value)
                oRecSet.MoveNext()
            Loop
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            coForm.Freeze(False)


        Catch ex As Exception
            coForm.Freeze(False)
            MsgBox("FrmtekDel. Fallo la carga previa del comboBox cargarComboChofi: " & ex.Message)
        Finally
            oCombo = Nothing
            oRecSet = Nothing
        End Try
    End Function


    '----- carga los procesos de carga
    Public Function AgregarLineas()
        Dim oGrid As SAPbouiCOM.Grid
        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim oCombo As SAPbouiCOM.ComboBoxColumn

        Try

            oGrid = coForm.Items.Item("11").Specific
            oGrid.DataTable.Clear()

            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "Select 1 as ""#"",'                         ' as ""Factura"", '          ' as ""Fecha Factura"", '          ' as ""Fecha Escaneo"", '          ' as ""Estatus"" from dummy"
            oGrid.DataTable.ExecuteQuery(stQuery)

            oGrid.Columns.Item(4).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oCombo = oGrid.Columns.Item(4)

            oCombo.ValidValues.Add("", "")
            oCombo.ValidValues.Add("Escaneado", "Escaneado")
            oCombo.ValidValues.Add("Retenido", "Retenido")
            oCombo.ValidValues.Add("Cancelado", "Cancelado")

            oGrid.Columns.Item(1).Editable = True
            oGrid.Columns.Item(2).Editable = True
            oGrid.Columns.Item(3).Editable = True

            oGrid.DataTable.Rows.Add(19)

            For i = 1 To 19
                oGrid.DataTable.SetValue("#", i, i + 1)
            Next

            ' Set columns size
            oGrid.Columns.Item(0).Width = 30
            oGrid.Columns.Item(1).Width = 100
            oGrid.Columns.Item(2).Width = 100
            oGrid.Columns.Item(3).Width = 100
            oGrid.Columns.Item(4).Width = 100
            oGrid.Columns.Item(0).Editable = False

            Return 0

        Catch ex As Exception

            MsgBox("FrmtekDel. fallo la carga previa de la forma AgregarLineas: " & ex.Message)

        Finally

            oGrid = Nothing

        End Try

    End Function

End Class
