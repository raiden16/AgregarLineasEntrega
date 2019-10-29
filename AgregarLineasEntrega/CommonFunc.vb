Option Strict Off
Option Explicit On
Imports Microsoft.Win32
Imports System.Windows.Forms
Imports System.Windows.Forms.MessageBox

Module CommonFunc

    Public Function loadFormXML(ByVal SBOApplication As SAPbouiCOM.Application, ByVal psFormUID As String, ByVal psFile As String) As Integer

        Dim loXMLDoc As MSXML2.DOMDocument
        Dim loForm As SAPbouiCOM.Form
        Try
            loXMLDoc = New MSXML2.DOMDocument
            '//BUSCA LA FORMA
            loForm = searchForm(SBOApplication, psFormUID)
            If (loForm Is Nothing) Then
                '//CARGA EL XML
                If (Not loXMLDoc.load(psFile)) Then
                    SBOApplication.MessageBox("No pudo abrir el archivo " & psFile)
                    Return -1
                End If

                SBOApplication.LoadBatchActions(loXMLDoc.xml)
                loadFormXML = 0
            End If
            '//MUEVE EL FOCO A LA FORMA
            loForm = SBOApplication.Forms.Item(psFormUID)
            loForm.Select()
        Catch ex As Exception
            SBOApplication.MessageBox("CommonFunc. Error al abrir la forma " & psFormUID & ". " & ex.Message)
            Return -1
        End Try

    End Function

    '//----- BUSCA UNA FORMA INDICADA EN LA APLICACION
    Public Function searchForm(ByVal SBOApplication As SAPbouiCOM.Application,
                               ByVal psFormUID As String) As SAPbouiCOM.Form
        Try
            searchForm = SBOApplication.Forms.Item(psFormUID)
        Catch ex As Exception
            searchForm = Nothing
        End Try
    End Function

    '//----- LEE EL PATH INDICADO
    Public Function ReadPath(ByVal psApplName As String) As String
        Dim sAns As String
        Dim sErr As String = ""
        sAns = My.Application.Info.DirectoryPath
        'sAns = RegValue(RegistryHive.CurrentUser, "BBConsulting", psApplName, sErr)
        ReadPath = sAns
        If Not (sAns <> "") Then
            MessageBox.Show("CommonFunc. Al obtener el valor del registro. " & sErr)
            ReadPath = ""
        End If
    End Function

End Module