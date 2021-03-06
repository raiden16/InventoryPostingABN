Public Class CatchingEvents


    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF
    Dim DocNum As String
    Dim conexionSQL As Sap.Data.Hana.HanaConnection

    Public Sub New()

        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

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
            End '//termina aplicación
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
            End '//termina aplicación
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
            End '//termina aplicación
        Finally
            loRecSet = Nothing
        End Try
    End Sub


    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try
            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx(1474000001) '// FORMA Recuento de Inventario

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
    '// METODOS PARA MANEJO DE EVENTOS ITEM
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent

        If pVal.Before_Action = True And pVal.FormTypeEx <> "" Then

            Select Case pVal.FormTypeEx

                Case 1474000001                           '////// FORMA Recuento de Inventario
                    frmOVPMControllerBefore(FormUID, pVal)

            End Select

        Else
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then
                Select Case pVal.FormTypeEx

                    Case 1474000001                           '////// FORMA Recuento de Inventario
                        frmOVPMControllerAfter(FormUID, pVal)

                End Select
            End If
        End If

    End Sub


    Private Sub frmOVPMControllerBefore(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim coForm As SAPbouiCOM.Form
        Dim stTabla As String
        Dim oDatatable As SAPbouiCOM.DBDataSource

        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case 1470000001

                            stTabla = "OINC"
                            coForm = SBOApplication.Forms.Item(FormUID)

                            oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            DocNum = oDatatable.GetValue("DocNum", 0)

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Facturacion Clientes. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA Recuento de Inventario
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmOVPMControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)
        Dim DocEntry, ObjType, LineNum, ItemCode, WhsCode, BatchNumber, Lote As String
        Dim CountQty, InWhsQty, Difference As Double
        Dim stQueryH1, stQueryH2 As String
        Dim oRecSetH1, oRecSetH2 As SAPbobsCOM.Recordset
        Dim oCS As SAPbobsCOM.CompanyService
        Dim oIPS As SAPbobsCOM.InventoryPostingsService
        Dim oIP As SAPbobsCOM.InventoryPosting
        Dim oIPLS As SAPbobsCOM.InventoryPostingLines
        Dim oIPL As SAPbobsCOM.InventoryPostingLine
        Dim oIPBNS As SAPbobsCOM.InventoryPostingBatchNumbers
        Dim oIPBN As SAPbobsCOM.InventoryPostingBatchNumber
        Dim oIPP As SAPbobsCOM.InventoryPostingParams
        Dim CantidadR, CantidadL As Double
        Dim tabla As DataTable
        Dim comm As New Sap.Data.Hana.HanaCommand
        Dim DA As New Sap.Data.Hana.HanaDataAdapter
        Dim ds As New DataSet

        oRecSetH1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oCS = SBOCompany.GetCompanyService
        oIPS = oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryPostingsService)
        oIP = oIPS.GetDataInterface(SAPbobsCOM.InventoryPostingsServiceDataInterfaces.ipsInventoryPosting)
        oIPLS = oIP.InventoryPostingLines

        Try

            Select Case pVal.EventType

                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case 1470000001

                            conectar()

                            stQueryH1 = "Select T1.""DocEntry"",T0.""ObjType"",T1.""LineNum"",T1.""ItemCode"",T1.""WhsCode"",T1.""CountQty"",T1.""InWhsQty"",T1.""Difference"",T2.""ManBtchNum"" from """ & SBOCompany.CompanyDB & """.OINC T0 Inner Join """ & SBOCompany.CompanyDB & """.INC1 T1 on T1.""DocEntry""=T0.""DocEntry"" Inner Join """ & SBOCompany.CompanyDB & """.OITM T2 on T2.""ItemCode""=T1.""ItemCode"" where T0.""DocNum""=" & DocNum
                            oRecSetH1.DoQuery(stQueryH1)
                            comm.CommandText = stQueryH1
                            comm.Connection = conexionSQL
                            DA.SelectCommand = comm
                            DA.Fill(ds)

                            If oRecSetH1.RecordCount > 0 Then

                                tabla = ds.Tables(0)

                                oIP.CountDate = DateTime.Now
                                oRecSetH1.MoveFirst()

                                For Each i As DataRow In tabla.Rows
                                    'i = 0 To oRecSetH1.RecordCount - 1

                                    DocEntry = oRecSetH1.Fields.Item("DocEntry").Value
                                    ObjType = oRecSetH1.Fields.Item("ObjType").Value
                                    LineNum = oRecSetH1.Fields.Item("LineNum").Value
                                    ItemCode = oRecSetH1.Fields.Item("ItemCode").Value
                                    WhsCode = oRecSetH1.Fields.Item("WhsCode").Value
                                    CountQty = oRecSetH1.Fields.Item("CountQty").Value
                                    InWhsQty = oRecSetH1.Fields.Item("InWhsQty").Value
                                    Difference = oRecSetH1.Fields.Item("Difference").Value
                                    Lote = oRecSetH1.Fields.Item("ManBtchNum").Value

                                    oIPL = oIPLS.Add()
                                    oIPL.BaseEntry = DocEntry
                                    oIPL.BaseReference = DocNum
                                    oIPL.BaseType = ObjType
                                    oIPL.BaseLine = LineNum
                                    oIPL.ItemCode = ItemCode
                                    oIPL.WarehouseCode = WhsCode
                                    oIPL.CountedQuantity = CountQty
                                    'oIPL.InWarehouseQuantity = Difference

                                    If Difference > 0 And Lote = "Y" Then

                                        stQueryH2 = "Select Top 1 ""BatchNum"" from OIBT where ""ItemCode""='" & ItemCode & "' AND ""WhsCode""='" & WhsCode & "' AND ""Direction""=0 order by ""CreateDate"" desc"
                                        oRecSetH2.DoQuery(stQueryH2)

                                        If oRecSetH2.RecordCount > 0 Then

                                            oRecSetH2.MoveFirst()

                                            oIPBNS = oIPL.InventoryPostingBatchNumbers
                                            oIPBN = oIPBNS.Add()

                                            BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                            oIPBN.BatchNumber = BatchNumber
                                            oIPBN.Quantity = CountQty
                                            oIPBN.BaseLineNumber = LineNum

                                        Else

                                            stQueryH2 = "Select Top 1 ""BatchNum"" from OIBT where ""ItemCode""='" & ItemCode & "' AND ""WhsCode""='001' AND ""Direction""=0 order by ""CreateDate"" desc"
                                            oRecSetH2.DoQuery(stQueryH2)

                                            If oRecSetH2.RecordCount > 0 Then

                                                oRecSetH2.MoveFirst()

                                                oIPBNS = oIPL.InventoryPostingBatchNumbers
                                                oIPBN = oIPBNS.Add()

                                                BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                                oIPBN.BatchNumber = BatchNumber
                                                oIPBN.Quantity = CountQty
                                                oIPBN.BaseLineNumber = LineNum

                                            End If

                                        End If

                                    ElseIf Difference < 0 And Lote = "Y" Then

                                        stQueryH2 = "Select T0.*,T1.""CreateDate"" from
                                                    (Select ""BatchNum"",""ItemCode"",""WhsCode"",
                                                    sum(case when ""Direction""=0 then ""Quantity"" else -1*""Quantity"" end) as ""CantidadLote"" 
                                                    from IBT1 where ""ItemCode""='" & ItemCode & "' AND ""WhsCode""='" & WhsCode & "'
                                                    Group by  ""BatchNum"",""ItemCode"",""WhsCode"") T0
                                                    Inner Join OBTN T1 on T1.""DistNumber""=T0.""BatchNum"" and T1.""ItemCode""=T0.""ItemCode""
                                                    where T0.""CantidadLote"">0
                                                    order by T1.""CreateDate"""
                                        oRecSetH2.DoQuery(stQueryH2)

                                        If oRecSetH2.RecordCount > 0 Then

                                            oRecSetH2.MoveFirst()
                                            oIPBNS = oIPL.InventoryPostingBatchNumbers
                                            CantidadR = Format(CountQty, "0.000")

                                            For l = 0 To oRecSetH2.RecordCount - 1

                                                CantidadL = Format(oRecSetH2.Fields.Item("CantidadLote").Value, "0.000")

                                                If CantidadR > CantidadL Then

                                                    CantidadR = Format(CantidadR - CantidadL, "0.000")

                                                    oIPBN = oIPBNS.Add()

                                                    BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                                    oIPBN.BatchNumber = BatchNumber
                                                    oIPBN.Quantity = CantidadL
                                                    oIPBN.BaseLineNumber = LineNum

                                                    l = 0

                                                Else

                                                    oIPBN = oIPBNS.Add()

                                                    BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value

                                                    oIPBN.BatchNumber = BatchNumber
                                                    oIPBN.Quantity = CantidadR
                                                    oIPBN.BaseLineNumber = LineNum

                                                    l = oRecSetH2.RecordCount - 1

                                                End If

                                                oRecSetH2.MoveNext()

                                            Next

                                        End If

                                    End If

                                    oRecSetH1.MoveNext()

                                Next

                                oIPP = oIPS.Add(oIP)
                                SBOApplication.MessageBox("Se creo con exito la contabilización de stocks.")

                                conexionSQL.Close()

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Recuento de Inventario. " & ex.Message)
            conexionSQL.Close()
        Finally
            'oPO = Nothing
        End Try
    End Sub


    Public Function conectar() As Boolean
        Dim stCadenaConexion As String
        Try

            conectar = False

            ''---- Cargamos datos de archivo de configuracion

            '---- objeto compañia
            conexionSQL = New Sap.Data.Hana.HanaConnection

            '---- armamos cadena de conexion
            stCadenaConexion = "DRIVER={B1CRHPROXY32};UID=" & My.Settings.UserSQL & ";PWD=" & My.Settings.PassSQL & ";SERVERNODE=" & My.Settings.Server

            '---- realizamos conexion
            conexionSQL = New Sap.Data.Hana.HanaConnection(stCadenaConexion)

            conexionSQL.Open()

        Catch ex As Exception
            SBOApplication.MessageBox("Error al conectar con HANA . " & ex.Message)
        End Try
    End Function


End Class
