VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strQuery  As String
Private m_Reporte As String
Private m_Query As String
Private m_Titulo As String
Private m_Empresa As String
Private m_Igv As String
Private m_Nombre1 As String
Private m_Cargo1 As String
Private m_Nombre2 As String
Private m_Cargo2 As String
Private m_Anomes As String
Private m_MontLetras As String
Private m_Glosa As String
Private m_TipoCambio As String
Private m_tipo As String
Private m_tipoDoc As String
Private m_Formulario As String
Private m_Estado As String
Private m_Fecha As String
Private m_Division As String
Private m_Titulos(30) As String
Private m_Tc(12) As Double
Private m_Cab1 As String
Private m_Cab2 As String
Private m_Cab3 As String
Private m_Cab4 As String
Private m_Cab5 As String
Private m_Cab6 As String
Private m_Cab7 As String
Private m_Cab8 As String
Private m_Cab9 As String
Private m_Cab10 As String
Private m_Cab11 As String
Private m_Cab12 As String
Private m_Cab13 As String
Private m_Cab14 As String
Private m_Cab15 As String
Private m_Sede As String
Private Param_Division  As String
Private m_Caso As String
Public Rs As ADODB.Recordset
Dim Cant As Integer
Dim RQ_FACT As MYSQL_RS
Dim RutaArchivo As String
Dim Rep_ECxA_monedaequiv As Double, Rep_ECxA_saldoanterior As Double
Dim Rep_ECxA_difanterior As Double
Dim RV_Descrip  As String, RV_Ruc  As String, RV_Fecha As String

Public Property Let T_Caso(valor As String)
  m_Caso = valor
End Property
Public Property Let TCambio(Indice As Integer, valor As Double)
  m_Tc(Indice) = valor
End Property
Public Property Let Division(valor As Integer)
  m_Division = valor
End Property
Public Property Let Estado(valor As String)
  m_Estado = valor
End Property
Public Property Let TituloLin(Indice As Integer, valor As String)
  m_Titulos(Indice) = valor
End Property
Public Property Let Formulario(valor As String)
  m_Formulario = valor
End Property
Public Property Let TipoBD(valor As String)
  m_tipo = valor
End Property
Public Property Let TipoDoc(valor As String)
  m_tipoDoc = valor
End Property
Public Property Let Titulo(valor As String)
  m_Titulo = valor
End Property
Public Property Let empresa(valor As String)
  m_Empresa = valor
End Property
Public Property Let fecha(valor As String)
  m_Fecha = UCase(valor)
End Property
Public Property Let MontLetras(valor As String)
  m_MontLetras = valor
End Property
Public Property Let Glosa(valor As String)
  m_Glosa = valor
End Property
Public Property Let TipoCambio(valor As String)
  m_TipoCambio = valor
End Property
Public Property Let Igv(valor As String)
  m_Igv = valor
End Property
Public Property Let Nombre1(valor As String)
  m_Nombre1 = valor
End Property
Public Property Let Nombre2(valor As String)
  m_Nombre2 = valor
End Property
Public Property Let Cargo1(valor As String)
  m_Cargo1 = valor
End Property
Public Property Let Cargo2(valor As String)
  m_Cargo2 = valor
End Property
Public Property Let Reporte(valor As String)
  m_Reporte = UCase(valor)
End Property
Public Property Let Query(valor As String)
  m_Query = valor
End Property
Public Property Let Cab1(valor As String)
  m_Cab1 = valor
End Property
Public Property Let Cab2(valor As String)
  m_Cab2 = valor
End Property
Public Property Let Cab3(valor As String)
  m_Cab3 = valor
End Property
Public Property Let Cab4(valor As String)
  m_Cab4 = valor
End Property
Public Property Let Cab5(valor As String)
  m_Cab5 = valor
End Property
Public Property Let Cab6(valor As String)
  m_Cab6 = valor
End Property
Public Property Let Cab7(valor As String)
  m_Cab7 = valor
End Property
Public Property Let Cab8(valor As String)
  m_Cab8 = valor
End Property
Public Property Let Cab9(valor As String)
  m_Cab9 = valor
End Property
Public Property Let Cab10(valor As String)
  m_Cab10 = valor
End Property
Public Property Let Cab11(valor As String)
  m_Cab11 = valor
End Property
Public Property Let Cab12(valor As String)
  m_Cab12 = valor
End Property
Public Property Let Cab13(valor As String)
  m_Cab13 = valor
End Property
Public Property Let Cab14(valor As String)
  m_Cab14 = valor
End Property
Public Property Let Cab15(valor As String)
  m_Cab15 = valor
End Property
Public Property Let sede(valor As String)
  m_Sede = valor
End Property
Public Function Convertir(myrs As MYSQL_RS) As ADODB.Recordset
    On Error GoTo SERROR
    Set Rs = New ADODB.Recordset
    Dim MyCols As Integer
    Dim cNombre As String
    Dim nDefinedSize As Integer
    Dim Tipo As DataTypeEnum
    myrs.MoveFirst
    For MyCols = 0 To myrs.FieldCount - 1
        Select Case myrs.Fields(MyCols).MySqlType
            Case MYSQL_SMALLINT: Tipo = adSmallInt
            Case MYSQL_INT: Tipo = adInteger
            Case MYSQL_DOUBLE: Tipo = adDouble
            Case MYSQL_VARCHAR: Tipo = adVarChar
            Case MYSQL_CHAR: Tipo = adChar
            Case MYSQL_TIME: Tipo = adDBTimeStamp
        End Select
        
        cNombre = myrs.Fields(MyCols).name
        nDefinedSize = myrs.Fields(MyCols).DefinedSize
        '---------------------------------------------------------------------------------------------
        ' para el caso de los campos agregados por query manualmente de tipo texto con valor en blanco
        If InStr(1, cNombre, " ") > 0 Then
            cNombre = Replace(cNombre, " ", "_")
            nDefinedSize = 50
            Tipo = adVarChar
        End If
        '---------------------------------------------------------------------------------------------
        Rs.Fields.Append cNombre, Tipo, nDefinedSize, adFldUpdatable
    Next
    Rs.Open
    Do While Not (myrs.EOF)
        Rs.AddNew
        For MyCols = 0 To myrs.FieldCount - 1
            Rs.Fields(MyCols).Value = CE(myrs.Fields(MyCols))
        Next
        Rs.Update
        myrs.MoveNext
    Loop
    Set Convertir = Rs
    Set Rs = Nothing

    Exit Function
SERROR:
    Mensajes err.Description
    Set Convertir = Rs
    Set Rs = Nothing

End Function


Public Sub ImprimeReporteGeneral(rsResultado As MYSQL_RS, rsResultado2 As MYSQL_RS, strNombreDelReporte As String, strTitulo As String, strSubTitulo As String, strEmpresa As String, Optional rsResultado3 As MYSQL_RS, Optional rsResultado4 As MYSQL_RS)
On Error GoTo controlError
    Dim crpAplication As New CRAXDRT.Application
    Dim crpReport As CRAXDRT.Report
    Dim CrDatabase As CRAXDRT.database
    Dim CrDatabaseTables As CRAXDRT.DatabaseTables
    Dim CrDatabaseTable As CRAXDRT.DatabaseTable
    Dim CrSections As CRAXDRT.Sections
    Dim CrSection As CRAXDRT.Section
    Dim CrReportObjs As CRAXDRT.ReportObjects
    Dim CrSubreportObj As CRAXDRT.SubreportObject
    Dim CrSubreport As CRAXDRT.Report
    Dim crpDatabase As CRAXDRT.database
    Dim crpParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crpParamDef As CRAXDRT.ParameterFieldDefinition
    Dim strRutaDelReporte As String
    If UCase(Right(strNombreDelReporte, 4)) = ".RPT" Then strNombreDelReporte = Mid(strNombreDelReporte, 1, Len(strNombreDelReporte) - 4)
    strRutaDelReporte = strNombreDelReporte & ".RPT"
    Set crpReport = crpAplication.OpenReport(strRutaDelReporte)
    crpReport.DiscardSavedData
    Set crpParamDefs = crpReport.ParameterFields
    For Each crpParamDef In crpParamDefs
        With crpParamDef
            Select Case .ParameterFieldName
                Case "TITULO"
                    .SetCurrentValue strTitulo
                Case "SUBTITULO"
                    .SetCurrentValue strSubTitulo
                Case "EMPRESA"
                    .SetCurrentValue strEmpresa
                Case "REPORTE"
                    .SetCurrentValue strNombreDelReporte
            End Select
        End With
    Next
    Set crpDatabase = crpReport.database
    Call crpDatabase.Tables(1).SetPrivateData(3, Convertir(rsResultado))
    Set CrSections = crpReport.Sections
    Dim X, Y
    For X = 1 To CrSections.Count
        Set CrSection = CrSections.Item(X)
        Set CrReportObjs = CrSection.ReportObjects
        For Y = 1 To CrReportObjs.Count
            If CrReportObjs.Item(Y).Kind = crSubreportObject Then
                Set CrSubreportObj = CrReportObjs.Item(Y)
                Set CrSubreport = CrSubreportObj.OpenSubreport
                Set CrDatabase = CrSubreport.database
                Set CrDatabaseTables = CrDatabase.Tables
                Set CrDatabaseTable = CrDatabaseTables.Item(1)
                Select Case CrSubreportObj.name
                    Case "Subreport1":
                        CrDatabaseTable.SetDataSource Convertir(rsResultado2), 3
                    Case "Subreport2":
                        CrDatabaseTable.SetDataSource Convertir(rsResultado3), 3
                    Case "Subreport3":
                        CrDatabaseTable.SetDataSource Convertir(rsResultado4), 3
                End Select
            End If
        Next
    Next
    Dim oPreviewReport As New frmReportPreview
    oPreviewReport.Caption = strNombreDelReporte
    oPreviewReport.SetReporte crpReport
    Set oPreviewReport = Nothing
    Set crpAplication = Nothing
    Set crpReport = Nothing
    Set crpDatabase = Nothing
    Set crpParamDefs = Nothing
    Set crpParamDef = Nothing
    Set Rs = Nothing
Exit Sub
controlError:
    Set oPreviewReport = Nothing
    Set crpAplication = Nothing
    Set crpReport = Nothing
    Set crpDatabase = Nothing
    Set crpParamDefs = Nothing
    Set crpParamDef = Nothing
    Set Rs = Nothing
End Sub
Public Sub EjecutaReporte(strRutaDelReporte As String, rsResultado As MYSQL_RS, Optional oRs As ADODB.Recordset, Optional bExportar As Boolean = False, Optional bSoloExportar As Boolean = False, Optional cNombreExportar As String)
    On Error GoTo controlError
    Screen.MousePointer = vbHourglass
    DoEvents
    Dim crpAplication As New CRAXDRT.Application
    Dim crpReport As New CRAXDRT.Report
    Dim CrDatabase As CRAXDRT.database
    Dim CrDatabaseTables As CRAXDRT.DatabaseTables
    Dim CrDatabaseTable As CRAXDRT.DatabaseTable
    Dim CrSections As CRAXDRT.Sections
    Dim CrSection As CRAXDRT.Section
    Dim CrReportObjs As CRAXDRT.ReportObjects
    Dim CrSubreportObj As CRAXDRT.SubreportObject
    Dim CrSubreport As CRAXDRT.Report
    Dim crpDatabase As CRAXDRT.database
    Dim crpParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crpParamDef As CRAXDRT.ParameterFieldDefinition
    Dim SQL As Integer
    Set crpReport = crpAplication.OpenReport(strRutaDelReporte)
    crpReport.DiscardSavedData
    Set crpParamDefs = crpReport.ParameterFields
    For Each crpParamDef In crpParamDefs
        With crpParamDef
            Select Case .ParameterFieldName
                Case "ANOMES"
                    .SetCurrentValue m_Anomes
                Case "TITULO"
                    .SetCurrentValue m_Titulo
                Case "EMPRESA"
                    .SetCurrentValue m_Empresa
                Case "DIVISION"
                    .SetCurrentValue m_Division
                Case "Fecha"
                    .SetCurrentValue m_Fecha
                Case "REPORTE"
                    .SetCurrentValue m_Reporte
                Case "Igv"
                    .SetCurrentValue m_Igv
                Case "Nombre1"
                    .SetCurrentValue m_Nombre1
                Case "Nombre2"
                    .SetCurrentValue m_Nombre2
                Case "Cargo1"
                    .SetCurrentValue m_Cargo1
                Case "Cargo2"
                    .SetCurrentValue m_Cargo2
                Case "MontLetras"
                    .SetCurrentValue m_MontLetras
                Case "TipoCambio"
                    .SetCurrentValue m_TipoCambio
                Case "TipoDoc"
                    .SetCurrentValue m_tipoDoc
                Case "Estado"
                    .SetCurrentValue m_Estado
                Case "Glosa"
                    .SetCurrentValue m_Glosa
                Case "Cab1"
                    .SetCurrentValue m_Cab1
                Case "Cab2"
                    .SetCurrentValue m_Cab2
                Case "Cab3"
                    .SetCurrentValue m_Cab3
                Case "Cab4"
                    .SetCurrentValue m_Cab4
                Case "Cab5"
                    .SetCurrentValue m_Cab5
                Case "Cab6"
                    .SetCurrentValue m_Cab6
                Case "Cab7"
                    .SetCurrentValue m_Cab7
                Case "Cab8"
                    .SetCurrentValue m_Cab8
                Case "Cab9"
                    .SetCurrentValue m_Cab9
                Case "Cab10"
                    .SetCurrentValue m_Cab10
                Case "Cab11"
                    .SetCurrentValue m_Cab11
                Case "Cab12"
                    .SetCurrentValue m_Cab12
                Case "Cab13"
                    .SetCurrentValue m_Cab13
                Case "Cab14"
                    .SetCurrentValue m_Cab14
                Case "Cab15"
                    .SetCurrentValue m_Cab15
                Case "Sede"
                    .SetCurrentValue m_Sede
            End Select
        End With
    Next
    Set crpDatabase = crpReport.database
    
    If oRs Is Nothing Then
        Call crpDatabase.Tables(1).SetPrivateData(3, Convertir(rsResultado))
    Else
        Call crpDatabase.Tables(1).SetPrivateData(3, oRs)
    End If
    
    Set CrSections = crpReport.Sections
    Dim X, Y
    For X = 1 To CrSections.Count
        Set CrSection = CrSections.Item(X)
        Set CrReportObjs = CrSection.ReportObjects
        For Y = 1 To CrReportObjs.Count
            If CrReportObjs.Item(Y).Kind = crSubreportObject Then
                Set CrSubreportObj = CrReportObjs.Item(Y)
                Set CrSubreport = CrSubreportObj.OpenSubreport
                Set CrDatabase = CrSubreport.database
                Set CrDatabaseTables = crpDatabase.Tables
                Set CrDatabaseTable = CrDatabaseTables.Item(1)
                
                If oRs Is Nothing Then
                    CrDatabaseTable.SetDataSource Convertir(rsResultado), 3
                Else
                    CrDatabaseTable.SetDataSource oRs, 3
                End If
            End If
        Next
    Next
    Dim oPreviewReport As frmReportPreview
    Set oPreviewReport = New frmReportPreview
    oPreviewReport.Caption = Mid(m_Reporte, 5, Len(m_Reporte) - 4) & " - " & NombreMes(strMesSistema, False)
    oPreviewReport.SetReporte crpReport
    
    Dim cResult As VbMsgBoxResult
    cResult = vbNo
    
    
    If bSoloExportar = False Then
        If bExportar = False Then
            cResult = vbNo
        Else
            cResult = MsgBox("Desea generar el reporte en formato PDF", vbYesNo + vbQuestion, "")
        End If
    Else
        cResult = vbYes
    End If
    
    If cResult = vbYes Then
        On Error Resume Next
        Dim cFile As String
        
        If bSoloExportar = True Then
            cFile = cNombreExportar
        Else
            cFile = Rep_Documents & "\" & m_Reporte & ".pdf"
        End If
        
        crpReport.ExportOptions.FormatType = crEFTPortableDocFormat
        crpReport.ExportOptions.DestinationType = crEDTDiskFile
        crpReport.ExportOptions.DiskFileName = cFile
        crpReport.Export False
        
        If bSoloExportar = False Then
            ShellExecute mdiInicio.hwnd, "open", cFile, vbNullString, vbNullString, 0
        End If
    End If
    
    Screen.MousePointer = vbNormal
    
    Set oPreviewReport = Nothing
    Set crpAplication = Nothing
    Set crpReport = Nothing
    Set crpDatabase = Nothing
    Set crpParamDefs = Nothing
    Set crpParamDef = Nothing
Exit Sub
controlError:
    Set oPreviewReport = Nothing
    Set crpAplication = Nothing
    Set crpReport = Nothing
    Set crpDatabase = Nothing
    Set crpParamDefs = Nothing
    Set crpParamDef = Nothing
    MsgBox "Ocurrió un error al tratar de Mostrar el Reporte " & vbCrLf & vbCrLf & _
            "Fuente: " & err.Source & vbCrLf & _
            "Número: " & err.Number & vbCrLf & _
            "Descripción: " & err.Description, vbCritical, "ERROR"
End Sub
Public Sub sp_Rep_CentroCostos(Auxiliar As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select trim(a.cenco) as cenco,a.descrip,b.descrip as encargado" & _
          " from cncosto as a Left Join cnauxil as b " & _
          " on (a.cod_auxil=b.codigo) where b.auxiliar='" & Auxiliar & "'" & _
          " order by a.cenco"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Usuarios()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select  a.usuario_id, a.nombre, b.descripcion as Perfil," & _
          "c.descrip as Area, d.anomes_bloq as Año_Mes_Bloqueo " & _
          "from 3cnuser as a LefT Join 1cnperfil as b on (a.perfil_id=b.perfil_id) " & _
          "Join cnarea as c on (a.area=c.idarea)  Join 7cia_user as d " & _
          "on(a.usuario_id=d.usuario_id)"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        Call MsgBox("No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento")
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_TiposDoc()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select coddoc, descrip, abrevia from cndocum T " & _
          "where (protegido = 'N' OR (SELECT permiso FROM docsusuario D WHERE D.coddoc=t.coddoc AND usuario = '" & strUsuarioId & "')=1)"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        Call MsgBox("No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento")
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_AccesoMenu(ByVal Nombre As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select  a.usuario_id, a.modulo, b.descrip, u.nombre " & _
          "from (5usuario_menu as a Left Join menu as b " & _
          "on (a.modulo=b.modulo and a.item=b.item) ) left Join  usuarios as u " & _
          "on (a.usuario_id=u.usuario_id) Where u.nombre='" & Nombre & "' AND a.visible='1'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        Call MsgBox("No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento")
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Auxiliares(Auxil As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select  DISTINCT a.Auxiliar , B.DESCRIP , a.codigo, a.DESCRIP as Aux," & _
          "a.DIRECC, a.DISTRI, a.Tipo from cnauxil as a RIGHT Join cntablas as b" & _
          " on (a.auxiliar=b.Tip_linea )WHERE a.auxiliar='" & Auxil & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        Call MsgBox("No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento")
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_Rep_ControlBC(ByVal AnioMes As String, ByVal DiaIni As String, ByVal DiaFin As String, ByVal Division As String)

    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
     
    Set oRs = ADO_LlenaRs("call RH_SP_BonosCampo('REPORTE','" & AnioMes & "','" & DiaIni & "','" & DiaFin & "','" & strUsuarioId & "','" & Division & "','');")

    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No hay datos de control de bonos", vbOKOnly + vbInformation, gsNomSW
            Exit Sub
        End If

        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, True
    End If

    Set oRs = Nothing

End Sub

Public Sub sp_Rep_ReporteAR(ByVal fecha As String, ByVal Division As String, ByVal Cliente As String, ByVal cDocumento As String, ByVal US As Double, ByVal OA As Double, ByVal cConv As String)

    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
     
    Set oRs = ADO_LlenaRs("call CC_SP_CuentasPorCobrar('REPORTE','" & Division & "','" & Cliente & "','" & fecha & "','" & strUsuarioId & "','" & cDocumento & "','" & US & "','" & OA & "','" & cConv & "','');")

    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No hay registros para el Reporte del AR", vbOKOnly + vbInformation, gsNomSW
            Exit Sub
        End If

        EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, oRSMYSQL, oRs, False
    End If

    Set oRs = Nothing

End Sub

Public Sub sp_Rep_OrdenCompra(CorreIni As String, CorreFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = " Select o.identificador,o.documento as N_Orden,o.Fec_Emision as Fecha, o.mon, " & _
          " o.total as Importe,o.auxiliar,o.desaux,o.Codigo,o.Descod as Nombre,o.Tipo,trim(o.cenco) as cenco, " & _
          " oc.tipo_cotizacion , oc.requisicion, (Select sum(a.total) + sum(a.otros_montos) from documento_contables as a left join" & _
          " movi_documento as b ON(a.identificador=b.identificador)" & _
          " where a.orden=oc.correl and  b.cod_estado<>'EL' and b.cod_estado<>'AN') as" & _
          " monto_factura, oc.otros_montos, o.division, a.cod_estado " & _
          " From (temprptdocumento as o " & _
          " left JOIN orden_compra as oc on (oc.Correl=o.documento)) " & _
          " left JOIN movi_documento as a on (a.identificador=o.identificador) where o.usu='" & strUsuarioId & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Varios()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select * from temprptdocumento where usu='" & strUsuarioId & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_VariosxCC(Div As String, cenco As String, Doc As String, aux As String, Cod As String, mon As String, Estado As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Delete from temprptdocumento where usu='" & strUsuarioId & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    Dim I As Integer
    For I = 1 To 4
        Select Case I
            Case 1
                SQL = " Insert into temprptdocumento select a.Identificador,1 as familia,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo, " & _
                      "concat(doc.serie,'-',doc.correl) AS Documento,doc.orden,doc.guia, doc.auxiliar AS auxiliar, " & _
                      "t.descrip as desaux,IF(k.coddoc<>'',doc.codigo,'') as Codigo,if(k.coddoc<>'',aux.descrip,'') as descod, " & _
                      "trim(doc.cenco) AS cenco,c.descrip as descenco,doc.mon AS mon,doc.dias_vcto,doc.fec_vcto, doc.fec_emision, " & _
                      "a.fecha_registro as fec_registro,if(k.coddoc<>'',doc.total,0) as total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                      " (SELECT U.AREA FROM 3cnuser AS U WHERE USUARIO_ID='" & strUsuarioId & "') AS area, m.prioridad AS prioridad, doc.division,m.usuario,'" & strUsuarioId & "' as usu,doc.obs" & _
                      " from (((((amarre_documento as a, movi_documento as m ,documento_contables AS doc)" & _
                      " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                      " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                      " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                      " on (doc.auxiliar=t.tip_linea ) LEFT join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N' OR " & _
                      "(SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K on (a.Cod_Tipo_Doc=k.coddoc)"
            Case 2
                SQL = " Insert into temprptdocumento select a.Identificador,4 as familia,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                      " concat(doc.serie,'-',doc.correl) AS Documento,'' as orden,'' as guia," & _
                      " doc.auxiliar AS auxiliar, t.descrip as desaux,if(k.coddoc<>'',doc.codigo,'') as Codigo,if(k.coddoc<>'',aux.descrip,'') as descod," & _
                      " trim(doc.cenco) AS cenco,c.descrip as descenco,'' as mon,'' as dias_vcto,'' as fec_vcto," & _
                      " doc.fec_emision,a.fecha_registro as fec_registro,0 as total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                      " (SELECT U.AREA FROM 3cnuser AS U WHERE USUARIO_ID='" & strUsuarioId & "') AS area, m.prioridad AS prioridad, doc.division,m.usuario,'" & strUsuarioId & "' as usu,doc.obs" & _
                      " from (((((amarre_documento as a, movi_documento as m ,documento_entidades AS doc)" & _
                      " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                      " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                      " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                      " on (doc.auxiliar=t.tip_linea ) LEFT join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N' OR " & _
                      "(SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K on (a.Cod_Tipo_Doc=k.coddoc)"
            Case 3
                SQL = " Insert into temprptdocumento select a.Identificador,6 as familia,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                      " doc.correl AS Documento,'' as orden,doc.guia," & _
                      " doc.auxiliar AS auxiliar, t.descrip as desaux,if(k.coddoc<>'',doc.codigo,'') as Codigo,if(k.coddoc<>'',aux.descrip,'') as descod," & _
                      " trim(doc.cenco) AS cenco,c.descrip as descenco,doc.mon,'' as dias_vcto,doc.fec_vcto," & _
                      " doc.fec_emision,a.fecha_registro as fec_registro,if(k.coddoc<>'',doc.total,0) as total,'' AS estado,m.Cod_Estado as desestado," & _
                      " (SELECT U.AREA FROM 3cnuser AS U WHERE USUARIO_ID='" & strUsuarioId & "') AS area, m.prioridad AS prioridad, doc.division,m.usuario,'" & strUsuarioId & "' as usu,doc.obs" & _
                      " from (((((amarre_documento as a, movi_documento as m ,orden_compra AS doc)" & _
                      " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                      " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                      " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                      " on (doc.auxiliar=t.tip_linea ) LEFT join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N' OR " & _
                      "(SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K on (a.Cod_Tipo_Doc=k.coddoc)"
            Case 4
                SQL = " Insert into temprptdocumento select a.Identificador,5 as familia,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                      " concat(doc.serie,'-',doc.correl) AS Documento,'' as orden,'' as guia," & _
                      " doc.auxiliar AS auxiliar, t.descrip as desaux,if(k.coddoc<>'',doc.codigo,'') as Codigo,if(k.coddoc<>'',aux.descrip,'') as descod," & _
                      " trim(doc.cenco) AS cenco,c.descrip as descenco,doc.mon,'' as dias_vcto,'' as fec_vcto," & _
                      " doc.fec_emision,a.fecha_registro as fec_registro,if(k.coddoc<>'',doc.total,0) as total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                      " (SELECT U.AREA FROM 3cnuser AS U WHERE USUARIO_ID='" & strUsuarioId & "') AS area, m.prioridad AS prioridad, doc.division,m.usuario,'" & strUsuarioId & "' as usu,doc.obs" & _
                      " from (((((amarre_documento as a, movi_documento as m ,proformas AS doc)" & _
                      " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                      " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                      " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                      " on (doc.auxiliar=t.tip_linea ) LEFT join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N' OR " & _
                      "(SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K on (a.Cod_Tipo_Doc=k.coddoc)"
        End Select
        SQL = SQL & " where  T.codtab = 1 and (a.identificador=m.identificador) and  (a.identificador=doc.identificador) "
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    Next
    SQL = "Select * from temprptdocumento where left(identificador,6)='" & strAnoSistema + strMesSistema & "' and usu='" & strUsuarioId & "'"
    If Div <> "" Then SQL = SQL & " and division='" & Div & "'"
    If cenco <> "" Then SQL = SQL & " and cenco='" & cenco & "'"
    If Doc <> "" Then SQL = SQL & " and tipo='" & Doc & "'"
    If aux <> "" Then SQL = SQL & " and auxiliar='" & aux & "'"
    If Cod <> "" Then SQL = SQL & " and codigo='" & Cod & "'"
    If mon <> "" Then SQL = SQL & " and mon='" & mon & "'"
    If Estado <> "" Then SQL = SQL & " and estado='" & Estado & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Doc(Tipo As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    If UCase(Tipo) = "O" Then
        SQL = " select t.identificador,t.tipo,t.destipo,t.documento,t.guia,o.solicitado,o.requisicion," & _
              "  o.encargado, t.auxiliar,t.desaux,t.codigo,t.descod,trim(t.cenco) as cenco,t.descenco,o.tipo_cotizacion,t.mon," & _
              "  t.fec_emision,t.fec_registro,t.total,t.estado,t.desestado,t.area,t.prioridad,t.division,t.usuario" & _
              "  from temprptdocumento as  t  left join orden_compra as o on (t.documento=o.correl) where t.tipo='" & Tipo & "' and usu='" & strUsuarioId & "'"
    Else
        SQL = "Select * from temprptdocumento where tipo='" & Tipo & "' and usu='" & strUsuarioId & "'"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_ModificarDoc(tabla As String, Tipo As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select DISTINCT a.Identificador,a.cod_tipo_doc,a.Cod_Fam," & _
          "CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar,cn.descrip as TipoAux," & _
          "dc.codigo,aux.descrip as NomAux,trim(dc.Cenco) as cenco,ct.descrip as NomCenco," & _
          "dc.encargado,enc.NomEncargado as encargado,dc.division,dc.Obs AS Observacion," & _
          "ae.Id_Area,ae.DESCRIP as NomArea,ae.Usuario,ae.fecha_movi,ae.Cod_Estado , ae.Descripcion," & _
          "ae.Obs, ae.anomes from (((((amarre_documento as a LEFT Join " & tabla & " as dc " & _
          "on (a.Identificador=dc.Identificador)) LEFT Join cncosto as ct " & _
          "on (dc.cenco=ct.cenco))LEFT Join encargados as enc " & _
          "on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux " & _
          "on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))" & _
          "LEFT join cntablas as cn on (dc.auxiliar=cn.tip_linea and cn.codtab='1')) " & _
          "LEFT join asignar_estados as ae on (a.Identificador=ae.identificador) " & _
          "where a.identificador='" & Tipo & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_DocumentoCont(indt As String, tabla As String, familia As FAMILIA_DOC)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select a.Identificador,a.Fecha_Registro, a.hora_Registro,a.Empresa," & _
          "a.Cod_Fam,a.cod_tipo_doc,a.Nombre_Mensajero,a.Tipo_Doc_Ide," & _
          "did.descrip,a.Ide_Mensajero,a.Flag,CONCAT(dc.Serie,'-',dc.Correl) " & _
          "as Documento,dc.Auxiliar,cn.descrip as TipoAux,dc.codigo," & _
          "aux.descrip as NomAux,trim(dc.Cenco) as cenco,ct.descrip as NomCenco,dc.encargado," & _
          "enc.NomEncargado,dc.Fec_Emision,dc.division, dc.Obs,"
    Select Case familia
        Case FAMILIA_DOC.CONTABLES
            SQL = SQL & " dc.Orden, dc.Guia,dc.mon,dc.Fec_Vcto,dc.dias_Vcto, dc.Fec_Pago," & _
                 "dc.Fec_Contabilizada ,dc.SubTotal,dc.IGV, dc.Otros_Montos,dc.Total," & _
                 "dc.Voucher as Voucher,' 'as Asunto,' 'as Requisicion,' 'as Entidad," & _
                 "' ' as Notificador,' ' as Notificado,' ' as DirigidoA,' ' as Atencion," & _
                 "' ' as Remitente,' ' as FechaServicio, ' ' as FechaEntrega,' ' as Factura," & _
                 "' ' as MontoFac, ' ' as MontoGuia, ' ' as TipoCotizacion"
        Case FAMILIA_DOC.ORDENES
            SQL = SQL & "' ' as Orden,dc.guia,dc.mon,dc.Fec_Vcto,' 'as dias_Vcto,' ' as Fec_Pago," & _
                 "' ' as Fec_Contabilizada,' ' as SubTotal,' ' as IGV,' ' as Otros_Montos," & _
                 "dc.total,' ' as Voucher,' ' as Asunto,dc.requisicion ,' ' as Entidad," & _
                 "' ' as Notificador,' ' as Notificado,' ' as DirigidoA,' ' as Atencion," & _
                 "' ' as Remitente,dc.Fec_Inicio_servicio ,dc.Fec_entrega ," & _
                 "dc.factura ,(Select sum(a.total) + sum(a.otros_montos) from documento_contables as a left join" & _
                 " movi_documento as b ON(a.identificador=b.identificador)" & _
                 " where a.orden=dc.correl and  b.cod_estado<>'EL' and b.cod_estado<>'AN') as" & _
                 " monto_factura ,dc.monto_guia ,dc.Tipo_cotizacion "
        Case FAMILIA_DOC.GENERALES
            SQL = SQL & " dc.Orden,' ' as Guia,dc.mon,' ' as Fec_Vcto,' ' as dias_Vcto,' ' as Fec_Pago," & _
                  "' ' as Fec_Contabilizada ,dc.SubTotal,dc.IGV,' ' as Otros_Montos,dc.Total,' ' as Voucher," & _
                  "dc.Asunto,dc.requisicion,' 'as Entidad,' ' as Notificador," & _
                  "' ' as Notificado,dc.DirigidoA,dc.Atencion,' ' as Remitente," & _
                  "' ' as FechaServicio,' ' as FechaEntrega,' ' as Factura,' ' as MontoFac," & _
                  "' ' as MontoGuia,' ' as TipoCotizacion"
        Case FAMILIA_DOC.ENTIDADES
            SQL = SQL & "' ' as Orden,' ' as Guia,' ' as  mon,' ' as Fec_Vcto,' ' as dias_Vcto,' ' as Fec_Pago," & _
                  "' ' as Fec_Contabilizada ,' ' as SubTotal,' ' as IGV,' ' as Otros_Montos,' ' as Total," & _
                  "' ' as Voucher,dc.Asunto,' ' as Requisicion,dc.Entidad,dc.Notificador,dc.Notificado," & _
                  "' ' as DirigidoA,' ' as Atencion,' ' as Remitente,' ' as FechaServicio, ' ' as FechaEntrega," & _
                  "' ' as Factura,' ' as MontoFac,' ' as MontoGuia, ' ' as TipoCotizacion"
    End Select
    SQL = SQL & " from (((((amarre_documento as a LEFT Join " & tabla & " as dc " & _
          "on (a.Identificador=dc.Identificador)) LEFT Join doc_identificacion as did " & _
          "on (a.Tipo_Doc_Ide=did.Tipo_Doc_Ide)) LEFT Join cncosto as ct " & _
          "on (dc.cenco=ct.cenco)) LEFT Join encargados as enc " & _
          "on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux " & _
          "on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))" & _
          "LEFT join cntablas as cn on (dc.auxiliar=cn.tip_linea and cn.codtab='1') " & _
          "where a.identificador='" & indt & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Busqueda(fam As String, Auxiliar As String, Codauxil As String, cenco As String, Division As String, fecha As String, Tipo As String, Estado As String, anomesini As String, anomesfin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim tabla As String
    Select Case fam
        Case FAMILIA_DOC.CONTABLES: tabla = "documento_contables"
            SQL = " select a.Identificador,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                  " concat(doc.serie,'-',doc.correl) AS Documento,doc.orden,doc.guia," & _
                  " doc.auxiliar AS auxiliar, t.descrip as desaux,doc.codigo as Codigo,aux.descrip as descod," & _
                  " trim(doc.cenco) AS cenco,c.descrip as descenco,doc.mon AS moneda,doc.dias_vcto,doc.fec_vcto," & _
                  " doc.fec_emision,a.fecha_registro,doc.total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                  " u.area, m.prioridad AS prioridad, doc.division,m.usuario" & _
                  " from ((((((amarre_documento as a, movi_documento as m ," & tabla & " AS doc)" & _
                  " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                  " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join 3cnuser as u" & _
                  " on (m.usuario = u.usuario_id)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                  " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                  " on (doc.auxiliar=t.tip_linea )" & _
                  " where (a.identificador=m.identificador) and  (a.identificador=doc.identificador) and" & _
                  " T.codtab = 1 and a.Cod_Fam=" & fam & " " & _
                  "and (cdoc.protegido = 'N' OR (SELECT permiso FROM docsusuario D WHERE D.coddoc=Cdoc.coddoc " & _
                  "and usuario = '" & strUsuarioId & "')=1) "
            If Tipo = "01" Then
                m_Reporte = "Rep_Busqueda_Fac.rpt"
            Else
                m_Reporte = "Rep_Contables.rpt"
            End If
        Case FAMILIA_DOC.ENTIDADES: tabla = "documento_entidades"
            SQL = " select a.Identificador,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                  " concat(doc.serie,'-',doc.correl) AS Documento,doc.Entidad," & _
                  " doc.auxiliar AS auxiliar, t.descrip as desaux,doc.codigo as Codigo,aux.descrip as descod," & _
                  " trim(doc.cenco) AS cenco,c.descrip as descenco," & _
                  " doc.fec_emision,a.fecha_registro,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                  " u.area, m.prioridad AS prioridad, doc.division,m.usuario" & _
                  " from ((((((amarre_documento as a, movi_documento as m ," & tabla & " AS doc)" & _
                  " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                  " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join 3cnuser as u" & _
                  " on (m.usuario = u.usuario_id)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                  " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                  " on (doc.auxiliar=t.tip_linea )" & _
                  " where (a.identificador=m.identificador) and  (a.identificador=doc.identificador) and" & _
                  " T.codtab = 1 and a.Cod_Fam=" & fam & " " & _
                  "and (cdoc.protegido = 'N' OR (SELECT permiso FROM docsusuario D WHERE D.coddoc=Cdoc.coddoc " & _
                  "and usuario = '" & strUsuarioId & "')=1) "
            m_Reporte = "Rep_EntidadesDiv.rpt"
        Case FAMILIA_DOC.GENERALES: tabla = "documento_generales"
            SQL = " select a.Identificador,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                  " concat(doc.serie,'-',doc.correl) AS Documento," & _
                  " doc.auxiliar AS auxiliar, t.descrip as desaux,doc.codigo as Codigo,aux.descrip as descod," & _
                  " doc.atencion,trim(doc.cenco) AS cenco,c.descrip as descenco,doc.mon AS moneda," & _
                  " doc.fec_emision,a.fecha_registro,doc.total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                  " u.area, m.prioridad AS prioridad, doc.division,m.usuario" & _
                  " from ((((((amarre_documento as a, movi_documento as m ," & tabla & " AS doc)" & _
                  " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                  " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join 3cnuser as u" & _
                  " on (m.usuario = u.usuario_id)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                  " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                  " on (doc.auxiliar=t.tip_linea )" & _
                  " where (a.identificador=m.identificador) and  (a.identificador=doc.identificador) and" & _
                  " T.codtab = 1 and a.Cod_Fam=" & fam & " " & _
                  "and (cdoc.protegido = 'N' OR (SELECT permiso FROM docsusuario D WHERE D.coddoc=Cdoc.coddoc " & _
                  "and usuario = '" & strUsuarioId & "')=1) "
             m_Reporte = "Rep_DocGenerales.rpt"
        Case FAMILIA_DOC.ORDENES: tabla = "orden_compra"
            SQL = " select a.Identificador,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                  " doc.correl AS Documento,doc.guia,doc.solicitado,doc.requisicion, doc.encargado," & _
                  " doc.auxiliar AS auxiliar, t.descrip as desaux,doc.codigo as Codigo,aux.descrip as descod," & _
                  " trim(doc.cenco) AS cenco,c.descrip as descenco,doc.tipo_cotizacion,doc.mon AS moneda," & _
                  " doc.fec_emision,a.fecha_registro,doc.total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                  " u.area, m.prioridad AS prioridad, doc.division,m.usuario" & _
                  " from ((((((amarre_documento as a, movi_documento as m ," & tabla & " AS doc)" & _
                  " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                  " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join 3cnuser as u" & _
                  " on (m.usuario = u.usuario_id)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                  " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                  " on (doc.auxiliar=t.tip_linea )" & _
                  " where (a.identificador=m.identificador) and  (a.identificador=doc.identificador) and" & _
                  " T.codtab = 1 and a.Cod_Fam=" & fam & " " & _
                  "and (cdoc.protegido = 'N' OR (SELECT permiso FROM docsusuario D WHERE D.coddoc=Cdoc.coddoc " & _
                  "and usuario = '" & strUsuarioId & "')=1) "
            m_Reporte = "Rep_Ordenes.rpt"
    End Select
    If Tipo <> Empty Then
        SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    End If
    If Estado <> Empty Then
        SQL = SQL & " and m.cod_estado='" & Estado & "'"
    End If
    If Auxiliar <> Empty Then
        SQL = SQL & " and doc.auxiliar='" & Auxiliar & "'"
    End If
    If Codauxil <> Empty Then
        SQL = SQL & " and doc.codigo='" & Codauxil & "'"
    End If
    If cenco <> Empty Then
        SQL = SQL & " and trim(doc.cenco)='" & cenco & "'"
    End If
    If Division <> Empty Then
        SQL = SQL & " and doc.division='" & Division & "'"
    End If
    If fecha <> "__/__/____" Then
        SQL = SQL & " and doc.Fec_emision='" & Format(fecha, "yyyy/mm/dd") & "' "
    End If
    If anomesini <> Empty Or anomesfin <> Empty Then
        If anomesini <> Empty And anomesfin = Empty Then
            SQL = SQL & " and a.anomes>='" & anomesini & "' and a.anomes<='" & anomesini & "'"
        End If
        If anomesini = Empty And anomesfin <> Empty Then
            SQL = SQL & " and a.anomes>='" & anomesfin & "' and a.anomes<='" & anomesfin & "'"
        End If
        If anomesini <> Empty And anomesfin <> Empty Then
            SQL = SQL & " and a.anomes>='" & anomesini & "' and a.anomes<='" & anomesfin & "'"
        End If
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No existe ninguna transaccion", vbInformation + vbOKOnly, "Movimiento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Firmas()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select * from cnfirmas"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de Responsable", vbInformation + vbOKOnly, " Firmas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_CuentasxPagarDoc(TipoDoc As String, Estado As String, aux As String, Tipo As String, FechaIni As String, FechaFin As String)
    Dim SQL As String, SQL1 As String
    Dim I As Integer
    Dim myrs As MYSQL_RS
    SQL1 = "Select '' as identifcador,'' as Fecha_Registro,a.Empresa,a.Cod_Fam,a.cod_tipo_doc,'' as Documento,dc.Auxiliar,cn.descrip as TipoAux, '' as codigo, " & _
           "(case a.cod_tipo_doc when 'LS' then 'LIQUIDACION POR TIEMPO DE SERVICIOS' when 'PL' then 'PLANILLA' " & _
           "when 'SS' then if(INSTR(dc.serie,'VAC')>0,'VACACIONES',dc.serie) else aux.descrip end) as NomAux, " & _
           "trim(dc.Cenco) as cenco,'' as Ref,ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.lote where idlote=ct.lote) as lote,(select descripcioncorta from novperuvhse.pozo where idpozo=ct.Pozo) as Pozo,dc.encargado,aux.descrip as NomEncargado, " & _
           "'' as Fec_Emision,dc.division,'' as Obs,'' as Orden,'' as Guia,dc.mon as mon,'' as Fec_Vcto,'' as dias_Vcto, " & _
           "'0' as diasVencidos,dc.Fec_Pago,'' as Fec_Contabilizada ,sum(dc.SubTotal) as subtotal, " & _
           "sum(dc.IGV) as igv, sum(dc.Otros_Montos) as otros_montos,sum(dc.Total) as total,'' as Voucher,mov.cod_Estado as cod_Estado,'' as Fecha_Movi " & _
           "from (((((amarre_documento as a LEFT Join " & _
           "documento_contables as dc on (a.Identificador=dc.Identificador)) " & _
           "LEFT Join cncosto as ct on (dc.cenco=ct.cenco)) RIGHT join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'S')) K " & _
           "on (a.Cod_Tipo_Doc=k.coddoc)) LEFT Join cnauxil as aux " & _
           "on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn " & _
           "on (dc.auxiliar=cn.tip_linea and cn.codtab='1'))LEFT join movi_documento as mov " & _
           "on (dc.Identificador=mov.Identificador) where a.flag=1 and mov.cod_estado<>'" & ANULADO & "' and mov.cod_estado<>'" & IMPRESO & "' and mov.cod_estado<>'" & ELIMINADO & "' and mov.cod_estado<>'" & CANCELADO & "' and "
    SQL = "Select a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          "cn.descrip as TipoAux,dc.codigo,(case a.cod_tipo_doc when 'LS' then 'LIQUIDACION POR TIEMPO DE SERVICIOS' when 'PL' then 'PLANILLA' " & _
           "when 'SS' then if(INSTR(dc.serie,'VAC')>0,'VACACIONES',dc.serie) else aux.descrip end) as NomAux,trim(dc.Cenco) as cenco,dc.Ref," & _
          "ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.pozo where idpozo=ct.Pozo) as Pozo,(select descripcioncorta from novperuvhse.lote where idlote=ct.lote) as lote,dc.encargado,aux.descrip as NomEncargado," & _
          "dc.Fec_Emision,dc.division,dc.Obs, dc.Orden, dc.Guia,dc.mon,dc.Fec_Vcto," & _
          "dc.dias_Vcto,'0' as diasVencidos,dc.Fec_Pago,dc.Fec_Contabilizada ,dc.SubTotal,dc.IGV," & _
          "dc.Otros_Montos,dc.Total,dc.Voucher as Voucher,mov.cod_Estado," & _
          "mov.Fecha_Movi from (((((amarre_documento as a LEFT Join " & _
          "documento_contables as dc on (a.Identificador=dc.Identificador)) " & _
          "LEFT Join cncosto as ct on (dc.cenco=ct.cenco)) RIGHT join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N')) K " & _
          "on (a.Cod_Tipo_Doc=k.coddoc)) LEFT Join cnauxil as aux " & _
          "on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn " & _
          "on (dc.auxiliar=cn.tip_linea and cn.codtab='1'))LEFT join movi_documento as mov " & _
          "on (dc.Identificador=mov.Identificador) where a.flag=1 and mov.cod_estado<>'" & ANULADO & "' and mov.cod_estado<>'" & IMPRESO & "' and mov.cod_estado<>'" & ELIMINADO & "' and mov.cod_estado<>'" & CANCELADO & "' and "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " dc.Fec_Pago>='" & FechaIni & "' and dc.Fec_Pago<='" & FechaFin & "'"
        SQL1 = SQL1 & " dc.Fec_Pago>='" & FechaIni & "' and dc.Fec_Pago<='" & FechaFin & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " dc.Fec_Pago>='" & FechaFin & "' and dc.Fec_Pago<>''"
            SQL1 = SQL1 & " dc.Fec_Pago>='" & FechaFin & "' and dc.Fec_Pago<>''"
            m_Titulo = "PROGRAMACION DE CUENTAS POR PAGAR DESDE " & FechaIni
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " dc.Fec_Pago<='" & FechaFin & "' and dc.Fec_Pago<>''"
                SQL1 = SQL1 & " dc.Fec_Pago<='" & FechaFin & "' and dc.Fec_Pago<>''"
                m_Titulo = "PROGRAMACION DE CUENTAS POR PAGAR AL " & FechaFin
            Else
                SQL = SQL & " Left(dc.Fec_Pago,7)>='" & strAnoSistema & "/" & strMesSistema & "' "
                SQL1 = SQL1 & " Left(dc.Fec_Pago,7)>='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If TipoDoc <> "" Then
        SQL = SQL & "and a.cod_tipo_doc='" & TipoDoc & "'"
        SQL1 = SQL1 & "and a.cod_tipo_doc='" & TipoDoc & "'"
    End If
    If Estado <> "" Then
        SQL = SQL & " and mov.cod_estado='" & Estado & "'"
        SQL1 = SQL1 & " and mov.cod_estado='" & Estado & "'"
    End If
    If Tipo <> "" Then
        SQL = SQL & "and dc.codigo='" & Tipo & "'"
        SQL1 = SQL1 & "and dc.codigo='" & Tipo & "'"
    End If
    If aux <> "" Then
        SQL = SQL & "and dc.Auxiliar='" & aux & "'"
        SQL1 = SQL1 & "and dc.Auxiliar='" & aux & "'"
    End If
    SQL1 = SQL1 & " group by nomaux"
    SQL = SQL1 & " union all " & SQL
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos ", vbInformation + vbOKOnly, " Cuentas xPagar"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_Rep_CuentasxPagarFec(TipoDoc As String, Estado As String, FechaIni As String, FechaFin As String, CodDiv As String, CodAux As String)
    Dim SQL As String, SQL1 As String, SQLDIV As String, SQLAUX As String, SQLRPT1 As String, SQLRPT2 As String, SQLRPT3 As String
    Dim I As Integer
    Dim myrs As MYSQL_RS
    ' if(dc.division<>'0000',dc.division,IFNULL((SELECT division FROM CNCOSTO WHERE COD_AUXIL=trim(dc.Cenco) LIMIT 1),'0000'))
    SQL1 = "Select dc.Identificador as identifcador,'' as Fecha_Registro,a.Empresa,a.Cod_Fam,a.cod_tipo_doc,'' as Documento,dc.Auxiliar,cn.descrip as TipoAux, '' as codigo, " & _
           "(case a.cod_tipo_doc when 'LS' then 'LIQUIDACION POR TIEMPO DE SERVICIOS' when 'PL' then 'PLANILLA' " & _
           "when 'SS' then if(INSTR(dc.serie,'VAC')>0,'VACACIONES',dc.serie) else aux.descrip end) as NomAux, " & _
           "trim(dc.Cenco) as cenco,'' as Ref,ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.pozo where idpozo=ct.Pozo) as Pozo,(select descripcioncorta from novperuvhse.lote where idlote=ct.Lote) as Lote,dc.encargado,enc.NomEncargado, " & _
           "'' as Fec_Emision,dc.division as division,'' as Obs,'' as Orden,'' as Guia,dc.mon as mon,'' as Fec_Vcto,'' as dias_Vcto, " & _
           "'0' as diasVencidos,dc.Fec_Pago,'' as Fec_Contabilizada ,sum(dc.SubTotal) as subtotal, " & _
           "sum(dc.IGV) as igv, sum(dc.Otros_Montos) as otros_montos,sum(dc.Total) as total,'' as Voucher,mov.cod_Estado as cod_Estado,'' as Fecha_Movi, " & _
           "(select if(f.Dias_Vcto='0','CONTADO',CONCAT(f.Dias_Vcto,' dias')) from DOCUMENTO_CONTABLES as f where f.identificador=dc.identificador) as TCredito  " & _
           "from ((((((amarre_documento as a LEFT Join documento_contables as dc on (a.Identificador=dc.Identificador)) " & _
           "LEFT Join cncosto as ct on (dc.cenco=ct.cenco)) LEFT Join encargados as enc on (dc.encargado=enc.codigo)) " & _
           "LEFT Join cnauxil as aux on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo)) " & _
           "LEFT join cntablas as cn on (dc.auxiliar=cn.tip_linea and cn.codtab='1')) LEFT join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'S')) K on (a.Cod_Tipo_Doc=k.coddoc)) " & _
           "LEFT join movi_documento as mov on (dc.Identificador=mov.Identificador) " & _
           "where a.flag=1  and mov.cod_estado<>'" & ANULADO & "' and mov.cod_estado<>'" & IMPRESO & "' and mov.cod_estado<>'" & ELIMINADO & "' and mov.cod_estado<>'" & CANCELADO & "' and "
           
    SQL = "Select a.Identificador as identifcador,a.Fecha_Registro, a.Empresa,a.Cod_Fam,a.cod_tipo_doc, " & _
          "CONCAT(dc.Serie,'-',dc.Correl) as Documento,if(dc.codigo in (select descrip from prov_serv),'0',dc.Auxiliar) as auxiliar,if(dc.codigo in (select descrip from prov_serv),'_SERVICIOS PUBLICOS',cn.descrip) as TipoAux,dc.codigo, " & _
          "if(k.coddoc<>'',aux.descrip,case a.cod_tipo_doc when 'LS' then 'LIQUIDACION POR TIEMPO DE SERVICIOS' " & _
          "when 'PL' then 'PLANILLA' else if(INSTR(dc.serie,'VAC')>0,'VACACIONES',dc.serie) end) as NomAux, " & _
          "trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.pozo where idpozo=ct.Pozo) as Pozo,(select descripcioncorta from novperuvhse.lote where idlote=ct.Lote) as Lote,dc.encargado,enc.NomEncargado,dc.Fec_Emision, " & _
          "dc.division as division,dc.Obs, dc.Orden, dc.Guia,dc.mon,dc.Fec_Vcto,dc.dias_Vcto,'0' as diasVencidos,dc.Fec_Pago, " & _
          "dc.Fec_Contabilizada,IF(k.coddoc<>'',dc.SubTotal,0) as subtotal,if(k.coddoc<>'',dc.IGV,0) as igv, " & _
          "if(k.coddoc<>'',dc.Otros_Montos,0) as otros_montos,if(k.coddoc<>'',dc.Total,0) as total, " & _
          "dc.Voucher as Voucher,mov.cod_Estado,mov.Fecha_Movi, " & _
          "(select if(f.Dias_Vcto='0','CONTADO',CONCAT(f.Dias_Vcto,' dias')) from DOCUMENTO_CONTABLES as f where f.identificador=dc.identificador) as TCredito  " & _
          "from ((((((amarre_documento as a LEFT Join " & _
          "documento_contables as dc on (a.Identificador=dc.Identificador)) " & _
          "LEFT Join cncosto as ct on (dc.cenco=ct.cenco)) LEFT Join encargados as enc " & _
          "on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux " & _
          "on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn " & _
          "on (dc.auxiliar=cn.tip_linea and cn.codtab='1')) LEFT join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N')) K " & _
          "on (a.Cod_Tipo_Doc=k.coddoc)) LEFT join movi_documento as mov " & _
          "on (dc.Identificador=mov.Identificador) where a.flag=1  and mov.cod_estado<>'" & ANULADO & "' and mov.cod_estado<>'" & IMPRESO & "' and mov.cod_estado<>'" & ELIMINADO & "' and mov.cod_estado<>'" & CANCELADO & "' and "
          
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & "dc.Fec_Pago>='" & FechaIni & "' and dc.Fec_Pago<='" & FechaFin & "'"
        SQL1 = SQL1 & "dc.Fec_Pago>='" & FechaIni & "' and dc.Fec_Pago<='" & FechaFin & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " dc.Fec_Pago>='" & FechaIni & "' and dc.Fec_Pago<>''"
            SQL1 = SQL1 & " dc.Fec_Pago>='" & FechaIni & "' and dc.Fec_Pago<>''"
            m_Titulo = "PROYECTADO AL " & FechaFin
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " dc.Fec_Pago<='" & FechaFin & "' and dc.Fec_Pago<>''"
                SQL1 = SQL1 & " dc.Fec_Pago<='" & FechaFin & "' and dc.Fec_Pago<>''"
                m_Titulo = "PROYECTADO AL " & Mid(FechaFin, 9, 2) & "/" & Mid(FechaFin, 6, 2) & "/" & Mid(FechaFin, 1, 4)
            Else
                SQL = SQL & " Left(dc.Fec_Pago,7)='" & strAnoSistema & "/" & strMesSistema & "' "
                SQL1 = SQL1 & " Left(dc.Fec_Pago,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If TipoDoc <> "" Then
        SQL = SQL & " and a.cod_tipo_doc='" & TipoDoc & "'"
        SQL1 = SQL1 & " and a.cod_tipo_doc='" & TipoDoc & "'"
    End If
    If Estado <> "" Then
        SQL = SQL & " and mov.cod_Estado='" & Estado & "'"
        SQL1 = SQL1 & " and mov.cod_Estado='" & Estado & "'"
    End If
    
    If CodDiv <> "0006" Then
     SQLDIV = "Z.division=" & CodDiv
    Else
     SQLDIV = "1=1"
    End If
    
    If CodAux <> "" Then
     SQLAUX = "Z.Auxiliar=" & CodAux
    Else
     SQLAUX = "1=1"
    End If
    
    'REPORTE POR DIVISION
    'SQL1 = SQL1 & " group by division,nomaux,fec_pago,dc.mon "
    SQLRPT1 = "Select Z.identifcador,Z.Fecha_Registro,Z.Empresa,Z.Cod_Fam,(select mid(descrip,1,16) from cndocum WHERE coddoc=Z.cod_tipo_doc) as cod_tipo_doc,Z.Documento,Z.Auxiliar,Z.TipoAux," & _
         "Z.codigo,Z.NomAux,Z.cenco,Z.Ref,Z.NomCenco,Z.Pozo,Z.Lote,Z.encargado,Z.NomEncargado,Z.Fec_Emision,(SELECT DESCRIP FROM CNMDEPAR WHERE CODDEP=Z.division) as division," & _
         "Z.Obs,Z.Orden,Z.Guia,Z.mon,Z.Fec_Vcto,Z.dias_Vcto,Z.diasVencidos,if(Z.Fec_Pago=Z.Fecha_Registro,DATE_ADD(Z.Fecha_Registro,INTERVAL 1 DAY),Z.Fec_Pago) as Fec_Pago,Z.Fec_Contabilizada,Z.subtotal," & _
         "Z.igv,Z.otros_montos,Z.total,Z.voucher,Z.cod_Estado,Z.Fecha_Movi,Z.TCredito from (" & _
         " " & SQL & " ) as Z where " & SQLDIV & " and  " & SQLAUX & " group by Z.auxiliar,Z.TipoAux,Z.division,Z.nomaux,Z.documento,Z.fec_pago,Z.mon HAVING Z.total>0 ORDER BY Z.auxiliar,Z.TipoAux,Z.division,Z.Fec_pago,Z.nomaux"
    Set myrs = oConexion.EjecutaSelectRS(SQLRPT1)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos ", vbInformation + vbOKOnly, " Cuentas xPagar"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    Set myrs = Nothing
     
     
    'REPORTE GENERAL
    SQLRPT2 = "Select Z.identifcador,Z.Fecha_Registro,Z.Empresa,Z.Cod_Fam,(select mid(descrip,1,16) from cndocum WHERE coddoc=Z.cod_tipo_doc) as cod_tipo_doc,Z.Documento,Z.Auxiliar,Z.TipoAux," & _
         "Z.codigo,concat(Z.codigo,'-',Z.NomAux) as NomAux,Z.cenco,Z.Ref,Z.NomCenco,Z.Pozo,Z.Lote,Z.encargado,Z.NomEncargado,Z.Fec_Emision,(SELECT DESCRIP FROM CNMDEPAR WHERE CODDEP=Z.division) as division," & _
         "Z.Obs,Z.Orden,Z.Guia,Z.mon,Z.Fec_Vcto,Z.dias_Vcto,Z.diasVencidos,if(Z.Fec_Pago=Z.Fecha_Registro,DATE_ADD(Z.Fecha_Registro,INTERVAL 1 DAY),Z.Fec_Pago) as Fec_Pago,Z.Fec_Contabilizada,Z.subtotal," & _
         "Z.igv,Z.otros_montos,Z.total,Z.voucher,Z.cod_Estado,Z.Fecha_Movi,Z.TCredito from (" & _
         " " & SQL & " ) as Z where " & SQLAUX & "  group by Z.auxiliar,Z.TipoAux,Z.nomaux,Z.Orden,Z.documento,Z.fec_pago,Z.mon HAVING Z.total>0  ORDER BY Z.auxiliar,Z.TipoAux,Z.Fec_pago,Z.nomaux"
    Set myrs = oConexion.EjecutaSelectRS(SQLRPT2)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos ", vbInformation + vbOKOnly, " Cuentas xPagar"
        Set myrs = Nothing
        Exit Sub
    End If
    m_Reporte = "Rep_CuetxPagarxFecha_general.rpt"
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    
    Set myrs = Nothing
    
    
    'REPORTE GENERAL SIN DOWNHOLE
    
'    m_Titulo = "PROGRAMACION DE PAGOS (SIN DOWNHOLE)  AL " & Mid(FechaFin, 9, 2) & "/" & Mid(FechaFin, 6, 2) & "/" & Mid(FechaFin, 1, 4)
'    SQLRPT3 = "Select Z.identifcador,Z.Fecha_Registro,Z.Empresa,Z.Cod_Fam,(select mid(descrip,1,16) from cndocum WHERE coddoc=Z.cod_tipo_doc) as cod_tipo_doc,Z.Documento,Z.Auxiliar,Z.TipoAux," & _
'         "Z.codigo,Z.NomAux,Z.cenco,Z.Ref,Z.NomCenco,Z.Pozo,Z.Lote,Z.encargado,Z.NomEncargado,Z.Fec_Emision,(SELECT DESCRIP FROM CNMDEPAR WHERE CODDEP=Z.division) as division," & _
'         "Z.Obs,Z.Orden,Z.Guia,Z.mon,Z.Fec_Vcto,Z.dias_Vcto,Z.diasVencidos,if(Z.Fec_Pago=Z.Fecha_Registro,DATE_ADD(Z.Fecha_Registro,INTERVAL 1 DAY),Z.Fec_Pago) as Fec_Pago,Z.Fec_Contabilizada,Z.subtotal," & _
'         "Z.igv,Z.otros_montos,Z.total,Z.voucher,Z.cod_Estado,Z.Fecha_Movi,Z.TCredito from (" & _
'         " " & SQL & " ) as Z where " & SQLAUX & " AND Z.division<>'0003' AND Z.division<>'0010' AND Z.division<>'0011' AND Z.division<>'0014' AND Z.division<>'0016' AND Z.division<>'0017' AND Z.division<>'2001' AND Z.division<>'9140' AND Z.division<>'9162'  group by Z.auxiliar,Z.TipoAux,Z.nomaux,Z.documento,Z.fec_pago,Z.mon ORDER BY Z.auxiliar,Z.TipoAux,Z.Fec_pago,Z.nomaux"
'    Set myrs = oConexion.EjecutaSelectRS(SQLRPT3)
'    If myrs.RecordCount <= 0 Then
'        MsgBox "No Existe Datos ", vbInformation + vbOKOnly, " Cuentas xPagar"
'        Set myrs = Nothing
'        Exit Sub
'    End If
'    m_Reporte = "Rep_CuetxPagarxFecha_general.rpt"
'    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
'
'    Set myrs = Nothing
End Sub
Public Sub SP_Rep_ChequesEmitidos(Ident As String)
    Dim SQL As String, fecha As String, banco As String, cuenta As String, cheque As String, Nombre As String
    Dim mon As String, obs As String, Importe As String ', monto As String
    Dim myrs As MYSQL_RS
    Dim regs As Integer, I As Integer
    oConexion.EjecutaInsertUpdateDelete "delete from TempCheques", TIPO_QUERY.Eliminar, False
    SQL = "Insert Into TempCheques Select a.Identificador,a.cod_tipo_doc,ch.cuenta,ch.anexo,ch.documento," & _
          " ch.referencia,ch.monto,dc.mon,ct.descripcion as NCuenta," & _
          " ct.Banco as Banco,dc.auxiliar,dc.codigo,dc.Obs,dc.Fec_Emision,dc.total as Importe," & _
          " dc.guia as cheque,aux.Descrip as Nombre " & _
          " from (((amarre_documento as a LEFT JOIN documento_contables as dc " & _
          " on (a.Identificador=dc.Identificador))LEFT JOIN detallecheque as ch " & _
          " on(dc.Identificador=ch.Identificador))LEFT JOIN cnauxil as aux " & _
          " on(dc.auxiliar=aux.auxiliar and dc.codigo=aux.codigo)) LEFT JOIN cuentas " & _
          " as ct on (dc.Orden=ct.codigo) where a.identificador='" & Ident & "' order by ch.item"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
    SQL = "Select * from TempCheques "
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount < 0 Then
        MsgBox "No Existe Datos para emitir cheque", vbInformation + vbOKOnly, " Cuentas xPagar"
        Set myrs = Nothing
        Exit Sub
    End If
    myrs.MoveFirst
    Do While Not myrs.EOF()
        If Len(myrs.Fields("referencia")) < 25 Then
            SQL = "Update tempcheques set referencia='" & Left(RTrim(myrs.Fields("referencia")) + Space(50), 50) & "'"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
        End If
        myrs.MoveNext
    Loop
    regs = (myrs.RecordCount / 2)
    myrs.MoveFirst
    fecha = myrs.Fields("Fec_emision")
    banco = myrs.Fields("banco")
    cuenta = myrs.Fields("ncuenta")
    obs = myrs.Fields("obs")
    Nombre = Left(Trim(myrs.Fields("nombre")) & " " & Replace(Space(75), " ", "*"), 75)
    mon = myrs.Fields("mon")
    cheque = myrs.Fields("cheque")
    Importe = myrs.Fields("Importe")
    For I = regs To 16
        myrs.AddNew
        myrs.Fields("Identificador") = Ident
        myrs.Fields("Fec_emision") = fecha
        myrs.Fields("Banco") = banco
        myrs.Fields("ncuenta") = cuenta
        myrs.Fields("cheque") = cheque
        myrs.Fields("mon") = mon
        myrs.Fields("nombre") = Nombre
        myrs.Fields("obs") = obs
        myrs.Fields("referencia") = Space(45)
        myrs.Fields("Importe") = Importe
        myrs.Update
    Next
    myrs.CloseRecordset
    SQL = "Select * from TempCheques "
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub SP_Rep_DocFacturacion(Ident As String, TipoDoc As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    
    SQL = "Select a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar,dc.Contrato," & _
          "cn.descrip as TipoAux,aux.ruc as codigo,aux.descrip as NomAux,aux.Direcc As Direccion," & _
          "trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,dt.Pozo as pozo,(select descripcioncorta from novperuvhse.lote where idlote=dt.lote) as lote,dc.encargado,enc.NomEncargado,dc.Fec_Emision," & _
          "dc.division,dc.Obs, dc.Orden, dc.Guia,dc.mon,dc.Fec_Vcto,dc.dias_Vcto," & _
          "dc.Fec_Pago,dc.Fec_Contabilizada ,dc.SubTotal,dc.IGV, dc.Otros_Montos,dc.Total," & _
          "dc.Voucher as Voucher,dt.item,dt.idsertar,dt.descripcion as Detalle,dt.Cantidad,dt.Monto,dt.Total as Valor " & _
          "from ((((amarre_documento as a LEFT Join documento_contables as dc " & _
          "on (a.Identificador=dc.Identificador)) LEFT Join cncosto as ct " & _
          "on (dc.cenco=ct.cenco)) LEFT Join encargados as enc " & _
          "on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux " & _
          "on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn " & _
          "on (dc.auxiliar=cn.tip_linea)LEFT Join detallefact as dt " & _
          "on (dc.Identificador=dt.Identificador) where cn.codtab='1' and a.identificador='" & Ident & "' and a.cod_tipo_doc='" & TipoDoc & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de documento registrado", vbInformation + vbOKOnly, " Facturacion"
        Set myrs = Nothing
        Exit Sub
    End If
   
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    
    Set myrs = Nothing
End Sub


Public Sub SP_Rep_DocFacturacion_Masivo(AnioMes As String, TipoDoc As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    
    SQL = "Select a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar,dc.Contrato," & _
          "cn.descrip as TipoAux,aux.ruc as codigo,aux.descrip as NomAux,aux.Direcc As Direccion," & _
          "trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,dt.Pozo as pozo,(select descripcioncorta from novperuvhse.lote where idlote=dt.lote) as lote,dc.encargado,enc.NomEncargado,dc.Fec_Emision," & _
          "dc.division,dc.Obs, dc.Orden, dc.Guia,dc.mon,dc.Fec_Vcto,dc.dias_Vcto," & _
          "dc.Fec_Pago,dc.Fec_Contabilizada ,dc.SubTotal,dc.IGV, dc.Otros_Montos,dc.Total," & _
          "dc.Voucher as Voucher,dt.item,dt.idsertar,dt.descripcion as Detalle,dt.Cantidad,dt.Monto,dt.Total as Valor " & _
          "from ((((amarre_documento as a LEFT Join documento_contables as dc " & _
          "on (a.Identificador=dc.Identificador)) LEFT Join cncosto as ct " & _
          "on (dc.cenco=ct.cenco)) LEFT Join encargados as enc " & _
          "on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux " & _
          "on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn " & _
          "on (dc.auxiliar=cn.tip_linea)LEFT Join detallefact as dt " & _
          "on (dc.Identificador=dt.Identificador) where dc.division='0007' and cn.codtab='1' and a.identificador in (SELECT identificador FROM facturacionWEB WHERE tipodoc= '" & TipoDoc & "'  and left(anomes,4)= left('" & AnioMes & "',4) and estado<>'AN' and estado<>'EL') and a.cod_tipo_doc='" & TipoDoc & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de documento registrado", vbInformation + vbOKOnly, " Facturacion"
        Set myrs = Nothing
        Exit Sub
    End If
   
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    
    Set myrs = Nothing
End Sub

Public Sub sp_EstadosCheques(Estado As String, aux As String, Tipo As String, mon As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select DISTINCT a.Identificador,a.Cod_Tipo_Doc AS Tipo,a.Fecha_Registro,a.anomes, " & _
          "CONCAT(doc.Serie,'-',doc.Correl) AS Documento,doc.Orden as CtaCte,if(k.coddoc<>'',doc.Fec_pago,'') as fec_pago, " & _
          "doc.guia as NCheque,doc.auxiliar AS auxiliar,ct.descrip as desaux,IF(k.coddoc<>'',doc.codigo,'') as Codigo, " & _
          "if(k.coddoc<>'',if(e.descripcion='ANULADO','CHEQUE ANULADO',aux.descrip),'') as descod,doc.mon AS moneda,doc.dias_vcto,doc.fec_vcto, " & _
          "doc.fec_emision,if(k.coddoc<>'',doc.total,0) as total,doc.Revisado,doc.Obs,doc.Voucher,mv.Cod_Estado AS estado, " & _
          "e.descripcion as desestado from (((((((amarre_documento as a " & _
          "LEFT Join documento_contables as doc on a.Identificador=doc.Identificador) " & _
          "LEFT Join cntablas as ct on doc.auxiliar=ct.tip_linea) " & _
          "LEFT Join cnauxil as aux on aux.codigo=doc.codigo and aux.auxiliar=doc.auxiliar) " & _
          "LEFT Join movi_documento as mv on doc.Identificador=mv.Identificador) " & _
          "LEFT Join doc_estado as e on mv.cod_estado=e.cod_estado) " & _
          "left join detallecheque d on a.identificador=d.identificador) " & _
          "left join (SELECT coddoc FROM cndocum C WHERE (c.protegido = 'N' OR " & _
          "(SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K " & _
          "on d.TipDoc=k.coddoc) where mv.cod_estado<>'EL' and mv.Cod_Estado<>'CA' and a.cod_tipo_doc='1' and  ct.codtab='1' and "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " doc.Fec_Emision>='" & FechaIni & "' and doc.Fec_Emision<='" & FechaFin & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " doc.Fec_Emision>='" & FechaIni & "' and doc.Fec_Emision<='" & FechaIni & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " doc.Fec_Emision>='" & FechaFin & "' and doc.Fec_Emision<='" & FechaFin & "'"
            Else
                SQL = SQL & " Left(doc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Estado <> "" Then SQL = SQL & " and mv.cod_Estado='" & Estado & "'"
    If Tipo <> "" Then SQL = SQL & " and doc.codigo='" & Tipo & "'"
    If aux <> "" Then SQL = SQL & " and doc.Auxiliar='" & aux & "'"
    If mon <> "" Then SQL = SQL & "and doc.mon='" & mon & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de documento registrado", vbInformation + vbOKOnly, " Reporte de Cheques"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_PagosTelewiesse(NOrden As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select mt.Identificador as Orden,mt.Fecha,mt.Auxiliar,ct.descrip as desaux,IF(K.CODDOC<>'',mt.codigo,'') as Codigo, " & _
          "IF(K.coddoc<>'',aux.descrip,'') as descod,mt.Moneda,mt.CtaCte,mt.TipDoc,mt.FolioRef, " & _
          "if(k.coddoc<>'',mt.Documento,'') as documento,if(k.coddoc<>'',mt.Importe,0) as importe, " & _
          "if(k.coddoc<>'',mt.Importeequ,0) as importeequ,if(k.coddoc<>'',mt.ctaaux,'') as ctaaux, " & _
          "if(k.coddoc<>'',mt.oficina,'') as oficina,mt.TipoPago, tp.descrip, mt.obs, mt.estado" & _
          " from (((mov_telewiese as mt LEFT JOIN cntablas as ct" & _
          " on (mt.auxiliar=ct.tip_linea)) left join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N' OR " & _
          "(SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K " & _
          "on (mt.tipdoc=k.coddoc)) LEFT JOIN cnauxil as aux" & _
          " on (aux.codigo=mt.codigo and aux.auxiliar=mt.auxiliar)) LEFT JOIN tipopago as tp" & _
          " on (mt.TipoPago=tp.Codpago) where ct.codtab='1'"
    If NOrden <> "" Then SQL = SQL & "  and mt.Identificador='" & NOrden & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de documento registrado", vbInformation + vbOKOnly, " Reporte de Cheques "
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_DetalleCtaxAuxDoc(aux As String, caux As String, tdoc As String, sdoc As String, ndoc As String, mon As String, ic As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select a.coddep,a.auxiliar,a.`codaux`," & _
          " (select descrip from `cnauxil` where auxiliar=a.auxiliar and codigo=a.codaux ) as nombre," & _
          " (select descrip from `cndocum` where coddoc=a.coddoc) as desdoc," & _
          " a.serdoc,a.numdoc,a.fecha,a.concepto," & IIf(mon = "E", "sum(a.cargod-a.abonod)", "sum(a.cargos-a.abonos)") & " as MSALDOACTUAL" & _
          " from cnmovi as a " & _
          " where cuenta='420" & ic & "' and a.anomes>='" & strAnoSistema & "00" & "' and a.auxiliar='" & aux & "' and a.codlib<>'06' and a.codaux='" & caux & "' GROUP BY DESDOC,a.serdoc,a.numdoc having  round(sum(a.cargod-a.abonod),2) =0 "  ' and a.coddoc='" & tdoc & "'" & " and a.numdoc='" & ndoc & "' " & _

          
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe datos de documento registrado", vbInformation + vbOKOnly, " Reporte de Cta. Cte Aux-Doc "
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Vacaiones(Mes As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim Anio As String, MesVaca As String
    If Len(Mes) > 2 Then
        MesVaca = Right(Mes, 2)
        Anio = Left(Mes, 4)
    Else
        MesVaca = Mes
        Anio = strAnoSistema
    End If
    SQL = "Select a.codigo,concat_ws(' ',a.apepat,a.apemat, a.nombre1,a.nombre2) as nombre," & _
          " a.codafp,c.gocehaber,c.fec_Salida,c.fec_Regreso,DATEDIFF(c.fec_Regreso,c.fec_Salida) + 1 as dias,  c.periodo" & _
          " from empleado as a left join calendario as c " & _
          " on(c.codemp=a.codigo) " & _
          " where a.tipo not in (3,4) and c.movemp='02' and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & Anio & MesVaca & "'" & _
          " and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & Anio & MesVaca & "'" & _
          " order by a.apepat,a.apemat,a.nombre1,a.nombre2,fec_salida"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No hay vacaciones registradas en el mes", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_Permisos(Mes As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim Anio As String, MesVaca As String
    If Len(Mes) > 2 Then
        MesVaca = Right(Mes, 2)
        Anio = Left(Mes, 4)
    Else
        MesVaca = Mes
        Anio = strAnoSistema
    End If
    SQL = "Select a.codigo,concat_ws(' ',a.apepat,a.apemat, a.nombre1,a.nombre2) as nombre,a.codafp," & _
          "c.fec_Salida,c.hora_salida,c.fec_Regreso,c.hora_regreso,DATEDIFF(c.fec_Regreso,c.fec_Salida) + 1 as dias, TIME_FORMAT( TIME(c.hora_regreso) - TIME(c.hora_salida) ,'%h:%i') as hora,c.observacion" & _
          " from empleado as a left join calendario as c " & _
          " on(c.codemp=a.codigo) " & _
          " where a.tipo not in (3,4) and c.movemp='03' and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & Anio & MesVaca & "'" & _
          " and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & Anio & MesVaca & "'" & _
          " order by a.apepat,a.apemat,a.nombre1,a.nombre2,fec_salida"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No hay permisos registradas en el mes", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_Licencias(Mes As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim Anio As String, MesVaca As String
    If Len(Mes) > 2 Then
        MesVaca = Right(Mes, 2)
        Anio = Left(Mes, 4)
    Else
        MesVaca = Mes
        Anio = strAnoSistema
    End If
    SQL = "Select a.codigo,concat_ws(' ',a.apepat,a.apemat, a.nombre1,a.nombre2) as nombre,a.codafp," & _
          "c.fec_Salida,c.hora_salida,c.fec_Regreso,c.hora_regreso,DATEDIFF(c.fec_Regreso,c.fec_Salida) + 1 as dias, TIME_FORMAT( TIME(c.hora_regreso) - TIME(c.hora_salida) ,'%h:%i') as hora,c.observacion" & _
          " from empleado as a left join calendario as c " & _
          " on(c.codemp=a.codigo) " & _
          " where a.tipo not in (3,4) and c.movemp='05' and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & Anio & MesVaca & "'" & _
          " and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & Anio & MesVaca & "'" & _
          " order by a.apepat,a.apemat,a.nombre1,a.nombre2,fec_salida"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No hay licencias registradas en el mes", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_Subsidios(Mes As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim Anio As String, MesVaca As String
    If Len(Mes) > 2 Then
        MesVaca = Right(Mes, 2)
        Anio = Left(Mes, 4)
    Else
        MesVaca = Mes
        Anio = strAnoSistema
    End If
    SQL = "Select a.codigo,concat_ws(' ',a.apepat,a.apemat, a.nombre1,a.nombre2) as nombre,a.codafp," & _
          "c.fec_Salida,c.hora_salida,c.fec_Regreso,c.hora_regreso,DATEDIFF(c.fec_Regreso,c.fec_Salida) + 1 as dias, TIME_FORMAT( TIME(c.hora_regreso) - TIME(c.hora_salida) ,'%h:%i') as hora,c.observacion" & _
          " from empleado as a left join calendario as c " & _
          " on(c.codemp=a.codigo) " & _
          " where a.tipo not in (3,4) and c.movemp='07' and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & Anio & MesVaca & "'" & _
          " and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & Anio & MesVaca & "'" & _
          " order by a.apepat,a.apemat,a.nombre1,a.nombre2,fec_salida"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No hay subsidios registradas en el mes", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_PagosTeleWxFecha(FechaIni As String, FechaFin As String, codRUC As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select mt.Identificador as Orden,mt.Fecha,mt.Auxiliar,ct.descrip as desaux,IF(k.coddoc<>'',mt.codigo,'') as Codigo, " & _
          "if(k.coddoc <>'',CONCAT('OC:',ifnull((select orden from documento_contables where identificador=mt.FolioRef limit 1),'-'),'  ',aux.descrip),'') as descod,mt.Moneda,CONCAT(mt.vou,'-',mt.CtaCte) as CtaCte,mt.TipDoc,mt.FolioRef, " & _
          "if(k.coddoc<>'',mt.Documento,'') as documento,if(k.coddoc<>'',mt.Importe,0) as importe, " & _
          "if(k.coddoc <>'',mt.Importeequ,0) as importeequ,if(k.coddoc<>'',mt.ctaaux,'') as ctaaux, " & _
          "if(k.coddoc<>'',mt.oficina,'') as oficina, mt.TipoPago, tp.descrip, mt.obs, mt.estado" & _
          " from ((((mov_telewiese as mt LEFT JOIN cntablas as ct" & _
          " on mt.auxiliar=ct.tip_linea) LEFT JOIN cnauxil as aux" & _
          " on aux.codigo=mt.codigo and aux.auxiliar=mt.auxiliar) LEFT JOIN tipopago as tp" & _
          " on mt.TipoPago=tp.Codpago) left join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N' OR " & _
          "(SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K " & _
          "on (mt.tipdoc=k.coddoc)) where mt.estado = 'TR' and ct.codtab='1' and "
    If codRUC <> "" Then
       SQL = SQL & " mt.codigo='" & Right("00000000000" & codRUC, 11) & "' ORDER BY mt.Fecha DESC"
       m_Titulo = "PROGRAMACION DE PAGOS TELEWISSE REALIZADOS POR FECHA DEL CODIGO:" & Right("00000000000" & codRUC, 11)
    Else
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & "  mt.fecha>='" & Format(FechaIni, "yyyy/mm/dd") & "' and mt.fecha<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " mt.fecha>='" & Format(FechaIni, "yyyy/mm/dd") & "' and mt.fecha<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " mt.fecha>='" & Format(FechaFin, "yyyy/mm/dd") & "' and mt.fecha<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " Left(mt.fecha,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    End If
    
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de documento registrado", vbInformation + vbOKOnly, " Reporte de Cheques"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_DocumentosRegistradosXFecha(FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select" & _
          " a.identificador,concat_ws('-',a.serie,a.correl) as numdoc,a.orden,a.auxiliar,a.codigo, " & _
          " d.descrip," & _
          " a.mon,a.fec_emision," & _
          " a.total,a.voucher,(select descrip from cnmdepar where coddep= a.division) as divi," & _
          " a.ref,(Select descrip from cndocum where coddoc=b.cod_tipo_doc) as doc,b.fecha_registro,b.usuario,C.COD_ESTADO" & _
          " from ((documento_contables AS A LEFT JOIN amarre_documento AS B" & _
          " ON (A.IDENTIFICADOR=B.IDENTIFICADOR)) left join movi_documento as C" & _
          " on (A.IDENTIFICADOR=C.IDENTIFICADOR)) left join cnauxil as D" & _
          " on (a.auxiliar=d.auxiliar and a.codigo=d.codigo)" & _
          " where b.flag='1' AND C.COD_ESTADO<>'EL' AND "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & "  B.fecha_registro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " B.fecha_registro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " B.fecha_registro>='" & Format(FechaFin, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " Left(B.fecha_registro,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de documento registrado", vbInformation + vbOKOnly, " Reporte de Cheques"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
   
End Sub

Public Sub sp_DocumentosRegistradosDiarioXFecha(FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    'Reporte Diario
    SQL = " Select " & _
          " (select descrip from cnmdepar where coddep=d.Division) as divi,d.orden," & _
          "concat(right(B.fecha_registro,2),'/',substring(B.fecha_registro,6,2),'/',left(B.fecha_registro,4)) as fecha_registro," & _
          "concat(right(d.Fec_Emision,2),'/',substring(d.Fec_Emision,3,2),'/',left(d.Fec_Emision,4)) as Fec_Emision," & _
          " concat(d.serie,'-',d.correl) as ndoc," & _
          " (select concat((select If (length(credito)=1,concat('0',credito),credito) from orden_compra where correl=d.orden),descrip) from cndocum where  coddoc=b.cod_tipo_doc)  as Docu," & _
          " (select descrip from cnauxil where auxiliar=d.auxiliar and codigo=d.codigo) as Proveedor, " & _
          " if(d.Mon='N',d.Total,0) as ImporteS, " & _
          " if(d.Mon='E',d.Total,0) as ImporteE,d.identificador as folio, " & _
          " 'XXXXXXX' as RecLogistica, " & _
          " 'XXXXXXX' as RecContab " & _
          " from (documento_contables as d left join  amarre_documento as b ON (d.IDENTIFICADOR=B.IDENTIFICADOR)) " & _
          " left join movi_documento as C on (d.identificador=c.identificador) " & _
          " Where " & _
          " b.flag='1' and d.Division not in ('0003','0010','0011','0014','0016','0017','2001','9140','9162') and  " & _
          " C.COD_ESTADO<>'EL' and b.cod_tipo_doc not in ('SG') and (B.usuario='KATHERINE' OR B.usuario='ROSA' OR B.usuario='MARCELA' OR B.usuario='BERLIZ') and "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & "  B.fecha_registro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " B.fecha_registro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " B.fecha_registro>='" & Format(FechaFin, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " Left(B.fecha_registro,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If

    SQL = SQL & " group by divi,B.fecha_registro,d.Fec_Emision,ndoc,Docu,Proveedor,ImporteS,ImporteE order by (select credito from orden_compra where correl=d.orden),d.identificador "
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de documento registrado", vbInformation + vbOKOnly, " Reporte de Cheques"
        Set myrs = Nothing
        Exit Sub
    End If
    
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
    
    
    'Reporte Control Documentario vs Ordenes Masterys
    m_Titulo = "DOCUMENTOS REGISTRADOS NOVADMIN DE ORDENES GENERADAS - MES  " & NombreMes(strMesSistema, False)
    m_Reporte = "Rep_ControlDiarioRegistroDocumvsOC.rpt"
    SQL = ""
    
    SQL = "Select " & _
    "a.Correl as Orden,a.Fec_Emision as FecOrden,a.cenco,a.mon,a.Total,a.Division,a.cta,o.identificador as folio,o.codigo as ruc,o.serie,o.correl,o.fec_emision,o.Fec_vcto,o.fec_pago,o.Total as TotFac,o.Dias_vcto, " & _
    "(select voucher from cnmovi  where codaux=o.codigo and serdoc=o.serie and numdoc=o.correl and codlib='01' limit 1) as voucher " & _
    "From " & _
    "(select Correl,Fec_Emision,cenco,mon,ValorCompra,IGV,Total,Division,cta " & _
    "From orden_compra " & _
    "where Fec_Emision>=concat('" & strAnoSistema & "','/','" & strMesSistema & "','/','01')  ) as a " & _
    "left join novperuhfm.documento_contables as o on (a.Correl=o.Orden) " & _
    "group by a.Correl"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de Ordenes ingresadas en Control Documentario", vbInformation + vbOKOnly, " Reporte de Cheques"
        Set myrs = Nothing
        Exit Sub
    End If
    
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
    
End Sub

Public Sub sp_DocumentosRegistradosDiarioXFecha_DH(FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    'Reporte Diario
    SQL = " Select" & _
          " (select descrip from cnmdepar where coddep=d.Division) as divi,d.orden," & _
          "concat(right(B.fecha_registro,2),'/',substring(B.fecha_registro,6,2),'/',left(B.fecha_registro,4)) as fecha_registro," & _
          "concat(right(d.Fec_Emision,2),'/',substring(d.Fec_Emision,3,2),'/',left(d.Fec_Emision,4)) as Fec_Emision," & _
          " concat(d.serie,'-',d.correl) as ndoc," & _
          " (select descrip from cndocum where  coddoc=b.cod_tipo_doc)  as Docu," & _
          " (select descrip from cnauxil where auxiliar=d.auxiliar and codigo=d.codigo) as Proveedor, " & _
          " if(d.Mon='N',d.Total,0) as ImporteS, " & _
          " if(d.Mon='E',d.Total,0) as ImporteE,d.identificador as folio, " & _
          " 'XXXXXXX' as RecLogistica, " & _
          " 'XXXXXXX' as RecContab " & _
          " from (documento_contables as d left join  amarre_documento as b ON (d.IDENTIFICADOR=B.IDENTIFICADOR)) " & _
          " left join movi_documento as C on (d.identificador=c.identificador) " & _
          " Where " & _
          " b.flag='1' and d.Division in ('0003','0010','0011','0014','0016','0017','2001','9140','9162') and " & _
          " C.COD_ESTADO<>'EL' and b.cod_tipo_doc not in ('SG') and (B.usuario='KATHERINE' OR B.usuario='ROSA' OR B.usuario='MARCELA' OR B.usuario='BERLIZ') and "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & "  B.fecha_registro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " B.fecha_registro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " B.fecha_registro>='" & Format(FechaFin, "yyyy/mm/dd") & "' and B.fecha_registro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " Left(B.fecha_registro,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    
    SQL = SQL & " group by divi,B.fecha_registro,d.Fec_Emision,ndoc,Docu,Proveedor,ImporteS,ImporteE"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de documento registrado", vbInformation + vbOKOnly, " Reporte de Cheques"
        Set myrs = Nothing
        Exit Sub
    End If
    
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
    
End Sub


Public Sub sp_ReporteVentasDoc(Tipo As String, aux As String, Div As String, alm As String, mone As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim I As Integer
    Dim myrs As MYSQL_RS
    SQL = "Select DISTINCT a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          " a.cod_tipo_doc,(select descrip from cndocum where coddoc=a.cod_tipo_doc) as ddoc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          " dc.Contrato,cn.descrip as TipoAux,dc.codigo,aux.descrip as NomAux," & _
          " aux.Direcc As Direccion,trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,dt.Pozo,dt.lote," & _
          " dc.encargado,enc.NomEncargado,dc.Fec_Emision,right(dc.division,4) as division,divi.descrip,dc.Obs, dc.Orden,dc.Guia," & _
          " dc.mon,dc.Fec_Vcto,dc.dias_Vcto,dc.Fec_Pago,dc.Fec_Contabilizada ," & _
          " (dc.SubTotal * cnd.factor ) as subtotal,(dc.IGV * cnd.factor) as IGV," & _
          " (dc.Otros_Montos * cnd.factor) as Otros_Montos,(dc.Total*cnd.factor) as total," & _
          " (dc.ImpEqui*cnd.factor) as ImpEqui,dc.Voucher as Voucher,mv.cod_Estado,(Select descrip from almacen where codigo=dc.codalmac) as almacen" & _
          " from (((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc" & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt" & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct" & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc" & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux" & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn" & _
          " on (dc.auxiliar=cn.tip_linea) LEFT join movi_documento as mv" & _
          " on (dc.Identificador=mv.Identificador))left join cnmdepar as divi on (dc.division=divi.coddep) where dc.auxiliar='2' and " & _
          " (a.cod_tipo_doc='01' or a.cod_tipo_doc='03'" & _
          " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08' or a.cod_tipo_doc='P') and mv.cod_estado<>'" & ELIMINADO & "' and a.flag='0' and cn.codtab='1' "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaFin, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and Left(dc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Tipo <> "" Then SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    If aux <> "" Then SQL = SQL & "  and dc.codigo='" & aux & "'"
    If Div <> "" Then SQL = SQL & "  and dc.division='" & Div & "'"
    If alm <> "" Then SQL = SQL & "  and dc.codalmac='" & alm & "'"
    If mone <> "" Then SQL = SQL & "  and dc.mon='" & mone & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Documentos para la consuta", vbInformation + vbOKOnly, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub


Public Sub sp_ReporteVentasxFec(Tipo As String, aux As String, Div As String, alm As String, mone As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select DISTINCT a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          "dc.Contrato,cn.descrip as TipoAux,dc.codigo,aux.descrip as NomAux," & _
          "aux.Direcc As Direccion,trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.pozo where idpozo=dt.Pozo) as Pozo,(select descripcioncorta from novperuvhse.lote where idlote=dt.Lote) as Lote," & _
          "dc.encargado,enc.NomEncargado,dc.Fec_Emision,dc.division,divi.descrip,dc.Obs, dc.Orden,dc.Guia," & _
          "dc.mon,dc.Fec_Vcto,dc.dias_Vcto,dc.Fec_Pago,dc.Fec_Contabilizada ," & _
          " (dc.SubTotal * cnd.factor ) as subtotal,(dc.IGV * cnd.factor) as IGV," & _
          " (dc.Otros_Montos * cnd.factor) as Otros_Montos,(dc.Total*cnd.factor) as total," & _
          " (dc.ImpEqui*cnd.factor) as ImpEqui,dc.Voucher as Voucher,mv.cod_Estado,(Select descrip from almacen where codigo=dc.codalmac) as almacen" & _
          " from (((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc" & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt" & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct" & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc" & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux" & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn" & _
          " on (dc.auxiliar=cn.tip_linea) LEFT join movi_documento as mv" & _
          " on (dc.Identificador=mv.Identificador))left join cnmdepar as divi on (dc.division=divi.coddep) where dc.auxiliar='2' and " & _
          " (a.cod_tipo_doc='01' or a.cod_tipo_doc='03'" & _
          " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08' or a.cod_tipo_doc='P' ) and mv.cod_estado<>'" & ELIMINADO & "' and a.flag='0' and cn.codtab='1' "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaFin, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and Left(dc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Tipo <> "" Then SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    If aux <> "" Then SQL = SQL & "  and dc.codigo='" & aux & "'"
    If Div <> "" Then SQL = SQL & "  and dc.division='" & Div & "'"
    If alm <> "" Then SQL = SQL & "  and dc.codalmac='" & alm & "'"
    If mone <> "" Then SQL = SQL & "  and dc.mon='" & mone & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Documentos para la consuta", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_ReporteVentasxCenco(Tipo As String, cenco As String, aux As String, Div As String, alm As String, mone As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select DISTINCT a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          "dc.Contrato,cn.descrip as TipoAux,dc.codigo,aux.descrip as NomAux," & _
          "aux.Direcc As Direccion,trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.pozo where idpozo=dt.Pozo) as Pozo,(select descripcioncorta from novperuvhse.lote where idlote=dt.Lote) as Lote," & _
          "dc.encargado,enc.NomEncargado,dc.Fec_Emision,dc.division,divi.descrip,dc.Obs, dc.Orden,dc.Guia," & _
          "dc.mon,dc.Fec_Vcto,dc.dias_Vcto,dc.Fec_Pago,dc.Fec_Contabilizada ," & _
          " (dc.SubTotal * cnd.factor ) as subtotal,(dc.IGV * cnd.factor) as IGV," & _
          " (dc.Otros_Montos * cnd.factor) as Otros_Montos,(dc.Total*cnd.factor) as total," & _
          " (dc.ImpEqui*cnd.factor) as ImpEqui,dc.Voucher as Voucher,mv.cod_Estado,(Select descrip from almacen where codigo=dc.codalmac) as almacen" & _
          " from (((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc" & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt" & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct" & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc" & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux" & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn" & _
          " on (dc.auxiliar=cn.tip_linea) LEFT join movi_documento as mv" & _
          " on (dc.Identificador=mv.Identificador))left join cnmdepar as divi on (dc.division=divi.coddep) where dc.auxiliar='2' and " & _
          " (a.cod_tipo_doc='01' or a.cod_tipo_doc='03'" & _
          " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08' or a.cod_tipo_doc='P') and mv.cod_estado<>'" & ELIMINADO & "' and a.flag='0' and cn.codtab='1' "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaFin, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and Left(dc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Tipo <> "" Then SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    If aux <> "" Then SQL = SQL & "  and dc.codigo='" & aux & "'"
    If Div <> "" Then SQL = SQL & "  and dc.division='" & Div & "'"
    If alm <> "" Then SQL = SQL & "  and dc.codalmac='" & alm & "'"
    If cenco <> "" Then SQL = SQL & "  and  trim(dc.cenco)='" & cenco & "'"
    If mone <> "" Then SQL = SQL & "  and  dc.mon='" & mone & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Documentos para la consuta", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_ReporteVentasxDivision(Tipo As String, aux As String, Div As String, alm As String, mone As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    
    SQL = "Select DISTINCT a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          "dc.Contrato,cn.descrip as TipoAux,dc.codigo,aux.descrip as NomAux," & _
          "aux.Direcc As Direccion,trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.pozo where idpozo=dt.Pozo) as Pozo,(select descripcioncorta from novperuvhse.lote where idlote=dt.Lote) as Lote," & _
          "dc.encargado,enc.NomEncargado,dc.Fec_Emision,dc.division,divi.descrip,dc.Obs, dc.Orden,dc.Guia," & _
          "dc.mon,dc.Fec_Vcto,dc.dias_Vcto,dc.Fec_Pago,dc.Fec_Contabilizada ," & _
          " (dc.SubTotal * cnd.factor ) as subtotal,(dc.IGV * cnd.factor) as IGV," & _
          " (dc.Otros_Montos * cnd.factor) as Otros_Montos,(dc.Total*cnd.factor) as total," & _
          " (dc.ImpEqui*cnd.factor) as ImpEqui,dc.Voucher as Voucher,mv.cod_Estado,(Select descrip from almacen where codigo=dc.codalmac) as almacen" & _
          " from (((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc" & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt" & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct" & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc" & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux" & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn" & _
          " on (dc.auxiliar=cn.tip_linea) LEFT join movi_documento as mv" & _
          " on (dc.Identificador=mv.Identificador))left join cnmdepar as divi on (dc.division=divi.coddep) where dc.auxiliar='2' and " & _
          " (a.cod_tipo_doc='01' or a.cod_tipo_doc='03'" & _
          " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08' or ((dt.codser <> '065' and a.cod_tipo_doc='P') and (dt.codser <> '068' and a.cod_tipo_doc='P'))) and mv.cod_estado<>'" & ELIMINADO & "' and a.flag='0' and cn.codtab='1' "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaFin, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and Left(dc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Tipo <> "" Then SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    If aux <> "" Then SQL = SQL & "  and dc.codigo='" & aux & "'"
    If Div <> "" And Div <> "0006" Then SQL = SQL & "  and dc.division='" & Div & "'   "
    If alm <> "" Then SQL = SQL & "  and dc.codalmac='" & alm & "'"
    If mone <> "" Then SQL = SQL & "  and dc.mon='" & mone & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Documentos para la consulta", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_ReporteVentasxAuxiliar(Tipo As String, aux As String, Div As String, alm As String, mone As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    
    SQL = "Select DISTINCT a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          "dc.Contrato,cn.descrip as TipoAux,dc.codigo,aux.descrip as NomAux," & _
          "aux.Direcc As Direccion,trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.pozo where idpozo=dt.Pozo) as Pozo,(select descripcioncorta from novperuvhse.lote where idlote=dt.Lote) as Lote," & _
          "dc.encargado,enc.NomEncargado,dc.Fec_Emision,right(dc.division,4) as division,divi.descrip,dc.Obs, dc.Orden,dc.Guia," & _
          "dc.mon,dc.Fec_Vcto,dc.dias_Vcto,dc.Fec_Pago,dc.Fec_Contabilizada ," & _
          " (dc.SubTotal * cnd.factor ) as subtotal,(dc.IGV * cnd.factor) as IGV," & _
          " (dc.Otros_Montos * cnd.factor) as Otros_Montos,(dc.Total*cnd.factor) as total," & _
          " (dc.ImpEqui*cnd.factor) as ImpEqui,dc.Voucher as Voucher,mv.cod_Estado,(Select descrip from almacen where codigo=dc.codalmac) as almacen" & _
          " from (((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc" & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt" & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct" & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc" & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux" & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn" & _
          " on (dc.auxiliar=cn.tip_linea) LEFT join movi_documento as mv" & _
          " on (dc.Identificador=mv.Identificador))left join cnmdepar as divi on (dc.division=divi.coddep) where dc.auxiliar='2' and " & _
          " (a.cod_tipo_doc='01' or a.cod_tipo_doc='03'" & _
          " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08' or a.cod_tipo_doc='P') and mv.cod_estado<>'" & ELIMINADO & "' and a.flag='0' and cn.codtab='1' "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaFin, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and Left(dc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Tipo <> "" Then SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    If aux <> "" Then SQL = SQL & "  and dc.codigo='" & aux & "'"
    If Div <> "" Then SQL = SQL & "  and dc.division='" & Div & "'"
    If alm <> "" Then SQL = SQL & "  and dc.codalmac='" & alm & "'"
    If mone <> "" Then SQL = SQL & "  and dc.mon='" & mone & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Documentos para la consuta", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_ReporteVentasxPozo(Tipo As String, pozo As String, aux As String, Div As String, alm As String, mone As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select DISTINCT a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          "dc.Contrato,cn.descrip as TipoAux,dc.codigo,aux.descrip as NomAux," & _
          "aux.Direcc As Direccion,trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco," & _
          "(select DescripcionCorta from novperuvhse.pozo where IdPozo=dt.Pozo) as Pozo," & _
          "(select DescripcionCorta from novperuvhse.lote where IdLote=dt.lote) as lote," & _
          "dc.encargado,enc.NomEncargado,dc.Fec_Emision,dc.division,divi.descrip,dc.Obs, dc.Orden,dc.Guia," & _
          "dc.mon,dc.Fec_Vcto,dc.dias_Vcto,dc.Fec_Pago,dc.Fec_Contabilizada ," & _
          " (dc.SubTotal * cnd.factor ) as subtotal,(dc.IGV * cnd.factor) as IGV," & _
          " (dc.Otros_Montos * cnd.factor) as Otros_Montos,(dc.Total*cnd.factor) as total," & _
          " (dc.ImpEqui*cnd.factor) as ImpEqui,dc.Voucher as Voucher,mv.cod_Estado,(Select descrip from almacen where codigo=dc.codalmac) as almacen" & _
          " from (((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc" & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt" & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct" & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc" & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux" & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn" & _
          " on (dc.auxiliar=cn.tip_linea) LEFT join movi_documento as mv" & _
          " on (dc.Identificador=mv.Identificador))left join cnmdepar as divi on (dc.division=divi.coddep) where dc.auxiliar='2' and " & _
          " (a.cod_tipo_doc='01' or a.cod_tipo_doc='03'" & _
          " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08' or a.cod_tipo_doc='P' ) and mv.cod_estado<>'" & ELIMINADO & "' and a.flag='0' and cn.codtab='1' "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaFin, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and Left(dc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Tipo <> "" Then SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    If aux <> "" Then SQL = SQL & "  and dc.codigo='" & aux & "'"
    If Div <> "" Then SQL = SQL & "  and dc.division='" & Div & "'"
    If alm <> "" Then SQL = SQL & "  and dc.codalmac='" & alm & "'"
    If pozo <> "" Then SQL = SQL & " and dt.Pozo='" & pozo & "'"
    If mone <> "" Then SQL = SQL & "  and dc.mon='" & mone & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Documentos para la consuta", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_ReporteVentasxLote(Tipo As String, Lote As String, aux As String, Div As String, alm As String, mone As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select DISTINCT a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          "a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          "dc.Contrato,cn.descrip as TipoAux,dc.codigo,aux.descrip as NomAux," & _
          "aux.Direcc As Direccion,trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco," & _
          "(select DescripcionCorta from novperuvhse.pozo where IdPozo=dt.Pozo) as Pozo," & _
          "(select DescripcionCorta from novperuvhse.lote where IdLote=dt.lote) as lote," & _
          "dc.encargado,enc.NomEncargado,dc.Fec_Emision,dc.division,divi.descrip,dc.Obs, dc.Orden,dc.Guia," & _
          "dc.mon,dc.Fec_Vcto,dc.dias_Vcto,dc.Fec_Pago,dc.Fec_Contabilizada ," & _
          " (dc.SubTotal * cnd.factor ) as subtotal,(dc.IGV * cnd.factor) as IGV," & _
          " (dc.Otros_Montos * cnd.factor) as Otros_Montos,(dc.Total*cnd.factor) as total," & _
          " (dc.ImpEqui*cnd.factor) as ImpEqui,dc.Voucher as Voucher,mv.cod_Estado,(Select descrip from almacen where codigo=dc.codalmac) as almacen" & _
          " from (((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc" & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt" & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct" & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc" & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux" & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn" & _
          " on (dc.auxiliar=cn.tip_linea) LEFT join movi_documento as mv" & _
          " on (dc.Identificador=mv.Identificador))left join cnmdepar as divi on (dc.division=divi.coddep) where dc.auxiliar='2' and " & _
          " (a.cod_tipo_doc='01' or a.cod_tipo_doc='03'" & _
          " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08' OR a.cod_tipo_doc='P') and mv.cod_estado<>'" & ELIMINADO & "' and a.flag='0' and cn.codtab='1' "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaFin, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and Left(dc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Tipo <> "" Then SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    If aux <> "" Then SQL = SQL & "  and dc.codigo='" & aux & "'"
    If Div <> "" Then SQL = SQL & "  and dc.division='" & Div & "'"
    If alm <> "" Then SQL = SQL & "  and dc.codalmac='" & alm & "'"
    If Lote <> "" Then SQL = SQL & " and dt.lote='" & Lote & "'"
    If mone <> "" Then SQL = SQL & " and dc.mon='" & mone & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
       MsgBox "No Existe Documentos para la consuta", vbInformation, "Reporte de Ventas"
       Set myrs = Nothing
       Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_ReporteVentasxClasificacion(Tipo As String, Claf As String, aux As String, Div As String, alm As String, mone As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    If mone = "" Then mone = "A"
    
    SQL = "Select a.Identificador,a.Fecha_Registro, a.Empresa,a.Cod_Fam," & _
          " a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Auxiliar," & _
          " dc.Contrato,cn.descrip as TipoAux,dc.codigo,aux.descrip as NomAux," & _
          " aux.Direcc As Direccion,trim(dc.Cenco) as cenco,dc.Ref,ct.descrip as NomCenco,(select descripcioncorta from novperuvhse.pozo where idpozo=dt.Pozo) as Pozo,(select descripcioncorta from novperuvhse.lote where idlote=dt.Lote) as Lote," & _
          " dc.encargado,enc.NomEncargado,dc.Fec_Emision,dc.division,divi.descrip,dc.Obs, dc.Orden,dc.Guia," & _
          " dc.mon,dc.Fec_Vcto,dc.dias_Vcto,dc.Fec_Pago,dc.Fec_Contabilizada ," & _
          " (SUM(dt.Total) * cnd.factor ) as subtotal,(dc.subTotal * cnd.factor ) as subtotalFAC,(dc.igv * cnd.factor) as IGV," & _
          " (dc.Otros_Montos * cnd.factor) as Otros_Montos,(dc.Total*cnd.factor) as total," & _
          " (dc.ImpEqui*cnd.factor) as ImpEqui,dc.Voucher as Voucher,mv.cod_Estado,(Select descrip from almacen where codigo=dc.codalmac) as almacen," & _
          " clf.Codigo as CodClf,clf.Descrip as Servicio,'" & Trim(mone) & "' as tmoneda, (select tipo_ven from cncambio where fecha=concat(right(dc.Fec_Emision,2),substring(dc.Fec_Emision,5,4),left(dc.Fec_Emision,4))) as tipcam " & _
          " from ((((((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc " & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt " & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct " & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc " & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux " & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn " & _
          " on (dc.auxiliar=cn.tip_linea)) LEFT join movi_documento as mv " & _
          " on (dc.Identificador=mv.Identificador))left join servicio as ser " & _
          " on(dt.codser=ser.codigo))left join clasif_serv as clf " & _
          " on(ser.codclasf=clf.codigo)) left join cnmdepar as divi on (dc.division=divi.coddep) " & _
          " where dt.codser<>'000' AND dc.auxiliar='2'  and mv.cod_estado<>'" & ELIMINADO & "' and a.flag='0' and cn.codtab='1' "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaIni, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " and dc.Fec_Emision>='" & Format(FechaFin, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and Left(dc.Fec_Emision,7)='" & strAnoSistema & "/" & strMesSistema & "' "
            End If
        End If
    End If
    If Tipo <> "" Then
        SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    Else
        SQL = SQL & " and (a.cod_tipo_doc='01' or a.cod_tipo_doc='03' " & _
              " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08')"
    End If
    If aux <> "" Then SQL = SQL & "  and dc.codigo='" & aux & "'"
    If Div <> "" Then SQL = SQL & "  and dc.division='" & Div & "'"
    If alm <> "" Then SQL = SQL & "  and dc.codalmac='" & alm & "'"
    If Claf <> "" Then
        SQL = SQL & " and clf.Codigo='" & Claf & "'"
        m_Anomes = "N"
    Else
        m_Anomes = "S"
    End If
    If mone <> "A" Then SQL = SQL & " and dc.mon='" & mone & "'"
    SQL = SQL & "  group by dc.Identificador,clf.Codigo"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Documentos para la consuta", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Reporte_Registros_Servicios()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select * from serv_tarif order by Codaux, IDSerTar"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Datos de Registros de Servicios y Tarifas", vbInformation, "Reporte de Servicios"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_ReporteDocCancelados(aux As String, CodAux As String, Tipo As String, serie As String, correl As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
    FechaIni = IIf(FechaIni = "__/__/____", "", FechaIni)
    FechaFin = IIf(FechaFin = "__/__/____", "", FechaFin)
    
    SQL = "call CP_SP_CuentasPorPagar('REPORTE','" & aux & "','" & CodAux & "','" & FechaIni & "','" & FechaFin & "','" & Tipo & "','" & serie & "','" & correl & "','','');"
    Set oRs = ADO_LlenaRs(SQL)

    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No Existe Datos de Docuemntos Cancelados", vbOKOnly + vbInformation, gsNomSW
            Exit Sub
        End If
        EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, oRSMYSQL, oRs, True
    End If
    Set oRs = Nothing
End Sub
Public Sub sp_RepDocxCobrar(Tipo As String, NumFolio As String, Credito As String, serie As String, correl As String, CodAux As String, Div As String, Lote As String, pozo As String, cenco As String, fecha As String)
Dim SQL As String
Dim myrs As MYSQL_RS
    SQL = "Select DISTINCT a.Identificador,a.cod_tipo_doc,CONCAT(dc.Serie,'-',dc.Correl) as Documento," & _
          "trim(dc.Cenco) as cenco,dc.division,dc.Auxiliar,dc.codigo,dc.Fec_Emision," & _
          "dc.mon , dc.total, mov.cod_Estado,aux.credito,dc.Fec_Vcto,dc.Fec_Pago " & _
          " from (((((amarre_documento as a LEFT Join documento_contables as dc " & _
          " on (a.Identificador=dc.Identificador))LEFT Join cncosto as ct " & _
          " on (dc.cenco=ct.cenco))LEFT join cntablas as cn " & _
          " on (dc.auxiliar=cn.tip_linea))LEFT join movi_documento as mov " & _
          " on (dc.Identificador=mov.Identificador))left join cnauxil as aux" & _
          " on (dc.codigo=aux.codigo)) LEFT JOIN ALMACEN AS ALMA ON (dc.codalmac=alma.codigo)" & _
          " where alma.descrip not like ('%EXT%') AND a.Flag='0' and aux.auxiliar='2' and cn.codtab='1' and (mov.Cod_estado='" & IMPRESO & "' or " & _
          " mov.Cod_estado='" & EMITIDO & "') and a.cod_tipo_doc='01' "
    If fecha <> "__/__/____" Then
        SQL = SQL & "and dc.Fec_Pago='" & Format(fecha, "YYYY/MM/DD") & "'"
    End If
    
    If Tipo <> "" Then
        SQL = SQL & " and a.cod_tipo_doc='" & Tipo & "'"
    Else
        SQL = SQL & " and (a.cod_tipo_doc='01'or a.cod_tipo_doc='03')"
    End If
    
    If NumFolio <> "" Then SQL = SQL & " and a.Identificador Like '%" & NumFolio & "%'"
    If Credito <> "0" Then SQL = SQL & " and aux.credito='" & Credito & "' "
    If correl <> "" Then SQL = SQL & " and dc.Serie='" & serie & "' and dc.Correl='" & correl & "'"
    If CodAux <> "" Then SQL = SQL & " and dc.codigo='" & CodAux & "'"
    If Div <> "" Then SQL = SQL & " and dc.division='" & Div & "'"
    If Lote <> "" Then SQL = SQL & " and ct.Lote in (Select idlote from novperuvhse.lote where DescripcionCorta like '" & Lote & "')"
    If pozo <> "" Then SQL = SQL & " and ct.Pozo in (Select idPozo from novperuvhse.pozo where DescripcionCorta like'" & pozo & "') "
    If cenco <> "" Then SQL = SQL & " and trim(dc.Cenco)='" & cenco & "'"
    SQL = SQL & " ORDER by dc.codigo,Documento "
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Datos de Documentos por Cobrar", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_CuentasxCobrar(F As Date, Optional Div As String, Optional Cliente As String, Optional RES As Integer)
    Dim SQL As String, str As String, str2 As String, str3 As String
    Dim myrs As MYSQL_RS
    Dim myrs2 As MYSQL_RS
    Dim fecha As String
    If Div <> "" And Div <> "0006" Then
        str = " and dc.division='" & Div & "' "
    Else
        str = ""
    End If
    If Cliente <> "" Then
        str2 = " and dc.codigo='" & Cliente & "' "
    Else
        str2 = ""
    End If
    
    If RES = vbNo Then
        str3 = " and (a.cod_tipo_doc='03' or a.cod_tipo_doc='01') "
    Else
        str3 = " and (a.cod_tipo_doc='03' or a.cod_tipo_doc='01' or a.cod_tipo_doc='P') "
    End If
    SQL = "delete from tmpcuentasxcobrar"
    
    oConexionMYSQL.Execute SQL
    
    SQL = "insert into tmpcuentasxcobrar" & _
          " select  dc.Identificador,dc.Fec_emision,dc.Fec_Pago as Fec_Cobro,dc.Fec_Vcto as Fec_Recepcion," & _
          " a.cod_tipo_doc,CONCAT(dc.serie,'-',dc.correl)as Documento,dc.auxiliar," & _
          " dc.codigo,aux.descrip as descod,trim(dc.cenco) as cenco,dc.division as Divs,dv.descrip," & _
          " dc.mon,if(dc.mon = 'N',(dc.total - dc.total_ref  - " & rptIgv("dc.Fec_emision") & " * ifnull((Select sum(valor) from factura_proforma" & _
          " where prof=CONCAT(dc.serie,'-',dc.correl)),0))/(select tipo_ven from cncambio where" & _
          " fecha = date_format(dc.fec_emision,'%d/%m/%Y')),(dc.total - dc.total_ref  - " & rptIgv("dc.Fec_emision") & " * ifnull((Select sum(valor) from" & _
          " factura_proforma where prof=CONCAT(dc.serie,'-',dc.correl)),0))) as Total,'' as Fec_Actual, (select numdias from forma_pago where codigo=aux.credito ) as  dias_vcto    ,0 as Monto_Actual," & _
          " 0 as Monto_3060,0 as Monto_6090,0 as Monto_90120,0 AS Monto_120,mov.cod_Estado,0 as Rango,aux.fechaconsulta,aux.resultadoconsulta," & _
          "(select contacto from contactoscliente cc where cc.auxiliar = aux.auxiliar and cc.codigo = aux.codigo and item = 1) as contacto," & _
          "(select contacto from contactoscliente cc where cc.auxiliar = aux.auxiliar and cc.codigo = aux.codigo and item = 2) as contacto2" & _
          " from ((((documento_contables as dc LEFT JOIN cnmdepar as dv " & _
          " on (dc.division=dv.coddep))LEFT JOIN amarre_documento AS a " & _
          " on (dc.IDENTIFICADOR=a.IDENTIFICADOR))LEFT Join cnauxil as aux " & _
          " on (aux.codigo=dc.codigo and aux.auxiliar=dc.auxiliar))Left Join " & _
          " movi_documento as mov on (dc.Identificador=mov.Identificador)) LEFT JOIN ALMACEN AS ALMA ON (dc.codalmac=alma.codigo)" & _
          " where alma.descrip not like ('%EXT%') and a.Flag='0' and aux.afiliada <> 'S' and dc.auxiliar='2' " & str3 & _
          " and (mov.cod_Estado <>'" & CANCELADO & "' and mov.cod_Estado <> '" & ELIMINADO & "' and mov.cod_Estado <>'" & ANULADO & "' " & _
          " and mov.cod_Estado <>'" & REGISTRADO & "')" & str & str2 & " order by dc.Fec_emision"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False

    SQL = "select Identificador,convert(DATE_FORMAT(adddate(CONVERT(Fec_Recepcion,date),convert(dias_vcto,unsigned )),'%Y/%m/%e' ),char) as Fec_Cobro,Fec_Recepcion from tmpcuentasxcobrar"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    Do While Not (myrs.EOF)
        If myrs.Fields("Fec_Cobro") <> "" Then
            SQL = " update tmpcuentasxcobrar set Fec_Actual='" & CDbl(CDate(myrs.Fields("Fec_Cobro"))) & "'" & _
                  " Where Identificador='" & myrs.Fields("Identificador") & "'"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
        End If
        myrs.MoveNext
    Loop
    SQL = " update tmpcuentasxcobrar set monto_actual=total,rango=0 " & _
          " Where " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned )))>= 0 and " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned ))) < 30"
    oConexionMYSQL.Execute SQL
    SQL = " update tmpcuentasxcobrar set Monto_3060=total,rango=1 " & _
          " Where " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned )))>= 30 and " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned ))) < 60"
    oConexionMYSQL.Execute SQL
    SQL = " update tmpcuentasxcobrar set Monto_6090=total,rango=2 " & _
          " Where " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned )))>= 60 and " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned ))) < 90"
    oConexionMYSQL.Execute SQL
    SQL = " update tmpcuentasxcobrar set Monto_90120=total,rango=3 " & _
          " Where " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned )))>= 90 and " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned ))) < 120"
    oConexionMYSQL.Execute SQL
    SQL = " update tmpcuentasxcobrar set Monto_120=total,rango=4 " & _
          " Where " & _
          " datediff(convert('" & Format(F, "yyyy/mm/dd") & "',date), adddate(CONVERT(Fec_Recepcion,date ), convert(dias_vcto,unsigned )))>= 120 "
    oConexionMYSQL.Execute SQL
    
    SQL = "delete from tmpcuentasxcobrar where fec_cobro=''"
    oConexionMYSQL.Execute SQL
    
End Sub
Public Sub sp_CENCO_CtasxCob()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim rpt As Integer
    rpt = MsgBox("¿Desea incluir Proformas en el Reporte?", vbYesNo + vbQuestion, "aviso")
    If rpt = vbYes Then
        SQL = " select DIVISION,DES_DIV,sum(Total),sum(Monto_Actual)," & _
              " sum(Monto_3060),sum(Monto_6090),Sum(Monto_90120)," & _
              " Sum(Monto_120) From tmpcuentasxcobrar " & _
              " group by division"
    Else
        SQL = " select DIVISION,DES_DIV,sum(Total),sum(Monto_Actual)," & _
              " sum(Monto_3060),sum(Monto_6090),Sum(Monto_90120)," & _
              " Sum(Monto_120) From tmpcuentasxcobrar where Cod_Tipo_Doc<>'P' " & _
              " group by division"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Datos de Docuemntos por Cobrar", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_CtasCenco_CtasxCob(Division As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim rpt As Integer
    rpt = MsgBox("¿Desea incluir Proformas en el Reporte?", vbYesNo + vbQuestion, "aviso")
    If rpt = vbYes Then
        SQL = "select * from tmpcuentasxcobrar where Cod_Tipo_Doc<>'XX'"
    Else
        SQL = "select * from tmpcuentasxcobrar where Cod_Tipo_Doc='01' or Cod_Tipo_Doc='03'"
    End If
    If Division <> "" Then SQL = SQL & " and division='" & Division & "'"
    SQL = SQL & " Order by DIVISION,FEC_RECEPCION"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Documentos por Cobrar", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_FlujoEfectivo()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select  Viernes,Movimiento,descripcion,Comentarios,mon,sum(total)," & _
          "Movto,Pintar from tmpflujodeefectivo where movto = 1 group by Viernes,Codigo,division,Comentarios " & _
          "union all select Viernes,Movimiento,descripcion,Comentarios,mon,sum(total)," & _
          "Movto,Pintar from tmpflujodeefectivo where movto = 2 group by Viernes,Codigo,division,Comentarios,movimiento"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Datos de flujo de Efectivo poryectado al " & frmFlujoEfectivo.dtpFecha.Text, vbInformation, "Reporte Flujo de Efectivo"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_FlujoEfectivoxDivision(Div As String, RES As Integer, Optional fec As String)
    Dim SQL As String, str As String, str2 As String, str3 As String
    Dim Str5 As String
    Dim myrs As MYSQL_RS
    If Trim(Div) <> "" And Trim(Div) <> "0006" Then
        str = " and Z.division='" & Div & "'"
        str2 = ",DIVISION, DES_DIV,"
    Else
        str = ""
        str2 = ",'0000' AS DIVISION,'TODAS' AS DES_DIV,"
    End If
    If fec <> "" Then
         Str5 = " and date_format(DATE_ADD(concat(substring(fecvcto,7,4),substring(fecvcto,3,4),substring(fecvcto,1,2)),INTERVAL " & _
               "if(weekday(concat(substring(fecvcto,7,4),substring(fecvcto,3,4),substring(fecvcto,1,2)))=4,0,weekday(concat(substring(fecvcto,7,4),substring(fecvcto,3,4),substring(fecvcto,1,2))) - " & _
               "(4-weekday(concat(substring(fecvcto,7,4),substring(fecvcto,3,4),substring(fecvcto,1,2)))+1)) DAY),'%Y/%m/%d') <='" & fec & "'"
    End If

        SQL = "Select IDENTIFICADOR,FEC_EMISION,FEC_COBRO,FEC_RECEPCION,COD_TIPO_DOC," & _
              " DOCUMENTO,AUXILIAR,CODIGO,RTRIM(PROVEEDOR),(select (select descrip from cnmdepar where coddep=c.Division) as nomdiv from documento_contables as c where c.identificador=Z.identificador) as cenco " & str2 & " MON," & _
              " TOTAL,Fec_Actual,Dias_Vcto,Monto_Actual,Monto_3060,Monto_6090," & _
              " Monto_90120,Monto_120,Estado,Rango,date_format(fecha,'%d/%m/%Y') as fecha,resultado,contacto,contacto2 from tmpcuentasxcobrar AS Z where Z.TOTAL >0 AND Z.Cod_Tipo_Doc<>'XX'" & str & str3

    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Datos de flujo de Efectivo para esta división", vbInformation, "Reporte Flujo de Efectivo"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_CobranzasVencidas(Div As String, Optional fec As String)
    Dim SQL As String, str As String, str2 As String
    Dim myrs As MYSQL_RS
    If Trim(Div) <> "" Then
        str = " and Z.division='" & Div & "'"
    Else
        str = ""
    End If
    If Trim(fec) <> "" Then
        str = " and Z.fec_cobro<='" & fec & "'"
    Else
        str = ""
    End If
    SQL = "select * from tmpcuentasxcobrar AS Z where Z.Cod_Tipo_Doc='03' or Z.Cod_Tipo_Doc='01'" & str & str2
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No existen cobranzas vencidas para esta división", vbInformation, "NOVADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_FlujoEfectivoxDivision_Resumen(Div As String, RES As Integer, Optional fec As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    If RES = vbYes Then
        SQL = "select * from tmpcuentasxcobrar where Cod_Tipo_Doc<>'XX' "
    Else
        SQL = "select * from tmpcuentasxcobrar where Cod_Tipo_Doc='01' or Cod_Tipo_Doc='03' "
    End If
    If Trim(Div) <> "" And Trim(Div) <> "0006" Then SQL = SQL & " and division='" & Div & "'"
    If Trim(fec) <> "" Then SQL = SQL & " and fec_cobro<>'' and fec_cobro<='" & fec & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Datos de flujo de Efectivo para esta división", vbInformation, "Reporte Flujo de Efectivo"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_LiquidCobranza(FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select" & _
          " a.identificador,a.fechacobro,a.cod_tipo_doc,a.ndocto,a.auxiliar,a.codigo,a.proveedor," & _
          " a.estado,a.total,b.cod as cod_tipo_ref,b.descrip as tipo_ref,b.numdoc as ndocto_ref, (b.importe)*-1 as total_ref,a.saldo,a.afecto," & _
          " a.Descripcion " & _
          " from  tmpLiquidaciondeCobranza as a LEFT join detalledctos as b on (a.identificador=b.identificador) " & _
          " where "
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " a.FechaCobro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and a.FechaCobro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " a.FechaCobro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and a.FechaCobro<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " a.FechaCobro>='" & Format(FechaFin, "yyyy/mm/dd") & "' and a.FechaCobro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " Left(a.FechaCobro,7)='" & strAnoSistema & "/" & strMesSistema & "'"
            End If
        End If
    End If
    SQL = SQL & " and seleccion = 'S' order by a.identificador"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Datos en Liquidación de Cobranzas ", vbInformation, "Reporte Liquidacion de Cobranzas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_OrdenFactura(ORD As String, Optional FecIni As String, Optional FecFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Delete from tmpOrdenesfacturadas "
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    SQL = "insert into tmpOrdenesfacturadas " & _
          "Select o.correl as N_Orden,o.Fec_Emision as Fecha,o.mon as monOrden," & _
          "o.Total as Importe,a.Codigo,a.Descrip as Nombre,trim(o.Cenco) as cenco,am.Cod_Tipo_Doc as Tipo," & _
          "dc.Identificador,CONCAT(dc.Serie,'-',dc.Correl) as Documento,dc.Fec_Emision ," & _
          "dc.mon, (dc.Total) as Total,o.Otros_Montos as Otros_Montos, dc.Guia,'0' as Cont,0.00 as SumOrden,O.Tunidades AS tu,O.URecibidas AS ur,mov.cod_estado    " & _
          " From (((orden_compra as o Inner JOIN documento_contables as dc " & _
          " on (o.Correl=dc.ORDEN)) LEFT JOIN cnauxil as a " & _
          " on (dc.auxiliar=a.auxiliar and dc.Codigo=a.Codigo)) " & _
          " left JOIN amarre_documento as am on (dc.Identificador=am.identificador)) " & _
          " LEFT JOIN movi_documento AS mov on(dc.Identificador=mov.identificador) " & _
          " Where am.Cod_Tipo_Doc IN (SELECT TIPO_DOC AS Cod_Tipo_Doc FROM configuracion_doc WHERE NOM_CAMPO='ORDEN' AND VISIBLE='1') " & _
          "        and mov.cod_estado<>'" & ANULADO & "' and mov.cod_estado<>'" & ELIMINADO & "'" & IIf(ORD = "", "", " and o.correl='" & ORD & "'")
    If IsDate(FecIni) And IsDate(FecFin) Then
        SQL = SQL & " and  o.fec_emision>='" & FecIni & "' and o.fec_emision<='" & FecFin & "'"
    End If
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
    oConexionMYSQL.Execute "Update tmpOrdenesfacturadas set total=total*-1 where tipodoc='07'"
    SQL = "select N_orden,count(N_Orden)as Cont,sum(Total) as SumOrden from tmpOrdenesfacturadas group by N_Orden"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    Do While Not (myrs.EOF)
        SQL = " Update tmpOrdenesfacturadas set Cont='" & myrs.Fields("Cont") & "'," & _
              " SumOrden='" & myrs.Fields("SumOrden") & "' where N_Orden='" & myrs.Fields("N_Orden") & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
        myrs.MoveNext
    Loop
End Sub
Public Sub sp_OrdenesFacturadas(orden As String, Cont As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim myrs2 As MYSQL_RS
    SQL = "select  a.N_Orden,a.Fecha,a.Mon_Orden,a.Importe,a.Codigo,a.Nombre," & _
          " trim(a.Cenco) as cenco,a.TipoDoc,a.Identificador,a.Documento,a.Fec_Emision,a.Mon, " & _
          " a.Total,a.Otros_Montos,a.Guia,a.Cont,a.SumOrden,A.TIPODOC as tipo,A.MON as serie,A.MON as numero, " & _
          " A.MON as fecr,a.ur as unidreci,A.MON as flaganu,A.MON as docint,a.tu,a.ur,a.estadodoc " & _
          " from tmpordenesfacturadas  as a where a.n_orden <>'' "
    If orden <> Empty Then SQL = SQL & " and a.N_orden ='" & orden & "'"
    If Cont <> Empty Then SQL = SQL & " and a.cont >='" & Cont & "' "
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    SQL = " select  a.correl as O_orden,a.fec_emision as Fecha,a.mon as Mon_Orden, " & _
         " a.total as Importe,a.Codigo,a.obs as Nombre, " & _
         " trim(a.Cenco) as cenco, a.visado as TipoDoc,a.Identificador, a.tipo as Documento, " & _
         " a.usuario as Fec_Emision,a.estado as Mon, " & _
         " a.valorcompra as Total,a.Otros_Montos,a.banco as Guia, " & _
         " a.credito as Cont,a.igv as SumOrden,b.tipo,b.serie,b.numero, " & _
         " b.fecha as frec,b.unidreci,b.flaganu,b.docint,a.tunidades,a.urecibidas " & _
         " from orden_compra as a " & _
         " left join orden_rec as B " & _
         " on (B.ORDEN=A.correl) where b.flaganu<>'S' "
    If orden <> Empty Then SQL = SQL & " and b.orden ='" & orden & "'"
    Set myrs2 = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Facturas Relacionada con esta Orden de Compras ", vbInformation, "Reporte Ordenes Facturadas"
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount <> 0 And myrs2.RecordCount <> 0 Then
        ImprimeReporteGeneral myrs, myrs2, App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, m_Titulo, "", m_Empresa
    Else
        m_Reporte = "Rep_OrdenesFacturadasSG.rpt"
        EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
    Set myrs2 = Nothing
End Sub
Public Sub sp_SaldoOrdenesFacturadas(fIni As String, fFin As String)
    Dim SQL As String, par As String, str As String
    Dim myrs As MYSQL_RS
    sp_OrdenFactura "", fIni, fFin
    par = InputBox("[A]biertas - [C]erraras - [N]egativas", "Ordenes", "A")
    Select Case par
        Case "A": str = "importe>SumOrden"
        Case "C": str = "importe=SumOrden"
        Case "N": str = "importe<SumOrden"
    End Select
    SQL = "select * from tmpOrdenesfacturadas where " & str
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Registros de Ordenes con Saldo ", vbInformation, "Reporte Ordenes Facturadas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_DocumentosSinSustento()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "SELECT LEFT(A.IDENTIFICADOR,6) AS ANOMES,RIGHT(A.IDENTIFICADOR,4) AS FOLIO," & _
          " A.FECHA_REGISTRO,B.ORDEN,A.COD_TIPO_DOC," & _
          " CONCAT(B.SERIE,'-',B.CORREL) AS DOCUMENTO,D.COD_ESTADO," & _
          " if(k.coddoc<>'',c.DESCRIP,'') as descrip,B.mon,if(k.coddoc<>'',B.total,0) as total " & _
          " FROM ((amarre_documento AS A LEFT JOIN documento_contables AS B " & _
          " ON (A.IDENTIFICADOR=B.IDENTIFICADOR)) LEFT JOIN cnauxil AS C " & _
          " ON(B.AUXILIAR=C.AUXILIAR AND B.CODIGO=C.CODIGO)) LEFT JOIN movi_documento AS D " & _
          " ON(A.IDENTIFICADOR=D.IDENTIFICADOR) " & _
          " left join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N' OR " & _
          " (SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = 'AUGUSTO')=1)) K" & _
          " on (a.Cod_Tipo_Doc=k.coddoc) " & _
          " WHERE A.OBS LIKE '%SUSTENTO%' AND D.COD_ESTADO<>'EL' " & _
          " ORDER BY A.IDENTIFICADOR"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Registros de Documentos sin Sustento", vbInformation, "Reporte Documentos sin Sustento"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_DocumentosSinOrden()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "SELECT LEFT(A.IDENTIFICADOR,6) AS ANOMES,RIGHT(A.IDENTIFICADOR,4) AS FOLIO," & _
          " A.FECHA_REGISTRO,B.ORDEN,A.COD_TIPO_DOC," & _
          " CONCAT(B.SERIE,'-',B.CORREL) AS DOCUMENTO,D.COD_ESTADO," & _
          " if(k.coddoc<>'',c.DESCRIP,'') as descrip,B.mon,if(k.coddoc<>'',B.total,0) as total " & _
          " FROM ((((amarre_documento AS A LEFT JOIN documento_contables AS B " & _
          " ON (A.IDENTIFICADOR=B.IDENTIFICADOR)) LEFT JOIN cnauxil AS C " & _
          " ON(B.AUXILIAR=C.AUXILIAR AND B.CODIGO=C.CODIGO)) LEFT JOIN movi_documento AS D " & _
          " ON(A.IDENTIFICADOR=D.IDENTIFICADOR)) LEFT JOIN configuracion_doc as cd on (cd.tipo_doc=a.cod_tipo_doc)) " & _
          " left join (SELECT coddoc FROM cndocum C WHERE (c.PROTEGIDO = 'N' OR " & _
          " (SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K" & _
          " on (a.Cod_Tipo_Doc=k.coddoc) " & _
          " WHERE cd.nom_campo='ORDEN' and cd.visible='1' and A.FLAG='1' AND left(a.identificador,6)='" & strAnoSistema & strMesSistema & "' and B.ORDEN='' AND D.COD_ESTADO<>'EL' " & _
          " ORDER BY A.IDENTIFICADOR"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Registros de Documentos sin Orden", vbInformation, "Reporte Documentos sin Orden"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_OrdenesImporteCero()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = " select a.fec_emision, a.codigo,a.correl,a.tipo,a.tipo_cotizacion," & _
          " a.guia, trim(a.cenco) as cenco, a.mon, a.total, b.descrip " & _
          " from orden_compra  as a left join cnauxil  as b " & _
          " on (a.auxiliar=b.auxiliar and a.codigo=b.codigo) where a.total =0 order by a.correl"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Ordenes Facturas con Importe Cero ", vbInformation, "Reporte Ordenes Facturadas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_TmpRegistroFacturacion(Optional par As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim myrs2 As MYSQL_RS
    Dim Descripcion As String
    Dim I As Integer
    Descripcion = ""
    I = 1
    SQL = "delete from tmpregistrosfacturacion where usuario = '" & strUsuarioId & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    If par = "" Then
        SQL = "Insert into tmpregistrosfacturacion(identificador,fecha,codtipo,ndoc,division,desdiv,auxiliar,codaux,nombre,estado,codser,tiposer,descripcion,total,clasificacion,lote,mon,tc,orden,ref,usuario) " & _
              " Select dc.Identificador,dc.Fec_Emision,am.cod_Tipo_Doc," & _
              " Concat(dc.Serie,'-',dc.Correl) as Documento,dc.Division," & _
              " dv.descrip as Des_Div,dc.auxiliar,dc.codigo,aux.Descrip as Nombre," & _
              " mov.cod_Estado,det.codser,det.descripcion as TipoSer," & _
              " cls.descrip as Descripcion,dc.SubTotal,cls.codigo as Clasf,(select DESCRIP from almacen where codigo=dc.codAlmac) as lote," & _
              " dc.mon,(select tipo_ven from cncambio where right(fecha,4)=left(dc.fec_emision,4) and left(fecha,2)=right(dc.fec_emision,2) and mid(fecha,4,2)=mid(dc.fec_emision,6,2)  ) as tc,dc.orden,dc.ref,'" & strUsuarioId & "' " & _
              " from ((((((documento_contables as dc left JOIN amarre_documento as am " & _
              " on (dc.Identificador=am.identificador)) LEFT JOIN cnauxil as aux " & _
              " on (dc.auxiliar=aux.auxiliar and dc.Codigo=aux.Codigo)) " & _
              " left join movi_documento as mov " & _
              " on(dc.Identificador=mov.Identificador))LEFT JOIN cnmdepar as dv " & _
              " on (dc.division=dv.coddep)) LEFT JOIN detallefact as det " & _
              " on (dc.Identificador=det.identificador)) LEFT JOIN servicio as ser " & _
              " on(det.codser=ser.codigo)) LEFT JOIN clasif_serv as cls " & _
              " on(ser.codclasf=cls.codigo) " & _
              " where mov.cod_estado <> '" & ELIMINADO & "' and am.flag='0' and" & _
              " (am.cod_tipo_doc='01' or am.cod_tipo_doc='P' or am.cod_tipo_doc='07' or am.cod_tipo_doc='03') " & _
              " and substring(dc.fec_emision,6,2)>='" & strMesSistema & "'" & _
              " and substring(dc.fec_emision,6,2)<='" & strMesSistema & "'" & _
              " and left(dc.fec_emision,4)='" & strAnoSistema & "'" & _
              " GROUP by dc.identificador,cls.codigo"
    Else
        SQL = "Insert into tmpregistrosfacturacion(identificador,fecha,codtipo,ndoc,division,desdiv,auxiliar,codaux,nombre,estado,codser,tiposer,descripcion,total,clasificacion,lote,mon,tc,orden,ref,usuario) " & _
              " Select dc.Identificador,dc.Fec_Emision,am.cod_Tipo_Doc," & _
              " Concat(dc.Serie,'-',dc.Correl) as Documento,dc.Division," & _
              " dv.descrip as Des_Div,dc.auxiliar,dc.codigo,aux.Descrip as Nombre," & _
              " mov.cod_Estado,det.codser,det.descripcion as TipoSer," & _
              " cls.descrip as Descripcion,dc.SubTotal,cls.codigo as Clasf,(select DESCRIP from almacen where codigo=dc.codAlmac) as lote," & _
              " dc.mon,dc.otros_montos as tc,dc.orden,dc.ref,'" & strUsuarioId & "' " & _
              " from ((((((documento_contables as dc left JOIN amarre_documento as am " & _
              " on (dc.Identificador=am.identificador))LEFT JOIN cnauxil as aux " & _
              " on (dc.auxiliar=aux.auxiliar and dc.Codigo=aux.Codigo)) " & _
              " left join movi_documento as mov " & _
              " on(dc.Identificador=mov.Identificador))LEFT JOIN cnmdepar as dv " & _
              " on (dc.division=dv.coddep))LEFT JOIN detallefact as det " & _
              " on (dc.Identificador=det.identificador))LEFT JOIN servicio as ser " & _
              " on(det.codser=ser.codigo))LEFT JOIN clasif_serv as cls " & _
              " on(ser.codclasf=cls.codigo) " & _
              " where mov.cod_estado <> '" & ELIMINADO & "' and am.flag='0' and am.cod_tipo_doc='P' " & _
              " GROUP by dc.identificador,cls.codigo"
    End If
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
    SQL = "SELECT Identificador,count(*) from tmpregistrosfacturacion " & _
          " WHERE ESTADO<>'" & ANULADO & "' and usuario = '" & strUsuarioId & "' group by identificador having count(*)>1"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    Do While Not (myrs.EOF)
        SQL = "select descripcion from tmpregistrosfacturacion where Identificador='" & myrs.Fields("Identificador") & "' and usuario = '" & strUsuarioId & "'"
        Set myrs2 = oConexion.EjecutaSelectRS(SQL)
        Descripcion = myrs2.Fields(0)
        myrs2.MoveNext
        For I = 1 To myrs2.RecordCount - 1
            If Descripcion = CE(myrs2.Fields("Descripcion")) <> "" Then
                Descripcion = Trim(Replace(Descripcion & "-" & CE(myrs2.Fields("Descripcion")), "SERVICIO", ""))
            End If
            myrs2.MoveNext
        Next
        SQL = "Update tmpregistrosfacturacion set descripcion='" & Descripcion & "' " & _
              "where Identificador='" & myrs.Fields("Identificador") & "' and usuario = '" & strUsuarioId & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
        Descripcion = ""
        myrs.MoveNext
    Loop
    Set myrs = Nothing
    Set myrs2 = Nothing
End Sub
Public Sub sp_RegistroFacturacionCorte()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select *  from tmpregistrosfacturacion " & _
          " where substring(Fecha,6,2)>='" & strMesSistema & "'" & _
          " and substring(Fecha,6,2)<='" & strMesSistema & "' and left(fecha,4)='" & strAnoSistema & "' and usuario = '" & strUsuarioId & "'"
    SQL = SQL & " group by Identificador "
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Registro de Facturación ", vbInformation, "Reporte Ordenes Facturadas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_RegistroFacturacionCierre()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select *  from tmpregistrosfacturacion where (CodTipo='01' OR CodTipo='07')" & _
         " and substring(Fecha,6,2)>='" & strMesSistema & "'" & _
         " and substring(Fecha,6,2)<='" & strMesSistema & "' and left(fecha,4)='" & strAnoSistema & "' and usuario = '" & strUsuarioId & "'"
    SQL = SQL & " group by Identificador "
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Registro de Facturación ", vbInformation, "Reporte Ordenes Facturadas"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_RegistroFacturacionProformas(Div As String, Cli As String, prof As String, mi As String, mf As String, mfac As String)
    Dim SQL As String, STRDiv As String, strcli As String, strprof As String, strMes As String, strmesfac As String, rpta As Integer, rpta2 As Integer, str2 As String, rptaF As Integer, strF As String, rptaOanda As String
    Dim myrs As MYSQL_RS
    Dim TCambioCierre As String
    
    If Div <> "" And Div <> "0006" Then
        'STRDiv = " and a.division='" & Div & "' "
        STRDiv = "AND IF((select ATIPO from cnmdepar where coddep='" & Div & "')='TUBULAR SERVICES', a.division in (select coddep from cnmdepar where ATIPO='TUBULAR SERVICES'), a.division='" & Div & "')"
    Else
        STRDiv = ""
    End If
    
    If Cli <> "" Then
        strcli = " and a.CodAux='" & Cli & "' "
    Else
        strcli = ""
    End If
    
    If prof <> "" Then
        prof = Right("00000" & Trim(Left(prof, InStr(1, prof, "-", vbBinaryCompare) - 1)), 5) & "-" & Right("000000000" & Trim(Right(prof, Len(prof) - InStr(1, prof, "-", vbBinaryCompare))), 9)
        strprof = " and a.NDoc='" & prof & "' "
    Else
        strprof = ""
    End If
    
    If mfac <> "" Then
        strmesfac = " and (substring(c.fec_emision,6,2)<='" & mfac & "' or c.fec_emision is null )   "
    Else
        strmesfac = ""
    End If
    
    If mi = "" And mf = "" Then
        strMes = " and LEFT(a.Fecha,7)<='" & mi & "'"
    Else
        If mi <> "" And mf = "" Then
            strMes = " and LEFT(a.Fecha,7)='" & mi & "'"
        Else
            If mi = "" And mf <> "" Then
                strMes = " and LEFT(a.Fecha,7)='" & mf & "'"
            Else
                If mi <> "" And mf <> "" Then
                    strMes = " and LEFT(a.Fecha,7)>='" & mi & "' and LEFT(a.Fecha,7)<='" & mf & "'"
                End If
            End If
        End If
    End If
    Set myrs = oConexion.EjecutaSelectRS("call CONT_InsertaTempProfFact ('" & strUsuarioId & "','" & mi & "','" & mf & "');")

    'and a.descripcion<>'REEMBOLSO'
    rpta2 = MsgBox("¿Desea incluir proformas por reembolso de gastos?", vbYesNo + vbQuestion, "NOVADMIN")
    If rpta2 = vbYes Then
        str2 = ""
    Else
        str2 = " and a.clasificacion<>'10' "
    End If
    
    rptaF = MsgBox("¿Desea visualizar las proformas con toda la Facturación?", vbYesNo + vbQuestion, "NOVADMIN")
    If rptaF = vbYes Then
        strF = "X"
    Else
        strF = ""
    End If
    TCambioCierre = InputBox("Ingrese el Tipo de Cambio de Cierre del Mes.", "Tipo de Cambio", 1)
    
    rpta = MsgBox("¿Desea generar solo las proformas con saldo?", vbYesNo + vbQuestion, "NOVADMIN")
    If rpta = vbYes Then
        SQL = "Select a.Identificador,a.Fecha,if(left(c.fec_emision,7) >'" & strF & mf & "',null,c.fec_emision) as fec_emision,a.CodTipo,if(left(c.fec_emision,7) >'" & strF & mf & "',null,(select cod_tipo_doc from amarre_documento where identificador=c.identificador)) " & _
              " as COD_TIPO_DOC,a.NDoc,a.Division," & _
              " a.DesDiv,a.Auxiliar, a.CodAux,ifnull((select descrip from cnauxil where auxiliar='2' and codigo=c.codigo),a.Nombre) as Nombre,a.Estado,a.Codser,a.TipoSer," & _
              " a.ref as Descripcion,a.Total,round((select IMPORTE from cncambiomes_ where tipo='OAN' and aniomes='" & strAnoSistema & strMesSistema & "'),6) as TotOanda,1.18*a.Total as TotalIGV, a.Clasificacion,a.lote,a.mon,a.tc,a.orden," & _
              " a.ref,if(left(c.fec_emision,7) >'" & strF & mf & "',null,concat(c.serie,'-',c.correl)) as fac,if(left(c.fec_emision,7) >'" & strF & mf & "',null,(c.subtotal * (select factor from cndocum where coddoc=(select cod_tipo_doc from amarre_documento where identificador=c.identificador)))) AS SubTotal, " & _
              " if(left(c.fec_emision,7) >'" & strF & mf & "',null,(select descrip from almacen where codigo=C.codalmac)) as almacen,if(left(c.fec_emision,7) >'" & strF & mf & "',null,((select sum(valor) from factura_proforma f where f.identificador=b.identificador AND F.PROF = A.NDOC) * (select factor from cndocum where coddoc=(select cod_tipo_doc from amarre_documento where identificador=c.identificador)))) as valor,if(a.descripcion<>'REEMBOLSO',' NO REEMBOLSO',a.descripcion) AS descriptipo, " & _
              "(select ifnull(PERSONAENC,'-') from segfactprof as s where s.IDENTIFICADOR=a.Identificador order by s.item desc limit 1) as personaenc, " & _
              "(select ifnull(SITUACIONPROF,'-') from segfactprof as s where s.IDENTIFICADOR=a.Identificador order by s.item desc limit 1)  as situacionprof," & _
              "(select ifnull(FECHAFACT,'-') from segfactprof as s where s.IDENTIFICADOR=a.Identificador order by s.item desc limit 1) as fechacierre," & TCambioCierre & " as TipoCambioCierre" & _
              " from  ((tmpregistrosfacturacion as a left join factura_proforma as b" & _
              " on a.ndoc=b.prof ) LEFT JOIN documento_contables AS C" & _
              " on (b.identificador=c.identificador))  RIGHT JOIN " & _
              " (select concat(b.serie,'-',b.correl) AS PROF,(b.subtotal - SUM(a.valor)) as saldo from ((documento_contables as B left join movi_documento as c" & _
              " on (b.identificador=c.identificador)) left join amarre_documento as d on (b.identificador=d.identificador)) left join factura_proforma as a on (a.prof=concat(b.serie,'-',b.correl))" & _
              " where c.cod_estado<>'AN' and c.cod_estado<>'EL' and d.cod_tipo_doc='P'" & _
              " group by concat(b.serie,'-',b.correl) having (saldo)>0 or (saldo) is null order by concat(b.serie,'-',b.correl)) " & _
              " AS ELM ON (A.NDOC=ELM.PROF) " & _
              " where a.usuario = '" & strUsuarioId & "' and a.total>0 AND a.CodTipo='P' and a.estado<>'AN'   " & STRDiv & strcli & strprof & strMes & strmesfac & str2 & _
              " group by descriptipo,c.fec_emision,a.ndoc,c.serie,c.correl"
    Else
        SQL = "Select a.Identificador,a.Fecha,if(left(c.fec_emision,7) >'" & strF & mf & "',null,c.fec_emision) as fec_emision,a.CodTipo,if(left(c.fec_emision,7) >'" & strF & mf & "',null,(select cod_tipo_doc from amarre_documento where identificador=c.identificador)) " & _
              " as COD_TIPO_DOC,a.NDoc,a.Division," & _
              " a.DesDiv,a.Auxiliar, a.CodAux, ifnull((select descrip from cnauxil where auxiliar='2' and codigo=c.codigo),a.Nombre) as Nombre,a.Estado,a.Codser,a.TipoSer," & _
              " a.ref as Descripcion,a.Total,Round((select IMPORTE from cncambiomes_ where tipo='OAN' and aniomes='" & strAnoSistema & strMesSistema & "'),6) as TotOanda,1.18*a.Total as TotalIGV, a.Clasificacion,a.lote,a.mon,a.tc,a.orden," & _
              " a.ref,if(left(c.fec_emision,7) >'" & strF & mf & "',null,concat(c.serie,'-',c.correl)) as fac,if(left(c.fec_emision,7) >'" & strF & mf & "',null,(c.subtotal * (select factor from cndocum where coddoc=(select cod_tipo_doc from amarre_documento where identificador=c.identificador)))) AS SubTotal, " & _
              " if(left(c.fec_emision,7) >'" & strF & mf & "',null,(select descrip from almacen where codigo=C.codalmac)) as almacen,if(left(c.fec_emision,7) >'" & strF & mf & "',null,((select sum(valor) from factura_proforma f where f.identificador=b.identificador AND F.PROF = A.NDOC) * (select factor from cndocum where coddoc=(select cod_tipo_doc from amarre_documento where identificador=c.identificador)))) as valor,if(a.descripcion<>'REEMBOLSO',' NO REEMBOLSO',a.descripcion) AS descriptipo," & _
              "(select ifnull(PERSONAENC,'-') from segfactprof as s where s.IDENTIFICADOR=a.Identificador order by s.item desc limit 1) as personaenc, " & _
              "(select ifnull(SITUACIONPROF,'-') from segfactprof as s where s.IDENTIFICADOR=a.Identificador order by s.item desc limit 1)  as situacionprof," & _
              "(select ifnull(FECHAFACT,'-') from segfactprof as s where s.IDENTIFICADOR=a.Identificador order by s.item desc limit 1) as fechacierre," & TCambioCierre & " as TipoCambioCierre" & _
              " from (tmpregistrosfacturacion as a left join factura_proforma as b" & _
              " on a.ndoc=b.prof ) LEFT JOIN documento_contables AS C" & _
              " on (b.identificador=c.identificador)" & _
              " where a.usuario = '" & strUsuarioId & "' AND a.CodTipo='P' and a.estado<>'AN'   " & STRDiv & strcli & strprof & strMes & strmesfac & str2 & _
              " group by descriptipo,c.fec_emision,a.ndoc,c.serie,c.correl"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Registro de Proformas ", vbInformation, "Reporte Ordenes Facturadas"
        Set myrs = Nothing
        Exit Sub
    End If
    rptaOanda = MsgBox("¿Desea generar el Reporte Con el tipo de Cambio Oanda?", vbYesNo + vbQuestion, "NOVADMIN")
    If rptaOanda = vbYes Then
        m_Reporte = "Rep_Proformas.rpt"
    Else
        m_Reporte = "Rep_Proformas_sinOanda.rpt"
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
   
     
End Sub
Public Sub sp_RegistroFacturacionProformasxMes(Div As String, Cli As String, prof As String, mi As String, mf As String, mfac As String)
    Dim SQL As String
      
    Dim oRSMYSQL As MYSQL_RS
    Dim oRs  As ADODB.Recordset
    Set oRs = ADO_LlenaRs("call FAC_SP_Proformas ('REP_PROFXMES','" & Div & "','" & Cli & "','" & prof & "','" & mi & "','" & mf & "','" & mfac & "');")
    
    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No Existe Registro de Proformas ", vbInformation, "Reporte Ordenes Facturadas"
            Set oRs = Nothing
            Exit Sub
        End If
        
        m_Cab1 = strAnoSistema & "/" & mfac
        m_Reporte = "Rep_ProformasxMes_New.rpt"
        EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, oRSMYSQL, oRs, False
        
    End If
    
    Set oRs = Nothing
    Screen.MousePointer = vbNormal
End Sub
Public Sub sp_Pagos_Descuentos()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select dc.Identificador,ddoc.Fecha as Fecha,a.cod_tipo_doc,'FACTURA'," & _
          "CONCAT(dc.serie,'-',dc.Correl) as Documento,dc.auxiliar,dc.codigo," & _
          "aux.descrip as Proveedor,mov.Cod_Estado,dc.Total," & _
          "ddoc.Cod as CodTipoRef,cn.descrip as Tipo_Ref,ddoc.Numdoc as NDocto_Ref," & _
          "ddoc.Importe as Total_Ref,(dc.Total-ddoc.Importe) as Saldo " & _
          " from ((((amarre_documento as a LEFT Join documento_contables as dc " & _
          " on (a.Identificador=dc.Identificador))Left Join movi_documento as mov " & _
          " on (mov.Identificador=dc.Identificador))Left Join detalledctos as ddoc " & _
          " on (dc.Identificador=ddoc.Identificador))LEFT Join cndocum as cn " & _
          " on(dc.Cod_Tipo_Ref=cn.CodDoc))left join cnauxil as aux " & _
          " on(dc.auxiliar=aux.auxiliar and dc.codigo=aux.codigo) " & _
          " where dc.auxiliar='2' and a.cod_tipo_doc='01' and mov.Cod_Estado <>'" & CANCELADO & "' " & _
          " and dc.Total_Ref<> 0"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Registro de Pagos de Descuentos ", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_BonoCampo()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select * from tmpbonoscampo order by nombre"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_BonoResumen(val As Double)
    Dim SQL As String, str As String
    Dim myrs As MYSQL_RS
    If val > 0 Then
        str = " where monto_bono= " & val & " "
    Else
        str = ""
    End If
    SQL = "SELECT ANIOMES,UCASE(NOMBRE) AS NOMBRE,CODEMP,MONTO_BONO," & _
          " sum(IF(d1<>'',1, 0 ))+sum(IF(d2<>'',1, 0 ))+ " & _
          " sum(IF(d3<>'',1, 0 ))+sum(IF(d4<>'',1, 0 ))+ " & _
          " sum(IF(d5<>'',1, 0 ))+sum(IF(d6<>'',1, 0 ))+ " & _
          " sum(IF(d7<>'',1, 0 ))+sum(IF(d8<>'',1, 0 ))+ " & _
          " sum(IF(d9<>'',1, 0 ))+sum(IF(d10<>'',1, 0 ))+ " & _
          " sum(IF(d11<>'',1, 0 ))+sum(IF(d12<>'',1, 0 ))+ " & _
          " sum(IF(d13<>'',1, 0 ))+sum(IF(d14<>'',1, 0 ))+ " & _
          " sum(IF(d15<>'',1, 0 ))+sum(IF(d16<>'',1, 0 ))+ " & _
          " sum(IF(d17<>'',1, 0 ))+sum(IF(d18<>'',1, 0 ))+ " & _
          " sum(IF(d19<>'',1, 0 ))+sum(IF(d20<>'',1, 0 ))+ " & _
          " sum(IF(d21<>'',1, 0 ))+sum(IF(d22<>'',1, 0 ))+ " & _
          " sum(IF(d23<>'',1, 0 ))+sum(IF(d24<>'',1, 0 ))+ " & _
          " sum(IF(d25<>'',1, 0 ))+sum(IF(d26<>'',1, 0 ))+ " & _
          " sum(IF(d27<>'',1, 0 ))+sum(IF(d28<>'',1, 0 ))+ " & _
          " sum(IF(d29<>'',1, 0 ))+sum(IF(d30<>'',1, 0 ))+ " & _
          " sum(IF(d31<>'',1, 0 )) as dias, (select descrip from cnmdepar where coddep=division) as division," & _
          " (select descrip from cnauxil where codigo=pozo and auxiliar='5') as pozo FROM tmpbonoscampo " & str & _
          " GROUP BY NOMBRE,CODEMP,MONTO_BONO"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_RRHH()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "Select * from tmprptdinamico "
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Tareo(AnoMes As String, plani As String, mon As String, emp As String)
    Dim SQL As String, str As String
    If emp <> "" Then
        str = " and a.emp='" & emp & "'"
    Else
        str = ""
    End If
    Dim myrs As MYSQL_RS
    SQL = "SELECT A.EMP, A.SBASICO,CONCAT(B.APEPAT,' ',B.APEMAT,' ',B.NOMBRE1,' ',B.NOMBRE2) AS NOMBRE,A.RUB," & _
          " R.DESCRIP,R.TIPO AS TRUB,TR.DESCRIP AS DTRUB,TR.TIPO AS TIP ,A.UNIDAD,A.CANT,(SELECT NOMBRE FROM AFP WHERE CODIGO=A.AFP) AS DESAFP " & _
          " FROM ((PL_TAREO  AS A LEFT JOIN EMPLEADO AS B" & _
          " ON (A.EMP=B.CODIGO)) LEFT JOIN pl_rubrosremunerativos AS R ON (A.RUB=R.CODIGO))" & _
          " LEFT JOIN pl_tiporubros AS TR ON (TR.CODIGO=R.TIPO)" & _
          " WHERE A.ANOMES='" & AnoMes & "' AND A.TIPO='" & plani & "' AND A.MONEDA='" & mon & "'" & str & _
          " ORDER BY TR.DESCRIP,R.DESCRIP"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    End If
End Sub

Public Sub sp_Rep_RtaQta(AnoMes As String, plani As String, mon As String, emp As String, Optional opcMostrar As Boolean)
    Dim SQL As String, str As String, Str1 As String
    Dim strTipo As String
    If emp <> "" And emp <> "00000000000" Then
        str = " and emp='" & emp & "'"
    Else
        str = ""
    End If
    If plani = "4" Then
        Str1 = " and emp in (Select emp from pl_tareo where anomes='" & AnoMes & "' and tipo='4' and moneda='" & mon & "')"
    End If
    If mon = "C" Or mon = "E" Then
        strTipo = " and c.tipoplani = 'S' "
    ElseIf mon = "N" Then
        strTipo = " and c.tipoplani = 'N' "
    Else
        If opcMostrar = True Then
            strTipo = ""
        Else
            strTipo = " and c.tipoplani = 'N' "
        End If
    End If
    Dim myrs As MYSQL_RS
    SQL = "SELECT (select CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) from" & _
          " empleado e where e.codigo=p.emp) AS NOMBRE," & _
          " (select descrip from pl_periodos where codigo=proceso) as planilla," & _
          " anomes,  proceso,  emp,  sueldo,  proy,  grati,  otrosing,  remant,  totrem,  impuit," & _
          " remneta, uit5,  uit20, uit27,  uit54,  uitexc,  tasa8,  tasa14,  tasa15,  tasa21,  tasa30,  totret,  retant,  retneta,  divisor," & _
          " rtaqta   FROM pl_rtaqta p left join empleado c on (c.codigo=p.emp) " & _
          " WHERE LEFT(ANOMES,4)='" & Left(AnoMes, 4) & "' AND  ANOMES<='" & AnoMes & "'" & str & Str1 & strTipo & _
          " order by ANOMES"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    End If
End Sub

Public Sub sp_Rep_Boleta(AnoMes As String, plani As String, mon As String, emp As String, Optional opcMostrar As Boolean)
    Dim rsprincipal As MYSQL_RS
    Dim rsIng As MYSQL_RS
    Dim rsDes As MYSQL_RS
    Dim rsApo As MYSQL_RS
    Dim SQL As String, strTipo As String
    Dim str As String, str2 As String, str3 As String, strmon As String
    If emp <> "" And emp <> "00000000000" Then
        str = " and a.emp='" & emp & "'"
    Else
        str = ""
    End If
    If mon = "C" Or mon = "E" Then
        strTipo = " and c.tipoplani = 'S' "
    ElseIf mon = "N" Then
        strTipo = " and c.tipoplani = 'N' "
    Else
        If opcMostrar = True Then
            strTipo = ""
        Else
            strTipo = " and c.tipoplani = 'N' "
        End If
    End If
    Select Case plani
        Case "1":
            str2 = " if(a.unidad='D',a.cant*(a.sbasico/30),a.cant) "
            str3 = " AND rub='121' "
        Case "2":
            str2 = " a.cant*.80*(a.sbasico/30) , 0 ))"
            str3 = " AND rub='121' "
        Case "4":
            str2 = " a.cant*(a.sbasico/30) , 0 ))"
            str3 = " AND rub='118' "
        Case "5":
            str2 = " if(a.unidad='D',a.cant*(a.sbasico/180),a.cant) "
            str3 = " AND rub='121' "
        Case "6"
            m_Reporte = "Rep_Boleta_CTS.rpt"
            SQL = "SELECT c.codafp, a.anomes,pp.fecha as fecproc," & _
                  " (select tipo_ven from cncambio where fecha=concat(right(pp.fecha,2),SUBSTRING(pp.fecha,5,4),left(pp.fecha,4)) ) as tc," & _
                  " a.emp,tr.tipo as tiprub , r.descrip," & _
                  " CONCAT(C.APEPAT,' ',C.APEMAT,' ',C.NOMBRE1,' ',C.NOMBRE2) AS NOMBRE," & _
                  " c.direccion,(select descrip from cncargos where codigo=(select codcargo from contrato where codemp=c.codigo order by codigo desc limit 1) ) as cargo," & _
                  " (select area from cncargos where codigo=(select codcargo from contrato where codemp=c.codigo order by codigo desc limit 1) ) as ger, C.FEC_INGRESO AS f_inicio," & _
                  " c.numdocide,c.fec_nac,c.sexo,a.tipo as proceso,a.moneda," & _
                  " (select descrip from pl_entidadfinanciera where codigo=c.ctsbanco) as bcts, " & _
                  " c.ctsmon,c.ctsnumcta," & _
                  " (select cant from pl_tareo where ANOMES='" & AnoMes & "' AND TIPO='6'" & _
                  " and moneda='" & IIf(mon = "C", "N", mon) & "'  AND rub='121'  and emp=a.emp) as cantid," & _
                  " a.rub,if(a.unidad='V',a.cant,a.sbasico) as cant ,a.sbasico, a.fecha" & _
                  " FROM ((((pl_tareo as a left join contrato as b on (a.emp=b.codemp))" & _
                  " LEFT JOIN empleado  AS C ON (a.emp=c.codigo))" & _
                  " left join pl_rubrosremunerativos as R on (a.rub=r.codigo))" & _
                  " left join pl_tiporubros as tr on (tr.codigo=r.tipo) ) left join pl_planiproc as pp" & _
                  " on (a.anomes=pp.anomes and a.tipo=pp.proceso and a.moneda=pp.mon)" & _
                  " WHERE B.CODIGO=A.CODCONTRATO AND  a.anomes='" & AnoMes & "' AND a.tipo='" & plani & "' AND a.moneda='" & IIf(mon = "C", "N", mon) & "'" & str & strTipo & _
                  " GROUP BY a.anomes,a.emp,a.tipo,a.moneda,a.fecha,a.rub" & _
                  " ORDER BY C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2,r.orden"
    End Select
    If plani <> "6" Then
    
    
    Dim cCadena As String
    
 
    cCadena = AnoMes
    ' VERIFICAR if(c.categoria='01',168,ifnull(dh.horas,0)) as horas CUANDO ES CATEGORIA DIRECCION EN EL RPT 168 HORAS

        SQL = "SELECT (select nombre from afp where codigo=c.codafp) as dafp ,c.codafp, '" & cCadena & "' as anomes, a.emp,tr.tipo as tiprub , r.descrip, CONCAT(C.APEPAT,' ',C.APEMAT,' ',C.NOMBRE1,' ',C.NOMBRE2) AS NOMBRE,c.direccion, " & _
          " (select descrip from cncargos where codigo=(select codcargo from contrato where codemp=c.codigo order by codigo desc limit 1) ) as cargo," & _
          " (select departamento from cncargos where codigo=(select codcargo from contrato where codemp=c.codigo order by codigo desc limit 1) ) as departamento," & _
          " (select area from cncargos where codigo=(select codcargo from contrato where codemp=c.codigo order by codigo desc limit 1) ) as area," & _
          " (select DESCRIP from CNMDEPAR where CODDEP=b.DIVISION) as DIVISION, c.fec_ingreso as f_inicio," & _
          " c.numdocide,c.fec_nac,c.sexo, " & _
          " (select descrip from nacionalidad where codigo=c.nacionalidad) as nacion, " & _
          " c.numafp,(select descrip from rh_categoria where codigo=c.categoria) as categor," & _
          " c.numseg,cal.fec_salida,cal.fec_regreso,cal.fec_fin,a.tipo as proceso,a.moneda," & _
          " ifnull(dh.dias,0) as cantid, if(c.categoria='01',ifnull(dh.horas,0),ifnull(dh.horas,0)) as horas ," & _
          " a.rub," & str2 & " as cant ,a.sbasico, a.fecha FROM (((((pl_tareo as a left join contrato as b on (a.emp=b.codemp))" & _
          " LEFT JOIN empleado  AS C ON (a.emp=c.codigo))left join pl_rubrosremunerativos as R on (a.rub=r.codigo))" & _
          " left join pl_tiporubros as tr on (tr.codigo=r.tipo) )" & _
          " left join (  select codemp," & _
          " if(replace(min(fec_salida),'/','') < '" & AnoMes & "01','" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01',  IF(left(min(fec_salida),4)='" & Left(AnoMes, 4) & "',  min(fec_salida),'') ) as fec_salida," & _
          " if(replace(max(fec_regreso),'/','') > '" & AnoMes & DiasMes(Right(AnoMes, 2), Left(AnoMes, 4)) & "',if( left(max(fec_regreso),4)='" & Left(AnoMes, 4) & "'  OR left(min(fec_salida),4)='" & Left(AnoMes, 4) & "' ,'" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & DiasMes(Right(AnoMes, 2), Left(AnoMes, 4)) & "','') , max(fec_regreso)) as fec_regreso,fec_regreso as fec_fin" & _
          " From calendario WHERE gocehaber='N' and substring(replace(fec_regreso,'/',''),1,6)>='" & AnoMes & "' and substring(if(replace(fec_salida,'/','') < '" & AnoMes & "01','" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01', fec_salida),6,2)<='" & Right(AnoMes, 2) & "'" & _
          " and movemp='02' " & _
          " and LEFT(fec_regreso,4)=  IF( LEFT(fec_regreso,4)='" & Left(AnoMes, 4) & "', '" & Left(AnoMes, 4) & "' , left(fec_regreso,4)) " & _
          " group by codemp) as cal on (a.emp=cal.codemp and a.tipo='1')) left join " & _
          " (Select emp,count(*) as dias,sum(if(tiposede='O',if(DATE_FORMAT(fecha, '%W')='Saturday',4,8),0)) + sum(if(tiposede='C',11,0)) as Horas " & _
          " From rh_entsalempleado where substring(fecha,1,7)='" & Trim(Left(AnoMes, 4) & "/" & Right(AnoMes, 2)) & "' and tipo='E' group by emp) as dh on (dh.emp=a.emp) " & _
          " WHERE a.anomes='" & AnoMes & "' AND a.tipo='" & plani & "' AND a.moneda='" & IIf(mon = "C", "N", mon) & "'" & str & strTipo & _
          " GROUP BY a.emp,a.tipo,a.moneda,a.fecha,a.rub" & _
          " ORDER BY C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2,r.orden"

          
    End If
    Set rsprincipal = oConexion.EjecutaSelectRS(SQL)
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, rsprincipal
    Set rsprincipal = Nothing
End Sub

Public Sub sp_Rep_CRQ(Anio As String, AnoMes As String, plani As String, mon As String, emp As String, Optional opcMostrar As Boolean)
    Dim rsprincipal As MYSQL_RS, OtrosIng As Double
    Dim strTipo As String, fec As String, Vac As Double, Grati As Double
    Dim str As String, str2 As String, str3 As String, strmon As String
    If emp <> "" And emp <> "00000000000" Then
        str = " and C.CODIGO='" & emp & "'"
    Else
        str = ""
    End If
    
'    If opcMostrar = True Then
'        strTipo = ""
'    Else
'        strTipo = " and c.tipoplani = 'N' "
'    End If
    
'    If mon = "C" Or mon = "E" Then
'        strTipo = " and c.tipoplani = 'S' "
'    Else
'        If mon = "N" Then
'            strTipo = " and c.tipoplani = 'N' "
'        Else
'            If opcMostrar = True Then
'                strTipo = ""
'            Else
'                strTipo = " and c.tipoplani = 'N' "
'            End If
'        End If
'    End If
    
    
    If emp <> "" And emp <> "00000000000" Then
        fec = InputBox("Ingrese la Fecha de Cese", "LIQUIDACION", CStr(Date))
        m_Fecha = fec
        OtrosIng = InputBox("Otros Ingresos", "LIQUIDACION", 0)
    End If
    Dim SQL As String, RQ As MYSQL_RS
    Dim FlgCalculo As Boolean
    Dim StrAnio As String
    SQL = "delete from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQL
    SQL = "insert into rh_templiquidacion(nombres,tiempo,codigo,fecingreso,ultgrati,porafp,mtoentregas,prestamos,adelantos, " & _
          " quinta,fecinicts,ctacts,bancocts,mtogracia,campo1,campo2,usuario,moncts,feccese)" & _
          " select CONCAT(C.APEPAT,' ',C.APEMAT,' ',C.NOMBRE1,' ',C.NOMBRE2) AS NOMBRE,c.numdocide,c.codigo,c.fec_ingreso,a.grati,a.impuit," & _
          " a.remneta,a.uit27,a.uit54,a.uitexc,a.tasa15,a.tasa21,a.tasa30,a.totret,0 as sueldoa," & _
          " (select sum(otrosing) from pl_rtaqta where emp=a.emp and proceso='1' and left(anomes,4)='" & Anio & "') as otrosinga,'" & strUsuarioId & "',c.situacion,c.fec_cese " & _
          " from pl_rtaqta as a left join empleado as c on (a.emp=c.codigo)" & _
          " where a.proceso='1' and a.anomes=(select max(anomes) from pl_rtaqta p where a.totret> 0 and p.emp=a.emp and left(anomes,4) = '" & Anio & "' and p.proceso = 1) " & str & strTipo & _
          " and (left(c.fec_cese,4)>'" & Anio & "' or c.fec_cese='' ) order by C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2"
    oConexionMYSQL.Execute SQL
    FlgCalculo = False
    SQL = "select * from rh_templiquidacion c where usuario = '" & strUsuarioId & "'" & str
    Set rsprincipal = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsprincipal.EOF
        FlgCalculo = False
        If rsprincipal.Fields("feccese") <> "" And Left(rsprincipal.Fields("feccese"), 4) = Anio Then
            fec = rsprincipal.Fields("feccese")
            FlgCalculo = True
            If fec = "" Then
                fec = "02/01/" & Anio + 1
                If Format(fec, "dddd") = "Sábado" Then fec = "04/01/" & Anio + 1
                If Format(fec, "dddd") = "Domingo" Then fec = "03/01/" & Anio + 1
                SQL = "update rh_templiquidacion set feccese='" & Format(fec, "dd/mm/yyyy") & "' where usuario = '" & strUsuarioId & "' " & _
                      "and codigo = '" & rsprincipal.Fields("codigo") & "'"
                oConexionMYSQL.Execute SQL
            End If
        Else
            If IsDate(fec) Then
                FlgCalculo = True
            Else
                SQL = "select ifnull(sum(if(p.unidad='V',cant,(sbasico * cant)/30)),0) AS TOT from pl_tareo p left join " & _
                      " pl_rubrosremunerativos r on (p.rub=r.codigo) left join pl_tiporubros t on (t.codigo=r.tipo) left join pl_rubros_afectos as d on (d.codrub= p.rub)" & _
                      " where left(anomes,4) = '" & Anio & "' and emp = '" & rsprincipal.Fields("codigo") & "' and d.codafec='10' and p.tipo = 1 and t.tipo='I' and p.rub NOT IN ('910','916')"
                Set RQ = oConexion.EjecutaSelectRS(SQL)
                If Not RQ.EOF() Then
                    SQL = "update rh_templiquidacion set CAMPO1 = " & CDbl(RQ.Fields("tot")) & " where usuario = '" & strUsuarioId & "' " & _
                          "and codigo = '" & rsprincipal.Fields("codigo") & "'"
                    oConexionMYSQL.Execute SQL
                End If
                SQL = "select ifnull(sum(if(p.unidad='V',cant,(sbasico * cant)/30)),0) as totgrati from pl_tareo p left join " & _
                      "pl_rubrosremunerativos r on (p.rub=r.codigo) left join pl_tiporubros t on (t.codigo=r.tipo) " & _
                      "where left(anomes,4) = '" & Anio & "' and emp = '" & rsprincipal.Fields("codigo") & "' and t.tipo='I' and ( (p.tipo = '1'  and (p.rub =  '403' or p.rub = '313' or p.rub = '910'  or p.rub = '1002' or p.rub = '1005' OR p.rub = '117' )) or  (p.tipo = '5' and p.rub = '1003'))"
                Set RQ = oConexion.EjecutaSelectRS(SQL)
                If Not RQ.EOF() Then
                    fec = "02/01/" & Anio + 1
                    If Format(fec, "dddd") = "Sábado" Then fec = "04/01/" & Anio + 1
                    If Format(fec, "dddd") = "Domingo" Then fec = "03/01/" & Anio + 1
                    SQL = "update rh_templiquidacion set CAMPO3 = " & CDbl(RQ.Fields("totgrati")) & ",feccese='" & Format(fec, "dd/mm/yyyy") & "' where usuario = '" & strUsuarioId & "' " & _
                          "and codigo = '" & rsprincipal.Fields("codigo") & "'"
                    oConexionMYSQL.Execute SQL
                    fec = ""
                End If
                SQL = "select max(anomes) as anomes from pl_rtaqta a where A.emp='" & rsprincipal.Fields("CODIGO") & "' and left(anomes,4) = '" & Anio & "' and A.proceso = 1"
                Set RQ = oConexion.EjecutaSelectRS(SQL)
                If Not RQ.EOF() Then
                    If val(Right(RQ.Fields("ANOMES"), 2)) >= 7 And val(Right(RQ.Fields("ANOMES"), 2)) < 12 Then
                        StrAnio = Anio & "07"
                        fec = Date
                    ElseIf val(Right(RQ.Fields("ANOMES"), 2)) > 7 And val(Right(RQ.Fields("ANOMES"), 2)) <= 12 Then
                        StrAnio = Anio & "12"
                    Else
                        fec = Date
                        StrAnio = ""
                    End If
                    If StrAnio <> "" Then
                        SQL = "select sum(grati) AS totgrati from pl_rtaqta p " & _
                              "where emp = '" & rsprincipal.Fields("CODIGO") & "' and PROCESO = 1 " & _
                              "and anomes in (" & StrAnio & ")"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If Not RQ.EOF() Then
                            SQL = "update rh_templiquidacion set toting = " & IIf(IsNull(RQ.Fields("totgrati")), 0, RQ.Fields("totgrati")) & " where usuario = '" & strUsuarioId & "' " & _
                                  "and codigo = '" & rsprincipal.Fields("codigo") & "'"
                            oConexionMYSQL.Execute SQL
                        End If
                    End If
                    If IsDate(fec) Then
                        If CDate(fec) = Date Then
                            SQL = "update rh_templiquidacion set feccese='" & Format(fec, "dd/mm/yyyy") & "' where usuario = '" & strUsuarioId & "' " & _
                                  "and codigo = '" & rsprincipal.Fields("codigo") & "'"
                            oConexionMYSQL.Execute SQL
                            fec = ""
                        End If
                    End If
                End If
            End If
        End If
        If FlgCalculo = True Then
            Dim RQ1 As MYSQL_RS
            Vac = 0
            Grati = 0
            SQL = "select ifnull(sum(if(p.unidad='V',cant,(sbasico * cant)/30)),0) AS TOT from pl_tareo p left join " & _
                  "pl_rubrosremunerativos r on (p.rub=r.codigo) left join pl_tiporubros t on (t.codigo=r.tipo)left join pl_rubros_afectos as d on (d.codrub= p.rub) " & _
                  "where left(anomes,4) = '" & Anio & "' and emp = '" & rsprincipal.Fields("codigo") & "' and d.codafec='10' and p.tipo = 1 and t.tipo='I' and p.rub NOT IN ('910','916')"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                SQL = "update rh_templiquidacion set CAMPO1 = " & CDbl(RQ.Fields("tot")) & " where usuario = '" & strUsuarioId & "' " & _
                      "and codigo = '" & rsprincipal.Fields("codigo") & "'"
                oConexionMYSQL.Execute SQL
            End If
            SQL = "select ifnull(sum(if(p.unidad='V',cant,(sbasico * cant)/30)),0) AS TOT from pl_tareo p left join " & _
                  "pl_rubrosremunerativos r on (p.rub=r.codigo) left join pl_tiporubros t on (t.codigo=r.tipo) " & _
                  "where left(anomes,4) = '" & Anio & "' and emp = '" & rsprincipal.Fields("codigo") & "' and t.tipo='I' and ( (p.tipo = '1'  and (p.rub = '403' or p.rub = '313'  or p.rub = '910'  or p.rub = '1002' or p.rub = '1005' OR p.rub = '117' )) or  (p.tipo = '5' and p.rub = '1003'))"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            
            If Not RQ.EOF() Then
                SQL = "update rh_templiquidacion set CAMPO3 = " & CDbl(RQ.Fields("tot")) & " where usuario = '" & strUsuarioId & "' " & _
                      "and codigo = '" & rsprincipal.Fields("codigo") & "'"
                oConexionMYSQL.Execute SQL
            End If
            
            If Month(fec) = 12 Then
                SQL = "select max(anomes) as anomes from pl_rtaqta a where A.emp='" & rsprincipal.Fields("CODIGO") & "' and left(anomes,4) = '" & Anio & "' and A.proceso = 1"
                Set RQ = oConexion.EjecutaSelectRS(SQL)
                If Not RQ.EOF() Then
                    If val(Right(RQ.Fields("ANOMES"), 2)) >= 7 And val(Right(RQ.Fields("ANOMES"), 2)) < 12 Then
                        StrAnio = Anio & "07"
                    ElseIf val(Right(RQ.Fields("ANOMES"), 2)) > 7 And val(Right(RQ.Fields("ANOMES"), 2)) <= 12 Then
                        StrAnio = Anio & "12"
                    Else
                        StrAnio = ""
                    End If
                    If StrAnio <> "" Then
                        SQL = "select sum(grati) AS totgrati from pl_rtaqta p " & _
                              "where emp = '" & rsprincipal.Fields("CODIGO") & "' and PROCESO = 1 " & _
                              "and anomes in (" & StrAnio & ")"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If Not RQ.EOF() Then
                            Grati = CDbl(RQ.Fields("totgrati"))
                        End If
                    End If
                End If
            Else
                SQL = "SELECT SUM(ifnull(((sbasico/cant)*180),0)) as cantidad from pl_tareo where emp = '" & rsprincipal.Fields("codigo") & "' and tipo = 5 and rub = 121 and left(anomes,4) = '" & Anio & "'"
                Set RQ = oConexion.EjecutaSelectRS(SQL)
                If Not RQ.EOF() Then
                    Grati = IIf(IsNull(RQ.Fields("cantidad")), 0, RQ.Fields("cantidad"))
                    SQL = "SELECT SUM(ifnull(cant,0)) as cantidad from pl_tareo where emp = '" & rsprincipal.Fields("codigo") & "' and tipo = '5' and rub = '201' and left(anomes,4) = '" & Anio & "'"
                    Set RQ = oConexion.EjecutaSelectRS(SQL)
                    If Not RQ.EOF() Then
                        Grati = Grati + IIf(IsNull(RQ.Fields("cantidad")), 0, RQ.Fields("cantidad"))
                        SQL = "SELECT SUM(ifnull(cant,0)) as cantidad from pl_tareo where emp = '" & rsprincipal.Fields("codigo") & "' and tipo = '5' and rub = '1001' and left(anomes,4) = '" & Anio & "'"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If Not RQ.EOF() Then
                            Grati = Grati + IIf(IsNull(RQ.Fields("cantidad")), 0, RQ.Fields("cantidad"))
                        End If
                    End If
                End If
            End If
            SQL = "update rh_templiquidacion set toting = " & Vac & ",mtoreten= " & Grati & " where usuario = '" & strUsuarioId & "' " & _
                  "and codigo = '" & rsprincipal.Fields("codigo") & "'"
            oConexionMYSQL.Execute SQL
            fec = ""
        End If
        rsprincipal.MoveNext
    Loop
    SQL = "select nombres as nombre,tiempo as numdocide,codigo,fecingreso,ultgrati as grati,(porafp *(-1)) as impuit,mtoentregas as remneta " & _
          ",prestamos as uit27,adelantos as uit54,quinta as uitexc,fecinicts as tasa15,ctacts as tasa21,bancocts as tasa30,mtogracia as totret " & _
          ",campo1 as sueldoa,campo2 as otrosinga,toting as vac,mtoreten as gratifi,campo3,feccese,'" & Anio & "' as anio from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
    Set rsprincipal = oConexion.EjecutaSelectRS(SQL)
    If rsprincipal.EOF Then
        Set rsprincipal = Nothing
        MsgBox "No existe Datos", vbOKOnly + vbInformation, "NOVADMIN"
        Exit Sub
    End If
    m_Titulo = "EJERCICIO GRAVABLE " & Anio
    m_Igv = CStr(OtrosIng)
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, rsprincipal
    Set rsprincipal = Nothing
End Sub
Public Sub sp_Rep_CAFP(AnoMes As String, plani As String, mon As String, emp As String, Optional opcMostrar As Boolean)
    Dim rsprincipal As MYSQL_RS
    Dim strTipo As String
    Dim str As String, str2 As String, str3 As String, strmon As String
    If emp <> "" And emp <> "00000000000" Then
        str = " and a.emp='" & emp & "'"
    Else
        str = ""
    End If
    If mon = "C" Or mon = "E" Then
        strTipo = " and c.tipoplani = 'S' "
    ElseIf mon = "N" Then
        strTipo = " and c.tipoplani = 'N' "
    Else
        If opcMostrar = True Then
            strTipo = ""
        Else
            strTipo = " and c.tipoplani = 'N' "
        End If
    End If
    Dim SQL As String
    SQL = "Select A.EMP,A.RUB,SUM(A.CANT)," & _
          "(select CONCAT(APEPAT,' ',APEMAT,' ',NOMBRE1,' ',NOMBRE2) FROM EMPLEADO WHERE CODIGO=A.EMP) AS NOMBRE, " & _
          " (select NUMDOCIDE FROM EMPLEADO WHERE CODIGO=A.EMP) AS NUMDOCIDE," & _
          " (select NUMAFP FROM EMPLEADO WHERE CODIGO=A.EMP) AS NUMFP," & _
          " (select DESCRIP FROM pl_rubrosremunerativos WHERE CODIGO=A.RUB) AS DESRUBRO," & _
          " (Select SUM(CANT) from pl_tareo where" & _
          " (rub='608' ) AND (TIPO='1' OR TIPO='5')" & _
          " AND LEFT(ANOMES,4)='" & Left(AnoMes, 4) & "' and emp=A.emp) as ingresos, " & _
          " (Select SUM(CANT) from pl_tareo where " & _
          " (rub='601' OR rub='606' OR rub='608' OR rub='611') AND (TIPO='1' OR TIPO='5')" & _
          " AND LEFT(ANOMES,4)='" & Left(AnoMes, 4) & "' and emp=A.emp) as total" & _
          " from pl_tareo AS A left join empleado c on (a.emp=c.codigo)" & _
           " Where " & _
          " (A.rub='601' OR " & _
          " A.rub='606' OR" & _
          " A.rub='608' OR " & _
          " A.rub='611') AND (A.TIPO='1' OR A.TIPO='5') AND LEFT(A.ANOMES,4)='" & Left(AnoMes, 4) & "' " & str & strTipo & _
          " GROUP BY A.EMP,A.RUB" & _
          " order by NOMBRE,DESRUBRO "
    Set rsprincipal = oConexion.EjecutaSelectRS(SQL)
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, rsprincipal
    Set rsprincipal = Nothing
End Sub

Public Function sp_ValidaTareo(AnoMes As String, TipPlani As String, moneda As String, Optional TextoPlanilla As String = "") As Boolean

    Screen.MousePointer = vbHourglass

    Dim SQL As String
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS

    SQL = "call PL_SP_Validacion_Tareo ('REPORTE_VALIDACION','" & AnoMes & "','" & TipPlani & "','" & moneda & "')"

    Set oRs = ADO_LlenaRs(SQL)

    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No existen datos", vbOKOnly + vbInformation, gsNomSW

            Screen.MousePointer = vbNormal
            Exit Function
        End If

        m_Titulo = "VALIDACION DEL TAREO: " & AnoMes & " " & TextoPlanilla

        EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, oRSMYSQL, oRs, False
    Else
        MsgBox "No existen datos", vbOKOnly + vbInformation, gsNomSW
    End If

    Screen.MousePointer = vbNormal
    Set oRs = Nothing

End Function

Public Sub sp_Rep_Planilla(AnoMes As String, plani As String, mon As String, flag As String, Cod As String, Optional opcMostrar As Boolean)
    Dim rsprincipal As MYSQL_RS, tc As Double, a As Variant
    Dim str As String, strmon As String, strflag As String, strcod As String
    Dim NomArchivoExcel As String, strTipo As String
    Dim SQL As String
    If mon <> "" Then
        strmon = " AND A.MONEDA='" & IIf(mon = "C", "N", mon) & "' "
    Else
        strmon = ""
    End If
    Select Case flag
        Case "AFP":
            strflag = "(select nombre from afp where codigo=a.afp) as dafp"
            strcod = " "
        Case "DIVI":
            strflag = "(select descrip from cnmdepar where codDep=t.division) as ddivi"
            If Cod <> "" Then
                strcod = " and t.division='" & Right("0000" & Trim(Cod), 4) & "'"
            Else
                strcod = ""
            End If
    End Select
    Select Case plani
        Case "1":
            str = " if(a.unidad='D',a.cant*(a.sbasico/30),a.cant) , '' ))"
            m_Titulo = "PLANILLA MENSUAL DE REMUNERACIONES " & NombreMes(Right(AnoMes, 2), False) & " - " & Left(AnoMes, 4)
        Case "2":
            str = "     a.cant*.80*(a.sbasico/30) , '' ))"
            m_Titulo = "PLANILLA QUINCENAL DE REMUNERACIONES " & NombreMes(Right(AnoMes, 2), False) & " - " & Left(AnoMes, 4)
        Case "4":
            str = " a.cant*(a.sbasico/30) , '' ))"
            m_Titulo = "PLANILLA VACACIONAL DE REMUNERACIONES " & NombreMes(Right(AnoMes, 2), False) & " - " & Left(AnoMes, 4)
        Case "5":
            str = " if(a.unidad='D',a.cant*(a.sbasico/180),a.cant) , '' ))"
            m_Titulo = "PLANILLA DE GRATIFICACION A " & NombreMes(Right(AnoMes, 2), False) & " - " & Left(AnoMes, 4)
        Case "6":
            str = " 360*(a.sbasico/360) , '' ))"
            m_Titulo = "PLANILLA DE CTS A " & NombreMes(Right(AnoMes, 2), False) & " - " & Left(AnoMes, 4)
            'm_Reporte = "Rep_PlanillaCTS.rpt"
            m_Reporte = "Rep_PlanillaCTS_Prueba08112018.rpt"
    End Select
    If mon = "C" Or mon = "E" Then
        strTipo = " and c.tipoplani = 'S' "
    Else
        If opcMostrar = True Then
            strTipo = ""
        Else
            strTipo = " and c.tipoplani = 'N' "
        End If
    End If
    If plani = "6" Then
        modPrincipal.TipoCambio
        a = InputBox("Ingrese el tipo de cambio", "TIPO DE CAMBIO", CStr(dblTipoCmbV))
        tc = CDbl(IIf(IsNumeric(a), a, 1))
        SQL = "SELECT A.FECHA,C.CTSMON,(Select descrip from  pl_entidadfinanciera where codigo=c.ctsbanco) as BANCO,A.EMP,CONCAT_WS(' ',C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2) as nombre," & _
              "C.FEC_INGRESO," & _
              " c.numdocide,C.CTSNUMCTA,(select cant from pl_tareo where ANOMES='" & AnoMes & "' AND TIPO='" & plani & "'" & _
              " AND rub='121' and unidad='D' and emp=a.emp) as cantid,a.sbasico," & _
              " sum( IF( a.rub ='121', " & str & " AS 'I121'," & _
              " sum( IF( a.rub ='1001', a.cant, 0 )) AS 'I1001'," & _
              " sum( IF( a.rub ='201', a.cant, 0 )) AS 'I201'," & _
              " sum( IF( a.rub ='401', a.cant, 0 )) AS 'I401'," & _
              " sum( IF( a.rub ='105', a.cant, 0 )) AS 'I105', " & _
              " sum( IF( a.rub ='1002', a.cant, 0 )) AS 'I1002'," & _
              " sum( IF( a.rub ='701', a.cant, 0 )) AS 'D701'," & _
              tc & " as tcambio  " & _
              " FROM pl_tareo as a LEFT JOIN" & _
              " empleado  AS c ON (a.EMP=c.CODIGO) left join contrato t on (t.codemp=C.codigo) and (t.estado = 'AP')" & _
              " WHERE a.ANOMES='" & AnoMes & "' AND a.TIPO='" & plani & "'" & strTipo & _
              " GROUP BY a.anomes,a.emp,a.tipo,a.fecha ORDER BY C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2"
    Else
        SQL = "SELECT " & strflag & " ,a.anomes,a.emp,CONCAT(C.APEPAT,' ',C.APEMAT,' ',C.NOMBRE1,' ',C.NOMBRE2) AS NOMBRE," & _
              " t.division, t.divgas , a.tipo as proceso,a.moneda,(select SUM(cant) from pl_tareo where ANOMES='" & AnoMes & "' AND TIPO='" & plani & "'" & _
              " AND (rub='121' or rub='117') and unidad='D' and emp=a.emp) as cantid,a.sbasico," & _
              " sum( IF( a.rub ='121', " & str & " AS 'I121'," & _
              " sum( IF( a.rub ='1001', a.cant, 0 )) AS 'I1001',sum( IF( a.rub ='1005', a.cant, 0 )) AS 'I1005',sum( IF( a.rub ='1008', a.cant, 0 )) AS 'I1008',sum( IF( a.rub ='1007', a.cant, 0 )) AS 'I1007',sum( IF( a.rub ='1006', a.cant, 0 )) AS 'I1006',(sum( IF( a.rub ='1002', a.cant, 0 )) + sum( IF( a.rub ='107', a.cant, 0 )))  AS 'I1002',sum( IF( a.rub ='1003', a.cant, 0 )) AS 'I1003'," & _
              " sum( IF( a.rub ='102', a.cant, 0 )) AS 'I102',sum( IF( a.rub ='105', a.cant, 0 )) AS 'I103'," & _
              " sum( IF( a.rub ='115', a.cant, 0 )) AS 'I115',sum( IF( a.rub ='116', a.cant, 0 )) AS 'I116',sum( IF( a.rub ='118', a.cant, 0 )) AS 'I118',sum( IF( a.rub ='117', " & str & " AS 'I117'," & _
              " sum( IF( a.rub ='201', a.cant, 0 )) AS 'I201',sum( IF( a.rub ='202', a.cant, 0 )) AS 'I202',sum( IF( a.rub ='313', a.cant, 0 )) AS 'I313'," & _
              " sum( IF( a.rub ='904', a.cant, 0 )) AS 'I904',sum( IF( a.rub ='909', a.cant, 0 )) AS 'I909',sum( IF( a.rub ='915', a.cant, 0 )) AS 'I915',sum( IF( a.rub ='916', a.cant, 0 )) AS 'I916',sum( IF( a.rub ='910', a.cant, 0 )) AS 'I910',sum( IF( a.rub ='911', a.cant, 0 )) AS 'I911',sum( IF( a.rub ='902', a.cant, 0 )) AS 'I902'," & _
              " sum( IF( a.rub ='401', a.cant, 0 )) AS 'I401',sum( IF( a.rub ='403', a.cant, 0 )) AS 'I403',sum( IF( a.rub ='404', a.cant, 0 )) AS 'I404',sum( IF( a.rub ='407', a.cant, 0 )) AS 'I407',sum( IF( a.rub ='504', a.cant, 0 )) AS 'I504'," & _
              " '' AS 'TOTAL INGRESOS', sum( IF( a.rub ='604', a.cant, 0 )) AS 'D604', " & _
              " sum( IF( a.rub ='605', a.cant, 0 )) AS 'D605',sum( IF( a.rub ='607', a.cant, 0 )) AS 'D607'," & _
              " sum( IF( a.rub ='608', a.cant, 0 )) AS 'D608',sum( IF( a.rub ='601', a.cant, 0 )) AS 'D601', sum( IF( a.rub ='606', a.cant, 0 )) AS 'D606'," & _
              " sum( IF( a.rub ='611', a.cant, 0 )) AS 'D611',sum( IF( a.rub ='701', a.cant, 0 )) AS 'D701'," & _
              " sum( IF( a.rub ='703', a.cant, 0 )) AS 'D703',sum( IF( a.rub ='704', a.cant, 0 )) AS 'D704', sum( IF( a.rub ='705', a.cant, 0 )) + sum( IF( a.rub ='706', a.cant, 0 )) AS 'D706'," & _
              " sum( IF( a.rub ='707', a.cant, 0 )) AS 'D707',sum( IF( a.rub ='708', a.cant, 0 )) AS 'D708'," & _
              " sum( IF( a.rub ='709', a.cant, 0 )) AS 'D709',sum( IF( a.rub ='710', a.cant, 0 )) AS 'D710'," & _
              " sum( IF( a.rub ='711', a.cant, 0 )) AS 'D711',sum( IF( a.rub ='712', a.cant, 0 )) AS 'D712',sum( IF( a.rub ='713', a.cant, 0 )) AS 'D713',sum( IF( a.rub ='714', a.cant, 0 )) AS 'D714','' AS 'TOTAL DESCUENTOS','' AS 'NETO A PAGAR' ,sum( IF( a.rub ='804', a.cant, 0 )) AS 'A804'," & _
              " sum( IF( a.rub ='806', a.cant, 0 )) AS 'A806',sum( IF( a.rub ='807', a.cant, 0 )) AS 'A807',sum( IF( a.rub ='805', a.cant, 0 )) AS 'A805'" & _
              " FROM pl_tareo as a LEFT JOIN" & _
              " empleado  AS c ON (a.EMP=c.CODIGO) LEFT JOIN CONTRATO T ON (T.CODEMP=C.CODIGO) AND (T.codigo=a.codcontrato) WHERE a.ANOMES='" & AnoMes & "' AND a.TIPO='" & plani & "'" & _
              strmon & strcod & strTipo & " GROUP BY a.anomes,a.emp,t.division, t.divgas , a.tipo,a.moneda ORDER BY C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2"
    End If
    Set rsprincipal = oConexion.EjecutaSelectRS(SQL)
    m_Empresa = "NATIONAL OILWELL VARCO PERU S.R.L"
    If frmRepPlanilla.chkExcel.Value = 1 Then
        If OpcExport = True Then
            NomArchivoExcel = Rep_Documents & "\PLANILLA" & AnoMes & "_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Mid(Format(Date, "dd/mm/yy"), 7, 8) & IIf(CantVersiones = 0, Empty, "_" & CantVersiones) & ".XLS"
        Else
            NomArchivoExcel = Rep_Documents & "\PLANILLA" & AnoMes & ".XLS"
        End If
        If Exportar_Excel_rs(NomArchivoExcel, rsprincipal) Then
        End If
    Else
        EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, rsprincipal
    End If
    Set rsprincipal = Nothing
End Sub
Public Sub sp_Rep_PlanillaTW(AnoMes As String, plani As String, mon As String)
    Dim SQL As String, num As Integer, Cant As Integer, str As String, tw As String, Tit As String
    Dim myrs As MYSQL_RS, rs_Horas As MYSQL_RS
    Set myrs = New MYSQL_RS
    Select Case plani
        Case "1":
            str = " if(a.unidad='D',a.cant*(a.sbasico/30),a.cant) "
            Tit = "MENSUAL  DE REMUNERACIONES"
        Case "2":
            str = " a.cant * .80 * (a.sbasico/30) "
            Tit = "QUINCENAL DE REMUNERACIONES"
        Case "4":
            str = " if(a.unidad='D',a.cant*(a.sbasico/30),a.cant) "
            Tit = " DE VACACIONES"
        Case "5":
            str = " if(a.unidad='D',a.cant*(a.sbasico/180),a.cant) "
            Tit = "DE GRATIFICAIONES "
        Case "6":
            str = " if(a.unidad='D',360*(a.sbasico/360) , a.cant ),0)"
            m_Titulo = "PAGO DE PLANILLA DE CTS A " & NombreMes(Right(AnoMes, 2), False) & " - " & Left(AnoMes, 4)
    End Select
    SQL = "Select * from pl_planiproc where anomes='" & AnoMes & "' and proceso='" & plani & "' and mon='" & mon & "'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.EOF And myrs.BOF Then
        tw = "SIN GENERAR"
    Else
        tw = myrs.Fields("tw") & ".txt"
    End If
    m_Titulo = "PAGO DE PLANILLA " & Tit & " MEDIANTE TELEWIESE " & tw
    If plani = "6" Then
        m_Titulo = "PAGO DE PLANILLA DE CTS A " & NombreMes(Right(AnoMes, 2), False) & " - " & Left(AnoMes, 4) & " MEDIANTE TELEWIESE " & tw
        SQL = "SELECT a.anomes,a.emp,'3' as tipcta_mn,c.ctsnumcta as numcta_mn,'3' as tipcta_me," & _
              " c.ctsnumcta as numcta_me,CONCAT_WS(' ',C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2) as nombre," & _
              " a.tipo  as proceso,c.ctsmon as moneda," & _
              "(select cant from pl_tareo where ANOMES='" & AnoMes & "' AND TIPO='" & plani & "'" & _
              " AND rub='121' and unidad='D' and emp=a.emp) *" & _
              " (sum(IF( a.rub ='121', " & str & ") + " & _
              " sum( IF( a.rub ='1001', a.cant, 0 )) + sum( IF( a.rub ='1002', a.cant, 0 )) + sum( IF( a.rub ='1007', a.cant, 0 )) + sum( IF( a.rub ='1005', a.cant, 0 )) + sum( IF( a.rub ='201', a.cant, 0 )) + sum( IF( a.rub ='105', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='401', a.cant, 0 )))/360 AS ingresos," & _
              " sum( IF( a.rub ='703', a.cant, 0 )) as descuentos,a.fecha" & _
              " FROM pl_tareo as a LEFT JOIN" & _
              " empleado  AS c ON (a.EMP=c.CODIGO)" & _
              " WHERE ctsbanco='09' and a.ANOMES='" & AnoMes & "' AND a.TIPO='" & plani & "'" & _
              " and c.ctsmon='" & mon & "' " & _
              " GROUP BY a.anomes,a.emp,a.tipo,c.ctsmon,a.fecha ORDER BY C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2"
    Else
        SQL = " SELECT a.anomes,a.emp,c.tipcta_mn,c.numcta_mn,c.tipcta_me,c.numcta_me," & _
              " CONCAT(C.APEPAT,' ',C.APEMAT,' ',C.NOMBRE1,' ',C.NOMBRE2) AS NOMBRE," & _
              " a.tipo as proceso,a.moneda,sum( IF( a.rub ='121', " & str & " , 0 )) + " & _
              " sum( IF( a.rub ='1001', a.cant, 0 )) + sum( IF( a.rub ='1005', a.cant, 0 )) + sum( IF( a.rub ='1008', a.cant, 0 )) +  sum( IF( a.rub ='1006', a.cant, 0 ))  + sum( IF( a.rub ='1007', a.cant, 0 ))  + " & _
              " sum( IF( a.rub ='1002', a.cant, 0 )) + sum( IF( a.rub ='1003', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='102', a.cant, 0 )) + sum( IF( a.rub ='103', a.cant, 0 )) + sum( IF( a.rub ='105', a.cant, 0 )) +  sum( IF( a.rub ='107', a.cant, 0 )) +" & _
              " sum( IF( a.rub ='115', a.cant, 0 )) + sum( IF( a.rub ='116', a.cant, 0 )) + sum( IF( a.rub ='118', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='117', " & str & " , 0 )) + " & _
              " sum( IF( a.rub ='201', a.cant, 0 )) + sum( IF( a.rub ='202', a.cant, 0 )) + sum( IF( a.rub ='313', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='910', a.cant, 0 )) + sum( IF( a.rub ='911', a.cant, 0 )) + sum( IF( a.rub ='902', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='915', a.cant, 0 )) + sum( IF( a.rub ='904', a.cant, 0 )) + sum( IF( a.rub ='909', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='916', a.cant, 0 )) + sum( IF( a.rub ='401', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='403', a.cant, 0 )) + sum( IF( a.rub ='504', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='404', a.cant, 0 )) + sum( IF( a.rub ='407', a.cant, 0 )) AS ingresos," & _
              " sum( IF( a.rub ='601', a.cant, 0 )) + sum( IF( a.rub ='604', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='605', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='607', a.cant, 0 )) + sum( IF( a.rub ='608', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='606', a.cant, 0 )) + sum( IF( a.rub ='611', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='701', a.cant, 0 )) + sum( IF( a.rub ='703', a.cant, 0 )) + sum( IF( a.rub ='704', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='705', a.cant, 0 )) + sum( IF( a.rub ='706', a.cant, 0 )) + sum( IF( a.rub ='707', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='708', a.cant, 0 )) + sum( IF( a.rub ='709', a.cant, 0 )) + " & _
              " sum( IF( a.rub ='710', a.cant, 0 )) + sum( IF( a.rub ='711', a.cant, 0 )) +  sum( IF( a.rub ='712', a.cant, 0 )) +  sum( IF( a.rub ='713', a.cant, 0 )) +  sum( IF( a.rub ='714', a.cant, 0 ))  as Descuentos,  a.fecha " & _
              " FROM pl_tareo as a left join empleado  as c ON (a.emp=c.codigo) WHERE a.anomes='" & AnoMes & "'" & _
              " AND a.tipo='" & plani & "' and a.moneda='" & mon & "' " & _
              " GROUP BY a.anomes,a.emp,a.tipo,a.moneda,a.fecha ORDER BY C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2"
    End If
    oConexionMYSQL.Execute SQL
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        Set myrs = Nothing
        MsgBox "No existe movimiento registrado esta fecha", vbOKOnly + vbInformation, "NOVADMIN"
        Exit Sub
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_InsiEntSal(fec As String)
    Dim SQL As String, num As Integer, Cant As Integer
    Dim myrs As MYSQL_RS, rs_Horas As MYSQL_RS
    Set myrs = New MYSQL_RS
    oConexionMYSQL.Execute ("delete from RH_TMPINSIDENCIASES where usuario='" & strUsuarioId & "'")
    SQL = " INSERT INTO RH_TMPINSIDENCIASES SELECT (select nombre from rh_estacionestrabajo where codigo= A.SEDE ) as sede," & _
          " (select concat(apepat,' ',apemat,' ',nombre1,' ',nombre2) from empleado where codigo=a.emp) as nomemp, a.emp,a.fecha," & _
          " '00:00:00' as e1,'00:00:00' as s1," & _
          " '00:00:00' as e2,'00:00:00' as s2," & _
          " '00:00:00' as e3,'00:00:00' as s3," & _
          " '00:00:00' as e4,'00:00:00' as s4," & _
          " '00:00:00' as e5,'00:00:00' as s5," & _
          " '00:00:00' as e6,'00:00:00' as s6," & _
          " '00:00:00' as e7,'00:00:00' as s7," & _
          " '00:00:00' as e8,'00:00:00' as s8,'" & strUsuarioId & "' as usuario,''" & _
          " FROM rh_entsalempleado as a  left join contrato as c" & _
          " on (a.emp=c.codemp ) WHERE a.FECHA='" & fec & "' GROUP BY a.emp,a.fecha" & _
          " ORDER BY a.emp,a.fecha"
    oConexionMYSQL.Execute SQL
    SQL = "Select * from RH_TMPINSIDENCIASES where usuario='" & strUsuarioId & "' order by emp"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        Set myrs = Nothing
        MsgBox "No existe movimiento registrado esta fecha", vbOKOnly + vbInformation, "NOVADMIN"
        Exit Sub
    End If
    SQL = "Select * from rh_entsalempleado where fecha='" & fec & "' order by emp,hor"
    Set rs_Horas = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount > 0 Then
        Do While Not myrs.EOF
            num = 0
            While rs_Horas.Fields("emp") = myrs.Fields("emp")
                If rs_Horas.Fields("tipo") = "E" Then
                    num = num + 1
                    SQL = "update rh_tmpinsidenciases set e" & Trim(CStr(num)) & "='" & Format(rs_Horas.Fields("hor")) & "'" & _
                          " where emp='" & myrs.Fields("emp") & "' and fecha='" & fec & "' and usuario='" & strUsuarioId & "'"
                    oConexionMYSQL.Execute SQL
                End If
                If rs_Horas.Fields("tipo") = "S" Then
                    SQL = "update rh_tmpinsidenciases set s" & Trim(CStr(num)) & "='" & Format(rs_Horas.Fields("hor")) & "'" & _
                          " where emp='" & myrs.Fields("emp") & "' and fecha='" & fec & "' and usuario='" & strUsuarioId & "'"
                    oConexionMYSQL.Execute SQL
                End If
                rs_Horas.MoveNext
            Wend
            myrs.MoveNext
        Loop
        myrs.CloseRecordset
        SQL = "Select * from rh_tmpinsidenciases where fecha='" & fec & "' and usuario='" & strUsuarioId & "' order by CONCAT(RIGHT(e1,4),left(e1,8)),emp"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_InsiFecha(Mes As String, empleado As String, sede As String)
    Dim SQL As String, num As Integer, Cant As Integer
    Dim myrs As MYSQL_RS, rs_Horas As MYSQL_RS
    Set myrs = New MYSQL_RS
    oConexionMYSQL.Execute ("delete from RH_TMPINSIDENCIASES where usuario='" & strUsuarioId & "'")
    SQL = " INSERT INTO RH_TMPINSIDENCIASES SELECT (select nombre from rh_estacionestrabajo where codigo= a.sede) as sede," & _
          " (select concat(apepat,' ',apemat,' ',nombre1,' ',nombre2) from empleado where codigo=a.emp ) as nomemp, a.emp,a.fecha," & _
          " '00:00:00' as e0,'00:00:00' as s0,'00:00:00' as e1,'00:00:00' as s1," & _
          " '00:00:00' as e2,'00:00:00' as s2," & _
          " '00:00:00' as e3,'00:00:00' as s3," & _
          " '00:00:00' as e4,'00:00:00' as s4," & _
          " '00:00:00' as e5,'00:00:00' as s5," & _
          " '00:00:00' as e6,'00:00:00' as s6," & _
          " '00:00:00' as e7,'00:00:00' as s7," & _
          " '00:00:00' as e8,'00:00:00' as s8,'" & strUsuarioId & "' as usuario,''" & _
          " FROM rh_entsalempleado as a  left join contrato as c" & _
          " on (a.emp=c.codemp ) WHERE a.fecha<>'x' "
    If Mes <> "" Then SQL = SQL & " and left(a.FECHA,7)='" & strAnoSistema & "/" & Mes & "'"
    If empleado <> "" Then SQL = SQL & " and a.emp='" & empleado & "'"
    If sede <> "" Then SQL = SQL & " and a.sede='" & sede & "'"
    SQL = SQL & " GROUP BY a.emp,a.fecha ORDER BY sede,a.emp,a.fecha"
    oConexionMYSQL.Execute SQL
    SQL = "Select * from RH_TMPINSIDENCIASES where usuario='" & strUsuarioId & "' order by emp"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        Set myrs = Nothing
        MsgBox "No existe movimiento registrado esta fecha", vbOKOnly + vbInformation, "NOVADMIN"
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        Do While Not myrs.EOF
            num = 0
            SQL = "Select * from rh_entsalempleado where emp='" & myrs.Fields("emp") & "' and fecha='" & myrs.Fields("fecha") & "' order by emp,hor"
            Set rs_Horas = oConexion.EjecutaSelectRS(SQL)
            While rs_Horas.Fields("emp") = myrs.Fields("emp")
                If rs_Horas.Fields("tipo") = "E" Then
                    num = num + 1
                    SQL = "update rh_tmpinsidenciases set e" & Trim(CStr(num)) & "='" & Format(rs_Horas.Fields("hor")) & "'" & _
                          " where emp='" & myrs.Fields("emp") & "' and fecha='" & myrs.Fields("fecha") & "' and usuario='" & strUsuarioId & "'"
                    oConexionMYSQL.Execute SQL
                End If
                If rs_Horas.Fields("tipo") = "S" Then
                    SQL = "update rh_tmpinsidenciases set s" & Trim(CStr(num)) & "='" & Format(rs_Horas.Fields("hor")) & "'" & _
                          " where emp='" & myrs.Fields("emp") & "' and fecha='" & myrs.Fields("fecha") & "' and usuario='" & strUsuarioId & "'"
                    oConexionMYSQL.Execute SQL
                End If
                rs_Horas.MoveNext
            Wend
            myrs.MoveNext
        Loop
        myrs.CloseRecordset
        
        Dim dia As Integer, MesSal As Integer
        SQL = "Select distinct sede,nomemp,emp from RH_TMPINSIDENCIASES where usuario='" & strUsuarioId & "' order by emp"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
        Do While Not myrs.EOF()
            SQL = "select fec_salida,fec_regreso,observacion from calendario d where d.codemp='" & myrs.Fields("emp") & "' " & _
                  "and movemp in ('03','05') order by fec_salida"
            Set rs_Horas = oConexion.EjecutaSelectRS(SQL)
            Do While Not rs_Horas.EOF()
                dia = Day(rs_Horas.Fields("fec_salida"))
                MesSal = Month(rs_Horas.Fields("fec_salida"))
                If Mes = MesSal Then
                    Do While dia <= Day(rs_Horas.Fields("fec_regreso"))
                        SQL = "INSERT INTO RH_TMPINSIDENCIASES(sede,nomemp,emp,fecha,e1,s1,e2,s2,e3,s3,e4,s4,e5," & _
                              "s5,e6,s6,e7,s7,e8,s8,usuario,OBS) values('" & myrs.Fields("sede") & "','" & myrs.Fields("nomemp") & "', " & _
                              "'" & myrs.Fields("emp") & "','" & strAnoSistema & "/" & Mes & "/" & Right("00" & dia, 2) & "','00:00:00','00:00:00'," & _
                              " '00:00:00','00:00:00','00:00:00','00:00:00','00:00:00','00:00:00','00:00:00','00:00:00'," & _
                              " '00:00:00','00:00:00','00:00:00','00:00:00','00:00:00','00:00:00','" & strUsuarioId & "','" & rs_Horas.Fields("observacion") & "')"
                        oConexionMYSQL.Execute SQL
                        dia = dia + 1
                    Loop
                End If
                rs_Horas.MoveNext
            Loop
            myrs.MoveNext
        Loop
        myrs.CloseRecordset
        SQL = "Select a.sede,a.nomemp,a.emp,a.fecha,a.e1,a.s1,a.e2," & _
              " a.s2,a.e3,a.s3,a.e4,a.s4,a.e5,a.s5,a.e6,a.s6,a.e7,a.s7,a.e8,a.s8,b.numdocide,c.area,c.descrip,d.descrip as categ,a.obs " & _
              " from ((rh_tmpinsidenciases as a" & _
              " left join empleado as b on (a.emp=b.codigo)) left join cncargos as c on " & _
              " (b.codcargo=c.codigo )) left join rh_categoria as d on (b.categoria=d.codigo) where a.usuario='" & strUsuarioId & "' order by a.e1,a.emp"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
        If empleado <> "" Then
            m_Sede = Trim(myrs.Fields("sede"))
        End If
        m_Nombre1 = "R.U.C.: " & strRucEmpresa
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
    Set rs_Horas = Nothing
End Sub
Public Sub sp_Rep_AsistenciaxDH(Mes As String, empleado As String)
    Dim SQL As String, mesanterior As String, anioanterior As String
    Dim diames As String, diamesanterior As String
    diames = Right("00" & Trim(str(DiasDelMes(Mes))), 2)
    mesanterior = IIf(Mes = "01", "12", Right("00" & Trim(str(val(Mes) - 1)), 2))
    diamesanterior = Right("00" & Trim(str(DiasDelMes(mesanterior))), 2)
    anioanterior = IIf(Mes = "01", str(val(strAnoSistema) - 1), strAnoSistema)
    Dim myrs As MYSQL_RS
    SQL = "Select a.emp,concat_Ws(' ',b.Apepat,b.apemat,b.nombre1,b.nombre2) as nombre," & _
          " count(*) as dias," & _
          " sum(if(a.tiposede='O',if(DATE_FORMAT(a.fecha, '%W')<>'Saturday',1,0),0)) as DiasO," & _
          " sum(if(a.tiposede='O',if(DATE_FORMAT(a.fecha, '%W')='Saturday',1,0),0)) as DiasSO," & _
          " sum(if(a.tiposede='O',if(DATE_FORMAT(a.fecha, '%W')='Saturday',4,IF(b.tipo = 3,7,8)),0)) as HorasO," & _
          " sum(if(a.tiposede='C',1,0)) as DiasC, sum(if(a.tiposede='C',11,0)) as HorasC," & _
          " DATEDIFF(LAST_DAY('" & strAnoSistema & "/" & Mes & "/" & diames & "'),'" & strAnoSistema & "/" & Mes & "/01')+1 - sum(if(a.tiposede='C',1,0)) " & _
          " - sum(if(a.tiposede='O',1,0)) - ifnull(vac.dias,0) as Descanso,ifnull(vac.dias, 0) As vac,b.tipo" & _
          " FROM (rh_entsalempleado as a left join empleado as b " & _
          " on (a.emp=b.codigo)) left JOIN (Select a.codigo," & _
          " c.fec_Salida,c.fec_Regreso, (DATEDIFF(c.fec_Regreso,c.fec_Salida) - " & _
          " if(c.fec_salida<='" & anioanterior & "/" & mesanterior & "/" & diamesanterior & "', DATEDIFF('" & anioanterior & "/" & mesanterior & "/" & diamesanterior & "',c.fec_salida)+1,0) - " & _
          " if(c.fec_Regreso>='" & strAnoSistema & "/" & Mes & "/" & diames & "', DATEDIFF(c.fec_Regreso,'" & strAnoSistema & "/" & Mes & "/" & diames & "'),0) + 1 ) as dias" & _
          " from empleado as a left join calendario as c  on(c.codemp=a.codigo)" & _
          " where a.tipo not in (3,4) and c.movemp='02' and gocehaber='N' and" & _
          " concat(left(fec_salida,4), substring(fec_salida,6,2))<='" & strAnoSistema & Mes & "' and" & _
          " concat(left(c.fec_regreso,4), substring(c.fec_regreso,6,2))>='" & strAnoSistema & Mes & "') as vac" & _
          " on (vac.codigo=a.emp)" & _
          " where a.fecha>='" & strAnoSistema & "/" & Mes & "/01' and a.fecha<='" & strAnoSistema & "/" & Mes & "/" & diames & "' and a.tipo='E'" & _
          " group by b.Apepat,b.apemat,b.nombre1,b.nombre2"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_AsistenciaxDHCampo(f1 As String, f2 As String, empleado As String)
    Dim sqlc As String, mesanterior As String, anioanterior As String
    Dim diames As String, diamesanterior As String
    Dim myrs As MYSQL_RS
    Dim str As String, Fingreso As String, FCese As String
    Dim TF1 As String, TF2 As String
    If empleado <> "" Then
        str = " and b.codigo='" & empleado & "'"
    Else
        str = ""
    End If
    TF1 = f1
    TF2 = f2
    sqlc = "delete from rh_tmpasisxdhcampo where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute sqlc
    sqlc = "insert into rh_tmpasisxdhcampo(division,situacion,codemp,nombres,diasper,dias,diaso,diasso,horaso,diasc,horasc, " & _
           "descanso,vac,descansoobli,usuario) Select c.division,b.situacion,a.emp,concat_Ws(' ',b.Apepat,b.apemat,b.nombre1,b.nombre2) as nombre," & _
           " DATEDIFF(if(if(b.fec_cese>='" & f1 & "' and b.fec_cese <= '" & f2 & "',b.fec_cese,'" & f2 & "') <'" & f2 & "',b.fec_cese,'" & f2 & "'),if(if(b.fec_ingreso>='" & f1 & "' and b.fec_ingreso <= '" & f2 & "',b.fec_ingreso,'" & f1 & "') > '" & f1 & "',b.fec_ingreso,'" & f1 & "'))+1 as diasper,count(*) as dias," & _
           " sum(if(a.tiposede='O',if(DATE_FORMAT(a.fecha, '%W')<>'Saturday',1,0),0)) as DiasO," & _
           " sum(if(a.tiposede='O',if(DATE_FORMAT(a.fecha, '%W')='Saturday',1,0),0)) as DiasSO," & _
           " sum(if(a.tiposede='O',if(DATE_FORMAT(a.fecha, '%W')='Saturday',4,8),0)) as HorasO," & _
           " sum(if(a.tiposede='C',1,0)) as DiasC, sum(if(a.tiposede='C',12,0)) as HorasC," & _
           " DATEDIFF(if(if(b.fec_cese>='" & f1 & "' and b.fec_cese <= '" & f2 & "',b.fec_cese,'" & f2 & "') <'" & f2 & "',b.fec_cese,'" & f2 & "'),if(if(b.fec_ingreso>='" & f1 & "' and b.fec_ingreso <= '" & f2 & "',b.fec_ingreso,'" & f1 & "') > '" & f1 & "',b.fec_ingreso,'" & f1 & "'))+1 - sum(if(a.tiposede='C',1,0)) " & _
           " - sum(if(a.tiposede='O',1,0)) - ifnull(vac.dias,0) as Descanso,ifnull(vac.dias, 0) As vac," & _
           " 154*(sum(if(a.tiposede='O',if(DATE_FORMAT(a.fecha, '%W')='Saturday',4,8),0))+sum(if(a.tiposede='C',12,0)))/231 as DescansoObli,'" & strUsuarioId & "' as usu " & _
           " FROM (rh_entsalempleado as a left join empleado as b " & _
           " on (a.emp=b.codigo)) left JOIN rh_tmptotvac as vac" & _
           " on (vac.emp=a.emp) left join contrato c on (c.codemp=b.codigo)" & _
           " where c.estado = 'AP' and a.fecha>='" & f1 & "' and a.fecha<='" & f2 & "' and a.tipo='E' and (vac.usuario='" & strUsuarioId & "' or vac.usuario is null) " & str & _
           " group by b.Apepat,b.apemat,b.nombre1,b.nombre2 HAVING sum(if(a.tiposede='C',1,0)) <>0"
    TotVacaciones IIf(f1 = "__/__/____", Date, f1), IIf(f2 = "__/__/____", Date, f2)
    oConexionMYSQL.Execute sqlc
    sqlc = "select * from rh_tmpasisxdhcampo where usuario = '" & strUsuarioId & "'"
    Set myrs = oConexion.EjecutaSelectRS(sqlc)
    Do While Not myrs.EOF
        Fingreso = Format(FechaPersonal(myrs.Fields("codemp")), "yyyy/mm/dd")
        FCese = Format(FechaPersonal(myrs.Fields("codemp"), 2), "yyyy/mm/dd")
        sqlc = "update rh_tmpasisxdhcampo SET fecingreso = '" & Fingreso & "',feccese = '" & FCese & "' where codemp = '" & myrs.Fields("codemp") & "' and usuario = '" & strUsuarioId & "'"
        oConexionMYSQL.Execute sqlc
        myrs.MoveNext
    Loop
    sqlc = "select division,situacion,codemp,nombres, " & _
           "DATEDIFF(if(if(feccese>='" & TF1 & "' and feccese <= '" & TF2 & "',feccese,'" & TF2 & "') <'" & TF2 & "',feccese,'" & TF2 & "'),if(if(fecingreso>='" & TF1 & "' and fecingreso <= '" & TF2 & "',fecingreso,'" & TF1 & "') > '" & TF1 & "',fecingreso,'" & TF1 & "'))+1 as diasper, " & _
           "dias,diaso,diasso,horaso,diasc,horasc, " & _
           "descanso,vac,descansoobli from rh_tmpasisxdhcampo where usuario = '" & strUsuarioId & "'"
    Set myrs = oConexion.EjecutaSelectRS(sqlc)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_ContratosTrab(CodEmp As String)
    Dim SQL As String, str As String
    Dim myrs As MYSQL_RS
        If CodEmp <> "" Then
        str = " and c.codemp='" & CodEmp & "'"
    Else
        str = ""
    End If
    SQL = "SELECT e.Codigo,Concat(ApePat, ' ', ApeMat, ' ', Nombre1, ' ', NOMBRE2) AS Nombres, " & _
          " IF(e.Situacion=0,'CESADO','ACTIVO'),e.Fec_Ingreso,e.Fec_Cese,c.codtipo, " & _
          " (select descrip from cncontrato where codigo = codtipo) AS descrip,c.bono,c.monto_bono, " & _
          " c.sbasico,0 as sbruto,date_format(c.f_inicio,'%d/%m/%Y'),date_format(c.f_termino,'%d/%m/%Y'),c.estado " & _
          " From empleado e INNER JOIN contrato c ON (e.codigo=c.codemp) where e.tipo not in (3,4) " & str
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_VacacionesTrab(CodEmp As String, fec As String, Division As String)
    Dim SQL As String, str As String, Str1 As String, str2 As String, per As String
    Dim myrs As MYSQL_RS
    Dim Ces As Boolean, FlgSaldo As Boolean
    oConexionMYSQL.Execute "Delete from tmpvac "
    If CodEmp <> "" Then
        str = " and a.codigo='" & CodEmp & "'"
    Else
        str = ""
    End If
    Str1 = ""
    FlgSaldo = True
    'If CodEmp = "" Then
    '    If MsgBox("¿Desea incluir a los Cesados?", vbQuestion + vbYesNo, "NOVBRANDT") = vbYes Then Ces = True
    'End If
    'If Ces = False Then Str1 = " and a.situacion = 1 "
    If Division <> "0006" And Division <> "" And Division <> "0000" Then
        str2 = " and c.division = '" & Division & "' "
    End If
    SQL = "Insert into tmPvac Select a.codigo,a.fec_cese,a.situacion," & _
           " CONCAT_WS(' ',a.apepat,a.apemat,a.nombre1,a.nombre2) as nombre,a.fec_ingreso," & _
           " b.fec_salida,b.fec_regreso,b.periodo,b.gocehaber,b.monto_viatico," & _
           " truncate(DATEDIFF(if(a.fec_cese='','" & fec & "',a.fec_cese),a.fec_ingreso)/360,0)*30 as diasVac, " & _
           " (select descriplocal from cnmdepar dp where dp.coddep = c.division) as division " & _
           " from (empleado as a left join  contrato c on(a.codigo=c.codemp)) " & _
           " left join (select * from calendario where movemp='02')  as b on (a.codigo=b.codemp) " & _
           " where a.tipo not in (3,4) and b.movemp='02' and c.codigo = (select max(codigo) from contrato t where t.codemp=a.codigo group by t.codemp) " & _
           " and b.fec_salida >= a.fec_ingreso " & str & Str1 & str2 & _
           " UNION Select a.codigo,a.fec_cese,a.situacion," & _
           " CONCAT_WS(' ',a.apepat,a.apemat,a.nombre1,a.nombre2) as nombre,'' as fec_ingreso," & _
           " '' as fec_salida,'' as fec_regreso,b.periodo,'P' as gocehaber,(30-sum(b.monto_viatico)) as mto_viatico," & _
           " truncate(DATEDIFF(if(a.fec_cese='','" & fec & "',a.fec_cese),a.fec_ingreso)/360,0)*30 as diasVac, " & _
           " (select descriplocal from cnmdepar dp where dp.coddep = c.division) as division " & _
           " from (empleado as a left join  contrato c on(a.codigo=c.codemp)) " & _
           " left join (select * from calendario where movemp='02')  as b on (a.codigo=b.codemp) " & _
           " where a.tipo not in (3,4) and b.movemp='02' and c.codigo = (select max(codigo) from contrato t where t.codemp=a.codigo group by t.codemp) " & _
           " and b.fec_salida >= a.fec_ingreso " & str & Str1 & str2 & " GROUP BY b.codemp,b.periodo having mto_viatico > 0 " & _
           " order by nombre,periodo"
   oConexionMYSQL.Execute SQL
   SQL = "select t.codigo,t.fec_cese,t.situacion,t.nombre,t.fec_ingreso,max(t.periodo) as periodo ,sum(t.monto_viatico) as dperiodo," & _
        " (30 - sum(t.monto_viatico)) as fperiodo," & _
        " (SELECT diasvac -SUM(MONTO_VIATICO) FROM `tmpvac` where codigo=t.codigo) as saldo,t.diasvac" & _
        " from tmpvac as t" & _
        " where t.periodo=(select max(periodo) from `tmpvac` where codigo=t.codigo )" & _
        " group by t.codigo"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    Do While Not myrs.EOF
        If myrs.Fields("fperiodo") > 0 Then
            SQL = "Insert into tmpvac (codigo,fec_cese,situacion,nombre,fec_ingreso," & _
                  " periodo,gocehaber,monto_viatico,diasvac) values (" & _
                  "'" & myrs.Fields("codigo") & "'," & _
                  "'" & myrs.Fields("fec_cese") & "'," & _
                  "'" & myrs.Fields("situacion") & "'," & _
                  "'" & myrs.Fields("nombre") & "'," & _
                  "'" & myrs.Fields("fec_ingreso") & "'," & _
                  "'" & myrs.Fields("periodo") & "'," & _
                  "'P'," & _
                  "" & myrs.Fields("fperiodo") & "," & _
                  "" & myrs.Fields("diasvac") & ")"
              oConexionMYSQL.Execute SQL
        End If
        If myrs.Fields("saldo") - myrs.Fields("fperiodo") = 30 Then
            per = Right(myrs.Fields("periodo"), 4) & "-" & CStr(val(Right(myrs.Fields("periodo"), 4)) + 1)
            SQL = "Insert into tmpvac (codigo,fec_cese,situacion,nombre,fec_ingreso," & _
                  " periodo,gocehaber,monto_viatico,diasvac) values (" & _
                  "'" & myrs.Fields("codigo") & "'," & _
                  "'" & myrs.Fields("fec_cese") & "'," & _
                  "'" & myrs.Fields("situacion") & "'," & _
                  "'" & myrs.Fields("nombre") & "'," & _
                  "'" & myrs.Fields("fec_ingreso") & "'," & _
                  "'" & per & "'," & _
                  "'P'," & _
                  "30," & _
                  "" & myrs.Fields("diasvac") & ")"
              oConexionMYSQL.Execute SQL
        End If
        If myrs.Fields("saldo") - myrs.Fields("fperiodo") = 60 Then
            per = Right(myrs.Fields("periodo"), 4) & "-" & CStr(val(Right(myrs.Fields("periodo"), 4)) + 1)
            SQL = "Insert into tmpvac (codigo,fec_cese,situacion,nombre,fec_ingreso," & _
                  " periodo,gocehaber,monto_viatico,diasvac) values (" & _
                  "'" & myrs.Fields("codigo") & "'," & _
                  "'" & myrs.Fields("fec_cese") & "'," & _
                  "'" & myrs.Fields("situacion") & "'," & _
                  "'" & myrs.Fields("nombre") & "'," & _
                  "'" & myrs.Fields("fec_ingreso") & "'," & _
                  "'" & per & "'," & _
                  "'P'," & _
                  "30," & _
                  "" & myrs.Fields("diasvac") & ")"
              oConexionMYSQL.Execute SQL
              per = CStr(val(Right(myrs.Fields("periodo"), 4)) + 1) & "-" & CStr(val(Right(myrs.Fields("periodo"), 4)) + 2)
              SQL = "Insert into tmpvac (codigo,fec_cese,situacion,nombre,fec_ingreso," & _
                  " periodo,gocehaber,monto_viatico,diasvac) values (" & _
                  "'" & myrs.Fields("codigo") & "'," & _
                  "'" & myrs.Fields("fec_cese") & "'," & _
                  "'" & myrs.Fields("situacion") & "'," & _
                  "'" & myrs.Fields("nombre") & "'," & _
                  "'" & myrs.Fields("fec_ingreso") & "'," & _
                  "'" & per & "'," & _
                  "'P'," & _
                  "30," & _
                  "" & myrs.Fields("diasvac") & ")"
              oConexionMYSQL.Execute SQL
        End If
        If myrs.Fields("saldo") - myrs.Fields("fperiodo") = 90 Then
            per = Right(myrs.Fields("periodo"), 4) & "-" & CStr(val(Right(myrs.Fields("periodo"), 4)) + 1)
            SQL = "Insert into tmpvac (codigo,fec_cese,situacion,nombre,fec_ingreso," & _
                  " periodo,gocehaber,monto_viatico,diasvac) values (" & _
                  "'" & myrs.Fields("codigo") & "'," & _
                  "'" & myrs.Fields("fec_cese") & "'," & _
                  "'" & myrs.Fields("situacion") & "'," & _
                  "'" & myrs.Fields("nombre") & "'," & _
                  "'" & myrs.Fields("fec_ingreso") & "'," & _
                  "'" & per & "'," & _
                  "'P'," & _
                  "30," & _
                  "" & myrs.Fields("diasvac") & ")"
              oConexionMYSQL.Execute SQL
              per = CStr(val(Right(myrs.Fields("periodo"), 4)) + 1) & "-" & CStr(val(Right(myrs.Fields("periodo"), 4)) + 2)
              SQL = "Insert into tmpvac (codigo,fec_cese,situacion,nombre,fec_ingreso," & _
                  " periodo,gocehaber,monto_viatico,diasvac) values (" & _
                  "'" & myrs.Fields("codigo") & "'," & _
                  "'" & myrs.Fields("fec_cese") & "'," & _
                  "'" & myrs.Fields("situacion") & "'," & _
                  "'" & myrs.Fields("nombre") & "'," & _
                  "'" & myrs.Fields("fec_ingreso") & "'," & _
                  "'" & per & "'," & _
                  "'P'," & _
                  "30," & _
                  "" & myrs.Fields("diasvac") & ")"
              oConexionMYSQL.Execute SQL
              per = CStr(val(Right(myrs.Fields("periodo"), 4)) + 2) & "-" & CStr(val(Right(myrs.Fields("periodo"), 4)) + 3)
               SQL = "Insert into tmpvac (codigo,fec_cese,situacion,nombre,fec_ingreso," & _
                  " periodo,gocehaber,monto_viatico,diasvac) values (" & _
                  "'" & myrs.Fields("codigo") & "'," & _
                  "'" & myrs.Fields("fec_cese") & "'," & _
                  "'" & myrs.Fields("situacion") & "'," & _
                  "'" & myrs.Fields("nombre") & "'," & _
                  "'" & myrs.Fields("fec_ingreso") & "'," & _
                  "'" & per & "'," & _
                  "'P'," & _
                  "30," & _
                  "" & myrs.Fields("diasvac") & ")"
              oConexionMYSQL.Execute SQL
        End If
       myrs.MoveNext
    Loop
    myrs.CloseRecordset
    SQL = "Select * from tmpvac UNION Select a.codigo,a.fec_cese,a.situacion, CONCAT_WS(' ',a.apepat,a.apemat,a.nombre1,a.nombre2) as nombre, " & _
           "a.fec_ingreso,'' AS fec_salida,'' AS fec_regreso,'" & strAnoSistema - 1 & "-" & strAnoSistema & "' AS periodo,'P' AS gocehaber,'30' AS monto_viatico, " & _
           "truncate(DATEDIFF(if(a.fec_cese='','" & fec & "',a.fec_cese),a.fec_ingreso)/360,0)*30 as diasVac, " & _
           "(select descriplocal from cnmdepar dp where dp.coddep = c.division) as division " & _
           "from (empleado as a left join  contrato c on(a.codigo=c.codemp)) where a.tipo not in (3,4) and c.codigo = (select max(codigo) from contrato t " & _
           "where t.codemp=a.codigo group by t.codemp) AND A.CODIGO NOT IN (select O.CODEMP from calendario O where movemp='02') " & _
           "and a.situacion = 1 " & str & str2 & " HAVING DIASVAC > 0 order by nombre,periodo,fec_salida,gocehaber"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
       EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub

Public Sub sp_Rep_PermisosNOVTrab(CodEmp As String, fec As String, Division As String)
    Dim SQL As String, str As String, Str1 As String, str2 As String, per As String
    Dim myrs As MYSQL_RS
    
    If CodEmp <> "" Then
        str = " and a.codigo='" & CodEmp & "'"
    Else
        str = ""
    End If
    Str1 = ""
    
    If Division <> "0006" And Division <> "" And Division <> "0000" Then
        str2 = " and l.division = '" & Division & "' "
    End If
   
   SQL = "Select a.codigo,a.situacion,a.fec_ingreso,a.fec_cese,(select descriplocal from cnmdepar dp where dp.coddep = l.division) as division,concat_ws(' ',a.apepat,a.apemat, a.nombre1,a.nombre2) as nombre, a.codafp, " & _
           "c.fec_Salida,c.hora_salida,c.fec_Regreso,c.hora_regreso,DATEDIFF(c.fec_Regreso,c.fec_Salida) + 1 as dias, TIME_FORMAT( TIME(c.hora_regreso) - TIME(c.hora_salida) ,'%h:%i') as hora,c.observacion " & _
           "from (empleado as a left join calendario as c on(c.codemp=a.codigo)) left join contrato as l on l.codemp=a.codigo and l.estado='AP'  " & _
           "where a.tipo not in (3,4) and c.movemp='03' " & str & str2 & " order by a.apepat,a.apemat,a.nombre1,a.nombre2,fec_salida"
   Set myrs = oConexion.EjecutaSelectRS(SQL)
   If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
   End If
   If myrs.RecordCount > 0 Then
       EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
   End If
   Set myrs = Nothing
End Sub

Public Sub sp_Rep_LicenciasNOVTrab(CodEmp As String, fec As String, Division As String)
    Dim SQL As String, str As String, Str1 As String, str2 As String, per As String
    Dim myrs As MYSQL_RS
    
    If CodEmp <> "" Then
        str = " and a.codigo='" & CodEmp & "'"
    Else
        str = ""
    End If
    Str1 = ""
    
    If Division <> "0006" And Division <> "" And Division <> "0000" Then
        str2 = " and l.division = '" & Division & "' "
    End If
   
   SQL = "Select a.codigo,a.situacion,a.fec_ingreso,a.fec_cese,(select descriplocal from cnmdepar dp where dp.coddep = l.division) as division,concat_ws(' ',a.apepat,a.apemat, a.nombre1,a.nombre2) as nombre, a.codafp, " & _
           "c.fec_Salida,c.hora_salida,c.fec_Regreso,c.hora_regreso,DATEDIFF(c.fec_Regreso,c.fec_Salida) + 1 as dias, TIME_FORMAT( TIME(c.hora_regreso) - TIME(c.hora_salida) ,'%h:%i') as hora,c.observacion " & _
           "from (empleado as a left join calendario as c on(c.codemp=a.codigo)) left join contrato as l on l.codemp=a.codigo and l.estado='AP'  " & _
           "where a.tipo not in (3,4) and c.movemp='05' " & str & str2 & " order by a.apepat,a.apemat,a.nombre1,a.nombre2,fec_salida"
   Set myrs = oConexion.EjecutaSelectRS(SQL)
   If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
   End If
   If myrs.RecordCount > 0 Then
       EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
   End If
   Set myrs = Nothing
End Sub

Public Sub sp_Rep_SubsidiosNOVTrab(CodEmp As String, fec As String, Division As String)
    Dim SQL As String, str As String, Str1 As String, str2 As String, per As String
    Dim myrs As MYSQL_RS
    
    If CodEmp <> "" Then
        str = " and a.codigo='" & CodEmp & "'"
    Else
        str = ""
    End If
    Str1 = ""
    
    If Division <> "0006" And Division <> "" And Division <> "0000" Then
        str2 = " and l.division = '" & Division & "' "
    End If
   
   SQL = "Select a.codigo,a.situacion,a.fec_ingreso,a.fec_cese,(select descriplocal from cnmdepar dp where dp.coddep = l.division) as division,concat_ws(' ',a.apepat,a.apemat, a.nombre1,a.nombre2) as nombre, a.codafp, " & _
           "c.fec_Salida,c.hora_salida,c.fec_Regreso,c.hora_regreso,DATEDIFF(c.fec_Regreso,c.fec_Salida) + 1 as dias, TIME_FORMAT( TIME(c.hora_regreso) - TIME(c.hora_salida) ,'%h:%i') as hora,c.observacion " & _
           "from (empleado as a left join calendario as c on(c.codemp=a.codigo)) left join contrato as l on l.codemp=a.codigo and l.estado='AP'  " & _
           "where a.tipo not in (3,4) and c.movemp='07' " & str & str2 & " order by a.apepat,a.apemat,a.nombre1,a.nombre2,fec_salida"
   Set myrs = oConexion.EjecutaSelectRS(SQL)
   If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
   End If
   If myrs.RecordCount > 0 Then
       EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
   End If
   Set myrs = Nothing
End Sub


Public Sub sp_Rep_ExtornosProf(Cond As Long)
    Dim SQL As String, SQTemp As String
    Dim RQ As MYSQL_RS
    SQTemp = ""
    If Cond > 0 Then
        SQTemp = " AND CONVERT(SUBSTRING(V.VOUCHER,3,6),UNSIGNED) > " & Cond & ""
    End If
    SQL = "SELECT DISTINCT M.SERDOC,M.NUMDOC,V.VOUCHER,V.FECHA,V.GLOSA From cnmovi M INNER JOIN cnvouc V ON (M.anomes=V.anomes) " & _
          "AND (M.voucher=V.voucher) WHERE V.ANOMES = '" & strAnoSistema & strMesSistema & "' AND V.CODLIB = '05' AND " & _
          "(TIPO = 'P' OR V.GLOSA LIKE 'REVERSION DE PROFORMAS%') " & SQTemp & _
          " ORDER BY SERDOC,NUMDOC,V.VOUCHER"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If RQ.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set RQ = Nothing
        Exit Sub
    End If
    If RQ.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, RQ
    End If
    Set RQ = Nothing
End Sub
Public Sub sp_Rep_PrestamosAdelantos(Opc As String, Optional emp As String)
    Dim SQL As String, str As String
    Dim myrs As MYSQL_RS
    
    If emp <> "" Then
        str = "and e.codigo='" & emp & "'"
    End If
    
    SQL = "Select d.solicita,d.fec_reg,d.documento,d.fec_ini,d.fec_vcto,d.mon,d.importe, " & _
          "d.cuotas,d.montopagado,(importe - montopagado) as falta, " & _
          "concat(apepat,' ',apemat,' ',nombre1,' ',nombre2) as descrip " & _
          "from documentos_rrhh d inner join empleado e on d.solicita=e.codigo " & _
          "where coddoc='" & Opc & "' and montopagado < importe AND " & _
          "DATE_FORMAT(d.fec_ini,'%Y%m') <= '" & strAnoSistema & strMesSistema & "' " & str & _
          "order by descrip,d.fec_reg"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_HistPrestamosAdelantosTrab(CodEmp As String, TipPlani As String, MesPlani As String, TipDoc As String, Doc As String, Todo As String, mon As String)

    Screen.MousePointer = vbHourglass
    
    Dim SQL As String
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
     
    If CodEmp <> "" Then
        CodEmp = Right("00000000000" & CodEmp, 11)
    End If
     
    Set oRs = ADO_LlenaRs("call RH_SP_Prestamos_Reporte('HISTORIAL','" & CodEmp & "','','" & TipPlani & "','" & Doc & "','" & Todo & "');")

    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No existen datos en el historial de prestamos", vbOKOnly + vbInformation, gsNomSW
            Exit Sub
        End If

        m_Titulo = "HISTORIAL DE DESCUENTOS POR PRESTAMOS/ADELANTOS POR TRABAJADOR"
        
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, False
    Else
        MsgBox "No existen datos en el historial de prestamos", vbOKOnly + vbInformation, gsNomSW
    End If

    Set oRs = Nothing


End Sub
Public Sub sp_Rep_Asistencias()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select item,nombres,D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15,D16,D17,D18,D19,D20,D21,D22,D23," & _
          "D24,D25,D26,D27,D28,D29,D30,D31,TipoSede,codigo,sede from rh_tmpasistencias where usuario ='" & strUsuarioId & "' order by item,tiposede desc"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_TotalContratos(CodEmp As String, fecha As String, AniosIni As String, AniosFin As String)
    On Error GoTo SERROR
    
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim RQ As MYSQL_RS
    Dim SAct As String, SAnios As String
    Dim SWhere As String, SFecha As String
    Screen.MousePointer = vbHourglass
    SQL = "delete from rh_tmpespacioscontratos where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQL
    SQL = ""
    SAct = " and e.situacion = 1 "
    If CodEmp = "" Then
        If MsgBox("¿Desea incluir a los cesados?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then
            SAct = ""
        End If
    Else
        SAct = ""
    End If
    If CodEmp <> "" Then
        SWhere = " and e.codigo = '" & CodEmp & "'"
    End If
    SFecha = "sysdate()"
    If fecha <> "__/__/____" Then
        If IsDate(fecha) Then
            SFecha = "'" & Format(fecha, "yyyy/mm/dd") & "'"
        Else
            Exit Sub
        End If
    End If
    SQL = "delete from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQL
    If MsgBox("¿Desea ver el Total de Contratos desde el Primer Ingreso?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then
        SQL = "insert into rh_templiquidacion(codigo,nombres,mesgrati,anios,meses,aniosvac,fecingreso,mesesvac, " & _
              "feccese,motivo,fecultcts,usuario,bancocts,moncts) select e.codigo,concat_ws(' ',e.apepat,e.apemat, e.nombre1,e.nombre2) as nombre," & _
              "IF(IFNULL(sum(k.dias),0)=0,0,count(k.codigo)) as TotalContratos, " & _
              "truncate(IFNULL(sum(k.dias),0) / 365,0) as años, " & _
              "truncate(TRUNCATE(IFNULL(sum(k.dias),0) MOD 365,0) / 30,0) as MESES, " & _
              "TRUNCATE(TRUNCATE(IFNULL(sum(k.dias),0) MOD 365,0) MOD 30,0) AS DIAS, " & _
              "if(E.SITUACION=0,'CESADOS','ACTIVOS') AS situacion,e.tipo,o.division, " & _
              "(select descriplocal from cnmdepar d where d.coddep = o.division) as divi, " & _
              "date_format(min(k.f_inicio),'%d/%m/%Y') AS fec_ingreso,'" & strUsuarioId & "', " & _
              "(select descrip from cncargos g where g.codigo=e.codcargo) as cargo,if(SUM(K.DIAS) <> DATEDIFF(DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),date_format(min(k.f_inicio),'%Y/%m/%d'))+1,'H','') as huecos " & _
              "from empleado e inner join (select a.codigo,c.f_inicio,IF(IFNULL(c.f_termino,'')='' or IFNULL(c.f_termino,'') > DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),IFNULL(c.f_termino,'')) as ffin, " & _
              "if((DATEDIFF(IF(IFNULL(c.f_termino,'')='' or IFNULL(c.f_termino,'') > DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),IFNULL(c.f_termino,'')),c.f_inicio)+1)<0,0,(DATEDIFF(IF(IFNULL(c.f_termino,'')='' or IFNULL(c.f_termino,'')> DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),IFNULL(c.f_termino,'')),c.f_inicio)+1)) dias,c.division " & _
              "from empleado as a left join contrato as c on(c.codemp=a.codigo) order by a.apepat,a.apemat,a.nombre1,a.nombre2,c.codigo) K on (e.codigo=k.codigo) left join contrato o on (o.codemp=e.codigo) " & _
              "where e.tipo not in (3,4) and o.codigo = (select max(codigo) from contrato t where t.codemp=e.codigo) " & SWhere & SAct & " GROUP by e.codigo order by divi,e.apepat,e.apemat,e.nombre1,e.nombre2"
        oConexionMYSQL.Execute SQL
        oReporte.Titulo = "REPORTE DE TIEMPO DE SERVICIOS ACUMULADO AL " & Format(IIf(fecha = "__/__/____", Date, fecha), "dd/mm/yyyy")
    Else
        Screen.MousePointer = vbHourglass
        SQL = "insert into rh_tmpespacioscontratos(nombres,usuario) " & _
              "select distinct e.codigo,'" & strUsuarioId & "' from empleado e left join contrato as c on(c.codemp=e.codigo) " & _
              "where 1=1 " & SAct & SWhere & " order by e.codigo"
        oConexionMYSQL.Execute SQL
        SQL = "select nombres from rh_tmpespacioscontratos where usuario = '" & strUsuarioId & "' order by nombres"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        Do While Not RQ.EOF()
            SQL = "update rh_tmpespacioscontratos set estado = '" & FechaPersonal(RQ.Fields("nombres")) & "' " & _
                  "where usuario = '" & strUsuarioId & "' and nombres = '" & RQ.Fields("nombres") & "'"
            oConexionMYSQL.Execute SQL
            RQ.MoveNext
        Loop
        SQL = "insert into rh_templiquidacion(codigo,nombres,mesgrati,anios,meses,aniosvac,fecingreso,mesesvac, " & _
              "feccese,motivo,fecultcts,usuario,bancocts,moncts) select e.codigo,concat_ws(' ',e.apepat,e.apemat, e.nombre1,e.nombre2) as nombre," & _
              "IF(IFNULL(sum(k.dias),0)=0,0,count(k.codigo)) as TotalContratos, " & _
              "truncate(IFNULL(sum(k.dias),0) / 365,0) as años, " & _
              "truncate(TRUNCATE(IFNULL(sum(k.dias),0) MOD 365,0) / 30,0) as MESES, " & _
              "TRUNCATE(TRUNCATE(IFNULL(sum(k.dias),0) MOD 365,0) MOD 30,0) AS DIAS, " & _
              "if(E.SITUACION=0,'CESADOS','ACTIVOS') AS situacion,e.tipo,o.division, " & _
              "(select descriplocal from cnmdepar d where d.coddep = o.division) as divi," & _
              "date_format(e.fec_ingreso,'%d/%m/%Y') as fecingreso,'" & strUsuarioId & "', " & _
              "(select descrip from cncargos g where g.codigo=e.codcargo) as cargo, " & _
              "if(SUM(K.DIAS) <> DATEDIFF(DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),date_format(min(k.f_inicio),'%Y/%m/%d'))+1,'H','') as huecos from empleado e inner join (select a.codigo,c.f_inicio,IF(IFNULL(c.f_termino,'')='' or IFNULL(c.f_termino,'') > DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),IFNULL(c.f_termino,'')) as ffin, " & _
              "if((DATEDIFF(IF(IFNULL(c.f_termino,'')='' or IFNULL(c.f_termino ,'') > DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),IFNULL(c.f_termino,'')),c.f_inicio)+1)<0,0,(DATEDIFF(IF(IFNULL(c.f_termino,'')='' or IFNULL(c.f_termino,'')> DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),DATE_FORMAT(" & SFecha & ",'%Y/%m/%d'),IFNULL(c.f_termino,'')),c.f_inicio)+1)) dias,c.division,T.ESTADO " & _
              "from empleado as a left join rh_tmpespacioscontratos t on (t.nombres=a.codigo) and (t.usuario = '" & strUsuarioId & "') left join contrato as c on(c.codemp=a.codigo) HAVING (f_inicio >= estado) order by a.apepat,a.apemat,a.nombre1,a.nombre2,c.codigo) K on (e.codigo=k.codigo) left join contrato o on (o.codemp=e.codigo) " & _
              "where e.tipo not in (3,4) and e.fec_ingreso <= " & SFecha & " and o.codigo = (select max(codigo) from contrato t where t.codemp=e.codigo) " & SWhere & SAct & SAnios & " GROUP by e.codigo order by divi,e.apepat,e.apemat,e.nombre1,e.nombre2"
        oConexionMYSQL.Execute SQL
    End If
    SQL = "select codigo,fecultcts as fec_ingreso,moncts from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    Dim DifAnios As Integer, DifMeses As Integer, Difdias As Integer
    Dim FecTemp As String, FecFin As String, FecIni As String, RQ1 As MYSQL_RS
    Dim I As Integer, SumAnios As Integer, SumMeses As Integer, Sumdias As Integer
    Do While Not RQ.EOF()
        If RQ.Fields("moncts") = "H" Then
            SumAnios = 0: SumMeses = 0: Sumdias = 0: I = 0
            SQL = "select codigo,f_Inicio,f_termino,o.ult from contrato C LEFT JOIN (SELECT codemp,MAX(CODIGO) as ult FROM " & _
                  "CONTRATO  group by codemp) as o on (c.codemp=o.codemp) where C.codemp = '" & RQ.Fields("codigo") & "' order by codigo"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            Do While Not RQ1.EOF()
                If I = 0 Then
                    FecIni = CDate(RQ1.Fields("f_inicio"))
                    If IsDate(RQ1.Fields("f_termino")) Then
                        FecFin = CDate(RQ1.Fields("f_termino"))
                    Else
                        FecFin = CDate(Date)
                    End If
                    RQ1.MoveNext
                    I = I + 1
                Else
                    If CDate(Format(FecFin, "dd/mm/yyyy")) + 1 = CDate(Format(RQ1.Fields("f_inicio"), "dd/mm/yyyy")) Then
                        If RQ1.Fields("ult") = RQ1.Fields("codigo") Then
                            FecFin = Date
                            GoTo AQUI
                        End If
                        FecFin = CDate(RQ1.Fields("f_termino"))
                        RQ1.MoveNext
                        I = I + 1
                    Else
AQUI:
                        FecTemp = Mid(Format(FecIni, "dd/mm/yyyy"), 1, 6) & Year(FecFin)
                        DifAnios = Year(FecFin) - Year(CDate(Format(FecIni, "dd/mm/yyyy")))
                        DifMeses = 0: Difdias = 0
                        If CDate(FecTemp) > FecFin Then
                            DifAnios = DifAnios - 1
                            FecTemp = Mid(FecTemp, 1, 6) & (Year(CDate(FecTemp)) - 1)
                        End If
                        If Year(CDate(FecTemp)) < Year(FecFin) Then
                            DifMeses = (12 - Month(CDate(FecTemp))) + Month(FecFin)
                        ElseIf Year(CDate(FecTemp)) = Year(FecFin) Then
                            DifMeses = Month(FecFin) - Month(CDate(FecTemp))
                        End If
                        FecTemp = Mid(FecTemp, 1, 3) & Right("00" & Month(CDate(FecFin)), 2) & "/" & Year(FecFin)
                        If Day(FecFin) > Day(CDate(FecTemp)) Then
                            Difdias = Day(FecFin) - Day(CDate(FecTemp))
                        ElseIf Day(FecFin) < Day(CDate(FecTemp)) Then
                            DifMeses = DifMeses - 1
                            If Day(CDate(FecTemp)) = 31 Then
                                If Day(DateSerial(Year(CDate("01/" & Right("00" & Month(CDate(FecFin)) - 1, 2) & "/" & Year(FecFin))), Month(CDate("01/" & Right("00" & Month(CDate(FecFin)) - 1, 2) & "/" & Year(FecFin))) + 1, 0)) < Day(CDate(FecTemp)) Then
                                    FecTemp = "30/" & Right("00" & Month(CDate(FecFin)) - 1, 2) & "/" & Year(FecFin)
                                End If
                            Else
                                If Month(CDate(FecFin)) = 1 Then
                                    FecTemp = Mid(FecTemp, 1, 3) & "12" & "/" & Year(FecFin) - 1
                                Else
                                    FecTemp = Mid(FecTemp, 1, 3) & Right("00" & Month(CDate(FecFin)) - 1, 2) & "/" & Year(FecFin)
                                End If
                            End If
                            Difdias = (Day(DateSerial(Year(CDate(FecTemp)), Month(CDate(FecTemp)) + 1, 0)) - Day(CDate(FecTemp))) + Day(FecFin)
                        End If
                        SumAnios = SumAnios + DifAnios
                        SumMeses = SumMeses + DifMeses
                        Sumdias = Sumdias + Difdias
                        FecIni = CDate(RQ1.Fields("f_inicio"))
                        If FecFin <> Date Then
                            If IsDate(RQ1.Fields("f_termino")) Then
                                FecFin = CDate(RQ1.Fields("f_termino"))
                            Else
                                FecFin = Date
                                RQ1.MovePrevious
                            End If
                        End If
                        RQ1.MoveNext
                        I = I + 1
                    End If
                End If
            Loop
            If Sumdias >= 30 Then
                SumMeses = SumMeses + Int(val(Sumdias / 30))
                Sumdias = Sumdias - (30 * Int(val(Sumdias / 30)))
            End If
            If SumMeses >= 12 Then
                SumAnios = SumAnios + Int(val(SumMeses / 12))
                SumMeses = SumMeses - (12 * Int(val(SumMeses / 12)))
            End If
            SQL = "update rh_templiquidacion set anios = " & SumAnios & ",meses = " & SumMeses & ",aniosvac=" & Sumdias & " " & _
                  "where usuario = '" & strUsuarioId & "' and codigo = '" & RQ.Fields("codigo") & "'"
            oConexionMYSQL.Execute SQL
        Else
            Dim Fec1 As String
            If IsDate(fecha) Then
                Fec1 = Format(fecha, "dd/mm/yyyy")
            Else
                Fec1 = Format(Date, "dd/mm/yyyy")
            End If
            FecTemp = Mid(RQ.Fields("fec_ingreso"), 1, 6) & Year(Format(Fec1, "dd/mm/yyyy"))
            DifAnios = Year(Format(Fec1, "dd/mm/yyyy")) - Year(CDate(RQ.Fields("fec_ingreso")))
            DifMeses = 0: Difdias = 0
            If CDate(FecTemp) > Format(Fec1, "dd/mm/yyyy") Then
                DifAnios = DifAnios - 1
                FecTemp = Mid(FecTemp, 1, 6) & (Year(CDate(FecTemp)) - 1)
            End If
            If Year(CDate(FecTemp)) < Year(Format(Fec1, "dd/mm/yyyy")) Then
                DifMeses = (12 - Month(CDate(FecTemp))) + Month(Format(Fec1, "dd/mm/yyyy"))
            ElseIf Year(CDate(FecTemp)) = Year(Format(Fec1, "dd/mm/yyyy")) Then
                DifMeses = Month(Format(Fec1, "dd/mm/yyyy")) - Month(CDate(FecTemp))
            End If
            FecTemp = Mid(FecTemp, 1, 3) & Right("00" & Month(CDate(Format(Fec1, "dd/mm/yyyy"))), 2) & "/" & Year(Format(Fec1, "dd/mm/yyyy"))
            If Day(Format(Fec1, "dd/mm/yyyy")) > Day(CDate(FecTemp)) Then
                Difdias = Day(Format(Fec1, "dd/mm/yyyy")) - Day(CDate(FecTemp))
            ElseIf Day(Format(Fec1, "dd/mm/yyyy")) < Day(CDate(FecTemp)) Then
                DifMeses = DifMeses - 1
                If Day(CDate(FecTemp)) = 31 Then
                    If Day(DateSerial(Year(CDate("01/" & Right("00" & Month(CDate(Format(Fec1, "dd/mm/yyyy"))) - 1, 2) & "/" & Year(Format(Fec1, "dd/mm/yyyy")))), Month(CDate("01/" & Right("00" & Month(CDate(Format(Fec1, "dd/mm/yyyy"))) - 1, 2) & "/" & Year(Format(Fec1, "dd/mm/yyyy")))) + 1, 0)) < Day(CDate(FecTemp)) Then
                        FecTemp = "30/" & Right("00" & Month(CDate(Format(Fec1, "dd/mm/yyyy"))) - 1, 2) & "/" & Year(Format(Fec1, "dd/mm/yyyy"))
                    End If
                Else
                    FecTemp = Mid(FecTemp, 1, 3) & Right("00" & IIf(Month(CDate(Format(Fec1, "dd/mm/yyyy"))) - 1 = 0, 12, Month(CDate(Format(Fec1, "dd/mm/yyyy"))) - 1), 2) & "/" & IIf(Month(CDate(Format(Fec1, "dd/mm/yyyy"))) - 1 = 0, Year(Format(Fec1, "dd/mm/yyyy")) - 1, Year(Format(Fec1, "dd/mm/yyyy")))
                End If
                Difdias = (Day(DateSerial(Year(CDate(FecTemp)), Month(CDate(FecTemp)) + 1, 0)) - Day(CDate(FecTemp))) + Day(Format(Fec1, "dd/mm/yyyy"))
            End If
            SQL = "update rh_templiquidacion set anios = " & DifAnios & ",meses = " & DifMeses & ",aniosvac=" & Difdias & " " & _
                  "where usuario = '" & strUsuarioId & "' and codigo = '" & RQ.Fields("codigo") & "'"
            oConexionMYSQL.Execute SQL
        End If
        RQ.MoveNext
    Loop
    SAnios = ""
    If Trim(AniosIni) <> "" And Trim(AniosFin) <> "" Then
        SAnios = " and (anios >= " & val(AniosIni) & " and anios <= " & val(AniosFin) & ")"
    End If
    If Trim(AniosIni) <> "" And Trim(AniosFin) = "" Then
        SAnios = " and (anios = " & val(AniosIni) & ")"
    End If
    If Trim(AniosIni) = "" And Trim(AniosFin) <> "" Then
        SAnios = " and (anios = " & val(AniosFin) & ")"
    End If
    SQL = "select r.codigo,nombres,mesgrati as totalcontratos,anios as años,meses as meses,aniosvac as dias,fecingreso as situacion, " & _
          "mesesvac as tipo,feccese as division,motivo as divi,fecultcts as fec_ingreso,bancocts as cargo from rh_templiquidacion r " & _
          "left join contrato c on (r.codigo=c.codemp) where usuario = '" & strUsuarioId & "' and c.codigo = (select max(codigo) from contrato t where t.codemp=r.codigo group by t.codemp) " & _
          SAnios
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
    
    Exit Sub
SERROR:
    
    Screen.MousePointer = vbDefault
    Resume Next
    Set myrs = Nothing
End Sub
Public Sub sp_TotalPrestamosAdelantos(CodEmp As String, Tipo As String, mon As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim SWhere As String, STipo As String
    Dim rpt As String, Smon As String
    If CodEmp <> "" Then SWhere = " and e.codigo = '" & CodEmp & "'"
    If Tipo <> "" Then STipo = " and coddoc = '" & IIf(Tipo = "A", "SS", "PR") & "'"
    If mon <> "" Then Smon = " and k.mon = '" & mon & "'"
    rpt = InputBox("¿Desea ver el Reporte [D]etallado ó [R]esumen", "Préstamos y Adelantos", "D")
    If rpt = "D" Then
        SQL = "select e.codigo,concat_ws(' ',e.apepat,e.apemat, e.nombre1,e.nombre2) as nombre,importe,montopagado, " & _
              "concat(substring(r.descrip,5,LENGTH(r.descrip)-1),if(gratijulio='S',if(k.tipcuota<>4,' /GRATI.JUL',''),''),IF(gratidic='S',if(k.tipcuota<>5,' /GRATI.DIC',''),''),IF(uti='S',if(k.tipcuota<>6,' /UTIL.',''),''),IF(vaca='S',if(k.tipcuota<>3,' /VACAC.',''),'')) AS formapago, " & _
              "cuotas , montocuota, ASUNTO,if(CODDOC='PR','PRESTAMO','ADELANTO') AS coddoc,date_format(k.fec_reg,'%d/%m/%Y') as fec_reg from empleado e inner join documentos_rrhh K on (k.solicita=E.codigo) " & _
              "inner join rh_formapago r on (r.codigo=k.tipcuota) Where Importe > MONTOPAGADO" & SWhere & STipo & Smon & _
              " order by nombre,coddoc"
    Else
        SQL = "select e.codigo,concat_ws(' ',e.apepat,e.apemat, e.nombre1,e.nombre2) as nombre,count(*) as totpr, " & _
              "(SUM(IMPORTE)-sum(montopagado)) as saldo,sum(montocuota) as montocuota " & _
              "from empleado e inner join documentos_rrhh K on (k.solicita=E.codigo) Where Importe > MONTOPAGADO " & _
              SWhere & STipo & Smon & " group by e.codigo order by nombre,coddoc"
        
        Reporte = "Rep_TotalPrestamosAdelantosRes.rpt"
        Titulo = "SALDO DE " & IIf(Tipo <> "", IIf(Tipo = "P", "PRESTAMOS", "ADELANTOS"), "PRESTAMOS Y ADELANTOS") & " AL PERSONAL"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_TotalPrestamosAdelantosCierre(Mes As String, CodEmp As String, Tipo As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim SWhere As String, STipo As String, SMes As String
    Dim Flg As Boolean
    If Mes <> "" Then
        SMes = "where (r.liquid = 'N') and anomesplani <= '" & strAnoSistema & Mes & "'"
        Flg = True
    End If
    If CodEmp <> "" Then
        If Flg = True Then
            SWhere = " and solicita = '" & CodEmp & "'"
            Flg = True
        Else
            SWhere = " where (r.liquid = 'N') and solicita = '" & CodEmp & "'"
        End If
    End If
    If Tipo <> "" Then
        If Flg = True Then
            STipo = " and tipo = '" & Tipo & "'"
        Else
            STipo = " where (r.liquid = 'N') and tipo = '" & Tipo & "'"
        End If
    End If
    SQL = "select e.codigo,concat_ws(' ',e.apepat,e.apemat, e.nombre1,e.nombre2) as nombre," & _
          "count(k.documento) as totpr,sum(saldo) as Saldo,IF(sum(k.montocuota)>sum(saldo),sum(saldo),sum(k.montocuota)) as montocuota " & _
          "from ((empleado e inner join (select solicita,sum(montoabonado) as abonado,r.documento," & _
          "(select importe from `documentos_rrhh` d where (r.anomes=d.anomes) and (r.documento=d.documento) " & _
          "and (r.solicita=d.solicita) and (r.fecha=d.fec_reg)) as imp,((select importe from `documentos_rrhh` d " & _
          "where (r.anomes=d.anomes) and (r.documento=d.documento) and (r.solicita=d.solicita) and " & _
          "(r.fecha=d.fec_reg))- sum(montoabonado)) as saldo,(select montocuota from `documentos_rrhh` d " & _
          "where (r.anomes=d.anomes) and (r.documento=d.documento) and (r.solicita=d.solicita) and " & _
          "(r.fecha=d.fec_reg)) as montocuota from `rh_pagosemp` r " & SMes & SWhere & STipo & _
          " GROUP by solicita,r.documento having sum(montoabonado) < imp " & _
          "order by solicita) as K on (k.solicita=E.codigo))) GROUP by e.codigo order by nombre"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_ContratosCancelados(CodEmp As String, Tipo As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim SCodEmp As String, SFecha As String, SWhere As String, Str1 As String
    Dim SCond As String
    If CodEmp <> "" Then
        SCodEmp = " and e.codigo = '" & CodEmp & "'"
    End If
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SFecha = " and c.codigo=(select max(codigo) from contrato t where t.codemp=c.codemp group by codemp) AND ((situacion = 0 and fec_cese BETWEEN '" & Format(FechaIni, "yyyy/mm/dd") & "' and '" & Format(FechaFin, "yyyy/mm/dd") & "') or (c.f_termino between '" & Format(FechaIni, "yyyy/mm/dd") & "' and '" & Format(FechaFin, "yyyy/mm/dd") & "' and situacion = 1 and c.estado = 'CA'))"
        Str1 = " and (c.f_termino between '" & Format(FechaIni, "yyyy/mm/dd") & "' and '" & Format(FechaFin, "yyyy/mm/dd") & "') "
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SFecha = " and c.codigo=(select max(codigo) from contrato t where t.codemp=c.codemp group by codemp) AND ((situacion = 0 and fec_cese = '" & Format(FechaIni, "yyyy/mm/dd") & "') or (c.f_termino = '" & Format(FechaIni, "yyyy/mm/dd") & "' and situacion = 1 and c.estado = 'CA'))"
            Str1 = " and (c.f_termino = '" & Format(FechaIni, "yyyy/mm/dd") & "') "
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SFecha = " and c.codigo=(select max(codigo) from contrato t where t.codemp=c.codemp group by codemp) AND ((situacion = 0 and fec_cese = '" & Format(FechaFin, "yyyy/mm/dd") & "') or (c.f_termino = '" & Format(FechaFin, "yyyy/mm/dd") & "' and situacion = 1 and c.estado = 'CA'))"
                Str1 = " and (c.f_termino = '" & Format(FechaFin, "yyyy/mm/dd") & "') "
            Else
                SFecha = " and c.codigo=(select max(codigo) from contrato t where t.codemp=c.codemp group by codemp) AND ((situacion = 0 and fec_cese = DATE_FORMAT(sysdate(),'Y%/%m/%d')) or (c.f_termino = DATE_FORMAT(sysdate(),'Y%/%m/%d') and situacion = 1 and c.estado = 'CA'))"
                Str1 = " and (c.f_termino = date_format(sysdate(),'%Y/%m/%d')) "
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass
    If Tipo = "H" Then
        SQL = "SELECT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo," & _
              "IFNULL(date_format(fec_cese,'%d/%m/%Y'),'') as fec_cese,DATE_FORMAT(max(f_inicio),'%d/%m/%Y') as f_inicio,date_format(max(f_termino),'%d/%m/%Y') as f_termino,'NO RENOVADOS' as situacion,e.tipo, " & _
              "C.DIVGAS,(select descrip from cnmdepar r where r.coddep=c.divgas) as descripdiv FROM empleado e inner join contrato c on (e.codigo=c.codemp) where 1=1 " & SFecha & SCodEmp & "  " & _
              "group by e.codigo,e.fec_cese,situacion union " & _
              "SELECT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo," & _
              "IFNULL(date_format(fec_cese,'%d/%m/%Y'),'') as fec_cese,DATE_FORMAT((SELECT f_inicio from contrato a where a.codemp=c.codemp and a.codigo=(select max(codigo) from contrato t where t.codemp=a.codemp group by codemp)),'%d/%m/%Y') as f_inicio,date_format((SELECT f_termino from contrato a where a.codemp=c.codemp and a.codigo=(select max(codigo) from contrato t where t.codemp=a.codemp group by codemp)),'%d/%m/%Y') as f_termino,'RENOVADOS' as situacion,e.tipo, " & _
              "C.DIVGAS,(select descrip from cnmdepar r where r.coddep=c.divgas) as descripdiv FROM empleado e inner join contrato c on (e.codigo=c.codemp) where 1=1 and c.codigo=(select CONCAT('CN',RIGHT(CONCAT('00',CONVERT(SUBSTRING(max(codigo),3,2),UNSIGNED) - 1),2)) from contrato t where t.codemp=c.codemp " & _
              "group by codemp)" & Str1 & SCodEmp & " and fec_cese = '' group by e.codigo,e.fec_cese,situacion order by descripdiv,situacion,fec_cese,f_termino,nombres"
    Else
        SQL = "SELECT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo," & _
              "IFNULL(date_format(fec_cese,'%d/%m/%Y'),'') as fec_cese,DATE_FORMAT(max(f_inicio),'%d/%m/%Y') as f_inicio,date_format(max(f_termino),'%d/%m/%Y') as f_termino,'NO RENOVADOS' as situacion,e.tipo, " & _
              "C.DIVision,(select descriplocal from cnmdepar r where r.coddep=c.division) as descripdiv FROM empleado e inner join contrato c on (e.codigo=c.codemp) where 1=1 " & SFecha & SCodEmp & "  " & _
              "group by e.codigo,e.fec_cese,situacion union " & _
              "SELECT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo," & _
              "IFNULL(date_format(fec_cese,'%d/%m/%Y'),'') as fec_cese,DATE_FORMAT((SELECT f_inicio from contrato a where a.codemp=c.codemp and a.codigo=(select CONCAT('CN',RIGHT(CONCAT('00',CONVERT(SUBSTRING(max(codigo),3,2),UNSIGNED) - 1),2)) from contrato t where t.codemp=a.codemp group by codemp)),'%d/%m/%Y') as f_inicio,date_format((SELECT f_termino from contrato a where a.codemp=c.codemp and a.codigo=(select CONCAT('CN',RIGHT(CONCAT('00',CONVERT(SUBSTRING(max(codigo),3,2),UNSIGNED) - 1),2)) from contrato t where t.codemp=a.codemp group by codemp)),'%d/%m/%Y') as f_termino,'RENOVADOS' as situacion,e.tipo, " & _
              "C.DIVision,(select descriplocal from cnmdepar r where r.coddep=c.division) as descripdiv FROM empleado e inner join contrato c on (e.codigo=c.codemp) where 1=1 and c.codigo=(select CONCAT('CN',RIGHT(CONCAT('00',CONVERT(SUBSTRING(max(codigo),3,2),UNSIGNED) - 1),2)) from contrato t where t.codemp=c.codemp " & _
              "group by codemp) " & Str1 & SCodEmp & " and fec_cese = '' group by e.codigo,e.fec_cese,situacion order by descripdiv,situacion,fec_cese,f_termino,nombres"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Dependientes(CodEmp As String, Parentesco As String, Sexo As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim SCodEmp As String, SSexo As String, SParent As String
    If CodEmp <> "" Then SCodEmp = " and b.codemp = '" & CodEmp & "'"
    If Sexo <> "" Then SSexo = " and b.sexo = '" & Sexo & "'"
    If Parentesco <> "" Then SParent = " and b.codtipo = '" & Parentesco & "'"
    Screen.MousePointer = vbHourglass
    SQL = "Select a.codigo,CONCAT_WS(' ',A.APEPAT,A.APEMAT,A.NOMBRE1,A.NOMBRE2) AS NOMEMP, " & _
          "C.DESCRIP,d.descrip as td,B.NUMDOCIDE,B.NOMBRE,B.APEPATERNO,B.APEMATERNO,B.FEC_NAC, " & _
          "B.EDAD,B.Sexo,o.divgas,r.nombre as nomtrabajo,o.esttrabajo,if(ifnull(C.DESCRIP,'')<>'','PERSONAL CON DEPENDIENTES','PERSONAL SIN DEPENDIENTES') as dep,a.tipo " & _
          "FROM ((empleado AS A LEFT JOIN familiares AS B ON (A.CODIGO=B.CODEMP)) " & _
          "LEFT JOIN pl_vinculofam AS C ON (B.CODTIPO=C.CODIGO)) left join doc_identificacion as d " & _
          "On (B.CODDOCIDE=D.TIPO_DOC_IDE) left join contrato o on (a.codigo=o.codemp) " & _
          "left join rh_estacionestrabajo r on (o.esttrabajo=r.codigo) " & _
          "Where a.tipo not in (3,4) and a.situacion = 1 and o.estado = 'AP' " & SCodEmp & SSexo & SParent & _
          " ORDER BY dep,A.APEPAT,A.APEMAT,A.NOMBRE1,A.NOMBRE2,C.DESCRIP,B.EDAD"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub

Public Sub sp_Rep_ReporteTrabajador(CodEmp As String)
    Screen.MousePointer = vbHourglass
    
    Dim SQL As String
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
     
    Dim cCesados As String
    
    cCesados = ""
    
    If Mensajes("Desea mostrar tambien los empreados cesados", vbYesNo + vbQuestion) = vbYes Then
        cCesados = "S"
    Else
        cCesados = "N"
    End If
     
    Set oRs = ADO_LlenaRs("call RH_SP_ReporteTrabajador('REPORTE','" & CodEmp & "','" & cCesados & "','','');")

    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No existen datos del trabajador", vbOKOnly + vbInformation, gsNomSW
            Exit Sub
        End If

        m_Titulo = "REPORTE DE TRABAJADOR"
        m_Reporte = "Rep_RepTrabajador.rpt"
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, False
    Else
        MsgBox "No existen datos del trabajador", vbOKOnly + vbInformation, gsNomSW
    End If

    Set oRs = Nothing
End Sub

Public Sub sp_Rep_DependientesEmp(CodEmp As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim SCodEmp As String, SSexo As String, SParent As String
    Dim NomArchivo As String
    If CodEmp <> "" Then SCodEmp = " and a.codigo = '" & CodEmp & "'"
    Screen.MousePointer = vbHourglass
    SQL = "Select a.codigo,if(IFNULL(C.DESCRIP,'')='','EMPLEADO',if(IFNULL(C.DESCRIP,'')='CONVIVIENTE','CONYUGE',IFNULL(C.DESCRIP,''))) AS DESCRIP," & _
          "IF(IFNULL(d.descrip,'')='',(select i.descrip from doc_identificacion i where i.tipo_doc_ide=a.coddocide),IFNULL(d.descrip,'')) as td, " & _
          "if(IFNULL(B.NUMDOCIDE,'')='',a.numdocide,B.NUMDOCIDE) as numdocide,if(IFNULL(B.APEPATERNO,'')='',a.apepat,B.APEPATERNO) as apepat, " & _
          "if(IFNULL(B.APEMATERNO,'')='',a.apemat,B.APEMATERNO) as apemat,if(IFNULL(B.NOMBRE,'')='',CONCAT_WS(' ',a.nombre1,a.nombre2),B.NOMBRE) as nombre, " & _
          "if(IFNULL(B.Sexo,'')='',a.sexo,B.Sexo) as sexo,if(ifnull(B.FEC_NAC,'')='',a.fec_nac,B.FEC_NAC) as fec_nac, " & _
          "if(ifnull(B.EDAD,'')='',a.edad,B.EDAD) as edad,CONCAT_WS(' ',A.APEPAT,A.APEMAT,A.NOMBRE1,A.NOMBRE2) AS NOMEMP, " & _
          "(select descrip from cnmdepar where coddep = O.division) as division,r.nombre as nomtrabajo,(select descrip from departamento where codigo = a.departamento) as depa,'1' as cant,if(a.tipo = 3,'PRACTICANTE','') as tipo " & _
          "FROM ((empleado AS A LEFT JOIN familiares AS B ON (A.CODIGO=B.CODEMP)) LEFT JOIN pl_vinculofam AS C ON (B.CODTIPO=C.CODIGO)) left join doc_identificacion as d " & _
          "On (B.CODDOCIDE=D.TIPO_DOC_IDE) left join contrato o on (a.codigo=o.codemp) left join rh_estacionestrabajo r on (o.esttrabajo=r.codigo) " & _
          "Where a.situacion = 1 and a.tipo not in (3,4) and o.estado = 'AP' " & SCodEmp & " Union All Select DISTINCT a.codigo,'EMPLEADO' AS DESCRIP, " & _
          "(select i.descrip from doc_identificacion i where i.tipo_doc_ide=a.coddocide) as td,a.numdocide,a.apepat,a.apemat,CONCAT_WS(' ',a.nombre1,a.nombre2) as nombre, " & _
          "a.sexo,a.fec_nac,a.edad,CONCAT_WS(' ',A.APEPAT,A.APEMAT,A.NOMBRE1,A.NOMBRE2) AS NOMEMP,(select descrip from cnmdepar where coddep = O.division) as division, " & _
          "r.nombre as nomtrabajo,(select descrip from departamento where codigo = a.departamento) as depa,'1' as cant,IF(a.tipo = 3,'PRACTICANTE','') as tipo FROM (empleado AS A inner JOIN familiares AS B ON (A.CODIGO=B.CODEMP)) left join doc_identificacion as d " & _
          "On (B.CODDOCIDE=D.TIPO_DOC_IDE) left join contrato o on (a.codigo=o.codemp) left join rh_estacionestrabajo r on (o.esttrabajo=r.codigo) " & _
          "Where a.tipo not in (3,4) and a.situacion = 1 and o.estado = 'AP' " & SCodEmp & " ORDER BY APEPAT,APEMAT,nombre"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    NomArchivo = Rep_Documents & "\DEPENDIENTES.XLS"
    If Exportar_Excel_rs(NomArchivo, myrs) Then
    End If
    Screen.MousePointer = vbDefault
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_AsistenciaxSede(Mes As String, sede As String)
    Dim SQL As String, Cant As Integer
    Dim RQ As MYSQL_RS
    Dim myrs As MYSQL_RS, CodEmp As String
    Set myrs = New MYSQL_RS
    oConexionMYSQL.Execute ("delete from rh_tmpasistencias where usuario='" & strUsuarioId & "'")
    SQL = "INSERT INTO rh_tmpasistencias(nombres,codigo,sede,envio,usuario) " & _
          "SELECT DISTINCT (select concat(apepat,' ',apemat,' ',nombre1,' ',nombre2) from empleado where codigo=a.emp and tipo not in (3,4)) as nomemp, " & _
          "a.emp,(select nombre from rh_estacionestrabajo where codigo= a.sede) as sede, " & _
          "a.sede,'" & strUsuarioId & "' as usuario " & _
          " FROM rh_entsalempleado as a WHERE a.fecha<>'x' "
    If Mes <> "" Then SQL = SQL & " and left(a.FECHA,7)='" & strAnoSistema & "/" & Mes & "'"
    If sede <> "" Then SQL = SQL & " and a.sede='" & sede & "'"
    SQL = SQL & " GROUP BY a.emp,fecha ORDER BY a.sede,nomemp"
    oConexionMYSQL.Execute SQL
    SQL = "Select * from rh_tmpasistencias where usuario='" & strUsuarioId & "' order by codigo,division"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        Set myrs = Nothing
        MsgBox "No existe movimiento registrado esta fecha", vbOKOnly + vbInformation, "NOVADMIN"
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        Do While Not myrs.EOF
            SQL = "select emp,fecha,tiposede,sede from rh_entsalempleado where emp = '" & myrs.Fields("codigo") & "'"
            If Mes <> "" Then SQL = SQL & " and left(FECHA,7)='" & strAnoSistema & "/" & Mes & "'"
            If sede <> "" Then SQL = SQL & " and sede='" & sede & "'"
            SQL = SQL & " GROUP BY emp,fecha ORDER BY fecha"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                Do While Not RQ.EOF()
                    If Trim(myrs.Fields("codigo")) = RQ.Fields("emp") Then
                        SQL = "update rh_tmpasistencias set D" & Day(RQ.Fields("fecha")) & "='" & RQ.Fields("tiposede") & "'" & _
                              " where codigo='" & myrs.Fields("codigo") & "' and envio = '" & RQ.Fields("sede") & "' and usuario='" & strUsuarioId & "'"
                        oConexionMYSQL.Execute SQL
                    End If
                    RQ.MoveNext
                Loop
            End If
            myrs.MoveNext
        Loop
        myrs.CloseRecordset
        SQL = "Select nombres,codigo,sede,d1,d2,d3,d4,d5,d6,d7,d8,d9,d10,d11,d12,d13,d14,d15,d16,d17,d18,d19,d20, " & _
              "d21,d22,d23,d24,d25,d26,d27,d28,d29,d30,d31,tiposede,division,envio,situacion,usuario " & _
              "from rh_tmpasistencias where usuario='" & strUsuarioId & "' order by envio,nombres"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
    Set RQ = Nothing
End Sub
'
Public Sub sp_RubrosEmp(plani As String, mesI As String, mesF As String, rub As String, emp As String, Div As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim SWhere As String, SFecha As String
    Dim rpt As String
    Screen.MousePointer = vbHourglass
    SQL = ""
    rpt = InputBox("¿Desea ver el Reporte [D]etallado ó [R]esumen", "Rubros Remunerativos", "R")
    If rpt = "R" Then
        SQL = "Select a.tipo as tipoP,p.descrip as plani,a.moneda,a.emp," & _
              "concat_ws(' ',b.apepat,apemat,nombre1,nombre2) as Empleado, " & _
              "a.afp,sum(a.sbasico) as sbasico,sum(if(t.tipo = 'I', " & _
              "case a.tipo when '1' then if(a.unidad= 'D',((a.sbasico/30)*a.cant),a.cant) " & _
              "when '5' then if(a.unidad= 'D',((a.sbasico/180)*a.cant),a.cant) " & _
              "when '6' then if(a.unidad= 'D',((a.sbasico/180)*a.cant),a.cant) end,0)) as ingresos, " & _
              "sum(if(t.tipo = 'D',a.cant,0)) as dsctos,sum(if(t.tipo = 'A',a.cant,0)) as aportes " & _
              "from ((((pl_tareo as a left join empleado as b on (a.emp=b.codigo)) " & _
              "left join pl_rubrosremunerativos as r on (a.rub=r.codigo)) " & _
              "left join pl_tiporubros as t on (r.tipo=t.codigo)) " & _
              "left join pl_periodos as p on (p.codigo=a.tipo)) left join contrato as c on (c.codemp=b.codigo) " & _
              "where a.tipo<>'x'"
    Else
        SQL = "Select a.anomes,a.tipo as tipoP,p.descrip as plani,a.moneda,a.emp," & _
              " concat_ws(' ',b.apepat,apemat,nombre1,nombre2) as Empleado, " & _
              " t.tipo as tipoR,t.descrip as descripTR,a.rub, " & _
              " r.descrip as descripR,a.unidad,a.cant,a.fecha,a.afp,a.sbasico" & _
              " from ((((pl_tareo as a left join empleado as b" & _
              " on (a.emp=b.codigo)) left join pl_rubrosremunerativos as r" & _
              " on (a.rub=r.codigo)) left join pl_tiporubros as t" & _
              " on (r.tipo=t.codigo)) left join pl_periodos as p" & _
              " on (p.codigo=a.tipo)) left join contrato as c on (c.codemp=b.codigo) where a.tipo<>'x' "
        Reporte = "Rep_RubrosEmp.rpt"
        Titulo = "REPORTE DETALLADO POR RUBROS REMUNERATIVOS"
    End If
    If plani <> "" Then
        SQL = SQL & " and a.tipo='" & plani & "' "
    End If
    If mesI <> "" And mesF <> "" Then
        SQL = SQL & " and a.anomes>='" & mesI & "' and a.anomes<='" & mesF & "' "
    Else
        If mesI <> "" And mesF = "" Then
            SQL = SQL & " and a.anomes='" & mesI & "'"
        Else
            If mesF <> "" And mesI = "" Then
                SQL = SQL & " and a.anomes='" & mesF & "'"
            End If
        End If
    End If
    If rub <> "" Then
        SQL = SQL & " and a.rub='" & rub & "' "
    End If
    If emp <> "" Then
        SQL = SQL & " and a.emp='" & emp & "' "
    End If
    
    If Div <> "" And Div <> "0006" Then
        SQL = SQL & " and c.division='" & Div & "' "
    End If
    
    If rpt = "D" Then
        SQL = SQL & " Group By a.anomes,a.emp,a.tipo,a.rub,a.unidad"
    Else
        SQL = SQL & " Group By a.emp,a.tipo,a.rub ORDER BY b.apepat,a.anomes,rub"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_FacturasExt(FactIni As String, FactFin As String, MesIni As String, MesFin As String, Cli As Integer, Cmd As CommonDialog)
    Dim SQL As String, tc As Double, a As Variant
    Dim myrs As MYSQL_RS
    Screen.MousePointer = vbHourglass
    Cant = Cant + 1
    If Cli = 0 Then
        modPrincipal.TipoCambio
        a = InputBox("Ingrese el tipo de cambio", "TIPO DE CAMBIO", CStr(dblTipoCmbV))
        tc = CDbl(IIf(IsNumeric(a), a, 1))
    End If
    If Cli = 1 Then
        SQL = "Select DISTINCT B.CODIGO,B.DESCRIP from documento_contables AS A LEFT JOIN cnauxil AS B " & _
              "ON (A.CODIGO=B.CODIGO) where A.serie='00001' AND A.AUXILIAR='2' AND B.AUXILIAR='2'"
    Else
        SQL = "Select CONCAT_WS('-',serie,correl) as Doc,'BRANDTADMIN' AS SD,A.CODIGO, " & _
              "IF(mon = 'N',A.SUBTOTAL/" & tc & ",A.SUBTOTAL) as subtotal,' ' P,' ' L, " & _
              "c.descrip,date_format(a.fec_emision,'%d/%m/%Y'),B.COD_ESTADO " & _
              "from (documento_contables AS A LEFT JOIN movi_documento AS B " & _
              "ON (a.identificador=b.identificador)) LEFT JOIN cnmdepar AS C " & _
              "ON (a.division=C.coddep) where auxiliar='2' AND serie='00001' "
    End If
    If Trim(FactFin) <> "" And Trim(FactIni) <> "" Then
        SQL = SQL & "and correl >='" & Trim(FactIni) & "' " & _
                    "and correl <='" & Trim(FactFin) & "'"
    Else
        If Trim(FactFin) <> "" And Trim(FactIni) = "" Then
            SQL = SQL & "and correl ='" & Trim(FactIni) & "' "
        Else
            If Trim(FactFin) = "" And Trim(FactIni) <> "" Then
                SQL = SQL & "and correl ='" & Trim(FactFin) & "' "
            End If
        End If
    End If
    If Trim(MesIni) <> "" And Trim(MesFin) <> "" Then
        SQL = SQL & "and date_format(fec_emision,'%Y%m') >='" & Trim(MesIni) & "' " & _
                    "AND date_format(fec_emision,'%Y%m') <='" & Trim(MesFin) & "'"
    Else
        If Trim(MesIni) <> "" And Trim(MesFin) = "" Then
            SQL = SQL & "and date_format(fec_emision,'%Y%m') ='" & Trim(MesIni) & "' "
        Else
            If Trim(MesIni) = "" And Trim(MesFin) <> "" Then
                SQL = SQL & "and date_format(fec_emision,'%Y%m') ='" & Trim(MesFin) & "' "
            End If
        End If
    End If
    If Cli = 1 Then SQL = SQL & " ORDER BY B.CODIGO"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If frmRepFact.chkExcel.Value = 1 Then
        If Cli = 0 Then
            Cmd.CancelError = True
            Cmd.Flags = cdlOFNFileMustExist Or _
            cdlOFNHideReadOnly Or _
            cdlOFNExplorer Or _
            cdlOFNLongNames
            On Error Resume Next
                Cmd.Filename = strAnoSistema & strMesSistema & "Invoice Details Layout"
                Cmd.ShowSave
                If err.Number = cdlCancel Or err.Number <> 0 Then Screen.MousePointer = vbDefault: Exit Sub
                RutaArchivo = Cmd.Filename
                If RutaArchivo = "" Then Screen.MousePointer = vbDefault: Exit Sub
        End If
        If Cant = 1 Then
            Set RQ_FACT = myrs
            sp_Rep_FacturasExt FactIni, FactFin, MesIni, MesFin, 1, Cmd
        Else
            Cant = 0
            Exportar_Excel_Fact RutaArchivo, RQ_FACT, myrs
        End If
    Else
        EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    End If
    Screen.MousePointer = vbDefault
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Refrigerios(valor As Double, FecIni As String, FecFin As String, Mes As String, Divi As String)
    Dim SQL As String, SVal As String, SDiv As String
    Dim myrs As MYSQL_RS
    If val(valor) > 0 Then
        SVal = " and monto = " & valor & ""
    End If
    If Trim(Divi) <> "0006" And Trim(Divi) <> "00" Then
        SDiv = " and C.division = '" & Divi & "'"
    End If
    Screen.MousePointer = vbHourglass
    SQL = "SELECT '" & strAnoSistema & "/" & Mes & "' as anomes,CONCAT(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) AS NOMBRE,k.codemp, " & _
          "k.monto,SUM(k.cantidad) AS cant,(select descrip from cnmdepar where coddep = C.division) as division,(SELECT DESCRIP FROM cnauxil C WHERE C.AUXILIAR=K.AUX AND C.CODIGO=K.CODAUX) AS CODAUX From " & _
          "(SELECT r.codemp,r.montoa AS monto,count(*) AS cantidad,(r.montoa * count(*)) AS subtotal,AUXILIARA AS AUX,codauxa AS codaux " & _
          "FROM rh_refriempleados R WHERE fecha BETWEEN '" & FecIni & "' AND '" & FecFin & "' " & _
          "GROUP BY codemp,codauxa,montoa) AS K INNER JOIN empleado E ON (K.CODEMP=E.CODIGO) LEFT JOIN CONTRATO C ON (C.CODEMP=E.CODIGO) where C.ESTADO = 'AP' " & SVal & SDiv & " GROUP BY K.CODEMP,K.CODAUX,K.MONTO " & _
          "ORDER BY K.CODAUX,E.APEPAT,E.APEMAT"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Bonos(valor As Double, FecIni As String, FecFin As String, Mes As String, Divi As String)
    Dim SQL As String, SVal As String, SDiv As String
    Dim myrs As MYSQL_RS
    If val(valor) > 0 Then
        SVal = " and monto = " & valor & ""
    End If
    If Trim(Divi) <> "0006" And Trim(Divi) <> "00" Then
        SDiv = " and C.division = '" & Divi & "'"
    End If
    Screen.MousePointer = vbHourglass
    SQL = "SELECT * from pl_tareo where anomes = '" & strAnoSistema & Mes & "' and tipo = 2"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.EOF() Then
        SQL = "SELECT '" & strAnoSistema & "/" & Mes & "' as anomes,CONCAT(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) AS NOMBRE," & _
              "k.codemp, k.monto,SUM(k.cantidad) AS cant,(select descrip from cnmdepar where coddep = C.division) as division,k.subtotal,' ' AS CODAUX " & _
              "From (SELECT r.codemp,r.bono AS monto,count(*) AS cantidad,sum(r.bono) AS subtotal " & _
              "FROM rh_bonosempleados R WHERE fecha BETWEEN '" & FecIni & "' AND '" & FecFin & "' GROUP BY codemp,r.bono) AS K " & _
              "INNER JOIN empleado E ON (K.CODEMP=E.CODIGO) LEFT JOIN CONTRATO C ON (C.CODEMP=E.CODIGO) where C.ESTADO = 'AP' " & SVal & SDiv & " GROUP BY K.CODEMP,k.monto ORDER BY E.APEPAT,E.APEMAT,e.nombre1,e.nombre2"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
    Else
        SQL = "SELECT * from pl_tareo where anomes = '" & strAnoSistema & Mes & "' and tipo = 1"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
        If myrs.EOF() Then
            SQL = "SELECT '" & strAnoSistema & "/" & Mes & "' AS anomes,CONCAT(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) AS NOMBRE, " & _
                  "k.codemp,k.monto,IF(ifnull(pl.cant,0)=0,0,ifnull(pl.cant,0)/k.monto) AS cant,(select descrip from cnmdepar " & _
                  "where coddep = C.division) as division,ifnull(pl.cant,0) as PRIMseg,'SEGUNDA' AS tipo From ((SELECT r.codemp, " & _
                  "r.bono AS monto,count(*) AS cantidad,sum(r.bono) AS subtotal FROM rh_bonosempleados R WHERE " & _
                  "fecha BETWEEN '" & FecIni & "' AND '" & FecFin & "' GROUP BY codemp,r.bono) AS K INNER JOIN empleado E ON " & _
                  "(K.CODEMP=E.CODIGO) LEFT JOIN CONTRATO C ON (C.CODEMP=E.CODIGO)) left join (select emp,cant,anomes from pl_tareo O " & _
                  "where tipo = 1 and anomes >= '" & strAnoSistema & Mes & "' and anomes <= '" & strAnoSistema & Mes & "' " & _
                  "and rub = '1001') as pl on (e.codigo=pl.emp) where c.codigo = (select max(codigo) from contrato r where r.codemp=e.codigo group by r.codemp) GROUP BY K.CODEMP,k.monto Union " & _
                  "SELECT '" & strAnoSistema & "/" & Mes & "' as anomes,CONCAT(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) AS NOMBRE," & _
                  "k.codemp, k.monto,SUM(k.cantidad) AS cant,(select descrip from cnmdepar where coddep = C.division) as division,k.subtotal as primseg,'PRIMERA' AS tipo " & _
                  "From (SELECT r.codemp,r.bono AS monto,count(*) AS cantidad,sum(r.bono) AS subtotal " & _
                  "FROM rh_bonosempleados R WHERE fecha BETWEEN '" & FecIni & "' AND '" & FecFin & "' GROUP BY codemp,r.bono) AS K " & _
                  "INNER JOIN empleado E ON (K.CODEMP=E.CODIGO) LEFT JOIN CONTRATO C ON (C.CODEMP=E.CODIGO) where C.ESTADO = 'AP' " & SVal & SDiv & " GROUP BY K.CODEMP,k.monto ORDER BY DIVISION,TIPO,nombre"
              Set myrs = oConexion.EjecutaSelectRS(SQL)
        Else
            SQL = "SELECT '" & strAnoSistema & "/" & Mes & "' AS anomes,CONCAT(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) AS NOMBRE, " & _
                  "k.codemp,k.monto,IF(ifnull(pl.cant,0)=0,0,ifnull(pl.cant,0)/k.monto) AS cant,(select descrip from cnmdepar " & _
                  "where coddep = C.division) as division,ifnull(pl.cant,0) as PRIMseg,'PRIMERA' AS tipo From ((SELECT r.codemp, " & _
                  "r.bono AS monto,count(*) AS cantidad,sum(r.bono) AS subtotal FROM rh_bonosempleados R WHERE " & _
                  "fecha BETWEEN '" & FecIni & "' AND '" & FecFin & "' GROUP BY codemp,r.bono) AS K INNER JOIN empleado E ON " & _
                  "(K.CODEMP=E.CODIGO) LEFT JOIN CONTRATO C ON (C.CODEMP=E.CODIGO)) left join (select emp,cant,anomes from pl_tareo O " & _
                  "where tipo = 2 and anomes >= '" & strAnoSistema & Mes & "' and anomes <= '" & strAnoSistema & Mes & "' " & _
                  "and rub = '1001') as pl on (e.codigo=pl.emp) where c.codigo = (select max(codigo) from contrato r where r.codemp=e.codigo group by r.codemp) GROUP BY K.CODEMP,k.monto Union " & _
                  "SELECT '" & strAnoSistema & "/" & Mes & "' AS anomes,CONCAT(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) AS NOMBRE, " & _
                  "k.codemp,k.monto,IF('" & FecFin & "' > '" & Year(FecFin) & "/" & Mid(FecFin, 6, 2) & "/15',IF(if(ifnull(pl.cant,0)=k.subtotal,0,k.subtotal-ifnull(pl.cant,0))=0,0,if(ifnull(pl.cant,0)=k.subtotal, " & _
                  "0,k.subtotal-ifnull(pl.cant,0))/k.monto),0) AS cant,(select descrip from cnmdepar where coddep = C.division) as division, " & _
                  "IF('" & FecFin & "' > '" & Year(FecFin) & "/" & Mid(FecFin, 6, 2) & "/15',if(ifnull(pl.cant,0)=k.subtotal,0,k.subtotal-ifnull(pl.cant,0)),0) AS primseg,'SEGUNDA' as tipo From " & _
                  "((SELECT r.codemp,r.bono AS monto,count(*) AS cantidad,sum(r.bono) AS subtotal FROM rh_bonosempleados R WHERE " & _
                  "fecha BETWEEN '" & FecIni & "' AND '" & FecFin & "' GROUP BY codemp,r.bono) AS K INNER JOIN empleado E ON " & _
                  "(K.CODEMP=E.CODIGO) LEFT JOIN CONTRATO C ON (C.CODEMP=E.CODIGO)) left join (select emp,cant,anomes from pl_tareo " & _
                  "where tipo = 2 and anomes >= '" & strAnoSistema & Mes & "' and anomes <= '" & strAnoSistema & Mes & "' and " & _
                  "rub = '1001') as pl on (e.codigo=pl.emp) where c.codigo = (select max(codigo) from contrato r where r.codemp=e.codigo group by r.codemp) GROUP BY K.CODEMP,k.monto ORDER BY DIVISION,TIPO,nombre"
            Set myrs = oConexion.EjecutaSelectRS(SQL)
        End If
    End If
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_EspaciosContratos(CodEmp As String, FechaIni As String, FechaFin As String)
    Dim SQL As String, SEmp As String
    Dim RQ As MYSQL_RS, emp As String, it As Integer, SSit As String
    Dim FecIni As String, I As Integer, CodCont As String, rpt As String
    If CodEmp <> "" Then SEmp = " and codemp = '" & CodEmp & "'"
    rpt = InputBox("¿Desea incluir a los empleados CESADOS?", "Contratos de Empleados", "N")
    If rpt = "N" Then SSit = " and situacion =1 " Else SSit = ""
    SQL = "Select CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres,codtipo, " & _
          "(select descrip from cncontrato t where t.codigo=codtipo) as descrip,f_inicio," & _
          "f_termino,c.codigo,Estado,e.Situacion,codemp from contrato c inner join empleado e on(e.codigo = c.codemp) " & _
          "where 1 = 1 and f_inicio >= '" & FechaIni & "' and f_inicio <= '" & FechaFin & "' " & SEmp & SSit & " order by nombres asc,c.codigo desc"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If RQ.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set RQ = Nothing
        Exit Sub
    End If
    If RQ.RecordCount > 0 Then
        I = 0
        it = 1
        SQL = "delete from rh_tmpespacioscontratos where usuario = '" & strUsuarioId & "'"
        oConexionMYSQL.Execute SQL
        Do While Not RQ.EOF()
            If I = 0 Then
                FecIni = RQ.Fields("f_inicio")
                CodCont = val(Right(RQ.Fields("codigo"), 2))
                emp = Trim(RQ.Fields("codemp"))
                I = I + 1
                RQ.MoveNext
            Else
                If Trim(RQ.Fields("codemp")) = emp Then
                    If CDate(FecIni) - CDate(IIf(Trim(RQ.Fields("f_termino")) = "", Date, RQ.Fields("f_termino"))) > 1 Then
                        SQL = "insert into rh_tmpespacioscontratos values(" & it & ",'" & Trim(RQ.Fields("nombres")) & "','" & val(Right(Trim(RQ.Fields("codigo")), 2)) & "  y  " & CodCont & "', " & _
                              "'" & CDate(RQ.Fields("f_termino")) & "  y  " & CDate(FecIni) & "', " & _
                              "'" & IIf(RQ.Fields("situacion") = 0, "CESADO", "ACTIVO") & "','" & strUsuarioId & "')"
                        oConexionMYSQL.Execute SQL
                        it = it + 1
                        FecIni = RQ.Fields("f_inicio")
                        CodCont = val(Right(RQ.Fields("codigo"), 2))
                    Else
                        FecIni = RQ.Fields("f_inicio")
                        CodCont = val(Right(RQ.Fields("codigo"), 2))
                    End If
                    I = I + 1
                    RQ.MoveNext
                Else
                    emp = Trim(RQ.Fields("codemp"))
                    I = 0
                End If
            End If
        Loop
        SQL = "SELECT item,nombres,contratos,fechas,estado FROM rh_tmpespacioscontratos order by estado,nombres,contratos"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, RQ
    End If
    Set RQ = Nothing
End Sub
Public Sub sp_Rep_EmpleadosxMes(anomesini As String, anomesfin As String, Division As String)
    Dim SQL As String, STRDiv As String
    Dim aini As String, AFin As String, Mes As String, Anio As String
    Dim myrs As MYSQL_RS, RQ As MYSQL_RS
    Screen.MousePointer = vbHourglass
    
    If Division <> "" And Division <> "0006" Then
        STRDiv = " and b.divgas = '" & Division & "'"
    End If
    SQL = "delete from tempdetctacte_2 where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQL
    If Right(anomesini, 2) = "01" Then
        Mes = "12"
        Anio = Left(anomesini, 4) - 1
    Else
        Mes = Right("00" & Right(anomesini, 2) - 1, 2)
        Anio = Left(anomesini, 4)
    End If
    
    aini = Anio & Mes
    AFin = aini
    
    sp_Rep_EmpleadosxMesCargaIni aini, AFin, Division
    sp_Rep_EmpleadosxMesCargaIni anomesini, anomesfin, Division, aini
    
    SQL = "delete from tempdetctacte where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQL
    SQL = "Select '1AL " & "01-" & Right(anomesini, 2) & "-" & Left(anomesini, 4) & "' as nombre,descripcuenta AS DIVISION,count(*) + " & _
          "ifnull((select count(*) FROM EMPLEADO E LEFT JOIN CONTRATO T ON (E.CODIGO=T.CODEMP) where fec_ingreso = '" & Left(anomesini, 4) & "/" & Right(anomesini, 2) & "/01" & "' AND " & _
          "E.SITUACION = 1 AND E.TIPO not in (3,4) AND T.ESTADO = 'AP' and t.divgas=t1.coddep group by t.divgas),0) as tot From tempdetctacte_2 T1 where abrevia = 'D' and " & _
          "usuario = '" & strUsuarioId & "' and auxiliar = '" & aini & "' AND IFNULL(CUENTA,'') <> 'CESADO' group by coddep Union " & _
          "select '2INGRESOS' as nombre,(SELECT DESCRIP FROM CNMDEPAR D WHERE D.CODDEP=T.DIVGAS) AS DIVISION,count(*) as tot FROM EMPLEADO E " & _
          "LEFT JOIN CONTRATO T ON (E.CODIGO=T.CODEMP) where fec_ingreso >= '" & Left(anomesini, 4) & "/" & Right(anomesini, 2) & "/02" & "' " & _
          "and fec_ingreso <= '" & Left(anomesini, 4) & "/" & Right(anomesini, 2) & "/" & Day(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0)) & "' " & _
          "AND E.SITUACION = 1 AND E.TIPO not in (3,4) AND T.ESTADO = 'AP' GROUP BY t.divgas Union " & _
          "select '3TRANSFERENCIAS' as nombre,descripcuenta as division,IFNULL(count(*),0) as tot from tempdetctacte_2 where abrevia = 'D' " & _
          "and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND IFNULL(desdoc,'') = 'TI' group by coddep Union " & _
          "select '4CESADOS AL " & Day(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0)) - 1 & "-" & Right(anomesini, 2) & "-" & Left(anomesini, 4) & "' as nombre,descripcuenta AS DIVISION,count(*) as tot from tempdetctacte_2 where abrevia = 'D' " & _
          "and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND IFNULL(CUENTA,'') = 'CESADO' and IFNULL(voucher,'') <> 'U' GROUP BY coddep Union " & _
          "select '5AL " & Day(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0)) & "-" & Right(anomesini, 2) & "-" & Left(anomesini, 4) & "' AS NOMBRE,descripcuenta AS DIVISION,count(*) as tot from tempdetctacte_2 where abrevia = 'D' and usuario = '" & strUsuarioId & "' " & _
          "and auxiliar = '" & anomesini & "' AND (IFNULL(CUENTA,'') <> 'CESADO' OR VOUCHER = 'U') group by coddep Union " & _
          "select '6CESES DEL " & Day(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0)) & "-" & Right(anomesini, 2) & "-" & Left(anomesini, 4) & "' as  NOMBRE,descripcuenta AS DIVISION,count(*) as tot from tempdetctacte_2 where abrevia = 'D' " & _
          "and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND IFNULL(CUENTA,'') = 'CESADO' and voucher = 'U' group by coddep Union " & _
          "select '701-" & Right("00" & val(Right(anomesini, 2)) + 1, 2) & "-" & Left(anomesini, 4) & "' as NOMBRE,descripcuenta AS DIVISION,count(*) + ifnull(k.tot,0) as tot from tempdetctacte_2 d left join (SELECT divgas,count(*) AS TOT FROM EMPLEADO E " & _
          "LEFT JOIN CONTRATO T ON (E.CODIGO=T.CODEMP) WHERE E.SITUACION = 1 AND E.TIPO not in (3,4) AND FEC_INGRESO ='" & Left(anomesini, 4) & "/" & Right(val(anomesini) + 1, 2) & "/01" & "' " & _
          "AND ESTADO = 'AP' group by divgas) as k on (k.divgas=d.coddep) where abrevia = 'D' and usuario = '" & strUsuarioId & "' " & _
          "and auxiliar = '" & anomesini & "' AND IFNULL(CUENTA,'') <> 'CESADO' group by coddep order by nombre,division"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then

On Error GoTo CtrlExporta

        Dim o_Excel As Object
        Dim o_Libro As Object
        Set o_Excel = New Excel.Application
        Set o_Libro = o_Excel.Workbooks.Add
        CantCta = 3
        Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "RESUMEN", anomesini
        
        
        SQL = "select CODAUX AS CODIGO,NOMBRE AS NOMBRES,descripcuenta AS DIVISION from tempdetctacte_2 where " & _
              "abrevia = 'D' and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND IFNULL(CUENTA,'') <> 'CESADO' " & _
              " UNION SELECT E.CODIGO,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS NOMBRES,(SELECT DESCRIP FROM CNMDEPAR D WHERE " & _
              "D.CODDEP=T.DIVGAS) AS DIVISION FROM EMPLEADO E LEFT JOIN CONTRATO T ON (E.CODIGO=T.CODEMP) WHERE E.SITUACION = 1 AND E.TIPO not in (3,4) AND " & _
              "FEC_INGRESO ='" & Left(anomesini, 4) & "/" & Right(val(anomesini) + 1, 2) & "/01" & "' AND ESTADO = 'AP' order by DIVISION,nombreS"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "01-" & Right("00" & val(Right(anomesini, 2)) + 1, 2) & "-" & Left(anomesini, 4)
        End If
        SQL = "SELECT E.CODIGO,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS NOMBRES,(SELECT DESCRIP FROM CNMDEPAR D WHERE " & _
              "D.CODDEP=T.DIVGAS) AS DIVISION FROM EMPLEADO E LEFT JOIN CONTRATO T ON (E.CODIGO=T.CODEMP) WHERE E.SITUACION = 1 AND E.TIPO not in (3,4) AND " & _
              "FEC_INGRESO ='" & Left(anomesini, 4) & "/" & Right(val(anomesini) + 1, 2) & "/01" & "' AND ESTADO = 'AP' order by DIVISION,nombreS"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "INGRESOS DEL 01-" & Right("00" & val(Right(anomesini, 2)) + 1, 2) & "-" & Left(anomesini, 4)
        End If
        SQL = "select CODAUX AS CODIGO,NOMBRE AS NOMBRES,descripcuenta AS DIVISION from tempdetctacte_2 where " & _
              "abrevia = 'D' and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND IFNULL(CUENTA,'') = 'CESADO' " & _
              "and voucher = 'U' order by descripcuenta,nombre"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "CESES DEL " & Day(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0)) & "-" & Right(anomesini, 2) & "-" & Left(anomesini, 4)
        End If
        SQL = "select CODAUX AS CODIGO,NOMBRE AS NOMBRES,descripcuenta AS DIVISION from tempdetctacte_2 where " & _
              "abrevia = 'D' and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND (IFNULL(CUENTA,'') <> 'CESADO' OR VOUCHER = 'U') " & _
              "order by descripcuenta,nombre"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, Day(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0)) & "-" & Right(anomesini, 2) & "-" & Left(anomesini, 4) 'AnoMesIni
        End If
        SQL = "select CODAUX AS CODIGO,NOMBRE AS NOMBRES,descripcuenta AS DIVISION from tempdetctacte_2 where " & _
              "abrevia = 'D' and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND IFNULL(CUENTA,'') = 'CESADO' " & _
              "and IFNULL(voucher,'') <> 'U' order by descripcuenta,nombre"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "CESES AL " & Day(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0)) - 1 & "-" & Right(anomesini, 2) & "-" & Left(anomesini, 4)
        End If
        SQL = "select CODAUX AS CODIGO,NOMBRE AS NOMBRES,descripcuenta AS DIVISION from tempdetctacte_2 where " & _
              "abrevia = 'D' and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND IFNULL(desdoc,'') = 'TI' " & _
              "order by descripcuenta,nombre"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "TRANSF.INTERNA"
        Else
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "TRANSF.INTERNA"
        End If
        SQL = "select e.CODIGO,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS NOMBRES,(SELECT DESCRIP FROM CNMDEPAR D WHERE " & _
              "D.CODDEP=T.DIVGAS) AS DIVISION FROM EMPLEADO E LEFT JOIN CONTRATO T ON (E.CODIGO=T.CODEMP) " & _
              "where fec_ingreso >= '" & Left(anomesini, 4) & "/" & Right(anomesini, 2) & "/02" & "' and " & _
              "fec_ingreso <= '" & Left(anomesini, 4) & "/" & Right(anomesini, 2) & "/" & Day(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0)) & "' " & _
              "AND E.SITUACION = 1 AND E.TIPO not in (3,4) AND T.ESTADO = 'AP' ORDER BY DIVISION,NOMBRES"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "INGRESOS"
        Else
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "INGRESOS"
        End If
        SQL = "select CODAUX AS CODIGO,NOMBRE AS NOMBRES,descripcuenta AS DIVISION from tempdetctacte_2 where " & _
              "abrevia = 'D' and usuario = '" & strUsuarioId & "' and auxiliar = '" & aini & "' AND IFNULL(CUENTA,'') <> 'CESADO' " & _
              "UNION ALL select e.CODIGO,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS NOMBRES, " & _
              "(SELECT DESCRIP FROM CNMDEPAR D WHERE D.CODDEP=T.DIVGAS) AS DIVISION FROM EMPLEADO E LEFT JOIN CONTRATO T " & _
              "ON (E.CODIGO=T.CODEMP) where fec_ingreso = '" & Left(anomesini, 4) & "/" & Right(anomesini, 2) & "/01" & "' " & _
              "AND E.SITUACION = 1 AND E.TIPO not in (3,4) AND T.ESTADO = 'AP' order by division,nombres"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Exportar_Excel_rs_RRHH RQ, o_Excel, o_Libro, "01-" & Right(anomesini, 2) & "-" & Left(aini, 4)
        End If
        o_Excel.Visible = True
        o_Libro.SaveAs Rep_Documents & "\" & "HFM" & Replace(Date, "/", "") & ".xls"
        If Not o_Excel Is Nothing Then Set o_Excel = Nothing
        If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    End If
    Screen.MousePointer = vbDefault
    Set myrs = Nothing
    Set RQ = Nothing
Exit Sub
CtrlExporta:
    Screen.MousePointer = vbDefault
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Set myrs = Nothing
    Set RQ = Nothing
End Sub
Public Sub ConsultaResumen(campo As String, Tipo As String, AnoMes As String, Opc As String, AM As String)
Dim SQL As String
    SQL = "insert into tempdetctacte (nombre,auxiliar,debe,haber,debeeq,habereq,monedaequiv,usuario) " & _
          "SELECT " & campo & "," & AM & ",(SELECT COUNT(*) FROM tempdetctacte_2 t1 WHERE USUARIO  = '" & strUsuarioId & "' AND ABREVIA = 'D' " & _
          "AND CODDEP = '0009' and t.auxiliar=t1.auxiliar " & Tipo & " GROUP BY AUXILIAR ORDER BY AUXILIAR) * " & Opc & " AS 'MDT', " & _
          "(SELECT COUNT(*) FROM tempdetctacte_2 t1 WHERE USUARIO  = '" & strUsuarioId & "' AND ABREVIA = 'D' AND CODDEP = '0003' " & _
          "and t.auxiliar=t1.auxiliar " & Tipo & " GROUP BY AUXILIAR ORDER BY AUXILIAR) * " & Opc & " AS 'RH', " & _
          "(SELECT COUNT(*) FROM tempdetctacte_2 t1 WHERE USUARIO  = '" & strUsuarioId & "' AND ABREVIA = 'D' AND CODDEP = '0008' " & _
          "and t.auxiliar=t1.auxiliar " & Tipo & " GROUP BY AUXILIAR ORDER BY AUXILIAR) * " & Opc & " AS 'SS', " & _
          "(SELECT COUNT(*) FROM tempdetctacte_2 t1 WHERE USUARIO  = '" & strUsuarioId & "' AND ABREVIA = 'D' AND CODDEP = '0001' " & _
          "and t.auxiliar=t1.auxiliar " & Tipo & " GROUP BY AUXILIAR ORDER BY AUXILIAR) * " & Opc & " AS 'CS', " & _
          "(SELECT COUNT(*) FROM tempdetctacte_2 t1 WHERE USUARIO  = '" & strUsuarioId & "' AND ABREVIA = 'D' AND CODDEP = '0002' " & _
          "and t.auxiliar=t1.auxiliar " & Tipo & " GROUP BY AUXILIAR ORDER BY AUXILIAR) * " & Opc & " AS 'IP','" & strUsuarioId & "' " & _
          "FROM tempdetctacte_2 T WHERE USUARIO  = '" & strUsuarioId & "'" & AnoMes & " GROUP BY AUXILIAr"
    oConexionMYSQL.Execute SQL
End Sub
Public Sub sp_Rep_EmpleadosxMesCargaIni(anomesini As String, anomesfin As String, Division As String, Optional CMP As String)
    Dim SQL As String, STRDiv As String
    Dim myrs As MYSQL_RS, RQ As MYSQL_RS
    If Division <> "" And Division <> "0006" Then
        STRDiv = " and b.divgas = '" & Division & "'"
    End If
    SQL = "SELECT * FROM PL_TAREO WHERE TIPO = 1 AND ANOMES >='" & anomesini & "' AND ANOMES <='" & anomesfin & "' " & _
          "AND RUB='121'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        SQL = "INSERT INTO tempdetctacte_2 (auxiliar,coddep,descripcuenta,MDEBE,usuario,abrevia) " & _
              "SELECT A.ANOMES,b.divgas,(select descrip from cnmdepar r where r.coddep=b.divgas) as descripdiv, " & _
              "COUNT(a.Tipo) As CANTIDAD,'" & strUsuarioId & "','R' FROM pl_tareo AS A LEFT JOIN contrato AS B ON (A.CODCONTRATO=B.CODIGO and a.emp=b.codemp) " & _
              "WHERE A.TIPO='1' AND A.ANOMES >='" & anomesini & "' AND A.ANOMES <='" & anomesfin & "' AND A.RUB='121' " & STRDiv & _
              "GROUP BY A.ANOMES,b.divgas ORDER BY A.ANOMES,b.divgas"
        oConexionMYSQL.Execute SQL
        SQL = "INSERT INTO tempdetctacte_2 (auxiliar,CODAUX,NOMBRE,coddep,descripcuenta,MDEBE,usuario,abrevia) " & _
              "SELECT A.ANOMES,c.codigo,CONCAT_WS(' ',c.apepat,c.apemat,C.nombre1,C.nombre2) as Nombre,B.divgas, " & _
              "(select descrip from cnmdepar r where r.coddep=b.divgas) as descripdiv,COUNT(A.TIPO) as cantidad,'" & strUsuarioId & "','D' " & _
              "From (pl_tareo AS A LEFT JOIN contrato AS B ON (A.CODCONTRATO=B.CODIGO and a.emp=b.codemp)) " & _
              "left join empleado as c ON (B.CODEMP=C.CODIGO) WHERE A.TIPO='1' AND A.ANOMES >='" & anomesini & "' " & _
              "AND A.ANOMES <='" & anomesfin & "' AND A.RUB='121' " & STRDiv & " GROUP BY A.ANOMES,a.emp,B.divgas " & _
              "ORDER BY A.ANOMES,CONCAT_WS(' ',c.apepat,c.apemat,C.nombre1,C.nombre2),B.divgas"
        oConexionMYSQL.Execute SQL
    Else
        SQL = "INSERT INTO tempdetctacte_2 (auxiliar,coddep,descripcuenta,MDEBE,usuario,abrevia) " & _
              "SELECT '" & anomesini & "' as anomes,b.divgas,(select descrip from cnmdepar r where r.coddep=b.divgas) as descripdiv, " & _
              "COUNT(*) As CANTIDAD,'" & strUsuarioId & "','R' FROM EMPLEADO C LEFT JOIN contrato AS B ON (C.CODIGO=b.codemp) " & _
              "WHERE C.SITUACION = 1 AND c.tipo not in (3,4) and  b.codigo = (select max(codigo) from contrato t where t.codemp=c.codigo group by t.codemp) " & STRDiv & _
              "GROUP BY anomes,b.divgas ORDER BY anomes,b.divgas"
        oConexionMYSQL.Execute SQL
        SQL = "INSERT INTO tempdetctacte_2 (auxiliar,CODAUX,NOMBRE,coddep,descripcuenta,MDEBE,usuario,abrevia) " & _
              "SELECT '" & anomesini & "' as anomes,c.codigo,CONCAT_WS(' ',c.apepat,c.apemat,C.nombre1,C.nombre2) as Nombre,B.divgas, " & _
              "(select descrip from cnmdepar r where r.coddep=b.divgas) as descripdiv,1 as cantidad,'" & strUsuarioId & "','D' " & _
              "From empleado as c LEFT JOIN contrato AS B ON (B.CODEMP=C.CODIGO) " & _
              "WHERE c.situacion = 1 and c.tipo not in (3,4) and b.codigo = (select max(codigo) from contrato t where t.codemp=c.codigo group by t.codemp) " & _
              STRDiv & " GROUP BY ANOMES,c.codigo,B.divgas " & _
              "ORDER BY ANOMES,CONCAT_WS(' ',c.apepat,c.apemat,C.nombre1,C.nombre2),B.divgas"
        oConexionMYSQL.Execute SQL
    End If
    SQL = "select auxiliar,CODAUX,NOMBRE,coddep,descripcuenta,MDEBE,usuario,abrevia from tempdetctacte_2 where abrevia = 'D' and usuario = '" & strUsuarioId & "' and AUXILIAR = '" & anomesini & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        Do While Not RQ.EOF()
            SQL = "SELECT e.codigo,c.divgas,e.fec_cese FROM empleado e inner join contrato c on (e.codigo=c.codemp) " & _
                  "where c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp " & _
                  "group by codemp)  and (situacion = 0 and fec_cese " & _
                  "BETWEEN '" & Left(anomesini, 4) & "/" & Right(anomesini, 2) & "/01" & "' and " & _
                  "'" & Left(anomesfin, 4) & "/" & Right(anomesfin, 2) & "/" & Day(DateSerial(Left(anomesfin, 4), val(Right(anomesfin, 2)) + 1, 0)) & "') " & _
                  "and e.codigo = '" & Trim(RQ.Fields("codaux")) & "'"
            Set myrs = oConexion.EjecutaSelectRS(SQL)
            If Not myrs.EOF() Then
                SQL = "update tempdetctacte_2 set cuenta = 'CESADO' WHERE CODAUX = '" & Trim(RQ.Fields("codaux")) & "' AND USUARIO = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "'"
                oConexionMYSQL.Execute SQL
                SQL = "update tempdetctacte_2 set MDEBE = MDEBE -1 WHERE coddep = '" & Trim(RQ.Fields("divgas")) & "' AND USUARIO = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "'"
                oConexionMYSQL.Execute SQL
                If CMP <> "" Then
                    If Format(Trim(myrs.Fields("fec_cese")), "dd/mm/yyyy") = Format(DateSerial(Left(anomesini, 4), val(Right(anomesini, 2)) + 1, 0), "dd/mm/yyyy") Then
                        SQL = "update tempdetctacte_2 set voucher = 'U' WHERE CODAUX = '" & Trim(RQ.Fields("codaux")) & "' AND USUARIO = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "'"
                        oConexionMYSQL.Execute SQL
                    End If
                End If
            End If
            If CMP <> "" Then
                SQL = "select coddep from tempdetctacte_2 where abrevia = 'D' and usuario = '" & strUsuarioId & "' and AUXILIAR = '" & CMP & "' and codaux = '" & Trim(RQ.Fields("codaux")) & "'"
                Set myrs = oConexion.EjecutaSelectRS(SQL)
                If Not myrs.EOF() Then
                    If Trim(RQ.Fields("coddep")) <> Trim(myrs.Fields("coddep")) Then
                        SQL = "update tempdetctacte_2 set desdoc = 'TI' WHERE CODAUX = '" & Trim(RQ.Fields("codaux")) & "' AND USUARIO = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "'"
                        oConexionMYSQL.Execute SQL
                        SQL = "update tempdetctacte_2 set numdoc = '-TI' WHERE CODAUX = '" & Trim(RQ.Fields("codaux")) & "' AND USUARIO = '" & strUsuarioId & "' and auxiliar = '" & CMP & "'"
                        oConexionMYSQL.Execute SQL
                    End If
                End If
            End If
            RQ.MoveNext
        Loop
    End If
    Screen.MousePointer = vbDefault
    Set myrs = Nothing
    Set RQ = Nothing
End Sub
Public Sub sp_Rep_EmpleadosxFechas(Division As String, Tipo As String, fecha As String)
    Dim SQL As String, STRDiv As String, STRDiv1 As String
    Dim myrs As MYSQL_RS, RQ As MYSQL_RS
    If Tipo = "H" Then
        If Division <> "0000" And Division <> "0006" Then
            STRDiv = " and C.divgas = '" & Division & "'"
            STRDiv1 = " and T.divgas = '" & Division & "'"
        End If
        SQL = "select e.codigo,CONCAT_WS(' ',e.apepat,e.apemat,e.nombre1,e.nombre2) as Nombre," & _
              "(select descrip from cnmdepar r where r.coddep=t.DIVGAS) as descripdiv,date_format(e.fec_ingreso,'%d/%m/%Y') as fecingreso,t.DIVGAS from (contrato t inner join " & _
              "(SELECT codemp,max(codigo) as cod FROM CONTRATO C WHERE C.F_INICIO <= '" & fecha & "' " & STRDiv & " group by codemp) as k " & _
              "on (t.codemp=k.codemp) and (t.codigo=k.cod)) left join empleado e on (t.codemp=e.codigo) where " & _
              "(t.F_TERMINO = '' or if(IfNULL(t.fechacese,'') = '',t.F_TERMINO >= '" & fecha & "',t.fechacese >= '" & fecha & "')) " & _
              "AND T.CODTIPO NOT IN ('05','06') " & STRDiv1 & _
              " order by CONCAT_WS(' ',e.apepat,e.apemat,e.nombre1,e.nombre2),T.DIVGAS"
    Else
        If Division <> "0000" And Division <> "0006" Then
            STRDiv = " and C.division = '" & Division & "'"
            STRDiv1 = " and T.division = '" & Division & "'"
        End If
        SQL = "select e.codigo,CONCAT_WS(' ',e.apepat,e.apemat,e.nombre1,e.nombre2) as Nombre," & _
              "(select descriplocal from cnmdepar r where r.coddep=t.division) as descripdiv,date_format(e.fec_ingreso,'%d/%m/%Y') as fecingreso,t.division from (contrato t inner join " & _
              "(SELECT codemp,max(codigo) as cod FROM CONTRATO C WHERE C.F_INICIO <= '" & fecha & "' " & STRDiv & " group by codemp) as k " & _
              "on (t.codemp=k.codemp) and (t.codigo=k.cod)) left join empleado e on (t.codemp=e.codigo) where " & _
              "(t.F_TERMINO = '' or if(IfNULL(t.fechacese,'') = '',t.F_TERMINO >= '" & fecha & "',t.fechacese >= '" & fecha & "')) " & _
              "AND T.CODTIPO NOT IN ('05','06') " & STRDiv1 & _
              " order by CONCAT_WS(' ',e.apepat,e.apemat,e.nombre1,e.nombre2),T.division"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Cesados(Division As String, Tipo As String, FechaIni As String, FechaFin As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim SCodEmp As String, SFecha As String, SWhere As String
    Screen.MousePointer = vbHourglass
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SFecha = " and (situacion = 0 and fec_cese BETWEEN '" & Format(FechaIni, "yyyy/mm/dd") & "' and '" & Format(FechaFin, "yyyy/mm/dd") & "')"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SFecha = " and (situacion = 0 and fec_cese = '" & Format(FechaIni, "yyyy/mm/dd") & "')"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SFecha = " and (situacion = 0 and fec_cese = '" & Format(FechaFin, "yyyy/mm/dd") & "')"
            Else
                If Division = "" Then SFecha = " and (situacion = 0 and fec_cese = DATE_FORMAT(sysdate(),'Y%/%m/%d'))"
            End If
        End If
    End If
    If Tipo = "H" Then
        If Division <> "" And Division <> "0006" Then
            SCodEmp = " and c.divgas = '" & Division & "'"
        End If
        SQL = "SELECT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo," & _
              "IFNULL(date_format(fec_cese,'%d/%m/%Y'),'') as fec_cese,DATE_FORMAT(max(f_inicio),'%d/%m/%Y') as f_inicio,date_format(max(f_termino),'%d/%m/%Y') as f_termino,IF(situacion=0,'CONTRATO(S) CANCELADO(S)','CONTRATO(S) POR VENCER') as situacion,e.tipo, " & _
              "C.divgas,(select descrip from cnmdepar r where r.coddep=c.divgas) as descripdiv,(select descrip from pl_tipocese p where p.codigo=c.codtipocese) as motivo FROM empleado e inner join contrato c on (e.codigo=c.codemp) where c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) " & SFecha & SCodEmp & "  " & _
              "group by e.codigo,e.fec_cese,situacion order by descripdiv,situacion,fec_cese,f_termino,nombres"
    Else
        If Division <> "" And Division <> "0006" Then
            SCodEmp = " and c.division = '" & Division & "'"
        End If
        SQL = "SELECT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo," & _
              "IFNULL(date_format(fec_cese,'%d/%m/%Y'),'') as fec_cese,DATE_FORMAT(max(f_inicio),'%d/%m/%Y') as f_inicio,date_format(max(f_termino),'%d/%m/%Y') as f_termino,IF(situacion=0,'CONTRATO(S) CANCELADO(S)','CONTRATO(S) POR VENCER') as situacion,e.tipo, " & _
              "C.division,(select descriplocal from cnmdepar r where r.coddep=c.division) as descripdiv,(select descrip from pl_tipocese p where p.codigo=c.codtipocese) as motivo FROM empleado e inner join contrato c on (e.codigo=c.codemp) where c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) " & SFecha & SCodEmp & "  " & _
              "group by e.codigo,e.fec_cese,situacion order by descripdiv,situacion,fec_cese,f_termino,nombres"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_EmpleadosPorMes(anomesini As String, anomesfin As String, Division As String)
    Dim SQL As String, STRDiv As String
    Dim myrs As MYSQL_RS
    Screen.MousePointer = vbHourglass
    If Division <> "0000" And Division <> "0006" Then
        STRDiv = " and b.division = '" & Division & "'"
    End If
    SQL = "select * from pl_tareo where anomes >= '" & anomesini & "' and anomes <= '" & anomesfin & "' and tipo = 1"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If Not myrs.EOF() Then
        SQL = "delete from tempdetctacte_2 where usuario = '" & strUsuarioId & "'"
        oConexionMYSQL.Execute SQL
        sp_Rep_EmpleadosxMesCargaIni anomesini, anomesfin, Division, ""
        SQL = "select CODAUX,NOMBRE,descripcuenta,date_format(e.fec_ingreso,'%d/%m/%Y') from tempdetctacte_2 t left join empleado e on (t.codaux=e.codigo) where " & _
              "abrevia = 'D' and usuario = '" & strUsuarioId & "' and auxiliar = '" & anomesini & "' AND IFNULL(CUENTA,'') <> 'CESADO' " & _
              "order by descripcuenta,nombre"
    Else
        SQL = "select e.codigo,concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres, " & _
              "(SELECT DESCRIP FROM cnmdepar d where d.coddep = c.division) as division, e.fec_ingreso " & _
              " from empleado e left join contrato c on (e.codigo=c.codemp) where c.estado = 'AP' and e.tipo not in (3,4) " & _
              " and e.situacion = 1 ORDER BY division,nombres"
        m_Titulo = "TOTAL DE EMPLEADOS AL " & Day(Date) & " DE " & NombreMes(Right("00" & Month(Date), 2), False) & " DEL " & Year(Date)
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_ContratosxVencer(Division As String, CodEmp As String, Tipo As String, FechaIni As String, FechaFin As String, dias As String)
Dim SQL As String, Str1 As String, str2 As String, str3 As String, Str4 As String, Str5 As String
Dim myrs As MYSQL_RS
Dim RQ As MYSQL_RS
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        str3 = " and If(c.fechacese<>'',c.fechacese,c.f_termino) >='" & Format(FechaIni, "yyyy/mm/dd") & "' and If(c.fechacese<>'',c.fechacese,c.f_termino) <= DATE_FORMAT(DATE_ADD('" & Format(FechaFin, "yyyy/mm/dd") & "',INTERVAL " & val(dias) & " DAY),'%Y/%m/%d') "
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            str3 = " and If(c.fechacese<>'',c.fechacese,c.f_termino) >='" & Format(FechaIni, "yyyy/mm/dd") & "' and If(c.fechacese<>'',c.fechacese,c.f_termino) <= DATE_FORMAT(DATE_ADD('" & Format(FechaIni, "yyyy/mm/dd") & "',INTERVAL " & val(dias) & " DAY),'%Y/%m/%d')"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                str3 = " and If(c.fechacese<>'',c.fechacese,c.f_termino) >='" & Format(FechaFin, "yyyy/mm/dd") & "' and If(c.fechacese<>'',c.fechacese,c.f_termino) <= DATE_FORMAT(DATE_ADD('" & Format(FechaFin, "yyyy/mm/dd") & "',INTERVAL " & val(dias) & " DAY),'%Y/%m/%d')"
            Else
                str3 = " and If(c.fechacese<>'',c.fechacese,c.f_termino) >='" & Format(Date, "yyyy/mm/dd") & "' and If(c.fechacese<>'',c.fechacese,c.f_termino) <= DATE_FORMAT(DATE_ADD('" & Format(Date, "yyyy/mm/dd") & "',INTERVAL " & val(dias) & " DAY),'%Y/%m/%d')"
            End If
        End If
    End If
    If Tipo = "H" Then
        If Division <> "000000000000" And Division <> "000000000006" Then
            Str1 = " and c.divgas = '" & Division & "'"
        End If
        Str4 = "c.divgas"
        Str5 = "descrip"
    Else
        If Division <> "000000000000" And Division <> "000000000006" Then
            Str1 = " and c.division = '" & Division & "'"
        End If
        Str4 = "c.division"
        Str5 = "descriplocal"
    End If
    If Division <> "000000000000" And Division <> "000000000006" Then
        If CodEmp <> "" Then str2 = " and e.codigo = '" & CodEmp & "'"
    End If
    Screen.MousePointer = vbHourglass
    SQL = "DELETE FROM rh_templiquidacion where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQL
    SQL = "insert into rh_templiquidacion(nombres,codigo,fecingreso,feccese,fecultcts,bancocts,fecinicts,motivo,moncts,basico,usuario,fecinivac,countcont) " & _
          "SELECT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo, " & _
          "DATE_FORMAT(max(f_inicio),'%d/%m/%Y') as f_inicio, " & _
          "date_format(max(If(c.fechacese<>'',c.fechacese,c.f_termino)),'%d/%m/%Y') as f_termino," & Str4 & ", " & _
          "(select " & Str5 & " from cnmdepar r where r.coddep=" & Str4 & ") as descripdiv,max(If(c.fechacese<>'',c.fechacese,c.f_termino)) as fin, " & _
          "IF(max(If(c.fechacese<>'',c.fechacese,c.f_termino)) > '" & FechaFin & "','SIGUIENTE(S) " & dias & " DIA(S)','DEL " & Format(FechaIni, "dd/mm/yyyy") & "  AL  " & Format(FechaFin, "dd/mm/yyyy") & "') as anomes,e.situacion, " & _
          "MONTH(max(If(c.fechacese<>'',c.fechacese,c.f_termino))) as mes,'" & strUsuarioId & "', " & _
          "(select min(f_inicio) as inicio from contrato o where sbruto = 1 and o.codemp =e.codigo) as finicio,(select count(codigo) from contrato o where sbruto = 1 and o.codemp =e.codigo) as countcont FROM empleado e inner join contrato c on (e.codigo=c.codemp) where 1=1 " & str3 & _
          "and e.situacion = 1 and c.estado='AP' and e.tipo not in (3,4) " & Str1 & str2 & _
          " group by e.codigo,e.fec_cese,situacion order by anomes,mes,FIN,nombres"
    oConexionMYSQL.Execute SQL
    SQL = "select codigo,DATE_FORMAT(fecinivac,'%d/%m/%Y') as fec_ingreso,feccese as f_termino,countcont from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    Dim DifAnios As Integer, DifMeses As Integer, Difdias As Integer
    Dim FecTemp As String, FecFin As String, FecIni As String, RQ1 As MYSQL_RS
    Dim I As Integer, SumAnios As Integer, SumMeses As Integer, Sumdias As Integer
    Do While Not RQ.EOF()
        SQL = "select a.codigo,sum(ifnull(DATEDIFF(ifnull(c.f_termino,''),c.f_inicio)+1,0)) as dias from empleado as a " & _
              "left join contrato as c on(c.codemp=a.codigo) Where c.sbruto = 1 and a.codigo = '" & RQ.Fields("codigo") & "' group by a.codigo"
        Set RQ1 = oConexion.EjecutaSelectRS(SQL)
        If Not RQ1.EOF() Then
            If CDbl(RQ1.Fields("dias")) <> CDbl(DateDiff("d", CDate(RQ.Fields("fec_ingreso")), CDate(RQ.Fields("f_termino")))) + 1 Then
                SumAnios = 0: SumMeses = 0: Sumdias = 0: I = 0
                SQL = "select codigo,f_Inicio,f_termino,o.ult from contrato C LEFT JOIN (SELECT codemp,MAX(CODIGO) as ult FROM " & _
                      "CONTRATO  group by codemp) as o on (c.codemp=o.codemp) where codtipo<>'02' and c.sbruto = 1 and C.codemp = '" & RQ.Fields("codigo") & "' order by codigo"
                Set RQ1 = oConexion.EjecutaSelectRS(SQL)
                Do While Not RQ1.EOF()
                    If I = 0 Then
                        FecIni = CDate(RQ1.Fields("f_inicio"))
                        FecFin = CDate(RQ1.Fields("f_termino"))
                        RQ1.MoveNext
                        I = I + 1
                    Else
aca:
                        If CDate(Format(FecFin, "dd/mm/yyyy")) + 1 = CDate(Format(RQ1.Fields("f_inicio"), "dd/mm/yyyy")) Then
                            If RQ1.Fields("ult") = RQ1.Fields("codigo") Then
                                FecFin = Format(RQ.Fields("f_termino"), "dd/mm/yyyy")
                                GoTo AQUI
                            End If
                            FecFin = CDate(RQ1.Fields("f_termino"))
                            RQ1.MoveNext
                            I = I + 1
                        Else
AQUI:
                            FecTemp = Mid(Format(FecIni, "dd/mm/yyyy"), 1, 6) & Year(FecFin)
                            DifAnios = Year(FecFin) - Year(CDate(Format(FecIni, "dd/mm/yyyy")))
                            DifMeses = 0: Difdias = 0
                            If CDate(FecTemp) > FecFin Then
                                DifAnios = DifAnios - 1
                                FecTemp = Mid(FecTemp, 1, 6) & (Year(CDate(FecTemp)) - 1)
                            End If
                            If Year(CDate(FecTemp)) < Year(FecFin) Then
                                DifMeses = (12 - Month(CDate(FecTemp))) + Month(FecFin)
                            ElseIf Year(CDate(FecTemp)) = Year(FecFin) Then
                                DifMeses = Month(FecFin) - Month(CDate(FecTemp))
                            End If
                            FecTemp = Mid(FecTemp, 1, 3) & Right("00" & Month(CDate(FecFin)), 2) & "/" & Year(FecFin)
                            If Day(FecFin) > Day(CDate(FecTemp)) Then
                                Difdias = Day(FecFin) - Day(CDate(FecTemp))
                            ElseIf Day(FecFin) < Day(CDate(FecTemp)) Then
                                DifMeses = DifMeses - 1
                                If Day(CDate(FecTemp)) = 31 Then
                                    If Day(DateSerial(Year(CDate("01/" & Right("00" & Month(CDate(FecFin)) - 1, 2) & "/" & Year(FecFin))), Month(CDate("01/" & Right("00" & Month(CDate(FecFin)) - 1, 2) & "/" & Year(FecFin))) + 1, 0)) < Day(CDate(FecTemp)) Then
                                        FecTemp = "30/" & Right("00" & Month(CDate(FecFin)) - 1, 2) & "/" & Year(FecFin)
                                    End If
                                Else
                                    If Month(CDate(FecFin)) = 1 Then
                                        FecTemp = Mid(FecTemp, 1, 3) & "12" & "/" & (Year(FecFin) - 1)
                                    Else
                                    
                                        FecTemp = Mid(FecTemp, 1, 3) & Right("00" & Month(CDate(FecFin)) - 1, 2) & "/" & Year(FecFin)
                                    End If
                                End If
                                Difdias = (Day(DateSerial(Year(CDate(FecTemp)), Month(CDate(FecTemp)) + 1, 0)) - Day(CDate(FecTemp))) + Day(FecFin)
                            End If
                            SumAnios = SumAnios + DifAnios
                            SumMeses = SumMeses + DifMeses
                            Sumdias = Sumdias + Difdias
                            FecIni = CDate(RQ1.Fields("f_inicio"))
                            If FecFin <> Format(RQ.Fields("f_termino"), "dd/mm/yyyy") Then
                                If IsDate(RQ1.Fields("f_termino")) Then
                                    FecFin = CDate(RQ1.Fields("f_termino"))
                                Else
                                    FecFin = Format(RQ.Fields("f_termino"), "dd/mm/yyyy")
                                    RQ1.MovePrevious
                                End If
                            End If
                            RQ1.MoveNext
                            I = I + 1
                        End If
                    End If
                Loop
                If I = 1 Then
                    RQ1.MovePrevious
                    GoTo aca
                End If
                If Sumdias >= 30 Then
                    SumMeses = SumMeses + Int(val(Sumdias / 30))
                    Sumdias = Sumdias - (30 * Int(val(Sumdias / 30)))
                End If
                If SumMeses >= 12 Then
                    SumAnios = SumAnios + Int(val(SumMeses / 12))
                    SumMeses = SumMeses - (12 * Int(val(SumMeses / 12)))
                End If
                SQL = "update rh_templiquidacion set anios = " & SumAnios & ",meses = " & SumMeses & ",aniosvac=" & Sumdias & " " & _
                      "where usuario = '" & strUsuarioId & "' and codigo = '" & RQ.Fields("codigo") & "'"
                oConexionMYSQL.Execute SQL
            Else
                FecTemp = Mid(RQ.Fields("fec_ingreso"), 1, 6) & Year(CDate(RQ.Fields("f_termino")))
                DifAnios = Year(CDate(RQ.Fields("f_termino"))) - Year(CDate(RQ.Fields("fec_ingreso")))
                DifMeses = 0: Difdias = 0
                If (CDate(FecTemp) > CDate(RQ.Fields("f_termino"))) Then
                    DifAnios = DifAnios - 1
                    FecTemp = Mid(FecTemp, 1, 6) & (Year(CDate(FecTemp)) - 1)
                End If
                If Year(CDate(FecTemp)) < Year(CDate(RQ.Fields("f_termino"))) Then
                    DifMeses = (12 - Month(CDate(FecTemp))) + Month(CDate(RQ.Fields("f_termino")))
                ElseIf Year(CDate(FecTemp)) = Year(CDate(RQ.Fields("f_termino"))) Then
                    DifMeses = Month(CDate(RQ.Fields("f_termino"))) - Month(CDate(FecTemp))
                End If
                FecTemp = Mid(FecTemp, 1, 3) & Right("00" & Month(CDate(RQ.Fields("f_termino"))), 2) & "/" & Year(CDate(RQ.Fields("f_termino")))
                
                ' si la fecha no es valida disminuir en uno
                If Not IsDate(FecTemp) Then
                    FecTemp = Mid(FecTemp, 1, 2) - 1 & "/" & Right("00" & Month(CDate(RQ.Fields("f_termino"))), 2) & "/" & Year(CDate(RQ.Fields("f_termino")))
                    ' si la fecha no es valida disminuir en uno
                    If Not IsDate(FecTemp) Then
                        FecTemp = Mid(FecTemp, 1, 2) - 1 & "/" & Right("00" & Month(CDate(RQ.Fields("f_termino"))), 2) & "/" & Year(CDate(RQ.Fields("f_termino")))
                    End If
                    
                    'sumar uno pero utilizando funcion de vb
                    FecTemp = DateAdd("d", 1, FecTemp)
                End If
                
                If Day(CDate(RQ.Fields("f_termino"))) > Day(CDate(FecTemp)) Then
                    Difdias = Day(CDate(RQ.Fields("f_termino"))) - Day(CDate(FecTemp))
                ElseIf Day(CDate(RQ.Fields("f_termino"))) < Day(CDate(FecTemp)) Then
                    DifMeses = DifMeses - 1
                    If Day(CDate(FecTemp)) = 31 Then
                        If Day(DateSerial(Year(CDate("01/" & Right("00" & Month(CDate(RQ.Fields("f_termino"))) - 1, 2) & "/" & Year(CDate(RQ.Fields("f_termino"))))), Month(CDate("01/" & Right("00" & Month(CDate(RQ.Fields("f_termino"))) - 1, 2) & "/" & Year(CDate(RQ.Fields("f_termino"))))) + 1, 0)) < Day(CDate(FecTemp)) Then
                            FecTemp = "30/" & Right("00" & Month(CDate(RQ.Fields("f_termino"))) - 1, 2) & "/" & Year(CDate(RQ.Fields("f_termino")))
                        End If
                    Else
                        If Month(CDate(RQ.Fields("f_termino"))) = 1 Then
                            FecTemp = Mid(FecTemp, 1, 3) & "12" & "/" & Year(CDate(RQ.Fields("f_termino"))) - 1
                        Else
                            FecTemp = Mid(FecTemp, 1, 3) & Right("00" & Month(CDate(RQ.Fields("f_termino"))) - 1, 2) & "/" & Year(CDate(RQ.Fields("f_termino")))
                        End If
                    End If
                    
                    Difdias = (Day(DateSerial(Year(CDate(FecTemp)), Month(CDate(FecTemp)) + 1, 0)) - Day(CDate(FecTemp))) + Day(CDate(RQ.Fields("f_termino")))
                End If
                
'                If (Year(CDate(RQ.Fields("f_termino"))) - Year(CDate(RQ.Fields("fec_ingreso")))) >= 1 And CInt(RQ.Fields("countcont")) = 1 Then
'                DifAnios = DifAnios + 1
'                End If
                
                SQL = "update rh_templiquidacion set anios = " & DifAnios & ",meses = " & DifMeses & ",aniosvac=" & Difdias & " " & _
                      "where usuario = '" & strUsuarioId & "' and codigo = '" & RQ.Fields("codigo") & "'"
                oConexionMYSQL.Execute SQL
            End If
        End If
        RQ.MoveNext
    Loop
    SQL = "select rtrim(r.nombres) as nombres,r.codigo,r.fecingreso as f_inicio,r.feccese as f_termino,r.fecultcts as division, r.bancocts as descripdiv, " & _
          "r.fecinicts as fin,r.motivo as anomes,r.moncts as situacion,r.basico as mes,r.anios,r.meses,r.aniosvac as dias,(select descrip from cncargos where codigo=c.codcargo) as cargo " & _
          "from rh_templiquidacion as r left join contrato as c on r.codigo=c.codemp where r.usuario = '" & strUsuarioId & "' and c.estado='AP' order by r.motivo,r.basico,r.fecinicts,r.nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Datos de Contratos por Vencer", vbInformation, "NOVADMIN"
        Screen.MousePointer = vbDefault
        Set myrs = Nothing
        Set RQ = Nothing
        Set RQ1 = Nothing
        Exit Sub
    End If
    If MsgBox("¿Desea ver el Reporte en excel?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then
        On Error GoTo CtrlExporta
        Dim o_Excel As Object
        Dim o_Libro As Object
        Dim Nom As String
        Set o_Excel = New Excel.Application
        Set o_Libro = o_Excel.Workbooks.Add
        SQL = "SELECT CODDEP," & Str5 & " AS DESCRIP FROM CNMDEPAR "
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        CantCta = 3
        Do While Not RQ.EOF()
            SQL = "SELECT (select " & Str5 & " from cnmdepar r where r.coddep=" & Str4 & ") as division,e.codigo," & _
                  "concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres," & _
                  "DATE_FORMAT(max(f_inicio),'%d/%m/%Y') as f_inicio_ult_cont, date_format(max(f_termino),'%d/%m/%Y') as f_termino_ult_cont, " & _
                  "IF(max(f_termino) > '" & FechaFin & "','SIGUIENTE(S) " & dias & " DIA(S)','DEL " & Format(FechaIni, "dd/mm/yyyy") & "  AL  " & Format(FechaFin, "dd/mm/yyyy") & "') as anomes,if(situacion=1,'ACTIVO','CESADO') as situ, " & _
                  "MONTH(max(f_termino)) as mes FROM empleado e inner join contrato c on (e.codigo=c.codemp) where 1=1 " & str3 & _
                  "and situacion = 1 and e.tipo not in (3,4) and " & Str4 & " = '" & Trim(RQ.Fields("coddep")) & "' " & str2 & _
                  " group by e.codigo,e.fec_cese,situacion order by anomes,division,f_termino_ult_cont ,nombres"
            Set myrs = oConexion.EjecutaSelectRS(SQL)
            If Not myrs.EOF() Then
                Exportar_Excel_rs_conLibro myrs, o_Excel, o_Libro
                Nom = cNombreHoja & CantCta + 1
                CantCta = CantCta + 1
                o_Libro.Sheets(Nom).name = Trim(RQ.Fields("DESCRIP"))
            End If
            RQ.MoveNext
        Loop
        SQL = "SELECT (select " & Str5 & " from cnmdepar r where r.coddep=" & Str4 & ") as division,e.codigo," & _
              "concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres," & _
              "DATE_FORMAT(max(f_inicio),'%d/%m/%Y') as f_inicio_ult_cont, date_format(max(f_termino),'%d/%m/%Y') as f_termino_ult_cont, " & _
              "IF(max(f_termino) > '" & FechaFin & "','SIGUIENTE(S) " & dias & " DIA(S)','DEL " & Format(FechaIni, "dd/mm/yyyy") & "  AL  " & Format(FechaFin, "dd/mm/yyyy") & "') as anomes, " & _
              "month(max(f_termino)) as mes FROM empleado e inner join contrato c on (e.codigo=c.codemp) where 1 = 1 " & str3 & _
              "and situacion = 1 and c.estado = 'AP' and e.tipo not in (3,4) " & Str1 & str2 & _
              " group by e.codigo,e.fec_cese,situacion order by anomes,division,f_termino_ult_cont ,nombres"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
        Exportar_Excel_rs_conLibro myrs, o_Excel, o_Libro
        Nom = cNombreHoja & CantCta + 1
        CantCta = CantCta + 1
        o_Libro.Sheets(Nom).name = "TODOS"
        
        o_Excel.Visible = True
        o_Libro.SaveAs Rep_Documents & "\" & "CONTRATOS_POR_VENCER.xls"
        If Not o_Excel Is Nothing Then Set o_Excel = Nothing
        If Not o_Libro Is Nothing Then Set o_Libro = Nothing
        Screen.MousePointer = vbDefault
Exit Sub
CtrlExporta:
        Screen.MousePointer = vbDefault
        If Not o_Libro Is Nothing Then: o_Libro.Close False
        If Not o_Excel Is Nothing Then: o_Excel.Quit
    Else
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
        Screen.MousePointer = vbDefault
    End If
    Set RQ = Nothing
    Set RQ1 = Nothing
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Liquidacion(CodEmp As String, CodMotivo As String)
    Dim SQL As String, fec As String, Anio As Long, PorAfp As Double, FlgIng As Boolean
    Dim Mes As String
    Dim Dif As Integer, I As Integer
    Dim DiasPend As Double, MesPend As Double, TotPend As Double, TotDif As Double
    Dim FecVaca As String, MtoGracia As Double, DiasPendientes As Integer
    Dim StrAnios As String, MtoBono As Double, Motivo As String, AsigFam As Double
    Dim toting As Double, MtoQuinta As Double, Mtoreten As Double, k As Long
    Dim RQ As MYSQL_RS
    SQL = "select * from pl_liquidacion where codemp = '" & CodEmp & "' and cerrado = 'S'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    
    On Error GoTo SERROR
    If RQ.EOF Then
        fec = InputBox("Ingrese la Fecha de Cese", "LIQUIDACION", CStr(Date))
        If fec = "" Then Exit Sub
        fec = Format(fec, "yyyy/mm/dd")
        Motivo = DescripcionesdeCodigos("TIPOCESE", CodMotivo, "DESCRIP")
        MtoBono = InputBox("Ingrese Monto de Entregas por Rendir del Trabajador", "LIQUIDACION", 0)
        MtoBono = IIf(IsNull(MtoBono) Or val(MtoBono) = 0, 0, MtoBono)
        Mtoreten = InputBox("Ingrese Monto para Dscto por Retenciones Judiciales", "LIQUIDACION", 0)
        Mtoreten = IIf(IsNull(Mtoreten) Or val(Mtoreten) = 0, 0, Mtoreten)
        MtoGracia = InputBox("Ingrese Monto a Título de Gracia", "LIQUIDACION", 0)
        MtoGracia = IIf(IsNull(MtoGracia) Or val(MtoGracia) = 0, 0, MtoGracia)
        DiasPendientes = InputBox("Ingrese Dias pendientes de pago", "LIQUIDACION", 0)
        DiasPendientes = IIf(IsNull(DiasPendientes) Or val(DiasPendientes) = 0, 0, DiasPendientes)
        SQL = "DELETE FROM rh_templiquidacion WHERE usuario = '" & strUsuarioId & "'"
        oConexionMYSQL.Execute SQL
        If Not IsDate(fec) Then
            MsgBox "No Existen Registros", vbInformation, gsNomSW
            Set RQ = Nothing
            Exit Sub
        End If
        
        '"truncate(PERIOD_DIFF(date_format('" & fec & "','%Y%m'),date_format(fec_ingreso,'%Y%m')) / 12,0) as anios," &
        
        SQL = "insert into rh_templiquidacion (codigo,nombres,fecingreso,feccese,tiempo,anios,meses,fecultcts,basico,ultgrati,usuario,mesgrati, " & _
              "ctacts,bancocts,moncts,fecult,afp,dni,mtoentregas,prestamos,adelantos,sctr,mtoreten,mtogracia,diaspen) " & _
              "select e.codigo,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres,fec_ingreso,'" & fec & "' AS fec_cese, " & _
              "PERIOD_ADD(date_format(fec_ingreso,'%Y%m'),PERIOD_DIFF(date_format('" & fec & "','%Y%m'),date_format(fec_ingreso,'%Y%m'))) as tiempo," & _
              "truncate(datediff('" & fec & "',fec_ingreso  ) / DAYOFYEAR(concat(year('" & fec & "'),'/12/31')),0)  as anios," & _
              "PERIOD_DIFF(date_format('" & fec & "','%Y%m'),date_format(fec_ingreso,'%Y%m')) - (truncate(PERIOD_DIFF(date_format('" & fec & "','%Y%m'),date_format(fec_ingreso,'%Y%m')) / 12,0) * 12) as meses," & _
              "DATE_FORMAT(if(MONTH('" & fec & "') < 5,LAST_DAY(concat(year('" & fec & "'),'/10/01')), " & _
              "if(MONTH('" & fec & "') > 5,LAST_DAY(concat(year('" & fec & "'),'/10/01')),IF(MONTH('" & fec & "') = 5,'" & strAnoSistema & "/04/30',''))),'%Y/%m/%d') as ultctsdep, " & _
              "c.sbasico,IFNULL((SELECT SUM(if(p.unidad = 'V',CANT,round((sbasico/180)*cant,2))) FROM PL_TAREO P " & _
              "inner join pl_rubrosremunerativos r on (p.rub=r.codigo) left join pl_tiporubros t on (r.tipo=t.codigo) " & _
              "Where p.emp = e.codigo AND ANOMES = IFNULL(if(month('" & fec & "') < 7,concat(year('" & fec & "')-1,'12'), " & _
              "if(month('" & fec & "') >= 7,concat(year('" & fec & "'),'07'),'')),'') AND p.tIPO = 5 and t.tipo = 'I' and rub not in ('1003')),0) as ultgrati,'" & strUsuarioId & "', " & _
              "IFNULL(if(month('" & fec & "') < 7,'12',if(month('" & fec & "') > 7,'07','')),'') as mesgrati,e.ctsnumcta, " & _
              "(select descrip from pl_entidadfinanciera l where l.codigo=e.ctsbanco) as ctsbanco,e.ctsmon, " & _
              "DATE_FORMAT(LAST_DAY(CONCAT(YEAR('" & fec & "'),'-',left(right('" & fec & "',5),2),'-',day('" & fec & "'))),'%Y/%m/%d') as ultfec, " & _
              "(select nombre from afp f where f.codigo=e.codafp) as afp,e.numdocide," & MtoBono & ",(SELECT IFNULL(SUM(IMPORTE)-SUM(MONTOPAGADO),0) AS MTO " & _
              "FROM DOCUMENTOS_RRHH WHERE solicita = e.codigo and CODDOC = 'PR' AND MONTOPAGADO < IMPORTE) as mtopr, " & _
              "(SELECT IFNULL(SUM(IMPORTE)-SUM(MONTOPAGADO),0) AS MTO FROM DOCUMENTOS_RRHH WHERE solicita = e.codigo and " & _
              "CODDOC = 'SS' AND MONTOPAGADO < IMPORTE) as mtoad,e.sctr," & Mtoreten & "," & MtoGracia & "," & DiasPendientes & " from empleado e left join contrato c on (e.codigo=c.codemp) inner join (select a.codigo, " & _
              "if((DATEDIFF('" & fec & "',fec_ingreso)+1)<0,0,(DATEDIFF('" & fec & "',fec_ingreso)+1)) dias " & _
              "from empleado as a order by a.apepat,a.apemat,a.nombre1,a.nombre2) AS K ON (e.codigo=k.codigo) " & _
              "where e.codigo = '" & CodEmp & "' and C.CODIGO = (SELECT MAX(CODIGO) FROM contrato t WHERE t.codemp=e.codigo)"
        oConexionMYSQL.Execute SQL
        
        If Month(fec) = 7 Then
            SQL = "UPDATE RH_TEMPLIQUIDACION SET" & _
                  " mesgrati=if((select ifnull(rem,0) from `pl_remanterior` where emp='" & CodEmp & "' and periodo='200907' and tipo='GJ')>0,0,mesgrati)" & _
                  ", ultgrati=if((select ifnull(rem,0) from `pl_remanterior` where emp='" & CodEmp & "' and periodo='200907' and tipo='GJ')>0,0,ultgrati) WHERE CODIGO='" & CodEmp & "' AND USUARIO='" & strUsuarioId & "' AND FECCESE='" & fec & "'"
        End If
        
        
        If Month(fec) = 12 Then
            SQL = "UPDATE RH_TEMPLIQUIDACION SET" & _
                  " mesgrati=if((select ifnull(rem,0) from `pl_remanterior` where emp='" & CodEmp & "' and periodo='200912' and tipo='GD')>0,0,mesgrati)" & _
                  ", ultgrati=if((select ifnull(rem,0) from `pl_remanterior` where emp='" & CodEmp & "' and periodo='200912' and tipo='GD')>0,0,ultgrati) WHERE CODIGO='" & CodEmp & "' AND USUARIO='" & strUsuarioId & "' AND FECCESE='" & fec & "'"
        End If
        oConexionMYSQL.Execute SQL
        
        SQL = "select * from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Dim RQ1 As MYSQL_RS
            Dim MontoBono1 As Double
            MtoBono = 0
            SQL = "SELECT IFNULL(sum(cant),0) as bonos FROM PL_TAREO P WHERE P.EMP='" & CodEmp & "' AND P.RUB = '1001' and " & _
                  " ANOMES = '" & Year(fec) & Right("00" & Month(fec), 2) & "' AND TIPO = 1"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            If Not RQ1.EOF() Then
                If RQ1.Fields("bonos") = 0 Then
                    If Day(CDate(fec)) > 15 Then
                        SQL = "SELECT IFNULL(sum(cant),0) as bonos FROM PL_TAREO P WHERE P.EMP='" & CodEmp & "' AND P.RUB = '1001' and " & _
                              " ANOMES = '" & Year(fec) & Right("00" & Month(fec), 2) & "' AND TIPO = 2"
                        Set RQ1 = oConexion.EjecutaSelectRS(SQL)
                        If Not RQ1.EOF() Then
                            MtoBono = RQ1.Fields("bonos")
                        End If
                    End If
                Else
                    MontoBono1 = RQ1.Fields("bonos")
                End If
            End If
            Mes = Left(Right(fec, 5), 2)
            If MontoBono1 > 0 Then
                Dif = val(Mes) - IIf(MtoBono > 0, 6, 7) + 1
            Else
                Dif = val(Mes) - IIf(MtoBono > 0, 6, 7)
            End If
            If Dif >= 0 Then
                Mes = Right("00" & (Dif + 1), 2)
                Anio = Year(fec)
            Else
                Mes = Right("00" & (Dif + 13), 2)
                Anio = Year(fec) - 1
            End If
            SQL = "SELECT IFNULL(promedio,0) as bonos FROM pl_prombonos P WHERE P.EMP='" & CodEmp & "'"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            If Not RQ1.EOF() Then
                MtoBono = Round(RQ1.Fields("bonos"))
            End If
            StrAnios = Right("00" & Day(RQ.Fields("fecingreso")), 2) & "/" & Right("00" & Trim(RQ.Fields("tiempo")), 2) & "/" & Left(Trim(RQ.Fields("tiempo")), 4)
            If CDate(StrAnios) > CDate(Format(fec, "dd/mm/yyyy")) Then
                If Mid(StrAnios, 1, 5) = Mid(Format(RQ.Fields("fecingreso"), "dd/mm/yyyy"), 1, 5) Then
                    I = DateDiff("d", CDate(Mid(StrAnios, 1, 3) & Right("00" & (val(Mid(StrAnios, 4, 2)) - 1), 2) & Mid(StrAnios, 6, 5)), CDate(Format(fec, "dd/mm/yyyy")))
                    StrAnios = IIf(val(RQ.Fields("anios")) = 0, "", IIf(val(RQ.Fields("anios")) = 1, "1 año ", RQ.Fields("anios") - 1 & " años ")) & IIf(val(RQ.Fields("meses")) = 0, "11 meses ", IIf(val(RQ.Fields("meses")) = 1, "", val(RQ.Fields("meses")) - 1 & " meses ")) & _
                               IIf(I = 0, "", IIf(I = 1, " 1 dia", I & " dias"))
                Else
                    I = DateDiff("d", CDate(val(Mid(StrAnios, 1, 2)) + 1 & "/" & Right("00" & (val(Mid(StrAnios, 4, 2)) - 1), 2) & Mid(StrAnios, 6, 5)), CDate(Format(fec, "dd/mm/yyyy"))) '(Day(CDate(StrAnios)) - Day(CDate(Format(fec, "dd/mm/yyyy")))) + 1 '(30 - Day(RQ.Fields("FecIngreso"))) + Day(fec)
                    StrAnios = IIf(val(RQ.Fields("anios")) = 0, "", IIf(val(RQ.Fields("anios")) = 1, "1 año ", RQ.Fields("anios") & " años ")) & IIf(val(RQ.Fields("meses")) = 0, "", IIf(val(RQ.Fields("meses")) = 1, "", val(RQ.Fields("meses")) - 1 & " meses ")) & _
                               IIf(I = 0, "", IIf(I = 1, " 1 dia", I & " dias"))
                End If
            ElseIf CDate(StrAnios) < CDate(Format(fec, "dd/mm/yyyy")) Then
                I = (IIf(Day(fec) = "31" Or (Day(fec) = 28 And Month(fec) = 2), "30", Day(fec)) - Day(RQ.Fields("fecingreso"))) + 1
                Dif = val(RQ.Fields("meses"))
                k = CLng(RQ.Fields("anios"))
                If I = 30 Then
                    I = 0
                    Dif = val(RQ.Fields("meses")) + 1
                    If Dif = 12 Then
                        Dif = 0
                        k = CLng(RQ.Fields("anios")) + 1
                    End If
                End If
                StrAnios = IIf(val(k) = 0, "", IIf(val(k) = 1, "1 año ", k & " años ")) & IIf(val(Dif) = 0, "", IIf(val(Dif) = 1, " 1 mes ", Dif & " meses ")) & _
                           IIf(I = 0, "", IIf(I = 1, " 1 día", I & " dias"))
            ElseIf CDate(StrAnios) = CDate(Format(fec, "dd/mm/yyyy")) Then
                StrAnios = IIf(val(RQ.Fields("anios")) = 0, "", IIf(val(RQ.Fields("anios")) = 1, "1 año ", RQ.Fields("anios") & " años ")) & IIf(val(RQ.Fields("meses")) = 0, "", IIf(val(RQ.Fields("meses")) = 1, " 1 mes ", RQ.Fields("meses") & " meses "))
            End If
            SQL = "SELECT ifnull(sum(monto_viatico),0) as numvac FROM CALENDARIO C WHERE C.MOVEMP = '02' AND C.CODEMP = '" & CodEmp & "' ORDER BY PERIODO"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            
            If Not RQ1.EOF() Then
                FecVaca = ""
                If (val(RQ1.Fields("numvac")) Mod 30) <> 0 Then
                
                    '--- calculo de las vacaciones truncas ---'
                    DiasPend = (val(RQ.Fields("anios")) * 30 - val(RQ1.Fields("numvac"))) * 12 'dias truncos expresados en 12 dias
                    MesPend = val(RQ.Fields("meses")) * 30  'meses pendientes truncos
                    TotPend = Abs(DiasPend + MesPend) + 1 'total de dias pendientes expresados en 12 dias
                    FecVaca = DateAdd("d", TotPend * -1, fec) 'retrocedemos los dias pendientes a la fecha de cese
                    '-----------------------------------------'
                    
                    'Dif = Int(val(RQ1.Fields("numvac")) / 30)
                    'I = Int((val(RQ1.Fields("numvac")) - (Dif * 30)) * 12 / 30)
                    'FecVaca = DateAdd("yyyy", Dif, Format(RQ.Fields("fecingreso"), "dd/mm/yyyy"))
                    'FecVaca = DateAdd("m", I, Format(FecVaca, "dd/mm/yyyy"))
                    
                Else
                    FecVaca = DateAdd("yyyy", val(RQ1.Fields("numvac")) / 30, Format(RQ.Fields("fecingreso"), "dd/mm/yyyy"))
                End If
            
            End If
            'Anio = Int(DateDiff("m", FecVaca, fec) / 12)
            Anio = Int(DateDiff("Y", FecVaca, fec) / Abs(DateDiff("Y", CStr(Year(fec)) & "/12/31", CStr(Year(fec)) & "/01/01")))
            
            Dif = DateDiff("m", FecVaca, fec) - (Anio * 12)
            Mes = DateAdd("m", DateDiff("m", FecVaca, fec), FecVaca)
            If CDate(Mes) > CDate(Format(fec, "dd/mm/yyyy")) Then
                I = DateDiff("d", CDate(val(Mid(Mes, 1, 2)) & "/" & Right("00" & (val(Mid(Mes, 4, 2)) - 1), 2) & Mid(Mes, 6, 5)), CDate(Format(fec, "dd/mm/yyyy")))
                'I = DateDiff("d", RQ.Fields("fecingreso"), fec)
                Dif = Dif - 1
            ElseIf CDate(Mes) < CDate(Format(fec, "dd/mm/yyyy")) Then
                I = (IIf(Day(fec) = "31" Or (Day(fec) = 28 And Month(fec) = 2), "30", Day(fec)) - (Day(FecVaca))) + 1
                If I = 30 Then
                    Dif = Dif + 1
                    I = 0
                    If Dif = 12 Then
                        Dif = 0
                        Anio = Anio + 1
                    End If
                End If
            ElseIf CDate(Mes) = CDate(Format(fec, "dd/mm/yyyy")) Then
                I = 0
            End If
            If I = Day(DateSerial(Year(fec), Month(fec) + 1, 0)) Then
                Dif = Dif + 1
                I = I - Day(fec)
            End If
            SQL = "select ifnull(porcdscto + porcseg + f.porccom + if((select sctr from empleado e where e.codigo = '" & CodEmp & "') = 'S', " & _
                  "2,0),0) as por from pl_afppar f where anomes = '" & Year(fec) & Right("00" & Month(fec), 2) & "' and " & _
                  "codafp = (select codafp from empleado e where e.codigo = '" & CodEmp & "')"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            If Not RQ1.EOF() Then
                PorAfp = RQ1.Fields("por")
            Else
                SQL = "select codafp from empleado e where e.codigo = '" & CodEmp & "'"
                Set RQ1 = oConexion.EjecutaSelectRS(SQL)
                If Not RQ1.EOF() Then
                    If RQ1.Fields("CODAFP") = "06" Then
                        PorAfp = 13
                    End If
                End If
            End If
            SQL = "SELECT asigfam from empleado where codigo = '" & CodEmp & "'"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            If Not RQ1.EOF() Then
                AsigFam = IIf(Trim(RQ1.Fields("asigfam")) = "S", 55, 0)
            End If
            Set RQ1 = Nothing
            MtoQuinta = 0
            SQL = "SELECT IFNULL(SUM(CANT),0) AS RTA FROM PL_TAREO P WHERE P.RUB = '605' AND EMP = '" & CodEmp & "' AND LEFT(ANOMES,4) = '" & Year(fec) & "' and tipo = 1"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            If Not RQ1.EOF() Then
                MtoQuinta = RQ1.Fields("RTA")
            End If
            SQL = "SELECT IFNULL(SUM(if(p.unidad = 'V',CANT,round((sbasico/if(p.tipo='1',30,180))*cant,2))),0) as cant FROM PL_TAREO P " & _
                  "LEFT JOIN pl_rubrosremunerativos R ON (P.RUB=R.CODIGO) LEFT JOIN pl_tiporubros T ON (R.TIPO=T.CODIGO) " & _
                  "LEFT JOIN pl_rubros_afectos A ON (P.RUB=A.CODRUB) WHERE T.TIPO = 'I' and (a.codafec = '10' OR p.rub='910' OR p.rub='911' OR p.rub='1003' ) " & _
                  "AND P.EMP = '" & CodEmp & "' AND (P.TIPO = '1' or p.tipo='5') AND LEFT(P.ANOMES,4) = '" & Year(fec) & "'"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            If Not RQ1.EOF() Then
                toting = Round(RQ1.Fields("cant"), 2) ' + ((RQ.Fields("basico") / 30) * Day(fec)) + ((AsigFam / 30) * Day(fec)) + Round(BonosCampo(Year(fec) & Right("00" & Month(fec), 2), "1", CodEmp, "N", Format(Mid(fec, 1, 8) & "01", "yyyy/mm/dd"), Format(fec, "yyyy/mm/dd")), 2), 2)
            End If
            SQL = "update rh_templiquidacion set motivo='" & Motivo & "',asignacion=" & AsigFam & ",bonos=" & MtoBono & ", " & _
                  "fecinicts = DATE_FORMAT(DATE_ADD(fecultcts,INTERVAL 1 DAY),'%Y/%m/%d'),tiempo='" & StrAnios & "', " & _
                  "fecinivac = '" & FecVaca & "',aniosvac=" & Anio & ",mesesvac=" & Dif & ",diasvac= " & I & ", " & _
                  "porafp = " & PorAfp & ",quinta = " & MtoQuinta & ",toting= " & toting & " where usuario = '" & strUsuarioId & "'"
            oConexionMYSQL.Execute SQL
        End If
        SQL = "select trim(nombres),date_format(fecingreso,'%d/%m/%Y') as fecingreso,date_format(feccese,'%d/%m/%Y') as feccese,tiempo, " & _
              "date_format(fecultcts,'%d/%m/%Y') as fecultcts,date_format(fecinicts,'%d/%m/%Y') as fecinicts,motivo,basico,asignacion, " & _
              "bonos,ultgrati,mesgrati,ctacts,bancocts,moncts,date_format(fecult,'%d/%m/%Y') as fecult,fecinivac,aniosvac,mesesvac,diasvac, " & _
              "round(porafp,2) as porafp,trim(afp) as afp,TRIM(dni),(mtoentregas *(-1)) as mto,(prestamos *(-1)) as pr,(adelantos * (-1)) as ad, " & _
              "sctr,quinta,toting,(mtoreten *(-1)) as mtoe,mtogracia,diaspen from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        FlgIng = False
    Else
        FlgIng = True
        SQL = "select CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres,date_format(fecingreso,'%d/%m/%Y') as fecingreso, " & _
              "date_format(feccese,'%d/%m/%Y') as feccese,tiempo,date_format(fecfinctsdep,'%d/%m/%Y') as fecultcts, " & _
              "date_format(fecinictsxpag,'%d/%m/%Y') as fecinicts,motivo,basico,p.asigfam,bono,ultgrati,mesgrati,e.ctsnumcta as ctacts, " & _
              "(select descrip from pl_entidadfinanciera f where f.codigo=e.ctsbanco) as bancocts,e.ctsmon as moncts, " & _
              "DATE_FORMAT(LAST_DAY(CONCAT(YEAR(feccese),'-',left(right(feccese,5),2),'-',day(feccese))),'%Y/%m/%d') as fecult, " & _
              "date_format(fecinivac,'%d/%m/%Y') AS fecinivac,aniosvac,mesvac,diasvac,porafp,afp,e.numdocide,entregas,prestamos,adelantos,if(mtosctr =0,'N','S') as sctr, " & _
              "quintatareo as quinta,toting,dsctoretenjud as mtoe,titulogracia as mtogracia,diaspen from pl_liquidacion p left join empleado e on (p.codemp=e.codigo) where p.cerrado='S' and codemp = '" & CodEmp & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
    End If

    If RQ.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set RQ = Nothing
        Exit Sub
    End If
    FlgLiquidacion = ""
    If RQ.RecordCount > 0 Then
        If FlgIng = False Then
            FlgLiquidacion = CodEmp
            RegistraLiquidacion CodEmp, RQ
        End If
        EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, RQ
    End If
    Set RQ = Nothing
    
    Exit Sub
SERROR:
    Call MsgBox(err.Description)
    Set RQ = Nothing

End Sub


Public Function CalculaRtaQta() As Double
'Dim Ing As Double
'Dim Uit7 As Long
'Dim Uit27 As Long
'Dim Neta As Double
'Dim Mto As Double
'Dim Mto1 As Double
'Dim Mto2 As Double
'Dim strRtaqta As String
'Dim quinta  As Double
'Dim toting  As Double
'Dim totvac  As Double
'Dim totgrati  As Double
'Dim boni  As Double
'
'quinta = 2307.98 'DESDE QUERY
'toting = 32652.85 'DESDE QUERY
'
'totvac = 1886.21 'CALCULADO
'totgrati = 1393.75 'CALCULADO
'boni = 125.44 'CALCULADO
'
'Uit7 = 24850 '7 * 3550 CONSTANTE
'Uit27 = 95850 '27 * 3550 CONSTANTE
'
'If quinta > 0 Then
'  Ing = toting + totvac + totgrati + boni
'  Neta = Ing - Uit7
'  If (Ing < Uit7) Then
'    Mto = 0
'  Else
'    If Neta <= Uit27 Then
'       Mto = Neta * 0.15
'     Else
'       Mto = Uit27 * 0.15
'        If (Neta - Uit27) <= Uit27 Then
'         Mto1 = (Neta - Uit27) * 0.3
'        Else
'         Mto1 = Uit27 * 0.3
'        End If
'     End If
'  End If
'
'  If quinta >= Mto + Mto1 Then
'      Mto2 = quinta - Mto + Mto1
'      strRtaqta = "Devolución exceso retención "
'    Else
'       Mto2 = (Mto + Mto1 - quinta) * -1
'       strRtaqta = "Retención"
'  End If
'Else
'  Mto2 = 0
'End If
'
'CalculaRtaQta = Mto2

End Function



Sub RegistraLiquidacion(CodEmp As String, RQ As MYSQL_RS)
On Error GoTo CtrlError
Dim PagoCts As Double, suma As Double, Mes As Integer, dias As Integer, PagoSctr As Double
Dim PagoVac As Double, FecIni As String, PagoGrati As Double, PagoEs As Double, PagoAfp As Double
Dim Ing As Double, Uit7 As Double, Uit27 As Long, Neta As Double, Mto As Double, Mto1 As Double, Mto2 As Double
Dim SQL As String, PagoVac1 As Double, PagoGrati1 As Double
    SQL = "delete from pl_liquidacion where codemp = '" & CodEmp & "' and cerrado = 'N'"
    oConexionMYSQL.Execute SQL
    suma = CDbl(RQ.Fields("basico")) + CDbl(RQ.Fields("asignacion")) + CDbl(RQ.Fields("bonos"))
    If (Day(CDate(RQ.Fields("feccese"))) = 31) Or (Day(CDate(RQ.Fields("feccese"))) = 30) Or (Day(CDate(RQ.Fields("feccese"))) = 28 And Month(CDate(RQ.Fields("feccese"))) = 2) Or (Day(CDate(RQ.Fields("feccese"))) = 29 And Month(CDate(RQ.Fields("feccese"))) = 2) Then
        If (CDate(RQ.Fields("feccese")) = CDate(RQ.Fields("fecinicts"))) Then
            Mes = DateDiff("m", CDate(RQ.Fields("fecinicts")), CDate(RQ.Fields("feccese")))
        Else
            Mes = DateDiff("m", CDate(RQ.Fields("fecinicts")), CDate(RQ.Fields("feccese"))) + 1
        End If
    Else
        Mes = DateDiff("m", CDate(RQ.Fields("fecinicts")), CDate(RQ.Fields("feccese")))
    End If
    If Day(CDate(RQ.Fields("feccese"))) = Day(CDate(RQ.Fields("fecult"))) Then
        dias = 0
    Else
        dias = DateDiff("d", CDate("01" & Mid(RQ.Fields("feccese"), 3, 4) & Right(RQ.Fields("feccese"), 4)), CDate(RQ.Fields("feccese"))) + 1
    End If
    PagoCts = Round((suma / 12) * Mes, 2) + Round((suma / 360) * dias, 2) + Round((((RQ.Fields("ultgrati") / 12) / 6) * Mes) + (((RQ.Fields("ultgrati") / 360) / 6) * dias), 2)
    PagoVac = (suma * RQ.Fields("aniosvac")) + ((suma / 12) * val(RQ.Fields("mesesvac"))) + ((suma / 360) * RQ.Fields("diasvac"))
    PagoVac1 = PagoVac
    PagoVac = PagoVac - Round(((Round(suma * RQ.Fields("aniosvac"), 2) + Round((suma / 12) * RQ.Fields("mesesvac"), 2) + Round((suma / 360) * RQ.Fields("diasvac"), 2)) * RQ.Fields("porafp")) / 100, 2)
    FecIni = IIf(CDate(RQ.Fields("feccese")) >= CDate("01/07/" & Mid(RQ.Fields("feccese"), 7, 4)), "01/07/" & Mid(RQ.Fields("feccese"), 7, 4), "01/01/" & Mid(RQ.Fields("feccese"), 7, 4))
    If (Day(CDate(RQ.Fields("feccese"))) = 31) Or (Day(CDate(RQ.Fields("feccese"))) = 30) Or (Day(CDate(RQ.Fields("feccese"))) = 28 And Month(CDate(RQ.Fields("feccese"))) = 2) Or (Day(CDate(RQ.Fields("feccese"))) = 29 And Month(CDate(RQ.Fields("feccese"))) = 2) Then
        Mes = DateDiff("m", CDate(FecIni), CDate(RQ.Fields("feccese"))) + 1
    Else
        Mes = DateDiff("m", CDate(FecIni), CDate(RQ.Fields("feccese")))
    End If
    If Day(CDate(RQ.Fields("feccese"))) = Day(CDate(RQ.Fields("fecult"))) Then
        dias = 0
    Else
        dias = DateDiff("d", CDate("01/" & Mid(RQ.Fields("feccese"), 4, 7)), CDate(RQ.Fields("feccese"))) + 1
    End If
    PagoGrati = Round((suma / 6) * Mes, 2) + Round((suma / 180) * dias, 2)
    PagoGrati1 = PagoGrati
    PagoGrati = PagoGrati - Round((((Round((suma / 6) * Mes, 2) + Round((suma / 180) * dias, 2)) * RQ.Fields("porafp")) / 100), 2)
    Uit7 = 7 * 3700
    Uit27 = 95850
    Mto2 = 0
    If val(RQ.Fields("quinta")) > 0 Then
        Ing = RQ.Fields("toting") + (Round(suma * RQ.Fields("aniosvac"), 2) + Round((suma / 12) * RQ.Fields("mesesvac"), 2) + Round((suma / 360) * RQ.Fields("mesesvac"), 2)) + (Round((suma / 6) * Mes, 2) + Round((suma / 180) * dias, 2))
        Neta = Ing - Uit7
        If (Ing < Uit7) Then
            Mto = 0
        Else
            If Neta <= Uit27 Then
                Mto = Neta * 0.15
            Else
                Mto = Uit27 * 0.15
                If (Neta - Uit27) <= Uit27 Then
                    Mto1 = (Neta - Uit27) * 0.3
                Else
                    Mto1 = Uit27 * 0.3
                End If
            End If
        End If
        Mto2 = (Mto + Mto1) - RQ.Fields("quinta") * (-1)
    Else
        Mto2 = 0
    End If
    PagoEs = Round((((Round(suma * RQ.Fields("aniosvac"), 2) + Round((suma / 12) * RQ.Fields("mesesvac"), 2) + Round((suma / 360) * RQ.Fields("diasvac"), 2)) + (Round((suma / 6) * Mes, 2) + Round((suma / 180) * dias, 2))) * 9) / 100, 2)
    PagoAfp = 0
    If RQ.Fields("afp") <> "ONP" Then
        PagoAfp = Round((((Round((suma / 6) * Mes, 2) + Round((suma / 180) * dias, 2)) + (Round(suma * RQ.Fields("aniosvac"), 2) + Round((suma / 12) * RQ.Fields("mesesvac"), 2) + Round((suma / 360) * RQ.Fields("diasvac"), 2))) * 2) / 100, 2)
    End If
    PagoSctr = 0
    If Trim(RQ.Fields("sctr")) = "S" Then
        PagoSctr = Round((((Round((suma / 6) * Mes, 2) + Round((suma / 180) * dias, 2)) + (Round(suma * RQ.Fields("aniosvac"), 2) + Round((suma / 12) * RQ.Fields("mesesvac"), 2) + Round((suma / 360) * RQ.Fields("diasvac"), 2))) * 0.53) / 100, 2)
    End If
    SQL = "insert into pl_liquidacion(Codemp,FecIngreso,FecCese,Tiempo,FecIniCtsDep,FecFinCtsDep,FecIniCtsxPag, " & _
          "Motivo,Afp,PorAfp,Basico,AsigFam,Bono,UltGrati,MtoCts,FecIniVac,aniosvac,mesvac,diasvac,MtoVac,FecIniGrat, " & _
          "MesGrati,MtoGrat,Prestamos,Adelantos,Renta,quintatareo,Entregas,DsctoRetenJud,TituloGracia,MtoEssalud,MtoSctr,MtoAfp, " & _
          "toting,fecemision,vacsinafp,gratisinafp) values ('" & CodEmp & "','" & Format(CDate(RQ.Fields("fecingreso")), "yyyy/mm/dd") & "', " & _
          "'" & Format(CDate(RQ.Fields("feccese")), "yyyy/mm/dd") & "','" & Trim(RQ.Fields("tiempo")) & "', " & _
          "'" & Format(CDate(RQ.Fields("fecingreso")), "yyyy/mm/dd") & "','" & Format(CDate(RQ.Fields("fecultcts")), "yyyy/mm/dd") & "', " & _
          "'" & Format(CDate(RQ.Fields("fecinicts")), "yyyy/mm/dd") & "','" & Trim(RQ.Fields("motivo")) & "', " & _
          "'" & Trim(RQ.Fields("afp")) & "'," & RQ.Fields("porafp") & "," & RQ.Fields("basico") & "," & RQ.Fields("asignacion") & ", " & _
          "" & RQ.Fields("bonos") & "," & RQ.Fields("ultgrati") & "," & PagoCts & ",'" & Format(CDate(RQ.Fields("fecinivac")), "yyyy/mm/dd") & "', " & _
          "" & RQ.Fields("aniosvac") & "," & RQ.Fields("mesesvac") & "," & RQ.Fields("diasvac") & "," & PagoVac & ", " & _
          "'" & Format(CDate(FecIni), "yyyy/mm/dd") & "','" & RQ.Fields("mesgrati") & "'," & PagoGrati & "," & RQ.Fields("PR") & ", " & _
          "" & RQ.Fields("ad") & "," & Mto2 & "," & RQ.Fields("quinta") & "," & RQ.Fields("mto") & "," & RQ.Fields("mtoe") & "," & RQ.Fields("mtogracia") & ", " & _
          "" & PagoEs & "," & PagoSctr & "," & PagoAfp & "," & RQ.Fields("toting") & ",'" & Date & "'," & PagoVac1 & "," & PagoGrati1 & ")"
    oConexionMYSQL.Execute SQL
Exit Sub
CtrlError:
    MsgBox err.Description, vbCritical, "Error registrando Liquidación"
End Sub
Public Sub sp_Rep_PrestamosRegistrados(CodEmp As String, Mes As String)
    Dim SQL As String, SVal As String, SDiv As String
    Dim myrs As MYSQL_RS
    If CodEmp <> "" Then
        SDiv = " and e.codigo = '" & CodEmp & "'"
    End If
    Screen.MousePointer = vbHourglass
    SQL = "select e.codigo,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres,date_format(fec_reg,'%d/%m/%Y'),coddoc,documento, " & _
          "ifnull((SELECT importe from documentos_rrhh r where r.solicita=d.solicita and r.documento=d.documento " & _
          "and mon = 'E'),0) as importedol,ifnull((SELECT montopagado from documentos_rrhh r where r.solicita=d.solicita " & _
          "and r.documento=d.documento and mon = 'E'),0) as pagadodol,ifnull((SELECT importe from documentos_rrhh r " & _
          "where r.solicita=d.solicita and r.documento=d.documento and mon = 'N'),0) as importesol,ifnull((SELECT montopagado from " & _
          "documentos_rrhh r where r.solicita=d.solicita and r.documento=d.documento and mon = 'N'),0) as pagadosol,mon, asunto from documentos_rrhh d left join empleado e on (d.solicita=e.codigo) " & _
          "where anomes = '" & strAnoSistema & Mes & "'" & SDiv & " order by nombres,fec_reg"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_PrestamosaDescontar(Mes As String, CodEmp As String, Tipo As String, plani As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim STipo As String, SMes As String, SEmp As String, SMes1 As String, STipo1 As String
    Dim SPlani As String, Splani1 As String
    If Mes <> "" Then
        SMes = " and e.anomesplani = '" & strAnoSistema & Mes & "'"
        SMes1 = " and r.anomes = '" & strAnoSistema & Mes & "'"
        If plani <> "" Then
            SMes1 = " and ((r.anomes = '" & strAnoSistema & Mes & "' and r.tipoplani = '" & plani & "') or (r.anomesdesc = '" & strAnoSistema & Mes & "' and r.tipoplanidesc = '" & plani & "'))"
        End If
    Else
        SMes = " and left(e.anomesplani,4) = '" & strAnoSistema & "'"
        SMes1 = " and left(r.anomes,4) = '" & strAnoSistema & "'"
    End If
    If CodEmp <> "" Then
        SEmp = " and r.solicita = '" & CodEmp & "'"
    End If
    If Tipo <> "" Then
        STipo = " and r.tipo = '" & Tipo & "'"
        STipo1 = " and e.tipo = '" & Tipo & "'"
    End If
    If plani <> "" Then
        SPlani = " and e.tipoplani = '" & plani & "'"
    End If
    SQL = "SELECT r.SOLICITA,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres, " & _
          "if(d.mon = 'N',SUM(r.MONTO),0) AS MTOS,if(d.mon = 'E',SUM(r.MONTO),0) AS MTOE," & _
          "IF(d.mon = 'N',ifnull(e.abonado,0),0) as descontados,IF(d.mon = 'E',ifnull(e.abonado,0),0) as descontadoe," & _
          "d.mon,r.tipo FROM (rh_divprestadel R Left Join " & _
          "(select sum(montoabonado) as abonado,solicita,documento,tipo,anomesplani from rh_pagosemp E where 1=1 " & SMes & _
          STipo1 & SPlani & " and liquid = 'N' group by solicita,e.tipo) as e " & _
          "on (r.solicita=e.solicita) and (r.documento=e.documento) and (r.tipo = e.tipo)) left join empleado a " & _
          "on (r.solicita=a.codigo) LEFT JOIN `documentos_rrhh` d on (r.solicita=d.solicita) and " & _
          "(r.documento=d.documento) and (r.tipo = IF(d.coddoc = 'SS','A','P')) WHERE 1=1 " & SMes1 & STipo & SEmp & _
          " AND r.LIQUI = 'N' GROUP BY r.SOLICITA,d.mon,r.tipo ORDER BY nombres,r.tipo"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_FoliosAutomaticos()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "SELECT d.identificador,a.fecha_registro,c.descrip,concat_ws('-',d.serie,d.correl) as numdoc,a.usuario, " & _
          "m.cod_estado,d.obs,d.orden FROM amarre_documento a left join documento_contables d on (a.identificador=d.identificador) " & _
          "left join movi_documento m on (d.identificador = m.identificador) left join cnauxil c on (d.auxiliar=c.auxiliar) " & _
          "and (d.codigo=c.codigo) where a.obs IN ('F.A.','F.A.MOD') and m.COD_ESTADO not in ('EL','CA') order by d.identificador"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de Folios Automáticos", vbInformation + vbOKOnly, "BRANDTADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_PrestamosaDescontarUti(CodEmp As String, Tipo As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim STipo As String, SMes As String, SEmp As String, SMes1 As String, STipo1 As String
    Dim SPlani As String, Splani1 As String
    If CodEmp <> "" Then
        SEmp = " and r.solicita = '" & CodEmp & "'"
    End If
    If Tipo <> "" Then
        STipo = " and r.tipo = '" & Tipo & "'"
        STipo1 = " and e.tipo = '" & Tipo & "'"
    End If
    SQL = "SELECT r.SOLICITA,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres, " & _
          "if(d.mon = 'N',SUM(r.MONTO),0) AS MTOS,if(d.mon = 'E',SUM(r.MONTO),0) AS MTOE," & _
          "IF(d.mon = 'N',ifnull(e.abonado,0),0) as descontados,IF(d.mon = 'E',ifnull(e.abonado,0),0) as descontadoe," & _
          "d.mon,r.tipo FROM (rh_divprestadel R Left Join " & _
          "(select sum(montoabonado) as abonado,e.solicita,tipo,anomesplani from documentos_rrhh dr left join rh_pagosemp E on (dr.solicita=e.solicita) " & _
          "and (IF(dr.coddoc = 'SS','A','P')=e.tipo) AND (DR.ANOMES=E.ANOMES) and (dr.documento=e.documento) where 1=1 " & _
          STipo1 & " and liquid = 'N' and dr.uti='S' and left(DR.anomes,4) = '" & strAnoSistema & "' group by solicita,e.tipo) as e " & _
          "on (r.solicita=e.solicita) and (r.tipo = e.tipo)) left join empleado a " & _
          "on (r.solicita=a.codigo) LEFT JOIN `documentos_rrhh` d on (r.solicita=d.solicita) and " & _
          "(r.documento=d.documento) and (r.tipo = IF(d.coddoc = 'SS','A','P')) WHERE 1=1 and d.uti = 'S' " & STipo & SEmp & _
          " AND r.LIQUI = 'N' GROUP BY r.SOLICITA,d.mon,r.tipo ORDER BY nombres,r.tipo"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_FoliosAutomaticosProg()
    Dim SQL As String
    Dim myrs As MYSQL_RS
    SQL = "select p.identificador,date_format(p.fec_ini,'%d/%m/%Y') as fecini,date_format(p.fec_fin,'%d/%m/%Y') as fecfin, " & _
          "p.dia,p.estado,p.anomes,c.descrip,d.obs,d.orden " & _
          "from prog_folios p left join documento_contables d on (p.identificador=d.identificador) " & _
          "left join cnauxil c on (d.auxiliar=c.auxiliar) and (d.codigo=c.codigo) order by p.identificador"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existe Datos de Folios Automáticos", vbInformation + vbOKOnly, "BRANDTADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\ControlDocumentario\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_SubsidiosTrab(CodEmp As String, CodMov As String)
    Dim SQL As String, str As String
    Dim myrs As MYSQL_RS
        If CodEmp <> "" Then
        str = " and c.codemp='" & CodEmp & "'"
    Else
        str = ""
    End If
    SQL = "SELECT e.Codigo,Concat(ApePat,' ', ApeMat,' ', Nombre1,' ', NOMBRE2) AS Nombres, " & _
          " e.Fec_Ingreso,date_format(c.fec_salida,'%d/%m/%Y'),date_format(c.fec_regreso,'%d/%m/%Y'),c.observacion,c.sinbono, " & _
          "(select descrip from pl_tipsuspensionaboral where codigo = c.dpto) as dpto From empleado e INNER JOIN calendario c ON (e.codigo=c.codemp) where e.tipo not in (3,4) and c.movemp = '" & CodMov & "' and " & _
          "left(c.codigo,4) = '" & strAnoSistema & "' " & str & " order by c.codigo"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_PermisosTrab(CodEmp As String, CodMov As String)
    Dim SQL As String, str As String
    Dim myrs As MYSQL_RS
        If CodEmp <> "" Then
        str = " and c.codemp='" & CodEmp & "'"
    Else
        str = ""
    End If
    SQL = "SELECT e.Codigo,Concat(ApePat,' ', ApeMat,' ', Nombre1,' ', NOMBRE2) AS Nombres, " & _
          " e.Fec_Ingreso,date_format(c.fec_salida,'%d/%m/%Y'),date_format(c.fec_regreso,'%d/%m/%Y'),c.observacion " & _
          " From empleado e INNER JOIN calendario c ON (e.codigo=c.codemp) where c.movemp = '" & CodMov & "' and " & _
          "left(c.codigo,4) = '" & strAnoSistema & "' " & str & " order by c.codigo"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existen Registros", vbInformation, gsNomSW
        Set myrs = Nothing
        Exit Sub
    End If
    If myrs.RecordCount > 0 Then
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_HistorialSueldos(CodEmp As String, Division As String)
    Dim SQL As String
    Dim myrs As MYSQL_RS
    Dim SEmp As String, Str1 As String, Ces As Boolean, Sl As Boolean, str2 As String
    If CodEmp <> "" Then SEmp = " and e.codigo = '" & CodEmp & "'"
    If Division <> "0000" And Division <> "0006" Then Str1 = " and c.division = '" & Division & "'"
    If CodEmp = "" Then
        If MsgBox("¿Desea incluir a los Cesados?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then Ces = True
    End If
    If Ces = False Then str2 = " and e.situacion = 1 "
    If MsgBox("¿Desea ver SOLO las variaciones en Sueldo?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then Sl = True
    Screen.MousePointer = vbHourglass
    If Sl = True Then
        Dim RQ As MYSQL_RS
        Dim RQ1 As MYSQL_RS
        SQL = "delete from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
        oConexionMYSQL.Execute SQL
        SQL = "insert into rh_templiquidacion(codigo,nombres,moncts,basico,motivo,sctr,usuario) " & _
              "select distinct e.codigo,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres, " & _
              "mon,sueldo,(select descriplocal from cnmdepar d where d.coddep=c.division) as divi, " & _
              "e.situacion,'" & strUsuarioId & "' from (empleado e left join rh_sueldos s on (e.codigo=s.codemp)) inner join contrato c " & _
              "on (c.codemp=e.codigo) and c.codigo = (select max(codigo) from contrato t where c.codemp=t.codemp) " & _
              "where e.tipo not in (3,4) " & SEmp & Str1 & str2 & " order by nombres,sueldo"
        oConexionMYSQL.Execute SQL
        SQL = "select codigo,basico from rh_templiquidacion where usuario = '" & strUsuarioId & "' order by codigo,basico"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        Do While Not RQ.EOF()
            SQL = "select codcont,fecha from rh_sueldos where codemp = '" & RQ.Fields("codigo") & "' and sueldo = " & RQ.Fields("basico") & " " & _
                  "order by codemp,codcont limit 1"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            If Not RQ1.EOF() Then
                SQL = "update rh_templiquidacion set fecingreso = '" & RQ1.Fields("fecha") & "',feccese = '" & RQ1.Fields("codcont") & "' " & _
                      "where codigo = '" & RQ.Fields("codigo") & "' and basico= " & RQ.Fields("basico") & " and usuario = '" & strUsuarioId & "'"
                oConexionMYSQL.Execute SQL
            End If
            RQ.MoveNext
        Loop
        Set RQ = Nothing
        Set RQ1 = Nothing
        SQL = "select codigo,nombres,feccese as codcont,0 as item,moncts as mon,basico as sueldo,date_format(fecingreso,'%d/%m/%Y') as fecha,motivo as divi, " & _
              "sctr as situacion from rh_templiquidacion where usuario = '" & strUsuarioId & "' order by nombres,feccese,fecingreso"
    Else
        SQL = "select e.codigo,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres, " & _
              "codcont,item,mon,sueldo,date_FORMAT(fecha,'%d/%m/%Y') as fecha,(select descriplocal from cnmdepar d where d.coddep=c.division) as divi, " & _
              "e.situacion from (empleado e left join rh_sueldos s on (e.codigo=s.codemp)) inner join contrato c " & _
              "on (c.codemp=e.codigo) and c.codigo = (select max(codigo) from contrato t where c.codemp=t.codemp) " & _
              "where e.tipo not in (3,4) " & SEmp & Str1 & str2 & " order by nombres,codcont,fecha"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Ventas(Optional Opc As Integer)
    Dim SQL As String
    Dim RQ As MYSQL_RS, RQ1 As MYSQL_RS
    Dim myrs As MYSQL_RS
    Screen.MousePointer = vbHourglass
    SQL = "delete from rh_templiquidacion where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQL
    SQL = "insert into rh_templiquidacion(fecingreso,tiempo,nombres,mesgrati,codigo,motivo,basico,ctacts,usuario,asignacion,moncts) " & _
          "Select DISTINCT dc.identificador,CONCAT(dc.Serie,'-',dc.Correl) as documento,divi.descrip,a.cod_tipo_doc, " & _
          "dc.codigo,aux.descrip as NomAux,((dc.SubTotal * cnd.factor )+dc.otros_montos) as subtotal, " & _
          "(Select descrip from almacen where codigo=dc.codalmac) as almacen,'" & strUsuarioId & "',dadi.mto,dadi.tipo" & _
          " from (((((((amarre_documento as a LEFT JOIN cndocum as cnd on (a.cod_tipo_doc=cnd.coddoc) )" & _
          " LEFT Join documento_contables as dc" & _
          " on (a.Identificador=dc.Identificador))right JOIN detallefact as dt" & _
          " on (dc.Identificador=dt.Identificador)) LEFT Join cncosto as ct" & _
          " on (dc.cenco=ct.cenco)) LEFT Join encargados as enc" & _
          " on (dc.encargado=enc.codigo))LEFT Join cnauxil as aux" & _
          " on (dc.Auxiliar=aux.auxiliar and dc.codigo=aux.codigo))LEFT join cntablas as cn" & _
          " on (dc.auxiliar=cn.tip_linea) LEFT join movi_documento as mv" & _
          " on (dc.Identificador=mv.Identificador))left join cnmdepar as divi on (dc.division=divi.coddep) " & _
          "LEFT JOIN (SELECT sum(MONTO) as mto,d.tipo as tipo,dc.codigo as codaux,dc.codalmac as almacen, " & _
          "divi.descrip as division FROM (((deducciones d left join documento_contables dc on (d.identificador=dc.identificador)) " & _
          "left join cnauxil c on (dc.Auxiliar=c.auxiliar and dc.codigo=c.codigo)) left join cnmdepar as divi on (dc.division=divi.coddep)) " & _
          "WHERE (CASAMATRIZ = 'S' AND ANOMES = '" & strAnoSistema & Right("00" & (val(strMesSistema) - 1), 2) & "') " & _
          "OR (ANOMES = '" & strAnoSistema & strMesSistema & "') group by d.tipo,c.descrip,almacen,divi.descrip " & _
          "order by divi.descrip,almacen,c.descrip) AS dadi on (dadi.division=divi.descrip) and (dadi.codaux=dc.codigo) " & _
          "and (dadi.almacen=dc.codalmac) where dc.auxiliar='2' and (a.cod_tipo_doc='01' or a.cod_tipo_doc='03'" & _
          " or a.cod_tipo_doc='07' or a.cod_tipo_doc='08' or ((dt.codser <> '065' and a.cod_tipo_doc='P') and " & _
          "(dt.codser <> '068' and a.cod_tipo_doc='P'))) and mv.cod_estado not in ('" & ELIMINADO & "','" & ANULADO & "') " & _
          "and a.flag='0' and cn.codtab='1' and dt.codser not in ('065','634','068','000') " & _
          "and dc.Fec_Emision>='" & Format("01/" & strMesSistema & "/" & strAnoSistema, "yyyy/mm/dd") & "' and dc.Fec_Emision<='" & Format(IIf(Day(DateSerial(strAnoSistema, strMesSistema + 1, 0)) = 31, "30/" & strMesSistema & "/" & strAnoSistema, DateSerial(strAnoSistema, strMesSistema + 1, 0)), "yyyy/mm/dd") & "' ORDER BY divi.descrip,documento"
    oConexionMYSQL.Execute SQL
    SQL = "select fecingreso as ident,basico as subtotal,tiempo as doc,moncts from rh_templiquidacion where usuario = '" & strUsuarioId & "' " & _
          "order by nombres,tiempo"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    Do While Not RQ.EOF()
        Dim SumProf As Double
        Dim Cant As Integer
        Dim strident As String
        SumProf = 0
        SQL = "select left(right(prof,6),2) as mesprof,valor from factura_proforma where identificador = '" & RQ.Fields("ident") & "'"
        Set RQ1 = oConexion.EjecutaSelectRS(SQL)
        If strident = Trim(RQ.Fields("ident")) Then
            Cant = 2
        Else
            strident = Trim(RQ.Fields("ident"))
            Cant = 1
        End If
        Do While Not RQ1.EOF
            SumProf = SumProf + CDbl(IIf(val(RQ1.Fields("valor")) = 0, 0, RQ1.Fields("valor")))
            RQ1.MoveNext
        Loop
        If SumProf > 0 Then
            RQ1.MovePrevious
            If val(RQ1.Fields("mesprof")) <> val(strMesSistema) Then
                If FormatNumber(SumProf, 2) = FormatNumber(CDbl(IIf(val(RQ.Fields("subtotal")) = 0, 0, RQ.Fields("subtotal"))), 2) Then
                    SQL = "delete from rh_templiquidacion where usuario = '" & strUsuarioId & "' and tiempo = '" & RQ.Fields("doc") & "' and fecingreso='" & RQ.Fields("ident") & "'"
                    oConexionMYSQL.Execute SQL
                Else
                    SQL = "update rh_templiquidacion set basico = " & RQ.Fields("subtotal") - SumProf & " where usuario = '" & strUsuarioId & "' and tiempo = '" & RQ.Fields("doc") & "' and fecingreso='" & RQ.Fields("ident") & "'"
                    oConexionMYSQL.Execute SQL
                End If
            End If
        End If
        If Cant > 1 Then
            SQL = "update rh_templiquidacion set basico = 0 where usuario = '" & strUsuarioId & "' and tiempo = '" & RQ.Fields("doc") & "' and fecingreso='" & RQ.Fields("ident") & "' and moncts= '" & RQ.Fields("moncts") & "'"
            oConexionMYSQL.Execute SQL
        End If
        RQ.MoveNext
    Loop
    Set RQ = Nothing
    Set RQ1 = Nothing
    Screen.MousePointer = vbDefault
    If Opc = 1 Then Exit Sub
    SQL = "select nombres as division,ctacts as lote,motivo as desaux,IF(mesgrati = '03',sum(basico)/1.19,sum(basico)) as monto, " & _
          "ifnull((select distinct asignacion as adicion from rh_templiquidacion r where usuario = '" & strUsuarioId & "' and moncts= 'A' " & _
          "and r.nombres=r1.nombres and r.ctacts=r1.ctacts and r.motivo=r1.motivo order by nombres),0) as adicion, " & _
          "ifnull((select distinct asignacion * (-1) as deduccion from rh_templiquidacion r where usuario = '" & strUsuarioId & "' and moncts= 'D' " & _
          "and r.nombres=r1.nombres and r.ctacts=r1.ctacts and r.motivo=r1.motivo order by nombres),0) as deduccion,tiempo,mesgrati " & _
          "from rh_templiquidacion r1 where usuario = '" & strUsuarioId & "' group by nombres,ctacts,motivo,tiempo,mesgrati " & _
          "union select distinct k.division as division,k.almacen as lote,k.nomaux as desaux,0 as monto, " & _
          "if(k.tip = 'A',k.mto,0) as adicion,if(k.tip = 'A',0,k.mto)*(-1) as deduccion,k.documento as tiempo,k.tipodoc as mesgrati " & _
          "from rh_templiquidacion r right join (SELECT divi.descrip as division,(Select descrip from almacen where codigo=dc.codalmac) as almacen, " & _
          "c.descrip as NomAux,sum(MONTO) as mto,d.tipo as tip,CONCAT(dc.Serie,'-',dc.Correl) as documento,ad.cod_tipo_doc as tipodoc " & _
          "FROM ((((deducciones d left join documento_contables dc on (d.identificador=dc.identificador)) left join amarre_documento ad " & _
          "on (ad.identificador=dc.identificador)) left join cnauxil c on (dc.Auxiliar=c.auxiliar and dc.codigo=c.codigo)) " & _
          "left join cnmdepar as divi on (dc.division=divi.coddep)) WHERE (CASAMATRIZ = 'S' AND " & _
          "d.ANOMES = '" & strAnoSistema & Right("00" & (val(strMesSistema) - 1), 2) & "') OR (d.ANOMES = '" & strAnoSistema & strMesSistema & "') " & _
          "group by d.tipo,c.descrip,dc.codalmac,divi.descrip order by divi.descrip,dc.codalmac,c.descrip) as k on (r.asignacion=k.mto) " & _
          "where asignacion is null order by DIVISION,lote,desaux,mesgrati,tiempo"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount = 0 Then
        MsgBox "No Existe Documentos para la consuta", vbInformation, "Reporte de Ventas"
        Set myrs = Nothing
        Exit Sub
    End If
    Exportar_Excel_VentasExterior Rep_Documents & "\Ventas_al_Exterior_" & Trim(NombreMes(strMesSistema, False)) & strAnoSistema, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_Letras(CodCli As String, NumLetra As String)
    Dim SQL As String, rpt As String
    Dim str As String, Str1 As String
    Dim myrs As MYSQL_RS
    If CodCli <> "" Then str = " AND l.codaux = '" & CodCli & "'"
    If NumLetra <> "" And NumLetra <> "000000" Then Str1 = " and (numero = '" & NumLetra & "' or ref = '" & NumLetra & "')"
    SQL = "select numero,c.descrip as cliente,fecgiro,fecvcto,l.importe,l.abono,l.tipo,l.interes,l.interesbco, " & _
          "l.ndebito,if(l.tipo='R',CONCAT(ref,'P'),CONCAT(numero,l.tipo)) as tipoletra, " & _
          "if(l.tipo='R',concat(numero,l.tipo),'') as tipoletraren from letra l left join cnauxil c on (l.codaux = c.codigo) " & _
          "where c.auxiliar = 2 and codestado <> 'EL' " & str & Str1 & " order by tipoletra,numero"
    oReporte.Titulo = "REPORTE DE LETRAS - DETALLE"
    oReporte.Reporte = "Rep_LetrasDetalle.rpt"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "BRANDTADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Facturacion\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_ResumenVacaciones(CodDiv As String, CodEmp As String, fecha As String)
    Dim SQL As String, rpt As String, I As Long
    Dim StrAnio1 As String, StrAnio2 As String, k As Integer, StrCampos As String
    Dim str As String, Str1 As String, str2 As String
    Dim myrs As MYSQL_RS
    str = ""
    Str1 = ""
    str2 = ""
    If CodEmp <> "" Then
        str = " and e.codigo = '" & CodEmp & "' "
    End If
    If CodDiv <> "000000000000" And CodDiv <> "000000000006" Then
        Str1 = " and o.division = '" & CodDiv & "' "
    End If
    If CodEmp = "" Then
        If MsgBox("¿Desea incluir a los cesados?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then
            str2 = ""
        Else
            str2 = " and e.situacion = 1 "
        End If
    End If
    If MsgBox("¿Desea ver el Reporte agrupados por Períodos?", vbYesNo + vbQuestion, "NOVADMIN") = vbYes Then
        StrCampos = ""
        StrAnio1 = ""
        StrAnio2 = ""
        k = 1
        Screen.MousePointer = vbHourglass
        For I = strAnoSistema - 2 To strAnoSistema - 1
            StrCampos = StrCampos & ",IFNULL(convert(P" & k & ".DIAS,UNSIGNED),IF(year(e.fec_ingreso)<='" & I & "',0,'N/A')) AS 'PV" & Right(I, 2) & Right(I + 1, 2) & "'," & _
                        "IFNULL(CONVERT(P" & k + 1 & ".DIAS,UNSIGNED),IF(year(e.fec_ingreso)<='" & I & "',0,'N/A')) AS 'PD" & Right(I, 2) & Right(I + 1, 2) & "', " & _
                        "IF(IFNULL(CONVERT(P" & k & ".DIAS,UNSIGNED),IF(year(e.fec_ingreso)<='" & I & "',0,'N/A'))='N/A' and IFNULL(CONVERT(P" & k + 1 & ".DIAS,UNSIGNED),IF(year(e.fec_ingreso)<='" & I & "',0,'N/A')) ='N/A','N/A',(30 - (IFNULL(P" & k & ".DIAS,'N/A')+IFNULL(P" & k + 1 & ".DIAS,'N/A')))) as 'PP" & Right(I, 2) & Right(I + 1, 2) & "'"
            StrAnio1 = StrAnio1 & " LEFT JOIN (SELECT SUM(monto_viatico) AS DIAS,CODEMP,a.PERIODO FROM calendario A " & _
                      "where movemp = '02' and gocehaber = 'S' and (a.periodo='" & I & "-" & I + 1 & "') " & _
                      "GROUP BY CODEMP,a.PERIODO) AS P" & k & " ON (P" & k & ".CODEMP=e.codigo) " & _
                      "LEFT JOIN (SELECT SUM(monto_viatico) AS DIAS,CODEMP,a.PERIODO FROM calendario A " & _
                      "where movemp = '02' and gocehaber IN ('N','X') and (a.periodo='" & I & "-" & I + 1 & "') " & _
                      "GROUP BY CODEMP,a.PERIODO) AS P" & k + 1 & " ON (P" & k + 1 & ".CODEMP=e.codigo)"
            k = k + 2
        Next
        SQL = "SELECT distinct CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as 'Apellidos y Nombres', " & _
              "(select descriplocal from cnmdepar r where r.coddep=o.division) as 'Division', " & _
              "date_format(e.fec_ingreso,'%d/%m/%Y') as 'Fec. Ingreso' " & StrCampos & _
              " FROM (empleado e left join contrato o on (e.codigo=o.codemp)) " & StrAnio1 & _
              "where o.codigo = (select max(codigo) from contrato n where n.codemp=E.codigo group by n.codemp) " & _
              "and e.situacion = 1 and e.tipo not in (3,4) " & str2 & Str1 & "ORDER BY CONCAT_WS(' ',apepat,apemat,nombre1,nombre2)"
        m_Empresa = strNombreEmpresa
        m_Titulo = "CONTROL DE VACACIONES POR DISFRUTAR AL " & fecha
        m_Reporte = "Rep_VacaionesFaltantes.rpt"
    Else
        Screen.MousePointer = vbHourglass
        SQL = "SELECT DISTINCT c.codemp,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres,C.periodo, " & _
              "IFNULL(K.DIAS,0) AS DIASCOMPRA,IFNULL(K1.DIAS,0) AS DIASFISICO,(30 - (IFNULL(K.DIAS,0)+IFNULL(K1.DIAS,0))) as pendiente, " & _
              "(select descriplocal from cnmdepar r where r.coddep=o.division) as division,date_format(e.fec_ingreso,'%d/%m/%Y') as fecini FROM ((calendario C " & _
              "left join empleado e on (c.codemp=e.codigo)) left join contrato o on (c.codemp=o.codemp)) " & _
              "LEFT JOIN (SELECT SUM(monto_viatico) AS DIAS,CODEMP,PERIODO FROM calendario A where movemp = '02' " & _
              "and gocehaber = 'S' GROUP BY CODEMP,PERIODO) AS K ON (K.CODEMP=C.CODEMP) AND (K.PERIODO=C.PERIODO) " & _
              "LEFT JOIN (SELECT SUM(monto_viatico) AS DIAS,CODEMP,PERIODO FROM calendario A where movemp = '02' " & _
              "and gocehaber IN ('N','X') GROUP BY CODEMP,PERIODO) AS K1 ON (K1.CODEMP=C.CODEMP) AND (K1.PERIODO=C.PERIODO) " & _
              "where movemp = '02' and o.codigo = (select max(codigo) from contrato n where n.codemp=c.codemp group by n.codemp) " & _
              str & Str1 & str2 & " order by nombres,PERIODO"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "BRANDTADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Screen.MousePointer = vbDefault
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_PlanillaPagoAportes(AnoMes As String, CodAfp As String)
    Dim SQL As String, Interes As Double, Interes1 As Double
    Dim str As String, Str1 As String, RES As Variant
    Dim myrs As MYSQL_RS
    str = ""
    If AnoMes = "" Then
        AnoMes = strAnoSistema & strMesSistema
    End If
    If CodAfp <> "" Then
        str = " and e.codafp = '" & CodAfp & "' "
    End If
    RES = InputBox("Ingrese Monto de Interés Moratorio para Total Fondo de Pensiones", "PLANILLA APORTES", 0)
    If RES = "" Then Exit Sub
    Interes = CDbl(RES)
    Interes = IIf(IsNull(Interes) Or val(Interes) = 0, 0, Interes)
    RES = InputBox("Ingrese Monto de Interés Moratorio para Total Retenciones", "PLANILLA APORTES", 0)
    If RES = "" Then Exit Sub
    Interes1 = CDbl(RES)
    Interes1 = IIf(IsNull(Interes1) Or val(Interes) = 0, 0, Interes1)
    If MsgBox("¿Emitir Reporte para Trabajadores de Ley?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then
        Str1 = " and e.sctr = 'S'"
        Select Case CodAfp
            Case "01"
                m_Reporte = "Rep_PlanillaPagoAportesLey_Hor.rpt"
            Case "02"
                m_Reporte = "Rep_PlanillaPagoAportesLey_Prof.rpt"
            Case "03"
                m_Reporte = "Rep_PlanillaPagoAportesLey_Int.rpt"
            Case "05"
                m_Reporte = "Rep_PlanillaPagoAportesLey_Pri.rpt"
            Case Else
                Exit Sub
        End Select
    Else
        Str1 = " and e.sctr = 'N'"
        Select Case CodAfp
            Case "01"
                m_Reporte = "Rep_PlanillaPagoAportes_Hor.rpt"
            Case "02"
                m_Reporte = "Rep_PlanillaPagoAportes_Prof.rpt"
            Case "03"
                m_Reporte = "Rep_PlanillaPagoAportes_Int.rpt"
            Case "05"
                m_Reporte = "Rep_PlanillaPagoAportes_Pri.rpt"
            Case Else
                Exit Sub
        End Select
    End If
    m_Anomes = Right(AnoMes, 2) & "/" & Left(AnoMes, 4)
    SQL = "select e.numseg,apepat,apemat,CONCAT_WS(' ',nombre1,nombre2) as nombres,'08' as tipo,IF(ifnull(e.fec_cese,'')='','',DATE_FORMAT(e.fec_cese,'%d/%m/%Y')) as feccese," & _
          "ifnull((vacsinafp),0) as rem,if(e.sctr = 'S',ifnull((vacsinafp)*0.02,0),0) as dospor, ifnull((vacsinafp)*if(e.sctr = 'S',0.12,0.10),0) as aporte,ifnull(round(((select porcseg from pl_afppar f where f.anomes = '" & AnoMes & "' and " & _
          "codafp = (select codigo from afp where nombre = p.afp)) *(vacsinafp))/100,2),0) as seg,ifnull(round(((select porccom from pl_afppar f where f.anomes = '" & AnoMes & "' and " & _
          "codafp = (select codigo from afp where nombre = p.afp)) *(vacsinafp))/100,2),0) as com,(select sum(round(vacsinafp,2)) from (empleado e inner join contrato c on (e.codigo=c.codemp)) " & _
          "left join pl_liquidacion p on (p.codemp=e.codigo) where p.cerrado='S' and c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) and (situacion = 0 and fec_cese BETWEEN '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01' and " & _
          "'" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & Day(DateSerial(Left(AnoMes, 4), val(Right(AnoMes, 2)) + 1, 0)) & "') " & str & Str1 & ") as totrem, " & _
          "(select sum((vacsinafp)*if(e.sctr = 'S',0.12,0.10)) from (empleado e inner join contrato c on (e.codigo=c.codemp)) " & _
          "left join pl_liquidacion p on (p.codemp=e.codigo) where p.cerrado='S' and c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t " & _
          "where t.codemp=c.codemp group by codemp) and (situacion = 0 and fec_cese BETWEEN '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01' and " & _
          "'" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & Day(DateSerial(Left(AnoMes, 4), val(Right(AnoMes, 2)) + 1, 0)) & "') " & str & Str1 & ") as totaporte, " & _
          "(select sum(round(((select porcseg from pl_afppar f where f.anomes = '" & AnoMes & "' and codafp = (select codigo from afp where nombre = p.afp)) *(vacsinafp))/100,2)) " & _
          "from (empleado e inner join contrato c on (e.codigo=c.codemp)) left join pl_liquidacion p on (p.codemp=e.codigo) where p.cerrado='S' and c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) " & _
          "and (situacion = 0 and fec_cese BETWEEN '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01' and '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & Day(DateSerial(Left(AnoMes, 4), val(Right(AnoMes, 2)) + 1, 0)) & "') " & str & Str1 & ") as totseg, " & _
          "(select sum(round(((select porccom from pl_afppar f where f.anomes = '" & AnoMes & "' and codafp = (select codigo from afp where nombre = p.afp)) *(vacsinafp))/100,2)) " & _
          "from (empleado e inner join contrato c on (e.codigo=c.codemp)) left join pl_liquidacion p on (p.codemp=e.codigo) where p.cerrado='S' and c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) " & _
          "and (situacion = 0 and fec_cese BETWEEN '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01' and '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & Day(DateSerial(Left(AnoMes, 4), val(Right(AnoMes, 2)) + 1, 0)) & "') " & str & Str1 & ") as totcom, " & _
          "(select sum(if(e.sctr = 'S',ROUND(ifnull((vacsinafp)*0.02,0),2),0)) " & _
          "from (empleado e inner join contrato c on (e.codigo=c.codemp)) left join pl_liquidacion p on (p.codemp=e.codigo) where p.cerrado='S' and c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) " & _
          "and (situacion = 0 and fec_cese BETWEEN '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01' and " & _
          "'" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & Day(DateSerial(Left(AnoMes, 4), val(Right(AnoMes, 2)) + 1, 0)) & "') " & str & Str1 & ") as totdospor," & Interes & " as mora1," & Interes1 & " as mora2 " & _
          "from (empleado e inner join contrato c on (e.codigo=c.codemp)) left join pl_liquidacion p on (p.codemp=e.codigo) " & _
          "where p.cerrado='S' and c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) and (situacion = 0 and fec_cese BETWEEN '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01' and " & _
          "'" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & Day(DateSerial(Left(AnoMes, 4), val(Right(AnoMes, 2)) + 1, 0)) & "') " & str & Str1 & " order by nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Exit Sub
    End If
    EjecutaReporte App.Path & "\Reportes\Planilla\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_VacContable(Anio As String)
    Dim SQL As String, StrCampos As String, I As Integer, Dif As Integer
    Dim k As Integer, StrCond As String, FecVaca As String
    Dim myrs As MYSQL_RS, RQ As MYSQL_RS, Ani As Long, Mes As String
    Dim Flg As Boolean, strTemp As String, StrCampos1 As String
    If MsgBox("¿Desea ver el Resumen?", vbQuestion + vbYesNo, gsNomSW) = vbYes Then
        If Anio >= "2008" Then
            Screen.MousePointer = vbHourglass
            SQL = "delete from rh_tempvacaciones where usuario = '" & strUsuarioId & "'"
            oConexionMYSQL.Execute SQL
            SQL = "insert into rh_tempvacaciones(codemp,enev,febv,marv,abrv,mayv,usuario) " & _
                  "select e.codigo,c.si" & Anio & ",convert(ifnull(k.trab,0),unsigned) as provision, " & _
                  "convert(if(ifnull(k1.trab,0) > 0 or ifnull(k2.trab,0) > 0,if(if(ifnull(e.fec_cese,'')='','',year(e.fec_cese))='" & Anio & "',0,360-ifnull(k.trab,0)),0),unsigned) as provaj, " & _
                  "convert(ifnull(k1.trab,0)*12*(-1),SIGNED) as fisicas,convert(ifnull(k2.Trab, 0) * 12 * (-1), SIGNED) As vendidas,'" & strUsuarioId & "' " & _
                  "from cnprovvacaciones c left join empleado e on(c.codemp=e.codigo) left join (select emp,sum(cant) as trab from pl_tareo p " & _
                  "where tipo = 1 and rub = 121 and left(anomes,4)='" & Anio & "' group by p.emp) k on (k.emp=c.codemp) " & _
                  "left join (select emp,sum(cant) as trab from pl_tareo p where tipo = 4 and rub = 121 and left(anomes,4)='" & Anio & "' group by emp) k1 on (k1.emp=c.codemp) " & _
                  "left join (select emp,sum(cant) as trab from pl_tareo p where tipo = 4 and rub = 117 and left(anomes,4)='" & Anio & "' group by emp) k2 on (k2.emp=c.codemp) " & _
                  "where ((e.fec_cese = '') or (year(e.fec_cese) >= '" & Anio & "')) and e.codigo not in (select p.codigo from empleado p where p.tipo not in (3,4) and (p.fec_ingreso='' and p.fec_cese= '') " & _
                  "Union select k.codemp from (select o.codemp,o.f_inicio from contrato o where o.codigo = 'CN01') as k where year(k.f_inicio) > '" & Anio & "') order by e.codigo"
            oConexionMYSQL.Execute SQL
            SQL = "select codigo,fec_ingreso,fec_cese from cnprovvacaciones c left join empleado e on (c.codemp=e.codigo) " & _
                  "where tipo not (3,4) order by codigo"
            Set myrs = oConexion.EjecutaSelectRS(SQL)
            Do While Not myrs.EOF()
                If Trim(myrs.Fields("FEC_CESE")) <> "" Then
                    If Year(CDate(myrs.Fields("FEC_CESE"))) = Anio Then
                        SQL = "SELECT ifnull(sum(monto_viatico),0) as numvac FROM CALENDARIO C WHERE C.MOVEMP = '02' AND C.CODEMP = '" & Trim(myrs.Fields("codigo")) & "' ORDER BY PERIODO"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If Not RQ.EOF() Then
                            FecVaca = ""
                            If (val(RQ.Fields("numvac")) Mod 30) <> 0 Then
                                Dif = Int(val(RQ.Fields("numvac")) / 30)
                                I = Int((val(RQ.Fields("numvac")) - (Dif * 30)) * 12 / 30)
                                FecVaca = DateAdd("yyyy", Dif, Format(myrs.Fields("fec_ingreso"), "dd/mm/yyyy"))
                                FecVaca = DateAdd("m", I, Format(FecVaca, "dd/mm/yyyy"))
                            Else
                                FecVaca = DateAdd("yyyy", val(RQ.Fields("numvac")) / 30, Format(myrs.Fields("fec_ingreso"), "dd/mm/yyyy"))
                            End If
                        End If
                        Ani = 0
                        Flg = False
                        FecVaca = IIf(Mid(FecVaca, 1, 5) = "02/01", "31/12/" & Year(FecVaca) - 1, FecVaca)
                        FecVaca = IIf(Day(FecVaca) = 31, CDate(FecVaca) - 1, FecVaca)
                        Do While CDate(DateAdd("yyyy", 1, FecVaca)) < CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))
                            FecVaca = DateAdd("yyyy", 1, FecVaca)
                            Ani = Ani + 1
                        Loop
                        Ani = Ani * 360
                        Dif = 0
                        If Day(FecVaca) = 30 Then Flg = True
                        Do While CDate(DateAdd("m", 1, FecVaca)) <= CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))
                            FecVaca = DateAdd("m", 1, FecVaca)
                            Dif = Dif + 1
                        Loop
                        Dif = Dif * 30
                        I = 0
                        If Flg = True Then
                            If Day(FecVaca) <> 30 Then
                                FecVaca = "30/" & Mid(FecVaca, 4, 7)
                            End If
                        End If
                        If CDate(FecVaca) <> CDate(IIf(Day(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy")) = 31, CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy")) - 1, Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))) Then
                            FecVaca = IIf(Flg = True, FecVaca, CDate(FecVaca) - 1)
                            FecVaca = IIf(Day(FecVaca) = 31, CDate(FecVaca) - 1, FecVaca)
                            If Month(FecVaca) < Month(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))) Then
                                I = 30 - IIf((Day(CDate(FecVaca)) = 29 And Month(CDate(FecVaca)) = 2) Or (Day(CDate(FecVaca)) = 28 And Month(CDate(FecVaca)) = 2), 30, Day(CDate(FecVaca)))
                                I = I + IIf(Day(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))) = 31, 30, Day(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))))
                            Else
                                I = IIf(Day(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))) = 31, 30, Day(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy")))) - Day(CDate(FecVaca))
                            End If
                        End If
                        SQL = "UPDATE rh_tempvacaciones SET junv=(-1)*" & Ani + Dif + I & " where codemp = '" & myrs.Fields("codigo") & "' and usuario = '" & strUsuarioId & "'"
                        oConexionMYSQL.Execute SQL
                    End If
                Else
                    If Anio = Year(Date) Then
                        SQL = "select * from pl_tareo where anomes = '" & Anio & Right("00" & Month(Date), 2) & "' and tipo = 1 AND emp = '" & myrs.Fields("codigo") & "'"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If RQ.EOF() Then
                            SQL = "update rh_tempvacaciones set febv = febv + " & Day(Date) & " where codemp = '" & myrs.Fields("codigo") & "' and usuario = '" & strUsuarioId & "'"
                            oConexionMYSQL.Execute SQL
                        End If
                    End If
                End If
                myrs.MoveNext
            Loop
            SQL = "select CONCAT_WS(' ',e.apepat,e.apemat,e.nombre1,e.nombre2) AS nom,date_format(e.fec_ingreso,'%d/%m/%Y') as fecing, " & _
                  "IF(e.fec_cese<>'',if(e.fec_cese > '" & Format(IIf(strAnoSistema = Anio, Date, "31/12/" & Anio & ""), "yyyy/mm/dd") & "','',date_format(e.fec_cese,'%d/%m/%Y')),'') as fecfin,convert(e.codigo,unsigned) as cod,enev as saldoini,febv as prov,marv as provaj, " & _
                  "abrv as fisicas,mayv as vendidas,IFNULL(junv,0) as liquidadas,IFNULL(if(e.fec_cese<>'',(-1)*(enev+febv+marv+abrv+mayv+junv),0),0) as ajuste " & _
                  "from rh_tempvacaciones r left join empleado e on (r.codemp=e.codigo) " & _
                  "where r.usuario = '" & strUsuarioId & "' order by nom"
            Set myrs = oConexion.EjecutaSelectRS(SQL)
            If myrs.RecordCount <= 0 Then
                MsgBox "No Existen Registros", vbInformation + vbOKOnly, "BRANDTADMIN"
                Set myrs = Nothing
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            Do While Not RQ.EOF()
                SQL = "update cnprovvacaciones set si" & Anio + 1 & "=" & CDbl(RQ.Fields("saldoini")) + CDbl(RQ.Fields("prov")) + CDbl(RQ.Fields("provaj")) + CDbl(RQ.Fields("fisicas")) + CDbl(RQ.Fields("vendidas")) + CDbl(RQ.Fields("liquidadas")) + CDbl(RQ.Fields("ajuste")) & " " & _
                      "where codemp = '" & Right("00000000000" & RQ.Fields("cod"), 11) & "'"
                oConexionMYSQL.Execute SQL
                RQ.MoveNext
            Loop
            Set RQ = Nothing
            m_Anomes = Anio
            Screen.MousePointer = vbDefault
            EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
        End If
    Else
        k = 1
        StrCampos = ""
        StrCampos1 = ""
        StrCond = ""
        Screen.MousePointer = vbHourglass
        For I = 1 To 12
            StrCampos = StrCampos & ",IFNULL(p" & I & ".cant,0) as '" & NombreMes(Right("00" & I, 2), True) & "V'"
            StrCampos1 = StrCampos1 & ",IFNULL(p" & I & ".cant1,0) as '" & NombreMes(Right("00" & I, 2), True) & "VtaVac'"
            StrCond = StrCond & " Left Join (select cant,emp,(((sbasico*cant)/30)+if((select COUNT(*) from pl_tareo e where e.tipo = 4 and " & _
                                "anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '121' and (e.emp=t.emp))<=0,(select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '201' and e.emp=t.emp),((select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '201' and e.emp=t.emp)*cant)/(select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub IN ('121','117') and e.emp=t.emp))+if((select COUNT(*) from pl_tareo e where e.tipo = 4 and " & _
                                "anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '121' and (e.emp=t.emp))<=0,(select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '1001' and e.emp=t.emp),((select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '1001' and e.emp=t.emp)*cant)/(select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub IN ('121','117') and e.emp=t.emp))) as cant1 from pl_tareo t where t.tipo = 4 and t.rub = 117 and " & _
                                "t.anomes= '" & Anio & Right("00" & I, 2) & "') as p" & I & " on (p" & I & ".emp=e.codigo) "
            k = k + 1
        Next
        SQL = "delete from rh_tempvacaciones where usuario = '" & strUsuarioId & "'"
        oConexionMYSQL.Execute SQL
        SQL = "insert into rh_tempvacaciones(Codemp,EneV,FebV,MarV,AbrV,MayV,JunV,JulV,AgoV,SepV,OctV,NovV,DicV, " & _
              "EneVtaVac,FebVtaVac,MarVtaVac,AbrVtaVac,MayVtaVac,JunVtaVac,JulVtaVac,AgoVtaVac,SepVtaVac,OctVtaVac,NovVtaVac,DicVtaVac,usuario) " & _
              "select DISTINCT e.codigo " & StrCampos & StrCampos1 & ",'" & strUsuarioId & "'" & _
              " from empleado e " & StrCond & " Where e.Tipo not in (3,4) and ((e.fec_cese = '') or (year(e.fec_cese) >= '" & Anio & "')) " & _
              "and e.codigo not in (select p.codigo from empleado p where p.tipo not in (3,4) and (p.fec_ingreso='' and p.fec_cese= '') " & _
              "Union select k.codemp from (select o.codemp,o.f_inicio from contrato o where o.codigo = 'CN01') as k where year(k.f_inicio) > '" & Anio & "')"
        oConexionMYSQL.Execute SQL
        
        k = 1
        StrCampos = ""
        StrCampos1 = ""
        StrCond = ""
        strTemp = ""
        For I = 1 To 12
            StrCampos = StrCampos & ",CONVERT(IFNULL(p" & I & ".cant,0),UNSIGNED) as '" & NombreMes(Right("00" & I, 2), True) & "T'"
            strTemp = strTemp & "CONVERT(IFNULL(p" & I & ".cant,0),UNSIGNED) + "
            StrCond = StrCond & " Left Join (select cant,emp from pl_tareo t where t.tipo = 1 and t.rub = 121 and " & _
                                "t.anomes= '" & Anio & Right("00" & I, 2) & "') as p" & I & " on (p" & I & ".emp=e.codigo) "
        Next
        strTemp = Mid(strTemp, 1, Len(strTemp) - 2) & " as TOTAL_T,''"
        k = I + 1
        For I = 1 To 12
            If I = 1 Then StrCampos = StrCampos & "," & strTemp: strTemp = ""
            StrCampos = StrCampos & ",CONVERT(IFNULL(p" & k & ".cant,0),UNSIGNED) as '" & NombreMes(Right("00" & I, 2), True) & "D'"
            StrCampos1 = StrCampos1 & ",IFNULL(p" & k & ".cant1,0) as '" & NombreMes(Right("00" & I, 2), True) & "RemVac',ifnull(r." & NombreMes(Right("00" & I, 2), True) & "VtaVac,0) as '" & NombreMes(Right("00" & I, 2), True) & ".Vta.Vac.'"
            strTemp = strTemp & "CONVERT(IFNULL(p" & k & ".cant,0),UNSIGNED) + "
            StrCond = StrCond & " Left Join (select cant,emp,(((sbasico*cant)/30)+if((select COUNT(*) from pl_tareo e where e.tipo = 4 and " & _
                                "anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '117' and (e.emp=t.emp))<=0,(select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '201' and e.emp=t.emp),((select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '201' and e.emp=t.emp)*cant)/(select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub IN ('121','117') and e.emp=t.emp))+if((select COUNT(*) from pl_tareo e where e.tipo = 4 and " & _
                                "anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '117' and (e.emp=t.emp))<=0,(select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '1001' and e.emp=t.emp),((select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub = '1001' and e.emp=t.emp)*cant)/(select ifnull(sum(cant),0) from pl_tareo e where e.tipo = 4 " & _
                                "and anomes = '" & Anio & Right("00" & I, 2) & "' and rub IN ('121','117') and e.emp=t.emp))) as cant1 from pl_tareo t where t.tipo = 4 and t.rub = 121 and " & _
                                "t.anomes= '" & Anio & Right("00" & I, 2) & "') as p" & k & " on (p" & k & ".emp=e.codigo) "
            k = k + 1
        Next
        strTemp = Mid(strTemp, 1, Len(strTemp) - 2) & " as TOTAL_D,''"
        StrCampos = StrCampos & "," & strTemp
        SQL = "select e.codigo,fec_ingreso,fec_cese from rh_tempvacaciones r left join empleado e on(r.codemp=e.codigo) where e.tipo not in (3,4) order by e.codigo"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
        Do While Not myrs.EOF()
            If Trim(myrs.Fields("FEC_CESE")) <> "" Then
                If Year(CDate(myrs.Fields("FEC_CESE"))) = Anio Then
                    SQL = "SELECT ifnull(sum(monto_viatico),0) as numvac FROM CALENDARIO C WHERE C.MOVEMP = '02' AND C.CODEMP = '" & Trim(myrs.Fields("codigo")) & "' ORDER BY PERIODO"
                    Set RQ = oConexion.EjecutaSelectRS(SQL)
                    If Not RQ.EOF() Then
                        FecVaca = ""
                        If (val(RQ.Fields("numvac")) Mod 30) <> 0 Then
                            Dif = Int(val(RQ.Fields("numvac")) / 30)
                            I = Int((val(RQ.Fields("numvac")) - (Dif * 30)) * 12 / 30)
                            FecVaca = DateAdd("yyyy", Dif, Format(myrs.Fields("fec_ingreso"), "dd/mm/yyyy"))
                            FecVaca = DateAdd("m", I, Format(FecVaca, "dd/mm/yyyy"))
                        Else
                            FecVaca = DateAdd("yyyy", val(RQ.Fields("numvac")) / 30, Format(myrs.Fields("fec_ingreso"), "dd/mm/yyyy"))
                        End If
                    End If
                    Ani = 0
                    Flg = False
                    FecVaca = IIf(Mid(FecVaca, 1, 5) = "02/01", "31/12/" & Year(FecVaca) - 1, FecVaca)
                    FecVaca = IIf(Day(FecVaca) = 31, CDate(FecVaca) - 1, FecVaca)
                    Do While CDate(DateAdd("yyyy", 1, FecVaca)) < CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))
                        FecVaca = DateAdd("yyyy", 1, FecVaca)
                        Ani = Ani + 1
                    Loop
                    Ani = Ani * 360
                    Dif = 0
                    If Day(FecVaca) = 30 Then Flg = True
                    Do While CDate(DateAdd("m", 1, FecVaca)) <= CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))
                        FecVaca = DateAdd("m", 1, FecVaca)
                        Dif = Dif + 1
                    Loop
                    Dif = Dif * 30
                    I = 0
                    If Flg = True Then
                        If Day(FecVaca) <> 30 Then
                            FecVaca = "30/" & Mid(FecVaca, 4, 7)
                        End If
                    End If
                    If CDate(FecVaca) <> CDate(IIf(Day(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy")) = 31, CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy")) - 1, Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))) Then
                        FecVaca = IIf(Flg = True, FecVaca, CDate(FecVaca) - 1)
                        FecVaca = IIf(Day(FecVaca) = 31, CDate(FecVaca) - 1, FecVaca)
                        If Month(FecVaca) < Month(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))) Then
                            I = 30 - IIf((Day(CDate(FecVaca)) = 29 And Month(CDate(FecVaca)) = 2) Or (Day(CDate(FecVaca)) = 28 And Month(CDate(FecVaca)) = 2), 30, Day(CDate(FecVaca)))
                            I = I + IIf(Day(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))) = 31, 30, Day(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))))
                        Else
                            I = IIf(Day(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))) = 31, 30, Day(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy")))) - Day(CDate(FecVaca))
                        End If
                    End If
                    SQL = "UPDATE rh_tempvacaciones SET " & NombreMes(Right("00" & Month(CDate(Format(myrs.Fields("fec_cese"), "dd/mm/yyyy"))), 2), True) & "L=" & Ani + Dif + I & " " & _
                          "where codemp = '" & myrs.Fields("codigo") & "'"
                    oConexionMYSQL.Execute SQL
                End If
            End If
            myrs.MoveNext
        Loop
        SQL = "select DISTINCT convert(e.codigo,UNSIGNED) as Codigo,date_format(e.fec_ingreso,'%d/%m/%Y') as 'Fecha Ingreso', " & _
              "IF(e.fec_cese<>'',if(e.fec_cese > '" & Format(IIf(strAnoSistema = Anio, Date, "31/12/" & Anio & ""), "yyyy/mm/dd") & "','',date_format(e.fec_cese,'%d/%m/%Y')),'') as 'Fecha Liquid', " & _
              "CONCAT_WS(' ',e.apepat,e.apemat,e.nombre1,e.nombre2) AS Trabajador " & StrCampos & _
              ",IFNULL(r.enev,0) as 'ENEV',IFNULL(r.febv,0) as 'FEBV',IFNULL(r.marv,0) as 'MARV',IFNULL(r.abrv,0) as 'ABRV', " & _
              "IFNULL(r.mayv,0) as 'MAYV',IFNULL(r.junv,0) as 'JUNV',IFNULL(r.julv,0) as 'JULV',IFNULL(r.agov,0) as 'AGOV', " & _
              "IFNULL(r.sepv,0) as 'SEPV',IFNULL(r.octv,0) as 'OCTV',IFNULL(r.novv,0) as 'NOVV',IFNULL(r.dicv,0) as 'DICV', " & _
              "(IFNULL(r.enev,0)+IFNULL(r.febv,0)+IFNULL(r.marv,0)+IFNULL(r.abrv,0)+IFNULL(r.mayv,0)+IFNULL(r.junv,0)+IFNULL(r.julv,0) " & _
              "+IFNULL(r.agov,0)+IFNULL(r.sepv,0)+IFNULL(r.octv,0)+IFNULL(r.novv,0)+IFNULL(r.dicv,0)) as Total_V,'' " & _
              ",IFNULL(r.enel,0) as 'ENEL',IFNULL(r.febl,0) as 'FEBL',IFNULL(r.marl,0) as 'MARL',IFNULL(r.abrl,0) as 'ABRL', " & _
              "IFNULL(r.mayl,0) as 'MAYL',IFNULL(r.junl,0) as 'JUNL',IFNULL(r.jull,0) as 'JULL',IFNULL(r.agol,0) as 'AGOL', " & _
              "IFNULL(r.sepl,0) as 'SEPL',IFNULL(r.octl,0) as 'OCTL',IFNULL(r.novl,0) as 'NOVL',IFNULL(r.dicl,0) as 'DICL', " & _
              "(IFNULL(r.enel,0)+IFNULL(r.febl,0)+IFNULL(r.marl,0)+IFNULL(r.abrl,0)+IFNULL(r.mayl,0)+IFNULL(r.junl,0)+IFNULL(r.jull,0) " & _
              "+IFNULL(r.agol,0)+IFNULL(r.sepl,0)+IFNULL(r.octl,0)+IFNULL(r.novl,0)+IFNULL(r.dicl,0)) as Total_L,'' " & StrCampos1 & _
              " from empleado e left join rh_tempvacaciones r on (e.codigo=r.codemp) " & StrCond & " Where e.Tipo not in (3,4) and r.usuario = '" & strUsuarioId & "' order by trabajador"
        Set myrs = oConexion.EjecutaSelectRS(SQL)
        If myrs.RecordCount <= 0 Then
            MsgBox "No Existen Registros", vbInformation + vbOKOnly, "BRANDTADMIN"
            Set myrs = Nothing
            Exit Sub
        End If
        Exportar_Excel_rs Rep_Documents & "\VacacionesCont.xls", myrs
        Screen.MousePointer = vbDefault
    End If
    Set myrs = Nothing
End Sub
Public Sub sp_Rep_TotalTrab(Division As String, fecha As String)
    Dim SQL As String, SQL1 As String, Tipo As String
    Dim myrs As MYSQL_RS, myrs1 As MYSQL_RS
    Dim SCodEmp As String, SFecha As String, SWhere As String
    Screen.MousePointer = vbHourglass
    If MsgBox("¿Desea ver el Reporte por divisiones para HCM?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then Tipo = "H" Else Tipo = "L"
    If Tipo = "H" Then
        If Division <> "" And Division <> "0006" Then
            SCodEmp = " and t.divhcm = '" & Division & "'"
            SWhere = " AND codigo = (select descrip from cnmdepar where coddep = '" & Division & "')"
        End If
        SQL = "select '' AS codigo,'' as nom,'' as division,tot,'' as fecingreso,tipo,codigo as nom1 from TOTTRABPORDIA where tipo = 'C' AND FECHA = '" & fecha & "' AND FECINGRESO = 'H' " & SWhere & " UNION " & _
              "select t.codigo,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nom, " & _
              "(select descrip from cnmdepar c where c.coddep=t.divhcm) as division,0 as tot, " & _
              "date_format(t.fecingreso,'%d/%m/%Y') as fecingreso,t.tipo,'' as nom1 from TOTTRABPORDIA t " & _
              "left join empleado e on (t.codigo=e.codigo) where t.tipo = 'D' " & SCodEmp & " AND FECHA = '" & fecha & "' order by codigo,nom1"
    Else
        If Division <> "" And Division <> "0006" Then
            SCodEmp = " and t.divlocal = '" & Division & "'"
            SWhere = " AND codigo = (select descriplocal from cnmdepar where coddep = '" & Division & "')"
        End If
        SQL = "select '' AS codigo,'' as nom,'' as division,tot,'' as fecingreso,tipo,codigo as nom1 from TOTTRABPORDIA where tipo = 'C' AND FECHA = '" & fecha & "' AND FECINGRESO = 'L' " & SWhere & " UNION " & _
              "select t.codigo,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nom, " & _
              "(select descriplocal from cnmdepar c where c.coddep=t.divlocal) as division,0 as tot, " & _
              "date_format(t.fecingreso,'%d/%m/%Y') as fecingreso,t.tipo,'' as nom1 from TOTTRABPORDIA t " & _
              "left join empleado e on (t.codigo=e.codigo) where t.tipo = 'D' " & SCodEmp & " AND FECHA = '" & fecha & "' order by codigo,nom1"
    End If
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    If myrs.RecordCount <= 0 Then
        MsgBox "No Existen Registros", vbInformation + vbOKOnly, "NOVADMIN"
        Set myrs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, myrs
    Set myrs = Nothing
End Sub
Public Sub CompararHCM_Local(Cmd As CommonDialog)
    Dim NomArchivo As String, columna As Integer, I As Integer, FechaIni As String, FechaFin As String
    Dim Hoja As String, ContTot As Integer, ContPer As Integer, Context As Integer, ContPerLocal As Integer, ContExtLocal As Integer
    Dim SQL As String, SQL1 As String, RQ As MYSQL_RS, Val1 As String, Val2 As String, FilaComp As Integer
    Dim NumHojas As Integer, FlgFec As Boolean, Pos As Integer, NOmCol As String
    Dim Fil As Integer, Col As Integer, SumTot As Double, SumPer As Double, SumExt As Double, SumPerLocal As Double, SumExtLocal As Double
    Cmd.CancelError = True
    Cmd.Flags = cdlOFNFileMustExist Or _
            cdlOFNHideReadOnly Or _
            cdlOFNExplorer Or _
            cdlOFNLongNames
    On Error Resume Next
    Cmd.ShowOpen
    If err.Number = cdlCancel Or err.Number <> 0 Then Exit Sub
    NomArchivo = Cmd.Filename
    Excel.Application.Workbooks.Open Filename:=NomArchivo
    Excel.Application.Visible = False
    SQL = "DELETE FROM RH_TMPASISTENCIAS where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQL
    For NumHojas = 1 To 4
        Hoja = Excel.Application.Sheets(NumHojas).name
        FlgFec = False
        Pos = 1: FechaIni = "": FechaFin = ""
AQUI:
        For I = 1 To 12
            If InStr(Pos, UCase(Trim(Worksheets(Hoja).Cells(1, 1).Value)), NombreMes(Right("00" & I, 2), True)) > 0 Then
                If FlgFec = False Then
                    Pos = InStr(Pos, UCase(Trim(Worksheets(Hoja).Cells(1, 1).Value)), NombreMes(Right("00" & I, 2), True)) + 11
                    FechaIni = FormatDateTime(Mid(Trim(Worksheets(Hoja).Cells(1, 1).Value), InStr(1, UCase(Trim(Worksheets(Hoja).Cells(1, 1).Value)), NombreMes(Right("00" & I, 2), True)) - 3, 11), vbShortDate)
                    FechaIni = Format(FechaIni, "yyyy/mm/dd")
                    FlgFec = True
                    GoTo AQUI
                Else
                    FechaFin = FormatDateTime(Mid(Trim(Worksheets(Hoja).Cells(1, 1).Value), InStr(Pos, UCase(Trim(Worksheets(Hoja).Cells(1, 1).Value)), NombreMes(Right("00" & I, 2), True)) - 3, 11), vbShortDate)
                    FechaFin = Format(FechaFin, "yyyy/mm/dd")
                    Exit For
                End If
            End If
        Next
        Screen.MousePointer = vbHourglass
        Select Case NumHojas
            Case 1:
                NOmCol = "Country Of Birth": Val1 = "PERU": Val2 = "COLOMBIA": FilaComp = 5
            Case 2:
                NOmCol = "Nationality": Val1 = "PERUVIAN": Val2 = "COLOMBIAN": FilaComp = 3
            Case 3:
                NOmCol = "Country": Val1 = "PE": Val2 = "CO": FilaComp = 5
            Case 4:
                NOmCol = "Work Location": Val1 = "PE": Val2 = "CO": FilaComp = 3
        End Select
        ContTot = 0: SumTot = 0
        For columna = 1 To 40
            If Trim(Worksheets(Hoja).Cells(FilaComp, columna).Value) = NOmCol Then
                I = FilaComp + 1
                ContPer = 0: Context = 0: SumPer = 0: SumExt = 0
                Do While Trim(Worksheets(Hoja).Cells(I, columna).Value) <> ""
                    If UCase(Mid(Trim(Worksheets(Hoja).Cells(I, columna).Value), 1, Len(Val1))) = Val1 Then
                        If NumHojas = 4 Then
                            If (Mid(Trim(Worksheets(Hoja).Cells(I, columna).Value), 1, 15) = "Suarez , Carlos") Or (Mid(Trim(Worksheets(Hoja).Cells(I, columna).Value), 1, 15) = "Leiva, Hernando") Then
                                SumExt = SumExt + CDbl(Worksheets(Hoja).Cells(I, columna + 10).Value)
                            Else
                                SumPer = SumPer + CDbl(Worksheets(Hoja).Cells(I, columna + 10).Value)
                            End If
                        Else
                            ContPer = ContPer + 1
                        End If
                    End If
                    If NumHojas < 4 Then
                        If UCase(Mid(Trim(Worksheets(Hoja).Cells(I, columna).Value), 1, Len(Val2))) = Val2 Then
                            Context = Context + 1
                        End If
                    End If
                    ContTot = ContTot + 1
                    If NumHojas = 4 Then SumTot = SumTot + CDbl(Worksheets(Hoja).Cells(I, columna + 10).Value)
                    I = I + 1
                Loop
                Exit For
            End If
        Next
        ContPerLocal = 0: ContExtLocal = 0: SumExtLocal = 0: SumPerLocal = 0
        Select Case NumHojas
            Case 1
                SQL = "select count(*) as cant from contrato c left join empleado e on (c.codemp=e.codigo) " & _
                      "where (('" & Format(FechaIni, "yyyy/mm/dd") & "' between f_inicio and if(fechacese='',f_termino,fechacese)) " & _
                      "or (f_termino = '' and f_inicio <= '" & Format(FechaIni, "yyyy/mm/dd") & "')) and e.tipo not in (3,4) and e.nacionalidad = '01'"
                SQL1 = "select count(*) as cant from contrato c left join empleado e on (c.codemp=e.codigo) " & _
                      "where (('" & Format(FechaIni, "yyyy/mm/dd") & "' between f_inicio and if(fechacese='',f_termino,fechacese)) " & _
                      "or (f_termino = '' and f_inicio <= '" & Format(FechaIni, "yyyy/mm/dd") & "')) and e.tipo not in (3,4) and e.nacionalidad = '02'"
            Case 2
                SQL = "SELECT count(*) as cant FROM EMPLEADO WHERE FEC_INGRESO >= '" & FechaIni & "' AND FEC_INGRESO <= '" & FechaFin & "' " & _
                      "AND tipo not in (3,4) and nacionalidad = '01'"
                SQL1 = "SELECT count(*) as cant FROM EMPLEADO WHERE FEC_INGRESO >= '" & FechaIni & "' AND FEC_INGRESO <= '" & FechaFin & "' " & _
                      "AND tipo not in (3,4) and nacionalidad = '02'"
            Case 3
                SQL = "SELECT COUNT(E.codigo) as cant FROM empleado e inner join contrato c on (e.codigo=c.codemp) " & _
                      "where e.nacionalidad = '01' and c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) " & _
                      "and (situacion = 0 and fec_cese BETWEEN '" & FechaIni & "' and '" & FechaFin & "')"
                SQL1 = "SELECT COUNT(E.codigo) as cant FROM empleado e inner join contrato c on (e.codigo=c.codemp) " & _
                       "where e.nacionalidad = '02' and c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) " & _
                       "and (situacion = 0 and fec_cese BETWEEN '" & FechaIni & "' and '" & FechaFin & "')"
            Case 4
                SQL = "select round(sum(sbasico*12),2) as cant from contrato c left join empleado e on (c.codemp=e.codigo) " & _
                      "where (('" & Format(FechaIni, "yyyy/mm/dd") & "' between f_inicio and if(fechacese='',f_termino,fechacese)) " & _
                      "or (f_termino = '' and f_inicio <= '" & Format(FechaIni, "yyyy/mm/dd") & "')) and e.tipo not in (3,4) and e.nacionalidad = '01'"
                SQL1 = "select round(sum(sbasico*12),2) as cant from contrato c left join empleado e on (c.codemp=e.codigo) " & _
                      "where (('" & Format(FechaIni, "yyyy/mm/dd") & "' between f_inicio and if(fechacese='',f_termino,fechacese)) " & _
                      "or (f_termino = '' and f_inicio <= '" & Format(FechaIni, "yyyy/mm/dd") & "')) and e.tipo not in (3,4) and e.nacionalidad = '02'"
        End Select
        If SQL <> "" Then
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                If NumHojas = 4 Then SumPerLocal = RQ.Fields("cant") Else ContPerLocal = RQ.Fields("cant")
            End If
        End If
        If SQL1 <> "" Then
            Set RQ = oConexion.EjecutaSelectRS(SQL1)
            If Not RQ.EOF() Then
                If NumHojas = 4 Then SumExtLocal = RQ.Fields("cant") Else ContExtLocal = RQ.Fields("cant")
            End If
        End If
        Set RQ = Nothing
        If ContTot > 0 Or SumTot > 0 Then
            Select Case NumHojas
                Case 1: Val1 = "TOTAL DE EMPLEADOS"
                Case 2: Val1 = "TOTAL DE INGRESOS"
                Case 3: Val1 = "TOTAL DE EGRESOS"
                Case 4: Val1 = "TOTAL DE SALARIO ANUAL"
            End Select
            SQL = "insert into RH_TMPASISTENCIAS(nombres,division,envio,codigo,sede,tiposede,situacion,item,usuario) " & _
                  "values('" & Val1 & "','" & Format(FechaIni, "dd/mm/yyyy") & "','" & Format(FechaFin, "dd/mm/yyyy") & "', " & _
                  "'" & IIf(NumHojas < 4, ContTot, SumTot) & "','" & IIf(NumHojas < 4, ContPer, SumPer) & "', " & _
                  "'" & IIf(NumHojas < 4, Context, SumExt) & "','" & IIf(NumHojas < 4, ContPerLocal, CStr(SumPerLocal)) & "', " & _
                  "" & IIf(NumHojas < 4, ContExtLocal, CStr(SumExtLocal)) & ",'" & strUsuarioId & "')"
            oConexionMYSQL.Execute SQL
            Dim Nom As String, NomDiv As String
            SQL = "select * from tottrabpordia where fecha = '" & FechaIni & "'and tipo = 'C' and fecingreso = 'H'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If NumHojas = 4 Then
                For columna = 1 To 40
                    I = 4
                    If Trim(Worksheets(Hoja).Cells(FilaComp, columna).Value) = NOmCol Then
                        ContPer = 0
                        Nom = ""
                        Do While Trim(Worksheets(Hoja).Cells(I, columna).Value) <> ""
                            If UCase(Mid(Trim(Worksheets(Hoja).Cells(I, columna).Value), 1, Len("PE"))) = "PE" Then
                                If Nom = Trim(Worksheets(Hoja).Cells(I, 3).Value) Then
                                    ContPer = ContPer + 1
                                    I = I + 1
                                Else
                                    If Nom <> "" Then
                                        SQL = "insert into RH_TMPASISTENCIAS(nombres,codigo,usuario) " & _
                                              "values('" & Nom & "','" & ContPer & "','" & strUsuarioId & "')"
                                        oConexionMYSQL.Execute SQL
                                        Select Case UCase(Mid(Mid(Nom, 1, IIf(InStr(9, Nom, "-") = 0, Len(Nom), InStr(9, Nom, "-"))), 9, Len(Mid(Nom, 1, IIf(InStr(9, Nom, "-") = 0, Len(Nom), InStr(9, Nom, "-")))) - 10))
                                            Case "SHARED SERVICES": NomDiv = "SHARED SERVICES"
                                            Case "INSPECTION": NomDiv = "TUBULAR SERVICES"
                                            Case "MD TOCTO": NomDiv = "MD TOTCO"
                                            Case "REED HYCALOG": NomDiv = "REED HYCALOG"
                                            Case "SOLIDS CONTROL": NomDiv = "SOLID CONTROL"
                                            Case Else
                                                 NomDiv = "MD TOTCO"
                                        End Select
                                        RQ.MoveFirst
                                        Do While Not RQ.EOF()
                                            If Trim(RQ.Fields("codigo")) = NomDiv Then
                                                SQL = "update RH_TMPASISTENCIAS set situacion = '" & RQ.Fields("tot") & "' where nombres = '" & Nom & "'"
                                                oConexionMYSQL.Execute SQL
                                                Exit Do
                                            End If
                                            RQ.MoveNext
                                        Loop
                                    End If
                                    ContPer = 0
                                    Nom = Trim(Worksheets(Hoja).Cells(I, 3).Value)
                                End If
                            Else
                                I = I + 1
                            End If
                        Loop
                        If Nom <> "" Then
                            SQL = "insert into RH_TMPASISTENCIAS(nombres,codigo,usuario) " & _
                                  "values('" & Nom & "','" & ContPer & "','" & strUsuarioId & "')"
                            oConexionMYSQL.Execute SQL
                            Select Case UCase(Mid(Mid(Nom, 1, IIf(InStr(9, Nom, "-") = 0, Len(Nom), InStr(9, Nom, "-"))), 9, Len(Mid(Nom, 1, IIf(InStr(9, Nom, "-") = 0, Len(Nom), InStr(9, Nom, "-")))) - 10))
                                Case "SHARED SERVICES": NomDiv = "SHARED SERVICES"
                                Case "INSPECTION": NomDiv = "TUBULAR SERVICES"
                                Case "MD TOCTO": NomDiv = "MD TOTCO"
                                Case "REED HYCALOG": NomDiv = "REED HYCALOG"
                                Case "SOLIDS CONTROL": NomDiv = "SOLID CONTROL"
                                Case Else
                                    NomDiv = "MD TOTCO"
                            End Select
                            RQ.MoveFirst
                            Do While Not RQ.EOF()
                                If Trim(RQ.Fields("codigo")) = NomDiv Then
                                    SQL = "update RH_TMPASISTENCIAS set situacion = '" & RQ.Fields("tot") & "' where nombres = '" & Nom & "'"
                                    oConexionMYSQL.Execute SQL
                                    Exit Do
                                End If
                                RQ.MoveNext
                            Loop
                        End If
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    Excel.Application.Quit
    Dim o_Excel As Object, o_Libro As Object, o_Hoja As Object
    Set o_Excel = New Excel.Application
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    SQL = "select nombres as descrip,division as fecini,envio as fecfin,codigo as total,sede as totper,tiposede as totcol,situacion as local,item as ext " & _
          "from RH_TMPASISTENCIAS where usuario = '" & strUsuarioId & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    Fil = 3
    Col = 3
    Do While Not RQ.EOF()
        If Fil = 3 Then
            o_Hoja.Cells(Fil - 1, Col + 1).Value = "FEC.INI"
            o_Hoja.Cells(Fil - 1, Col + 2).Value = "FEC.FIN"
            o_Hoja.Cells(Fil - 1, Col + 3).Value = "TOTAL HCM"
            o_Hoja.Cells(Fil - 1, Col + 4).Value = "HCM PERU"
            o_Hoja.Cells(Fil - 1, Col + 5).Value = "HCM COL."
            o_Hoja.Cells(Fil - 1, Col + 6).Value = "LOCAL PERU"
            o_Hoja.Cells(Fil - 1, Col + 7).Value = "LOCAL EXT."
            o_Excel.Worksheets(1).Range(Trim(Chr(68)) & ":" & Trim(Chr(68))).NumberFormat = "@"
            o_Excel.Worksheets(1).Range(Trim(Chr(69)) & ":" & Trim(Chr(69))).NumberFormat = "@"
        End If
        o_Hoja.Cells(Fil, Col).Value = Trim(RQ.Fields("descrip"))
        o_Hoja.Cells(Fil, Col + 1).Value = Trim(RQ.Fields("fecini"))
        o_Hoja.Cells(Fil, Col + 2).Value = Trim(RQ.Fields("fecfin"))
        o_Hoja.Cells(Fil, Col + 3).Value = val(RQ.Fields("total"))
        o_Hoja.Cells(Fil, Col + 4).Value = val(RQ.Fields("totper"))
        o_Hoja.Cells(Fil, Col + 5).Value = val(RQ.Fields("totcol"))
        o_Hoja.Cells(Fil, Col + 6).Value = val(RQ.Fields("local"))
        o_Hoja.Cells(Fil, Col + 7).Value = val(RQ.Fields("ext"))
        Fil = Fil + 1
        RQ.MoveNext
    Loop
    o_Excel.Visible = True
    o_Hoja.SaveAs Rep_Documents & "\HCM_LOCAL.xls"
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
    Screen.MousePointer = vbDefault
Exit Sub
ErrorAbrir:
    Screen.MousePointer = vbDefault
    MsgBox "Ocurrió un error al momento de generar el archivo de Excel, " & vbNewLine & _
           "consulte con el administrador del sistema", vbExclamation + vbOKOnly, "NOVADMIN"
End Sub




Public Sub VerificaPlanilla_RRHH_Contabilidad(AnoMes As String)
Dim o_Excel As Object, o_Libro As Object, o_Hoja As Object, SQL As String, columna As Integer
Dim Fil As Integer, Col As Integer, myrs As MYSQL_RS, Sum1 As Double, Sum2 As Double, dia As String
Set o_Excel = New Excel.Application
Set o_Libro = o_Excel.Workbooks.Add
Set o_Hoja = o_Libro.Worksheets.Add

On Error GoTo ErrorAbrir
    
Dim cNombreHoja As String
cNombreHoja = BuscarNombreHoja


    Screen.MousePointer = vbHourglass
    'SUBSIDIO
    SQL = "select anomes,moneda,emp as codigo,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS nombres,convert(p.rub,unsigned) as rubro,p.Cant as monto " & _
          "from pl_tareo p left join empleado e on (p.emp=e.codigo) where anomes='" & AnoMes & "' and (rub='915' OR RUB='916') " & _
          "and (p.tipo='1' or p.tipo='5')"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 1).Value = "RRHH"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 2).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 1).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(70)) & ":" & Trim(Chr(70))).NumberFormat = "#,##0.00"
        o_Excel.Worksheets(1).Range(Trim(Chr(67)) & ":" & Trim(Chr(67))).NumberFormat = "@"
        o_Excel.Worksheets(1).Range(Trim(Chr(69)) & ":" & Trim(Chr(69))).NumberFormat = "@"
        Fil = 4: Col = 1: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("rubro"))
            o_Hoja.Cells(Fil, Col + 5).Value = CDbl(myrs.Fields("monto"))
            Sum1 = Sum1 + CDbl(myrs.Fields("monto"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 5).Value = Sum1
    End If
    SQL = "select a.anomes,a.voucher,a.auxiliar,a.codaux as codigo,a.moneda,a.coddoc as tipodoc,a.numdoc,a.concepto,a.Cargos " & _
          "from cnmovi as a where a.anomes='" & AnoMes & "' AND (cuenta='16803') AND a.numdoc like '%JUL%'"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 9).Value = "CONTABILIDAD"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 10).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 9).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(81)) & ":" & Trim(Chr(81))).NumberFormat = "#,##0.00"
        For columna = 0 To 8
            o_Excel.Worksheets(1).Range(Trim(Chr(73 + columna)) & ":" & Trim(Chr(73 + columna))).NumberFormat = "@"
        Next
        Fil = 4: Col = 9: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("voucher"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("auxiliar"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 5).Value = Trim(myrs.Fields("tipodoc"))
            o_Hoja.Cells(Fil, Col + 6).Value = Trim(myrs.Fields("numdoc"))
            o_Hoja.Cells(Fil, Col + 7).Value = Trim(myrs.Fields("concepto"))
            o_Hoja.Cells(Fil, Col + 8).Value = CDbl(myrs.Fields("cargos"))
            Sum1 = Sum1 + CDbl(myrs.Fields("cargos"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 8).Value = Sum1
    End If
    FormatoEx o_Excel, cNombreHoja & "4", "SUBSIDIO"
    Set o_Hoja = o_Libro.Worksheets.Add
    'DSCTO M.JUDICIAL
    SQL = "select anomes,moneda,emp as codigo,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS nombres,p.rub as rubro,p.cant as monto " & _
          "from pl_tareo p left join empleado e on (p.emp=e.codigo) where anomes='" & AnoMes & "' and " & _
          "(rub='703') and (p.tipo='1' or p.tipo='5')"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 1).Value = "RRHH"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 2).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 1).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(70)) & ":" & Trim(Chr(70))).NumberFormat = "#,##0.00"
        o_Excel.Worksheets(1).Range(Trim(Chr(67)) & ":" & Trim(Chr(67))).NumberFormat = "@"
        o_Excel.Worksheets(1).Range(Trim(Chr(69)) & ":" & Trim(Chr(69))).NumberFormat = "@"
        Fil = 4: Col = 1: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("rubro"))
            o_Hoja.Cells(Fil, Col + 5).Value = CDbl(myrs.Fields("monto"))
            Sum1 = Sum1 + CDbl(myrs.Fields("monto"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 5).Value = Sum1
    End If
    SQL = "select a.anomes,a.voucher,a.auxiliar,a.codaux as codigo,a.moneda,a.coddoc as tipodoc,a.numdoc,a.concepto," & _
          "a.Cargos,a.abonos from cnmovi as a where a.anomes='" & AnoMes & "' and CODLIB='02' AND (cuenta='46902')"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 9).Value = "CONTABILIDAD"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 10).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 9).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(81)) & ":" & Trim(Chr(81))).NumberFormat = "#,##0.00"
        o_Excel.Worksheets(1).Range(Trim(Chr(82)) & ":" & Trim(Chr(82))).NumberFormat = "#,##0.00"
        For columna = 0 To 8
            o_Excel.Worksheets(1).Range(Trim(Chr(73 + columna)) & ":" & Trim(Chr(73 + columna))).NumberFormat = "@"
        Next
        Fil = 4: Col = 9: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("voucher"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("auxiliar"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 5).Value = Trim(myrs.Fields("tipodoc"))
            o_Hoja.Cells(Fil, Col + 6).Value = Trim(myrs.Fields("numdoc"))
            o_Hoja.Cells(Fil, Col + 7).Value = Trim(myrs.Fields("concepto"))
            o_Hoja.Cells(Fil, Col + 8).Value = CDbl(myrs.Fields("cargos"))
            o_Hoja.Cells(Fil, Col + 9).Value = CDbl(myrs.Fields("abonos"))
            Sum1 = Sum1 + CDbl(myrs.Fields("abonos"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 9).Value = Sum1
    End If
    FormatoEx o_Excel, cNombreHoja & "5", "DSCTO. M.JUDICIAL"
    Set o_Hoja = o_Libro.Worksheets.Add
    'PRESTAMOS Y ADELANTOS
    SQL = "select anomes,moneda,emp as codigo,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS nombres, " & _
          "p.rub as rubro,sum(p.cant) as monto from pl_tareo p left join empleado e on (p.emp=e.codigo) " & _
          "where anomes='" & AnoMes & "' and (rub='701' OR RUB='709') and (p.tipo='1' or p.tipo='5') group by emp,rub order by nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 1).Value = "RRHH-PLANILLA"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 2).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 1).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        For columna = 0 To 5
            o_Excel.Worksheets(1).Range(Trim(Chr(65 + columna)) & ":" & Trim(Chr(65 + columna))).NumberFormat = "@"
        Next
        o_Excel.Worksheets(1).Range(Trim(Chr(70)) & ":" & Trim(Chr(70))).NumberFormat = "#,##0.00"
        Fil = 4: Col = 1: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("rubro"))
            o_Hoja.Cells(Fil, Col + 5).Value = CDbl(myrs.Fields("monto"))
            Sum1 = Sum1 + CDbl(myrs.Fields("monto"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 5).Value = Sum1
    End If
    SQL = "select a.anomes,a.voucher,a.auxiliar,a.codaux as codigo,c.descrip as nombres,a.moneda, " & _
          "a.coddoc as tipodoc,a.numdoc,a.concepto,sum(a.abonos) as abonos " & _
          "from cnmovi as a left join cnauxil c on (a.auxiliar=c.auxiliar and a.codaux=c.codigo) where " & _
          "a.anomes='" & AnoMes & "' and (cuenta='14102' or cuenta='14101' or cuenta='38501' or cuenta='38502') " & _
          "and codlib='02' group by a.codaux,left(cuenta,2) order by nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 9).Value = "CONTABILIDAD"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 10).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 9).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(82)) & ":" & Trim(Chr(82))).NumberFormat = "#,##0.00"
        For columna = 0 To 8
            o_Excel.Worksheets(1).Range(Trim(Chr(73 + columna)) & ":" & Trim(Chr(73 + columna))).NumberFormat = "@"
        Next
        Fil = 4: Col = 9: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("voucher"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("auxiliar"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 5).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 6).Value = Trim(myrs.Fields("tipodoc"))
            o_Hoja.Cells(Fil, Col + 7).Value = Trim(myrs.Fields("numdoc"))
            o_Hoja.Cells(Fil, Col + 8).Value = Trim(myrs.Fields("concepto"))
            o_Hoja.Cells(Fil, Col + 9).Value = CDbl(myrs.Fields("abonos"))
            Sum1 = Sum1 + CDbl(myrs.Fields("abonos"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 9).Value = Sum1
    End If
    FormatoEx o_Excel, cNombreHoja & "6", "PRESTAMOS Y ADELANTOS"
    Set o_Hoja = o_Libro.Worksheets.Add
    'CESES
    dia = IIf(val(Right(AnoMes, 2)) < Month(Date), Day(DateSerial(Left(AnoMes, 4), val(Right(AnoMes, 2)) + 1, 0)), Day(Date))
    SQL = "SELECT e.codigo,concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres," & _
          "IFNULL(date_format(fec_cese,'%d/%m/%Y'),'') as fec_cese,(select descrip from pl_tipocese p " & _
          "where p.codigo=c.codtipocese) as motivo FROM empleado e inner join contrato c on (e.codigo=c.codemp) " & _
          "where c.codigo = (SELECT max(codigo) as cod FROM CONTRATO t where t.codemp=c.codemp group by codemp) " & _
          "and (situacion = 0 and fec_cese BETWEEN '" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01' and " & _
          "'" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & dia & "') group by e.codigo,e.fec_cese " & _
          "order by nombres,fec_cese"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 1).Value = "RRHH"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 2).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 1).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        For columna = 0 To 4
            o_Excel.Worksheets(1).Range(Trim(Chr(65 + columna)) & ":" & Trim(Chr(65 + columna))).NumberFormat = "@"
        Next
        Fil = 4: Col = 1: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("fec_cese"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("motivo"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
    End If
    SQL = "select a.anomes,a.voucher,a.auxiliar,a.codaux as codigo,c.descrip as nombres,a.moneda, " & _
          "a.coddoc as tipodoc,a.numdoc,a.concepto,a.abonos " & _
          "from cnmovi as a left join cnauxil c on (a.auxiliar=c.auxiliar and a.codaux=c.codigo) where " & _
          "a.anomes='" & AnoMes & "' and cuenta='46903' " & _
          "and codlib='05' order by nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 7).Value = "CONTABILIDAD"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 8).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 7).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(80)) & ":" & Trim(Chr(80))).NumberFormat = "#,##0.00"
        For columna = 0 To 8
            o_Excel.Worksheets(1).Range(Trim(Chr(72 + columna)) & ":" & Trim(Chr(72 + columna))).NumberFormat = "@"
        Next
        Fil = 4: Col = 7: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("voucher"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("auxiliar"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 5).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 6).Value = Trim(myrs.Fields("tipodoc"))
            o_Hoja.Cells(Fil, Col + 7).Value = Trim(myrs.Fields("numdoc"))
            o_Hoja.Cells(Fil, Col + 8).Value = Trim(myrs.Fields("concepto"))
            o_Hoja.Cells(Fil, Col + 9).Value = CDbl(myrs.Fields("abonos"))
            Sum1 = Sum1 + CDbl(myrs.Fields("abonos"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 9).Value = Sum1
    End If
    FormatoEx o_Excel, cNombreHoja & "7", "CESES"
    Set o_Hoja = o_Libro.Worksheets.Add
    'VACACIONES
    SQL = "select anomes,moneda,emp as codigo,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS nombres, " & _
          "p.rub as rubro,sum(p.cant) as monto from pl_tareo p left join empleado e on (p.emp=e.codigo) " & _
          "where anomes='" & AnoMes & "' and (rub='117' OR RUB='118') AND P.TIPO='1' group by emp,rub order by nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 1).Value = "RRHH-PLANILLA"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 2).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 1).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        For columna = 0 To 4
            o_Excel.Worksheets(1).Range(Trim(Chr(65 + columna)) & ":" & Trim(Chr(65 + columna))).NumberFormat = "@"
        Next
        o_Excel.Worksheets(1).Range(Trim(Chr(70)) & ":" & Trim(Chr(70))).NumberFormat = "#,##0.00"
        Fil = 4: Col = 1: Sum1 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("rubro"))
            o_Hoja.Cells(Fil, Col + 5).Value = CDbl(myrs.Fields("monto"))
            Sum1 = Sum1 + val(myrs.Fields("monto"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 5).Value = Sum1
    End If
    SQL = "select a.anomes,a.voucher,a.auxiliar,a.codaux as codigo,b.descrip as nombres,a.moneda,a.coddoc as tipodoc, " & _
          "a.numdoc,a.concepto,a.cargos from cnmovi as a left join cnauxil as b on (a.auxiliar=b.auxiliar and a.codaux=b.codigo) " & _
          "where a.anomes='" & AnoMes & "' and cuenta='41201' and codlib='02' and coddoc='7' and a.cargos<>0 order by nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 9).Value = "CONTABILIDAD"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 10).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 10).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(83)) & ":" & Trim(Chr(83))).NumberFormat = "#,##0.00"
        For columna = 0 To 8
            o_Excel.Worksheets(1).Range(Trim(Chr(73 + columna)) & ":" & Trim(Chr(73 + columna))).NumberFormat = "@"
        Next
        Fil = 4: Col = 9: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("voucher"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("auxiliar"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 5).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 6).Value = Trim(myrs.Fields("tipodoc"))
            o_Hoja.Cells(Fil, Col + 7).Value = Trim(myrs.Fields("numdoc"))
            o_Hoja.Cells(Fil, Col + 8).Value = Trim(myrs.Fields("concepto"))
            o_Hoja.Cells(Fil, Col + 9).Value = CDbl(myrs.Fields("cargos"))
            Sum1 = Sum1 + CDbl(myrs.Fields("cargos"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 9).Value = Sum1
    End If
    FormatoEx o_Excel, cNombreHoja & "8", "VACACIONES"
    Set o_Hoja = o_Libro.Worksheets.Add
    'BONOS

    SQL = "call RH_SP_BonosCampo('EXPORTXLS','" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "','','','','','');"
    Dim RsB As ADODB.Recordset

    Set RsB = ADO_LlenaRs(SQL)
    o_Hoja.Cells(1, 1).Value = "OPERACIONES"
    If Not (RsB.EOF And RsB.BOF) Then
        For columna = 0 To RsB.Fields.Count - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 1).Value = Trim(RsB.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(68)) & ":" & Trim(Chr(68))).NumberFormat = "#,##0.00"
        o_Excel.Worksheets(1).Range(Trim(Chr(65)) & ":" & Trim(Chr(65))).NumberFormat = "@"
        Dim xDia As Integer, xSumDias As Integer
        Dim emp As String, FlgSal As Boolean, FlgIng As Boolean
        Dim Tipo As String, dias As Integer, xdiaini As Integer, xdiafin As Integer, I As Integer
        Fil = 4: Col = 1: Sum1 = 0: Sum2 = 0
        emp = Trim(RsB.Fields("codigo"))
        Do While Not (RsB.EOF)
            o_Hoja.Cells(Fil, Col).Value = Trim(RsB.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(RsB.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(RsB.Fields("dias"))
            o_Hoja.Cells(Fil, Col + 3).Value = CDbl(RsB.Fields("monto_bono"))
            Fil = Fil + 1
            
            RsB.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 2).Value = "=SUM(C4:C" & Fil - 1 & ")"
        o_Hoja.Cells(Fil, Col + 3).Value = "=SUM(D4:D" & Fil - 1 & ")"
     Else
         o_Hoja.Cells(1, 2).Value = "SIN MOVIMIENTO"
     End If
    Set RsB = Nothing
    SQL = "select e.codigo,CONCAT_WS(' ',APEPAT,APEMAT,NOMBRE1,NOMBRE2) AS nombres,convert((p.cant/c.monto_bono),unsigned) as dias, " & _
          "p.CANT as monto from (pl_tareo p left join empleado e on (p.emp=e.codigo)) LEFT JOIN CONTRATO C on (e.codigo=c.codemp) " & _
          "where p.tipo = 1 and p.anomes = '" & AnoMes & "' and rub = '1001' and c.CODIGO = (SELECT MAX(CODIGO) FROM CONTRATO T " & _
          "WHERE T.CODEMP=E.CODIGO GROUP BY T.CODEMP) ORDER BY nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 7).Value = "PLANILLA"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 8).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 7).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(74)) & ":" & Trim(Chr(74))).NumberFormat = "#,##0.00"
        o_Excel.Worksheets(1).Range(Trim(Chr(71)) & ":" & Trim(Chr(71))).NumberFormat = "@"
        Fil = 4: Col = 7: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 2).Value = val(CEN(myrs.Fields("dias")))
            o_Hoja.Cells(Fil, Col + 3).Value = CDbl(CEN(myrs.Fields("monto")))
            Sum1 = Sum1 + val(CE(myrs.Fields("dias")))
            Sum2 = Sum2 + CDbl(CE(myrs.Fields("monto")))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 2).Value = Sum1
        o_Hoja.Cells(Fil, Col + 3).Value = Sum2
    End If
    FormatoEx o_Excel, cNombreHoja & "9", "BONOS"
    Set o_Hoja = o_Libro.Worksheets.Add
    'REFRIGERIOS
    dia = IIf(val(Right(AnoMes, 2)) < Month(Date), Day(DateSerial(Left(AnoMes, 4), val(Right(AnoMes, 2)) + 1, 0)), Day(Date))
    SQL = "select codemp AS codigo,(Select CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) from empleado " & _
          "where codigo=codemp) as nombres,codauxa as codaux,facturaa as factura,montoa as monto,count(*) as cantidad,montoa*count(*) as Valor " & _
          "from `rh_refriempleados` where fecha>='" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01' and " & _
          "fecha<='" & Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/" & dia & "' " & _
          "group by codemp,codauxa,facturaa,montoa order by codauxa,facturaa,nombres"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 1).Value = "RRHH"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 2).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 1).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        For columna = 0 To 4
            o_Excel.Worksheets(1).Range(Trim(Chr(65 + columna)) & ":" & Trim(Chr(65 + columna))).NumberFormat = "@"
        Next
        o_Excel.Worksheets(1).Range(Trim(Chr(69)) & ":" & Trim(Chr(69))).NumberFormat = "#,##0.00"
        o_Excel.Worksheets(1).Range(Trim(Chr(71)) & ":" & Trim(Chr(71))).NumberFormat = "#,##0.00"
        Fil = 4: Col = 1: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("codaux"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("factura"))
            o_Hoja.Cells(Fil, Col + 4).Value = CDbl(myrs.Fields("monto"))
            o_Hoja.Cells(Fil, Col + 5).Value = CDbl(myrs.Fields("cantidad"))
            o_Hoja.Cells(Fil, Col + 6).Value = CDbl(myrs.Fields("valor"))
            Sum1 = Sum1 + val(myrs.Fields("valor"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 5).Value = "Fact.": o_Hoja.Cells(Fil, Col + 6).Value = Sum1
        o_Hoja.Cells(Fil + 1, Col + 5).Value = "Igv": o_Hoja.Cells(Fil + 1, Col + 6).Value = Round(Sum1 * 0.18, 2)
        o_Hoja.Cells(Fil + 2, Col + 5).Value = "Total": o_Hoja.Cells(Fil + 2, Col + 6).Value = Sum1 + Round(Sum1 * 0.18, 2)
    End If
    SQL = "select a.anomes,a.voucher,a.auxiliar,a.codaux,a.moneda,a.coddoc as tipodoc,a.numdoc,a.concepto,a.cargos " & _
          "from cnmovi as a left join cnauxil as b on (a.auxiliar=b.auxiliar and a.codaux=b.codigo) " & _
          "where a.anomes='" & AnoMes & "' and b.suspeN='REFRI' and a.codlib='01' order by numdoc,concepto"
    Set myrs = oConexion.EjecutaSelectRS(SQL)
    o_Hoja.Cells(1, 10).Value = "CONTABILIDAD"
    If myrs.RecordCount <= 0 Then
        o_Hoja.Cells(1, 11).Value = "SIN MOVIMIENTO"
    Else
        For columna = 0 To myrs.FieldCount - 1
            If myrs.Fields(columna).name <> "" Then
                o_Hoja.Cells(3, columna + 10).Value = Trim(myrs.Fields(columna).name)
            End If
        Next columna
        o_Excel.Worksheets(1).Range(Trim(Chr(83)) & ":" & Trim(Chr(83))).NumberFormat = "#,##0.00"
        For columna = 0 To 8
            o_Excel.Worksheets(1).Range(Trim(Chr(74 + columna)) & ":" & Trim(Chr(74 + columna))).NumberFormat = "@"
        Next
        Fil = 4: Col = 10: Sum1 = 0: Sum2 = 0
        Do While Not myrs.EOF()
            o_Hoja.Cells(Fil, Col).Value = Trim(myrs.Fields("anomes"))
            o_Hoja.Cells(Fil, Col + 1).Value = Trim(myrs.Fields("voucher"))
            o_Hoja.Cells(Fil, Col + 2).Value = Trim(myrs.Fields("auxiliar"))
            o_Hoja.Cells(Fil, Col + 3).Value = Trim(myrs.Fields("codigo"))
            o_Hoja.Cells(Fil, Col + 4).Value = Trim(myrs.Fields("nombres"))
            o_Hoja.Cells(Fil, Col + 5).Value = Trim(myrs.Fields("moneda"))
            o_Hoja.Cells(Fil, Col + 6).Value = Trim(myrs.Fields("tipodoc"))
            o_Hoja.Cells(Fil, Col + 7).Value = Trim(myrs.Fields("numdoc"))
            o_Hoja.Cells(Fil, Col + 8).Value = Trim(myrs.Fields("concepto"))
            o_Hoja.Cells(Fil, Col + 9).Value = CDbl(myrs.Fields("cargos"))
            Sum1 = Sum1 + CDbl(myrs.Fields("cargos"))
            Fil = Fil + 1
            myrs.MoveNext
        Loop
        o_Hoja.Cells(Fil, Col + 9).Value = Sum1
    End If
    FormatoEx o_Excel, cNombreHoja & "10", "REFRIGERIOS"
    o_Excel.Visible = True
    o_Hoja.SaveAs Rep_Documents & "\Verificacion Planilla_RRHH_Contabilidad.xls"
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
    Set myrs = Nothing
    Screen.MousePointer = vbDefault
Exit Sub
ErrorAbrir:
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
    Set myrs = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Ocurrió un error al momento de generar el archivo de Excel, " & vbNewLine & _
           "consulte con el administrador del sistema", vbExclamation + vbOKOnly, "NOVADMIN"
End Sub
Sub FormatoEx(o_Excel As Excel.Application, Nom As String, Nom1 As String)
    o_Excel.Cells.Select
    With o_Excel.Selection.Font
        .name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    o_Excel.Sheets(Nom).Select
    o_Excel.Sheets(Nom).name = Nom1
    o_Excel.Cells.EntireColumn.AutoFit
    o_Excel.Rows("1:3").Select
    o_Excel.Selection.Font.Bold = True
End Sub

Public Sub TareasProgramadas(CodEmp As String, CodTarea As String, Codprioridad As Integer, dia As String, Numdia As String, Mes As String)
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
    Screen.MousePointer = vbHourglass
    DoEvents
                    
    Set oRs = ADO_LlenaRs("call RH_SP_Tareas_Empleado('REPORTE_SIMPLE','" & strAnoSistema & "','" & "" & "','','','','','','');")
    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No Existen Registros", vbOKOnly + vbInformation, gsNomSW
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, False
    End If
    Set oRs = Nothing
    
    Screen.MousePointer = vbNormal
End Sub


Public Sub sp_Rep_CtasPagarPreview(TxtOrdenPago As String, TFolio As String)
    Dim SQL As String, str As String
    Dim myrs As MYSQL_RS
    Dim dia As Integer
    Dim oRs  As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    Dim Compra, Venta, Oanda
    Dim flagBD As String
    Dim FlagDH As String
    Dim fecCorte As String
    Dim FlagSoles As String
    
    SQL = "Call  cn_buscar_liquidacion_rep(" & TxtOrdenPago & ",'" & TFolio & "');"
    Set oRs = ADO_LlenaRs(SQL)
    If Not oRs Is Nothing Then
       If oRs.RecordCount = 0 Then
          MsgBox "No existen datos", vbOKOnly + vbInformation, gsNomSW
          Exit Sub
       End If
       
       EjecutaReporte App.Path & "\Reportes\CuentasxPagar\" & m_Reporte, myrs, oRs, False
    End If
End Sub

Public Sub sp_Rep_CtasCobrarAnexo(TC_Compra As Double, TC_Venta As Double)
    Dim SQL As String, str As String
    Dim myrs As MYSQL_RS
    Dim dia As Integer
    Dim oRs  As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    Dim Compra, Venta, Oanda
    Dim flagBD As String
    Dim FlagDH As String
    Dim fecCorte As String
    Dim FlagSoles As String
    
    
    Compra = DameParametrosReporte("C")
    Venta = DameParametrosReporte("V")
    Oanda = DameParametrosReporte("O")
    fecCorte = CE(InputBox("Ingrese la fehca con la cual va a comparar?", "", "30/11/2019"))
    
    
    SQL = "Call  CONT_Rep_TemplateAging_FacturasSOX_ADM('" & Compra & "','" & Venta & "','" & Oanda & "','" & strMesSistema & "','" & strAnoSistema & "','1','" & fecCorte & "','L','N','N','');"
    Set oRs = ADO_LlenaRs(SQL)
    If Not oRs Is Nothing Then
            If oRs.RecordCount = 0 Then
                MsgBox "No existen datos", vbOKOnly + vbInformation, gsNomSW
                Exit Sub
            End If
     
            EjecutaReporte App.Path & "\Reportes\CuentasxCobrar\" & m_Reporte, myrs, oRs, False
    End If
    
'    SQL = "delete from TempTempplateD where usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute SQL
'    SQL = "insert into TempTempplateD(usuario,impme,impmn,moneda,campo1,campo2,campo3,coddep,descuenta,division,coddoc,numdoc,cuenta,aux,codaux,desaux,mone,monn) " & _
'          " SELECT '" & strUsuarioId & "',0 as IMPME,0 AS IMPMN,w.moneda,t.string1 as dia,t.string2 as fecha,t.string3 as mesdia,f.abrevia,w.descrip as descuenta, d.descrip as division , a.coddoc," & _
'          " a.numdoc,a.cuenta,a.auxiliar,a.codaux,b.descrip as desaux,sum(a.cargod-a.abonod) as mone," & _
'          " sum(a.cargos-a.abonos)  as monn  " & _
'          " FROM ((((cnmovi as a left join cnauxil as b on a.auxiliar=b.auxiliar and a.codaux=b.codigo) " & _
'          " left join cnmdepar as d on a.coddep=d.coddep) left join cnmayor as w on a.cuenta=w.cuenta) " & _
'          " left join cndocum as f  on a.coddoc=f.coddoc) LEFT JOIN CNTEMP AS T ON (T.indice = '0')  " & _
'          " WHERE b.afiliada='N' and (left(a.cuenta,5)='121301' or left(a.cuenta,5)='121302' or left(a.cuenta,5)='12301' or left(a.cuenta,5)='12302' OR left(a.cuenta,6)='121101') AND A.AUXILIAR='2' " & _
'          " and a.anomes>='" & strAnoSistema & "00' and a.anomes<='" & strAnoSistema & strMesSistema & "' " & _
'          " group by d.descrip,a.cuenta, a.auxiliar, a.codaux, a.coddoc,  a.numdoc, " & _
'          " b.descrip,w.descrip ,f.abrevia,w.moneda" & _
'          " Having round(sum(a.cargod - a.abonod),2)<>0.00 and round(sum(a.cargos - a.abonos),2) <> 0.00 "
'    oConexionMYSQL.Execute (SQL)
'    SQL = "update TempTempplateD set impme=round(monn/" & TC_Compra & ",2) where moneda='N' and monn>0 and usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute (SQL)
'    SQL = "update TempTempplateD set impme=round(monn/" & TC_Venta & ",2) where moneda='N' and monn<0 and usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute (SQL)
'    SQL = "update TempTempplateD set impme=round(mone,2) where moneda='N' and monn=0 and usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute (SQL)
'    SQL = "update TempTempplateD set impmn=round(monn,2) where moneda='N' and usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute (SQL)
'    SQL = "update TempTempplateD set impmn=round(mone*" & TC_Compra & ",2) where moneda<>'N' and mone>0 and usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute (SQL)
'    SQL = "update TempTempplateD set impmn=round(mone*" & TC_Venta & ",2) where moneda<>'N' and mone<0 and usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute (SQL)
'    SQL = "update TempTempplateD set impmn=round(monn,2) where moneda<>'N' and mone=0 and usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute (SQL)
'    SQL = "update TempTempplateD set impme=round(mone,2) where moneda<>'N' and usuario='" & strUsuarioId & "'"
'    oConexionMYSQL.Execute (SQL)
'    SQL = " SELECT campo1 as dia,campo2 as fecha,campo3 as fecrecepcion,coddoc,numdoc " & _
'          " FROM TempTempplateD WHERE usuario='" & strUsuarioId & "' group by division,cuenta,aux,codaux,coddoc,numdoc, " & _
'          " desaux,descuenta,coddep"
'    Set Rs = oConexion.EjecutaSelectRS(SQL)
'    If Rs.EOF And Rs.BOF Then
'        MsgBox "No tiene movimiento", vbInformation + vbOKOnly, "Movimiento"
'        Exit Sub
'    End If
'    Dim f1 As String, f2 As String, dias As String
'    Do While Not Rs.EOF
'        If Rs.Fields("coddoc") = "L" Then
'            f1 = ObtenerFechaRecepcion(Rs.Fields("numdoc"), Rs.Fields("coddoc"))
'        Else
'            f1 = ObtenerFechaRecepcion(Right(Mid(Rs.Fields("numdoc"), 1, InStr(1, Rs.Fields("numdoc"), "-") - 1), 3) & "-" & CDbl(IIf(InStr(1, Rs.Fields("numdoc"), "-") > 0, Mid(Rs.Fields("numdoc"), InStr(1, Rs.Fields("numdoc"), "-") + 1, Len(Rs.Fields("numdoc")) - InStr(1, Rs.Fields("numdoc"), "-")), "")), Rs.Fields("coddoc"))
'        End If
'        f2 = Mid(f1, InStr(1, f1, ",") + 1, Len(f1) - InStr(1, f1, ","))
'        f1 = Mid(f1, 1, 10)
'        dia = Date - CDate(IIf(f1 = "", Date, f1))
'        If dia = 0 Then
'            dias = "A - Current 0 DAYS"
'        ElseIf dia >= 1 And dia <= 30 Then
'            dias = "B - Past Due 1 - 30 DAYS"
'        ElseIf dia > 30 And dia <= 60 Then
'            dias = "C - Past Due 31 - 60 DAYS"
'        ElseIf dia > 60 And dia <= 90 Then
'            dias = "D - Past Due 61 - 90 DAYS"
'        ElseIf dia > 90 Then
'            dias = "E - Past Due > 90 DAYS"
'        End If
'        If Rs.Fields("coddoc") = "P" Then dias = "Unbilled"
'        SQL = "update TempTempplateD set campo2 = '" & f1 & "',campo3 = '" & f2 & "',campo1 = '" & dias & "' where " & _
'              "numdoc = '" & Trim(Rs.Fields("numdoc")) & "' and coddoc = '" & Trim(Rs.Fields("coddoc")) & "'"
'        oConexionMYSQL.Execute SQL
'        Rs.MoveNext
'    Loop
'    CantCta = 0
'    Dim RQ As MYSQL_RS
'    Dim o_Excel As Object
'    Dim o_Libro As Object
'    Set o_Excel = New Excel.Application
'    Set o_Libro = o_Excel.Workbooks.Add
'    SQL = "SELECT DIVISION,SUM(i1) '0 DOL',SUM(i2) '0 SOL',SUM(i3) '1 - 30 DOL',SUM(i4) '1 - 30 SOL', " & _
'          "Sum(i5) '31 - 60 DOL',SUM(i6) '31 - 60 SOL',SUM(i7) '61 - 90 DOL',SUM(i8) '61 - 90 SOL', " & _
'          "Sum(i9) '> 90 DOL',SUM(i10) '> 90 SOL',SUM(i11) 'Unbilled DOL',SUM(i12) 'Unbilled SOL' " & _
'          "FROM (SELECT DISTINCT division,if(instr(CAMPO1,'A - Current 0 DAYS')>0,SUM(impme),0) I1, " & _
'          "if(instr(CAMPO1,'A - Current 0 DAYS')>0,SUM(impmn),0) I2,if(instr(CAMPO1,'1 - 30')>0,SUM(impme),0) I3, " & _
'          "if(instr(CAMPO1,'1 - 30')>0,SUM(impmn),0) I4,if(instr(CAMPO1,'31 - 60')>0,SUM(impme),0) I5, " & _
'          "if(instr(CAMPO1,'31 - 60')>0,SUM(impmn),0) I6,if(instr(CAMPO1,'61 - 90')>0,SUM(impme),0) I7, " & _
'          "if(instr(CAMPO1,'61 - 90')>0,SUM(impmn),0) I8,if(instr(CAMPO1,'> 90')>0,SUM(impme),0) I9, " & _
'          "if(instr(CAMPO1,'> 90')>0,SUM(impmn),0) I10,if(instr(CAMPO1,'Unbilled')>0,SUM(impme),0) I11, " & _
'          "if(instr(CAMPO1,'Unbilled')>0,SUM(impmn),0) I12 From TempTempplateD WHERE usuario='" & strUsuarioId & "' " & _
'          "group by division,CAMPO1) AS K group by division"
'    Set RQ = oConexion.EjecutaSelectRS(SQL)
'    If RQ.EOF And RQ.BOF Then
'        MsgBox "No tiene movimiento", vbInformation + vbOKOnly, "Movimiento"
'        Exit Sub
'    End If
'    Call Exportar_Excel_CtasxCobrar(RQ, o_Excel, o_Libro)
'    SQL = "SELECT DISTINCT DIVISION FROM TempTempplateD WHERE usuario='" & strUsuarioId & "'"
'    Set RQ = oConexion.EjecutaSelectRS(SQL)
'    Do While Not RQ.EOF()
'        SQL = "SELECT division,campo2 as 'FECHA EMISION',campo3 as 'FECHA RECEPCION',coddep AS DOC,numdoc AS 'NUM.DOC'," & _
'              "desaux AS CLIENTE,impme DOLARES,impmn AS SOLES,MONEDA " & _
'              "FROM TempTempplateD WHERE usuario='" & strUsuarioId & "' and division = '" & Trim(RQ.Fields("division")) & "' " & _
'              "group by division,cuenta,aux,codaux,coddoc,numdoc,desaux,coddep " & _
'              "order by division,concat(right(campo2,4),substring(campo2,3,4),left(campo2,2)),numdoc"
'        Set Rs = oConexion.EjecutaSelectRS(SQL)
'        Call Exportar_Excel_CtasxCobrar(Rs, o_Excel, o_Libro)
'        RQ.MoveNext
'    Loop
'    o_Excel.Visible = True
'    o_Libro.SaveAs Rep_Documents & "\" & "CtasxCobrar " & strMesSistema & strAnoSistema
'    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
'    If Not o_Libro Is Nothing Then Set o_Libro = Nothing

End Sub

Private Function ObtenerFechaRecepcion(Numdoc As String, TipoDoc As String) As String
Dim SQL As String, RQ As MYSQL_RS
Dim cad As String
If TipoDoc = "L" Then
    SQL = "select fecgiro,fecvcto from letra where numero = '" & Right("000000" & Numdoc, 6) & "'"
Else
    SQL = "select fec_emision,fec_vcto " & _
          "from ((documento_contables dc left join movi_documento m on (dc.identificador=m.identificador)) " & _
          "left join amarre_documento a on (dc.identificador=a.identificador)) " & _
          "where CONCAT(right(dc.serie,3),'-',convert(dc.correl,unsigned)) = '" & Numdoc & "' and " & _
          "a.cod_tipo_doc = '" & TipoDoc & "' and dc.auxiliar = 2"
End If
Set RQ = oConexion.EjecutaSelectRS(SQL)
If Not RQ.EOF() Then
    If TipoDoc = "L" Then
        cad = Trim(RQ.Fields("fecgiro")) & "," & Trim(RQ.Fields("fecvcto"))
    Else
        cad = Right(RQ.Fields("fec_emision"), 2) & Mid(RQ.Fields("fec_emision"), 5, 4) & Left(RQ.Fields("fec_emision"), 4)
        cad = cad & "," & Right(RQ.Fields("fec_vcto"), 2) & Mid(RQ.Fields("fec_vcto"), 5, 4) & Left(RQ.Fields("fec_vcto"), 4)
    End If
    ObtenerFechaRecepcion = cad
End If
Set RQ = Nothing
End Function

Public Function FechaUSA(cFecha As String) As String
    FechaUSA = Mid(cFecha, 4, 2) & Left(cFecha, 2) & Right(cFecha, 4)
End Function

Public Sub PersonalCS(cFechaIni As String, cFechaFin As String)
    Dim NomArchivo As String
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
    Dim cRetorno As String
    Screen.MousePointer = vbHourglass
    DoEvents
    Dim SQL As String
    Dim msj As String
    m_Cab1 = cFechaIni
    m_Cab2 = cFechaFin
    
    Dim Email  As String
    Email = "Mary.Horst@nov.com"
    
    cRetorno = CE(RetornaValor("call RH_SP_ReporteTrabajador('CORREO_ENVIADO','','','" & cFechaIni & "','')"))
    
    If cRetorno = "" Then
        m_Reporte = "Rep_PersonalCS_Det.rpt"
        Set oRs = ADO_LlenaRs("call RH_SP_ReporteTrabajador('PERSONAL_SOLIDOS_DET','','','" & cFechaIni & "','" & cFechaFin & "')")
        If Not oRs Is Nothing Then
            If oRs.RecordCount = 0 Then
                MsgBox "No existen datos", vbOKOnly + vbInformation, gsNomSW
                Exit Sub
            End If
    
            EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, False
        End If
        
        '---------------------------------------------
        m_Reporte = "Rep_PersonalCS.rpt"
        Set oRs = ADO_LlenaRs("call RH_SP_ReporteTrabajador('PERSONAL_SOLIDOS','','','" & cFechaIni & "','" & cFechaFin & "')")
        If Not oRs Is Nothing Then
            If oRs.RecordCount = 0 Then
                MsgBox "No Existen Registros", vbOKOnly + vbInformation, gsNomSW
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
            
            msj = "Reporte Semanal del personal de Control de solidos al " & cFechaFin & "." & _
                  vbNewLine & vbNewLine & "Saludos," & _
                  vbNewLine & vbNewLine & "Augusto Mendoza"
            msj = ""
            
            NomArchivo = "SOLID_CONTROL_DETAIL_" & FechaUSA(cFechaIni) & ".PDF"
            
            If ExisteArchivo(Rep_Documents & "\" & NomArchivo, False) = True Then
                Call EliminaArchivo(Rep_Documents & "\" & NomArchivo, True)
            End If
            
            EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, True, True, Rep_Documents & "\" & NomArchivo
            
            
'            If strUsuarioId = "CARLAC" Then
                If Mensajes("Desea enviarlo por correo electronico", vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    MensajeNuevoOutlook Email, 2, msj, "SC WEEKLY HEADCOUNT 131", "maria.ocana@nov.com;augusto.mendoza@nov.com", NomArchivo
    
                    If Mensajes("El correo fue enviado a sus destinatarios", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                    
                        SQL = "insert into rh_envio_csemp( anio,mes,fecha,hora) " & _
                              "values('" & CE(Year(Date)) & "','" & Right("00" & CE(Month(Date)), 2) & "','" & CE(Date) & "','" & CE(Time) & "')"
                          
                        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                    End If
                End If
'            End If
        End If
        Set oRs = Nothing
        
        Screen.MousePointer = vbNormal
    Else
        Screen.MousePointer = vbNormal
    End If
End Sub


Public Sub PersonalReductions(cFechaIni As String, cFechaFin As String)
    

    Dim NomArchivo As String
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
    Dim cRetorno As String
    Screen.MousePointer = vbHourglass
    DoEvents
    Dim SQL As String
    Dim msj As String
    m_Cab1 = cFechaIni
    m_Cab2 = cFechaFin
    
    
    Dim Email  As String
    Email = "Mary.Horst@nov.com"
    
    cRetorno = CE(RetornaValor("call RH_SP_EmpleadosxMes('REPORTE_REDUCTIONS','','','','" & cFechaIni & "','" & cFechaFin & "')"))
    
    If cRetorno = "" Then
        m_Reporte = "Rep_PersonalCS_Det.rpt"
        Set oRs = ADO_LlenaRs("call RH_SP_ReporteTrabajador('PERSONAL_SOLIDOS_DET','','','" & cFechaIni & "','" & cFechaFin & "')")
        If Not oRs Is Nothing Then
            If oRs.RecordCount = 0 Then
                MsgBox "No existen datos", vbOKOnly + vbInformation, gsNomSW
                Exit Sub
            End If
    
            EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, False
        End If
        
        '---------------------------------------------
        m_Reporte = "Rep_PersonalCS.rpt"
        Set oRs = ADO_LlenaRs("call RH_SP_ReporteTrabajador('PERSONAL_SOLIDOS','','','" & cFechaIni & "','" & cFechaFin & "')")
        If Not oRs Is Nothing Then
            If oRs.RecordCount = 0 Then
                MsgBox "No Existen Registros", vbOKOnly + vbInformation, gsNomSW
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
            
            msj = "Reporte Semanal del personal de Control de solidos al " & cFechaFin & "." & _
                  vbNewLine & vbNewLine & "Saludos," & _
                  vbNewLine & vbNewLine & "Augusto Mendoza"
            msj = ""
            
            NomArchivo = "SOLID_CONTROL_DETAIL_" & FechaUSA(cFechaIni) & ".PDF"
            
            If ExisteArchivo(Rep_Documents & "\" & NomArchivo, False) = True Then
                Call EliminaArchivo(Rep_Documents & "\" & NomArchivo, True)
            End If
            
            EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, True, True, Rep_Documents & "\" & NomArchivo
            
            
'            If strUsuarioId = "CARLAC" Then
                If Mensajes("Desea enviarlo por correo electronico", vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    MensajeNuevoOutlook Email, 2, msj, "SC WEEKLY HEADCOUNT 131", "maria.ocana@nov.com;augusto.mendoza@nov.com", NomArchivo
    
                    If Mensajes("El correo fue enviado a sus destinatarios", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                    
                        SQL = "insert into rh_envio_csemp( anio,mes,fecha,hora) " & _
                              "values('" & CE(Year(Date)) & "','" & Right("00" & CE(Month(Date)), 2) & "','" & CE(Date) & "','" & CE(Time) & "')"
                          
                        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                    End If
                End If
'            End If
        End If
        Set oRs = Nothing
        
        Screen.MousePointer = vbNormal
    Else
        Screen.MousePointer = vbNormal
    End If
End Sub

Public Sub sp_CantEmpDiv(opcRpt As String, opcDiv As String, mi As String, mf As String, f1 As String, f2 As String, f3 As String, f4 As String, f5 As String, f6 As String, f7 As String, f8 As String, f9 As String, f10 As String, f11 As String, f12 As String)
    Dim SQL As String
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
    Screen.MousePointer = vbHourglass
    DoEvents

    
    SQL = "Call RH_SP_CantEmpFec ( '" & opcRpt & "','" & opcDiv & "','','" & mi & "','" & mf & "','" & f1 & "','" & f2 & "','" & f3 & "','" & f4 & "','" & f5 & "','" & f6 & "','" & f7 & "','" & f8 & "','" & f9 & "','" & f10 & "','" & f11 & "','" & f12 & "');"
    
    Set oRs = ADO_LlenaRs(SQL)
    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No Existen Registros", vbOKOnly + vbInformation, gsNomSW
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, False
    End If
    Set oRs = Nothing
    
    Screen.MousePointer = vbNormal

End Sub

Public Sub sp_CantEmpDivDet(opcRpt As String, opcDiv As String, mi As String, mf As String, f1 As String, f2 As String, f3 As String, f4 As String, f5 As String, f6 As String, f7 As String, f8 As String, f9 As String, f10 As String, f11 As String, f12 As String)
    Dim SQL As String
    Dim oRs  As ADODB.Recordset
    Dim oRSMYSQL As MYSQL_RS
    Screen.MousePointer = vbHourglass
    DoEvents

    
    SQL = "Call RH_SP_CantEmpFecDet ( '" & opcRpt & "','" & opcDiv & "','','" & mi & "','" & mf & "','" & f1 & "','" & f2 & "','" & f3 & "','" & f4 & "','" & f5 & "','" & f6 & "','" & f7 & "','" & f8 & "','" & f9 & "','" & f10 & "','" & f11 & "','" & f12 & "');"
    
    Set oRs = ADO_LlenaRs(SQL)
    If Not oRs Is Nothing Then
        If oRs.RecordCount = 0 Then
            MsgBox "No Existen Registros", vbOKOnly + vbInformation, gsNomSW
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        EjecutaReporte App.Path & "\Reportes\RecursosHumanos\" & m_Reporte, oRSMYSQL, oRs, False
    End If
    Set oRs = Nothing
    
    Screen.MousePointer = vbNormal

End Sub



Function DameParametrosReporte(ByVal vTipo As String) As String
    Dim RQ As MYSQL_RS
    DameParametrosReporte = "0"
    Dim SQL As String
    
    If vTipo = "C" Then
     SQL = "select importecpa as Res from cncambiocierre where aniomes='" & strAnoSistema & strMesSistema & "'"
    End If
    
    If vTipo = "V" Then
     SQL = "select importevta as Res from cncambiocierre where aniomes='" & strAnoSistema & strMesSistema & "'"
    End If
    
    If vTipo = "O" Then
     SQL = " select importe as Res from cncambiomes_ where tipo='OAN' and aniomes='" & strAnoSistema & strMesSistema & "' "
    End If
    
    If vTipo = "F" Then
     SQL = " select concat(novperuhfm.ULTIMO_DIA_CC('" & strAnoSistema & strMesSistema & "'),'/','" & strMesSistema & "','/','" & strAnoSistema & "') as Res; "
    End If
    
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    
    If RQ.RecordCount > 0 Then
        DameParametrosReporte = RQ.Fields("Res")
    End If
    
    
    Set Rs = Nothing
End Function
