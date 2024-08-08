VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private m_Campo1 As String
'Private m_Campo2 As String
'Private m_Campo3  As String
'Private m_Caso As Integer
'
'Public Property Let pCampo1(Valor As String)
'  m_Campo1 = Valor
'End Property
'Public Property Let pCampo2(Valor As String)
'  m_Campo2 = Valor
'End Property
'Public Property Let pCampo3(Valor As Integer)
'  m_Campo3 = Valor
'End Property
'
'Public Property Let pCaso(Valor As Integer)
'  m_Caso = Valor
'End Property
'
'Public Property Get Caso() As Integer
'  Caso = m_Caso
'End Property

'Public Function CargarLista(Sql As String) As ADODB.Recordset
'
'    Dim STR As String, strCampos As String, str2 As String, i As Integer
'    Dim RS As ADODB.Recordset
'    Set RS = New ADODB.Recordset
'
'    Dim Sql As String
'    'MSHConsultas.Cols = 2
'    'ADOConexion.CursorLocation = adUseClient
'    RS.CursorLocation = adUseClient
'    RS.CursorType = adOpenStatic
'    RS.LockType = adLockBatchOptimistic
'
'    Select Case m_Caso
'        Case 1
'            Sql = "SELECT DISTINCT LEFT(CODDEP,2) as UNIDAD , DESCRIP as Descripción FROM CNMDEPAR where coddep<>'0000' "
'        Case 22
'
'            Sql = "SELECT DISTINCT rigHt(CODDEP,2) as UNIDAD , AREA as Descripción FROM CNMDEPAR where coddep<>'0000' "
'
'        Case 1
'            Sql = "SELECT DISTINCT CODLIB as Libro,DESCRIP as Descripción FROM CNLIBROS where codlib<>'00'"
'        Case 2
'            Sql = "SELECT DISTINCT Cuenta,DESCRIP  as Descripción FROM CNMAYOR WHERE TIPO='D' ORDER BY 1"
'        Case 3
'            Sql = "SELECT DISTINCT Codigo,DESCRIP  as Descripción FROM CNAUXIL WHERE AUXILIAR = '" & m_codAuxil & "' Order By 1"
'        Case 4
'            Sql = "SELECT DISTINCT CENCOS as CCosto,DESCRIP  as Descripción , RIGHT(CENCOS,3) AS CEN FROM CNMCOSTO WHERE CENCOS<>'0000' ORDER BY 1"
'        Case 5
'            Sql = "SELECT DISTINCT CODDEP as Division,DESCRIP+ ' - ' + AREA  as Descripción FROM CNMDEPAR  WHERE CODDEP<>'0000' ORDER BY 1"
'        Case 6
'            Sql = "SELECT DISTINCT Codigo,DESCRIP  as Descripción FROM CNAUXIL WHERE AUXILIAR = '" & CodAuxil & "' and codigo<>'00000000000' Order By 1"
'        Case 7
'            Sql = "SELECT DISTINCT CODDOC as Documento ,DESCRIP as Descripción ,Abrevia FROM CNDOCUM  ORDER BY 1"
'        Case 8
'            Sql = "SELECT DISTINCT a.CODAUX as Código, b.descrip  as Descripción FROM CNMOVI as a left join cnauxil as b on a.codaux=b.codigo and a.auxiliar=b.auxiliar  WHERE a.CUENTA>= '" & m_CtaIni & "' AND a.CUENTA<= '" & m_CtaFin & "' order by 1"
'        Case 9
'            Sql = "SELECT DISTINCT tip_linea as Auxiliar ,DESCRIP  as Descripción FROM cntablas  WHERE codtab='1' Order By 1"
'    End Select
'
'  If Sql <> Empty Or Sql <> "" Then
'
'    Set RS = ADOConexionEmp.Execute(Sql)
'    numreg = RS.RecordCount
'
'    If RS.BOF And RS.EOF Then
'        MsgBox "No hay registros para la consulta", vbInformation, "Sistema Contable"
'    Else
'        MSHConsultas.Clear
'        RS.MoveFirst
'        Do While Not RS.EOF
'          If RS.Fields.Count > 1 Then
'            Select Case m_Caso
'                Case 11
'                    MSHConsultas.AddItem CStr(Trim(CE(RS(0)))) & vbTab & CStr(Trim(CE(RS(1)))) & " " & CStr(Trim(CE(RS(2))))
'                Case Else
'                    MSHConsultas.AddItem CStr(Trim(CE(RS(0)))) & vbTab & CStr(Trim(CE(RS(1))))
'            End Select
'          Else
'            MSHConsultas.AddItem CStr(Trim(CE(RS(0))))
'          End If
'          RS.MoveNext
'        Loop
'        lblNro = MSHConsultas.Rows
'    End If
'
'    If MSHConsultas.Rows > 1 Then
'        MSHConsultas.FixedRows = 1
'    End If
'
'    MSHConsultas.TextMatrix(0, 0) = "Codigo"
'    MSHConsultas.TextMatrix(0, 1) = "Descripcion"
'
'    RS.Close
'
'  End If
'Set RS = Nothing
'Set RS_mon = Nothing
'End Function














