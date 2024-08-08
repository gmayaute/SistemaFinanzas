VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ConectarMYSQL(Host As String) As Boolean
On Error GoTo ControlErrores:
    ConectarMYSQL = False
    If oConexionMYSQL.State = MY_CONN_OPEN Then oConexionMYSQL.CloseConnection
        With oConexionMYSQL
            .ConnectionTimeout = 600
            .CursorLocation = adUseServer
            
             strAnioConex = gsBD
            
            If Host = gsServidor Then
                .OpenConnection "localhost", "root", "bdserverbr", strAnioConex
            Else
                .OpenConnection gsServidor, "root", "bdserverbr", strAnioConex
            End If
            ConectarMYSQL = True
            '----------- CREAR ODBC PARA ADO CON MYSQL ------------------'
            gsDSN = "BRANDTMYSQLHFM"
            Call CreateODBCUsuarioSistema(True, gsServidor, strAnioConex, gsDSN, "root", "bdserverbr")
            strADOConection = "DSN=" & gsDSN
            '------------------------------------------------------------'
        End With
         
Exit Function
ControlErrores:
    If oConexionMYSQL.State = MY_CONN_OPEN Then oConexionMYSQL.CloseConnection
    PantallaDeError
End Function

Public Function ConectarEmpresa() As Boolean
    On Error GoTo ControlErrores:
    If ADOConexion.State = adStateOpen Then ADOConexion.Close
    
    '----------- CREAR ODBC PARA ADO CON MYSQL ------------------'
    If strODBCEmp = "BRANDTMYSQLHFM" Then
    
        Call LeeConfiguracion
        gsDSN = "BRANDTMYSQLHFM"
    Else
        gsBD = strODBCEmp
        gsDSN = strODBCEmp
        
        If strODBCEmp = "BRANDTMYSQLEX" Then
            gsBD = "brandtexcont"
        End If
        
    End If
    
    strADOConection = "DSN=" & gsDSN
    
    '---------------------------------------------------------------------------
    
    Call CreateODBCUsuarioSistema(True, gsServidor, gsBD, strODBCEmp, "root", "bdserverbr")
    With ADOConexion
        
        .CursorLocation = adUseServer
        
        .ConnectionString = "DSN=" & strODBCEmp & ";uid=root; pwd=bdserverbr"
        .Open
        ConectarEmpresa = True
        ADOCommand.ActiveConnection = ADOConexion
    End With
    ConectarEmpresa = True
    
    
    Call CreateODBCUsuarioSistema(True, gsServidor, gsBD, gsDSN, "root", "bdserverbr")
    
    '------------------------------------------------------------'
    
    Exit Function
    
    
ControlErrores:
    If ADOConexion.State = adStateOpen Then ADOConexion.Close
    PantallaDeError
End Function
Public Function DesConectarSQL() As Boolean
On Error GoTo ControlErrores:
    DesConectarSQL = False
    If ADOConexion.State = adStateOpen Then ADOConexion.Close
    DesConectarSQL = True
Exit Function
ControlErrores:
End Function
Public Function IniciarTransaccion() As Boolean
On Error GoTo ControlErrores:
    IniciarTransaccion = False
    oConexionMYSQL.BeginTrans
    IniciarTransaccion = True
Exit Function
ControlErrores:
End Function
Public Function FinalizarTransaccion() As Boolean
On Error GoTo ControlErrores:
    FinalizarTransaccion = False
    oConexionMYSQL.CommitTrans
    FinalizarTransaccion = True
Exit Function
ControlErrores:
End Function
Public Function DeshacerTransaccion() As Boolean
On Error GoTo ControlErrores:
    DeshacerTransaccion = False
    oConexionMYSQL.RollbackTrans
    DeshacerTransaccion = True
Exit Function
ControlErrores:
End Function
Private Sub PantallaDeError()
Dim oError As ADODB.ERROR
Dim SERROR As String
Dim I As Integer
    With ADOConexion
        For Each oError In .Errors
            SERROR = ""
            SERROR = SERROR & "Numero" & vbTab & vbTab & ": " & err.Number & vbCrLf
            SERROR = SERROR & "Descripcion" & vbTab & ": " & err.Description & vbCrLf
            SERROR = SERROR & "Proveedor" & vbTab & ": " & err.Source & vbCrLf
            SERROR = SERROR & "Ejecutable" & vbTab & ": " & App.EXEName & vbCrLf
            SERROR = SERROR & "Titulo" & vbTab & vbTab & ": " & App.Title & vbCrLf
            If Len(Trim(err.Description)) > 0 Then MsgBox SERROR, vbOKOnly + vbCritical, "Error SQL Server"
        Next
        .Errors.Clear
    End With
End Sub
Public Function EjecutaSelect(sVista As String) As MYSQL_RS
On Error GoTo ControlErrores:
    Set EjecutaSelect = New MYSQL_RS
    With EjecutaSelect
        If .State = adStateOpen Then .CloseRecordset
        .CursorLocation = adUseServer
        .OpenRs "select * from " & sVista, oConexionMYSQL, adOpenDynamic, adLockOptimistic
    End With
Exit Function
ControlErrores:
End Function
Public Function EjecutaSelectRS(sQuery As String) As MYSQL_RS
    Set EjecutaSelectRS = New MYSQL_RS
On Error GoTo ControlErrores:
    With EjecutaSelectRS
        If .State = adStateOpen Then .CloseRecordset
        .CursorLocation = adUseServer
        .OpenRs sQuery, oConexionMYSQL, adOpenDynamic, adLockOptimistic
        
    End With
Exit Function
ControlErrores:
End Function
Public Function EjecutaInsertUpdateDelete(sQuery As String, Trans As Integer, lMensajeOK As Boolean) As Boolean
    Dim nRegistros As Long
    Dim sMess As String
    On Error GoTo ControlErrores:
    EjecutaInsertUpdateDelete = False
    IniciarTransaccion
    nRegistros = 0
    oConexionMYSQL.Execute sQuery, nRegistros
    FinalizarTransaccion
    If lMensajeOK Then
        Select Case Trans
            Case 1  'Insert
                sMess = "Se agregaron " & Trim(CStr(nRegistros)) & " registros."
            Case 2  'Update
                sMess = "Se actualizaron " & Trim(CStr(nRegistros)) & " registros."
            Case 3  'Delete
                sMess = "Se eliminaron " & Trim(CStr(nRegistros)) & " registros."
            Case Else
                sMess = "Ingrese el número de transacción"
        End Select
        MsgBox sMess, vbOKOnly + vbInformation, "Informe..."
    End If
    If nRegistros > 0 Then EjecutaInsertUpdateDelete = True
Exit Function
ControlErrores:
    DeshacerTransaccion
End Function

Public Function EjecutaInsertUpdateDelete2(sQuery As String, Trans As Integer, lMensajeOK As Boolean) As Long
    Dim nRegistros As Long
    Dim sMess As String
    On Error GoTo ControlErrores:
    EjecutaInsertUpdateDelete2 = 0
    IniciarTransaccion
    nRegistros = 0
    oConexionMYSQL.Execute sQuery, nRegistros
    FinalizarTransaccion
    If lMensajeOK Then
        Select Case Trans
            Case 1  'Insert
                sMess = "Se agregaron " & Trim(CStr(nRegistros)) & " registros."
            Case 2  'Update
                sMess = "Se actualizaron " & Trim(CStr(nRegistros)) & " registros."
            Case 3  'Delete
                sMess = "Se eliminaron " & Trim(CStr(nRegistros)) & " registros."
            Case Else
                sMess = "Ingrese el número de transacción"
        End Select
        MsgBox sMess, vbOKOnly + vbInformation, "Informe..."
    Else
        EjecutaInsertUpdateDelete2 = nRegistros
    End If
Exit Function
ControlErrores:
    DeshacerTransaccion
End Function




