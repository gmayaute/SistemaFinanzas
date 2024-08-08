VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADODataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************
' DECLARACIONES GENERALES
'**************************************************************************

' *** Timers para la Apertura de la Conexion
Public Enum TEL_CONNECT_TIMERS_ENUM
    TEL_CONNECT_SHORT_TIMEOUT = 50000
    TEL_CONNECT_NORMAL_TIMEOUT = 60000
    TEL_CONNECT_LONG_TIMEOUT = 2000000
End Enum

' *** Tipos de Bloqueo para la Apertura de Recordsets
Public Enum TEL_LOCK_TYPES
    TEL_LOCK_READ_ONLY = ADODB.adLockReadOnly
    TEL_LOCK_OPTIMISTIC = ADODB.adLockOptimistic
    TEL_LOCK_PESSIMISTIC = ADODB.adLockPessimistic
    TEL_LOCK_BATCH_OPTIMISTIC = ADODB.adLockBatchOptimistic
End Enum

' *** Ubicacion del Cursor
Public Enum TEL_CURSOR_LOCATION
    TEL_CURSOR_NONE = ADODB.adUseNone
    TEL_CURSOR_CLIENT = ADODB.adUseClient
    TEL_CURSOR_SERVER = ADODB.adUseServer
End Enum

' *** Tipos de Cursores para Recordsets
Public Enum TEL_CURSOR_TYPE
    TEL_TYPE_FORWARD_ONLY = ADODB.adOpenForwardOnly
    TEL_TYPE_KEYSET = ADODB.adOpenKeyset
    TEL_TYPE_DYNAMIC = ADODB.adOpenDynamic
    TEL_TYPE_STATIC = ADODB.adOpenStatic
End Enum

' *** Columnas del Buffer de Parametros
Public Enum TEL_SP_PARAM_BUFFER
    TEL_SP_PARAMET_NAME = 0
    TEL_SP_PARAMET_TYPE = 1
    TEL_SP_PARAMET_DIRECTION = 2
    TEL_SP_PARAMET_SIZE = 3
    TEL_SP_PARAMET_VALUE = 4
End Enum

' *** Direcciones de los Parametros
Public Enum TEL_SP_PARAM_DIRECTION
    TEL_PARAM_INPUT = ADODB.adParamInput
    TEL_PARAM_OUTPUT = ADODB.adParamOutput
    TEL_PARAM_INPUT_OUTPUT = ADODB.adParamInputOutput
    TEL_PARAM_RETURN_VALUE = ADODB.adParamReturnValue
End Enum

Private Const msOBJECT_NAME As String = "IDataAccess"
' *** Variables para Albergar Objetos de ADO
Private moCachedConnection As ADODB.Connection

Public Property Get Id() As String
    Id = msOBJECT_NAME
End Property

Private Sub Class_Terminate()
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.Close
        Set moCachedConnection = Nothing
    End If
End Sub

'**************************************************************************
' Nombre....: CacheConnection()
' Proposito.: Esta funcion abre una conexion de ADO y la mantiene Abierta.
' Inputs....: Opcional Cadena de Connección y Valor para ConnectTimeOut.
' Retorno...: Retorna TRUE si la Función es Satisfactoria
' Efectos...: Cuando es llamada sin argumentos CIERRA la Conexion.
'**************************************************************************
Public Function CacheConnection(Optional ByVal sConnect As String) As Boolean

    
    On Error GoTo ErrorHandler
    Dim lErrNo      As Long
    Dim sErrDesc    As String
    Dim sErrSrc     As String
    
    
    CacheConnection = False
    
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.Close
        Set moCachedConnection = Nothing
    End If

    If sConnect <> vbNullString Then
        Set moCachedConnection = New ADODB.Connection
        moCachedConnection.ConnectionTimeout = TEL_CONNECT_TIMERS_ENUM.TEL_CONNECT_NORMAL_TIMEOUT
        moCachedConnection.ConnectionString = sConnect
        moCachedConnection.CursorLocation = adUseClient
        moCachedConnection.Open
    End If
    
    CacheConnection = True
    
    Exit Function
    
ErrorHandler:
    lErrNo = err.Number
    sErrDesc = err.Description
    sErrSrc = err.Source
    Set moCachedConnection = Nothing
End Function

'**************************************************************************
' Nombre....: DestroyConecction()
' Proposito.: Esta Función Destruye la Conexión
' Inputs....: Ninguno
' Se asume..: Ninguno
' Retorno...: Retorna TRUE si la Función es Satisfactoria
' Efectos...: Ninguno
'**************************************************************************
Public Function DestroyConecction() As Boolean

    On Error GoTo ErrorHandler
    Dim lErrNo      As Long
    Dim sErrDesc    As String
    Dim sErrSrc     As String
    
    DestroyConecction = False
    
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.Close
    End If

    Set moCachedConnection = Nothing
    
    DestroyConecction = True
    
    Exit Function
    
ErrorHandler:
    lErrNo = err.Number
    sErrDesc = err.Description
    sErrSrc = err.Source
    
    Set moCachedConnection = Nothing
End Function

'**************************************************************************
' Nombre....: IniciarTransaccion()
' Proposito.: Realiza un Begin en la Transacción
' Inputs....: N/A
' Se asume..: N/A
' Retorno...: N/A
' Efectos...: N/A
'**************************************************************************
Public Function IniciarTransaccion()
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.BeginTrans
    End If
End Function

'**************************************************************************
' Nombre....: FinalizarTransaccion()
' Proposito.: Realiza un Commit en la Transacción
' Inputs....: N/A
' Se asume..: N/A
' Retorno...: N/A
' Efectos...: N/A
'**************************************************************************
Public Function FinalizarTransaccion()
On Error GoTo ERROR:
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.CommitTrans
    End If
ERROR:
End Function

'**************************************************************************
' Nombre....: CancelarTransaccion()
' Proposito.: Realiza un RollBack en la Transacción
' Inputs....: N/A
' Se asume..: N/A
' Retorno...: N/A
' Efectos...: N/A
'**************************************************************************
Public Function CancelarTransaccion()
    On Error GoTo ERROR:
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.RollbackTrans
    End If
    Exit Function
ERROR:
    
End Function

'**************************************************************************
' Nombre....: ExecuteActionSP()
' Proposito.: Ejecuta un sp que no devuelve filas (Query de acciones)
' Inputs....: Stored Procedure, Array de Parametros
' Se asume..: N/A
' Retorno...: Retorna TRUE si la función tiene exito
' Efectos...: Retorna en el mismo arreglo de parametros los valores
'             de los parametros del SP definidos como OUTPUT.
'**************************************************************************
Public Function ExecuteActionSP(ByVal cCadena As String) As Boolean
                    
    On Error GoTo ErrorHandler
    Dim cadena As String
    Dim lErrNo As Long
    Dim sErrDesc As String
    Dim sErrSrc     As String
    
    Dim oCommand As ADODB.Command
    Dim oParam As ADODB.Parameter
    
    ExecuteActionSP = False
    
    Set oCommand = New ADODB.Command
    Set oCommand.ActiveConnection = moCachedConnection
    oCommand.CommandText = cCadena
    oCommand.CommandTimeout = TEL_CONNECT_TIMERS_ENUM.TEL_CONNECT_NORMAL_TIMEOUT
    oCommand.Execute
    
    Set oCommand = Nothing
    Set oParam = Nothing

    ExecuteActionSP = True
    
    Exit Function
    
ErrorHandler:
            
            lErrNo = err.Number
            sErrDesc = UCase(err.Description)
            sErrSrc = err.Source
            
            ExecuteActionSP = False
            
            On Error Resume Next
            
            Set oCommand = Nothing
            Set oParam = Nothing
            
            On Error GoTo 0
            Dim msgErr As String
            Dim PosIni As Integer
            Dim PosFin As Integer
            
            Select Case lErrNo
                Case -2147217873
                
                        If InStr(1, sErrDesc, "PRIMARY KEY") > 0 Then
                            msgErr = "Esta intentado insertar un registro que ya EXISTE." & Chr(13)
                        ElseIf InStr(1, sErrDesc, "INSERT") > 0 Then
                            If InStr(1, sErrDesc, "NULL") > 0 Then
                                msgErr = "Esta intentado insertar un valor NULO en un campo obligatorio."
                            ElseIf InStr(1, sErrDesc, "FOREIGN KEY") > 0 Then
                                PosIni = InStr(1, sErrDesc, "'FK")
                                PosFin = InStr(PosIni + 1, sErrDesc, "'")
                                cadena = Mid(sErrDesc, PosIni, PosFin + 1 - PosIni)
                                msgErr = "Esta intentado insertar un valor que no se encuentra la tabla maestra." & Salto(2) & "Mensaje : " & cadena
                                
                            Else
                                msgErr = "ERROR AL TRATAR DE INSERTAR UN REGISTRO"
                            End If
                        ElseIf InStr(1, sErrDesc, "UPDATE") > 0 Then
                            If InStr(1, sErrDesc, "NULL") > 0 Then
                                msgErr = "Esta intentado insertar un valor NULO en un campo obligatorio."
                            ElseIf InStr(1, sErrDesc, "FOREIGN KEY") > 0 Then
                                PosIni = InStr(1, sErrDesc, "'FK")
                                PosFin = InStr(PosIni + 1, sErrDesc, "'")
                                cadena = Mid(sErrDesc, PosIni, PosFin + 1 - PosIni)
                            
                                msgErr = "Esta intentado actualizar con un valor que no se encuentra la tabla maestra." & Salto(2) & "Mensaje : " & cadena
                            Else
                                msgErr = "ERROR AL TRATAR DE ACTUALIZAR EL REGISTRO"
                            End If
                        ElseIf InStr(1, sErrDesc, "DELETE") > 0 Then
                            If InStr(1, sErrDesc, "REFERENCE") > 0 Then
                                msgErr = "Esta intentado borrar un registro del cual dependen otros datos."
                            Else
                                msgErr = "ERROR AL TRATAR DE ELIMINAR EL REGISTRO"
                            End If
                        Else
                            msgErr = err.Description
                        End If

                Case Else
                        msgErr = LCase(sErrDesc)
            End Select
            
            Mensajes msgErr, vbInformation + vbOKOnly
            
End Function


'**************************************************************************
' Nombre....: GetRecordsetSP()
' Proposito.: Obtiene un Recordset desconectado ejecutando un sp
' Inputs....:
' Se asume..: N/A
' Retorno...: Retorna TRUE si la función tiene exito
' Efectos...: Retrona en el mismo arreglo de parametros los valores
'             de los parametros del SP definidos como OUTPUT.
'**************************************************************************
' Cambio Mensaje de Error
' Optional blnMensaje As Boolean = True
Public Function GetRecordSetSP(ByVal cCadena As String, _
                                ByRef oRecordset As ADODB.Recordset, _
                                Optional blnMensaje As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lErrNo As Long
    Dim sErrDesc As String
    Dim sErrSrc As String
    
    Dim oCommand As ADODB.Command
    Dim oParam As ADODB.Parameter
    
    GetRecordSetSP = False

    Set oRecordset = New ADODB.Recordset
    oRecordset.CursorLocation = TEL_CURSOR_LOCATION.TEL_CURSOR_CLIENT
    oRecordset.CursorType = TEL_CURSOR_TYPE.TEL_TYPE_STATIC
    oRecordset.LockType = TEL_LOCK_TYPES.TEL_LOCK_BATCH_OPTIMISTIC
    
    Set oCommand = New ADODB.Command
    Set oCommand.ActiveConnection = moCachedConnection
    oCommand.CommandText = cCadena
    oCommand.CommandTimeout = TEL_CONNECT_TIMERS_ENUM.TEL_CONNECT_NORMAL_TIMEOUT

    
    Set oRecordset = oCommand.Execute()
    
    Set oCommand = Nothing
    Set oParam = Nothing
    
    GetRecordSetSP = True
    
    Exit Function
    
ErrorHandler:
        lErrNo = err.Number
        sErrDesc = err.Description
        sErrSrc = err.Source
        
        Set oRecordset = Nothing
        Set oCommand = Nothing
        Set oParam = Nothing
        
        If blnMensaje Then
            If lErrNo = 3709 Then
                MsgBox "No se puede establecer la conexión, intentelo nuevamente ", vbInformation, ""
                
            Else
                Mensajes "Numero de Error: " & CStr(lErrNo) & vbCrLf & "Modulo: " & msOBJECT_NAME & ".ExecuteActionSP()" & vbCrLf & sErrSrc & vbCrLf & sErrDesc, vbInformation
            End If
            
        End If

End Function






