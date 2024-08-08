VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADOConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CadenaConexion As String
Private m_NombreSP As String
Private m_FlagCnx As Boolean
Private m_EjecutarConsulta As Boolean
Private objDatos As clsADODataAccess
Private m_Tabla As ADODB.Recordset

Public Property Get CadenaDeConexion() As String
    CadenaDeConexion = m_CadenaConexion
End Property

Public Property Let CadenaDeConexion(ByVal strData As String)
    m_CadenaConexion = strData
End Property

Public Property Get NombreSP() As String
    NombreSP = m_NombreSP
End Property

Public Property Let NombreSP(ByVal strData As String)
    m_NombreSP = strData
End Property

Private Sub Class_Initialize()
    m_FlagCnx = False
End Sub

Private Sub CerrarRecordSet(ByRef rsParam As Recordset)
    If rsParam Is Nothing Then Exit Sub
    If rsParam.State <> 0 Then rsParam.Close
    Set rsParam = Nothing
End Sub

Private Sub Class_Terminate()
    CerrarRecordSet m_Tabla
    
    If Not objDatos Is Nothing Then
        objDatos.DestroyConecction
    End If
End Sub

Public Function InicializaClase()
    Set m_Tabla = New ADODB.Recordset
    Set objDatos = New clsADODataAccess
    objDatos.CacheConnection (CadenaDeConexion)
End Function

Public Function BeginTrans()
    objDatos.IniciarTransaccion
End Function

Public Function CommitTrans()
    objDatos.FinalizarTransaccion
End Function

Public Function CancelTrans()
    objDatos.CancelarTransaccion
End Function

Public Function FinalizaClase()
    objDatos.DestroyConecction
End Function

Public Function MantenimientoDeTablas(ByVal cCnx As String, ByVal cCadena As String, _
                                      Optional ByVal blnStatus As Boolean, _
                                      Optional ByVal bConTransacciones As Boolean = True, _
                                      Optional ByVal bForzarCommit As Boolean = False) As Boolean
    
    
    CadenaDeConexion = cCnx
    
    If Not m_FlagCnx Then
        InicializaClase
        
        If bConTransacciones Then
            Me.BeginTrans
        End If
        
        m_FlagCnx = True
    End If
    
    m_EjecutarConsulta = False
    
    
    MantenimientoDeTablas = EjecutarTransaccion(cCadena, False)
    
    If bConTransacciones Then
        If MantenimientoDeTablas = True Then
            If blnStatus Then ' si se ejecuto la transaccion y el blnstatus = true entonces fuerza a cerrar la clase
                Me.CommitTrans
                Me.FinalizaClase
            End If
        Else
            If bForzarCommit = True Then
                Me.CommitTrans 'si sale un error igual hacer un commit a la transaccion
            Else
                Me.CancelTrans 'si sale un error hace un rollback de la transaccion
            End If
            
            If blnStatus = True Then
                Me.FinalizaClase
            End If
        End If
    Else
        If blnStatus = True Then
            Me.FinalizaClase
        End If
    End If
    
End Function

Public Function ConsultaDatosTabla(ByVal cCnx As String, ByVal cCadena As String, Optional blnMensaje As Boolean = True) As ADODB.Recordset

    CadenaDeConexion = cCnx
    
    Call InicializaClase
    
    m_EjecutarConsulta = True
    
    If EjecutarTransaccion(cCadena, blnMensaje) Then
        Set ConsultaDatosTabla = Clone(m_Tabla, adLockReadOnly)
    Else
        Set ConsultaDatosTabla = Nothing
    End If

    Call FinalizaClase
End Function

Private Function Clone(ByVal oRS As ADODB.Recordset, _
    Optional ByVal LockType As ADODB.LockTypeEnum = adLockUnspecified) As ADODB.Recordset
    
    Dim oStream As ADODB.Stream
    Dim oRsClone As ADODB.Recordset
    
    Set oStream = New ADODB.Stream
    oRS.Save oStream
    
    Set oRsClone = New ADODB.Recordset
    oRsClone.Open oStream, , , LockType
    
    Set Clone = oRsClone
    
    Set oRsClone = Nothing
End Function


Private Function EjecutarTransaccion(cCadena As String, Optional blnMensaje As Boolean = True) As Boolean
    
    On Error GoTo ErrHandler
    Dim intNroFil As Integer
    Dim intDireccion As Integer
    Dim lngNroFilas As Long
    
    Dim blnResultado As Boolean
        
    If Not m_EjecutarConsulta Then
        If m_Tabla.State = adStateOpen Then m_Tabla.Close
        blnResultado = objDatos.ExecuteActionSP(cCadena)
        If blnResultado = True Then
            EjecutarTransaccion = True
        Else
            EjecutarTransaccion = False
        End If
    Else
        If m_Tabla.State = adStateOpen Then m_Tabla.Close
    
        objDatos.GetRecordSetSP cCadena, m_Tabla, blnMensaje
        
        
        If m_Tabla Is Nothing Then
            Mensajes "No se encontraron datos"
        Else
            If m_Tabla.RecordCount > 0 Then
                EjecutarTransaccion = True
            Else
                EjecutarTransaccion = False
            End If
        End If
    End If
        
    Exit Function

ErrHandler:

    If blnMensaje Then
        Mensajes "Error al ejecutar la transaccion, intente nuevamente..." & Salto(2) & err.Source & Salto(1) & err.Description, vbInformation

    End If
    EjecutarTransaccion = False
End Function
    


