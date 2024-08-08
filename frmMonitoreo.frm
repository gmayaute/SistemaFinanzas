VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMonitoreo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor del Sistema"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   Icon            =   "frmMonitoreo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10080
   Begin MSFlexGridLib.MSFlexGrid flxMonitor 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   8916
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMonitoreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ConfigMonitor()
    With flxMonitor
        .Clear
        .Cols = 6
        .Rows = 1
        .row = 0
        .Col = 0
        .CellForeColor = &H80000002
        .ColWidth(0) = 800
        .ColAlignment(0) = MSHFLEXGRID_ALINEACION.Centro
        .TextMatrix(0, 0) = "Item"
        
        .row = 0
        .Col = 1
        .CellForeColor = &H80000002
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = Space(5) + "Usuario"
        
        .row = 0
        .Col = 2
        .CellForeColor = &H80000002
        .ColWidth(2) = 3550
        .ColAlignment(2) = MSHFLEXGRID_ALINEACION.Centro
        .TextMatrix(0, 2) = "Empresa"
        
        .row = 0
        .Col = 3
        .CellForeColor = &H80000002
        .ColWidth(3) = 1200
        .ColAlignment(3) = MSHFLEXGRID_ALINEACION.Centro
        .TextMatrix(0, 3) = "Fec. Conexión"
        
        .row = 0
        .Col = 4
        .CellForeColor = &H80000002
        .ColWidth(4) = 1200
        .ColAlignment(4) = MSHFLEXGRID_ALINEACION.Centro
        .TextMatrix(0, 4) = "Hor. Conexión"
        
        .row = 0
        .Col = 5
        .CellForeColor = &H80000002
        .ColWidth(5) = 2000
        .ColAlignment(5) = MSHFLEXGRID_ALINEACION.Centro
        .TextMatrix(0, 5) = "Host de Conexión"
    End With
End Sub

Private Sub MuestraConeciones()
    Dim I As Integer
    Dim SQL As String
    Dim rs As MYSQL_RS
    Me.Left = 0
    Me.Top = 0
    I = 1
    Set rs = oConexion.EjecutaSelect("cia_user_connect")
    With flxMonitor
        Do While Not (rs.EOF)
            .Rows = .Rows + 1
            .TextMatrix(I, 0) = CStr(I)
            .TextMatrix(I, 1) = rs.Fields("usuario_id")
            .TextMatrix(I, 2) = rs.Fields("codcia")
            .TextMatrix(I, 3) = rs.Fields("fec_conexion")
            .TextMatrix(I, 4) = rs.Fields("hor_conexion")
            .TextMatrix(I, 5) = rs.Fields("host_conexion")
            I = I + 1
            rs.MoveNext
        Loop
    End With
    rs.CloseRecordset
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Call WheelHook(frmMonitoreo)
    ConfigMonitor
    MuestraConeciones
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    
    On Error Resume Next
    
    With flxMonitor
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then
            Lstep = 10
        End If
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then
                NewValue = 0
            End If
        Else
            NewValue = .TopRow + Lstep
            If NewValue > .Rows - 1 Then
                NewValue = .Rows - 1
            End If
        End If
        .TopRow = NewValue
    End With
End Sub
