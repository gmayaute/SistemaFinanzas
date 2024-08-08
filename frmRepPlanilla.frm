VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmRepPlanilla 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de boletas de remuneraciones"
   ClientHeight    =   2385
   ClientLeft      =   7320
   ClientTop       =   7230
   ClientWidth     =   4530
   Icon            =   "frmRepPlanilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEmp 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   930
      TabIndex        =   9
      Top             =   1590
      Width           =   2115
   End
   Begin VB.CheckBox chkExcel 
      BackColor       =   &H009F5539&
      Caption         =   "En Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   3090
      TabIndex        =   5
      Top             =   480
      Width           =   1395
   End
   Begin VB.CheckBox chkVac 
      BackColor       =   &H009F5539&
      Caption         =   "Rep. Vacac."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   3090
      TabIndex        =   4
      Top             =   240
      Width           =   1395
   End
   Begin VB.OptionButton optTipo 
      BackColor       =   &H009F5539&
      Caption         =   "Pre-impreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   1
      Left            =   3090
      TabIndex        =   1
      Top             =   30
      Width           =   1365
   End
   Begin VB.OptionButton optTipo 
      BackColor       =   &H009F5539&
      Caption         =   "Impreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   0
      Left            =   3030
      TabIndex        =   0
      Top             =   -210
      Width           =   1395
   End
   Begin Proyecto1.chameleonButton chBtnReporte 
      Height          =   375
      Left            =   3090
      TabIndex        =   3
      ToolTipText     =   "Ver Reporte"
      Top             =   765
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepPlanilla.frx":030A
      PICN            =   "frmRepPlanilla.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdadjuntar 
      Height          =   375
      Left            =   3570
      TabIndex        =   6
      ToolTipText     =   "Adjuntar Archivo"
      Top             =   765
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepPlanilla.frx":0868
      PICN            =   "frmRepPlanilla.frx":0884
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdVer 
      Height          =   315
      Left            =   3150
      TabIndex        =   7
      Top             =   1155
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "&Archivos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepPlanilla.frx":120E
      PICN            =   "frmRepPlanilla.frx":122A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton chBtnSalir 
      Height          =   375
      Left            =   4035
      TabIndex        =   8
      Top             =   765
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepPlanilla.frx":35AC
      PICN            =   "frmRepPlanilla.frx":35C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   15
      TabIndex        =   18
      Top             =   30
      Width           =   900
   End
   Begin MSForms.ComboBox CboAnio 
      Height          =   315
      Left            =   930
      TabIndex        =   17
      Top             =   15
      Width           =   2115
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3731;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   15
      TabIndex        =   16
      Top             =   824
      Width           =   900
   End
   Begin MSForms.ComboBox cboMon 
      Height          =   315
      Left            =   930
      TabIndex        =   15
      Top             =   809
      Width           =   2115
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3731;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   15
      TabIndex        =   14
      Top             =   1214
      Width           =   900
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   15
      TabIndex        =   13
      Top             =   427
      Width           =   900
   End
   Begin MSForms.ComboBox cboMes 
      Height          =   315
      Left            =   930
      TabIndex        =   12
      Top             =   412
      Width           =   2115
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3731;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboProceso 
      Height          =   300
      Left            =   930
      TabIndex        =   11
      Top             =   1206
      Width           =   2115
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3731;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empleado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   15
      TabIndex        =   10
      Top             =   1605
      Width           =   900
   End
   Begin VB.Label lblEmp 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   1995
      Width           =   4410
   End
End
Attribute VB_Name = "frmRepPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim RES As Integer
Dim Rs As MYSQL_RS
Dim RutaDestino As String
Private oConsulta As FrmConsultas
Dim IdentArch As String
Dim ContVer As Integer

Sub RepVaca()
    Set oReporte = New clsReporte
    If chkVac.Value = 1 Then
        DoEvents
        oReporte.Titulo = "RELACION DE VACACIONES GOZADAS EN EL MES DE " & NombreMes(cboMes.List(cboMes.ListIndex, 2), False) & " " & strAnoSistema
        oReporte.Reporte = "Rep_Vacaciones.rpt"
        oReporte.sp_Vacaiones cboMes.List(cboMes.ListIndex, 2)
    End If
End Sub

Private Sub cboMes_Change()
    IdentificadorArchivos
End Sub

Public Sub chBtnReporte_Click()
    Dim UsuAceptado As Boolean, opcMostrar As Boolean
    Dim RQ As MYSQL_RS, NumAut As Integer
    UsuAceptado = False
    If cboMon.List(cboMon.ListIndex, 1) <> "N" And cboMon.List(cboMon.ListIndex, 1) <> "" Then
        SQL = "select * from autorizaciones where codigo = 3"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        Do While Not RQ.EOF
            If Trim(RQ.Fields("usuario")) = strUsuarioId Then
                UsuAceptado = True
                opcMostrar = True
                Exit Do
            End If
            RQ.MoveNext
        Loop
        If UsuAceptado = False Then
            SQL = "select * from configuracion_acceso where codigo = 3"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                NumAut = RQ.Fields("numaut")
            End If
            SQL = "select IFNULL(count(*),0) as cantN from rh_tempacceso where codigo = 3 and usuario = '" & strUsuarioId & "' and autorizado = 'S'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                If val(RQ.Fields("CANTN")) < NumAut Then
                    If MsgBox("Usted no se encuentra autorizado a visualizar esta planilla. ¿Desea enviar una solicitud de autorización?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then
                        SQL = "SELECT IFNULL(COUNT(*),0) AS CANTAUT FROM rh_tempacceso WHERE USUARIO = '" & strUsuarioId & "' AND CODIGO = 3"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If Not RQ.EOF Then
                            If val(RQ.Fields("CANTAUT")) < NumAut Then
                                For i = 1 To NumAut
                                    SQL = "insert into rh_tempacceso(codigo,usuario,autorizado) values(3,'" & strUsuarioId & "','N')"
                                    oConexionMYSQL.Execute SQL
                                Next
                            End If
                        End If
                    End If
                    Exit Sub
                End If
            Else
                If MsgBox("Usted no se encuentra autorizado a visualizar esta planilla. ¿Desea enviar una solicitud de autorización?", vbQuestion + vbYesNo, "NOVADMIN") = vbYes Then
                    SQL = "SELECT IFNULL(COUNT(*),0) AS CANTAUT FROM rh_tempacceso WHERE USUARIO = '" & strUsuarioId & "' AND CODIGO = 3"
                    Set RQ = oConexion.EjecutaSelectRS(SQL)
                    If Not RQ.EOF Then
                        If val(RQ.Fields("CANTAUT")) < NumAut Then
                            For i = 1 To NumAut
                                SQL = "insert into rh_tempacceso(codigo,usuario,autorizado) values(3,'" & strUsuarioId & "','N')"
                                oConexionMYSQL.Execute SQL
                            Next
                        End If
                    End If
                End If
                Exit Sub
            End If
        End If
    Else
        SQL = "select * from autorizaciones where codigo = 3"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        Do While Not RQ.EOF
            If Trim(RQ.Fields("usuario")) = strUsuarioId Then
                UsuAceptado = True
                opcMostrar = True
                Exit Do
            End If
            RQ.MoveNext
        Loop
        SQL = "select * from configuracion_acceso where codigo = 3"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            NumAut = RQ.Fields("numaut")
        End If
        SQL = "select IFNULL(count(*),0) as cantN from rh_tempacceso where codigo = 3 and usuario = '" & strUsuarioId & "' and autorizado = 'S'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            If val(RQ.Fields("CANTN")) = NumAut Then
                opcMostrar = True
            End If
        End If
    End If
    Set oReporte = New clsReporte
    Select Case Me.Caption
        Case "Boletas de remuneraciones"
            'oReporte.Reporte = "Rep_Boleta.rpt"
            oReporte.Reporte = "Rep_Boleta_2017.rpt"
            oReporte.sp_Rep_Boleta cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), cboMon.List(cboMon.ListIndex, 1), Right("00000000000" & Trim(txtEmp), 11), opcMostrar
        Case "Certificado de Renta de Quinta"
            oReporte.Reporte = "Rep_Certificado_RtaQta.rpt"
            oReporte.sp_Rep_CRQ cboAnio.List(cboAnio.ListIndex, 1), cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), cboMon.List(cboMon.ListIndex, 1), Right("00000000000" & Trim(txtEmp), 11), opcMostrar
        Case "Certificado de Aporte de Pensiones"
            oReporte.Reporte = "Rep_Certificado_ApoPre.rpt"
            oReporte.sp_Rep_CAFP cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), cboMon.List(cboMon.ListIndex, 1), Right("00000000000" & Trim(txtEmp), 11), opcMostrar
        Case "Planilla x AFP"
            oReporte.Reporte = "Rep_PlanillaAFP.rpt"
            oReporte.sp_Rep_Planilla cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), cboMon.List(cboMon.ListIndex, 1), "AFP", "", opcMostrar
            
            RepVaca
            
            If frmRepPlanilla.chkExcel.Value = vbUnchecked Then
                oReporte.Reporte = "Rep_ValidacionesTareo.rpt"
                oReporte.sp_ValidaTareo cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), cboMon.List(cboMon.ListIndex, 1), CE(cboProceso.Text)
            End If
            
        Case "Planilla x División"
            oReporte.Reporte = "Rep_PlanillaDivi.rpt"
            oReporte.sp_Rep_Planilla cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), cboMon.List(cboMon.ListIndex, 1), "DIVI", txtEmp, opcMostrar
            RepVaca
            
        Case "Hoja de Renta de Quinta"
            oReporte.Reporte = "Rep_HojaTrabRtaQta.rpt"
            oReporte.Titulo = "HOJA DE TRABAJO DE RENTA DE QUINTA CATEGORIA A " & NombreMes(cboMes.List(cboMes.ListIndex, 2), False) & " " & strAnoSistema
            oReporte.sp_Rep_RtaQta cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), Left(cboMon, 1), Right("00000000000" & Trim(txtEmp), 11), opcMostrar
        Case "Asiento de Planilla"
            Dim RES As Integer
            RES = MsgBox("Esta seguro de generar el asiento contable para esta planilla", vbQuestion + vbYesNo, "NOVADMIN")
            On Error GoTo asiento
                If RES = vbYes Then
                    Select Case cboProceso.List(cboProceso.ListIndex, 2)
                        Case "1":
                            AsientoPlanilla cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), IIf(Left(cboMon, 1) = "C", "N", Left(cboMon, 1))
                        Case "4":
                            AsientoPlanillaProvision cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), "4", IIf(Left(cboMon, 1) = "C", "N", Left(cboMon, 1))
                        Case "5":
                            AsientoPlanillaProvision cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), "5", IIf(Left(cboMon, 1) = "C", "N", Left(cboMon, 1))
                        Case "6":
                            AsientoPlanillaProvision cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), "6", IIf(Left(cboMon, 1) = "C", "N", Left(cboMon, 1))
                    End Select
                    MsgBox "Se generó el asiento corrrectamente", vbInformation + vbOKOnly, "NOVPeru"
                End If
Exit Sub
asiento:
            MsgBox "Se produjo un error al generar el asiento, consulte con el administrador del sistema", vbCritical + vbOKOnly, "NOVADMIN"
    End Select
End Sub

Private Sub chBtnSalir_Click()
    Unload Me
End Sub

Private Sub cmdadjuntar_Click()
Dim RutaOrigen As String
Dim NomArchivo As String
Dim Ident As String
On Error GoTo CtrlError
    If chkExcel.Value = 1 And Me.Caption = "Planilla x AFP" And cboMes.ListIndex > -1 Then
        Acceso = True
        Screen.MousePointer = vbHourglass
        OpcExport = True
        CantVersiones = 0
        
        NomArchivo = "PLANILLA" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Mid(Format(Date, "dd/mm/yy"), 7, 8) & ".XLS"
        
        If EncuentraArchivo(NomArchivo, Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2)) Then
            CantVersiones = DevCantVer(NomArchivo, cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), ".XLS", 21)
            NomArchivo = "PLANILLA" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Mid(Format(Date, "dd/mm/yy"), 7, 8) & IIf(CantVersiones = 0, Empty, "_" & CantVersiones) & ".XLS"
            Ident = IdentArch
        Else
            Ident = GeneraIdentificadorAd("6", cboMes.List(cboMes.ListIndex, 1), cboMes.List(cboMes.ListIndex, 2))
            IdentArch = Ident
        End If
        
        Call chBtnReporte_Click
        
        RutaOrigen = Rep_Documents & "\" & NomArchivo
        CopiarArchivos RutaOrigen, RutaDestino, Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2)
        GrabarData "6", Ident, Replace(RutaDestino & "\" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "\", "\", "*"), NomArchivo
        SetAttr RutaDestino & "\" & Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2) & "\" & NomArchivo, vbReadOnly
        
        Call Kill(RutaOrigen)
        
        NomArchivo = "PLANI" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Mid(Format(Date, "dd/mm/yy"), 7, 8) & ".XLS"
        If GeneraReporteXLsPDF(NomArchivo, Ident) Then
            GeneraArchivoTareo cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboMon.List(cboMon.ListIndex, 1), cboProceso.List(cboProceso.ListIndex, 2)
            Screen.MousePointer = vbDefault
            MsgBox "Archivos adjuntados", vbInformation, "NOVADMIN"
            OpcExport = False
            CantVersiones = 0
            
            chkExcel.Value = 0
        Else
            MsgBox "No se adjuntaron todos los archivos. Consulte con el Administrador de Sistemas", vbInformation, "NOVADMIN"
        End If
    Else
        MsgBox "Debe seleccionar el check EN EXCEL para adjuntar los archivos de la planilla", vbInformation, "NOVADMIN"
    End If
    Exit Sub
CtrlError:
    MsgBox "Error Adjuntando archivos", vbCritical, "NOVADMIN"
End Sub

Sub GeneraArchivoTareo(AnoMes As String, mon As String, tipoplani As String)
Dim RQ As MYSQL_RS
Dim NomArchivo As String
    SQL = "select * from pl_tareo where anomes = '" & AnoMes & "'"
    If mon <> "" Then SQL = SQL & " and mon = '" & mon & "'"
    If tipoplani <> "" Then SQL = SQL & " and tipo = '" & tipoplani & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        NomArchivo = "TAREO" & AnoMes & "_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Mid(Format(Date, "dd/mm/yy"), 7, 8) & ".XLS"
        CantVersiones = DevCantVer(NomArchivo, AnoMes, ".XLS", 18)
        NomArchivo = "TAREO" & AnoMes & "_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Mid(Format(Date, "dd/mm/yy"), 7, 8) & IIf(CantVersiones = 0, Empty, "_" & CantVersiones) & ".XLS"
        If Exportar_Excel_rs(Rep_Documents & "\" & NomArchivo, RQ) Then
            CopiarArchivos Rep_Documents & "\" & NomArchivo, RutaDestino, AnoMes
            GrabarData "6", IdentArch, Replace(RutaDestino & "\" & AnoMes & "\", "\", "*"), NomArchivo
            SetAttr RutaDestino & "\" & AnoMes & "\" & NomArchivo, vbReadOnly
            Kill Rep_Documents & "\" & NomArchivo
        End If
    End If
    Set RQ = Nothing
End Sub

Function GeneraReporteXLsPDF(NomArchivo As String, identificador As String) As Boolean
Dim RutaOrigen As String
Dim Ident As String
Dim TempNomArc As String
On Error GoTo CtrlError
    GeneraReporteXLsPDF = False
    'chkExcel.Value = 0
    chBtnReporte_Click
    Dim frmRep As frmReportPreview
    Set frmRep = New frmReportPreview
    TempNomArc = Replace(NomArchivo, "XLS", "PDF")
    CantVersiones = DevCantVer(NomArchivo, cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), ".XLS", 18)
    NomArchivo = "PLANI" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Mid(Format(Date, "dd/mm/yy"), 7, 8) & IIf(CantVersiones = 0, Empty, "_" & CantVersiones) & ".XLS"
    frmRep.ExportaExcel NomArchivo
    RutaOrigen = Rep_Documents & "\" & NomArchivo
    CopiarArchivos RutaOrigen, RutaDestino, Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2)
    GrabarData "6", identificador, Replace(RutaDestino & "\" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "\", "\", "*"), NomArchivo
    SetAttr RutaDestino & "\" & Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2) & "\" & NomArchivo, vbReadOnly
    Kill RutaOrigen
    CantVersiones = DevCantVer(TempNomArc, cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), ".PDF", 18)
    NomArchivo = "PLANI" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "_" & Right("00" & Day(Date), 2) & Right("00" & Month(Date), 2) & Mid(Format(Date, "dd/mm/yy"), 7, 8) & IIf(CantVersiones = 0, Empty, "_" & CantVersiones) & ".PDF"
    frmRep.ExportaPdf NomArchivo
    RutaOrigen = Rep_Documents & "\" & NomArchivo
    CopiarArchivos RutaOrigen, RutaDestino, Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2)
    GrabarData "6", identificador, Replace(RutaDestino & "\" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "\", "\", "*"), NomArchivo
    SetAttr RutaDestino & "\" & Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2) & "\" & NomArchivo, vbReadOnly
    Kill RutaOrigen
    GeneraReporteXLsPDF = True
Exit Function
CtrlError:
    GeneraReporteXLsPDF = False
    MsgBox "Error Adjuntando archivos", vbCritical, "NOVADMIN"
End Function

Function DevCantVer(Nom As String, AnoMes As String, Exten As String, Pos As Integer) As Integer
    DevCantVer = False
    Dim SQL As String
    Dim RQ As MYSQL_RS
    SQL = "select count(*) as cant from archivosadjuntos where modulo = 6 and left(identificador,6) = '" & AnoMes & "' " & _
          "and ruta = '" & Replace(RutaDestino & "\" & Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2) & "\", "\", "*") & "' " & _
          "and LEFT(nombre," & Pos & ") = '" & Mid(Nom, 1, InStr(1, Nom, Exten) - 1) & "' and RIGHT(nombre,4) = '" & Exten & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        DevCantVer = RQ.Fields("CANT")
    End If
    Set RQ = Nothing
End Function

Function EncuentraArchivo(Nom As String, AnoMes As String) As Boolean
    EncuentraArchivo = False
    Dim SQL As String
    Dim RQ As MYSQL_RS
    SQL = "select * from archivosadjuntos where modulo = 6 and left(identificador,6) = '" & AnoMes & "' " & _
          "and ruta = '" & Replace(RutaDestino & "\" & Trim(cboMes.List(cboMes.ListIndex, 1)) & cboMes.List(cboMes.ListIndex, 2) & "\", "\", "*") & "' and nombre = '" & Nom & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        EncuentraArchivo = True
    End If
    Set RQ = Nothing
End Function

Sub GrabarData(Modulo As String, Ident As String, Ruta As String, NomAr As String)
    SQL = "call Insert_ArchAdjuntos('" & Modulo & "','" & Ident & "','" & Ruta & "','" & NomAr & "')"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
End Sub

Private Sub cmdver_Click()
    IdentificadorArchivos
    frmArchivosAdjuntos.IdentificadorAr = IdentArch
    frmArchivosAdjuntos.AnioSel = cboMes.List(cboMes.ListIndex, 1)
    frmArchivosAdjuntos.MesSel = cboMes.List(cboMes.ListIndex, 2)
    frmArchivosAdjuntos.Show
End Sub

Private Sub Form_Activate()
    Select Case Me.Caption
        Case "Boletas de remuneraciones"
            Me.Height = 2790
            Lbl(3).Caption = "Empleado"
        Case "Certificado de Renta de Quinta"
            Me.Height = 2790
            Lbl(3).Caption = "Empleado"
        Case "Certificado de Aporte de Pensiones"
            Me.Height = 2790
            Lbl(3).Caption = "Empleado"
        Case "Hoja de Renta de Quinta"
            Me.Height = 2790
            Lbl(3).Caption = "Empleado"
        Case "Planilla x AFP"
            Me.Height = 2205
            Lbl(3).Visible = False
            txtEmp.Visible = False
        Case "Planilla x División"
            Me.Height = 2790
            Lbl(3).Caption = "División"
        Case "Asiento de Planilla"
            Me.Height = 2205
            Lbl(3).Visible = False
            txtEmp.Visible = False
            cmdadjuntar.Visible = False
            cmdver.Visible = False
    End Select
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    RutaDestino = "\\" & gsServidorApp & "\ddsi$\RRHH\PLANILLAS"
    Set oConsulta = New FrmConsultas
    LlenarMesP cboMes
    LlenarMonedaP cboMon
    LlenarAnio
    cboMon.AddItem "CONFIDENCIAL"
    cboMon.List(3, 1) = "C"
    LlenarProcesos cboProceso
End Sub

Public Sub LlenarAnio()
    Dim i As Integer, anioctual As Integer
    cboAnio.Clear
    cboAnio.AddItem "Seleccionar..."
    cboAnio.List(0, 1) = strAnoSistema
    For i = 1 To 3
        cboAnio.AddItem Trim(str(CDbl(strAnoSistema) - (i - 1)))
        cboAnio.List(i, 1) = Trim(str(CDbl(strAnoSistema) - (i - 1)))
        If Trim(Year(Date)) = Trim(cboAnio.List(i, 1)) Then
            anioctual = i
        End If
    Next
    cboAnio.ListIndex = anioctual
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BorrarSolicitud
End Sub

Sub BorrarSolicitud()
    SQL = "delete from rh_tempacceso where usuario = '" & strUsuarioId & "' and codigo = 3"
    oConexionMYSQL.Execute SQL
End Sub

Private Sub txtEmp_Change()
    If IsNumeric(txtEmp) Then
        lblEmp = DescripcionesdeCodigos("EMPLEADO", Right("00000000000" & Trim(txtEmp), 11))
    Else
        lblEmp = ""
    End If
     Select Case Me.Caption
        Case "Planilla x AFP"
           lblEmp = ""
        Case "Planilla x División"
           lblEmp = DescripcionesdeCodigos("DES_DIVISION", Right("0000" & Trim(txtEmp), 4))
        Case Else
           lblEmp = DescripcionesdeCodigos("EMPLEADO", Right("00000000000" & Trim(txtEmp), 11))
    End Select
End Sub

Private Sub txtEmp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Empleados"
            .pForm = FORM_REPPLANI
            .pCaso = LABEL_EMP
            .Show
        End With
    End If
End Sub

Sub IdentificadorArchivos()
    Dim SQL As String
    Dim RQ As MYSQL_RS
    IdentArch = ""
    SQL = "select distinct identificador from archivosadjuntos where modulo = 6 and left(identificador,6) = '" & cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        IdentArch = Left(Trim(RQ.Fields("identificador")), 6)
    End If
    Set RQ = Nothing
End Sub

