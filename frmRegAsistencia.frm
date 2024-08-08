VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmRegAsistencia 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de incidencias de Entrada y Salida de personal"
   ClientHeight    =   5745
   ClientLeft      =   6015
   ClientTop       =   2925
   ClientWidth     =   4080
   ControlBox      =   0   'False
   Icon            =   "frmRegAsistencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmlistado 
      BackColor       =   &H009F5539&
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8220
      Begin Proyecto1.chameleonButton cmdingresar 
         Height          =   375
         Left            =   1178
         TabIndex        =   20
         Top             =   5310
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Aceptar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         FCOL            =   12648384
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRegAsistencia.frx":030A
         PICN            =   "frmRegAsistencia.frx":0326
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.ListBox LstSede 
         Height          =   5250
         Left            =   45
         TabIndex        =   21
         Top             =   0
         Width           =   3945
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "6959;9260"
         MatchEntry      =   0
         ListStyle       =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7740
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   1635
      Left            =   2220
      TabIndex        =   2
      Top             =   180
      Width           =   1485
      Begin VB.Label lblfoto 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIN FOTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   615
         Left            =   390
         TabIndex        =   15
         Top             =   510
         Width           =   705
      End
      Begin VB.Image imgFoto 
         Height          =   1425
         Left            =   60
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1365
      End
   End
   Begin Proyecto1.chameleonButton cmdAceptar 
      Height          =   375
      Left            =   210
      TabIndex        =   3
      Top             =   1440
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      FCOL            =   12648384
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRegAsistencia.frx":0480
      PICN            =   "frmRegAsistencia.frx":049C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnReporte 
      Height          =   345
      Left            =   7680
      TabIndex        =   13
      Top             =   1590
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   609
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
      MICON           =   "frmRegAsistencia.frx":05F6
      PICN            =   "frmRegAsistencia.frx":0612
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpIngreso 
      Height          =   285
      Left            =   6210
      TabIndex        =   16
      Top             =   1620
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      _Version        =   393216
      Format          =   94109697
      CurrentDate     =   39241.3699537037
   End
   Begin Proyecto1.chameleonButton cmdEnviar 
      Height          =   345
      Left            =   3870
      TabIndex        =   17
      Top             =   1590
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Enviar Registro"
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
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRegAsistencia.frx":0B54
      PICN            =   "frmRegAsistencia.frx":0B70
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox meCodigo 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   990
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "########"
      Mask            =   "########"
      PromptChar      =   "_"
   End
   Begin Proyecto1.chameleonButton btnSalir 
      Height          =   345
      Left            =   7680
      TabIndex        =   22
      Top             =   1995
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   609
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
      MICON           =   "frmRegAsistencia.frx":110A
      PICN            =   "frmRegAsistencia.frx":1126
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   3495
      Left            =   30
      TabIndex        =   6
      Top             =   2340
      Width           =   8145
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRegistro 
         Height          =   3255
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   5741
         _Version        =   393216
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Categoría:"
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
      Left            =   3750
      TabIndex        =   18
      Top             =   1290
      Width           =   885
   End
   Begin VB.Image imgLlave 
      Height          =   240
      Left            =   1890
      Picture         =   "frmRegAsistencia.frx":14EC
      Top             =   1530
      Width           =   240
   End
   Begin MSForms.Label lblSede 
      Height          =   285
      Left            =   30
      TabIndex        =   14
      ToolTipText     =   "Hora Actual"
      Top             =   2040
      Width           =   8145
      ForeColor       =   12632256
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "AQUI NOMBRE DE SEDE"
      Size            =   "14367;503"
      BorderColor     =   -2147483639
      FontName        =   "Arial"
      FontEffects     =   1073741829
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Horario:"
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
      Left            =   3750
      TabIndex        =   12
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sede:"
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
      Height          =   165
      Left            =   3750
      TabIndex        =   11
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo:"
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
      Height          =   225
      Left            =   3750
      TabIndex        =   10
      Top             =   630
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Left            =   3750
      TabIndex        =   9
      Top             =   420
      Width           =   825
   End
   Begin MSForms.Label lblFecha 
      Height          =   360
      Left            =   5280
      TabIndex        =   8
      ToolTipText     =   "Hora Actual"
      Top             =   30
      Width           =   1635
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "01/06/2007"
      Size            =   "2884;635"
      BorderColor     =   -2147483639
      BorderStyle     =   1
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label lblHora 
      Height          =   390
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Hora Actual"
      Top             =   60
      Width           =   1935
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "08:30:00 AM"
      Size            =   "3413;688"
      BorderColor     =   -2147483639
      BorderStyle     =   1
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ListBox lstDatosEmp 
      Height          =   1095
      Left            =   4680
      TabIndex        =   4
      Top             =   450
      Width           =   3465
      BackColor       =   10442041
      ForeColor       =   12648447
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "6112;1931"
      MatchEntry      =   0
      BorderColor     =   8421631
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      ForeColor       =   8421631
      BackColor       =   10442041
      VariousPropertyBits=   276824091
      Caption         =   "CODIGO"
      Size            =   "2778;582"
      BorderColor     =   8421631
      BorderStyle     =   1
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmRegAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Dim nFila As Integer
Private WithEvents mFTP As cFTP
Attribute mFTP.VB_VarHelpID = -1

Private Sub btnReporte_Click()
    Dim sede As String
    sede = CodSede
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    oReporte.Titulo = "REPORTE DE INCIDENCIAS DE ENTRADA Y SALIDA"
    oReporte.sede = DescripcionesdeCodigos("SEDES", sede, "NOMBRE")
    oReporte.Reporte = "Rep_InsidenciaEntSal.rpt"
    oReporte.sp_Rep_InsiEntSal Format(dtpIngreso.Value, "yyyy/mm/dd")
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo SERROR
    Dim SQL As String, I As Integer, codempleado As String, llave As String
    Dim rsSupervisor As MYSQL_RS, rsLlave As MYSQL_RS
    Dim supervisor As String, Clave As String
    Dim Permiso As Integer
        SQL = "select llave from rh_estacionestrabajo where codigo='" & lblSede.tag & "'"
        Set rsLlave = oConexion.EjecutaSelectRS(SQL)
        llave = rsLlave.Fields("llave")
        If llave = "S" Then
            supervisor = InputBox("Ingrese su Usuario de Supervisor", "NOVPeru")
            Clave = InputBox("Ingrese su Clave de Supervisor", "NOVPeru")
            SQL = "Select clave from 3cnuser where usuario_id='" & supervisor & "' and perfil_id='0006'"
            Set rsSupervisor = oConexion.EjecutaSelectRS(SQL)
            If rsSupervisor.RecordCount > 0 Then
                If Clave = rsSupervisor.Fields("CLAVE") Then
                    codempleado = cEmpleado(meCodigo)
                    If codempleado <> "" Then
                        If Not VerificaDiasVaca(codempleado) Then
                            Permiso = VerificaPermisos(codempleado)
                            If val(Permiso) = 0 Then
                                    RegistrarAsistencia codempleado, Format(lblFecha, "yyyy/mm/dd"), lblHora
                                    mshRegistro.row = nFila
                                    If nFila < 10 Then
                                        mshRegistro.TopRow = 1
                                    Else
                                        mshRegistro.TopRow = nFila - 5
                                    End If
                                    mshRegistro.Col = 0
                                    mshRegistro.ColSel = mshRegistro.Cols - 1
                                    meCodigo = "________"
                                    meCodigo.SetFocus
                            Else
                                MsgBox "Usted se encuentra de " & IIf(val(Permiso) = 3, "permiso", "licencia") & ". No se registrará su asistencia", vbInformation, "NOVPeru"
                                cmdAceptar.SetFocus
                            End If
                        Else
                            MsgBox "Usted se encuentra en Vacaciones. No se registrará su asistencia", vbInformation, "NOVPeru"
                            cmdAceptar.SetFocus
                        End If
                    End If
                Else
                    MsgBox "Clave incorrecta, por favor vuelva intentarlo", vbOKOnly + vbInformation, "NOVPeru"
                    cmdAceptar.SetFocus
                End If
            Else
                MsgBox "El usuario no existe o no es un personal de confianza" & vbNewLine & _
                       "vuelva intentarlo o consulte con el administrador del sistema", vbOKOnly + vbInformation, "NOVPeru"
                       cmdAceptar.SetFocus
            End If
        Else
            codempleado = cEmpleado(meCodigo)
            If codempleado <> "" Then
                If Not VerificaDiasVaca(codempleado) Then
                    Permiso = VerificaPermisos(codempleado)
                    If val(Permiso) = 0 Then
                            RegistrarAsistencia codempleado, Format(lblFecha, "yyyy/mm/dd"), lblHora
                            mshRegistro.row = nFila
                            If nFila < 10 Then
                                mshRegistro.TopRow = 1
                            Else
                                mshRegistro.TopRow = nFila - 5
                            End If
                            mshRegistro.Col = 0
                            mshRegistro.ColSel = mshRegistro.Cols - 1
                            meCodigo = "________"
                            meCodigo.SetFocus
                    Else
                        MsgBox "Usted se encuentra de " & IIf(val(Permiso) = 3, "permiso", "licencia") & ". No se registrará su asistencia", vbInformation, "NOVPeru"
                        cmdAceptar.SetFocus
                    End If
                Else
                    MsgBox "Usted se encuentra en Vacaciones. No se registrará su asistencia", vbInformation, "NOVPeru"
                    cmdAceptar.SetFocus
                End If
            End If
        End If
    
    Set rsSupervisor = Nothing
    Set rsLlave = Nothing
    
    Exit Sub
SERROR:
    Mensajes err.Description
End Sub

Function VerificaDiasVaca(Cod As String) As Boolean
    Dim SQ As String, RQ As MYSQL_RS
    Dim DiaIniV As Integer, DiaFinV As Integer
    Dim MesIniV As Integer, MesFinV As Integer
    Dim AnioIniV As String, AnioFinV As String
    SQ = "Select c.fec_Salida,c.fec_Regreso from empleado as a left join calendario as c " & _
         "on(c.codemp=a.codigo) where c.movemp='02' and " & _
         "concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & Year(Date) & Right("00" & Month(Date), 2) & "' and " & _
         "concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & Year(Date) & Right("00" & Month(Date), 2) & "' and " & _
         "gocehaber = 'N' and c.codemp = '" & Cod & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    If Not RQ.EOF() Then
        Do While Not RQ.EOF
            If CDate(lblFecha) >= Format(RQ.Fields("fec_salida"), "dd/mm/yyyy") Then
                If CDate(lblFecha) <= Format(RQ.Fields("fec_regreso"), "dd/mm/yyyy") Then
                    VerificaDiasVaca = True
                    Exit Do
                End If
            End If
            RQ.MoveNext
        Loop
    End If
    Set RQ = Nothing
End Function

Function VerificaPermisos(Cod As String) As Integer
    Dim SQ As String, RQ As MYSQL_RS
    Dim DiaIniV As Integer, DiaFinV As Integer
    Dim MesIniV As Integer, MesFinV As Integer
    Dim AnioIniV As String, AnioFinV As String
    VerificaPermisos = 0
    SQ = "Select c.fec_Salida,c.fec_Regreso,c.movemp from empleado as a left join calendario as c " & _
         "on(c.codemp=a.codigo) where c.movemp in ('03','05','07') and " & _
         "concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & Year(Date) & Right("00" & Month(Date), 2) & "' and " & _
         "concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & Year(Date) & Right("00" & Month(Date), 2) & "' and " & _
         "c.codemp = '" & Cod & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    If Not RQ.EOF() Then
        Do While Not RQ.EOF
            If CDate(lblFecha) >= Format(RQ.Fields("fec_salida"), "dd/mm/yyyy") Then
                If CDate(lblFecha) <= Format(RQ.Fields("fec_regreso"), "dd/mm/yyyy") Then
                    VerificaPermisos = RQ.Fields("movemp")
                    Exit Do
                End If
            End If
            RQ.MoveNext
        Loop
    End If
    Set RQ = Nothing
End Function

Function cEmpleado(ndi As String) As String
    Dim SQL As String
    Dim rscod As MYSQL_RS
    SQL = "Select codigo from empleado where numdocide='" & Right("00000000" & Trim(ndi), 8) & "'"
    Set rscod = oConexion.EjecutaSelectRS(SQL)
    If rscod.RecordCount = 0 Then
        MsgBox "El código ingresado no existe, inténtelo denuevo ", vbOKOnly + vbExclamation, "NOVPeru"
        meCodigo = "________"
        cEmpleado = ""
    Else
        cEmpleado = rscod.Fields("CODIGO")
    End If
    Set rscod = Nothing
End Function

Sub RegistrarAsistencia(CodEmp As String, fec As String, hor As String)
    On Error GoTo SERROR
    Dim SQL As String, I As Integer, dia As Integer, Foto As String
    Dim Tipo As String
    Dim HoraReg As Integer, HoraSis As Integer, DifH As Integer
    Dim MinReg As Integer, MinSis As Integer, DifM As Integer
    Dim rsES As MYSQL_RS, rsEmp As MYSQL_RS
    SQL = "Select tipo,hor from rh_entsalempleado where fecha='" & fec & "' and emp='" & CodEmp & "' order by hor desc limit 1"
    Set rsES = oConexion.EjecutaSelectRS(SQL)
    If rsES.RecordCount > 0 Then
        HoraReg = Left(Format(rsES.Fields("hor"), "hh:mm"), 2)
        HoraSis = Left(Format(lblHora, "hh:mm"), 2)
        MinReg = Right(Format(rsES.Fields("hor"), "hh:mm"), 2)
        MinSis = Right(Format(lblHora, "hh:mm"), 2)
        DifH = HoraSis - HoraReg
        DifM = MinSis - MinReg
        If Abs(val(DifM)) >= 3 Then
            If rsES.Fields("tipo") = "R" Then Tipo = "S"
            If rsES.Fields("tipo") = "E" Then Tipo = "S"
            If rsES.Fields("tipo") = "S" Then Tipo = "R"
        Else
            If Abs(val(DifH)) > 0 Then
                If rsES.Fields("tipo") = "R" Then Tipo = "S"
                If rsES.Fields("tipo") = "S" Then Tipo = "R"
                If rsES.Fields("tipo") = "E" Then Tipo = "S"
            Else
                MsgBox "No puede registrar " & IIf(rsES.Fields("tipo") = "R", "la Salida", "el Reingreso") & " hasta pasado 3 minutos", vbInformation, "NOVPeru"
                Exit Sub
            End If
        End If
    Else
        Tipo = "E"
        If ValidaHora(CodEmp) = False Then
            cmdAceptar.SetFocus
            Exit Sub
        End If
    End If
    
    SQL = "Select a.foto,a.Nombre1,a.nombre2,a.apepat,a.apemat,b.descrip,c.horlab,c.estTrabajo,a.categoria" & _
          " from (empleado as a left join cncargos as b on (a.codcargo=b.codigo)) left join contrato as c on (a.codigo=c.codemp)" & _
          " where a.codigo='" & CodEmp & "' and c.estado='AP'"
    Set rsEmp = oConexion.EjecutaSelectRS(SQL)
    If rsEmp.RecordCount < 1 Then
        MsgBox "No hay contrato habilitado para este trabajador", vbInformation + vbOKOnly, "NOVPeru"
        Exit Sub
    End If
    If rsEmp.Fields("categoria") = "01" Then
        MsgBox "Para Personal de Dirección, el registro es automático. No es necesario su Registro por este medio.", vbInformation + vbOKOnly, "NOVPeru"
        Exit Sub
    End If
    With lstDatosEmp
        .Clear
        .AddItem rsEmp.Fields("Nombre1") & " " & rsEmp.Fields("Nombre2") & " " & rsEmp.Fields("ApePat") & " " & rsEmp.Fields("ApeMat")
        .AddItem rsEmp.Fields("DESCRIP")
        .AddItem DescripcionesdeCodigos("SEDES", rsEmp.Fields("estTrabajo"), "NOMBRE")
        .AddItem rsEmp.Fields("horlab")
    End With
    SQL = "Insert into rh_entsalempleado (sede,emp,fecha,hor,tipo) values ('" & lblSede.tag & "','" & CodEmp & "','" & _
          fec & "','" & Format(hor, "HH:MM:SS") & "','" & Tipo & "')"
    oConexionMYSQL.Execute SQL
   
    CargaRegistrodeInsidencias CodEmp, fec, hor
    
    
    Exit Sub
SERROR:
    Mensajes err.Description
    
End Sub


Private Function CampoHorario(dia As Integer, H As String) As String
    Select Case dia
        Case 1: CampoHorario = "Lu" & H
        Case 2: CampoHorario = "Ma" & H
        Case 3: CampoHorario = "Mi" & H
        Case 4: CampoHorario = "Ju" & H
        Case 5: CampoHorario = "Vi" & H
        Case 6: CampoHorario = "Sa" & H
        Case 7: CampoHorario = "Do" & H
    End Select
End Function

Private Function ConectarFTP(servidor As String, User As String, pass As String) As Boolean
    Set mFTP = New cFTP
    mFTP.SetModeActive
    mFTP.SetTransferBinary
    If mFTP.OpenConnection(servidor, User, pass) Then
        ConectarFTP = True
    Else
        If mFTP.OpenConnection("192.168.1.2", User, pass) Then
            ConectarFTP = True
        Else
            ConectarFTP = False
        End If
   End If
End Function

Private Function GenerarTxt() As Boolean
    Dim I As Integer
    Dim NomArchivo As String
    Dim filetemp As Integer
    Dim separador As String
    Dim Rs As MYSQL_RS
    Dim aux As Boolean
    Dim SQL As String
    Dim RES As Integer
    Dim Cont As Integer
    separador = "|"
    filetemp = FreeFile()
    SQL = "Select * from rh_entsalempleado where envio='N' order by fecha,emp"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    NomArchivo = Rep_Documents & "\RRHH\" & "S" & Trim(lblSede.tag) & "_" & Replace(Format(Date, "yyyy/mm/dd"), "/", "")
    On Error GoTo ErrorAbrir
    Open NomArchivo & ".txt" For Output As #filetemp
    Do While Not (Rs.EOF)
        Print #filetemp, Rs.Fields("sede") & separador & _
              Rs.Fields("emp") & separador & _
              Rs.Fields("fecha") & separador & _
              Rs.Fields("hor") & separador & _
              Rs.Fields("tipo") & separador
        Cont = Cont + 1
        Rs.MoveNext
    Loop
    Close filetemp
    SQL = "Update rh_entsalempleado set envio='S'" & _
          " where envio='N'"
    oConexionMYSQL.Execute SQL
    GenerarTxt = True
    MsgBox "La Generación de archivos se realizó satisfactoriamente." & vbNewLine & Cont & " Registro(s) generado(s)", vbInformation, gsNomSW
    Set Rs = Nothing
ErrorAbrir:
    If err.Number = 76 Then
        RES = MsgBox("No se encuentra la ruta: " & Rep_Documents & "\RRHH" & vbNewLine & vbNewLine & ",Desea crearla?", vbQuestion + vbYesNo, gsNomSW)
        If RES = 6 Then
            MkDir (Rep_Documents & "\RRHH\")
            Open NomArchivo & ".txt" For Output As #filetemp
            Resume Next
        Else
            GenerarTxt = False
        End If
    End If
End Function

Private Sub cmdingresar_Click()
    lblSede.Caption = LstSede.List(LstSede.ListIndex, 0)
    lblSede.tag = LstSede.List(LstSede.ListIndex, 1)
    If LstSede.ListIndex > 0 Then
        If MsgBox("Usted ha seleccionado la sede " & lblSede.Caption & ". ¿ESTA SEGURO QUE DESEA CONTINUAR?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
            frmlistado.Visible = False
            Me.Width = 8325
            Me.Height = 6300
        End If
    Else
        MsgBox "SELECCIONE CUIDADOSAMENTE LA SEDE EN LA QUE SE ENCUENTRA," & vbNewLine & _
               "DE LO CONTRARIO SU ASISTENCIA SERA INCORRECTA", vbCritical, "NOVPeru"
    End If
End Sub

Private Sub Form_Activate()
    lblFecha = Format(Date, "dd/mm/yyyy")
    dtpIngreso.Value = Format(Date, "dd/mm/yyyy")
    lblHora.Caption = Format(CStr(Time()))
    CargarAsistencia Format(lblFecha, "yyyy/mm/dd")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim SQL As String
    Dim rsSede As MYSQL_RS
    Me.Width = 4080
    Me.Height = 6135
    Me.Top = 0
    Me.Left = 0
    Timer1.Enabled = True
    Sedes
    Me.BorderStyle = vbFixedToolWindow
End Sub

Sub Sedes()
    Dim RQ As MYSQL_RS
    Dim SQ As String, I As Integer
    SQ = "SELECT * from rh_estacionestrabajo order by codigo"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    LstSede.Clear
    I = 0
    If Not RQ.EOF() Then
        Do While Not RQ.EOF()
            LstSede.AddItem Trim(RQ.Fields("nombre"))
            LstSede.List(I, 1) = Trim(RQ.Fields("codigo"))
            LstSede.List(I, 2) = Trim(RQ.Fields("tipo"))
            I = I + 1
            RQ.MoveNext
        Loop
        LstSede.ListIndex = 0
    End If
    Set RQ = Nothing
End Sub

Private Sub imgLlave_DblClick()
    Dim SQL As String, I As Integer, resp As Integer, llave As String
    Dim rsSupervisor As MYSQL_RS, rsLlave As MYSQL_RS
    Dim supervisor As String, Clave As String
    supervisor = InputBox("Ingrese su Usuario de Supervisor", "NOVPeru")
    Clave = InputBox("Ingrese su Clave de Supervisor", "NOVPeru")
    SQL = "Select clave from 3cnuser where usuario_id='" & supervisor & "' and perfil_id='0006'"
    Set rsSupervisor = oConexion.EjecutaSelectRS(SQL)
    If rsSupervisor.RecordCount > 0 Then
        If Clave = rsSupervisor.Fields("CLAVE") Then
                SQL = "Select llave from rh_estacionestrabajo where codigo='" & lblSede.tag & "'"
                Set rsLlave = oConexion.EjecutaSelectRS(SQL)
                llave = rsLlave.Fields("llave")
                If llave = "N" Then
                    resp = MsgBox("¿Desea habilitar el control de registro manual de asistencia?", vbYesNo + vbQuestion, "NOVPeru")
                    If resp = vbYes Then
                        SQL = "update rh_estacionestrabajo set llave='S' where codigo='" & lblSede.tag & "'"
                    End If
                Else
                    resp = MsgBox("¿Desea deshabilitar el control de registro manual de asistencia?", vbYesNo + vbQuestion, "NOVPeru")
                    If resp = vbYes Then
                        SQL = "update rh_estacionestrabajo set llave='N' where codigo='" & lblSede.tag & "'"
                    End If
                End If
                oConexionMYSQL.Execute SQL
        Else
            MsgBox "Clave incorrecta, por favor vuelva intentarlo", vbOKOnly + vbInformation, "NOVPeru"
            cmdAceptar.SetFocus
        End If
    Else
        MsgBox "El usuario no existe o no es un personal de confianza" & vbNewLine & _
               "vuelva intentarlo o consulte con el administrador del sistema", vbOKOnly + vbInformation, "NOVPeru"
               cmdAceptar.SetFocus
    End If
    Set rsSupervisor = Nothing
    Set rsLlave = Nothing
End Sub

Private Sub meCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If meCodigo <> "________" Then
            cmdAceptar_Click
        End If
    End If
End Sub

Function ValidaHora(Optional CodEmp As String) As Boolean
    On Error GoTo SERROR
    Dim SQL As String, dia As Integer, hora As String
    Dim rsHor As MYSQL_RS
    ValidaHora = False
    SQL = "Select b.LuE,b.MaE,b.MiE,b.JuE,b.ViE,b.SaE,b.DoE,TolEntrada from contrato as a left join rh_horarios as b on (a.horlab=b.nombre) where a.codemp='" & CodEmp & "' and a.estado='AP'"
    Set rsHor = oConexion.EjecutaSelectRS(SQL)
    If Not (rsHor.EOF And rsHor.BOF) Then
        dia = Weekday(Date, vbMonday) - 1
        hora = Format(TimeValue(rsHor.Fields(dia)) - TimeValue("00:" & Right("00" & Trim(IIf(val(rsHor.Fields(7)) = 0, 5, val(rsHor.Fields(7)))), 2) & ":00"), "hh:mm:ss")
        If Hour(Time) > Hour(hora) Then ValidaHora = True
        If Hour(Time) = Hour(hora) Then
            If Minute(lblHora) >= Minute(hora) Then
                ValidaHora = True
            End If
        End If
    End If
    
    If ValidaHora = False Then
        MsgBox "La Asistencia se registrará a partir de las " & hora, vbInformation, "NOVPeru"
    End If
    
    Exit Function
SERROR:
    Mensajes err.Description
End Function

Private Sub mshRegistro_RowColChange()
    Dim SQL As String
    Dim rsEmp As MYSQL_RS
    If mshRegistro.row > 0 Then
        SQL = "Select a.foto,a.Nombre1,a.nombre2,a.apepat,a.apemat,b.descrip,c.horlab,c.estTrabajo,d.descrip as categ" & _
              " from ((empleado as a left join cncargos as b on (a.codcargo=b.codigo))" & _
              " left join contrato as c on (a.codigo=c.codemp))" & _
              " left join rh_categoria as d on (a.categoria=d.codigo)" & _
              " where a.codigo='" & mshRegistro.TextMatrix(mshRegistro.row, 1) & _
              "' and c.estado='AP'"
        Set rsEmp = oConexion.EjecutaSelectRS(SQL)
        With lstDatosEmp
            .Clear
            .AddItem rsEmp.Fields("Nombre1") & " " & rsEmp.Fields("Nombre2") & " " & rsEmp.Fields("ApePat") & " " & rsEmp.Fields("ApeMat")
            .AddItem rsEmp.Fields("DESCRIP")
            .AddItem DescripcionesdeCodigos("SEDES", rsEmp.Fields("estTrabajo"), "NOMBRE")
            .AddItem rsEmp.Fields("horlab")
            .AddItem rsEmp.Fields("categ")
        End With
       
    End If
End Sub

Private Sub Timer1_Timer()
    I = I + 1
    lblHora.Caption = Format(CStr(Time()))
    If I = 60 Then
        I = 0
        CargarAsistencia Format(lblFecha, "yyyy/mm/dd")
        meCodigo.SetFocus
    End If
End Sub

Sub CargaRegistrodeInsidencias(Cod As String, fec As String, hor As String)
    Dim SQL As String, I As Integer, fila As Integer
    Dim rs_Registro As MYSQL_RS
    Dim rs_Emp As MYSQL_RS
    fila = 0
    SQL = "Select count(*) as insi from rh_entsalempleado where emp='" & Cod & "' and fecha='" & fec & "'"
    Set rs_Emp = oConexion.EjecutaSelectRS(SQL)
    With mshRegistro
        If rs_Emp.RecordCount > 0 Then
            If 3 + rs_Emp.Fields("insi") > .Cols Then
                .Cols = 3 + rs_Emp.Fields("insi")
                .ColAlignment = flexAlignCenterCenter
            End If
            If .Cols < 5 Then
                .ColWidth(3) = 1250
               .TextMatrix(0, 3) = "ENTRADA"
               .ColAlignmentFixed(3) = flexAlignCenterCenter
            Else
                .ColWidth(2 + rs_Emp.Fields("insi")) = 1250
                .ColAlignmentFixed(2 + rs_Emp.Fields("insi")) = flexAlignCenterCenter
               .TextMatrix(0, 2 + rs_Emp.Fields("insi")) = IIf(.TextMatrix(0, rs_Emp.Fields("insi") + 1) = "ENTRADA", "SALIDA", "ENTRADA")
            End If
            For I = 0 To .Rows - 1
                If .TextMatrix(I, 1) = Cod Then
                    .TextMatrix(I, rs_Emp.Fields("insi") + 2) = hor
                    fila = I
                    nFila = I
                End If
            Next
            If fila = 0 Then
                .Rows = .Rows + 1
                .FixedRows = 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = Cod
                .TextMatrix(.Rows - 1, rs_Emp.Fields("insi") + 2) = hor
            End If
        End If
    End With
    Set rs_Registro = Nothing
    Set rs_Emp = Nothing
End Sub

Sub CargarAsistencia(fec As String)
    Dim I As Integer, columna As Integer, fila As Integer
    On Error GoTo NADA
    Dim rscol As MYSQL_RS
    Dim rsfilas As MYSQL_RS
    Dim SQL As String, cemp As String
    SQL = "SELECT MAX(A.INSIDENCIAS) AS CUENTA FROM" & _
          " (SELECT COUNT(*) AS INSIDENCIAS FROM rh_entsalempleado" & _
          " WHERE FECHA='" & fec & "' GROUP BY FECHA,EMP ) AS A"
    Set rscol = oConexion.EjecutaSelectRS(SQL)
    mshRegistro.Rows = 1
    mshRegistro.Cols = 3
    mshRegistro.TextMatrix(0, 0) = "N°"
    mshRegistro.ColWidth(0) = 300
    mshRegistro.ColAlignment(0) = flexAlignCenterCenter
    mshRegistro.TextMatrix(0, 1) = "CODIGO"
    mshRegistro.ColWidth(1) = 1250
    mshRegistro.ColAlignment(1) = flexAlignCenterCenter
    mshRegistro.TextMatrix(0, 2) = "ORDEN"
    mshRegistro.ColWidth(2) = 0
    mshRegistro.ColAlignmentFixed(0) = flexAlignCenterCenter
    mshRegistro.ColAlignmentFixed(1) = flexAlignCenterCenter
    If Not IsNull(rscol.Fields("cuenta")) Then
        mshRegistro.Cols = 3 + rscol.Fields("cuenta")
        For I = 0 To rscol.Fields("cuenta") - 1
            If I = 0 Then
                mshRegistro.TextMatrix(0, 3 + I) = "ENTRADA"
                mshRegistro.ColWidth(3 + I) = 1250
                mshRegistro.ColAlignmentFixed(3 + I) = flexAlignCenterCenter
                mshRegistro.ColAlignment(3 + I) = flexAlignCenterCenter
            Else
                mshRegistro.ColAlignment(3 + I) = flexAlignCenterCenter
                mshRegistro.ColAlignmentFixed(3 + I) = flexAlignCenterCenter
                mshRegistro.ColWidth(3 + I) = 1250
                mshRegistro.TextMatrix(0, 3 + I) = IIf(mshRegistro.TextMatrix(0, 2 + I) = "ENTRADA", "SALIDA", "ENTRADA")
            End If
        Next
       
    End If
    SQL = "SELECT * FROM rh_entsalempleado" & _
         " WHERE FECHA='" & fec & "'AND ENVIO = 'N' order by emp,hor"
    Set rsfilas = oConexion.EjecutaSelectRS(SQL)
    cemp = ""
    columna = 3
    fila = 0
    Do While Not (rsfilas.EOF)
        If cemp <> rsfilas.Fields("emp") Then
            cemp = rsfilas.Fields("emp")
            mshRegistro.Rows = mshRegistro.Rows + 1
            mshRegistro.FixedRows = 1
            fila = fila + 1
            columna = 3
            mshRegistro.TextMatrix(fila, 0) = fila
            mshRegistro.TextMatrix(fila, 1) = rsfilas.Fields("emp")
            mshRegistro.TextMatrix(fila, 2) = Format(rsfilas.Fields("hor"), "HH:MM:SS")
            mshRegistro.TextMatrix(fila, columna) = Format(rsfilas.Fields("hor"))
        Else
            columna = columna + 1
            mshRegistro.TextMatrix(fila, columna) = Format(rsfilas.Fields("hor"), "HH:MM:SS")
            mshRegistro.TextMatrix(fila, columna) = Format(rsfilas.Fields("hor"))
        End If
        rsfilas.MoveNext
    Loop
    If fila > 1 Then
        mshRegistro.Col = 2
        mshRegistro.Sort = 7
        For I = 1 To mshRegistro.Rows - 1
            mshRegistro.TextMatrix(I, 0) = I
        Next
        mshRegistro.Col = 0
        mshRegistro.row = mshRegistro.Rows - 1
        DoEvents
        mshRegistro.ColSel = mshRegistro.Cols - 1
        mshRegistro.Refresh
        Call keybd_event(vbKeyHome, 0, 0, 0)
        mshRegistro_RowColChange
    End If
    If mshRegistro.Rows > 12 Then mshRegistro.TopRow = mshRegistro.Rows - 5
    Set rscol = Nothing
    Set rsfilas = Nothing
Exit Sub
NADA:
    Exit Sub
End Sub
