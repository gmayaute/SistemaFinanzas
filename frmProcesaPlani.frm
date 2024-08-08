VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmProcesaPlani 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesa Planilla"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmProcesaPlani.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin Proyecto1.chameleonButton cmdAceptar 
      Height          =   375
      Left            =   1530
      TabIndex        =   6
      Top             =   930
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
      MICON           =   "frmProcesaPlani.frx":030A
      PICN            =   "frmProcesaPlani.frx":0326
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
      BackColor       =   &H009F5539&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Lbl 
      BackColor       =   &H009F5539&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   495
   End
   Begin VB.Label Lbl 
      BackColor       =   &H009F5539&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   2310
      TabIndex        =   3
      Top             =   60
      Width           =   765
   End
   Begin MSForms.ComboBox cboMes 
      Height          =   315
      Left            =   540
      TabIndex        =   2
      Top             =   30
      Width           =   1695
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2990;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboProceso 
      Height          =   315
      Left            =   1830
      TabIndex        =   1
      Top             =   480
      Width           =   2355
      VariousPropertyBits=   746604569
      DisplayStyle    =   7
      Size            =   "4154;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboMon 
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   60
      Width           =   1545
      VariousPropertyBits=   746604569
      DisplayStyle    =   7
      Size            =   "2725;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmProcesaPlani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim Rs As MYSQL_RS

Private Sub cmdAceptar_Click()
    Dim rpta As Integer
    SQL = "Insert into pl_planiproc (anomes,proceso,mon,contab,fecha ) values " & _
        "('" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "','" & _
        cboProceso.List(cboProceso.ListIndex, 2) & "','" & _
        Left(cboMon.Text, 1) & "','N','" & _
        Format(Date, "yyyy/mm/dd") & "')"
    rpta = MsgBox("Esta seguro de procesar esta planilla," & vbNewLine & _
                 "ya no podrá modificar el tareo", vbYesNo + vbQuestion, "NOVPeru")
    If rpta = vbYes Then
        oConexionMYSQL.Execute SQL
        ActualizarDocRRHH strAnoSistema & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), Left(cboMon.Text, 1), Format(DtpFecha, "yyyy/mm/dd")
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    LlenarMesP cboMes
    LlenarMonedaP cboMon
    LlenarProcesos cboProceso
End Sub
