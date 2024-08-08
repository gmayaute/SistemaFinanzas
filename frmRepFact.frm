VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmRepFact 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Facturas al Exterior"
   ClientHeight    =   1185
   ClientLeft      =   6735
   ClientTop       =   8580
   ClientWidth     =   6255
   Icon            =   "frmRepFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   6255
   Begin MSComDlg.CommonDialog CMD 
      Left            =   5580
      Top             =   -90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtmesal 
      Height          =   315
      Left            =   975
      MaxLength       =   6
      TabIndex        =   3
      Top             =   795
      Width           =   795
   End
   Begin VB.TextBox txtmesdel 
      Height          =   315
      Left            =   960
      MaxLength       =   6
      TabIndex        =   2
      Top             =   427
      Width           =   795
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
      Left            =   5010
      TabIndex        =   5
      Top             =   442
      Width           =   1140
   End
   Begin VB.TextBox txtfactal 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3030
      TabIndex        =   1
      Top             =   60
      Width           =   1845
   End
   Begin VB.TextBox txtfactdel 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   975
      TabIndex        =   0
      Top             =   60
      Width           =   1845
   End
   Begin Proyecto1.chameleonButton cmdsalir 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5565
      TabIndex        =   6
      ToolTipText     =   "Salir"
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
      MICON           =   "frmRepFact.frx":030A
      PICN            =   "frmRepFact.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdreporte 
      Height          =   375
      Left            =   5070
      TabIndex        =   7
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
      MICON           =   "frmRepFact.frx":06EC
      PICN            =   "frmRepFact.frx":0708
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
      Caption         =   "Mes Al"
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
      Left            =   45
      TabIndex        =   11
      Top             =   810
      Width           =   855
   End
   Begin VB.Label LblDes 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   1
      Left            =   1830
      TabIndex        =   10
      Top             =   795
      Width           =   3045
   End
   Begin VB.Label LblDes 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   0
      Left            =   1830
      TabIndex        =   9
      Top             =   420
      Width           =   3045
   End
   Begin VB.Label Lbl 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mes Del"
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
      Left            =   45
      TabIndex        =   8
      Top             =   442
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2850
      X2              =   3060
      Y1              =   210
      Y2              =   225
   End
   Begin VB.Label Lbl 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factura"
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
      Left            =   45
      TabIndex        =   4
      Top             =   75
      Width           =   855
   End
End
Attribute VB_Name = "frmRepFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConsulta As FrmConsultas

Private Sub cmdreporte_Click()
    Set oReporte = New clsReporte
    oReporte.Reporte = "Rep_FacturasExt.rpt"
    oReporte.Titulo = "REPORTE DETALLE DE FACTURAS"
    oReporte.empresa = "NATIONAL OILWELL VARCO PERU S.R.L"
    oReporte.sp_Rep_FacturasExt Mid(Trim(txtfactdel), 7, Len(Trim(txtfactdel))), Mid(Trim(txtfactal), 7, Len(Trim(txtfactal))), Trim(txtmesdel), Trim(txtmesal), 0, CmD
    If chkExcel.Value = False Then
        oReporte.Reporte = "Rep_FacturasExtCli.rpt"
        oReporte.Titulo = "REPORTE CLIENTES"
        oReporte.sp_Rep_FacturasExt Mid(Trim(txtfactdel), 7, Len(Trim(txtfactdel))), Mid(Trim(txtfactal), 7, Len(Trim(txtfactal))), Trim(txtmesdel), Trim(txtmesal), 1, CmD
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Set oConsulta = New FrmConsultas
End Sub

Private Sub txtfactal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtfactal = BuscaNumFact(Right("000000000" & Trim(txtfactal), 9))
        txtmesdel.SetFocus
    End If
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 3
            .pCol = 0: .pAnchoCol = 1700
            .pCol = 1: .pAnchoCol = 3800
            .pCol = 2: .pAnchoCol = 900
            .pTitulo = "Documentos del Mes " & NombreMes(strMesSistema, False)
            .pForm = FORM_REPFACT
            .pCaso = LABEL_FACTURA
            .Show
        End With
    End If
End Sub

Private Sub txtfactal_LostFocus()
    txtfactal = BuscaNumFact(Right("000000000" & Trim(txtfactal), 9))
    txtmesdel.SetFocus
End Sub

Private Sub txtfactdel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtfactdel = BuscaNumFact(Right("000000000" & Trim(txtfactdel), 9))
        txtfactal.SetFocus
    End If
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 3
            .pCol = 0: .pAnchoCol = 1700
            .pCol = 1: .pAnchoCol = 3800
            .pCol = 2: .pAnchoCol = 900
            .pTitulo = "Documentos del Mes " & NombreMes(strMesSistema, False)
            .pForm = FORM_REPFACT
            .pCaso = LABEL_DOCxFACT
            .Show
        End With
    End If
End Sub

Function BuscaNumFact(NumCorrel As String) As String
    Dim SQL As String
    Dim RQ As MYSQL_RS
    SQL = "select serie from documento_contables where correl = '" & NumCorrel & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF Then
        BuscaNumFact = RQ.Fields("serie") & "-" & NumCorrel
    End If
    Set RQ = Nothing
End Function

Private Sub txtfactdel_LostFocus()
    txtfactdel = BuscaNumFact(Right("000000000" & Trim(txtfactdel), 9))
    txtfactal.SetFocus
End Sub

Private Sub txtmesal_Change()
    LblDes(1).Caption = NombreMes(Right("00" & Trim(txtmesal), 2), False)
End Sub

Private Sub txtmesdel_Change()
    LblDes(0).Caption = NombreMes(Right("00" & Trim(txtmesdel), 2), False)
End Sub

Private Sub txtmesdel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtmesal.SetFocus
End Sub
