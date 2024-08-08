VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmInterfazTw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TeleWiese"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "frmInterfazTw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7290
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   3105
      Left            =   0
      TabIndex        =   4
      Top             =   -60
      Width           =   7275
      Begin VB.Frame Frame2 
         BackColor       =   &H009F5539&
         Caption         =   "Opciones de Orden"
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
         Height          =   1365
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   7035
         Begin VB.TextBox txtCodOpcional 
            Height          =   315
            Left            =   2250
            MaxLength       =   11
            TabIndex        =   14
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkCodOpcional 
            BackColor       =   &H009F5539&
            Caption         =   "Orden con Código:"
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
            Height          =   225
            Left            =   180
            TabIndex        =   13
            Top             =   285
            Width           =   2025
         End
         Begin VB.CheckBox chkDocOpcional 
            BackColor       =   &H009F5539&
            Caption         =   "Orden con Doc.:"
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
            Height          =   225
            Left            =   180
            TabIndex        =   12
            Top             =   720
            Width           =   1965
         End
         Begin VB.TextBox txtDocOpcional 
            Height          =   315
            Left            =   2250
            MaxLength       =   15
            TabIndex        =   11
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkPreparar 
            BackColor       =   &H009F5539&
            Caption         =   "Preparar Orden"
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
            Height          =   225
            Left            =   5160
            TabIndex        =   10
            Top             =   930
            Width           =   1815
         End
         Begin VB.CheckBox chkPagoUnico 
            BackColor       =   &H009F5539&
            Caption         =   "Pago Unico"
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
            Height          =   225
            Left            =   5160
            TabIndex        =   9
            Top             =   630
            Width           =   1575
         End
         Begin MSForms.ComboBox cmbTipoOrden 
            Height          =   315
            Left            =   5130
            TabIndex        =   16
            Top             =   240
            Width           =   1635
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2884;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label15 
            BackColor       =   &H009F5539&
            Caption         =   "Tipo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   4680
            TabIndex        =   15
            Top             =   300
            Width           =   375
         End
      End
      Begin Proyecto1.chameleonButton BtnGenerar 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   "Eliminar"
         Top             =   2520
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Generar Tw"
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
         MICON           =   "frmInterfazTw.frx":2372
         PICN            =   "frmInterfazTw.frx":238E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnSalir 
         Height          =   375
         Left            =   3690
         TabIndex        =   18
         Top             =   2520
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Salir"
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
         MICON           =   "frmInterfazTw.frx":4710
         PICN            =   "frmInterfazTw.frx":472C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblNumDocTw 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2910
         TabIndex        =   17
         Top             =   585
         Width           =   4245
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Auxiliar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   3840
         TabIndex        =   6
         Top             =   210
         Width           =   1095
      End
      Begin MSForms.TextBox txtCodigo 
         Height          =   315
         Left            =   5010
         TabIndex        =   1
         Top             =   210
         Width           =   2145
         VariousPropertyBits=   746604571
         MaxLength       =   11
         Size            =   "3784;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro. Orden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   570
         Width           =   1095
      End
      Begin MSForms.TextBox txtOrden 
         Height          =   345
         Left            =   1290
         TabIndex        =   2
         Top             =   570
         Width           =   1545
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   4
         Size            =   "2725;609"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cboauxiliar 
         Height          =   315
         Left            =   1290
         TabIndex        =   0
         Top             =   210
         Width           =   2445
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   128
         DisplayStyle    =   7
         Size            =   "4313;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmInterfazTw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConsulta As FrmConsultas
Private Sub BtnGenerar_Click()
    Dim TipOrden As String
    If txtOrden <> Empty And txtCodigo <> Empty And cboauxiliar.List(cboauxiliar.ListIndex, 0) <> "Seleccionar..." Then
        If cmbTipoOrden.ListCount > 1 Then
            TipOrden = cmbTipoOrden.List(cmbTipoOrden.ListIndex, 1)
        Else
            TipOrden = "0"
        End If
        If GeneraTxt(txtOrden, cboauxiliar.List(cboauxiliar.ListIndex, 1), txtCodOpcional, txtDocOpcional, chkPreparar.Value, TipOrden, chkPagoUnico.Value) Then
            txtCodigo = Empty
            txtOrden = Empty
            txtCodOpcional = Empty
            txtDocOpcional = Empty
            LLenarCbo cboauxiliar
            CargarTipoOrden
            chkCodOpcional.Value = False
            chkDocOpcional.Value = False
            chkPagoUnico.Value = False
            chkPreparar.Value = False
            cboauxiliar.SetFocus
        End If
    Else
        If cboauxiliar.List(cboauxiliar.ListIndex, 0) = "Seleccionar..." Then cboauxiliar.SetFocus: Exit Sub
        If txtCodigo = Empty Then txtCodigo.SetFocus: Exit Sub
        If txtOrden = Empty Then txtOrden.SetFocus: Exit Sub
    End If
End Sub

Private Sub CargarTipoOrden()
    With cmbTipoOrden
        .Clear
        .AddItem "Selecionar"
        .List(0, 1) = "0"
        .AddItem "Planilla"
        .List(1, 1) = "1"
        .AddItem "Factura"
        .List(2, 1) = "2"
        .ListIndex = 0
    End With
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    LLenarCbo cboauxiliar
    CargarTipoOrden
    Set oConsulta = New FrmConsultas
End Sub
Private Sub LLenarCbo(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rscbo As MYSQL_RS
    cbo.Clear
    SQL = "Select tip_linea,descrip From CNTABLAS where codtab='1' and " & _
            "(tip_linea='3' or tip_linea='2' or tip_linea='5') Order By 1"
    Set rscbo = oConexion.EjecutaSelectRS(SQL)
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = "00"
    Do While Not rscbo.EOF
        cbo.AddItem rscbo.Fields(1)
        cbo.List(cbo.ListCount - 1, 1) = rscbo.Fields(0)
        rscbo.MoveNext
    Loop
    cbo.ListIndex = 0
    rscbo.CloseRecordset
    Set rscbo = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set oConsulta = Nothing
End Sub

Private Sub txtCodigo_GotFocus()
    mark1 txtCodigo
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim SQL As String
    Dim RS As MYSQL_RS
    If KeyCode.Value = vbKeyF1 Then
        If cboauxiliar.List(cboauxiliar.ListIndex, 0) <> "Seleccionar..." Then
            strTipoAuxiliar = cboauxiliar.List(cboauxiliar.ListIndex, 1)
            With oConsulta
                .pCols = 4
                .pCol = 0: .pAnchoCol = 1500
                .pCol = 1: .pAnchoCol = 3000
                .pCol = 2: .pAnchoCol = 0
                .pCol = 3: .pAnchoCol = 0
                .pTitulo = "Codigos" & DescripcionesdeCodigos("CNAUXIL", cboauxiliar.List(cboauxiliar.ListIndex, 1))
                .pForm = FORM_INTERFAZTW
                .pCaso = Label_Descrip_Auxil
                .Show
            End With
            Else
            cboauxiliar.SetFocus
        End If
    End If
    If KeyCode = 13 Then
        If cboauxiliar.List(cboauxiliar.ListIndex, 0) <> "Seleccionar..." Then
            txtCodigo = Right("00000000000" & Trim(txtCodigo), 11)
            SQL = "AUXIL where Auxiliar='" & cboauxiliar.List(cboauxiliar.ListIndex, 1) & "' " & _
                  " and codigo = '" & txtCodigo & "'"
            Set RS = oConexion.EjecutaSelect(SQL)
            If Not RS.EOF Then
               txtOrden.SetFocus
            Else
            MsgBox "No se encuentra el codigo ingresado", vbInformation, gsNomSW: txtCodigo = Empty: txtCodigo.SetFocus: Exit Sub
            End If
        Else
        cboauxiliar.SetFocus
        End If
    End If
End Sub

Private Sub txtOrden_GotFocus()
    mark1 txtOrden
End Sub

Private Sub txtOrden_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = vbKeyF1 Then
        If txtCodigo <> Empty And cboauxiliar.List(cboauxiliar.ListIndex, 0) <> "Seleccionar..." Then
            With oConsulta
                .pCols = 5
                .pCol = 0: .pAnchoCol = 800
                .pCol = 1: .pAnchoCol = 800
                .pCol = 2: .pAnchoCol = 500
                .pTitulo = "Ordenes-TeleWiese"
                .pForm = FORM_INTERFAZTW
                .pCaso = LABEL_ORDENTW
                .Show
            End With
            Else
            cboauxiliar.SetFocus
        End If
    End If
     If KeyCode = 13 Then
        If txtOrden <> Empty Then
            txtOrden = Right("0000" & txtOrden, 4)
        End If
    End If
End Sub
