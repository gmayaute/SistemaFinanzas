VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmModLetraRen 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Letras Renovadas"
   ClientHeight    =   1815
   ClientLeft      =   9480
   ClientTop       =   11790
   ClientWidth     =   5820
   Icon            =   "frmModLetraRen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5820
   Begin MSMask.MaskEdBox meFecGiro 
      Height          =   300
      Left            =   4170
      TabIndex        =   0
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meFecvcto 
      Height          =   300
      Left            =   4170
      TabIndex        =   2
      Top             =   480
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Proyecto1.chameleonButton BtnEliminar 
      Height          =   345
      Left            =   1230
      TabIndex        =   11
      ToolTipText     =   "Eliminar"
      Top             =   1380
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Eliminar"
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
      MICON           =   "frmModLetraRen.frx":014A
      PICN            =   "frmModLetraRen.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnGrabar 
      Height          =   345
      Left            =   75
      TabIndex        =   5
      ToolTipText     =   "Guardar"
      Top             =   1380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Grabar"
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
      MICON           =   "frmModLetraRen.frx":05A8
      PICN            =   "frmModLetraRen.frx":05C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton BtnSalir 
      Height          =   345
      Left            =   5250
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   1365
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
      MICON           =   "frmModLetraRen.frx":0A06
      PICN            =   "frmModLetraRen.frx":0A22
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox txtinteresbco 
      Height          =   330
      Left            =   4170
      TabIndex        =   4
      Top             =   885
      Width           =   1575
      VariousPropertyBits=   746604571
      MaxLength       =   10
      Size            =   "2778;582"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interés Bco."
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
      Height          =   300
      Left            =   2670
      TabIndex        =   14
      Top             =   900
      Width           =   1470
   End
   Begin MSForms.TextBox txtinteres 
      Height          =   330
      Left            =   1050
      TabIndex        =   3
      Top             =   885
      Width           =   1575
      VariousPropertyBits=   746604571
      MaxLength       =   10
      Size            =   "2778;582"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interés"
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
      Height          =   300
      Left            =   30
      TabIndex        =   13
      Top             =   900
      Width           =   990
   End
   Begin MSForms.Label letra 
      Height          =   300
      Left            =   1050
      TabIndex        =   10
      Top             =   75
      Width           =   1575
      ForeColor       =   128
      BackColor       =   12632256
      Size            =   "2778;529"
      BorderColor     =   64
      SpecialEffect   =   2
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label9 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Letra:"
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
      Height          =   300
      Left            =   30
      TabIndex        =   9
      Top             =   75
      Width           =   990
   End
   Begin VB.Label Label1 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha de Giro:"
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
      Height          =   300
      Left            =   2670
      TabIndex        =   8
      Top             =   90
      Width           =   1470
   End
   Begin VB.Label Label3 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha de Vcto.:"
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
      Height          =   300
      Left            =   2670
      TabIndex        =   7
      Top             =   480
      Width           =   1470
   End
   Begin VB.Label Label6 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CobBanco"
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
      Height          =   300
      Left            =   30
      TabIndex        =   6
      Top             =   480
      Width           =   990
   End
   Begin MSForms.TextBox txtCodBco 
      Height          =   330
      Left            =   1050
      TabIndex        =   1
      Top             =   465
      Width           =   1575
      VariousPropertyBits=   746604571
      MaxLength       =   10
      Size            =   "2778;582"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmModLetraRen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumND As String

Private Sub btnEliminar_Click()
Dim SQL As String
    If MsgBox("¿Seguro que desea eliminar la Letra " & letra & "?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
        SQL = "update letra set codestado = 'EL' where numero = '" & letra & "'"
        oConexionMYSQL.Execute SQL
        frmLetras.CargarLetrasRenovadas frmLetras.mshLetras.TextMatrix(frmLetras.mshLetras.row, 1)
        Unload Me
    End If
End Sub

Private Sub btnGrabar_Click()
Dim SQL As String
Dim FlgInt As Boolean
    If CDbl(txtinteres) <> CDbl(DevInteresLetra(letra)) Then
        FlgInt = False
        If MsgBox("¿Desea que el Monto de Interés se modifique también en la Nota de Débito N° ? " & NumND, vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
            FlgInt = True
        End If
    End If
    SQL = "update letra set fecgiro= '" & meFecGiro & "',fecvcto='" & meFecvcto & "', " & _
          "codbco='" & txtCodBco & "',codestado = 'MO',interesbco=" & CDbl(txtinteresbco) & ",interes=" & CDbl(txtinteres) & " " & _
          "where numero = '" & Trim(letra) & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
    
    If FlgInt = True Then
        SQL = "update documento_contables set total= " & CDbl(txtinteres) & ",subtotal='" & CDbl(FormatNumber(txtinteres / 1.18, 2)) & "', " & _
              "igv='" & CDbl(txtinteres - FormatNumber(txtinteres / 1.18, 2)) & "' " & _
              "where numero = '" & Trim(letra) & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
    End If
    frmLetras.CargarLetrasRenovadas frmLetras.mshLetras.TextMatrix(frmLetras.mshLetras.row, 1)
End Sub

Function DevInteresLetra(NumLetra As String) As Double
Dim SQL As String
Dim RQ As MYSQL_RS
    DevInteresLetra = False
    SQL = "select l.* from letra l left join documento_contables d on (l.ndebito=CONCAT(d.SERIE,'-',d.correl)) " & _
          "left join amarre_documento a on(d.identificador=a.identificador) where l.numero = '" & NumLetra & "' " & _
          "and a.cod_tipo_doc = '08'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        If RQ.Fields("ndebito") <> "" Then
            NumND = Trim(RQ.Fields("ndebito"))
            DevInteresLetra = FormatNumber(RQ.Fields("interes"), 2)
        End If
    Else
        NumND = ""
    End If
    Set RQ = Nothing
End Function

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Me.Width) / 2
    Top = (Screen.Height - Me.Height) / 2
End Sub
