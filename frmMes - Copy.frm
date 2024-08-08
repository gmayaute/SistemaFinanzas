VERSION 5.00
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmMes 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de los meses"
   ClientHeight    =   3960
   ClientLeft      =   3510
   ClientTop       =   3765
   ClientWidth     =   3840
   Icon            =   "frmMes.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3840
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Caption         =   "Seleccione el mes a utilizar en el sistema"
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
      Height          =   3285
      Left            =   15
      TabIndex        =   14
      Top             =   105
      Width           =   3765
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "C&ierre"
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
         Height          =   285
         Index           =   13
         Left            =   2040
         TabIndex        =   13
         Top             =   2760
         Width           =   1110
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Diciembre"
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
         Height          =   285
         Index           =   12
         Left            =   2040
         TabIndex        =   12
         Top             =   2340
         Width           =   1440
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Noviembre"
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
         Height          =   285
         Index           =   11
         Left            =   2040
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Octubre"
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
         Height          =   285
         Index           =   10
         Left            =   2040
         TabIndex        =   10
         Top             =   1500
         Width           =   1335
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Setiembre"
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
         Height          =   285
         Index           =   9
         Left            =   2040
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "A&gosto"
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
         Height          =   285
         Index           =   8
         Left            =   2040
         TabIndex        =   8
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "J&ulio"
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
         Height          =   285
         Index           =   7
         Left            =   2040
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Junio"
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
         Height          =   285
         Index           =   6
         Left            =   270
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Mayo"
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
         Height          =   285
         Index           =   5
         Left            =   270
         TabIndex        =   5
         Top             =   2340
         Width           =   1110
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "A&bril"
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
         Height          =   285
         Index           =   4
         Left            =   270
         TabIndex        =   4
         Top             =   1920
         Width           =   1110
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Marzo"
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
         Height          =   285
         Index           =   3
         Left            =   270
         TabIndex        =   3
         Top             =   1500
         Width           =   1110
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Febrero"
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
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   2
         Top             =   1080
         Width           =   1110
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "&Enero"
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
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   1
         Top             =   660
         Width           =   1110
      End
      Begin VB.OptionButton optMes 
         BackColor       =   &H009F5539&
         Caption         =   "A&pertura"
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
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   0
         Top             =   300
         Width           =   1110
      End
   End
   Begin Proyecto1.chameleonButton cmdAceptar 
      Height          =   405
      Left            =   330
      TabIndex        =   15
      Top             =   3450
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Aceptar"
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
      MICON           =   "frmMes.frx":08CA
      PICN            =   "frmMes.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdCancelar 
      Height          =   405
      Left            =   2010
      TabIndex        =   16
      Top             =   3450
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      MICON           =   "frmMes.frx":0A40
      PICN            =   "frmMes.frx":0A5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAceptar_Click()
  SQL = "Call CONT_Rep_Proc_Genericos('7cia_user_m','" & Right("00" & Trim(MesSistema), 2) & "','" & CodigoEmpresa & "','" & UsuarioActivo & "','','','','','','','' );"
  ADO_EjecutaQry (SQL)
  Actualiza_Status_bar
  Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Set RS = New ADODB.Recordset
    SQL = "Call CONT_Rep_Proc_Genericos('7cia_user_s','" & Right("00" & Trim(CodigoEmpresa), 2) & "','" & UsuarioActivo & "','','','','','','','','' );"
    Set RS = ADO_LlenaRs(SQL)
    
    MesSistema = Right("00" & RTrim(RS.Fields("mes_actual")), 2)
    optMes(MesSistema).Value = True
    Set RS = Nothing
End Sub
Private Sub optMes_Click(Index As Integer)
    MesSistema = Right("00" & RTrim(Index), 2)
End Sub
                    



