VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMensaje 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ingresando..."
   ClientHeight    =   1800
   ClientLeft      =   3600
   ClientTop       =   3495
   ClientWidth     =   5310
   Icon            =   "frmMensaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H009F5539&
      BorderStyle     =   0  'None
      FillColor       =   &H00800000&
      Height          =   1725
      Left            =   30
      ScaleHeight     =   1725
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   30
      Width           =   5235
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   -30
         Top             =   1320
      End
      Begin MSForms.Label lblUserInavlid 
         Height          =   495
         Left            =   690
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   4275
         ForeColor       =   16777215
         BackColor       =   10442041
         VariousPropertyBits=   8388627
         Caption         =   "No está autorizado para ingresar al Sistema Integrado Administrativo"
         Size            =   "7541;873"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblUserOk 
         Height          =   525
         Left            =   720
         TabIndex        =   2
         Top             =   390
         Visible         =   0   'False
         Width           =   4395
         ForeColor       =   16777215
         BackColor       =   10442041
         VariousPropertyBits=   8388627
         Caption         =   "Bienvenido al Sistema Integrado Administrativo... "
         Size            =   "7752;926"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label lblEmpresa 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NATIONAL OILWELL VARCO PERU S.A."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   300
         TabIndex        =   1
         Top             =   1110
         Width           =   4605
      End
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    lblUserInavlid.Visible = False
    'lblCerrando.Visible = False
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If lblUserOk.Visible = True Then
        'Call Actualiza_Status_bar
        mdiInicio.Show

        Unload Me
    Else
        Unload frmLoginUSer
        Unload Me
        End
    End If
End Sub
