VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProg 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programación de Folios"
   ClientHeight    =   2580
   ClientLeft      =   8205
   ClientTop       =   6135
   ClientWidth     =   2895
   Icon            =   "frmProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   2895
   Begin MSMask.MaskEdBox meDia 
      Height          =   345
      Left            =   1260
      TabIndex        =   7
      Top             =   1620
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   2
      Format          =   "##"
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox chkEstado 
      BackColor       =   &H009F5539&
      Caption         =   "Activar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   300
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dpIni 
      Height          =   345
      Left            =   1260
      TabIndex        =   0
      Top             =   660
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      Format          =   60424193
      CurrentDate     =   39867
   End
   Begin MSComCtl2.DTPicker dpFin 
      Height          =   345
      Left            =   1260
      TabIndex        =   1
      Top             =   1140
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      Format          =   60424193
      CurrentDate     =   39867
   End
   Begin Proyecto1.chameleonButton btnEliminar 
      Height          =   345
      Left            =   2190
      TabIndex        =   8
      ToolTipText     =   "Eliminar"
      Top             =   2130
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
      MICON           =   "frmProg.frx":030A
      PICN            =   "frmProg.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblFolio 
      Alignment       =   2  'Center
      BackColor       =   &H009F5539&
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
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   300
      TabIndex        =   6
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label Label2 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Día:"
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
      Left            =   300
      TabIndex        =   4
      Top             =   1650
      Width           =   765
   End
   Begin VB.Label Label1 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fin:"
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
      Left            =   300
      TabIndex        =   3
      Top             =   1140
      Width           =   765
   End
   Begin VB.Label Label6 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio:"
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
      Height          =   345
      Left            =   300
      TabIndex        =   2
      Top             =   660
      Width           =   765
   End
End
Attribute VB_Name = "frmProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEliminar_Click()
    oConexionMYSQL.Execute "Delete from prog_folios where identificador='" & lblFolio & "'"
End Sub

Private Sub chkEstado_Click()
On Error GoTo DUPLI
    If chkEstado.Value = 1 Then
        oConexionMYSQL.Execute "Insert into prog_folios (identificador,fec_ini,fec_fin,dia,estado,anomes) values ('" & _
                                lblFolio & "','" & Format(dpIni.Value, "yyyy/mm/dd") & "','" & Format(dpFin, "yyyy/mm/dd") & "','" & meDia & "','0','')"
        oConexionMYSQL.Execute "Update prog_folios set estado='1' where identificador='" & lblFolio & "'"
    Else
        oConexionMYSQL.Execute "Update prog_folios set estado='0' where identificador='" & lblFolio & "'"
    End If
    Exit Sub
DUPLI:
    Resume Next
End Sub

Private Sub Form_Load()
    If strIdentificador <> "" Then
        dpIni.Enabled = True
        dpFin.Enabled = True
        meDia.Enabled = True
        chkEstado.Enabled = True
        btnEliminar.Enabled = True
        lblFolio = strIdentificador
        cargaProg strIdentificador
    Else
        dpIni.Enabled = False
        dpFin.Enabled = False
        meDia.Enabled = False
        chkEstado.Enabled = False
        btnEliminar.Enabled = False
    End If
End Sub

Sub cargaProg(id As String)
    Dim rsprog As New MYSQL_RS
    Set rsprog = oConexion.EjecutaSelectRS("Select * from prog_folios where identificador='" & id & "'")
    Do While Not rsprog.EOF
        dpIni = rsprog.Fields("fec_ini")
        dpFin = rsprog.Fields("fec_fin")
        meDia = rsprog.Fields("dia")
        If rsprog.Fields("estado") = "0" Then
            chkEstado.Value = 0
        Else
            chkEstado.Value = 1
        End If
        rsprog.MoveNext
    Loop
End Sub
