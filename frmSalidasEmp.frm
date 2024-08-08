VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmSalidasEmp 
   BackColor       =   &H009F5539&
   Caption         =   "Movimiento del Empleado"
   ClientHeight    =   9660
   ClientLeft      =   3045
   ClientTop       =   3720
   ClientWidth     =   15240
   ForeColor       =   &H00000000&
   Icon            =   "frmSalidasEmp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   15240
   Begin Proyecto1.chameleonButton btnReporte 
      Height          =   465
      Left            =   12975
      TabIndex        =   66
      Top             =   30
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   820
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
      MICON           =   "frmSalidasEmp.frx":014A
      PICN            =   "frmSalidasEmp.frx":0166
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
      Height          =   585
      Left            =   30
      TabIndex        =   5
      Top             =   -60
      Width           =   5265
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1365
         TabIndex        =   67
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2007
         BuddyControl    =   "txtAnio"
         BuddyDispid     =   196621
         OrigLeft        =   1230
         OrigTop         =   180
         OrigRight       =   1470
         OrigBottom      =   555
         Max             =   2012
         Min             =   2007
         SyncBuddy       =   -1  'True
         BuddyProperty   =   -517
         Enabled         =   -1  'True
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   3
         Left            =   3960
         Max             =   11
         TabIndex        =   10
         Top             =   195
         Value           =   11
         Width           =   825
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmSalidasEmp.frx":02C0
         Left            =   2295
         List            =   "frmSalidasEmp.frx":02C2
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   165
         Width           =   1605
      End
      Begin Proyecto1.chameleonButton btnRefrescar 
         Height          =   375
         Left            =   4785
         TabIndex        =   8
         ToolTipText     =   "Refrescar Búsqueda - F5"
         Top             =   135
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSalidasEmp.frx":02C4
         PICN            =   "frmSalidasEmp.frx":02E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1845
         TabIndex        =   12
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   225
         Width           =   375
      End
      Begin MSForms.TextBox txtAnio 
         Height          =   285
         Left            =   525
         TabIndex        =   7
         Top             =   180
         Width           =   810
         VariousPropertyBits=   746604571
         Size            =   "1429;503"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin Proyecto1.chameleonButton CmdColor 
      Height          =   405
      Left            =   10980
      TabIndex        =   13
      Top             =   60
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Configurar Leyenda"
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
      MICON           =   "frmSalidasEmp.frx":0CF2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   585
      Left            =   5295
      TabIndex        =   0
      Top             =   -60
      Width           =   5655
      Begin Proyecto1.chameleonButton btnBuscar 
         Height          =   345
         Left            =   6390
         TabIndex        =   9
         ToolTipText     =   "Buscar"
         Top             =   150
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSalidasEmp.frx":0D0E
         PICN            =   "frmSalidasEmp.frx":0D2A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mov."
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
         Left            =   2235
         TabIndex        =   4
         Top             =   210
         Width           =   435
      End
      Begin MSForms.ComboBox cboMovEmp 
         Height          =   315
         Left            =   2670
         TabIndex        =   3
         Top             =   180
         Width           =   2955
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "5212;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtEmpleado 
         Height          =   315
         Left            =   885
         TabIndex        =   2
         Top             =   180
         Width           =   1320
         VariousPropertyBits=   746604571
         MaxLength       =   11
         Size            =   "2328;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
   End
   Begin Proyecto1.chameleonButton btnDia 
      Height          =   315
      Index           =   1
      Left            =   2535
      TabIndex        =   14
      Top             =   570
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Lunes"
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
      BCOL            =   14737632
      BCOLO           =   15309923
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSalidasEmp.frx":30AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnDia 
      Height          =   315
      Index           =   2
      Left            =   4695
      TabIndex        =   15
      Top             =   570
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Martes"
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
      BCOL            =   14737632
      BCOLO           =   15309923
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSalidasEmp.frx":30C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnDia 
      Height          =   315
      Index           =   3
      Left            =   6840
      TabIndex        =   16
      Top             =   570
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Miércoles"
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
      BCOL            =   14737632
      BCOLO           =   15309923
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSalidasEmp.frx":30E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnDia 
      Height          =   315
      Index           =   4
      Left            =   9000
      TabIndex        =   17
      Top             =   570
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Jueves"
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
      BCOL            =   14737632
      BCOLO           =   15309923
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSalidasEmp.frx":3100
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnDia 
      Height          =   315
      Index           =   6
      Left            =   13290
      TabIndex        =   18
      Top             =   570
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Sábado"
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
      BCOL            =   14737632
      BCOLO           =   15309923
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSalidasEmp.frx":311C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnDia 
      Height          =   315
      Index           =   0
      Left            =   420
      TabIndex        =   19
      Top             =   570
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Domingo"
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
      BCOL            =   14737632
      BCOLO           =   15309923
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSalidasEmp.frx":3138
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnDia 
      Height          =   315
      Index           =   5
      Left            =   11130
      TabIndex        =   20
      Top             =   570
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Viernes"
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
      BCOL            =   14737632
      BCOLO           =   15309923
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSalidasEmp.frx":3154
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCalendario 
      Height          =   1395
      Left            =   1275
      TabIndex        =   21
      Top             =   11790
      Visible         =   0   'False
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   2461
      _Version        =   393216
      Rows            =   7
      Cols            =   7
      FixedCols       =   0
      WordWrap        =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H009F5539&
      Height          =   9500
      Left            =   45
      TabIndex        =   22
      Top             =   810
      Width           =   17865
      Begin MSFlexGridLib.MSFlexGrid GridMenu 
         Height          =   2205
         Left            =   10485
         TabIndex        =   65
         Top             =   2115
         Visible         =   0   'False
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   3889
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         FocusRect       =   2
         GridLines       =   0
         GridLinesFixed  =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   4
         Left            =   8580
         TabIndex        =   23
         Top             =   90
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   40
         Left            =   10725
         TabIndex        =   24
         Top             =   6870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   39
         Left            =   8580
         TabIndex        =   25
         Top             =   6870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   37
         Left            =   4290
         TabIndex        =   26
         Top             =   6870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   36
         Left            =   2145
         TabIndex        =   27
         Top             =   6870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   35
         Left            =   0
         TabIndex        =   28
         Top             =   6870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   34
         Left            =   12870
         TabIndex        =   29
         Top             =   5520
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   33
         Left            =   10725
         TabIndex        =   30
         Top             =   5520
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   32
         Left            =   8580
         TabIndex        =   31
         Top             =   5520
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   30
         Left            =   4290
         TabIndex        =   32
         Top             =   5520
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   29
         Left            =   2145
         TabIndex        =   33
         Top             =   5520
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   28
         Left            =   0
         TabIndex        =   34
         Top             =   5520
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   27
         Left            =   12870
         TabIndex        =   35
         Top             =   4155
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   26
         Left            =   10725
         TabIndex        =   36
         Top             =   4155
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   25
         Left            =   8580
         TabIndex        =   37
         Top             =   4155
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   23
         Left            =   4290
         TabIndex        =   38
         Top             =   4155
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   22
         Left            =   2145
         TabIndex        =   39
         Top             =   4155
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   21
         Left            =   0
         TabIndex        =   40
         Top             =   4155
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   20
         Left            =   12870
         TabIndex        =   41
         Top             =   2805
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   19
         Left            =   10725
         TabIndex        =   42
         Top             =   2805
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   18
         Left            =   8580
         TabIndex        =   43
         Top             =   2805
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   16
         Left            =   4290
         TabIndex        =   44
         Top             =   2805
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   15
         Left            =   2145
         TabIndex        =   45
         Top             =   2805
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   14
         Left            =   0
         TabIndex        =   46
         Top             =   2805
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   13
         Left            =   12870
         TabIndex        =   47
         Top             =   1440
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   12
         Left            =   10725
         TabIndex        =   48
         Top             =   1440
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   11
         Left            =   8580
         TabIndex        =   49
         Top             =   1440
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   9
         Left            =   4290
         TabIndex        =   50
         Top             =   1440
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   8
         Left            =   2145
         TabIndex        =   51
         Top             =   1440
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   7
         Left            =   0
         TabIndex        =   52
         Top             =   1440
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   6
         Left            =   12870
         TabIndex        =   53
         Top             =   90
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   2
         Left            =   4290
         TabIndex        =   54
         Top             =   90
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   1
         Left            =   2145
         TabIndex        =   55
         Top             =   90
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   41
         Left            =   12870
         TabIndex        =   56
         Top             =   6870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   0
         Left            =   0
         TabIndex        =   57
         Top             =   90
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   5
         Left            =   10725
         TabIndex        =   58
         Top             =   90
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   3
         Left            =   6435
         TabIndex        =   59
         Top             =   90
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   38
         Left            =   6435
         TabIndex        =   60
         Top             =   6870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   31
         Left            =   6435
         TabIndex        =   61
         Top             =   5520
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   24
         Left            =   6435
         TabIndex        =   62
         Top             =   4155
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   17
         Left            =   6435
         TabIndex        =   63
         Top             =   2805
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDia 
         Height          =   1305
         Index           =   10
         Left            =   6435
         TabIndex        =   64
         Top             =   1440
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2302
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSalidasEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private datToday As Date
Private intThisYear As Date
Private intThisMonth As Date
Private intThisDay As Date
Private strMonthName As String
Private datFirstDay As Date
Private intFirstWeekDay As Integer
Private intLastDay As Integer
Private intprintday As Integer
Private Color(0 To 10) As String
Private Movi_emp(0 To 7, 0 To 2) As String
Private oConsulta As FrmConsultas
Dim IndGrid As Integer, IndCol As Integer, IndFila As Integer
Dim FilaIni As String, FilaFin As String, IndColFin As Integer

Private Sub btnBuscar_Click()
    CargarCalen
End Sub

Private Sub btnRefrescar_Click()
    cboMes.ListIndex = Month(Now) - 1
    txtAnio = CInt(Year(Now))
    txtEmpleado = Empty
    cboMovEmp.ListIndex = 0
    CargarCalen
    btnRefrescar.SetFocus
End Sub

Private Sub btnReporte_Click()
    Dim SQL As String, Str1 As String
    Dim RQ As MYSQL_RS
    Screen.MousePointer = vbHourglass
    If IndGrid >= 0 Then
        Str1 = " and if(fec_regreso<>'',fec_regreso,fec_salida) = '" & txtAnio.Text & "/" & Right("00" & MonthNumber(cboMes.List(cboMes.ListIndex)), 2) & "/" & Right("00" & flexDia(IndGrid).TextMatrix(0, 0), 2) & "'"
    Else
        Str1 = " and if(fec_regreso<>'',fec_regreso,fec_salida) >= '" & txtAnio.Text & "/" & Right("00" & MonthNumber(cboMes.List(cboMes.ListIndex)), 2) & "/01' and if(fec_regreso<>'',fec_regreso,fec_salida) <= '" & txtAnio.Text & "/" & Right("00" & MonthNumber(cboMes.List(cboMes.ListIndex)), 2) & "/31'"
    End If
    SQL = "SELECT l.descripcion as Situacion,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as Nombres, " & _
          "if(fec_regreso<>'','I','S') AS Tipo,hora_salida as 'Hor.Sal.', " & _
          "(SELECT DESCRIP FROM DEPARTAMENTO D WHERE D.CODIGO=C.DPTO) AS Dpto,(select l.descripcioncorta from novperuvhse.lote as l where l.idlote=C.LOTE) as Lote," & _
          "(select p.descripcioncorta from novperuvhse.pozo as p where p.idPozo=C.POZO) as Pozo,OBSERVACION as Observaciones,  " & _
          "IFNULL((SELECT DESCRIP FROM CNMDEPAR E WHERE E.CODDEP=C.PERIODO),'') AS Division,I.descrip as 'Tip.Linea',a.descrip as Agencia,s.descrip as Estancia " & _
          "FROM CALENDARIO C LEFT JOIN EMPLEADO E ON (C.CODEMP=E.CODIGO)left join leyendatraslados l on (c.sinbono=l.color) " & _
          "left join linea as I on (I.codigo=c.codlinea) left join estancia as S on(S.codigo=c.codestancia) " & _
          "left join agencia a on (c.codagencia=a.codigo) WHERE MOVEMP = '01' " & _
          Str1 & " ORDER BY TIPO,SITUACION,hora_salida,nombres"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        Exportar_Excel_Calendario Rep_Documents & "\Salidas.xls", RQ
    End If
    Screen.MousePointer = vbDefault
    Set RQ = Nothing
End Sub

Private Sub cboMes_Click()
    If txtAnio <> Empty Then CargarCalen
End Sub

Private Sub CargarCalen()
    Dim Indice As Integer
    Dim intLoopDay As Integer
    Dim i As Integer
    Frame3.Visible = False
    ConfigGrillas
    For i = 0 To flexDia.Count - 1
        flexDia(i).Clear
        flexDia(i).Visible = False
    Next

    datToday = DateSerial(CInt(txtAnio), MonthNumber(cboMes.List(cboMes.ListIndex)), 15)     'today's date
    intThisYear = Year(datToday)              'the current year
    intThisMonth = Month(datToday)            'the current month
    intThisDay = Day(datToday)                'the current day
    strMonthName = MonthName(intThisMonth)    'the name of the current month
    datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
    intFirstWeekDay = weekday(datFirstDay, vbSunday)
    intLastDay = GetLastDay(datToday)
    intprintday = 1   'the value of the day number to print in the page
    flxCalendario.Clear
    flxCalendario.row = 0
    While intprintday <= intLastDay
        flxCalendario.row = flxCalendario.row  'increment row
        For intLoopDay = 1 To 7
            flxCalendario.Col = intLoopDay - 1
            If intFirstWeekDay > 1 Then
                flxCalendario.Text = "-"  ' put a - if outside the month's range
                intFirstWeekDay = intFirstWeekDay - 1
            Else
                If intprintday > intLastDay Then
                    flxCalendario.Text = "-"
                Else
                    Indice = val(Trim(flxCalendario.row * 7 + flxCalendario.Col))
                    If cboMovEmp.List(cboMovEmp.ListIndex, 1) = "01" Then
                        flexDia(Indice).Col = 0
                        flexDia(Indice).row = 0
                        flexDia(Indice).CellFontBold = True
                        flexDia(Indice).TextMatrix(0, 0) = intprintday
                        flexDia(Indice).TextMatrix(1, 0) = "Ingresa"
                        flexDia(Indice).TextMatrix(1, 4) = "Sale"
                        flexDia(Indice).Col = 0
                        flexDia(Indice).row = 1
                        flexDia(Indice).CellFontSize = 9
                        flexDia(Indice).CellFontBold = True
                        flexDia(Indice).CellAlignment = 4
                        flexDia(Indice).Col = 4
                        flexDia(Indice).row = 1
                        flexDia(Indice).CellFontSize = 9
                        flexDia(Indice).CellFontBold = True
                        flexDia(Indice).CellAlignment = 4
                        flexDia(Indice).Col = 4
                        flexDia(Indice).row = 0
                        If flxCalendario.Col = 0 Or flxCalendario.Col = 6 Then
                            flexDia(Indice).CellFontSize = 8
                            flexDia(Indice).CellForeColor = &H372398      'rojito
                            flexDia(Indice).CellFontBold = True
                        Else
                            flexDia(Indice).CellFontSize = 8
                            flexDia(Indice).CellForeColor = &H800000     'azulito
                            flexDia(Indice).CellFontBold = True
                        End If
                    Else
                        flexDia(Indice).TextMatrix(0, 0) = intprintday
                    End If
                    flexDia(Indice).Col = 0
                    flexDia(Indice).row = 0
                    If flxCalendario.Col = 0 Or flxCalendario.Col = 6 Then
                        flexDia(Indice).CellFontSize = 8
                        flexDia(Indice).CellForeColor = &H372398      'rojito
                        flexDia(Indice).CellFontBold = True
                    Else
                        flexDia(Indice).CellFontSize = 8
                        flexDia(Indice).CellForeColor = &H800000     'azulito
                        flexDia(Indice).CellFontBold = True
                    End If
                    If flexDia(Indice).TextMatrix(0, 0) <> "" Then flexDia(Indice).Visible = True
                End If
                intprintday = intprintday + 1
            End If
        Next
        flxCalendario.row = flxCalendario.row + 1
    Wend
    LlenarCalendario
    Frame3.Visible = True
    DoEvents
End Sub

Private Sub cboMovEmp_Change()
    If cboMovEmp.List(cboMovEmp.ListIndex, 1) = "01" Then
        CmdColor.Visible = True
        btnReporte.Visible = True
        CargarMenu
    Else
        CmdColor.Visible = False
        btnReporte.Visible = False
    End If
    CargarCalen
End Sub

Private Sub CmdColor_Click()
    frmLeyenda.Show
End Sub

Private Sub flexDia_Click(Index As Integer)
    flexDia_RowColChange Index
    IndGrid = Index
End Sub

Private Sub flexDia_DblClick(Index As Integer)
    If cboMovEmp.List(cboMovEmp.ListIndex) = "CONTRATO" Then
        AbrirFormulario Index, 1
    Else
        AbrirFormulario Index, 2
    End If
End Sub

Private Sub ConfigGrillas()
    Dim i As Integer, J As Integer
    For i = 0 To 41
        For J = 0 To flexDia(i).Rows - 1
            flexDia(i).RowHeight(J) = 200
        Next
        If cboMovEmp.List(cboMovEmp.ListIndex, 1) = "01" Then
            flexDia(i).Cols = 8
            flexDia(i).ColWidth(0) = 930
            flexDia(i).ColWidth(4) = 930
            flexDia(i).ColWidth(5) = 0
            flexDia(i).ColWidth(6) = 0
            flexDia(i).ColWidth(7) = 0
            flexDia(i).Rows = 6
            flexDia(i).FixedRows = 2
        Else
            flexDia(i).Cols = 4
            flexDia(i).ColWidth(0) = 1900
            flexDia(i).Rows = 5
            flexDia(i).FixedRows = 1
        End If
        flexDia(i).ColWidth(1) = 0
        flexDia(i).ColWidth(2) = 0
        flexDia(i).ColWidth(3) = 0
        flexDia(i).BackColorBkg = &HC0FFFF
        flexDia(i).BorderStyle = flexBorderNone
        flexDia(i).GridLines = flexGridNone
    Next
    'MovGrillas
End Sub

Private Sub MovGrillas()
    Dim i As Integer
    Dim Columnas As Integer
    Dim X As Integer
    Dim Y As Integer
    Columnas = 7
    For i = 1 To flexDia.Count - 1
        With flexDia(i)
            If i Mod Columnas = 0 Then
                Y = flexDia(i - Columnas).Top + flexDia(0).Height + 85
                X = flexDia(0).Left
            Else
                Y = flexDia(i - 1).Top
                X = flexDia(i - 1).Left + flexDia(i - 1).Width + 59
            End If
            .Move X, Y
        End With
    Next
    MoverBtn
End Sub

Private Sub MoverBtn()
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    For i = 1 To btnDia.Count - 1
        With btnDia(i)
            Y = btnDia(0).Top
            X = flexDia(0).Width + (btnDia(i - 1).Left) + 35
            .Move X, Y
        End With
    Next
End Sub

Public Sub LlenarCalendario()
    VectorMovEmp
    Dim i As Integer, k As Integer, J As Integer
    Dim FechaCal As Date, FEC_CONS As String
    Dim AnioMes As String
    Dim SQL As String, sqlwhere As String, TipoMov As String
    Dim rscalen As MYSQL_RS
    If cboMovEmp.List(cboMovEmp.ListIndex, 0) <> "CONTRATO" Then
        SQL = "Select a.codigo,a.codemp,a.movemp,a.fec_salida,a.hora_salida,a.fec_regreso," & _
              " a.hora_regreso,a.dpto,(select descripcioncorta from novperuvhse.lote where idlote=a.lote) as lote," & _
              "(select descripcioncorta from novperuvhse.pozo where idpozo=a.pozo) as pozo,a.codagencia,a.codlinea,a.tipoboleto," & _
              " a.mon_boleto,a.monto_boleto,a.codestancia,a.pagoestancia,a.mon_estancia," & _
              " a.monto_estancia,a.mon_viatico,a.monto_viatico,a.observacion,a.sinbono," & _
              " a.periodo, a.gocehaber, a.autorizado from calendario as a left join empleado as b on (a.codemp=b.codigo)"
    Else
        SQL = "Select a.codigo,a.anomes,a.codtipo,a.codemp,a.f_inicio,a.f_termino," & _
              " a.mon_sueldo,a.sbasico,a.bono, a.mon_bono,a.monto_bono,a.estado,a.HorLab," & _
              " a.EstTrabajo from contrato as a left join empleado as b " & _
              " on (a.codeMp=b.codigo)  where  a.estado <>'" & PENDIENTE & "'"
    End If
    sqlwhere = " where"
    If txtEmpleado <> Empty Then
        sqlwhere = sqlwhere & " a.codemp = '" & txtEmpleado & "'"
    End If
    If cboMovEmp.List(cboMovEmp.ListIndex, 0) <> "CONTRATO" Then
        If sqlwhere <> " where" Then
              sqlwhere = sqlwhere & " and"
        End If
            sqlwhere = sqlwhere & " a.movemp = '" & cboMovEmp.List(cboMovEmp.ListIndex, 1) & "'"
    End If
    If sqlwhere = " where" Then
        sqlwhere = Empty
    End If
    If cboMovEmp.List(cboMovEmp.ListIndex, 0) = "NINGUNO" Then
        SQL = Empty
    Else
        If cboMovEmp.List(cboMovEmp.ListIndex, 0) = "TRASLADOS Y MOVILIDADES" Then
            SQL = SQL & sqlwhere & " order by a.codigo"
        Else
            SQL = SQL & sqlwhere & " order by b.apepat,b.apemat,b.nombre1,b.nombre2,a.codemp"
        End If
    End If
    Set rscalen = oConexion.EjecutaSelectRS(SQL)
    If cboMovEmp.List(cboMovEmp.ListIndex, 0) <> Empty Then
        TipoMov = cboMovEmp.List(cboMovEmp.ListIndex, 1)
    Else
        TipoMov = "00"
    End If
    If cboMovEmp.List(cboMovEmp.ListIndex, 0) <> "CONTRATO" Then
        If cboMovEmp.List(cboMovEmp.ListIndex, 0) = "TRASLADOS Y MOVILIDADES" Then
            Do While Not rscalen.EOF
               AnioMes = txtAnio & "/" & Right("00" & Trim(str(MonthNumber(cboMes.List(cboMes.ListIndex)))), 2)
               If rscalen.Fields("fec_regreso") <> "" Then FEC_CONS = rscalen.Fields("FEC_REGRESO") Else FEC_CONS = rscalen.Fields("FEC_SALIDA")
               If Replace(AnioMes, "/", "") = Replace(Left(FEC_CONS, 7), "/", "") Then
                    For i = 0 To 41
                         If val(Trim(flexDia(i).TextMatrix(0, 0))) > 0 Then
                            If IsDate(FEC_CONS) Then
                                FechaCal = CDate(AnioMes & "/" & Right("00" & Trim(flexDia(i).TextMatrix(0, 0)), 2))
                                If FechaCal = CDate(FEC_CONS) Then
                                    If rscalen.Fields("fec_regreso") = "" Then
                                        MarcaDiaCol4 i, Trim(rscalen.Fields("CodEmp")), Trim(rscalen.Fields("Codigo")), TipoMov, 0, Trim(rscalen.Fields("sinbono"))
                                    Else
                                        MarcaDia i, Trim(rscalen.Fields("CodEmp")), Trim(rscalen.Fields("Codigo")), TipoMov, 0, Trim(rscalen.Fields("sinbono"))
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                rscalen.MoveNext
            Loop
        Else
            Do While Not rscalen.EOF
               AnioMes = txtAnio & "/" & Right("00" & Trim(str(MonthNumber(cboMes.List(cboMes.ListIndex)))), 2)
               If Replace(AnioMes, "/", "") >= Replace(Left(rscalen.Fields("FEC_SALIDA"), 7), "/", "") And Replace(AnioMes, "/", "") <= Replace(Left(rscalen.Fields("FEC_REGRESO"), 7), "/", "") Then
                    For i = 0 To 41
                         If val(Trim(flexDia(i).TextMatrix(0, 0))) > 0 Then
                            If IsDate(rscalen.Fields("FEC_SALIDA")) And IsDate(rscalen.Fields("FEC_REGRESO")) Then
                                FechaCal = CDate(AnioMes & "/" & Right("00" & Trim(flexDia(i).TextMatrix(0, 0)), 2))
                                If FechaCal >= CDate(rscalen.Fields("FEC_SALIDA")) And FechaCal <= CDate(rscalen.Fields("FEC_REGRESO")) Then
                                     MarcaDia i, Trim(rscalen.Fields("CodEmp")), Trim(rscalen.Fields("Codigo")), TipoMov, 0, Trim(rscalen.Fields("sinbono"))
                                End If
                            End If
                        End If
                    Next
                End If
                rscalen.MoveNext
            Loop
        End If
    Else
        Do While Not rscalen.EOF
           AnioMes = txtAnio & "/" & Right("00" & Trim(str(MonthNumber(cboMes.List(cboMes.ListIndex)))), 2)
           If Replace(AnioMes, "/", "") >= Replace(Left(rscalen.Fields("F_INICIO"), 7), "/", "") And Replace(AnioMes, "/", "") <= Replace(Left(rscalen.Fields("F_TERMINO"), 7), "/", "") Then
                For i = 0 To 41
                     If val(Trim(flexDia(i).TextMatrix(0, 0))) > 0 Then
                        If IsDate(rscalen.Fields("F_INICIO")) And IsDate(rscalen.Fields("F_TERMINO")) Then
                            FechaCal = CDate(AnioMes & "/" & Right("00" & Trim(flexDia(i).TextMatrix(0, 0)), 2))
                            If FechaCal = CDate(rscalen.Fields("F_INICIO")) Then
                                MarcaDia i, Trim(rscalen.Fields("CodEmp")), Trim(rscalen.Fields("Codigo")), TipoMov, "I", Trim(rscalen.Fields("sinbono")), 0
                            End If
                            If FechaCal = CDate(rscalen.Fields("F_TERMINO")) And rscalen.Fields("Estado") <> CANCELADO Then
                                MarcaDia i, Trim(rscalen.Fields("CodEmp")), Trim(rscalen.Fields("Codigo")), TipoMov, "F", Trim(rscalen.Fields("sinbono")), 1
                            End If
                        End If
                    End If
                Next
            End If
            If rscalen.Fields("F_TERMINO") = "" Then
                For i = 0 To 41
                     If val(Trim(flexDia(i).TextMatrix(0, 0))) > 0 Then
                        If IsDate(rscalen.Fields("F_INICIO")) Then
                            FechaCal = CDate(AnioMes & "/" & Right("00" & Trim(flexDia(i).TextMatrix(0, 0)), 2))
                            If FechaCal = CDate(rscalen.Fields("F_INICIO")) Then
                                MarcaDia i, Trim(rscalen.Fields("CodEmp")), Trim(rscalen.Fields("Codigo")), TipoMov, "I", Trim(rscalen.Fields("sinbono")), 0
                            End If
                        End If
                    End If
                Next
            End If
            rscalen.MoveNext
        Loop
    End If
End Sub

Private Sub flexDia_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If cboMovEmp.List(cboMovEmp.ListIndex, 1) = "01" Then
            If flexDia(Index).TextMatrix(flexDia(Index).row, flexDia(Index).Col) <> "" Then
                If MsgBox("¿Está seguro que desea eliminar al empleado " & flexDia(Index).TextMatrix(flexDia(Index).row, flexDia(Index).Col) & " del día " & flexDia(Index).TextMatrix(0, 0) & " de " & cboMes.List(cboMes.ListIndex) & "?", vbQuestion + vbYesNo, gsNomSW) = vbYes Then
                    EliminarEmp flexDia(Index).TextMatrix(flexDia(Index).row, flexDia(Index).Col + 1), flexDia(Index).TextMatrix(flexDia(Index).row, flexDia(Index).Col + 2)
                    flexDia(Index).RemoveItem flexDia(Index).row
                    CargarCalen
                End If
            End If
        End If
    End If
   If KeyCode = 67 And Shift = 2 Then
       Copiar_Clipboard flexDia(Index)
    End If
End Sub

Sub EliminarEmp(CodEmp As String, CodCal As String)
    Dim SQL As String
    SQL = "delete from calendario where movemp = '01' and codemp = '" & CodEmp & "' and codigo = '" & CodCal & "'"
    oConexionMYSQL.Execute SQL
End Sub

Private Sub flexDia_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    FilaIni = flexDia(Index).Rowsel
    IndCol = flexDia(Index).Col
    If Button = vbRightButton Then
        If cboMovEmp.List(cboMovEmp.ListIndex, 1) = "01" Then
            Dim Flg As Boolean
            Flg = False
            If flexDia(Index).Col = 0 Then
                If Trim(flexDia(Index).TextMatrix(flexDia(Index).row, 0)) <> "" Then Flg = True
            End If
            If flexDia(Index).Col = 4 Then
                If Trim(flexDia(Index).TextMatrix(flexDia(Index).row, 4)) <> "" Then Flg = True
            End If
            If Flg = True Then
                IndGrid = Index
                IndCol = flexDia(Index).Col
                IndFila = flexDia(Index).row
                GridMenu.Top = flexDia(Index).Top + (flexDia(Index).row * flexDia(Index).CellHeight)
                If flexDia(Index).Col = 0 Then
                    GridMenu.Left = flexDia(Index).Left + flexDia(Index).CellWidth
                ElseIf flexDia(Index).Col = 4 Then
                    GridMenu.Left = flexDia(Index).Left + (2 * flexDia(Index).CellWidth)
                End If
                Select Case Index
                 Case 6, 13, 20, 27, 34, 41
                    GridMenu.Left = GridMenu.Left - GridMenu.Width - flexDia(Index).Width + 1200
                 
                End Select
                GridMenu.Visible = True
            End If
        End If
    End If
End Sub

Private Sub flexDia_RowColChange(Index As Integer)
    If flexDia(Index).Col = 0 Then
        If flexDia(Index).TextMatrix(flexDia(Index).row, 1) <> "" Then
            frmSalidasEmp.Caption = "Movimiento del Empleado..." & " [" & UCase(DescripcionesdeCodigos("EMPLEADO", flexDia(Index).TextMatrix(flexDia(Index).row, 1))) & "]" & " [" & UCase(flexDia(Index).TextMatrix(flexDia(Index).row, 1)) & "]"
        Else
            frmSalidasEmp.Caption = "Movimiento del Empleado..."
        End If
    Else
        If flexDia(Index).Col = 4 Then
            If flexDia(Index).TextMatrix(flexDia(Index).row, 5) <> "" Then
                frmSalidasEmp.Caption = "Movimiento del Empleado..." & " [" & UCase(DescripcionesdeCodigos("EMPLEADO", flexDia(Index).TextMatrix(flexDia(Index).row, 5))) & "]" & " [" & UCase(flexDia(Index).TextMatrix(flexDia(Index).row, 5)) & "]"
            Else
                frmSalidasEmp.Caption = "Movimiento del Empleado..."
            End If
        Else
            frmSalidasEmp.Caption = "Movimiento del Empleado..."
        End If
    End If
End Sub

Sub CargarMenu()
    Dim SQL As String, i As Integer
    Dim RQ As MYSQL_RS
    
    SQL = "select * from leyendatraslados"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    GridMenu.Clear
    GridMenu.Cols = 3
    GridMenu.Rows = 0
    GridMenu.ColWidth(0) = 280
    GridMenu.ColWidth(1) = 4000
    GridMenu.ColWidth(2) = 0
    i = 0
    Do While Not RQ.EOF()
        GridMenu.AddItem "" & vbTab & Trim(RQ.Fields("descripcion")) & vbTab & Trim(RQ.Fields("color"))
        GridMenu.Col = 0
        GridMenu.row = i
        GridMenu.CellBackColor = Trim(RQ.Fields("color"))
        i = i + 1
        RQ.MoveNext
    Loop
    Set RQ = Nothing
End Sub

Private Sub flexDia_SelChange(Index As Integer)
    FilaFin = flexDia(Index).Rowsel
    IndColFin = flexDia(Index).ColSel
End Sub

Private Sub Form_Activate()
    CargarCalen
    CargarMenu
End Sub
Private Sub Form_Click()
    IndGrid = -1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        GridMenu.Visible = False
    End If
End Sub
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    IndGrid = -1
    CargarCbo
    txtAnio = CInt(Year(Now))
    HScroll1.Value = cboMes.ListIndex
    MovEmp cboMovEmp
    flxCalendario.BandExpandable(0) = True
    frmSalidasEmp.Height = 9390
    frmSalidasEmp.Width = 15360
    Set oConsulta = New FrmConsultas
    Me.KeyPreview = True
End Sub

Private Sub CargarCbo()
    cboMes.AddItem "Enero"
    cboMes.AddItem "Febrero"
    cboMes.AddItem "Marzo"
    cboMes.AddItem "Abril"
    cboMes.AddItem "Mayo"
    cboMes.AddItem "Junio"
    cboMes.AddItem "Julio"
    cboMes.AddItem "Agosto"
    cboMes.AddItem "Septiembre"
    cboMes.AddItem "Octubre"
    cboMes.AddItem "Noviembre"
    cboMes.AddItem "Diciembre"
    cboMes.ListIndex = Month(Now) - 1  'set to the present month
End Sub

Function GetLastDay(datTheDate As Variant) As Integer
    GetLastDay = Day(DateAdd("m", 1, DateSerial(CInt(txtAnio), MonthNumber(cboMes.Text), 1)) - 1) 'last day of the month
End Function

Function MonthName(X As Variant) As String
    Select Case X
        Case 1
            MonthName = "Enero"
        Case 2
            MonthName = "Febrero"
        Case 3
            MonthName = "Marzo"
        Case 4
            MonthName = "Abril"
        Case 5
            MonthName = "Mayo"
        Case 6
            MonthName = "Junio"
        Case 7
            MonthName = "Julio"
        Case 8
            MonthName = "Agosto"
        Case 9
            MonthName = "Septiembre"
        Case 10
            MonthName = "Octubre"
        Case 11
            MonthName = "Noviembre"
        Case 12
            MonthName = "Diciembre"
    End Select
End Function

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Frame3.Width = Me.Width - 200
        Frame3.Height = Me.Height - Frame3.Top - 450
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConsulta = Nothing
End Sub

Private Sub GridMenu_Click()
    flexDia(IndGrid).row = IndFila
    flexDia(IndGrid).Col = IndCol
    flexDia(IndGrid).CellBackColor = GridMenu.TextMatrix(GridMenu.Rowsel, 2)
    ActualizarenCalendario GridMenu.TextMatrix(GridMenu.Rowsel, 2)
    GridMenu.Visible = False
End Sub

Sub ActualizarenCalendario(Color As String)
    Dim SQL As String
    SQL = "update calendario set sinbono ='" & Color & "' where movemp = '01' and  " & _
          "codemp = '" & flexDia(IndGrid).TextMatrix(IndFila, IndCol + 1) & "' and " & _
          "codigo = '" & flexDia(IndGrid).TextMatrix(IndFila, IndCol + 2) & "'"
    oConexionMYSQL.Execute SQL
End Sub

Private Sub HScroll1_Change()
    cboMes.ListIndex = CInt(HScroll1.Value)
End Sub


Private Sub txtAnio_Change()
    CargarCalen
End Sub

Private Sub colores()
    Color(0) = &HE0E0E0 'gris
    Color(1) = &HDBD5FB 'rosado
    Color(2) = &HC0E0FF 'naranja
    Color(3) = &HF2DBD0 'AZUL
    Color(4) = &HC0FFC0 'verde
    Color(5) = &HFFC0FF 'lila
    Color(6) = &HDFEFD3 'VERDE ++
    Color(7) = &H95DDEA 'AMARIYELLOW
    Color(8) = &HFFFFC0 'Celeste
    Color(9) = &HFFFFFF 'blanco
End Sub

Private Sub MarcaDia(dia As Integer, emp As String, CodCalen As String, TipoMov As String, TipFec As String, Optional Bono As String, Optional val As Integer)
    Dim i As Integer, J As Integer, INI As Integer
    Dim strColor As String
    If cboMovEmp.List(cboMovEmp.ListIndex, 1) = "01" Then INI = 2 Else INI = 1
    For J = INI To flexDia(dia).Rows - 1
        If flexDia(dia).TextMatrix(J, 0) = Empty Then
            flexDia(dia).row = J
            If J > 2 Then
                flexDia(dia).Rows = flexDia(dia).Rows + 1
            End If
            For i = 0 To 7
                If Movi_emp(i, 0) = TipoMov Then
                    If TipoMov = "04" Then
                        If val = 0 Then
                            strColor = Movi_emp(i, 2)
                        Else
                            strColor = Color(2)
                        End If
                    Else
                        If J Mod 2 = 0 Then
                            strColor = vbWhite
                            Exit For
                        Else
                            strColor = Movi_emp(i, 2)
                        End If
                    End If
                End If
            Next
            flexDia(dia).Col = 0
            flexDia(dia).GridLines = flexGridInset
            flexDia(dia).CellBackColor = strColor
            If cboMovEmp.List(cboMovEmp.ListIndex, 1) = "01" Then
                flexDia(dia).CellBackColor = IIf(Bono = "", vbWhite, Bono)
                flexDia(dia).TextMatrix(J, 0) = LCase(DescripcionesdeCodigos("EMPLEADOABREV", emp, "Nom"))
            Else
                flexDia(dia).TextMatrix(J, 0) = LCase(DescripcionesdeCodigos("EMPLEADO", emp))
            End If
            flexDia(dia).TextMatrix(J, 1) = emp
            flexDia(dia).TextMatrix(J, 2) = CodCalen
            flexDia(dia).TextMatrix(J, 3) = TipFec
            Exit For
        End If
    Next
End Sub

Private Sub MarcaDiaCol4(dia As Integer, emp As String, CodCalen As String, TipoMov As String, TipFec As String, Optional sinbono As String)
    Dim i As Integer, J As Integer
    Dim strColor As String
    For J = 2 To flexDia(dia).Rows - 1
        If flexDia(dia).TextMatrix(J, 4) = Empty Then
            flexDia(dia).row = J
            If J > 2 Then
                flexDia(dia).Rows = flexDia(dia).Rows + 1
            End If
            flexDia(dia).Col = 4
            flexDia(dia).GridLines = flexGridInset
            flexDia(dia).CellBackColor = sinbono
            flexDia(dia).TextMatrix(J, 4) = LCase(DescripcionesdeCodigos("EMPLEADOABREV", emp, "Nom"))
            flexDia(dia).TextMatrix(J, 5) = emp
            flexDia(dia).TextMatrix(J, 6) = CodCalen
            flexDia(dia).TextMatrix(J, 7) = TipFec
            Exit For
        End If
    Next
End Sub

Private Sub txtEmpleado_Change()
    If txtEmpleado = Empty Then
        frmSalidasEmp.Caption = "Movimiento del Empleado"
        txtEmpleado.ToolTipText = Empty
    End If
    CargarCalen
End Sub

Private Sub txtEmpleado_GotFocus()
    mark1 txtEmpleado
End Sub

Private Sub txtEmpleado_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtEmpleado = Right("00000000000" & Trim(txtEmpleado), 11)
        txtEmpleado.ToolTipText = DescripcionesdeCodigos("EMPLEADO", txtEmpleado)
        frmSalidasEmp.Caption = "Movimiento del Empleado - [ " & Trim(txtEmpleado.ToolTipText) & " ]"
        cboMovEmp.SetFocus
    End If
    If KeyCode.Value = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1200
            .pCol = 1: .pAnchoCol = 3800
            .pTitulo = "Empleados"
            .pForm = FORM_SALIDASEMP
            .pCaso = LABEL_EMPREG_F1
            .Show
        End With
    End If
End Sub

Public Sub AbrirFormulario(Indice As Integer, i As Integer)
    Dim fecha  As String
    Dim fechaI  As Date
    Dim fechaF As Date
    Select Case i
           Case 1
                fecha = txtAnio & "/" & Right("00" & MonthNumber(cboMes.List(cboMes.ListIndex)), 2) & "/" & Right("00" & Trim(flexDia(Indice).TextMatrix(0, 0)), 2)
                With frmContrato
                    If flexDia(Indice).TextMatrix(0, 0) <> Empty Then
                        If flexDia(Indice).row > 0 Then
                            If Trim(flexDia(Indice).TextMatrix(flexDia(Indice).row, 0)) = Empty Then
                                 Exit Sub
                            Else
                                .lblModo = "Consulta"
                                .DatosContrato flexDia(Indice).TextMatrix(flexDia(Indice).row, 1), _
                                               flexDia(Indice).TextMatrix(flexDia(Indice).row, 2), fecha, flexDia(Indice).TextMatrix(flexDia(Indice).row, 3)
                                .Show
                            End If
                        End If
                    End If
                End With
            Case 2
                With frmPrograma
                    If flexDia(Indice).TextMatrix(0, 0) <> Empty Then
                        .tag = UCase(cboMovEmp.List(cboMovEmp.ListIndex, 0))
                        If UCase(cboMovEmp.List(cboMovEmp.ListIndex, 0)) = "TRASLADOS Y MOVILIDADES" Then
                            .lblModo = "Consulta"
                            .SSTab1.TabVisible(0) = True
                            .SSTab1.TabVisible(1) = False
                            .SSTab1.Tab = 0
                            .Frame3.Caption = IIf(flexDia(Indice).Col = 0, "Destino", "Origen")
                            .Label2.Caption = IIf(flexDia(Indice).Col = 0, "Salida", "Llegada")
                            .lbltipo = IIf(flexDia(Indice).Col = 0, "I", "S")
                            .dpsalida = flexDia(Indice).TextMatrix(0, 0) & "/" & MonthNumber(cboMes.List(cboMes.ListIndex)) & "/" & Trim(txtAnio.Text)
                            .lblcolor.Caption = IIf(CStr(flexDia(Indice).CellBackColor) = "0", "16777215", flexDia(Indice).CellBackColor)
                            If flexDia(Indice).Col = 0 Then
                                .CargaTab0 flexDia(Indice).TextMatrix(flexDia(Indice).row, 2), flexDia(Indice).TextMatrix(flexDia(Indice).row, 1)
                            Else
                                .CargaTab0 flexDia(Indice).TextMatrix(flexDia(Indice).row, 6), flexDia(Indice).TextMatrix(flexDia(Indice).row, 5)
                            End If
                            .btnModificar.Visible = True
                        Else
                            .SSTab1.TabVisible(0) = False
                            .SSTab1.TabVisible(1) = True
                            .SSTab1.Tab = 1
                            .CargaTab1 flexDia(Indice).TextMatrix(flexDia(Indice).row, 2), flexDia(Indice).TextMatrix(flexDia(Indice).row, 1)
                            .lblModo = "Consulta"
                        End If
                    End If
                End With
    End Select
End Sub

Private Sub VectorMovEmp()
    Dim SQL As String
    Dim rsmov As MYSQL_RS
    Dim i As Integer
    SQL = "Select * from movi_emp order by DESCRIP"
    colores
    For i = 0 To 7
        Movi_emp(i, 0) = Empty
        Movi_emp(i, 1) = Empty
        Movi_emp(i, 2) = Empty
    Next
    i = 0
    Set rsmov = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsmov.EOF
        Movi_emp(i, 0) = rsmov.Fields("codigo")
        Movi_emp(i, 1) = rsmov.Fields("descrip")
        Movi_emp(i, 2) = Color(i)
        i = i + 1
        If i = 9 Then i = 0
        rsmov.MoveNext
    Loop
    Set rsmov = Nothing
End Sub

Sub Copiar_Clipboard(flexDia As MSFlexGrid)
On Error GoTo errSub
Dim columna As Long
Dim fila As Long
Dim Datos As String
Dim FlgIng As Boolean
    Datos = vbNullString
    If flexDia.Rows < 1 Then
        MsgBox " No hay datos en el FlexGrid para copiar", vbExclamation
        Exit Sub
    End If
    FlgIng = False
    For fila = FilaIni To FilaFin
        For columna = IndCol To IndColFin
            If FlgIng = False Then
                FlgIng = True
                If (IndCol = IndColFin) And (IndCol = 0) Then
                    Datos = Datos & "Ingresa" & vbNewLine
                ElseIf (IndCol = IndColFin) And (IndCol = 4) Then
                    Datos = Datos & "Sale" & vbNewLine
                ElseIf (IndCol < IndColFin) Then
                    Datos = Datos & "Ingresa" & vbTab & vbTab & "Sale" & vbNewLine
                End If
            End If
            If columna = 0 Or columna = 4 Then
                Datos = Datos & flexDia.TextMatrix(fila, columna) & vbTab  '& ";"
            End If
        Next columna
        Datos = Datos & vbNewLine
    Next fila
    Clipboard.Clear
    Clipboard.SetText Datos
    FilaIni = ""
    FilaFin = ""
    IndCol = 0
    IndColFin = 0
    'MsgBox Datos
Exit Sub
errSub:
MsgBox " Número de error: " & err.Number & vbNewLine & _
       " Descripción del error: " & err.Description, vbCritical
End Sub

