VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegEmpleado 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Empleados"
   ClientHeight    =   8445
   ClientLeft      =   3240
   ClientTop       =   6435
   ClientWidth     =   16650
   Icon            =   "frmRegEmpleado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   16650
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   2745
      Left            =   14100
      TabIndex        =   206
      Top             =   630
      Width           =   2415
      Begin MSComctlLib.ProgressBar pbProgreso 
         Height          =   225
         Left            =   240
         TabIndex        =   207
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Image imgFoto 
         Height          =   2550
         Left            =   30
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2295
      End
   End
   Begin Proyecto1.chameleonButton CmdDirectorio 
      Height          =   345
      Left            =   7695
      TabIndex        =   191
      ToolTipText     =   "Directorio de Correos Actuales"
      Top             =   7965
      Width           =   405
      _ExtentX        =   714
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
      MICON           =   "frmRegEmpleado.frx":014A
      PICN            =   "frmRegEmpleado.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdenviar 
      Height          =   345
      Left            =   8685
      TabIndex        =   171
      ToolTipText     =   "Enviar Email Solicitud Autorizaciones"
      Top             =   7980
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
      MICON           =   "frmRegEmpleado.frx":02C0
      PICN            =   "frmRegEmpleado.frx":02DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7200
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   12700
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   1058
      BackColor       =   10442041
      TabCaption(0)   =   "   Datos del Empleado"
      TabPicture(0)   =   "frmRegEmpleado.frx":0C52
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "    Datos de Dependientes"
      TabPicture(1)   =   "frmRegEmpleado.frx":5A54
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "   Condiciones del Empleado"
      TabPicture(2)   =   "frmRegEmpleado.frx":5D6E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(2)=   "Frame7"
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(4)=   "Frame6"
      Tab(2).Control(5)=   "FrCont"
      Tab(2).Control(6)=   "Frame13"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "   Otros"
      TabPicture(3)   =   "frmRegEmpleado.frx":6088
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         BackColor       =   &H009F5539&
         Height          =   6570
         Left            =   -120
         TabIndex        =   83
         Top             =   1200
         Width           =   13875
         Begin VB.TextBox txtcargo 
            Height          =   285
            Left            =   930
            TabIndex        =   245
            Top             =   5640
            Width           =   735
         End
         Begin VB.TextBox lblcargo 
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
            Height          =   285
            Left            =   1770
            TabIndex        =   244
            Top             =   5640
            Width           =   3615
         End
         Begin VB.Frame frmformedu 
            BackColor       =   &H009F5539&
            Height          =   5415
            Left            =   150
            TabIndex        =   241
            Top             =   150
            Visible         =   0   'False
            Width           =   13755
            Begin NOVAdmin.flxEdit flxformEduEmp 
               Height          =   4815
               Left            =   60
               TabIndex        =   242
               Top             =   180
               Width           =   13605
               _ExtentX        =   23998
               _ExtentY        =   8493
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CellFontName    =   "MS Sans Serif"
               CellFontSize    =   8.25
               BackColorSel    =   -2147483635
               BackColorFixed  =   -2147483633
               CellPicture     =   "frmRegEmpleado.frx":63A2
               ConfirmarBorradoLinea=   0   'False
               ColWidth0       =   960
               ColAlignment0   =   9
               FixedAlignment0 =   9
               ColWidth1       =   960
               ColAlignment1   =   9
               FixedAlignment1 =   9
               ForeColorSel    =   -2147483634
               ForeColorFixed  =   -2147483630
               GridColorFixed  =   12632256
               MouseIcon       =   "frmRegEmpleado.frx":63BE
               RowHeight0      =   240
               RowHeight1      =   240
            End
            Begin Proyecto1.chameleonButton btnEliminarDoc 
               Height          =   345
               Left            =   12720
               TabIndex        =   243
               ToolTipText     =   "Eliminar"
               Top             =   4920
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   609
               BTYPE           =   14
               TX              =   "Eliminar"
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
               MICON           =   "frmRegEmpleado.frx":63DA
               PICN            =   "frmRegEmpleado.frx":63F6
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
         Begin VB.TextBox txtHCMEmpleado 
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
            Height          =   345
            Left            =   2190
            MaxLength       =   20
            TabIndex        =   205
            Top             =   180
            Width           =   795
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H009F5539&
            Caption         =   "Grado de Instrucción"
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
            Height          =   2925
            Left            =   6120
            TabIndex        =   112
            Top             =   3240
            Width           =   7695
            Begin VB.CheckBox OptRegEduPriv 
               Caption         =   "Check1"
               Height          =   255
               Left            =   5040
               TabIndex        =   40
               Top             =   840
               Width           =   255
            End
            Begin VB.CheckBox OptRegEduPub 
               Caption         =   "Check1"
               Height          =   255
               Left            =   3360
               TabIndex        =   251
               Top             =   840
               Width           =   255
            End
            Begin VB.OptionButton OptPeruNo 
               Caption         =   "Option1"
               Height          =   255
               Left            =   5040
               TabIndex        =   228
               Top             =   480
               Width           =   255
            End
            Begin VB.OptionButton OptPeruSi 
               Caption         =   "opt2"
               Height          =   255
               Left            =   3360
               TabIndex        =   227
               Top             =   480
               Width           =   255
            End
            Begin Proyecto1.chameleonButton btnGrabarFormEduc 
               Height          =   345
               Left            =   6240
               TabIndex        =   237
               ToolTipText     =   "Guardar"
               Top             =   2280
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
               MICON           =   "frmRegEmpleado.frx":6550
               PICN            =   "frmRegEmpleado.frx":656C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin Proyecto1.chameleonButton cmdformedu 
               Height          =   330
               Left            =   7200
               TabIndex        =   238
               ToolTipText     =   "Visualizar Historial de Sueldos"
               Top             =   2280
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   582
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
               MICON           =   "frmRegEmpleado.frx":69AE
               PICN            =   "frmRegEmpleado.frx":69CA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSForms.ComboBox cboEstPrin 
               Height          =   315
               Left            =   3360
               TabIndex        =   240
               Top             =   2520
               Width           =   2055
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "3625;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtAnhoEgr 
               Height          =   255
               Left            =   3360
               TabIndex        =   236
               Top             =   2160
               Width           =   2055
               VariousPropertyBits=   746604571
               Size            =   "3625;450"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboCarrera 
               Height          =   255
               Left            =   3360
               TabIndex        =   235
               Top             =   1800
               Width           =   4275
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "7541;450"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboNomInst 
               Height          =   255
               Left            =   3360
               TabIndex        =   234
               Top             =   1440
               Width           =   4275
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "7541;450"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboTipInst 
               Height          =   255
               Left            =   3360
               TabIndex        =   233
               Top             =   1080
               Width           =   4275
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "7541;450"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboFormSup 
               Height          =   315
               Left            =   3360
               TabIndex        =   222
               Top             =   120
               Width           =   4275
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "7541;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label75 
               BackStyle       =   0  'Transparent
               Caption         =   "Formación Educativa Principal:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   239
               Top             =   2520
               Width           =   3135
            End
            Begin VB.Label Label74 
               BackStyle       =   0  'Transparent
               Caption         =   "Privada"
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
               Left            =   5400
               TabIndex        =   232
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label Label73 
               BackStyle       =   0  'Transparent
               Caption         =   "Pública"
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
               Left            =   3720
               TabIndex        =   231
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label Label72 
               BackStyle       =   0  'Transparent
               Caption         =   "No"
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
               Left            =   5400
               TabIndex        =   230
               Top             =   480
               Width           =   525
            End
            Begin VB.Label Label71 
               BackStyle       =   0  'Transparent
               Caption         =   "Si"
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
               Left            =   3720
               TabIndex        =   229
               Top             =   480
               Width           =   525
            End
            Begin VB.Label Label70 
               BackStyle       =   0  'Transparent
               Caption         =   "Año de Egreso:"
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
               Left            =   120
               TabIndex        =   226
               Top             =   2280
               Width           =   1365
            End
            Begin VB.Label Label69 
               BackStyle       =   0  'Transparent
               Caption         =   "Carrera:"
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
               Left            =   120
               TabIndex        =   225
               Top             =   1920
               Width           =   3165
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre de la  Institución Educativa:"
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
               Left            =   120
               TabIndex        =   224
               Top             =   1560
               Width           =   3165
            End
            Begin VB.Label Label68 
               BackStyle       =   0  'Transparent
               Caption         =   "Régimen de la Institución Educativa:"
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
               Left            =   120
               TabIndex        =   223
               Top             =   840
               Width           =   3165
            End
            Begin VB.Label Label48 
               BackStyle       =   0  'Transparent
               Caption         =   "Estudió en una institución del Perú:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   480
               Width           =   3135
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Formación Superior Completa:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   240
               Width           =   2775
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Institución Educativa:"
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
               Left            =   120
               TabIndex        =   113
               Top             =   1200
               Width           =   2685
            End
         End
         Begin MSMask.MaskEdBox txtEmpleado 
            Height          =   315
            Left            =   120
            TabIndex        =   84
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            ForeColor       =   128
            MaxLength       =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtRuc 
            Height          =   315
            Left            =   7410
            TabIndex        =   85
            Top             =   6600
            Visible         =   0   'False
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            ForeColor       =   128
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtEdad 
            Height          =   315
            Left            =   3930
            TabIndex        =   6
            Top             =   1075
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtNumHijos 
            Height          =   315
            Left            =   9330
            TabIndex        =   27
            Top             =   1080
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   1
            Mask            =   "#"
            PromptChar      =   " "
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H009F5539&
            Caption         =   "Contacto de Emergencia"
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
            Height          =   2115
            Left            =   6120
            TabIndex        =   108
            Top             =   1320
            Width           =   7695
            Begin MSForms.TextBox txtemailApo 
               Height          =   315
               Left            =   840
               TabIndex        =   248
               Top             =   1680
               Width           =   4305
               VariousPropertyBits=   746604571
               MaxLength       =   150
               Size            =   "7594;556"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboTipoDocPar 
               Height          =   315
               Left            =   840
               TabIndex        =   213
               Top             =   1320
               Width           =   795
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "1402;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboParenContacto 
               Height          =   315
               Left            =   3840
               TabIndex        =   212
               Top             =   1320
               Width           =   1275
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "2249;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtNroDocParen 
               Height          =   285
               Left            =   1680
               TabIndex        =   209
               Top             =   1320
               Width           =   1065
               VariousPropertyBits=   746604571
               MaxLength       =   20
               Size            =   "1879;503"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtMovilApo 
               Height          =   255
               Left            =   3360
               TabIndex        =   31
               Top             =   960
               Width           =   1755
               VariousPropertyBits=   746604571
               MaxLength       =   12
               Size            =   "3096;450"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtFijoApo 
               Height          =   255
               Left            =   810
               TabIndex        =   30
               Top             =   960
               Width           =   1905
               VariousPropertyBits=   746604571
               MaxLength       =   12
               Size            =   "3360;450"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtDirApo 
               Height          =   255
               Left            =   810
               TabIndex        =   29
               Top             =   600
               Width           =   4305
               VariousPropertyBits=   746604571
               MaxLength       =   150
               Size            =   "7594;450"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtApoderado 
               Height          =   255
               Left            =   840
               TabIndex        =   28
               Top             =   240
               Width           =   4305
               VariousPropertyBits=   746604571
               MaxLength       =   150
               Size            =   "7594;450"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label76 
               BackStyle       =   0  'Transparent
               Caption         =   "Email:"
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
               Height          =   195
               Left            =   120
               TabIndex        =   247
               Top             =   1800
               Width           =   585
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "Parentesco:"
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
               Height          =   195
               Left            =   2760
               TabIndex        =   211
               Top             =   1320
               Width           =   1035
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Doc.:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   210
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label64 
               BackStyle       =   0  'Transparent
               Caption         =   "Móvil:"
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
               Height          =   195
               Left            =   2790
               TabIndex        =   129
               Top             =   960
               Width           =   555
            End
            Begin VB.Label Label52 
               BackStyle       =   0  'Transparent
               Caption         =   "Telf.:"
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
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label51 
               BackStyle       =   0  'Transparent
               Caption         =   "Direc:"
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
               Height          =   195
               Left            =   120
               TabIndex        =   110
               Top             =   600
               Width           =   585
            End
            Begin VB.Label Label50 
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
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   120
               TabIndex        =   109
               Top             =   210
               Width           =   705
            End
         End
         Begin MSMask.MaskEdBox txtCalzado 
            Height          =   315
            Left            =   2940
            TabIndex        =   22
            Top             =   4965
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "##.#"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtEstatura 
            Height          =   315
            Left            =   990
            TabIndex        =   21
            Top             =   4920
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "#.##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.DTPicker dtpNacimiento 
            Height          =   315
            Left            =   1020
            TabIndex        =   5
            Top             =   1075
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   123928577
            CurrentDate     =   38637
         End
         Begin MSMask.MaskEdBox txtPeso 
            Height          =   315
            Left            =   990
            TabIndex        =   24
            Top             =   5280
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "###.##"
            PromptChar      =   " "
         End
         Begin MSForms.TextBox txtmailper 
            Height          =   315
            Left            =   960
            TabIndex        =   250
            Top             =   4200
            Width           =   5055
            VariousPropertyBits=   746604571
            Size            =   "8916;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtPasaporte 
            Height          =   315
            Left            =   990
            TabIndex        =   18
            Top             =   4575
            Width           =   1965
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "3466;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtFonoFijo 
            Height          =   315
            Left            =   1020
            TabIndex        =   15
            Top             =   3427
            Width           =   1965
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "3466;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtmail 
            Height          =   315
            Left            =   1020
            TabIndex        =   17
            Top             =   3819
            Width           =   4995
            VariousPropertyBits=   746604571
            Size            =   "8811;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtTipoDoc 
            Height          =   315
            Left            =   1020
            TabIndex        =   10
            Top             =   1859
            Width           =   525
            VariousPropertyBits=   746604571
            MaxLength       =   2
            Size            =   "926;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDireccion 
            Height          =   315
            Left            =   1020
            TabIndex        =   12
            Top             =   2251
            Width           =   4995
            VariousPropertyBits=   746604571
            MaxLength       =   250
            Size            =   "8811;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDistrito 
            Height          =   315
            Left            =   1020
            TabIndex        =   14
            Top             =   3035
            Width           =   855
            VariousPropertyBits=   746604571
            MaxLength       =   6
            Size            =   "1508;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDpto 
            Height          =   315
            Left            =   1020
            TabIndex        =   13
            Top             =   2640
            Width           =   855
            VariousPropertyBits=   746604571
            MaxLength       =   3
            Size            =   "1508;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboNacion 
            Height          =   315
            Left            =   1020
            TabIndex        =   8
            Top             =   1467
            Width           =   1965
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3466;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtCarnetExt 
            Height          =   315
            Left            =   3960
            TabIndex        =   9
            Top             =   1470
            Width           =   2085
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "3678;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBrevete 
            Height          =   285
            Left            =   4800
            TabIndex        =   20
            Top             =   4590
            Width           =   1185
            VariousPropertyBits=   746604571
            MaxLength       =   12
            Size            =   "2090;503"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtFonoMov 
            Height          =   315
            Left            =   3870
            TabIndex        =   16
            Top             =   3427
            Width           =   2145
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "3784;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboTMame 
            Height          =   285
            Left            =   2910
            TabIndex        =   25
            Top             =   5340
            Width           =   1815
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3201;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtNombre2 
            Height          =   285
            Left            =   1440
            TabIndex        =   2
            Top             =   720
            Width           =   1545
            VariousPropertyBits=   746604571
            ForeColor       =   128
            MaxLength       =   65
            Size            =   "2725;503"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.ComboBox cboGSanguineo 
            Height          =   285
            Left            =   4800
            TabIndex        =   23
            Top             =   4980
            Width           =   1185
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2090;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboBreveteCat 
            Height          =   285
            Left            =   3840
            TabIndex        =   19
            Top             =   4590
            Width           =   915
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "1614;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboEstCivil 
            Height          =   315
            Left            =   6960
            TabIndex        =   26
            Top             =   1080
            Width           =   1755
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3096;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtNumDoc 
            Height          =   315
            Left            =   3930
            TabIndex        =   11
            Top             =   1859
            Width           =   2085
            VariousPropertyBits=   746604571
            MaxLength       =   11
            Size            =   "3678;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboGenero 
            Height          =   315
            Left            =   4860
            TabIndex        =   7
            Top             =   1075
            Width           =   1155
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2037;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtApePat 
            Height          =   285
            Left            =   3120
            TabIndex        =   3
            Top             =   720
            Width           =   2955
            VariousPropertyBits=   746604571
            ForeColor       =   128
            MaxLength       =   65
            Size            =   "5212;503"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtApeMat 
            Height          =   285
            Left            =   6090
            TabIndex        =   4
            Top             =   720
            Width           =   2715
            VariousPropertyBits=   746604571
            ForeColor       =   128
            MaxLength       =   65
            Size            =   "4789;503"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtNombre1 
            Height          =   285
            Left            =   30
            TabIndex        =   1
            Top             =   720
            Width           =   1455
            VariousPropertyBits=   746604571
            ForeColor       =   128
            MaxLength       =   65
            Size            =   "2566;503"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label77 
            BackStyle       =   0  'Transparent
            Caption         =   " E-mail Per:"
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
            Height          =   255
            Left            =   0
            TabIndex        =   249
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label54 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   246
            Top             =   5640
            Width           =   735
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "HCM:"
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
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   204
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "N° C.E.:"
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
            Left            =   3090
            TabIndex        =   149
            Top             =   1512
            Width           =   765
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   " Nación.:"
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
            Left            =   0
            TabIndex        =   148
            Top             =   1515
            Width           =   885
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Pasap N°:"
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
            Left            =   60
            TabIndex        =   122
            Top             =   4620
            Width           =   1245
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Brevete:"
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
            Left            =   3060
            TabIndex        =   121
            Top             =   4620
            Width           =   735
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "G. Sangre:"
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
            Left            =   3810
            TabIndex        =   120
            Top             =   5010
            Width           =   1095
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "T. Calzado:"
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
            Left            =   1680
            TabIndex        =   119
            Top             =   5010
            Width           =   1065
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "T. Mameluco:"
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
            Left            =   1680
            TabIndex        =   118
            Top             =   5400
            Width           =   1185
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "Peso:"
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
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   117
            Top             =   5400
            Width           =   735
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Estatura:"
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
            Left            =   60
            TabIndex        =   116
            Top             =   5010
            Width           =   795
         End
         Begin VB.Label lblTipoDoc 
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
            Left            =   1590
            TabIndex        =   106
            Top             =   1859
            Width           =   1395
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   " Doc. Ide.:"
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
            Left            =   0
            TabIndex        =   105
            Top             =   1905
            Width           =   915
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "N° DI:"
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
            Left            =   3090
            TabIndex        =   104
            Top             =   1904
            Width           =   645
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   " E-mail:"
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
            Height          =   255
            Left            =   0
            TabIndex        =   103
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Móvil:"
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
            Left            =   3090
            TabIndex        =   102
            Top             =   3472
            Width           =   585
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   " Telef.:"
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
            Left            =   0
            TabIndex        =   101
            Top             =   3465
            Width           =   645
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   " Dirección:"
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
            Left            =   0
            TabIndex        =   100
            Top             =   2295
            Width           =   1035
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "E. Civil:"
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
            Left            =   6060
            TabIndex        =   99
            Top             =   1155
            Width           =   675
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo:"
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
            Left            =   4350
            TabIndex        =   98
            Top             =   1120
            Width           =   495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Edad:"
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
            Left            =   3090
            TabIndex        =   97
            Top             =   1120
            Width           =   615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Fec. Nac.:"
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
            Left            =   90
            TabIndex        =   96
            Top             =   1120
            Width           =   915
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Hijos:"
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
            Left            =   8820
            TabIndex        =   95
            Top             =   1155
            Width           =   465
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   " Distrito:"
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
            Left            =   0
            TabIndex        =   94
            Top             =   3075
            Width           =   795
         End
         Begin VB.Label lblDistrito 
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
            Left            =   1920
            TabIndex        =   93
            Top             =   3035
            Width           =   4095
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   " Dpto:"
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
            Left            =   0
            TabIndex        =   92
            Top             =   2685
            Width           =   765
         End
         Begin VB.Label lblDpto 
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
            Left            =   1920
            TabIndex        =   91
            Top             =   2643
            Width           =   4095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ape. Mat."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6930
            TabIndex        =   90
            Top             =   450
            Width           =   885
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Ape. Pat."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4020
            TabIndex        =   89
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s)"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1110
            TabIndex        =   88
            Top             =   510
            Width           =   945
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   9480
            TabIndex        =   87
            Top             =   780
            Width           =   105
         End
         Begin VB.Label lblEmpleado 
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
            Height          =   195
            Left            =   13890
            TabIndex        =   86
            Top             =   390
            Visible         =   0   'False
            Width           =   45
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H009F5539&
         Caption         =   "Retenciones Judiciales"
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
         Height          =   6615
         Left            =   -75000
         TabIndex        =   154
         Top             =   600
         Width           =   13995
         Begin VB.CheckBox chkmovi 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1800
            TabIndex        =   220
            Top             =   6000
            Width           =   255
         End
         Begin VB.TextBox txtMontoMov 
            Height          =   285
            Left            =   3960
            TabIndex        =   219
            Top             =   6000
            Width           =   1575
         End
         Begin NOVAdmin.flxEdit flxret 
            Height          =   4275
            Left            =   0
            TabIndex        =   194
            Top             =   240
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   7541
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CellFontName    =   "MS Sans Serif"
            CellFontSize    =   8.25
            BackColorSel    =   -2147483643
            BackColorFixed  =   9868950
            CellPicture     =   "frmRegEmpleado.frx":72E2
            ColAlignment0   =   9
            FixedAlignment0 =   9
            ForeColorSel    =   16711680
            ForeColorFixed  =   14474460
            MouseIcon       =   "frmRegEmpleado.frx":72FE
            RowHeight0      =   240
         End
         Begin MSForms.ComboBox cboMonMovi 
            Height          =   285
            Left            =   2160
            TabIndex        =   218
            Top             =   6000
            Width           =   1680
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2963;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movilidad"
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
            Height          =   195
            Left            =   840
            TabIndex        =   221
            Top             =   6000
            Width           =   825
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Monto    S/."
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
            Height          =   195
            Left            =   5130
            TabIndex        =   158
            Top             =   4890
            Width           =   1545
         End
         Begin VB.Label lbltotalmto 
            Alignment       =   1  'Right Justify
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
            Left            =   6705
            TabIndex        =   157
            Top             =   4830
            Width           =   1410
         End
         Begin VB.Label lbltotalpor 
            Alignment       =   2  'Center
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
            Left            =   3945
            TabIndex        =   156
            Top             =   4830
            Width           =   795
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total %"
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
            Height          =   195
            Left            =   3255
            TabIndex        =   155
            Top             =   4890
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H009F5539&
         Caption         =   "Generales"
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
         Height          =   1005
         Left            =   -74970
         TabIndex        =   42
         Top             =   690
         Width           =   13965
         Begin MSComCtl2.DTPicker dtpCese 
            Height          =   285
            Left            =   2910
            TabIndex        =   80
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   123928577
            CurrentDate     =   38637
         End
         Begin MSComCtl2.DTPicker dtpIngreso 
            Height          =   285
            Left            =   660
            TabIndex        =   81
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   123928577
            CurrentDate     =   38637
         End
         Begin MSForms.TextBox txtAsigFam 
            Height          =   285
            Left            =   11385
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   60
            VariousPropertyBits=   746604571
            Size            =   "106;503"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkJubil 
            DataSource      =   "cb"
            Height          =   285
            Left            =   7890
            TabIndex        =   147
            Top             =   600
            Width           =   1125
            BackColor       =   10442041
            ForeColor       =   16777215
            DisplayStyle    =   4
            Size            =   "1984;503"
            Value           =   "0"
            Caption         =   "Jubilado"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CheckBox chkSctr 
            DataSource      =   "cb"
            Height          =   285
            Left            =   3765
            TabIndex        =   141
            Top             =   600
            Width           =   1065
            BackColor       =   10442041
            ForeColor       =   16777215
            DisplayStyle    =   4
            Size            =   "1879;503"
            Value           =   "0"
            Caption         =   "SENATI"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CheckBox chkSvl 
            DataSource      =   "cb"
            Height          =   285
            Left            =   660
            TabIndex        =   123
            Top             =   600
            Width           =   795
            BackColor       =   10442041
            ForeColor       =   16777215
            DisplayStyle    =   4
            Size            =   "1402;503"
            Value           =   "0"
            Caption         =   "S.V.L"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.ComboBox cboPersonal 
            Height          =   285
            Left            =   5280
            TabIndex        =   126
            Top             =   240
            Width           =   1530
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2699;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboTipo 
            Height          =   285
            Left            =   9600
            TabIndex        =   124
            Top             =   600
            Width           =   1635
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2884;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkAsigFam 
            DataSource      =   "cb"
            Height          =   285
            Left            =   6900
            TabIndex        =   79
            Top             =   600
            Width           =   945
            BackColor       =   10442041
            ForeColor       =   16777215
            DisplayStyle    =   4
            Size            =   "1667;503"
            Value           =   "0"
            Caption         =   "A. Fam"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.ComboBox cboSituacion 
            Height          =   285
            Left            =   9600
            TabIndex        =   44
            Top             =   240
            Width           =   1635
            VariousPropertyBits=   746604569
            BackColor       =   14737632
            DisplayStyle    =   7
            Size            =   "2884;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboCategoria 
            Height          =   285
            Left            =   6885
            TabIndex        =   146
            Top             =   240
            Width           =   1725
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3043;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblnumsctr 
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
            Left            =   4845
            TabIndex        =   193
            Top             =   585
            Width           =   1965
         End
         Begin VB.Label lblnumsvl 
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
            Left            =   1695
            TabIndex        =   192
            Top             =   585
            Width           =   1965
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Personal:"
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
            Left            =   4455
            TabIndex        =   127
            Top             =   285
            Width           =   810
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   9150
            TabIndex        =   125
            Top             =   645
            Width           =   450
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. Cese:"
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
            Left            =   2190
            TabIndex        =   47
            Top             =   285
            Width           =   720
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ing:"
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
            Left            =   60
            TabIndex        =   46
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Situación:"
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
            Left            =   8730
            TabIndex        =   43
            Top             =   285
            Width           =   870
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H009F5539&
         Caption         =   "CTS"
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
         Height          =   1260
         Left            =   -69840
         TabIndex        =   70
         Top             =   1740
         Width           =   8865
         Begin MSForms.TextBox txtBancoCTS 
            Height          =   315
            Left            =   990
            TabIndex        =   60
            Top             =   180
            Width           =   645
            VariousPropertyBits=   746604571
            MaxLength       =   2
            Size            =   "1138;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox CboMonCTS 
            Height          =   315
            Left            =   990
            TabIndex        =   61
            Top             =   525
            Width           =   1860
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3281;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtCTSCta 
            Height          =   285
            Left            =   3615
            TabIndex        =   62
            Top             =   540
            Width           =   2385
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "4207;503"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro:"
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
            Left            =   3060
            TabIndex        =   82
            Top             =   600
            Width           =   435
         End
         Begin VB.Label lblBancoCTS 
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
            Left            =   1680
            TabIndex        =   73
            Top             =   180
            Width           =   4320
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
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
            Height          =   315
            Left            =   210
            TabIndex        =   72
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda:"
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
            Left            =   210
            TabIndex        =   71
            Top             =   615
            Width           =   765
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H009F5539&
         Height          =   5670
         Left            =   -74955
         TabIndex        =   107
         Top             =   645
         Width           =   13935
         Begin NOVAdmin.flxEdit flxDependientes 
            Height          =   5505
            Left            =   0
            TabIndex        =   197
            Top             =   120
            Width           =   13845
            _ExtentX        =   19764
            _ExtentY        =   9710
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CellFontName    =   "MS Sans Serif"
            CellFontSize    =   8.25
            BackColorSel    =   -2147483643
            BackColorFixed  =   9868950
            CellPicture     =   "frmRegEmpleado.frx":731A
            ColAlignment0   =   9
            FixedAlignment0 =   9
            ForeColorSel    =   16711680
            ForeColorFixed  =   14474460
            MouseIcon       =   "frmRegEmpleado.frx":7336
            RowHeight0      =   240
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H009F5539&
         Caption         =   "Seguro Médico"
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
         Height          =   1965
         Left            =   -69840
         TabIndex        =   51
         Top             =   3030
         Width           =   8865
         Begin NOVAdmin.flxEdit msfseg 
            Height          =   1095
            Left            =   90
            TabIndex        =   196
            Top             =   450
            Width           =   5355
            _ExtentX        =   9393
            _ExtentY        =   2037
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CellFontName    =   "MS Sans Serif"
            CellFontSize    =   8.25
            BackColorSel    =   -2147483635
            BackColorFixed  =   -2147483633
            CellPicture     =   "frmRegEmpleado.frx":7352
            ConfirmarBorradoLinea=   0   'False
            ColWidth0       =   960
            ColAlignment0   =   9
            FixedAlignment0 =   9
            ColWidth1       =   960
            ColAlignment1   =   9
            FixedAlignment1 =   9
            ForeColorSel    =   -2147483634
            ForeColorFixed  =   -2147483630
            GridColorFixed  =   12632256
            MouseIcon       =   "frmRegEmpleado.frx":736E
            RowHeight0      =   240
            RowHeight1      =   240
         End
         Begin MSForms.CheckBox chkver 
            DataSource      =   "cb"
            Height          =   390
            Left            =   6300
            TabIndex        =   190
            Top             =   900
            Width           =   870
            BackColor       =   10442041
            ForeColor       =   16777215
            DisplayStyle    =   4
            Size            =   "1535;688"
            Value           =   "0"
            Caption         =   "Activo"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.ComboBox cboGenFlx 
         Height          =   315
         Left            =   -65520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1020
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H009F5539&
         Caption         =   "Cuentas Bancarias"
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
         Height          =   1305
         Left            =   -74970
         TabIndex        =   53
         Top             =   1710
         Width           =   5100
         Begin MSForms.TextBox txtNumCtaMn 
            Height          =   315
            Left            =   2910
            TabIndex        =   57
            Top             =   570
            Width           =   1995
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "3519;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboTipCtaMn 
            Height          =   315
            Left            =   960
            TabIndex        =   56
            Top             =   570
            Width           =   1905
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3360;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboTipCtaMe 
            Height          =   315
            Left            =   960
            TabIndex        =   58
            Top             =   930
            Width           =   1905
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3360;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtNumCtaMe 
            Height          =   315
            Left            =   2910
            TabIndex        =   59
            Top             =   930
            Width           =   1995
            VariousPropertyBits=   746604571
            MaxLength       =   20
            Size            =   "3519;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBanco 
            Height          =   315
            Left            =   960
            TabIndex        =   55
            Top             =   210
            Width           =   615
            VariousPropertyBits=   1015040027
            MaxLength       =   2
            Size            =   "1085;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nacional"
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
            Height          =   195
            Left            =   60
            TabIndex        =   68
            Top             =   630
            Width           =   765
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Extranjera"
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
            Height          =   195
            Left            =   60
            TabIndex        =   65
            Top             =   960
            Width           =   870
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
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
            Height          =   195
            Left            =   60
            TabIndex        =   63
            Top             =   240
            Width           =   555
         End
         Begin VB.Label lblBanco 
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
            Left            =   1605
            TabIndex        =   54
            Top             =   210
            Width           =   3300
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H009F5539&
         Caption         =   "Fondo de Pensiones"
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
         Height          =   1995
         Left            =   -74970
         TabIndex        =   48
         Top             =   3015
         Width           =   5115
         Begin MSComCtl2.DTPicker dtpFecIns 
            Height          =   315
            Left            =   3315
            TabIndex        =   67
            Top             =   630
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   123928577
            CurrentDate     =   38637
         End
         Begin MSForms.OptionButton optComiAFP 
            Height          =   285
            Index           =   1
            Left            =   2370
            TabIndex        =   215
            Top             =   1230
            Width           =   1935
            BackColor       =   10442041
            ForeColor       =   -2147483643
            DisplayStyle    =   5
            Size            =   "3413;503"
            Value           =   "0"
            Caption         =   "Comisión Mixta"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton optComiAFP 
            Height          =   285
            Index           =   0
            Left            =   210
            TabIndex        =   214
            Top             =   1200
            Width           =   1845
            BackColor       =   10442041
            ForeColor       =   -2147483643
            DisplayStyle    =   5
            Size            =   "3254;503"
            Value           =   "0"
            Caption         =   "Comisión x Flujo"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtNumAfp 
            Height          =   315
            Left            =   435
            TabIndex        =   66
            Top             =   660
            Width           =   1845
            VariousPropertyBits=   746604571
            MaxLength       =   14
            Size            =   "3254;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtAfp 
            Height          =   315
            Left            =   435
            TabIndex        =   64
            Top             =   210
            Width           =   735
            VariousPropertyBits=   746604571
            MaxLength       =   2
            Size            =   "1296;556"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ins."
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
            Left            =   2385
            TabIndex        =   128
            Top             =   690
            Width           =   915
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F.P"
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
            Height          =   195
            Left            =   90
            TabIndex        =   52
            Top             =   270
            Width           =   300
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro"
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
            Height          =   195
            Left            =   75
            TabIndex        =   50
            Top             =   720
            Width           =   315
         End
         Begin VB.Label lblAfp 
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
            Left            =   1215
            TabIndex        =   49
            Top             =   210
            Width           =   3690
         End
      End
      Begin VB.Frame FrCont 
         BackColor       =   &H009F5539&
         Caption         =   "Contrato"
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
         Height          =   2265
         Left            =   -74970
         TabIndex        =   41
         Top             =   4980
         Width           =   13995
         Begin VB.TextBox lblcargoC 
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
            Height          =   285
            Left            =   7470
            TabIndex        =   203
            Top             =   1800
            Width           =   3525
         End
         Begin VB.TextBox txtcargoC 
            Height          =   285
            Left            =   6660
            TabIndex        =   202
            Top             =   1800
            Width           =   795
         End
         Begin VB.Frame frmsueldos 
            BackColor       =   &H009F5539&
            Height          =   1935
            Left            =   1560
            TabIndex        =   153
            Top             =   0
            Visible         =   0   'False
            Width           =   4155
            Begin NOVAdmin.flxEdit flxsueldos 
               Height          =   1935
               Left            =   0
               TabIndex        =   195
               Top             =   120
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   3413
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CellFontName    =   "MS Sans Serif"
               CellFontSize    =   8.25
               BackColorSel    =   -2147483635
               BackColorFixed  =   -2147483633
               CellPicture     =   "frmRegEmpleado.frx":738A
               ConfirmarBorradoLinea=   0   'False
               ColWidth0       =   960
               ColAlignment0   =   9
               FixedAlignment0 =   9
               ColWidth1       =   960
               ColAlignment1   =   9
               FixedAlignment1 =   9
               ForeColorSel    =   -2147483634
               ForeColorFixed  =   -2147483630
               GridColorFixed  =   12632256
               MouseIcon       =   "frmRegEmpleado.frx":73A6
               RowHeight0      =   240
               RowHeight1      =   240
            End
         End
         Begin VB.Frame frcese 
            BackColor       =   &H009F5539&
            Caption         =   "Otros Tipos de Cese del Empleado"
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
            Height          =   990
            Left            =   5160
            TabIndex        =   164
            Top             =   0
            Visible         =   0   'False
            Width           =   5430
            Begin MSComCtl2.DTPicker dtfechacese 
               Height          =   315
               Left            =   1110
               TabIndex        =   167
               Top             =   600
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648384
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   123928577
               CurrentDate     =   38637
            End
            Begin MSForms.ComboBox Cbotipocese 
               Height          =   315
               Left            =   1110
               TabIndex        =   165
               Top             =   240
               Width           =   4260
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "7514;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fec. Cese:"
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
               Height          =   195
               Left            =   60
               TabIndex        =   168
               Top             =   270
               Width           =   930
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipos Cese:"
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
               Height          =   195
               Left            =   -450
               TabIndex        =   166
               Top             =   600
               Width           =   1020
            End
         End
         Begin Proyecto1.chameleonButton btnContrato 
            Height          =   345
            Left            =   11490
            TabIndex        =   74
            ToolTipText     =   "Generar Contrato"
            Top             =   240
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   "Contrato"
            ENAB            =   0   'False
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
            MICON           =   "frmRegEmpleado.frx":73C2
            PICN            =   "frmRegEmpleado.frx":73DE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnRenovar 
            Height          =   345
            Left            =   11490
            TabIndex        =   75
            ToolTipText     =   "Renovar Contrato"
            Top             =   690
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   "Renovar"
            ENAB            =   0   'False
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
            MICON           =   "frmRegEmpleado.frx":9EE8
            PICN            =   "frmRegEmpleado.frx":9F04
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnHabilitar 
            Height          =   345
            Left            =   12660
            TabIndex        =   76
            ToolTipText     =   "Habilitar Contrato"
            Top             =   210
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   "Habilitar"
            ENAB            =   0   'False
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
            MICON           =   "frmRegEmpleado.frx":A05E
            PICN            =   "frmRegEmpleado.frx":A07A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnGrabarC 
            Height          =   345
            Left            =   11760
            TabIndex        =   77
            ToolTipText     =   "Guardar"
            Top             =   1260
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
            MICON           =   "frmRegEmpleado.frx":A1D4
            PICN            =   "frmRegEmpleado.frx":A1F0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnCancelCont 
            Height          =   345
            Left            =   12660
            TabIndex        =   78
            ToolTipText     =   "Cancelar Contrato"
            Top             =   690
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   "Cancelar"
            ENAB            =   0   'False
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
            MICON           =   "frmRegEmpleado.frx":A632
            PICN            =   "frmRegEmpleado.frx":A64E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton cmdotros 
            Height          =   345
            Left            =   12660
            TabIndex        =   163
            ToolTipText     =   "Cesar al Empleado"
            Top             =   1230
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   "Ceses"
            ENAB            =   0   'False
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
            MICON           =   "frmRegEmpleado.frx":A7A8
            PICN            =   "frmRegEmpleado.frx":A7C4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame frmcont 
            BackColor       =   &H009F5539&
            BorderStyle     =   0  'None
            Height          =   2850
            Left            =   120
            TabIndex        =   172
            Top             =   240
            Width           =   5445
            Begin MSComCtl2.DTPicker dtpFin 
               Height          =   315
               Left            =   3720
               TabIndex        =   173
               Top             =   360
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12640511
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   123928577
               CurrentDate     =   38637
            End
            Begin MSComCtl2.DTPicker dtpInicio 
               Height          =   315
               Left            =   960
               TabIndex        =   174
               Top             =   360
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               CalendarBackColor=   12648384
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   123928577
               CurrentDate     =   38637
            End
            Begin Proyecto1.chameleonButton cmdver 
               Height          =   330
               Left            =   2535
               TabIndex        =   175
               ToolTipText     =   "Visualizar Historial de Sueldos"
               Top             =   1020
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   582
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
               MICON           =   "frmRegEmpleado.frx":AC16
               PICN            =   "frmRegEmpleado.frx":AC32
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSForms.ComboBox cboContratos 
               Height          =   315
               Left            =   2640
               TabIndex        =   183
               Top             =   0
               Width           =   2625
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "4630;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtContrato 
               Height          =   285
               Left            =   840
               TabIndex        =   181
               Top             =   0
               Width           =   1095
               VariousPropertyBits=   746604571
               ForeColor       =   128
               MaxLength       =   4
               Size            =   "1931;503"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.ComboBox cboMonSueldo 
               Height          =   315
               Left            =   900
               TabIndex        =   180
               Top             =   690
               Width           =   1590
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "2805;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtMontoBono 
               Height          =   315
               Left            =   3735
               TabIndex        =   179
               Top             =   1365
               Width           =   1590
               VariousPropertyBits=   746604575
               Size            =   "2805;556"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboMonBono 
               Height          =   285
               Left            =   1860
               TabIndex        =   178
               Top             =   1380
               Width           =   1680
               VariousPropertyBits=   746604571
               DisplayStyle    =   7
               Size            =   "2963;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.CheckBox chkBono 
               DataSource      =   "cb"
               Height          =   285
               Left            =   -45
               TabIndex        =   177
               Top             =   1395
               Width           =   1845
               VariousPropertyBits=   1015023643
               BackColor       =   10442041
               ForeColor       =   16777215
               DisplayStyle    =   4
               Size            =   "3254;503"
               Value           =   "0"
               Caption         =   "Bonos de Campo:"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "S. Básico:"
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
               Height          =   195
               Left            =   -30
               TabIndex        =   189
               Top             =   1095
               Width           =   885
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "F. Término:"
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
               Height          =   195
               Left            =   2640
               TabIndex        =   188
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "F. Inicio:"
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
               Height          =   195
               Left            =   0
               TabIndex        =   187
               Top             =   480
               Width           =   765
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número:"
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
               Height          =   195
               Left            =   0
               TabIndex        =   186
               Top             =   30
               Width           =   720
            End
            Begin VB.Label Label37 
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   2040
               TabIndex        =   185
               Tag             =   " "
               Top             =   30
               Width           =   525
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda:"
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
               Height          =   195
               Left            =   -30
               TabIndex        =   184
               Tag             =   " "
               Top             =   750
               Width           =   750
            End
            Begin VB.Label lblEstContrato 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ULTIMO CONTRATO APROBADO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   225
               Left            =   1155
               TabIndex        =   182
               Top             =   1740
               Width           =   3105
            End
            Begin VB.Label lblsbasico 
               Alignment       =   2  'Center
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
               Height          =   285
               Left            =   900
               TabIndex        =   176
               Top             =   1050
               Width           =   1590
            End
         End
         Begin VB.TextBox txtdivc 
            Height          =   285
            Left            =   6660
            TabIndex        =   199
            Top             =   1140
            Width           =   765
         End
         Begin VB.TextBox lbldivc 
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
            Height          =   285
            Left            =   7470
            TabIndex        =   200
            Top             =   1140
            Width           =   3555
         End
         Begin MSForms.ComboBox CboTipIng 
            Height          =   315
            Left            =   9360
            TabIndex        =   217
            Top             =   480
            Width           =   1635
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2884;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtcencos 
            Height          =   330
            Left            =   6600
            TabIndex        =   161
            Top             =   795
            Width           =   1305
            VariousPropertyBits=   746604571
            Size            =   "2302;582"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cbodiv 
            Height          =   315
            Left            =   6660
            TabIndex        =   150
            Top             =   1455
            Width           =   4350
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "7673;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboEstTrabajo 
            Height          =   315
            Left            =   6660
            TabIndex        =   145
            Top             =   120
            Width           =   4350
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "7673;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboHorLab 
            Height          =   315
            Left            =   6660
            TabIndex        =   143
            Top             =   480
            Width           =   2205
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3889;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboModal 
            Height          =   315
            Left            =   11280
            TabIndex        =   140
            Top             =   1350
            Visible         =   0   'False
            Width           =   165
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "291;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   8880
            TabIndex        =   216
            Top             =   600
            Width           =   885
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   5520
            TabIndex        =   201
            Top             =   1860
            Width           =   570
         End
         Begin VB.Label lblcencos 
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
            Height          =   330
            Left            =   7920
            TabIndex        =   162
            Top             =   795
            Width           =   3060
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cen. Costo:"
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
            Height          =   195
            Left            =   5550
            TabIndex        =   160
            Top             =   863
            Width           =   1005
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Div. Local:"
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
            Height          =   165
            Left            =   5550
            TabIndex        =   159
            Top             =   1215
            Width           =   1035
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Div. HCM:"
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
            Height          =   195
            Left            =   5550
            TabIndex        =   151
            Top             =   1515
            Width           =   885
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. Trabajo:"
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
            Height          =   195
            Index           =   0
            Left            =   5550
            TabIndex        =   144
            Top             =   195
            Width           =   1110
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Hor. Lab.:"
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
            Height          =   195
            Left            =   5550
            TabIndex        =   142
            Top             =   528
            Width           =   885
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H009F5539&
         BorderStyle     =   0  'None
         Height          =   2040
         Left            =   -74970
         TabIndex        =   152
         Top             =   5190
         Width           =   11295
      End
   End
   Begin Proyecto1.chameleonButton btnSalir 
      Height          =   345
      Left            =   12150
      TabIndex        =   34
      Top             =   7965
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmRegEmpleado.frx":B54A
      PICN            =   "frmRegEmpleado.frx":B566
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnEliminar 
      Height          =   345
      Left            =   3960
      TabIndex        =   36
      ToolTipText     =   "Eliminar"
      Top             =   7920
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Eliminar"
      ENAB            =   0   'False
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
      MICON           =   "frmRegEmpleado.frx":B92C
      PICN            =   "frmRegEmpleado.frx":B948
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnModificar 
      Height          =   345
      Left            =   2640
      TabIndex        =   37
      ToolTipText     =   "Modificar"
      Top             =   7920
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Modificar"
      ENAB            =   0   'False
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
      MICON           =   "frmRegEmpleado.frx":BAA2
      PICN            =   "frmRegEmpleado.frx":BABE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnCancelar 
      Height          =   345
      Left            =   5610
      TabIndex        =   38
      ToolTipText     =   "Deshacer"
      Top             =   7965
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmRegEmpleado.frx":BEEC
      PICN            =   "frmRegEmpleado.frx":BF08
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   8820
      Top             =   9030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Proyecto1.chameleonButton btnGrabar 
      Height          =   345
      Left            =   6060
      TabIndex        =   32
      ToolTipText     =   "Guardar"
      Top             =   7965
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Grabar"
      ENAB            =   0   'False
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
      MICON           =   "frmRegEmpleado.frx":C44A
      PICN            =   "frmRegEmpleado.frx":C466
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnNuevo 
      Height          =   345
      Left            =   1320
      TabIndex        =   69
      ToolTipText     =   "Nuevo"
      Top             =   7920
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Nuevo"
      ENAB            =   0   'False
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
      MICON           =   "frmRegEmpleado.frx":C8A8
      PICN            =   "frmRegEmpleado.frx":C8C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H009F5539&
      Height          =   465
      Left            =   9045
      TabIndex        =   130
      Top             =   7485
      Width           =   3345
      Begin Proyecto1.chameleonButton btnUltimo 
         Height          =   285
         Left            =   2280
         TabIndex        =   131
         ToolTipText     =   "Ultimo"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRegEmpleado.frx":CC2E
         PICN            =   "frmRegEmpleado.frx":CC4A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnPrimero 
         Height          =   285
         Left            =   420
         TabIndex        =   132
         ToolTipText     =   "Primero"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRegEmpleado.frx":CFCC
         PICN            =   "frmRegEmpleado.frx":CFE8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnSgt 
         Height          =   285
         Left            =   1860
         TabIndex        =   133
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRegEmpleado.frx":D36D
         PICN            =   "frmRegEmpleado.frx":D389
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnPrevio 
         Height          =   285
         Left            =   840
         TabIndex        =   134
         ToolTipText     =   "Previo"
         Top             =   150
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRegEmpleado.frx":D6F5
         PICN            =   "frmRegEmpleado.frx":D711
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblPrimero 
         Alignment       =   2  'Center
         BackColor       =   &H009F5539&
         Caption         =   "1"
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
         Left            =   90
         TabIndex        =   137
         Top             =   180
         Width           =   345
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H009F5539&
         Caption         =   "100"
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
         Height          =   195
         Left            =   2670
         TabIndex        =   136
         Top             =   180
         Width           =   495
      End
      Begin VB.Label lblCuenta 
         Alignment       =   2  'Center
         BackColor       =   &H009F5539&
         Caption         =   "50"
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
         Left            =   1260
         TabIndex        =   135
         Top             =   180
         Width           =   465
      End
   End
   Begin Proyecto1.chameleonButton cmdElimSoli 
      Height          =   345
      Left            =   9150
      TabIndex        =   170
      ToolTipText     =   "Eliminar Solicitudes"
      Top             =   7980
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmRegEmpleado.frx":DA7D
      PICN            =   "frmRegEmpleado.frx":DA99
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
      Left            =   11640
      TabIndex        =   198
      Top             =   7980
      Visible         =   0   'False
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
      MICON           =   "frmRegEmpleado.frx":DEDB
      PICN            =   "frmRegEmpleado.frx":DEF7
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
      Height          =   360
      Left            =   14310
      TabIndex        =   208
      Top             =   3660
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   635
      BTYPE           =   14
      TX              =   "&Adjuntar Archivos"
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
      MICON           =   "frmRegEmpleado.frx":E439
      PICN            =   "frmRegEmpleado.frx":E455
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnGenAsiento 
      Height          =   585
      Left            =   14040
      TabIndex        =   252
      ToolTipText     =   "Modificar"
      Top             =   7680
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "Generar Asiento Practicantes/Proveedores"
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
      MICON           =   "frmRegEmpleado.frx":EDDF
      PICN            =   "frmRegEmpleado.frx":EDFB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnActualizarBenef 
      Height          =   585
      Left            =   14040
      TabIndex        =   253
      ToolTipText     =   "Modificar"
      Top             =   6720
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1032
      BTYPE           =   14
      TX              =   "Actualizar Pract/RetensJud"
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
      MICON           =   "frmRegEmpleado.frx":F229
      PICN            =   "frmRegEmpleado.frx":F245
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox txtBusqueda 
      Height          =   315
      Left            =   3375
      TabIndex        =   139
      Top             =   7485
      Width           =   2055
      VariousPropertyBits=   746604571
      Size            =   "3625;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboBusqueda 
      Height          =   315
      Left            =   1320
      TabIndex        =   138
      Top             =   7485
      Width           =   2025
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3572;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblautoriza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POR AUTORIZAR: CESE 28/02/2009 - TERMINO CONTRATO / SUELDO : 5000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   5475
      TabIndex        =   169
      Top             =   7455
      Width           =   3570
   End
   Begin VB.Label lblSituacEmp 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
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
      Left            =   9900
      TabIndex        =   45
      Top             =   8070
      Width           =   1620
   End
   Begin VB.Label lblModo 
      BackStyle       =   0  'Transparent
      Caption         =   "Acción"
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
      Height          =   225
      Left            =   9750
      TabIndex        =   39
      Top             =   7875
      Visible         =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "frmRegEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConsulta As New FrmConsultas
Dim mFileSize As Long
Dim arrByte() As Byte
Private strFoto As String
Private tmp
Private intNumHijos As Integer
Private generocod As Boolean
Private rsgral As MYSQL_RS
Private Parentesco(0 To 7) As String
Dim FlgCont As Boolean
Dim sbasico As Double
Dim ComisionAFP As String
Dim OptPeruflg As String
Dim OptRegEduflg As String


Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub ConfiguracboTipo()
    CboTipIng.Clear
    CboTipIng.AddItem "Fijo"
    CboTipIng.AddItem "Variable"
End Sub

Private Sub btnActualizarBenef_Click()
    FrmPrac.Show
End Sub

Private Sub btnCancelar_Click()
    ConfigurarBotones cfgCancelar
    If txtEmpleado.Enabled Then
        txtEmpleado.SetFocus
    Else
        txtNombre1.SetFocus
    End If
End Sub

Private Sub TiposContratos(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim Rs As MYSQL_RS
    Dim I As Integer
    SQL = "Select * from cncontrato"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = "00"
    I = 1
    Do While Not Rs.EOF
        cbo.AddItem CE(Rs.Fields("DESCRIP"))
        cbo.List(I, 1) = CE(Rs.Fields("CODIGO"))
        I = I + 1
        Rs.MoveNext
    Loop
    cbo.ListIndex = 0
    Set Rs = Nothing
End Sub

Private Sub Categoria(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim Rs As MYSQL_RS
    Dim I As Integer
    SQL = "Select * from rh_categoria order by descrip"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    cbo.Clears
    'cbo.AddItem "Seleccionar..."
    'cbo.List(0, 1) = "00"
    I = 0
    Do While Not Rs.EOF
        cbo.AddItem CE(Rs.Fields("DESCRIP"))
        cbo.List(I, 1) = CE(Rs.Fields("CODIGO"))
        I = I + 1
        Rs.MoveNext
    Loop
    cbo.ListIndex = 0
    Set Rs = Nothing
End Sub

Private Sub btnCancelCont_Click()
    Dim SQL As String
    Dim RES As Integer
    Dim UsuAceptado As Boolean
    UsuAceptado = False
    RES = MsgBox("¿Está seguro que desea cancelar el Contrato del empleado " & vbNewLine & " con código Nro. " & Trim(txtEmpleado) & " ?", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        If cboSituacion.List(cboSituacion.ListIndex, 1) = 0 Then
            SQL = " Update contrato set estado = '" & CANCELADO & "' where codigo = '" & Trim(txtContrato) & "'" & _
                  " and codemp = '" & Trim(txtEmpleado) & "'"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        Else
            SQL = " Update contrato set estado = '" & CANCELADO & "',f_termino = '" & Format(Date, "yyyy/mm/dd") & "' where codigo = '" & Trim(txtContrato) & "'" & _
                  " and codemp = '" & Trim(txtEmpleado) & "'"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        End If
        Consulta
        MoverPosRs Trim(txtEmpleado)
        CargarDatos rsgral
        ModoFormulario modConsulta
        btnContrato.Enabled = False
        btnRenovar.Enabled = False
        btnCancelCont.Enabled = False
    End If
    
    
End Sub

Private Sub btnContrato_Click()
    txtContrato = GenNumContrato(txtEmpleado)
    txtContrato.tag = PENDIENTE
    txtContrato.Locked = True
    txtContrato.BackColor = ColorDeshabilitado
    btnGrabarC.Enabled = True
    With flxsueldos
        .TextMatrix(.Rows - 1, 0) = val(.Rows - 1)
        .TextMatrix(.Rows - 1, 1) = FormatNumber(lblsbasico, 2)
    End With
End Sub

Private Sub Eliminar(CodEmp As String)
    Dim SQL As String
    SQL = "Delete from cnauxil where codigo = '" & CodEmp & "' and auxiliar = '3'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    SQL = "Call Delete_Cont('" & txtEmpleado & "');"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    SQL = "Call Delete_Familiares ('" & CodEmp & "');"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    SQL = "Call Delete_Empleado ('" & CodEmp & "');"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, True
End Sub

Private Sub btnEliminar_Click()
    Dim RES As Integer
    RES = MsgBox("¿Esta seguro que desea Eliminar al empleado, " & vbNewLine & " con código Nro. " & Trim(txtEmpleado) & " ?", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        Eliminar Trim(txtEmpleado)
        ModoFormulario modAccion
        txtBusqueda = Empty
    End If
End Sub



Private Sub btnEliminarDoc_Click()
 Dim SQL As String
 
 With flxformEduEmp
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 10) = strChecked Then
             SQL = " Delete from pl_empformedu where id=" & _
                    "'" & .TextMatrix(I, 9) & "' "
              oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
            End If
        Next
 End With
 
 MsgBox "Formación Educativa Eliminada. ", vbOKOnly, gsNomSW
End Sub

Private Sub btnGenAsiento_Click()
    Dim folioRef As String
    Dim folioPago As String
    Dim flagtipo As String

    folioRef = InputBox("Ingrese la Fecha de Generación del Voucher  p.e:20130613 ", "Voucher Terceros", "20130613")
    If folioRef = "" Then Exit Sub
    folioPago = InputBox("Ingrese el Nro de la Orden de Pago  p.e:2013060127 ", "Voucher Terceros", "2013060127")
    If folioPago = "" Then Exit Sub
    flagtipo = InputBox("Ingrese el Flag de Tipo de Voucher  p.e:[P] si es practicantes,[R] si es ReJud ", "Voucher Terceros", "P")
    If flagtipo = "" Then Exit Sub
                 
    GenerarfolioPract folioRef, folioPago, flagtipo
End Sub

Private Sub btnGrabar_Click()
    Dim RES As Integer
    Dim SQL As String
    If lblModo = "Nuevo" Then
        If Grabar Then
            Consulta
            MoverPosRs Trim(txtEmpleado)
            CargarDatos rsgral
            ModoFormulario modConsulta
            RES = MsgBox("¿Desea Generar Contrato al empleado " & vbNewLine & " con código Nro. " & Trim(txtEmpleado) & " ?", vbQuestion + vbYesNo, gsNomSW)
            If RES = 6 Then
                BloqueoControles True
                ConfigurarBotones cfgNuevo
                btnContrato.Enabled = True
                SSTab1.Tab = 2
                btnGrabar.Enabled = False
                btnContrato.SetFocus
            End If
        End If
        
    End If
    If lblModo = "Modificar" Then
        If Actualizar Then
            Consulta
            MoverPosRs Trim(txtEmpleado)
            CargarDatos rsgral
            ModoFormulario modConsulta
            txtBusqueda = Empty
        End If
    End If
End Sub


Private Sub VerFoto(di As String)
    On Error GoTo NADA
    Dim root As String
    
    root = "\\SRVPERFS01\OT-PER-PUBLIC$\NOV\Empleados\" & di & ".jpg"
    imgFoto.Picture = LoadPicture(root)
     
NADA:
    Exit Sub
End Sub


Private Sub MoverPosRs(codigo As String)
    Dim I As Integer
    Dim Pos As Integer
    If rsgral.State = MY_RS_OPEN Then
        For I = 1 To rsgral.RecordCount
            Do While Not rsgral.EOF
                If rsgral.Fields("CODIGO") = codigo Then
                     I = rsgral.AbsolutePosition
                     rsgral.AbsolutePosition = I
                     ConfigBtnsBusq I, rsgral.RecordCount
                     Exit Sub
                End If
                rsgral.MoveNext
            Loop
        Next
    End If
End Sub

Private Function Grabar() As Boolean
    Dim SQL As String
    Dim I As Integer
    Dim FecNac As String
    Dim FecIng As String
    Dim FecCese As String
    Dim FecIngAfp As String
    Dim ComisionAFP As String

    
    Grabar = False
    If Validar Then
        If strFoto = "" Then
            strFoto = "0"
        End If
        FecNac = IIf(Not IsDate(dtpNacimiento.Value), Empty, dtpNacimiento.Value)
        FecIng = IIf(Not IsDate(dtpIngreso.Value), Empty, dtpIngreso.Value)
        FecCese = IIf(Not IsDate(dtpCese.Value), Empty, dtpCese.Value)
        FecIngAfp = IIf(Not IsDate(dtpFecIns.Value), Empty, dtpFecIns.Value)
        ComisionAFP = IIf(optComiAFP(0).Value = True, "F", "M")
         
        SQL = "Call Insert_Emp ('" & Trim(txtEmpleado) & "','" & Right("000" & Trim(txtcargo), 3) & "'," & _
              "'','','0','" & cboTipo.List(cboTipo.ListIndex, 1) & _
              "','00','" & cboCategoria.List(cboCategoria.ListIndex, 1) & "','" & Right("00" & Trim(txtTipoDoc), 2) & "'," & _
              "'" & cboPersonal.List(cboPersonal.ListIndex, 1) & "','" & Trim(txtNumDoc) & "','" & _
              Trim(txtCarnetExt) & "','" & Trim(txtPasaporte) & "','" & cboBreveteCat.List(cboBreveteCat.ListIndex, 1) & "','" & _
              Trim(txtBrevete) & "','" & Trim(txtNombre1) & "','" & Trim(txtNombre2) & "','" & _
              Trim(txtApePat) & "','" & Trim(txtApeMat) & "','" & Format(FecNac, "yyyy/mm/dd") & "'," & Right("00" & Trim(txtEdad.Text), 2) & _
              ",'" & cboGenero.List(cboGenero.ListIndex, 1) & "','" & _
              cboGSanguineo.List(cboGSanguineo.ListIndex, 1) & "','" & cboEstCivil.List(cboEstCivil.ListIndex, 1) & "'," & _
              Right("0" & Trim(txtNumHijos.Text), 1) & ",'" & Trim(txtDireccion) & _
              "','" & Right("000000" & Trim(txtDistrito), 6) & "'," & "'" & txtDpto & "'," & _
              " '" & cboNacion.List(cboNacion.ListIndex, 1) & "','" & Trim(txtFonoFijo.Text) & _
              "','" & Trim(txtFonoMov.Text) & "','" & Trim(txtmail) & "','" & Trim(txtmailper) & "'," & Trim(txtEstatura.Text) & "," & Trim(txtPeso) & "," & _
              " " & Trim(txtCalzado.Text) & ",'" & cboTMame.List(cboTMame.ListIndex, 1) & "'," & _
              " " & strFoto & ",'" & IIf(chkAsigFam.Value = True, "S", "N") & "','" & IIf(chkJubil.Value = True, "S", "N") & "'" & _
              ",'" & IIf(chkSctr.Value = True, "S", "N") & "','" & IIf(chkSvl.Value = True, "S", "N") & "','00','' " & _
              ",'" & Right("00" & Trim(txtAfp), 2) & "', " & "'" & Trim(txtNumAfp) & _
              "','" & Right("00" & Trim(txtBanco), 2) & "', '" & cboTipCtaMn.List(cboTipCtaMn.ListIndex, 1) & _
              "','" & Trim(txtNumCtaMn) & "', '" & cboTipCtaMe.List(cboTipCtaMe.ListIndex, 1) & _
              "','" & Trim(txtNumCtaMe) & "','" & _
              "','" & Right("00" & Trim(txtBancoCTS), 2) & "'," & "'" & IIf(CboMonCTS.ListIndex <> 0, CboMonCTS.List(CboMonCTS.ListIndex, 1), "0") & _
              "','" & Trim(txtCTSCta) & "','" & Format(FecIng, "yyyy/mm/dd") & "','" & Format(FecCese, "yyyy/mm/dd") & "','" & Trim(txtApoderado) & "','" & Trim(txtDirApo) & "','" & Trim(txtFijoApo) & "','" & Trim(txtMovilApo) & "', " & _
              "'" & Format(FecIngAfp, "yyyy/mm/dd") & "','" & Trim(txtHCMEmpleado.Text) & "','" & ComisionAFP & "');"
        If oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, False) = True Then
              SQL = " Insert into cnauxil(AUXILIAR,CODIGO,RUC,DESCRIP,DIRECC,DISTRI,EMAIL,TELE1,TELE2,TIPO,AGT_RETEN,TIPCTA_MN,NUMCTA_MN,TIPCTA_ME,NUMCTA_ME)" & _
                    " values ('3','" & Trim(txtEmpleado) & "','" & Trim(txtEmpleado) & "', '" & Trim(txtApePat) & " " & Trim(txtApeMat) & " " & Trim(txtNombre1) & " " & Trim(txtNombre2) & "' ," & _
                    " '" & Trim(txtDireccion) & "', '" & Right("000000" & Trim(txtDistrito), 6) & "'," & _
                    " '" & Trim(txtEmail) & "', '" & Trim(txtFonoFijo) & "'," & _
                    " '" & Trim(txtFonoMov) & "','N','N','" & cboTipCtaMn.List(cboTipCtaMn.ListIndex, 1) & _
                    "','" & Trim(txtNumCtaMn) & "', '" & cboTipCtaMe.List(cboTipCtaMe.ListIndex, 1) & _
                    "','" & Trim(txtNumCtaMe) & "')"
              oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
              
              
              SQL = " Insert into empleado_apoderado(codemp,NomApo,DirecApo,FonoApo,MovilApo,tipodocApo,nrodocApo,relacionApo,emailApo)" & _
                    " values ('" & Trim(txtEmpleado) & "','" & Trim(txtApoderado) & "', '" & Trim(txtDirApo) & "', '" & Trim(txtFijoApo) & "', '" & Trim(txtMovilApo) & "', '" & cboTipoDocPar.List(cboTipoDocPar.ListIndex, 1) & "' ," & _
                    "'" & Trim(txtNroDocParen) & "','" & cboParenContacto.List(cboParenContacto.ListIndex, 1) & "','" & Trim(txtemailApo) & "')"
              oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        End If
        With flxDependientes
            For I = 1 To .Rows - 1
                If .TextMatrix(I, 0) <> Empty Then
                    SQL = "Call Insert_Familiares ('" & .TextMatrix(I, 0) & "', '" & txtEmpleado & "'," & _
                            " '" & .TextMatrix(I, 10) & "','" & .TextMatrix(I, 7) & "', '" & .TextMatrix(I, 8) & "','" & .TextMatrix(I, 1) & "'," & _
                            " '" & .TextMatrix(I, 2) & "', '" & .TextMatrix(I, 3) & "', '" & Format(.TextMatrix(I, 4), "yyyy/mm/dd") & "', " & _
                            " '" & .TextMatrix(I, 5) & "', '" & gen(.TextMatrix(I, 6)) & "', 'S');"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                    SQL = "insert into rh_verificaessalud (codemp,item,verificado,fecha,familiar) values " & _
                          "('" & txtEmpleado & "'," & .TextMatrix(I, 0) & ",'" & IIf(.TextMatrix(I, 11) = strChecked, "S", "N") & "', " & _
                          "'" & Format(Date, "dd/mm/yyyy") & "','F')"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                End If
            Next
        End With
        With flxret
            For I = 1 To .Rows - 1
                If val(.TextMatrix(I, 3)) <> 0 Or val(.TextMatrix(I, 4)) <> 0 Then
                    SQL = "Call Insert_Retenciones ('" & Trim(.TextMatrix(I, 0)) & "','" & Trim(txtEmpleado) & "'," & _
                          " '" & Trim(.TextMatrix(I, 1)) & "'," & CDbl(IIf(val(.TextMatrix(I, 3)) = 0, 0, .TextMatrix(I, 3))) & ",'" & .TextMatrix(I, 2) & "'," & _
                          " " & CDbl(IIf(val(.TextMatrix(I, 4)) = 0, 0, .TextMatrix(I, 4))) & ");"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                End If
            Next
        End With
        
        With msfseg
            EnumerarItems msfseg
            For I = 1 To .Rows - 1
                If Trim(.TextMatrix(I, 1)) <> "" Then
                    If I = 1 Then
                        SQL = "update empleado set codseg='" & Trim(.TextMatrix(I, 1)) & "',numseg='" & Trim(.TextMatrix(I, 3)) & "' where codigo ='" & Trim(txtEmpleado) & "'"
                        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
                    End If
                    SQL = "Call Insert_Seguros ('" & Trim(txtEmpleado) & "'," & Trim(.TextMatrix(I, 0)) & "," & _
                          "'" & Trim(.TextMatrix(I, 1)) & "','" & Trim(.TextMatrix(I, 3)) & "','" & Format(Trim(.TextMatrix(I, 4)), "yyyy/mm/dd") & "'," & _
                          " '" & Format(Trim(.TextMatrix(I, 5)), "yyyy/mm/dd") & "');"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                End If
            Next
        End With
        SQL = "insert into rh_verificaessalud (codemp,item,verificado,fecha,familiar) values ('" & txtEmpleado & "',0,'" & IIf(chkver.Value = True, "S", "N") & "','" & Format(Date, "dd/mm/yyyy") & "','T')"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        Grabar = True
    End If
End Function


Private Sub contrato(CodEmp As String, contrato As String)
    Dim SQL As String
    Dim rscont As MYSQL_RS
    Dim Bono As String
    Dim RES As Integer, UsuAcepFec As Boolean, UsuAcepSue As Boolean
    Dim FecCese As String, esta As String
    Dim contAnterior As String
    SQL = "Select * from contrato where codemp = '" & CodEmp & "' and codigo = '" & contrato & "'"
    Set rscont = oConexion.EjecutaSelectRS(SQL)
    If chkBono.Value = True Then
        Bono = "S"
    Else
        Bono = "N"
    End If
    FecCese = IIf(Not IsDate(dtpFin.Value), Empty, Format(dtpFin, "yyyy/mm/dd"))
    If rscont.RecordCount = 1 Then
        If ValidaContr Then
            UsuAcepFec = False
            UsuAcepSue = False
            If IsDate(dtfechacese) Then
                If InStr(1, lblautoriza, "CESE") = 0 Then
                    If UsuarioAceptado(4, strUsuarioId, "cesar al empleado", txtEmpleado, 0, Format(dtfechacese.Value, "yyyy/mm/dd"), Cbotipocese.List(Cbotipocese.ListIndex, 1), "") = True Then
                        UsuAcepFec = True
                    End If
                End If
            End If
            If sbasico > 0 Then
                If InStr(1, lblautoriza, "SUELDO") = 0 Then
                    If UsuarioAceptado(7, strUsuarioId, "modificar sueldo al empleado", txtEmpleado, sbasico, "", "", "") = True Then
                        UsuAcepSue = True
                        lblsbasico = sbasico
                    End If
                End If
            End If
            If UsuAcepSue = True Then
                Sueldos
            End If
            
            SQL = " Call Update_Cont ('" & contrato & "','" & cboContratos.List(cboContratos.ListIndex, 1) & "', '" & CodEmp & "'," & _
                  " '" & Format(dtpInicio.Value, "yyyy/mm/dd") & "', '" & FecCese & "'," & _
                  " " & CDbl(lblsbasico) & ", '" & Bono & "' , " & CDbl(txtMontoBono) & "," & _
                  " '" & cboMonSueldo.List(cboMonSueldo.ListIndex, 1) & _
                  "', '" & IIf(chkBono.Value = True, cboMonBono.List(cboMonBono.ListIndex, 1), "") & _
                  "','" & cboEstTrabajo.List(cboEstTrabajo.ListIndex, 1) & "','" & _
                  cboHorLab.List(cboHorLab.ListIndex, 0) & "','" & Trim(txtcencos) & "', " & _
                  "'" & cboDiv.List(cboDiv.ListIndex, 1) & "','" & Trim(txtdivc) & "','" & Trim(txtcargoC) & "','" & DameTipoSuel(CboTipIng.Value) & "');"
            
            If IsDate(dtpFin.Value) And Format(dtpFin.Value, "yyyy/mm/dd") <= Format(Date, "yyyy/mm/dd") Then
                If Format(dtpFin.Value, "yyyy/mm/dd") >= Format(dtpInicio.Value, "yyyy/mm/dd") Then
                    oConexionMYSQL.Execute "Update contrato set estado='CA' where codemp='" & CodEmp & "' and codigo='" & contrato & "'"
                Else
                    MsgBox "No se actualizó el estado de contrato porque" & vbNewLine & _
                           " la fecha de término no puede ser menor a la fecha de inicio", vbOKOnly + vbExclamation, "NOVPeru"
                End If
            Else
                If Format(dtpInicio.Value, "yyyy/mm/dd") > Format(Date, "yyyy/mm/dd") And IsDate(dtpFin.Value) Then
                    oConexionMYSQL.Execute "Update contrato set estado='PN' where codemp='" & CodEmp & "' and codigo='" & contrato & "'"
                End If
            End If
            ContinuidadContratos CodEmp
            If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Modificar, False) Then
            Else
                If FecCese < Format(Date, "yyyy/mm/dd") And FecCese <> "" Then
                Else
                    If Trim(txtContrato.Text) <> "CN01" Then
                        oConexionMYSQL.Execute "Update contrato set estado='AP' where codemp='" & CodEmp & "' and codigo='" & contrato & "'"
                    ElseIf Trim(txtContrato.Text) = "CN01" And lblEstContrato.Caption <> "ULTIMO CONTRATO APROBADO" Then
                        oConexionMYSQL.Execute "Update contrato set estado='PN' where codemp='" & CodEmp & "' and codigo='" & contrato & "'"
                    End If
                End If
                btnGrabar.Enabled = False
                btnContrato.Enabled = False
            End If
            If IsDate(dtfechacese.Value) And Cbotipocese.ListIndex > -1 And UsuAcepFec = True Then
                oConexionMYSQL.Execute "Update contrato set fechacese='" & Format(dtfechacese.Value, "yyyy/mm/dd") & "',codtipocese = '" & Cbotipocese.List(Cbotipocese.ListIndex, 1) & "' where codemp='" & CodEmp & "' and codigo='" & contrato & "'"
                SQL = " Update empleado set SITUACION = '0', FEC_CESE= '" & Format(dtfechacese.Value, "yyyy/mm/dd") & "'" & _
                      " WHERE codigo='" & Trim(CodEmp) & "'"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                If lblEstContrato.Caption <> "ULTIMO CONTRATO CANCELADO" Then
                    If MsgBox("¿Desea cancelar su contrato?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                        oConexionMYSQL.Execute "Update contrato set estado='CA' where codemp='" & CodEmp & "' and codigo='" & contrato & "'"
                    End If
                End If
            Else
                oConexionMYSQL.Execute "Update contrato set fechacese='',codtipocese='' where codemp='" & CodEmp & "' and codigo='" & contrato & "'"
            End If
            If Trim(txtContrato.tag) = "PN" Then
                oConexionMYSQL.Execute "delete from rh_tempacceso where codigo = 5 and usuario = '" & strUsuarioId & "' and codemp='" & CodEmp & "' and autorizado = 'S'"
            End If
            Consulta
            MoverPosRs Trim(CodEmp)
            CargarDatos rsgral
        End If
    End If
    If rscont.RecordCount = 0 And Trim(txtContrato) <> Empty Then
        If ValidaContr Then
            SQL = " Call Insert_Cont ('" & contrato & "', '" & strAnoSistema & strMesSistema & "' , '" & cboContratos.List(cboContratos.ListIndex, 1) & "'," & _
                  " '" & CodEmp & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") & "', '" & Format(dtpFin.Value, "yyyy/mm/dd") & "'," & _
                  " " & CDbl(lblsbasico) & " ,1, '" & Trim(Bono) & "' , " & CDbl(txtMontoBono) & ", '" & txtContrato.tag & "'," & _
                  " '" & cboMonSueldo.List(cboMonSueldo.ListIndex, 1) & "', '" & IIf(chkBono.Value = True, cboMonBono.List(cboMonBono.ListIndex, 1), "") & _
                  "','" & cboEstTrabajo.List(cboEstTrabajo.ListIndex, 1) & "','" & cboHorLab.List(cboHorLab.ListIndex, 0) & "', " & _
                  "'" & Trim(txtcencos) & "','" & cboDiv.List(cboDiv.ListIndex, 1) & "','" & Trim(txtdivc) & "','','','" & Trim(txtcargoC) & "','" & DameTipoSuel(CboTipIng.Value) & "');"
            If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, False) Then
                MsgBox "El Contrato no se guardó correctamente", vbInformation, gsNomSW
            Else
                ContinuidadContratos CodEmp
                contAnterior = Right("00" & Trim(val(Right(Trim(contrato), 2)) - 1), 2)
                If contAnterior <> "00" Then
                    SQL = " Update contrato set Estado = '" & CANCELADO & "' and F_TERMINO='" & Format(Date, "yyyy/dd/mm") & "'" & _
                          " where codigo= '" & Left(Trim(contrato), 2) & contAnterior & "' and codemp = '" & CodEmp & "'"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
                End If
                If Trim(txtContrato) <> "CN01" Then
                    SQL = " Update empleado set SITUACION = '1', FEC_CESE= '" & IIf(FecCese < Date, FecCese, "") & "'" & _
                          " where codigo = '" & Trim(CodEmp) & "'"
                    If oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, False) Then
                        Sueldos
                        btnGrabar.Enabled = False
                        btnContrato.Enabled = False
                        btnHabilitar.Enabled = False
                        btnRenovar.Enabled = False
                    End If
                End If
            End If
            If Trim(txtContrato) = "CN01" Then
                Consulta
                MoverPosRs Trim(CodEmp)
                CargarDatos rsgral
                btnGrabar.Enabled = False
                btnContrato.Enabled = False
                btnHabilitar.Enabled = False
                btnRenovar.Enabled = False
            Else
                If Trim(txtContrato.tag) = "PN" Then
                    oConexionMYSQL.Execute "delete from rh_tempacceso where codigo = 5 and usuario = '" & strUsuarioId & "' and codemp='" & CodEmp & "' and autorizado = 'S'"
                End If
                RES = MsgBox("¿Desea habilitar el Contrato del empleado " & vbNewLine & " con código Nro. " & Trim(txtEmpleado) & " ?", vbQuestion + vbYesNo, gsNomSW)
                If RES = 6 Then
                    btnHabilitar.Enabled = True
                    btnHabilitar_Click
                Else
                    Consulta
                    MoverPosRs Trim(CodEmp)
                    CargarDatos rsgral
                    btnGrabar.Enabled = False
                    btnContrato.Enabled = False
                    btnHabilitar.Enabled = False
                    btnRenovar.Enabled = False
                End If
            End If
        End If
    End If
    Set rscont = Nothing
End Sub

Sub Sueldos()
    Dim SQL As String
    SQL = "Call Delete_Sueldos ('" & Trim(txtEmpleado) & "','" & Trim(txtContrato) & "');"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    With flxsueldos
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 1) <> Empty Then
                SQL = "Call Insert_Sueldos ('" & Trim(txtEmpleado) & "', '" & Trim(txtContrato) & "'," & _
                      "" & I & ",'" & cboMonSueldo.List(cboMonSueldo.ListIndex, 1) & "', " & CDbl(.TextMatrix(I, 1)) & ", " & _
                      "'" & IIf(.TextMatrix(I, 2) = "__/__/____" Or .TextMatrix(I, 2) = "", "", Format(Trim(.TextMatrix(I, 2)), "yyyy/mm/dd")) & "');"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
            End If
        Next
        CargarSueldos Trim(txtEmpleado), Trim(txtContrato)
    End With
End Sub

Private Function ValidaContr() As Boolean
    ValidaContr = True
    If IsDate(meFecIni) And IsDate(meFecFin) Then
        If meFecIni <> "  /  /    " And meFecFin <> "  /  /    " Then
            If CDate(meFecIni) > CDate(meFecFin) Then
                ValidaContr = False
                MsgBox "La Fecha de inicio del contrato no puede ser mayor que la de cese", vbInformation, gsNomSW
                meFecIni.SetFocus
                Exit Function
            End If
        End If
    End If
    If lblsbasico = Empty Then ValidaContr = False: MsgBox "Debe Ingresar el Sueldo Básico del empleado", vbInformation, gsNomSW: SSTab1.Tab = 2: Exit Function
    If cboHorLab.ListIndex = 0 Then ValidaContr = False: MsgBox "Debe Ingresar el horario laboral del empleado", vbInformation, gsNomSW: SSTab1.Tab = 2: Exit Function ' cboHorLab.SetFocus: Exit Function
    If chkBono.Value = True Then
        If txtMontoBono = Empty Then ValidaContr = False: MsgBox "Debe Ingresar el Monto del Bono de campo", vbInformation, gsNomSW: SSTab1.Tab = 2: txtMontoBono.SetFocus: Exit Function
        If cboMonBono.List(cboMonBono.ListIndex, 0) = "Seleccionar..." Then
            ValidaContr = False
            MsgBox "Debe Seleccionar la moneda correspondiente al Monto del Bono de Campo", vbInformation, gsNomSW
            If cboMonBono.Enabled = True Then
                cboMonBono.SetFocus
            End If
            Exit Function
        End If
    End If
    If chkBono.Value = False Then
        If txtMontoBono = Empty Then txtMontoBono = "0.00"
    End If
End Function

Private Function gen(s As String) As String
    If s = "Femenino" Then gen = "F": Exit Function
    If s = "Masculino" Then gen = "M": Exit Function
    If UCase(s) = "F" Then gen = "Femenino": Exit Function
    If UCase(s) = "M" Then gen = "Masculino"
End Function

Private Function Actualizar() As Boolean
    Dim SQL As String
    Dim I As Integer
    Dim FecNac As String
    Dim FecIng As String
    Dim FecCese As String
    Dim FecIngAfp As String
    Dim ComisionAFP As String
    Actualizar = False
    If Validar Then
        If strFoto <> "" Then
            If Left(strFoto, 2) <> "0x" Then
                strFoto = ToHex(strFoto)
            End If
        End If
        FecNac = IIf(Not IsDate(dtpNacimiento.Value), Empty, dtpNacimiento.Value)
        FecIng = IIf(Not IsDate(dtpIngreso.Value), Empty, dtpIngreso.Value)
        FecCese = IIf(Not IsDate(dtpCese.Value), Empty, dtpCese.Value)
        FecIngAfp = IIf(Not IsDate(dtpFecIns.Value), Empty, dtpFecIns.Value)
        ComisionAFP = IIf(optComiAFP(0).Value = True, "F", "M")
        SQL = "Call Update_Emp ('" & Trim(txtEmpleado) & "','" & Right("000" & Trim(txtcargo), 3) & "'," & _
              "'','','" & cboSituacion.List(cboSituacion.ListIndex, 1) & "','" & cboTipo.List(cboTipo.ListIndex, 1) & _
              "','00','" & cboCategoria.List(cboCategoria.ListIndex, 1) & "','" & Right("00" & Trim(txtTipoDoc), 2) & "'," & _
              "'" & cboPersonal.List(cboPersonal.ListIndex, 1) & "','" & Trim(txtNumDoc) & "','" & _
              Trim(txtCarnetExt) & "','" & Trim(txtPasaporte) & "','" & cboBreveteCat.List(cboBreveteCat.ListIndex, 1) & "','" & _
              Trim(txtBrevete) & "','" & Trim(txtNombre1) & "','" & Trim(txtNombre2) & "','" & _
              Trim(txtApePat) & "','" & Trim(txtApeMat) & "','" & Format(FecNac, "yyyy/mm/dd") & "'," & Right("00" & Trim(txtEdad.Text), 2) & _
              ",'" & cboGenero.List(cboGenero.ListIndex, 1) & "','" & _
              cboGSanguineo.List(cboGSanguineo.ListIndex, 1) & "','" & cboEstCivil.List(cboEstCivil.ListIndex, 1) & "'," & _
              Right("0" & Trim(txtNumHijos.Text), 1) & ",'" & Trim(txtDireccion) & _
              "','" & Right("000000" & Trim(txtDistrito), 6) & "'," & "'" & txtDpto & "'," & _
              " '" & cboNacion.List(cboNacion.ListIndex, 1) & "','" & Trim(txtFonoFijo.Text) & _
              "','" & Trim(txtFonoMov.Text) & "','" & Trim(txtmail) & "','" & Trim(txtmailper) & "'," & Trim(txtEstatura.Text) & "," & Trim(txtPeso) & "," & _
              " " & Trim(txtCalzado.Text) & ",'" & cboTMame.List(cboTMame.ListIndex, 1) & "'," & _
              " '" & strFoto & "','" & IIf(chkAsigFam.Value = True, "S", "N") & "','" & IIf(chkJubil.Value = True, "S", "N") & "'" & _
              ",'" & IIf(chkSctr.Value = True, "S", "N") & "','" & IIf(chkSvl.Value = True, "S", "N") & "','00','' " & _
              ",'" & Right("00" & Trim(txtAfp), 2) & "', " & "'" & Trim(txtNumAfp) & _
              "','" & Right("00" & Trim(txtBanco), 2) & "', '" & cboTipCtaMn.List(cboTipCtaMn.ListIndex, 1) & _
              "','" & Trim(txtNumCtaMn) & "', '" & cboTipCtaMe.List(cboTipCtaMe.ListIndex, 1) & _
              "','" & Trim(txtNumCtaMe) & "','" & _
              "','" & Right("00" & Trim(txtBancoCTS), 2) & "'," & "'" & IIf(CboMonCTS.ListIndex <> 0, CboMonCTS.List(CboMonCTS.ListIndex, 1), "0") & _
              "','" & Trim(txtCTSCta) & "','" & Format(FecIng, "yyyy/mm/dd") & "','" & Format(FecCese, "yyyy/mm/dd") & "','" & Trim(txtApoderado) & "','" & Trim(txtDirApo) & "','" & Trim(txtFijoApo) & "','" & Trim(txtMovilApo) & "', " & _
              "'" & Format(FecIngAfp, "yyyy/mm/dd") & "','" & Trim(txtHCMEmpleado) & "','" & ComisionAFP & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
        
        SQL = "Call Update_ApoderadoEmp ('" & Trim(txtEmpleado) & "','" & txtApoderado & "'," & _
              "'" & txtDirApo & "','" & txtFijoApo & "','" & txtMovilApo & "','" & cboTipoDocPar.List(cboTipoDocPar.ListIndex, 1) & _
              "','" & txtNroDocParen & "','" & cboParenContacto.List(cboParenContacto.ListIndex, 1) & "','" & txtemailApo & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
        
        SQL = "Call Delete_Familiares ('" & Trim(txtEmpleado) & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
        SQL = "delete from rh_verificaessalud where codemp = '" & Trim(txtEmpleado) & "' and familiar = 'F'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
        With flxDependientes
            For I = 1 To .Rows - 1
                 If .TextMatrix(I, 0) <> Empty Then
                    SQL = "Call Insert_Familiares ('" & .TextMatrix(I, 0) & "', '" & Trim(txtEmpleado) & "'," & _
                          " '" & IIf(.TextMatrix(I, 10) <> "", .TextMatrix(I, 10), "00") & "','" & IIf(.TextMatrix(I, 7) <> "", .TextMatrix(I, 7), "00") & "', '" & .TextMatrix(I, 8) & "','" & .TextMatrix(I, 1) & "'," & _
                          " '" & .TextMatrix(I, 2) & "', '" & .TextMatrix(I, 3) & "', '" & Format(.TextMatrix(I, 4), "yyyy/mm/dd") & "', " & _
                          "" & IIf(.TextMatrix(I, 5) <> "", .TextMatrix(I, 5), "0") & " , '" & gen(.TextMatrix(I, 6)) & "', 'S');"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
                    SQL = "insert into rh_verificaessalud (codemp,item,verificado,fecha,familiar) values " & _
                          "('" & txtEmpleado & "'," & .TextMatrix(I, 0) & ",'" & IIf(.TextMatrix(I, 11) = strChecked, "S", "N") & "', " & _
                          "'" & Format(Date, "dd/mm/yyyy") & "','F')"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                End If
            Next
             
        End With
        SQL = "Call Delete_Retenciones('" & Trim(txtEmpleado) & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
        With flxret
            For I = 1 To .Rows - 1
                If val(.TextMatrix(I, 3)) <> 0 Or val(.TextMatrix(I, 4)) <> 0 Then
                    SQL = "Call Insert_Retenciones ('MO','" & Trim(txtEmpleado) & "'," & _
                          " '" & Trim(.TextMatrix(I, 1)) & "'," & CDbl(IIf(val(.TextMatrix(I, 3)) = 0, 0, .TextMatrix(I, 3))) & ",'" & .TextMatrix(I, 2) & "'," & _
                          " " & CDbl(IIf(val(.TextMatrix(I, 4)) = 0, 0, .TextMatrix(I, 4))) & ");"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                End If
            Next
        End With
        ' Asignación de Movilidades
        SQL = "Call Delete_Movilidades('" & Trim(txtEmpleado) & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
        
        SQL = "Call Insert_Movilidades ('00','" & Trim(txtEmpleado) & "'," & _
                          " '" & txtNombre1 & " " & txtNombre2 & " " & txtApePat & " " & txtApeMat & "','" & cboMonMovi.List(cboMonMovi.ListIndex, 1) & "'," & _
                          " " & txtMontoMov.Text & ");"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        
        
        SQL = "Call Delete_Seguros('" & Trim(txtEmpleado) & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
        With msfseg
            EnumerarItems msfseg
            For I = 1 To .Rows - 1
                If Trim(.TextMatrix(I, 1)) <> "" Then
                    If I = 1 Then
                        SQL = "update empleado set codseg='" & Trim(.TextMatrix(I, 1)) & "',numseg='" & Trim(.TextMatrix(I, 3)) & "' where codigo ='" & Trim(txtEmpleado) & "'"
                        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
                    End If
                    SQL = "Call Insert_Seguros ('" & Trim(txtEmpleado) & "'," & Trim(.TextMatrix(I, 0)) & "," & _
                          "'" & Trim(.TextMatrix(I, 1)) & "','" & Trim(.TextMatrix(I, 3)) & "','" & Format(Trim(.TextMatrix(I, 4)), "yyyy/mm/dd") & "'," & _
                          " '" & Format(Trim(.TextMatrix(I, 5)), "yyyy/mm/dd") & "');"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                End If
            Next
        End With
        SQL = "update rh_verificaessalud set verificado = '" & IIf(chkver.Value = False, "N", "S") & "',fecha = '" & Format(Date, "dd/mm/yyyy") & "' " & _
              "where codemp = '" & Trim(txtEmpleado) & "' and familiar = 'T'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
        Actualizar = True
    End If
End Function

Private Function Validar() As Boolean
    Dim I As Integer, J As Integer, numhij As Integer
    Validar = True
    Exit Function
    If txtEmpleado = Empty Then Validar = False: MsgBox "Debe ingresar el código del empleado", vbInformation, gsNomSW: SSTab1.Tab = 0: txtEmpleado.SetFocus: Exit Function
    If txtNombre1 = Empty Then Validar = False: MsgBox "Debe ingresar el Nombre del empleado", vbInformation, gsNomSW: SSTab1.Tab = 0: txtNombre1.SetFocus: Exit Function
    If dtpCese.Value <> Nulo And dtpIngreso <> Nulo Then
        If CDate(dtpCese) < CDate(dtpIngreso) Then
            Validar = False
            MsgBox "La Fecha de Cese, no puede ser menor a la Fecha de Ingreso", vbInformation, gsNomSW
            Exit Function
        End If
    End If
    With flxDependientes
        If cboEstCivil.List(cboEstCivil.ListIndex, 1) <> "C" Then
            For J = 1 To .Rows - 1
                If Trim(.TextMatrix(J, 10)) = CODCONYUG Then
                    MsgBox "Estado Civil es: " & cboEstCivil.List(cboEstCivil.ListIndex, 0) & vbNewLine & _
                           "No se puede ingresar a la Lista de Dependientes", vbInformation, gsNomSW
                    Validar = False
                    SSTab1.Tab = 1
                    .SetFocus
                    .row = J
                    .ColSel = 9
                    Exit Function
                End If
            Next
        End If
        If cboEstCivil.List(cboEstCivil.ListIndex, 1) = "C" Then
            Dim encontro As Boolean
            Dim numConyug As Integer
            encontro = False
            numConyug = 0
            For J = 1 To .Rows - 1
                If Trim(.TextMatrix(J, 10)) = CODCONYUG Then
                    encontro = True
                    numConyug = numConyug + 1
                End If
            Next
            If numConyug > 1 Then
                MsgBox "No se puede ingresar a más de un cónyuge ", vbInformation, gsNomSW
                Validar = False
                SSTab1.Tab = 1
                .SetFocus
                .row = 1
                .ColSel = 9
                Exit Function
            End If
        End If
        If str(Right("0" & Trim(txtNumHijos), 1)) = 0 Then
            For J = 1 To .Rows - 1
                If Trim(.TextMatrix(J, 10)) = CODHIJO Then
                    MsgBox "Número de hijos igual a 0," & vbNewLine & _
                           "No se pueden ingresar a la Lista de Dependientes", vbInformation, gsNomSW
                    SSTab1.Tab = 1
                    .SetFocus
                    .row = 1
                    .ColSel = 9
                    Validar = False
                End If
            Next
        End If
    End With
End Function

Private Sub TipCta(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rscta As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from tipopago where codpago <> '0'"
    Set rscta = oConexion.EjecutaSelectRS(SQL)
    Do While Not rscta.EOF
        cbo.AddItem CE(rscta.Fields("DESCRIP"))
        cbo.List(I, 1) = CE(rscta.Fields("CODPAGO"))
        I = I + 1
        rscta.MoveNext
    Loop
    Set rscta = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub Tipo(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsTipo As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from tipoemp order by codigo"
    Set rsTipo = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsTipo.EOF
        cbo.AddItem CE(rsTipo.Fields("descrip"))
        cbo.List(I, 1) = CE(rsTipo.Fields("codigo"))
        I = I + 1
        rsTipo.MoveNext
    Loop
    Set rsTipo = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub ModEmp(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsModEmp As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from modemp order by codigo"
    Set rsModEmp = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsModEmp.EOF
        cbo.AddItem CE(rsModEmp.Fields("descrip"))
        cbo.List(I, 1) = CE(rsModEmp.Fields("codigo"))
        I = I + 1
        rsModEmp.MoveNext
    Loop
    Set rsModEmp = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub Nacionalidad(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsNacion As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from nacionalidad order by codigo"
    Set rsNacion = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsNacion.EOF
        cbo.AddItem CE(rsNacion.Fields("descrip"))
        cbo.List(I, 1) = CE(rsNacion.Fields("codigo"))
        I = I + 1
        rsNacion.MoveNext
    Loop
    Set rsNacion = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub Tallas(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsTallas As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from tallas order by codigo"
    Set rsTallas = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsTallas.EOF
        cbo.AddItem CE(rsTallas.Fields("descrip"))
        cbo.List(I, 1) = CE(rsTallas.Fields("codigo"))
        I = I + 1
        rsTallas.MoveNext
    Loop
    Set rsTallas = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub TipoDoc(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsTipoDoc As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from doc_identificacion where tipo_doc_ide<>'00' order by tipo_doc_ide"
    Set rsTipoDoc = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsTipoDoc.EOF
        cbo.AddItem CE(rsTipoDoc.Fields("DESCRIP"))
        cbo.List(I, 1) = CE(rsTipoDoc.Fields("TIPO_DOC_IDE"))
        I = I + 1
        rsTipoDoc.MoveNext
    Loop
    Set rsTipoDoc = Nothing
    cbo.ListIndex = 0
End Sub


Private Sub TipoParen(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsTipoParen As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from parentesco order by paren_ide"
    Set rsTipoParen = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsTipoParen.EOF
        cbo.AddItem CE(rsTipoParen.Fields("DESCRIP"))
        cbo.List(I, 1) = CE(rsTipoParen.Fields("PAREN_IDE"))
        I = I + 1
        rsTipoParen.MoveNext
    Loop
    Set rsTipoParen = Nothing
    cbo.ListIndex = 0
End Sub


Private Sub SituacionEmp(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsSitua As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from situacionemp order by codigo"
    Set rsSitua = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsSitua.EOF
        cbo.AddItem CE(rsSitua.Fields("descrip"))
        cbo.List(I, 1) = CE(rsSitua.Fields("codigo"))
        I = I + 1
        rsSitua.MoveNext
    Loop
    Set rsSitua = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub TBrevetes(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsBreve As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from tiposbrevete order by codigo"
    Set rsBreve = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsBreve.EOF
        cbo.AddItem CE(rsBreve.Fields("descrip"))
        cbo.List(I, 1) = CE(rsBreve.Fields("codigo"))
        I = I + 1
        rsBreve.MoveNext
    Loop
    Set rsBreve = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub Gsangre(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsGsangre As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from GSangre order by codigo"
    Set rsGsangre = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsGsangre.EOF
        cbo.AddItem CE(rsGsangre.Fields("descrip"))
        cbo.List(I, 1) = CE(rsGsangre.Fields("codigo"))
        I = I + 1
        rsGsangre.MoveNext
    Loop
    Set rsGsangre = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub Personal(cbo As MSForms.ComboBox) ' DIRECCION
    Dim SQL As String
    Dim rsPerso As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from tippersonal order by codigo"
    Set rsPerso = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsPerso.EOF
        cbo.AddItem CE(rsPerso.Fields("descrip"))
        cbo.List(I, 1) = CE(rsPerso.Fields("codigo"))
        I = I + 1
        rsPerso.MoveNext
    Loop
    Set rsPerso = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub btnGrabarC_Click()
    contrato Trim(txtEmpleado), Trim(txtContrato)
    ActualizaFechaIngreso Trim(txtEmpledo)
End Sub

Private Sub btnGrabarFormEduc_Click()
Dim SQL As String

    
'Formación Educativa
 If OptPeruSi.Value = True Then
   OptPeruflg = "S"
 Else
   OptPeruflg = "N"
 End If
              
 If OptRegEduPub.Value = 1 Then
   OptRegEduflg = "N"
 Else
   OptRegEduflg = "P"
 End If
 
'FALTA VALIDAR QUE NO SE GRABE DOS VECES
If ValidaFormacionEducativa(Trim(txtEmpleado)) = False Then
 SQL = " Insert into pl_empformedu(codemp, codformsup,flgEstPeru,flgRegInst,flgTipoInst,coduni,codcarr,anhoegr,visible)" & _
                    " values ('" & Trim(txtEmpleado) & "','" & cboFormSup.List(cboFormSup.ListIndex, 1) & "', '" & OptPeruflg & "', '" & OptRegEduflg & "', '" & cboTipInst.List(cboTipInst.ListIndex, 1) & "'," & _
                "'" & cboNomInst.List(cboNomInst.ListIndex, 1) & "','" & cboCarrera.List(cboCarrera.ListIndex, 1) & "', '" & txtAnhoEgr & "', '" & cboEstPrin.List(cboEstPrin.ListIndex, 1) & "')"
 oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
 
  MsgBox "Formación Educativa Ingresada", vbInformation, "NOV"
Else
  MsgBox "Ya se encuentra ingresada la Formación Educativa, verifique los datos", vbInformation, "NOV"
End If


 
End Sub


Private Sub btnHabilitar_Click()
    Dim SQL As String
    Dim RES As Integer
    Dim UsuAceptado As Boolean
    UsuAceptado = False
    If Trim(txtContrato) = "CN01" Then
        UsuAceptado = False
        If InStr(1, lblautoriza, "ACTIVAR") = 0 Then
            If UsuarioAceptado(2, strUsuarioId, "activar al empleado y habilitar su primer contrato", txtEmpleado, 0, "", "", Format(dtpIngreso.Value, "yyyy/mm/dd")) = True Then
                UsuAceptado = True
            End If
        End If
    Else
        UsuAceptado = True
    End If
    
    If UsuAceptado = True Then
    RES = MsgBox("¿Esta Seguro que desea habilitar el Contrato del empleado " & vbNewLine & " con código Nro. " & Trim(txtEmpleado) & " ?", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        SQL = " Update contrato set estado = '" & APROBADO & "' where codigo = '" & Trim(txtContrato) & "'" & _
              " and codemp = '" & Trim(txtEmpleado) & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        Consulta
        MoverPosRs Trim(txtEmpleado)
        CargarDatos rsgral
        ModoFormulario modConsulta
    End If
End If
End Sub

Private Sub btnModificar_Click()
     ModoFormulario modEditar
End Sub

Private Sub btnNuevo_Click()
    Dim RES As Integer
    ModoFormulario modNuevo
    
    cboFormSup.ListIndex = 0
    cboTipInst.ListIndex = 0
    cboNomInst.ListIndex = 0
    cboCarrera.ListIndex = 0
    cboEstPrin.ListIndex = 0
    txtAnhoEgr.Text = ""
    
    RES = MsgBox("Desea Generar Código?", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        generocod = True
        txtEmpleado = GenCodEmp
        txtEmpleado.SetFocus
        txtEmpleado.SelStart = 11
        txtRuc.Enabled = False
        txtRuc.BackColor = ColorDeshabilitado
    Else
        txtEmpleado.SetFocus
        generocod = False
    End If
End Sub

Private Sub btnPrevio_Click()
    txtBusqueda = Empty
    If rsgral.AbsolutePosition > 0 Then
        If Not rsgral.AbsolutePosition = 1 Then
            rsgral.MovePrevious
            CargarDatos rsgral
        End If
        ConfigBtnsBusq rsgral.AbsolutePosition, rsgral.RecordCount
    End If
End Sub

Private Sub btnPrimero_Click()
    txtBusqueda = Empty
    rsgral.MoveFirst
    CargarDatos rsgral
    ConfigBtnsBusq rsgral.AbsolutePosition, rsgral.RecordCount
End Sub

Private Sub btnRenovar_Click()
    Dim UsuAceptado As Boolean
    Dim SQL As String
    UsuAceptado = False
    txtContrato = GenNumContrato(txtEmpleado)
    txtContrato.Locked = True
    txtContrato.BackColor = ColorDeshabilitado
    txtContrato.tag = PENDIENTE
    dtpFin.CheckBox = True
    dtpFin.Value = Date
    dtpFin.Value = Null
    btnGrabarC.Enabled = True
    btnGrabar.Enabled = False
    btnCancelCont.Enabled = False
    ConfiguraGrillaSueldos
    With flxsueldos
        .TextMatrix(.Rows - 1, 0) = val(.Rows - 1)
        .TextMatrix(.Rows - 1, 1) = FormatNumber(lblsbasico, 2)
    End With
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub btnSgt_Click()
    txtBusqueda = Empty
    If rsgral.AbsolutePosition > 0 Then
        If rsgral.AbsolutePosition <> rsgral.RecordCount Then
            rsgral.MoveNext
            CargarDatos rsgral
        End If
        ConfigBtnsBusq rsgral.AbsolutePosition, rsgral.RecordCount
    End If
    If ComisionAFP = "0" Then
         optComiAFP(0).Value = True
    Else
         optComiAFP(1).Value = True
    End If
End Sub

Private Sub btnUltimo_Click()
    txtBusqueda = Empty
    rsgral.MoveLast
    CargarDatos rsgral
    ConfigBtnsBusq rsgral.AbsolutePosition, rsgral.RecordCount
    If ComisionAFP = "0" Then
         optComiAFP(0).Value = True
    Else
         optComiAFP(1).Value = True
    End If
End Sub

Private Sub cboBreveteCat_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then txtBrevete.SetFocus
End Sub

Private Sub cboBusqueda_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then txtBusqueda.SetFocus
End Sub

Private Sub cboBusqueda_LostFocus()
    txtBusqueda.SetFocus
End Sub

Private Sub cboCategoria_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then cboTipo.SetFocus
End Sub

Private Sub cboContratos_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then dtpInicio.SetFocus
End Sub

Private Sub cboEstCivil_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then txtNumHijos.SetFocus
End Sub

Private Sub cboEstTrabajo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cboHorLab.SetFocus
End Sub

Private Sub cboGenero_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cboNacion.SetFocus
End Sub

Private Sub cboGenFlx_DropDown()
    With cboGenFlx
        .Clear
        .AddItem "Femenino"
        .AddItem "Masculino"
    End With
End Sub

Private Sub cboGenFlx_GotFocus()
    If cboGenFlx.ListCount > 0 Then cboGenFlx.ListIndex = 0
    Call keybd_event(vbKeyHome, 0, 0, 0)
End Sub

Private Sub cboGenFlx_LostFocus()
    With flxDependientes
        .TextMatrix(.row, 6) = cboGenFlx.List(cboGenFlx.ListIndex)
        cboGenFlx.Visible = False
    End With
End Sub

Public Function ToHex(cadena As String) As String
    Dim I As Long
    Dim Tamanio As Long
    Tamanio = Len(cadena)
    If Tamanio < 61000 Then
        For I = 1 To Tamanio
            ToHex = ToHex & Right("00" & Trim(Hex(Asc(Mid(cadena, I, 1)))), 2)
            pbProgreso.Value = I * 100 / Tamanio
        Next
    Else
        MsgBox "Tamaño de imagen no permitida, " & vbNewLine & vbNewLine & _
               "Recomendado: 400x500 pxls", vbInformation, gsNomSW
        ToHex = "0x"
        pbProgreso.Visible = False
        Exit Function
    End If
    pbProgreso.Visible = False
    ToHex = "0x" & ToHex
End Function

Private Sub cboGSanguineo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cboEstCivil.SetFocus
End Sub

Private Sub cboHorLab_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then txtcencos.SetFocus
End Sub

Private Sub cboMonBono_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then txtMontoBono.SetFocus
End Sub

Private Sub cboMonMovi_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then txtMontoMov.SetFocus
End Sub

Private Sub cboMonMovi_KeyPress(KeyAscii As MSForms.ReturnInteger)
    cboMonMovi.Locked = False
End Sub


Private Sub CboMonCTS_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then txtCTSCta.SetFocus
End Sub


Private Sub cboMonSueldo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then chkBono.SetFocus
End Sub

Private Sub cboNacion_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then txtCarnetExt.SetFocus
End Sub

Private Sub cboPersonal_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then cboCategoria.SetFocus
End Sub

Private Sub cboSituacion_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cboPersonal.SetFocus
End Sub

Private Sub cboTipCtaMe_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then txtNumCtaMe.SetFocus
End Sub

Private Sub cboTipCtaMn_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then txtNumCtaMn.SetFocus
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then chkSvl.SetFocus
End Sub

Private Sub cboTMame_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cboGSanguineo.SetFocus
End Sub





Private Sub chkAsigFam_Click()
    If chkAsigFam.Value = True And lblModo = "Modificar" Then
        txtAsigFam.Locked = False
        txtAsigFam.BackColor = ColorHabilitado
    Else
        txtAsigFam.Locked = True
        txtAsigFam.BackColor = ColorDeshabilitado
    End If
End Sub

Private Sub chkJubil_Click()
    If chkJubil.Value = True And lblModo = "Modificar" Then
        chkJubil.Locked = False
    Else
        chkJubil.Locked = True
    End If
End Sub

Private Sub chkBono_Click()
    If chkBono.Value = True And lblModo = "Modificar" Then
        cboMonBono.Locked = False
        cboMonBono.BackColor = ColorHabilitado
        txtMontoBono.Locked = False
        txtMontoBono.BackColor = ColorHabilitado
    Else
        cboMonBono.Locked = True
        cboMonBono.BackColor = ColorDeshabilitado
        cboMonBono.ListIndex = 0
        txtMontoBono.Locked = True
        txtMontoBono.BackColor = ColorDeshabilitado
        txtMontoBono = "0.00"
    End If
End Sub

Private Sub chkmovi_Click()
    cboMonMovi.Locked = False
    cboMonMovi.BackColor = ColorHabilitado
    txtMontoMov.Locked = False
    txtMontoMov.BackColor = ColorHabilitado
End Sub


Private Sub chkBono_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        If chkBono.Value = True Then
            cboMonBono.SetFocus
        End If
    End If
End Sub

'Private Sub chkmovi_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        If chkmovi.Value = True Then
'            cboMonMovi.SetFocus
'        End If
'    End If
'End Sub

Private Sub chkSctr_Click()
    If chkSctr.Value = True And lblModo = "Modificar" Then
        chkSctr.Locked = False
    Else
        chkSctr.Locked = True
    End If
End Sub

Private Sub chkSvl_Click()
     If chkSvl.Value = True And lblModo = "Modificar" Then
        chkSvl.Locked = False
    Else
        chkSvl.Locked = True
    End If
End Sub

Private Sub CmdCopiar_Click()
    Clipboard.SetText TxtDirectorio
    MsgBox "Copiado", vbInformation
End Sub

Private Sub cmdadjuntar_Click()
    frmArchivosAdjuntos.IdentificadorAr = txtEmpleado.Text
    frmArchivosAdjuntos.AnioSel = strAnoSistema
    frmArchivosAdjuntos.MesSel = strMesSistema
    frmArchivosAdjuntos.Show
End Sub

Private Sub CmdDirectorio_Click()
Dim Emails As String
    Emails = CargarEmails
    
    If Emails <> "" Then
'        MensajeNuevoOutlook Emails, 1

         If EnviarEmail("Eduardo.madrid@nov.com;", "", "", "gustavo.mayaute@nov.com;Rocio.RodasRojas@nov.com;", Emails, "CORREOS NOV", "", "") Then
           MsgBox "Base de Datos de Correos Enviado.", vbOKOnly + vbInformation, gsNomSW
         End If
    End If
    
End Sub

Function CargarEmails() As String
    Dim SQL As String, cad As String, Cad1 As String
    Dim RQ As MYSQL_RS
    SQL = "select mail,concat_ws(' ',nombre1,nombre2,apepat,apemat) as nombres FROM empleado where mail like '%nov.com%' and situacion = 1 order by mail"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    Do While Not RQ.EOF()
        If Trim(RQ.Fields("mail")) <> "" Then
            cad = cad & RQ.Fields("mail") & ";"
        Else
            Cad1 = Cad1 & Trim(RQ.Fields("nombres")) & Chr(13)
        End If
        RQ.MoveNext
    Loop
    If Cad1 <> "" Then
        CargarEmails = ""
        MsgBox "El(Los) siguiente(s) empleado(s) no cuentan con Correo Electrónico: " & Chr(13) & Cad1, vbInformation, "NOVPeru"
    Else
        cad = Mid(cad, 1, Len(cad) - 1)
        CargarEmails = cad
    End If
    Set RQ = Nothing
End Function

Private Sub cmdElimSoli_Click()
    Dim SQL As String
    If MsgBox("¿Está seguro de eliminar la(s) solicitud(es) enviada(s) para este empleado?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
        SQL = "delete from rh_tempacceso where usuario = '" & strUsuarioId & "' AND codemp = '" & Trim(txtEmpleado) & "'"
        oConexionMYSQL.Execute SQL
        CargaAutorizaciones Trim(txtEmpleado)
        MsgBox "Se eliminaron sus solicitudes", vbInformation, "NOVPeru"
    End If
End Sub

Private Sub cmdenviar_Click()
Dim Email As String, ps As String, mensaje As String
Dim Nom As String, Ape As String, SQL As String
Dim RQ As MYSQL_RS
Dim cantc As Integer, canta As Integer, CantM As Integer
Dim msc As String, msa As String, msm As String
    If MsgBox("¿Seguro desea enviar un Email para las autorizaciones?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
        Nom = LCase(Mid(strNombreUsuario, 1, InStr(1, strNombreUsuario, " ") - 1))
        Ape = LCase(Mid(strNombreUsuario, InStr(1, strNombreUsuario, " ") + 1, IIf(InStr(InStr(1, strNombreUsuario, " ") + 1, strNombreUsuario, " ") = 0, Len(strNombreUsuario), InStr(InStr(1, strNombreUsuario, " ") + 1, strNombreUsuario, " ") - InStr(1, strNombreUsuario, " ") - 1)))
        Email = Nom & "." & Ape & "@nov.com"
        SQL = "select numdocide, fec_nac from empleado where apepat = '" & Ape & "' and nombre1 = '" & Nom & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            ps = Mid(Trim(RQ.Fields("numdocide")), 1, 1) & Year(CDate(RQ.Fields("fec_nac"))) & Mid(Trim(RQ.Fields("numdocide")), Len(Trim(RQ.Fields("numdocide"))), 1)
        End If
        mensaje = Chr(13) & "Se solicita la autorización de las siguientes acciones: " & Chr(13) & Chr(13)
        SQL = "select r.usuario,c.descrip,(select CONCAT_WS(' ',apepat,apemat,nombre1) from empleado e where e.codigo=r.codemp) as nomemp, " & _
              "(select tipo from empleado e where e.codigo=r.codemp) AS prc," & _
              "DATE_FORMAT(r.fechacese,'%d/%m/%Y') as fechacese,(select descrip from pl_tipocese p where p.codigo=r.tipocese) as tipocese,r.sueldo, " & _
              "(select sbasico from contrato t where (t.codemp=r.codemp) and codigo = (select max(codigo) from contrato o where o.codemp=t.codemp group by o.codemp)) as basico " & _
              "from rh_tempacceso r inner join configuracion_acceso c on (r.codigo=c.codigo) LEFT JOIN autorizaciones a on (a.codigo=c.codigo) " & _
              "where AUTORIZADO = 'N' and a.usuario = 'MOA' and r.usuario = '" & strUsuarioId & "' order by nomemp"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            Do While Not RQ.EOF()
                Select Case Mid(RQ.Fields("descrip"), 1, InStr(1, RQ.Fields("descrip"), " ") - 1)
                    Case "ACTIVAR"
                        If canta = 0 Then msa = "ACTIVAR TRABAJADORES: " & Chr(13)
                        If canta > 0 Then
                            msa = msa & "- " & Trim(RQ.Fields("nomemp")) & IIf(RQ.Fields("PRC") = 3, "   Practicante", "") & Chr(13)
                        End If
                        canta = canta + 1
                        If canta > 1 Then RQ.MoveNext
                    Case "CESAR"
                        If cantc = 0 Then msc = "CESAR TRABAJADORES: " & Chr(13)
                        If cantc > 0 Then
                            msc = msc & "- " & Trim(RQ.Fields("nomemp")) & IIf(RQ.Fields("PRC") = 3, "   Practicante", "") & Space(6) & "Cese: " & Trim(RQ.Fields("fechacese")) & Space(4) & Trim(RQ.Fields("tipocese")) & Chr(13)
                        End If
                        cantc = cantc + 1
                        If cantc > 1 Then RQ.MoveNext
                    Case "MODIFICAR"
                        If CantM = 0 Then msm = "MODIFICAR SUELDOS DE: " & Chr(13)
                        If CantM > 0 Then
                            msm = msm & "- " & Trim(RQ.Fields("nomemp")) & Chr(9) & "Sueldo Ant.: " & RQ.Fields("basico") & Space(4) & "A: " & RQ.Fields("sueldo") & Chr(13)
                        End If
                        CantM = CantM + 1
                        If CantM > 1 Then RQ.MoveNext
                End Select
            Loop
                            
            'If EnviarEmail("Maria.OcanaAngeles@nov.com", "", "", "gustavo.mayaute@nov.com", mensaje & msa & msc & msm, "NOVADMIN:SOLICITUD DE AUTORIZACIONES RRHH", "Augusto.MendozaTalla@nov.com;Carla.CerdanTijera@nov.com", "") Then
            
            If EnviarEmail("gustavo.mayaute@nov.com", "", "", "gustavo.mayaute@nov.com", mensaje & msa & msc & msm, "NOVADMIN:SOLICITUD DE AUTORIZACIONES RRHH", "Carla.CerdanTijera@nov.com", "") Then
                MsgBox "Email enviado", vbInformation, "NOV"
            End If
            
        Else
            MsgBox "No tiene solicitudes de autorización", vbInformation, "NOVPeru"
        End If
    End If
    Set RQ = Nothing
End Sub

Private Sub cmdformedu_Click()
    With frmformedu
        If .Visible = True Then
            .Visible = False
        Else
            .Visible = True
        End If
    End With
    
    CargarFormacEduc txtEmpleado
End Sub

Private Sub cmdotros_Click()
    With frcese
        If .Visible = True Then
            .Visible = False
        Else
            .Top = 825
            .Left = 5820
            .Visible = True
        End If
    End With
End Sub

Private Sub cmdver_Click()
    With frmsueldos
        If .Visible = True Then
            .Visible = False
        Else
            '.Top = 240
            '.Left = 3150
            .Visible = True
        End If
    End With
End Sub

Sub ConfiguraGrillaSueldos()
    With flxsueldos
        .Clear
        .Cols = 3
        .Rows = 2
        .FixedCols = 1
        .ColWidth(0) = 300
        .TextMatrix(0, 0) = "Item"
        .ColType(0) = Numero
        .ColMaxLength(0) = 4
        .ColWidth(1) = 1150
        .TextMatrix(0, 1) = Space(4) & "Sueldo"
        .ColType(1) = Numero
        .ColMaxLength(1) = 15
        .CaracteresValidos(1) = "1234567890."
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = Space(4) & "Fecha"
        .ColType(2) = fecha
        .ColMaxLength(2) = 10
        .CaracteresValidos(2) = "1234567890/"
    End With
End Sub

Private Sub dtpCese_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cboPersonal.SetFocus
    End If
End Sub

Private Sub dtpFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cboMonSueldo.SetFocus
    End If
End Sub

Private Sub dtpIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dtpCese.SetFocus
    End If
End Sub

Private Sub dtpInicio_Change()
    With flxsueldos
        If IsDate(dtpInicio.Value) Then
            .TextMatrix(.Rows - 1, 2) = dtpInicio.Value
        Else
            .TextMatrix(.Rows - 1, 2) = ""
        End If
    End With
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dtpFin.SetFocus
    End If
End Sub

Private Sub dtpNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtEdad = "  "
        txtEdad = Right("00" & Trim(val(Year(Date) - Right(Trim(dtpNacimiento), 4))), 2)
        If val(txtEdad) <= 60 And val(txtEdad) >= 18 Then
            txtEdad = txtEdad
        Else
            txtEdad = "00"
        End If
        txtEdad.SetFocus
    End If
End Sub

Private Sub flxDependientes_Click()
    If flxDependientes.Col = 6 And (lblModo = "Nuevo" Or lblModo = "Modificar") Then
        With cboGenFlx
            .Top = flxDependientes.CellTop + flxDependientes.Top
            .Left = flxDependientes.CellLeft + flxDependientes.Left
            .Width = flxDependientes.CellWidth
            .Visible = True
        End With
    Else
        cboGenFlx.Visible = False
    End If
    If flxDependientes.Col = 11 And (lblModo = "Nuevo" Or lblModo = "Modificar") Then
        With flxDependientes
            .TextMatrix(.row, 11) = IIf(.TextMatrix(.row, 11) = strChecked, strUnChecked, strChecked)
        End With
    End If
End Sub

Private Sub GridEditCombo()
    flxDependientes.Col = 6
    cboGenFlx.Left = flxDependientes.CellLeft + flxDependientes.Left
    cboGenFlx.Top = flxDependientes.CellTop + flxDependientes.Top
    cboGenFlx.Width = flxDependientes.CellWidth
    cboGenFlx.Visible = True
End Sub
  
Private Sub flxDependientes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And flxDependientes.Col = 9 And (lblModo = "Nuevo" Or lblModo = "Modificar") Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1000
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Tipos de Dependientes"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_DEPEND
            .Show
        End With
    End If
    If KeyCode = vbKeyF1 And flxDependientes.Col = 7 And (lblModo = "Nuevo" Or lblModo = "Modificar") Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1000
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Tipos de Documentos"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_TIP_DOC
            .Show
        End With
    End If
End Sub

Private Sub flxDependientes_KeyPress(KeyAscii As Integer)
    Dim I As Integer
    If KeyAscii = 13 Then
        On Error GoTo pregunta
        With flxDependientes
            If .Col = 9 And .row = .Rows - 1 Then 'Parentesco, penultima columna
                If Trim(.TextMatrix(.row, 4)) <> Empty Then
                    If Not IsDate(Trim(.TextMatrix(.row, 4))) Then
                        .Col = 3
                        .SetFocus
                        Exit Sub
                    End If
                End If
                If Trim(.TextMatrix(.row, 9)) <> Empty And Trim(.TextMatrix(.row, 10)) <> Empty Then
                    EnumerarItems flxDependientes
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .Col = 0
                End If
                If Trim(.TextMatrix(.row, 9)) = Empty Then
                    .TextMatrix(.row, 10) = Empty
                End If
            End If
            If .Col <> 4 Then 'Fecha de Nac
                TipodeCampo = cadena
            End If
            If .Col = 3 Then '
                'SendKeys "{f2}"
            End If
            If .Col = 4 And Trim(.TextMatrix(.row, 4)) <> Empty Then
                .TextMatrix(.row, 4) = Space(0) & Trim(.TextMatrix(.row, 4))
                TipodeCampo = fecha
                'SendKeys "{f2}"
            Else
                TipodeCampo = cadena
            End If
            If .Col = 5 Then ' Edad
                If Trim(.TextMatrix(.row, 4)) <> Empty Then
                     TipodeCampo = cadena
                     .TextMatrix(.row, 5) = IIf(Month(Trim(.TextMatrix(.row, 4))) < Month(Date), Format(Date, "yyyy") - Format(Trim(.TextMatrix(.row, 4)), "yyyy"), IIf(Month(Trim(.TextMatrix(.row, 4))) = Month(Date), IIf(Day(Trim(.TextMatrix(.row, 4))) <= Day(Date), (Format(Date, "yyyy") - Format(Trim(.TextMatrix(.row, 4)), "yyyy")), (Format(Date, "yyyy") - Format(Trim(.TextMatrix(.row, 4)), "yyyy")) - 1), (Format(Date, "yyyy") - Format(Trim(.TextMatrix(.row, 4)), "yyyy")) - 1))
                    If Not (Trim(.TextMatrix(.row, 5)) >= 0 And Trim(.TextMatrix(.row, 5)) <= 100) Then
                        .TextMatrix(.row, 5) = "0"
                    End If
                End If
               ' SendKeys "{F2}"
            End If
            If .Col = 8 Then 'Nro
                    .TextMatrix(.row, 8) = Right("00000000" & Trim(.TextMatrix(.row, 8)), 8)
            End If
             If .Col = 7 Then 'TipoDocIde
                .TextMatrix(.row, 7) = Right("00" & Trim(.TextMatrix(.row, 7)), 2)
                'SendKeys "{f2}"
            End If
            If .Col = 6 Then 'sexo
                If .TextMatrix(.row, 6) <> Empty Then
                    .TextMatrix(.row, 6) = gen(Trim(.TextMatrix(.row, 6)))
                End If
            End If
        End With
    End If
Exit Sub
pregunta:
    MsgBox "Uno de los datos es incorrecto," & vbNewLine & "por favor revise o consulte con su administrador", vbOKOnly + vbInformation, "NOVPeru"
    Exit Sub
End Sub

Private Sub flxDependientes_RowColChange()
    With flxDependientes
        If .Col = 4 Then
            TipodeCampo = fecha
        Else
            TipodeCampo = cadena
        End If
    End With
End Sub

Private Sub flxret_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        CalculoTotalRet
    End If
End Sub

Private Sub flxret_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With flxret
            If .Col = 3 Then
                .TextMatrix(.row, 4) = "0.00"
                'SendKeys "{f2}"
            End If
            If .Col = 4 Then
                .TextMatrix(.row, 4) = FormatNumber(IIf(val(.TextMatrix(.row, 4)) = 0, 0, Trim(.TextMatrix(.row, 4))), 2)
                If val(.TextMatrix(.row, 4)) > 0 Or val(.TextMatrix(.row, 3)) > 0 Then
                    EnumerarItems flxret
                    CalculoTotalRet
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .Col = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub flxsueldos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With flxsueldos
            If .Col = 1 Then
                .TextMatrix(.row, 1) = FormatNumber(IIf(val(.TextMatrix(.row, 1)) = 0, 0, Trim(.TextMatrix(.row, 1))), 2)
                If cboContratos.List(cboContratos.ListIndex, 1) = "03" Then
                    lblsbasico = .TextMatrix(.row, 1)
                Else
                    sbasico = .TextMatrix(.row, 1)
                End If
            End If
        End With
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmsueldos.Visible = False
        frcese.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim SQL As String
    'SQL = "Call sp_ActualizaEdad();"
    ConfiguracboTipo
    MsgBox "PASA EL 1", vbOKOnly + vbInformation, "NOVPeru"
    'oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
    Me.Top = 0
    Me.Left = 0
    ConfiguraGrilla
    MsgBox "PASA LA CONFIG GRILLA", vbOKOnly + vbInformation, "NOVPeru"
    'ModoFormulario modAccion
    TipodeCampo = cadena
    strFoto = ""
    'Call WheelHook(frmRegEmpleado)
    Set oConsulta = New FrmConsultas
    SSTab1.Tab = 0
    FlgCont = False
    If VerContratoySueldos Then
        frmcont.Height = 2820
        FlgCont = True
    Else
        frmcont.Height = 0
        FlgCont = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RES As Integer
    RES = MsgBox("¿Desea salir del formulario?", vbYesNo + vbQuestion, "Aviso")
    If RES = vbNo Then
        Cancel = 1
    Else
        WheelUnHook
        Set oConsulta = Nothing
        Set rsgral = Nothing
    End If
End Sub

Private Sub FrCont_DblClick()
Dim Usu As String
Dim pass As String
Dim SQL As String
Dim RQ As MYSQL_RS
    If FlgCont = False Then
        Usu = InputBox("Ingrese su Usuario de Acceso al Módulo", "NOVPeru")
        pass = InputBox("Ingrese su Clave de Acceso al Módulo", "NOVPeru")
        SQL = "Select clave from 3cnuser where usuario_id='" & Usu & "' and area='0005'"
        SQL = "select * from 3cnuser where ((area = '0005' and perfil_id='0006') OR perfil_id='0004')  and usuario_id= '" & Usu & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            If pass = DecodificarClave(Trim(RQ.Fields("clave"))) Then
                frmcont.Height = 2820
            Else
                MsgBox "Clave Incorrecta", vbInformation, "NOVPeru"
            End If
        Else
            MsgBox "El usuario no existe o no pertenece al Area de R.R.H.H." & vbNewLine & _
                   "vuelva intentarlo o consulte con el administrador del sistema", vbOKOnly + vbInformation, "NOVPeru"
            frmcont.Height = 0
        End If
    End If
    Set RQ = Nothing
End Sub


Private Sub lblnumsctr_DblClick()
    If chkSctr.Value = True Then
        frmPolizas.TipoPoliza = "R"
        frmPolizas.Show
    End If
End Sub

Private Sub lblnumsvl_DblClick()
    If chkSvl.Value = True Then
        frmPolizas.TipoPoliza = "L"
        frmPolizas.Show
    End If
End Sub

Private Sub msfseg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And msfseg.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1000
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Seguros Médicos"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_SEGU
            .Show
        End With
    End If
End Sub

Private Sub msfseg_KeyPress(KeyAscii As Integer)
    With msfseg
        If KeyAscii = 13 And .BackColor = ColorHabilitado Then
            .TextMatrix(.row, 2) = DescripcionesdeCodigos("SEGURO", .TextMatrix(.row, 1))
            If .Col = 5 Then
                If Trim(.TextMatrix(.row, 1)) <> "" And Trim(.TextMatrix(.row, 3)) <> "" Then
                    If .row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub msfseg_RowColChange()
    With msfseg
        Select Case .Col
            Case 1, 2, 3
                TipodeCampo = cadena
            Case 4, 5
                TipodeCampo = fecha
        End Select
    End With
End Sub

Private Sub OptRegEduPriv_Click()
  OptRegEduPub.Value = 0
End Sub

Private Sub OptRegEduPub_Click()
  OptRegEduPriv.Value = 0
End Sub

Private Sub txtAfp_Change()
    If txtAfp = Empty Then
        lblAfp = Empty
    End If
End Sub

Private Sub txtAfp_GotFocus()
    mark1 txtAfp
End Sub

Private Sub txtAfp_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 And txtAfp.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 3
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 4000
            .pTitulo = "AFP's"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_AFP
            .Show
        End With
    End If
    If KeyCode = 13 Then
        txtAfp = Right("00" & Trim(txtAfp), 2)
        lblAfp = DescripcionesdeCodigos("AFP", txtAfp, "Descrip")
        txtNumAfp.SetFocus
    End If
    If KeyCode <> vbKeySpace Then
        If KeyCode <> vbKeyDelete Then
            If KeyCode <> 8 Then
                If Not (KeyCode >= 96 And KeyCode <= 109) Then
                    If Not IsNumeric("0" & Chr(KeyCode)) Then
                        Beep
                        KeyCode = 0
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtApeMat_Change()
    txtApeMat = UCase$(txtApeMat)
End Sub

Private Sub txtApeMat_GotFocus()
    mark1 txtApeMat
End Sub

Private Sub txtApeMat_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        dtpNacimiento.SetFocus
        Exit Sub
    End If
    If KeyCode <> vbKeySpace Then
        If KeyCode <> vbKeyDelete Then
            If KeyCode <> 8 Then
                If IsNumeric("0" & Chr(KeyCode)) Or (KeyCode >= 96 And KeyCode <= 109) Then
                    Beep
                    KeyCode = 0
                End If
            End If
        End If
    End If
   End If
End Sub

Private Sub txtApePat_Change()
    txtApePat = UCase$(txtApePat)
End Sub

Private Sub txtApePat_GotFocus()
    mark1 txtApePat
End Sub

Private Sub txtApePat_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtApeMat.SetFocus
        Exit Sub
    End If
    If KeyCode <> vbKeySpace Then
        If KeyCode <> vbKeyDelete Then
            If KeyCode <> 8 Then
                If IsNumeric("0" & Chr(KeyCode)) Or (KeyCode >= 96 And KeyCode <= 109) Then
                    Beep
                    KeyCode = 0
                End If
            End If
        End If
    End If
    End If
End Sub

Private Sub txtApoderado_GotFocus()
    mark1 txtApoderado
End Sub

Private Sub txtApoderado_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtDirApo.SetFocus
    End If
 End If
End Sub

Private Sub txtBanco_Change()
    If txtBanco = Empty Then
        lblBanco = Empty
    Else
        lblBanco = DescripcionesdeCodigos("BANCO", Right("00" & Trim(txtBanco), 2))
    End If
End Sub

Private Sub txtBanco_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 And txtBanco.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 4000
            .pTitulo = "Cargos"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_BANCO
            .Show
        End With
    End If
    If KeyCode = 13 Then
        txtBanco = Right("00" & Trim(txtBanco), 2)
        lblBanco = DescripcionesdeCodigos("BANCO", Trim(txtBanco))
        cboTipCtaMn.SetFocus
    End If
    If KeyCode <> vbKeySpace Then
        If KeyCode <> vbKeyDelete Then
            If KeyCode <> 8 Then
                If Not (KeyCode >= 96 And KeyCode <= 109) Then
                    If Not IsNumeric("0" & Chr(KeyCode)) Then
                        Beep
                        KeyCode = 0
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtBanco_LostFocus()
    txtBanco = Right("00" & Trim(txtBanco), 2)
End Sub

Private Sub txtBancoCTS_Change()
    If txtBancoCTS = Empty Then
        lblBancoCTS = Empty
    Else
        lblBancoCTS = DescripcionesdeCodigos("BANCO", Right("00" & Trim(txtBancoCTS), 2))
    End If
End Sub

Private Sub txtBancoCTS_GotFocus()
    mark1 txtBancoCTS
End Sub

Private Sub txtBancoCTS_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 And txtBanco.BackColor = ColorHabilitado Then
         With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 4000
            .pTitulo = "Cargos"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_BANCOCTS
            .Show
        End With
    End If
    If KeyCode = 13 Then
        CboMonCTS.SetFocus
    End If
     If KeyCode <> vbKeySpace Then
        If KeyCode <> vbKeyDelete Then
            If KeyCode <> 8 Then
                If Not (KeyCode >= 96 And KeyCode <= 109) Then
                    If Not IsNumeric("0" & Chr(KeyCode)) Then
                        Beep
                        KeyCode = 0
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtBancoCTS_LostFocus()
    txtBancoCTS = Right("00" & Trim(txtBancoCTS), 2)
End Sub

Private Sub txtBrevete_GotFocus()
    mark1 txtBrevete
End Sub

Private Sub txtBrevete_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtEstatura.SetFocus
    End If
 End If
End Sub

Private Sub txtBusqueda_GotFocus()
    mark1 txtBusqueda
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 And (lblModo = "Acción" Or lblModo = "Consulta") Then
        txtBusqueda = UCase$(txtBusqueda)
        Busqueda txtBusqueda
        cboBusqueda.SetFocus
        
        If ComisionAFP = "0" Then
         optComiAFP(0).Value = True
        Else
         optComiAFP(1).Value = True
        End If
    End If
 End If
End Sub


Private Sub txtCalzado_GotFocus()
    mark1 txtCalzado
End Sub

Private Sub txtCalzado_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        cboTMame.SetFocus
    End If
 End If
End Sub


Private Sub txtCargo_GotFocus()
    mark1 txtcargo
End Sub

Private Sub txtcargoC_GotFocus()
     mark1 txtcargoC
End Sub

Private Sub txtCargoC_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtcargoC = Right("000" & txtcargo, 3)
        lblcargoC = DescripcionesdeCodigos("CARGOS", Trim(txtcargo))
        'txtObs.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyF1 Then
         With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 4000
            .pTitulo = "Cargos"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_CARGOS
            .Show
        End With
    End If
  End If
End Sub



Private Sub txtCargo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtcargo = Right("000" & txtcargo, 3)
        lblcargo = DescripcionesdeCodigos("CARGOS", Trim(txtcargo))
        txtObs.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyF1 And txtcargo.BackColor = ColorHabilitado Then
        Exit Sub
    End If
  End If
End Sub

Private Sub txtCarnetExt_GotFocus()
    mark1 txtCarnetExt
End Sub

Private Sub txtCarnetExt_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtTipoDoc.SetFocus
    End If
 End If
End Sub

Private Sub txtcencos_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 And txtcencos.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1200
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Centros de Costos"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_CENCOS
            .Show
        End With
    End If
End Sub

Private Sub txtcencos_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        lblcencos = DescripcionesdeCodigos("CENCO", Trim(txtcencos), "1")
        txtdivc.SetFocus
    End If
End Sub

Private Sub txtCTSCta_GotFocus()
    mark1 txtCTSCta
End Sub

Private Sub txtCTSCta_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtAfp.SetFocus
    End If
End Sub

Private Sub txtDirApo_GotFocus()
    mark1 txtDirApo
End Sub

Private Sub txtDirApo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtFijoApo.SetFocus
    End If
 End If
End Sub



Private Sub txtdivc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 And txtdivc.BackColor = ColorHabilitado Then
         With oConsulta
            .pCols = 6
            .pCol = 0: .pAnchoCol = 1000
            .pCol = 1: .pAnchoCol = 0
            .pCol = 2: .pAnchoCol = 0
            .pCol = 3: .pAnchoCol = 0
            .pCol = 4: .pAnchoCol = 0
            .pCol = 5: .pAnchoCol = 3500
            .pTitulo = "ccHFM"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_DIVISIONES_LOCAL
            .Show
        End With
    End If
    If KeyCode = 13 Then
        txtdivc = Right("0000" & Trim(txtdivc), 4)
        lbldivc = DescripcionesdeCodigos("DES_DIVISIONLOCAL", Trim(txtdivc))
    End If
End Sub

Private Sub txtdivc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboDiv.SetFocus
End Sub



Private Sub txtDpto_Change()
    If txtDpto = Empty Then
        lblDpto = Empty
    End If
End Sub

Private Sub txtDpto_GotFocus()
    mark1 txtDpto
End Sub

Private Sub txtDpto_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtDpto = Right("00" & txtDpto, 2)
        lblDpto = DescripcionesdeCodigos("DEPARTAMENTOS", Trim(txtDpto))
        txtDistrito.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyF1 And txtDpto.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 4000
            .pTitulo = "DPTOS"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_DPTO
            .Show
        End With
    End If
  End If
End Sub

Private Sub txtEdad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboGenero.SetFocus
    End If
End Sub

Private Sub txtEdad_LostFocus()
    If dtpNacimiento.Value <> Nulo Then
         txtEdad = Right("00" & Trim(val(Year(Date) - Right(Trim(dtpNacimiento), 4))), 2)
        If val(txtEdad) <= 80 And val(txtEdad) >= 18 Then
            txtEdad = txtEdad
        Else
            txtEdad = "00"
        End If
    Else
        txtEdad = "00"
    End If
End Sub

Private Sub txtEmpleado_GotFocus()
    mark1 txtEmpleado
End Sub

Private Sub BusqxCod(codigo As String)
    Dim SQL As String
    Dim rsdatos As MYSQL_RS
    Dim I As Integer
    SQL = "Select * from empleado where codigo = '" & codigo & "' and codigo <> '00000000000' "
    Set rsdatos = oConexion.EjecutaSelectRS(SQL)
    If rsdatos.RecordCount = 1 Then
        rsgral.MoveFirst
        For I = 1 To rsgral.RecordCount
            Do While Not rsgral.EOF
                If rsgral.Fields("CODIGO") = codigo Then
                     I = rsgral.AbsolutePosition
                     rsgral.AbsolutePosition = I
                     ConfigBtnsBusq I, rsgral.RecordCount
                     Exit For
                End If
                rsgral.MoveNext
            Loop
        Next
        CargarDatos rsdatos
    Else
        CargaCnAuxil codigo
    End If
    Set rsdatos = Nothing
End Sub

Private Sub Busqueda(Dato As String)
    Dim SQL As String
    Dim rsdatos As MYSQL_RS
    Dim I As Integer
    Dim criterio As String
    If cboBusqueda.ListCount > 1 Then
        criterio = cboBusqueda.List(cboBusqueda.ListIndex, 1)
    Else
        criterio = "0"
    End If
    
    
    Select Case criterio
            Case "APEPAT"
                 SQL = "Select e.*,a.* from empleado as e left join empleado_apoderado as a on e.codigo = a.codemp where " & criterio & " like concat('" & UCase(Dato) & "','%') order by apepat,apemat,nombre1,nombre2"
            Case "APEMAT"
                 SQL = "Select e.*,a.* from empleado as e left join empleado_apoderado as a on e.codigo = a.codemp where " & criterio & " like concat('" & UCase(Dato) & "','%') order by apepat,apemat,nombre1,nombre2"
            Case "NOMBRE1"
                 SQL = "Select e.*,a.* from empleado as e left join empleado_apoderado as a on e.codigo = a.codemp where " & criterio & " like concat('" & UCase(Dato) & "','%') order by apepat,apemat,nombre1,nombre2"
            Case "NOMBRE2"
                 SQL = "Select e.*,a.* from empleado as e left join empleado_apoderado as a on e.codigo = a.codemp where " & criterio & " like concat('" & UCase(Dato) & "','%') order by apepat,apemat,nombre1,nombre2"
            Case "NUMDOCIDE"
                 SQL = "Select e.*,a.* from empleado as e left join empleado_apoderado as a on e.codigo = a.codemp where " & criterio & " like concat('" & UCase(Dato) & "','%') order by apepat,apemat,nombre1,nombre2"
            Case "CODIGOHCM"
                 SQL = "Select e.*,a.* from empleado as e left join empleado_apoderado as a on e.codigo = a.codemp where " & criterio & " like concat('" & UCase(Dato) & "','%') order by apepat,apemat,nombre1,nombre2"
            Case "POS"
                 If Dato <> "" Then
                    If str(Dato) >= 1 And str(Dato) <= rsgral.RecordCount Then
                       If str(Dato) = 1 Then
                          rsgral.MoveFirst
                       Else
                          rsgral.Move CDbl(str(Dato) - 1), 1
                       End If
                    Else
                       Exit Sub
                    End If
                 End If
            Case "0"
                 Exit Sub
    End Select
    If criterio <> "POS" Then
        Set rsdatos = oConexion.EjecutaSelectRS(SQL)
        If Not rsdatos.EOF Then
            rsgral.MoveFirst
            For I = 1 To rsgral.RecordCount
                If UCase(rsgral.Fields(criterio)) Like UCase(Dato) & "*" Then
                     I = rsgral.AbsolutePosition
                     rsgral.AbsolutePosition = I
                     ConfigBtnsBusq I, rsgral.RecordCount
                     Exit For
                End If
                rsgral.MoveNext
            Next
            CargarDatos rsdatos
        End If
        If rsdatos.RecordCount = 0 Then
            MsgBox "No se encuentra el dato buscado", vbOKOnly + vbInformation, "NOVPeru"
            rsgral.MoveFirst
            ModoFormulario modConsulta
        End If
    Else
        CargarDatos rsgral
        ConfigBtnsBusq str(Dato), rsgral.RecordCount
    End If
    Set rsdatos = Nothing
End Sub

Private Sub ConfigBtnsBusq(AbsolutPos As Integer, RecCount As Integer)
    If AbsolutPos = 1 Then btnPrimero.Enabled = False: btnPrevio.Enabled = False: btnUltimo.Enabled = True: btnSgt.Enabled = True: Exit Sub
    If AbsolutPos = RecCount Then btnUltimo.Enabled = False: btnSgt.Enabled = False: btnPrimero.Enabled = True: btnPrevio.Enabled = True: Exit Sub
    If AbsolutPos > 1 And AbsolutPos < RecCount Then btnUltimo.Enabled = True: btnSgt.Enabled = True: btnPrimero.Enabled = True: btnPrevio.Enabled = True: Exit Sub
End Sub

Private Sub CargarDatos(rsdatos As MYSQL_RS)
    With rsdatos
        LimpiarDatos
        txtHCMEmpleado = CE(.Fields("codigoHCM"))
        txtEmpleado = Right("00000000000" & .Fields("codigo"), 11)
        txtEmpleado.tag = CE(.Fields("codigo"))
        lblEmpleado = UCase(CE(.Fields("apepat"))) & " " & UCase(CE(.Fields("apemat"))) & " " & UCase(CE(.Fields("nombre1"))) & " " & UCase(CE(.Fields("nombre2")))
        frmRegEmpleado.Caption = "Mantenimiento de Personal" & " - [" & lblEmpleado.Caption & "]"
        txtNombre1 = UCase(CE(.Fields("nombre1")))
        txtNombre2 = UCase(CE(.Fields("nombre2")))
        txtApePat = UCase(CE(.Fields("apepat")))
        txtApeMat = UCase(CE(.Fields("apemat")))
        If CE(.Fields("cafp")) = "F" Then
            optComiAFP(0).Value = True
            optComiAFP(1).Value = False
            ComisionAFP = "0"
        End If
        If CE(.Fields("cafp")) = "M" Then
            optComiAFP(0).Value = False
            optComiAFP(1).Value = True
            ComisionAFP = "1"
        End If
        If IsDate(.Fields("FEC_NAC")) Then
            dtpNacimiento = Format(.Fields("FEC_NAC"), "dd/mm/yyyy")
        Else
            dtpNacimiento = Empty
        End If
        
        txtEdad.Text = Right("00" & CEN(.Fields("edad")), 2)
        txtTipoDoc = CE(.Fields("CODDOCIDE"))
        lblTipoDoc = DescripcionesdeCodigos("TIPODOCIDE", Trim(txtTipoDoc))
        txtNumDoc = CE(.Fields("NUMDOCIDE"))
        txtCarnetExt = CE(.Fields("carnetext"))
        txtDireccion = CE(.Fields("DIRECCION"))
        txtDistrito = CE(.Fields("distrito"))
        lblDistrito = DescripcionesdeCodigos("DISTRITO", Trim(txtDistrito))
        txtDpto = CE(.Fields("departamento"))
        lblDpto = DescripcionesdeCodigos("DEPARTAMENTOS", Trim(txtDpto))
        'txtObs = CE(.Fields("obs"))
        txtmail = CE(.Fields("mail"))
        txtmailper = CE(.Fields("mailper"))
        'txtGrado = CE(.Fields("codgrado"))
        'lblGrado = DescripcionesdeCodigos("GRADOS", Trim(txtGrado))
        'txtTitulo = CE(.Fields("codtit"))
        'lblTitulo = DescripcionesdeCodigos("TITULOS", Trim(txtTitulo))
        lblSituacEmp = DescripcionesdeCodigos("SITUACIONEMP", CE(.Fields("SITUACION")))
        txtNumHijos.Text = Right("0" & Trim(CEN(.Fields("NUM_HIJOS"))), 1)
        txtFonoFijo.Text = Replace(CE(.Fields("fonofijo")), "-", "")
        txtFonoMov.Text = CE(.Fields("fonomovil"))
        txtPasaporte = CE(.Fields("pasaporte"))
        txtEstatura = FormatNumber(IIf(CE(.Fields("estatura")) = "", 0, CE(.Fields("estatura"))), 2)
        txtPeso.Text = Format(IIf(CE(.Fields("peso")) = "", 0, CE(.Fields("peso"))), "00#.00")
        txtCalzado = Format(IIf(CE(.Fields("calzado")) = "", 0, CE(.Fields("calzado"))), "0#.0")
        txtBrevete = CE(.Fields("brevete"))
        txtApoderado = CE(.Fields("nomapo"))
        txtDirApo = CE(.Fields("direcapo"))
        txtFijoApo = CE(.Fields("fonoapo"))
        txtMovilApo = CE(.Fields("movilapo"))
        txtNroDocParen = CE(.Fields("nrodocApo"))
        txtemailApo = CE(.Fields("emailApo"))
        For I = 1 To cboNacion.ListCount - 1
            If .Fields("nacionalidad") = cboNacion.List(I, 1) Then
                cboNacion.ListIndex = I
                Exit For
            Else
                cboNacion.ListIndex = 0
            End If
        Next
        For I = 0 To cboBreveteCat.ListCount - 1
            If .Fields("tipbrevete") = cboBreveteCat.List(I, 1) Then
                cboBreveteCat.ListIndex = I
                Exit For
             Else
                cboBreveteCat.ListIndex = 0
            End If
        Next
        For I = 0 To cboGSanguineo.ListCount - 1
            If .Fields("gsangre") = cboGSanguineo.List(I, 1) Then
                cboGSanguineo.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To cboTMame.ListCount - 1
            If .Fields("mameluco") = cboTMame.List(I, 1) Then
                cboTMame.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To cboGenero.ListCount - 1
            If .Fields("sexo") = cboGenero.List(I, 1) Then
                cboGenero.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To cboEstCivil.ListCount - 1
            If .Fields("Est_Civil") = cboEstCivil.List(I, 1) Then
                cboEstCivil.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To cboCategoria.ListCount - 1
            If .Fields("Categoria") = cboCategoria.List(I, 1) Then
                cboCategoria.ListIndex = I
                Exit For
            End If
        Next
        For I = 1 To cboTipoDocPar.ListCount - 1
            If .Fields("tipodocApo") = cboTipoDocPar.List(I, 1) Then
                cboTipoDocPar.ListIndex = I
                Exit For
            Else
                cboTipoDocPar.ListIndex = 0
            End If
        Next
        For I = 1 To cboParenContacto.ListCount - 1
            If .Fields("relacionApo") = cboParenContacto.List(I, 1) Then
                cboParenContacto.ListIndex = I
                Exit For
            Else
                cboParenContacto.ListIndex = 0
            End If
        Next
        
        
        dtpInicio = Empty
        dtpFin = Empty
        dtpFecIns = Empty
        If IsDate(.Fields("FEC_INGRESO")) Then
            dtpIngreso = Format(.Fields("FEC_INGRESO"), "dd/mm/yyyy")
        Else
            dtpIngreso = Empty
        End If
        If IsDate(.Fields("FEC_CESE")) Then
            dtpCese = Format(.Fields("FEC_CESE"), "dd/mm/yyyy")
        Else
            dtpCese = Empty
        End If
        If IsDate(.Fields("FECINGAFP")) Then
            dtpFecIns = Format(.Fields("FECINGAFP"), "dd/mm/yyyy")
        Else
            dtpFecIns = Empty
        End If
        txtAfp = CE(.Fields("CODAFP"))
        txtNumAfp = CE(.Fields("NUMAFP"))
        txtAsigFam = CE(.Fields("ASIGFAM"))
        If txtAsigFam <> "N" Then
            chkAsigFam.Value = True
        Else
            chkAsigFam.Value = False
        End If
        If .Fields("sctr") = "S" Then
            chkSctr.Value = True
        Else
            chkSctr.Value = False
        End If
        If .Fields("svl") = "S" Then
            chkSvl.Value = True
        Else
            chkSvl.Value = False
        End If
        If .Fields("jubilado") = "S" Then
            chkJubil.Value = True
        Else
            chkJubil.Value = False
        End If
        txtBanco = CE(.Fields("CODBANCO"))
        txtNumCta = CE(.Fields("NUMCTA"))
        lblBanco = DescripcionesdeCodigos("BANCO", txtBanco)
     
        For I = 1 To cboTipCtaMe.ListCount - 1
            If .Fields("TIPCTA_ME") = cboTipCtaMe.List(I, 1) Then
                cboTipCtaMe.ListIndex = I
                Exit For
            End If
        Next
        For I = 1 To cboTipCtaMn.ListCount - 1
            If .Fields("TIPCTA_MN") = cboTipCtaMn.List(I, 1) Then
                cboTipCtaMn.ListIndex = I
                Exit For
            End If
        Next
        For I = 0 To cboSituacion.ListCount - 1
            If CE(.Fields("SITUACION")) = Trim(cboSituacion.List(I, 1)) Then
                cboSituacion.ListIndex = I
                Exit For
            Else
                cboSituacion.ListIndex = 0
            End If
        Next
        For I = 0 To cboTipo.ListCount - 1
            If CE(.Fields("tipo")) = Trim(cboTipo.List(I, 1)) Then
                cboTipo.ListIndex = I
                Exit For
            Else
                cboTipo.ListIndex = 0
            End If
        Next
        For I = 0 To cboPersonal.ListCount - 1
            If CE(.Fields("personal")) = Trim(cboPersonal.List(I, 1)) Then
                cboPersonal.ListIndex = I
                Exit For
            Else
                cboPersonal.ListIndex = 0
            End If
        Next
        For I = 1 To CboMonCTS.ListCount - 1
            If CE(.Fields("CTSMON")) = Trim(CboMonCTS.List(I, 1)) Then
                CboMonCTS.ListIndex = I
                Exit For
            Else
                CboMonCTS.ListIndex = 0
            End If
        Next
        txtBancoCTS = .Fields("CTSBANCO")
        lblBancoCTS = DescripcionesdeCodigos("BANCO", txtBancoCTS)
        txtCTSCta = .Fields("CTSNUMCTA")
        txtNumCtaMn = CE(.Fields("NUMCTA_MN"))
        txtNumCtaMe = CE(.Fields("NUMCTA_ME"))
        lblAfp = DescripcionesdeCodigos("AFP", txtAfp, "Descrip")
        strFoto = .Fields("Foto")
        
        
        VerFoto txtNumDoc
        

        If rsgral.RecordCount > 1 Then
            lblPrimero = "1"
            lblTotal = rsgral.RecordCount
            lblCuenta = rsgral.AbsolutePosition '& " / " & rsgral.RecordCount
        End If
        CargarFlx Trim(txtEmpleado)
        ModoFormulario modConsulta
        CargaContrato (Trim(txtEmpleado))
        
        txtcargo = CE(txtcargoC)
        lblcargo = DescripcionesdeCodigos("CARGOS", Trim(txtcargo))
        
        CargarSueldos Trim(txtEmpleado), Trim(txtContrato)
        CargarRetenciones Trim(txtEmpleado)
        CargarMovilidades Trim(txtEmpleado)
        CargarFormacionEducativa Trim(txtEmpleado)
        CargaAutorizaciones Trim(txtEmpleado)
        CargaVerificado Trim(txtEmpleado)
        CargarSeguros Trim(txtEmpleado)
        If chkSvl.Value = True Then MostrarPoliza "L"
        If chkSctr.Value = True Then MostrarPoliza "R"
        If BtnEliminar.tag <> "" Then BtnEliminar.Enabled = BtnEliminar.tag Else: BtnEliminar.Enabled = False
    End With
End Sub

Sub MostrarPoliza(Tipo As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    SQL = "Select * from polizas where tipo = '" & Tipo & "' order by fecini"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        If Tipo = "L" Then
            lblnumsvl = Trim(RQ.Fields("numpoliza"))
        Else
            lblnumsctr = Trim(RQ.Fields("numpoliza"))
        End If
    End If
    Set RQ = Nothing
End Sub

Private Sub CargaVerificado(codigo As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    SQL = "Select * from rh_verificaessalud where codemp = '" & codigo & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        chkver.Value = IIf(RQ.Fields("verificado") = "S", True, False)
    End If
    Set RQ = Nothing
End Sub

Private Sub CargaAutorizaciones(codigo As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    lblautoriza = "POR AUTORIZAR: "
    SQL = "Select c.descrip,r.codigo,r.sueldo,r.tipocese,r.fechacese,p.descrip as cese from rh_tempacceso r left join configuracion_acceso c " & _
          "on(r.codigo=c.codigo) left join pl_tipocese p on(r.tipocese=p.codigo) where codemp = '" & codigo & "' and autorizado = 'N' order by r.codigo"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        Do While Not RQ.EOF
            Select Case RQ.Fields("codigo")
                Case 2:
                    lblautoriza = lblautoriza & " ACTIVAR EMPLEADO /"
                Case 4
                    lblautoriza = lblautoriza & " CESE:" & Format(Trim(RQ.Fields("fechacese")), "dd/mm/yyyy") & " " & Trim(RQ.Fields("cese")) & " /"
                Case 7
                    lblautoriza = lblautoriza & " SUELDO:" & FormatNumber(RQ.Fields("SUELDO"), 2)
            End Select
            RQ.MoveNext
        Loop
        lblautoriza = IIf(Mid(lblautoriza, Len(lblautoriza), 1) = "/", Mid(lblautoriza, 1, Len(lblautoriza) - 1), lblautoriza)
        cmdElimSoli.Enabled = True
    Else
        lblautoriza = ""
        cmdElimSoli.Enabled = False
    End If
    Set RQ = Nothing
End Sub

Private Sub moneda(cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = "0"
    cbo.AddItem "Nacional"
    cbo.List(1, 1) = "N"
    cbo.AddItem "Extranjera"
    cbo.List(2, 1) = "E"
    cbo.ListIndex = 0
End Sub

Private Sub monedaCCI(cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = "0"
    cbo.AddItem "Nacional"
    cbo.List(1, 1) = "N"
    cbo.AddItem "Extranjera"
    cbo.List(2, 1) = "E"
    cbo.AddItem "CCI"
    cbo.List(3, 1) = "C"
    cbo.ListIndex = 0
End Sub

'upload de la información del contrato,carga el ultimo contrato activo
Private Sub CargaContrato(CodEmp As String)
    Dim SQL As String
    Dim rscont As MYSQL_RS
    Dim contrato As String
    Dim I As Integer
    Dim Divi As String
    cboContratos.ListIndex = 0
    cboMonBono.ListIndex = 0
    cboMonMovi.ListIndex = 0
    cboMonSueldo.ListIndex = 0
    cboDiv.ListIndex = 0
    ConfiguracboTipo
    SQL = "Select MAX(CODIGO) as maximo from contrato where codemp = '" & CodEmp & "'"
    Set rscont = oConexion.EjecutaSelectRS(SQL)
    If Not IsNull(rscont.Fields("MAXIMO")) Then
        contrato = CE(rscont.Fields("maximo"))
    End If
    If IsNull(rscont.Fields("MAXIMO")) Then
        If (lblModo = "Acción" Or lblModo = "Consulta") Then
            txtContrato.tag = ""
            btnContrato.tag = "True"
            BtnEliminar.tag = "True"
            btnRenovar.tag = "False"
            btnCancelCont.tag = "False"
            btnHabilitar.tag = "False"
        End If
        If (lblModo = "Nuevo") Then
            txtContrato.tag = ""
            btnContrato.tag = "True"
            BtnEliminar.tag = "True"
            btnRenovar.tag = "False"
            btnCancelCont.tag = "False"
            btnHabilitar.tag = "False"
        End If
        Set rscont = Nothing
        lblEstContrato = "ULTIMO CONTRATO PENDIENTE"
        Exit Sub
    End If
    Set rscont = Nothing
    SQL = "Select * from contrato where codigo = '" & contrato & "'  and codemp = '" & CodEmp & "'"
    Set rscont = oConexion.EjecutaSelectRS(SQL)
    If rscont.RecordCount = 1 Then
        With rscont
            txtContrato = CE(.Fields("CODIGO"))
            txtContrato.tag = CE(.Fields("ESTADO"))
            txtTipoCont = CE(.Fields("CODTIPO"))
            lblTipoCont = DescripcionesdeCodigos("CNCONTRATO", Trim(CE(.Fields("CODTIPO"))))
            lblsbasico = FormatNumber(CEN(.Fields("SBASICO")), 2)
            If IsDate(Format(.Fields("F_INICIO"), "dd/mm/yyyy")) Then
                dtpInicio.Value = Format(.Fields("F_INICIO"), "dd/mm/yyyy")
            End If
            If IsDate(Format(.Fields("F_TERMINO"), "dd/mm/yyyy")) Then
                dtpFin.Value = Format(.Fields("F_TERMINO"), "dd/mm/yyyy")
            End If
            If CE(.Fields("BONO")) = "S" Then
                chkBono.Value = True
            Else
                chkBono.Value = False
            End If
            For I = 0 To cboContratos.ListCount - 1
                If cboContratos.List(I, 1) = CE(.Fields("CODTIPO")) Then
                    cboContratos.ListIndex = I
                    Exit For
                End If
            Next
            For I = 0 To cboMonSueldo.ListCount - 1
                If cboMonSueldo.List(I, 1) = CE(.Fields("MON_SUELDO")) Then
                    cboMonSueldo.ListIndex = I
                    Exit For
                End If
            Next
            For I = 0 To cboMonBono.ListCount - 1
                If cboMonBono.List(I, 1) = CE(.Fields("MON_BONO")) Then
                    cboMonBono.ListIndex = I
                    Exit For
                End If
            Next
            For I = 0 To cboHorLab.ListCount - 1
                If cboHorLab.List(I, 0) = CE(.Fields("HorLab")) Then
                    cboHorLab.ListIndex = I
                    Exit For
                End If
            Next
            For I = 0 To CboTipIng.ListCount - 1
                If .Fields("TipSueldo") = DameTipoSuel(CboTipIng.List(I, 0)) Then
                    CboTipIng.ListIndex = I
                    Exit For
                End If
            Next
            
            For I = 0 To cboEstTrabajo.ListCount - 1
                If .Fields("EstTrabajo") = cboEstTrabajo.List(I, 1) Then
                    cboEstTrabajo.ListIndex = I
                    Exit For
                End If
            Next
            
            Divi = .Fields("divgas")
            For I = 0 To cboDiv.ListCount - 1
                If Divi = cboDiv.List(I, 1) Then
                    cboDiv.ListIndex = I
                    Exit For
                End If
            Next
            txtdivc = CE(.Fields("division"))
            lbldivc = DescripcionesdeCodigos("DES_DIVISIONLOCAL", Trim(txtdivc))
            'Act. Cargo en Contrato
            txtcargoC = CE(.Fields("codcargo"))
            lblcargoC = DescripcionesdeCodigos("CARGOS", Trim(txtcargoC))
            
            txtcencos = CE(.Fields("cencos"))
            lblcencos = DescripcionesdeCodigos("CENCO", Trim(txtcencos), "1")
            txtMontoBono = FormatNumber(CEN(.Fields("MONTO_BONO")), 2)
            For I = 0 To Cbotipocese.ListCount - 1
                If .Fields("codtipocese") = Cbotipocese.List(I, 1) Then
                    Cbotipocese.ListIndex = I
                    Exit For
                End If
            Next
            If IsDate(Format(.Fields("fechacese"), "dd/mm/yyyy")) Then
                dtfechacese.Value = Format(.Fields("fechacese"), "dd/mm/yyyy")
            End If
            Select Case CE(.Fields("ESTADO"))
                 Case PENDIENTE
                    btnHabilitar.tag = "True"
                    btnCancelCont.tag = "True"
                    btnRenovar.tag = "False"
                    btnContrato.tag = "False"
                    BtnEliminar.tag = "True"
                    btnGrabarC.tag = "True"
                    lblEstContrato = "ULTIMO CONTRATO PENDIENTE"
                    cmdotros.Enabled = True
                Case APROBADO
                    btnHabilitar.tag = "False"
                    btnCancelCont.tag = "True"
                    btnContrato.tag = "False"
                    If IsDate(meFecFin) Then
                        If CDate(meFecFin) < Date Then
                            btnRenovar.tag = "True"
                        Else
                            btnRenovar.tag = "False"
                        End If
                    Else
                        btnRenovar.tag = "True"
                    End If
                    BtnEliminar.tag = "False"
                    lblEstContrato = "ULTIMO CONTRATO APROBADO"
                    cmdotros.Enabled = True
                Case CANCELADO
                    btnHabilitar.tag = "False"
                    btnCancelCont.tag = "False"
                    btnContrato.tag = "False"
                    btnRenovar.tag = "True"
                    lblEstContrato = "ULTIMO CONTRATO CANCELADO"
                    cmdotros.Enabled = True
                Case Else
                    btnHabilitar.tag = "False"
                    btnCancelCont.tag = "False"
                    btnRenovar.tag = "False"
                    btnContrato.tag = "True"
                    lblEstContrato = "SIN CONTRATO"
                    cmdotros.Enabled = True
            End Select
        End With
    Else
          lblEstContrato = "SIN CONTRATO"
    End If
    Set rscont = Nothing
End Sub

Private Sub BotonesContrato()
    If btnHabilitar.tag <> "" Then btnHabilitar.Enabled = btnHabilitar.tag Else: btnHabilitar.Enabled = False
    If btnHabilitar.Enabled Then btnGrabarC.Enabled = True
    If btnCancelCont.tag <> "" Then btnCancelCont.Enabled = btnCancelCont.tag Else: btnCancelCont.Enabled = False
    If btnRenovar.tag <> "" Then btnRenovar.Enabled = btnRenovar.tag Else: btnRenovar.Enabled = False
    If btnContrato.tag <> "" Then btnContrato.Enabled = btnContrato.tag Else: btnContrato.Enabled = False
End Sub

Private Function CargaCnAuxil(codigo As String) As Boolean
    Dim SQL As String
    Dim rscod As MYSQL_RS
    Dim RES As Integer
    CargaCnAuxil = False
    SQL = "Select * from cnauxil where codigo = '" & codigo & "' and auxiliar = '3' and codigo <> '00000000000' "
    Set rscod = oConexion.EjecutaSelectRS(SQL)
    If rscod.RecordCount = 1 Then
        With rscod
            ModoFormulario modNuevo
            txtEmpleado = codigo
            If CE(.Fields("RUC")) <> Empty Then
                txtRuc = CE(.Fields("ruc"))
            End If
            lblEmpleado = CE(.Fields("descrip"))
            txtDireccion = CE(.Fields("direcc"))
            CargaCnAuxil = True
            Exit Function
            Set rscod = Nothing
        End With
    End If
    If rscod.RecordCount = 0 And lblModo = "Nuevo" Then
        CargaCnAuxil = False
        Exit Function
    End If
    If rscod.RecordCount = 0 Then
        If lblModo = "Consulta" Or lblModo = "Acción" Then
            MsgBox "El código ingresado, no se encuentra registrado", vbInformation, gsNomSW
            LimpiarDatos
            txtEmpleado = txtEmpleado.tag
            BusqxCod Trim(txtEmpleado.tag)
            txtEmpleado.SetFocus
        End If
    End If
End Function

Private Function GenCodEmp() As String
    Dim SQL As String
    Dim rsgen As MYSQL_RS
    Dim Cont As Integer
    Dim CodEmp As String
    SQL = "SELECT MAX(CODIGO) AS MAXIMO FROM EMPLEADO "
    Set rsgen = oConexion.EjecutaSelectRS(SQL)
    If rsgen.RecordCount > 0 Then
        GenCodEmp = Right("00000000000" + RTrim(CStr(CDbl(rsgen.Fields("MAXIMO")) + 1)), 11)
    Else
        GenCodEmp = "00000000001"
    End If
    Set rsgen = Nothing
End Function

Private Sub CargarFlx(codigo As String)
    Dim SQL As String
    Dim I As Integer
    Dim rsflx As MYSQL_RS
    SQL = "select distinct f.*,r.verificado from familiares f left join rh_verificaessalud r " & _
          "on (f.codemp=r.codemp) AND (f.item=r.item) where f.codemp = '" & codigo & "' and r.familiar = 'F' order by f.codemp,item"
    'sql = "Select * from familiares where codemp = '" & codigo & "' order by item"
    Set rsflx = oConexion.EjecutaSelectRS(SQL)
    ConfiguraGrilla
    Do While Not rsflx.EOF
        With flxDependientes
            For I = 1 To rsflx.RecordCount
                .TextMatrix(I, 1) = CE(rsflx.Fields("nombre"))
                .TextMatrix(I, 2) = CE(rsflx.Fields("apepaterno"))
                .TextMatrix(I, 3) = CE(rsflx.Fields("apematerno"))
                .TextMatrix(I, 4) = Format(rsflx.Fields("fec_nac"), "dd/mm/yyyy")
                .TextMatrix(I, 5) = CE(rsflx.Fields("edad"))
                .TextMatrix(I, 6) = gen(CE(rsflx.Fields("sexo")))
                .TextMatrix(I, 7) = CE(rsflx.Fields("coddocide"))
                .TextMatrix(I, 8) = CE(rsflx.Fields("numdocide"))
                .TextMatrix(I, 9) = DescripcionesdeCodigos("PARENTESCO", CE(rsflx.Fields("codtipo")))
                .TextMatrix(I, 10) = CE(rsflx.Fields("codtipo"))
                .Col = 11: .row = I
                .CellFontName = "Wingdings"
                .CellFontSize = 14
                .TextMatrix(I, 11) = IIf(rsflx.Fields("verificado") = "S", strChecked, strUnChecked)
                EnumerarItems flxDependientes
                .Rows = .Rows + 1
                rsflx.MoveNext
            Next
        End With
    Loop
    Set rsflx = Nothing
End Sub

Private Sub CargarRetenciones(codigo As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    SQL = "Select * from retenciones where codemp = '" & codigo & "' order by item"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    ConfiguraGrillaRet
    Do While Not RQ.EOF
        With flxret
            For I = 1 To RQ.RecordCount
                .TextMatrix(I, 1) = CE(RQ.Fields("nombres"))
                .TextMatrix(I, 2) = CE(RQ.Fields("moneda"))
                .TextMatrix(I, 3) = CEN(RQ.Fields("porcentaje"))
                .TextMatrix(I, 4) = CE(RQ.Fields("monto"))
                EnumerarItems flxret
                .Rows = .Rows + 1
                RQ.MoveNext
            Next
        End With
    Loop
    CalculoTotalRet
    Set RQ = Nothing
End Sub

Private Sub CargarMovilidades(codigo As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    txtMontoMov.Text = ""
    chkmovi.Value = "0"
    
    SQL = "Select * from Movilidades where codemp = '" & codigo & "' order by item"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If RQ.EOF = False Then
        chkmovi.Value = "1"
        txtMontoMov.Text = CE(RQ.Fields("monto"))
        For I = 0 To cboMonMovi.ListCount - 1
          If cboMonMovi.List(I, 1) = CE(RQ.Fields("moneda")) Then
            cboMonMovi.ListIndex = I
            Exit For
          End If
        Next
    End If
    
    Set RQ = Nothing
End Sub


Private Sub CargarFormacionEducativa(codigo As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    
    SQL = "Select * from pl_empformedu where codemp = '" & codigo & "' and visible='S' limit 1"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    
    txtAnhoEgr.Text = ""
    
    If RQ.EOF = False Then
        
        cboFormSup.ListIndex = 0
        For I = 0 To cboFormSup.ListCount - 1
          If cboFormSup.List(I, 1) = CE(RQ.Fields("codformsup")) Then
            cboFormSup.ListIndex = I
            Exit For
          End If
        Next
        
        I = 0
        cboTipInst.ListIndex = 0
        For I = 0 To cboTipInst.ListCount - 1
          If cboTipInst.List(I, 1) = CE(RQ.Fields("flgTipoInst")) Then
            cboTipInst.ListIndex = I
            Exit For
          End If
        Next
        
        I = 0
        cboNomInst.ListIndex = 0
        For I = 0 To cboNomInst.ListCount - 1
          If cboNomInst.List(I, 1) = CE(RQ.Fields("coduni")) Then
            cboNomInst.ListIndex = I
            Exit For
          End If
        Next
        
        I = 0
        cboCarrera.ListIndex = 0
        For I = 0 To cboCarrera.ListCount - 1
          If cboCarrera.List(I, 1) = CE(RQ.Fields("codcarr")) Then
            cboCarrera.ListIndex = I
            Exit For
          End If
        Next
        
        I = 0
        cboEstPrin.ListIndex = 0
        For I = 0 To cboEstPrin.ListCount - 1
          If cboEstPrin.List(I, 1) = CE(RQ.Fields("visible")) Then
            cboEstPrin.ListIndex = I
            Exit For
          End If
        Next
        
        txtAnhoEgr.Text = CE(RQ.Fields("anhoegr"))
        
         If CE(RQ.Fields("flgEstPeru")) = "S" Then
           OptPeruSi.Value = True
           OptPeruNo.Value = False
         Else
           OptPeruSi.Value = False
           OptPeruNo.Value = True
         End If
         
         If CE(RQ.Fields("flgRegInst")) = "N" Then
           OptRegEduPub.Value = 1
           OptRegEduPriv.Value = 0
         Else
           OptRegEduPub.Value = 0
           OptRegEduPriv.Value = 1
         End If
    Else
         TipoInstEduc cboTipInst
         EstudPrin cboEstPrin
         FormSup cboFormSup
         NomInst cboNomInst
         DameCarrera cboCarrera
         txtAnhoEgr.Text = ""
    End If
    
    Set RQ = Nothing
End Sub

Sub CalculoTotalRet()
Dim SumPor As Double, SumMto As Double
    With flxret
        For I = 1 To .Rows - 1
            SumPor = SumPor + CDbl(IIf(val(.TextMatrix(I, 3)) = 0, 0, .TextMatrix(I, 3)))
            SumMto = SumMto + CDbl(IIf(val(.TextMatrix(I, 4)) = 0, 0, .TextMatrix(I, 4)))
        Next
        lbltotalpor = SumPor
        lbltotalmto = FormatNumber(SumMto, 2)
    End With
End Sub

Private Sub CargarSueldos(codigo As String, contrato As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    SQL = "Select * from rh_sueldos where codemp = '" & codigo & "' and codcont <= '" & contrato & "' order by FECHA"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    ConfiguraGrillaSueldos
    I = 1
    Do While Not RQ.EOF
        With flxsueldos
           .TextMatrix(I, 1) = FormatNumber(CEN(RQ.Fields("SUELDO")), 2)
           .TextMatrix(I, 2) = Format(RQ.Fields("fecha"), "dd/mm/yyyy")
           .Rows = .Rows + 1
           I = I + 1
            RQ.MoveNext
        End With
    Loop
    EnumerarItems flxsueldos
    Set RQ = Nothing
End Sub

Sub ConfiguraGrillaSeguros()
    With msfseg
        .Clear
        .Cols = 6
        .Rows = 2
        .FixedCols = 1
        .ColWidth(0) = 300
        .TextMatrix(0, 0) = "Item"
        .TextMatrix(1, 0) = "1"
        .ColType(0) = Numero
        .TextMatrix(1, 0) = "1"
        .ColWidth(1) = 350
        .TextMatrix(0, 1) = "Cod"
        .ColType(1) = cadena
        .ColMaxLength(1) = 2
        .CaracteresValidos(1) = "1234567890"
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = Space(4) & "Seguro"
        .ColType(2) = cadena
        .ColWidth(3) = 1600
        .TextMatrix(0, 3) = Space(10) & "Número"
        .ColType(3) = cadena
        .ColMaxLength(3) = 15
        .CaracteresValidos(3) = "ABCDEFGHIJKLMÑNOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz1234567890"
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Inicio Vig."
        .ColType(4) = fecha
        .ColMaxLength(4) = 10
        .CaracteresValidos(4) = "0123456789/"
        .ColWidth(5) = 1000
        .TextMatrix(0, 5) = Space(3) & "Fin Vig."
        .ColType(5) = fecha
        .ColMaxLength(5) = 10
        .CaracteresValidos(5) = "0123456789/"
    End With
End Sub

Private Sub CargarSeguros(codigo As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    SQL = "Select * from rh_segmedicoemp where codemp = '" & codigo & "' order by item"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    ConfiguraGrillaSeguros
    I = 1
    Do While Not RQ.EOF
        With msfseg
                .TextMatrix(I, 1) = Trim(RQ.Fields("codseg"))
                .TextMatrix(I, 2) = DescripcionesdeCodigos("SEGURO", Trim(RQ.Fields("codseg")))
                .TextMatrix(I, 3) = Trim(RQ.Fields("numpoliza"))
                .TextMatrix(I, 4) = Format(Trim(RQ.Fields("fecinivig")), "dd/mm/yyyy")
                .TextMatrix(I, 5) = Format(Trim(RQ.Fields("fecfinvig")), "dd/mm/yyyy")
                I = I + 1
                .Rows = .Rows + 1
                RQ.MoveNext
            
        End With
    Loop
    EnumerarItems msfseg
    Set RQ = Nothing
End Sub

Private Sub txtContrato_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        a.SetFocus
    End If
End Sub

Private Sub txtDireccion_Change()
    txtDireccion = UCase$(txtDireccion)
End Sub

Private Sub txtDireccion_GotFocus()
    mark1 txtDireccion
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then txtDpto.SetFocus
 End If
End Sub

Private Sub txtDistrito_Change()
    If txtDistrito = Empty Then lblDistrito = Empty
End Sub

Private Sub txtDistrito_GotFocus()
    mark1 txtDistrito
End Sub

Private Sub txtDistrito_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = vbKeyF1 And txtDistrito.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1000
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Distritos"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_DISTRITO
            .Show
        End With
    End If
    If KeyCode = 13 Then
        lblDistrito = Space(2) & DescripcionesdeCodigos("DISTRITO", Trim(txtDistrito))
        If lblDistrito = Empty Then
            txtDistrito = "0000"
        End If
        txtFonoFijo.SetFocus
    End If
  End If
End Sub

Private Sub txtEdad_GotFocus()
    mark1 txtEdad
End Sub

Private Sub txtEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And lblModo = "Nuevo" Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 4000
            .pTitulo = "Empleados"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_EMP
            .Show
        End With
    End If
    If KeyCode = vbKeyF1 And lblModo = "Acción" Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 4000
            .pTitulo = "Empleados Registrados"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_EMPREG
            .Show
        End With
    End If
    If KeyCode = vbKeyLeft Then
        If btnPrevio.Enabled Then
            btnPrevio.SetFocus
        End If
    End If
End Sub

Private Sub txtEmpleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And lblModo = "Nuevo" And Not generocod Then
        txtEmpleado = Right("00000000000" & Trim(txtEmpleado), 11)
        If Not CargaCnAuxil(txtEmpleado) Then
            txtRuc = txtEmpleado
            lblEmpleado = Empty
            txtNombre1.SetFocus
        Else
            BusqxCod txtEmpleado
            ModoFormulario modConsulta
            txtNombre1.SetFocus
        End If
    End If
    If KeyAscii = 13 And lblModo = "Nuevo" And generocod Then
        txtEmpleado = Right("00000000000" & Trim(txtEmpleado), 11)
        txtEmpleado.SelStart = 11
        txtNombre1.SetFocus
    End If
    If KeyAscii = 13 And (lblModo = "Acción" Or lblModo = "Consulta") Then
        txtEmpleado = Right("00000000000" & Trim(txtEmpleado), 11)
        BusqxCod txtEmpleado
        txtEmpleado.SetFocus
        txtEmpleado.SelStart = 0
        txtEmpleado.SelLength = Len(txtEmpleado.Text)
    End If
End Sub

Private Sub txtEstatura_GotFocus()
    mark1 txtEstatura
End Sub

Private Sub txtEstatura_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then txtPeso.SetFocus
 End If
End Sub

Private Sub txtFijoApo_GotFocus()
    mark1 txtFijoApo
End Sub

Private Sub txtFijoApo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then txtMovilApo.SetFocus
 End If
End Sub

Private Sub txtFonoFijo_GotFocus()
    mark1 txtFonoFijo
End Sub

Private Sub txtFonoFijo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then txtFonoMov.SetFocus
 End If
End Sub

Private Sub txtFonoMov_GotFocus()
    mark1 txtFonoMov
End Sub

Private Sub txtFonoMov_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then txtmail.SetFocus
 End If
End Sub

'Private Sub txtGrado_Change()
'    If txtGrado = Empty Then lblGrado = Empty
'End Sub

'Private Sub txtGrado_GotFocus()
'    mark1 txtGrado
'End Sub

'Private Sub txtGrado_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
' If KeyCode <> 17 And Shift <> 2 Then
'    If KeyCode = 13 Then
'        txtGrado = Right("00" & txtGrado, 2)
'        lblGrado = DescripcionesdeCodigos("GRADOS", Trim(txtGrado))
'        txtcargo.SetFocus
'        Exit Sub
'    End If
'    If KeyCode = vbKeyF1 And txtGrado.BackColor = ColorHabilitado Then
'         With oConsulta
'            .pCols = 2
'            .pCol = 0: .pAnchoCol = 1500
'            .pCol = 1: .pAnchoCol = 4000
'            .pTitulo = "Grados"
'            .pForm = FORM_REGEMP
'            .pCaso = LABEL_GRADOS
'            .Show
'        End With
'    End If
'  End If
'End Sub

Private Sub txtmail_GotFocus()
    mark1 txtmail
End Sub

Private Sub txtmail_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then txtmailper.SetFocus
 End If
End Sub

Private Sub txtmail_LostFocus()
    If txtmail <> Empty Then txtmail = LCase$(txtmail)
End Sub

Private Sub txtMontoBono_GotFocus()
    mark1 txtMontoBono
End Sub

Private Sub txtMontoBono_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtMontoBono = FormatNumber(txtMontoBono, 2)
        cboEstTrabajo.SetFocus
    End If
End Sub

Private Sub txtMontoBono_LostFocus()
    If txtMontoBono <> Empty And txtMontoBono.BackColor = ColorHabilitado Then
        txtMontoBono = FormatNumber(txtMontoBono, 2)
    End If
    If txtMontoBono = Empty Then txtMontoBono = "0.00"
End Sub

Private Sub txtMovilApo_GotFocus()
    mark1 txtMovilApo
End Sub
'
'Private Sub txtMovilApo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
' If KeyCode <> 17 And Shift <> 2 Then
'    If KeyCode = 13 Then txtTitulo.SetFocus
' End If
'End Sub

Private Sub txtNombre1_Change()
    txtNombre1 = UCase$(txtNombre1)
End Sub

Private Sub txtNombre2_Change()
    txtNombre2 = UCase$(txtNombre2)
End Sub

Private Sub genero(cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem "Seleccionar"
    cbo.List(0, 1) = "0"
    cbo.AddItem "Femenino"
    cbo.List(1, 1) = "F"
    cbo.AddItem "Masculino"
    cbo.List(2, 1) = "M"
    cbo.ListIndex = 0
End Sub

Private Sub TipoInstEduc(cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem "Seleccionar"
    cbo.List(0, 1) = "0"
    cbo.AddItem "Colegio"
    cbo.List(1, 1) = "C"
    cbo.AddItem "Instituto"
    cbo.List(2, 1) = "I"
    cbo.AddItem "Universidad"
    cbo.List(3, 1) = "U"
    cbo.AddItem "Curso"
    cbo.List(4, 1) = "C"
    cbo.AddItem "Diplomado"
    cbo.List(5, 1) = "D"
    cbo.AddItem "Especialización"
    cbo.List(6, 1) = "E"
    cbo.AddItem "CEO"
    cbo.List(7, 1) = "A"
    cbo.AddItem "CETPRO"
    cbo.List(8, 1) = "T"
    cbo.AddItem "Otro"
    cbo.List(9, 1) = "O"
    cbo.ListIndex = 0
End Sub

Private Sub EstudPrin(cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem "Seleccionar"
    cbo.List(0, 1) = "0"
    cbo.AddItem "No"
    cbo.List(1, 1) = "N"
    cbo.AddItem "Si"
    cbo.List(2, 1) = "S"
    cbo.ListIndex = 0
End Sub

Private Sub Estado(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsEstEmp As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from estadocivil order by codigo"
    Set rsEstEmp = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsEstEmp.EOF
        cbo.AddItem CE(rsEstEmp.Fields("descrip"))
        cbo.List(I, 1) = CE(rsEstEmp.Fields("codigo"))
        I = I + 1
        rsEstEmp.MoveNext
    Loop
    Set rsEstEmp = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub FormSup(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsEstEmp As MYSQL_RS
    Dim k As Integer
    k = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from pl_siteduc order by codsit"
    Set rsEstEmp = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsEstEmp.EOF
        cbo.AddItem CE(rsEstEmp.Fields("descrip"))
        cbo.List(k, 1) = CE(rsEstEmp.Fields("codsit"))
        k = k + 1
        rsEstEmp.MoveNext
    Loop
    Set rsEstEmp = Nothing
    cbo.ListIndex = 0
End Sub
 
Private Sub NomInst(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsEstEmp As MYSQL_RS
    Dim k As Integer
    k = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from pl_universia order by descrip"
    Set rsEstEmp = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsEstEmp.EOF
        cbo.AddItem CE(rsEstEmp.Fields("descrip"))
        cbo.List(k, 1) = CE(rsEstEmp.Fields("coduni"))
        k = k + 1
        rsEstEmp.MoveNext
    Loop
    Set rsEstEmp = Nothing
    cbo.ListIndex = 0
End Sub
 
Private Sub DameCarrera(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsEstEmp As MYSQL_RS
    Dim k As Integer
    k = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from pl_carrprof order by descrip"
    Set rsEstEmp = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsEstEmp.EOF
        cbo.AddItem CE(rsEstEmp.Fields("descrip"))
        cbo.List(k, 1) = CE(rsEstEmp.Fields("codcarr"))
        k = k + 1
        rsEstEmp.MoveNext
    Loop
    Set rsEstEmp = Nothing
    cbo.ListIndex = 0
End Sub
             
 
Private Sub Situacion(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsSitEmp As MYSQL_RS
    Dim I As Integer
    I = 1
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    SQL = "Select * from situacionemp order by codigo"
    Set rsSitEmp = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsSitEmp.EOF
        cbo.AddItem CE(rsSitEmp.Fields("descrip"))
        cbo.List(I, 1) = CE(rsSitEmp.Fields("codigo"))
        I = I + 1
        rsSitEmp.MoveNext
    Loop
    Set rsSitEmp = Nothing
    cbo.ListIndex = 0
End Sub

Private Sub ConfiguraGrilla()
    With flxDependientes
        .Clear
        .Refresh
        .Rows = 2
        .Cols = 12
        .RowHeight(1) = 315
        .ColWidth(0) = 400
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColType(0) = cadena
        .ColMaxLength(0) = 15
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(1) = 1500
        .TextMatrix(0, 1) = Space(9) + "Nombre"
        .ColType(1) = cadena
        .ColMaxLength(1) = 64
        .CaracteresValidos(1) = "ab cd" & Chr(13) & "efghijklmnopqrstuvwxyz" & UCase("abcdefghijklmnopqrstuvwxyz") & ""
        .ColWidth(2) = 2000
        .TextMatrix(0, 2) = Space(9) + "Apellido Pat."
        .ColType(2) = cadena
        .ColMaxLength(2) = 64
        .CaracteresValidos(2) = "ab cd" & Chr(13) & "efghijklmnopqrstñuvwxyz" & UCase("abcdefghiñjklmnopqrstuvwxyz") & ""
        .ColWidth(3) = 2000
        .TextMatrix(0, 3) = Space(9) + "Apellido Mat."
        .ColType(3) = cadena
        .ColMaxLength(3) = 64
        .CaracteresValidos(3) = "ab cd" & Chr(13) & "efghijklmnopqrstuvwxyz" & UCase("abcdefghijklmnopqrstuvwxyz") & ""
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = Space(1) + "Fecha Nac"
        .ColType(4) = cadena
        .ColMaxLength(4) = 10
        .CaracteresValidos(4) = "1234567890./"
        .ColWidth(5) = 500
        .TextMatrix(0, 5) = Space(0) + "Edad"
        .ColType(5) = cadena
        .ColMaxLength(5) = 2
        .CaracteresValidos(5) = "1234567890.,"
        .ColWidth(6) = 1000
        .TextMatrix(0, 6) = Space(5) + "Sexo"
        .ColType(6) = cadena
        .ColMaxLength(6) = 1
        .CaracteresValidos(6) = "FMfm"
        .ColWidth(7) = 400
        .TextMatrix(0, 7) = Space(0) + "Doc Identidad"
        .ColType(7) = cadena
        .ColMaxLength(7) = 2
        .CaracteresValidos(7) = "0123456789"
        .ColWidth(8) = 800
        .TextMatrix(0, 8) = Space(4) + "Nro."
        .ColType(8) = cadena
        .ColMaxLength(8) = 8
        .CaracteresValidos(8) = "0123456789"
        .ColWidth(9) = 1000
        .TextMatrix(0, 9) = Space(1) + "Parentesco"
        .ColType(9) = cadena
        .ColMaxLength(9) = 10
        .CaracteresValidos(9) = "ab cd" & Chr(13) & "efghijklmnopqrstuvwxyz" & UCase("abcdefghijklmnopqrstuvwxyz") & ""
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = Space(0) + "CodParent"
        .ColType(10) = cadena
        .ColMaxLength(10) = 10
        .CaracteresValidos(10) = "0123456789"
        .ColWidth(11) = 330
        .TextMatrix(0, 11) = "Activo"
        .ColType(11) = cadena
    End With
End Sub

Private Sub ConfiguraGrillaRet()
    With flxret
        .Clear
        .Refresh
        .Rows = 2
        .Cols = 5
        .RowHeight(1) = 315
        .ColWidth(0) = 400
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColType(0) = cadena
        .ColMaxLength(0) = 15
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(1) = 4500
        .TextMatrix(0, 1) = Space(41) + "Nombre"
        .ColType(1) = cadena
        .ColMaxLength(1) = 64
        .CaracteresValidos(1) = "ab cd.()" & Chr(13) & "efghijklmnopqrstuvwxyzñ" & UCase("abcdefghijklmnopqrstuvwxyzñ") & ""
        .ColWidth(2) = 500
        .TextMatrix(0, 2) = Space(1) + "Mon."
        .ColType(2) = cadena
        .ColMaxLength(2) = 10
        .CaracteresValidos(2) = "NE"
        .ColWidth(3) = 600
        .TextMatrix(0, 3) = Space(5) + "%"
        .ColType(3) = cadena
        .ColMaxLength(3) = 15
        .CaracteresValidos(3) = "1234567890"
        .ColWidth(4) = 1600
        .TextMatrix(0, 4) = Space(10) + "Monto"
        .ColType(4) = cadena
        .ColMaxLength(4) = 10
        .CaracteresValidos(4) = "1234567890."
    End With
End Sub

Private Function GenNumContrato(CodEmp As String) As String
    Dim SQL As String
    Dim rscont As MYSQL_RS
    SQL = " Select MAX(RIGHT(CODIGO,2)) as maximo from contrato where" & _
          " codemp = '" & CodEmp & "'"
    Set rscont = oConexion.EjecutaSelectRS(SQL)
    If Not IsNull(rscont.Fields("maximo")) Then
        SQL = " UPDATE contrato SET ESTADO='" & CANCELADO & "' where" & _
              " codemp = '" & CodEmp & "' AND CODIGO='CN" & rscont.Fields("MAXIMO") & "'"
        oConexionMYSQL.Execute SQL
        GenNumContrato = "CN" & Right("00" & Trim(str(val(rscont.Fields("maximo") + 1))), 2)
    End If
    If IsNull(rscont.Fields("maximo")) Then
        GenNumContrato = "CN01"
    End If
    Set rscont = Nothing
End Function

Private Sub LimpiarDatos()
    sbasico = 0
    txtEmpleado = "           "
    txtNombre1 = Empty
    txtNombre2 = Empty
    txtApePat = Empty
    txtApeMat = Empty
    meFecNac = "  /  /    "
    txtEdad = "  "
    txtTipoDoc = Empty
    txtDireccion = Empty
    txtCarnetExt = Empty
    txtDpto = Empty
    txtPasaporte = Empty
    txtBrevete = Empty
    txtFonoFijo = Empty
    txtFonoMov = Empty
    txtDireccion = Empty
    txtEstatura = "0.00"
    txtCalzado = "00.0"
    txtPeso = "000.00"
    txtApoderado = Empty
    txtDirApo = Empty
    txtFijoApo = Empty
    txtdivc = Empty
    txtmail = Empty
    txtmailper = Empty
    txtNumDoc = Empty
    txtNumHijos = " "
    txtcargo = Empty
'    txtGrado = Empty
'    txtTitulo = Empty
'    txtObs = Empty
    txtDistrito = Empty
    txtAfp = Empty
    txtBanco = Empty
    txtNumAfp = Empty
    lblsbasico = "0.00"
    txtMontoBono = "0.00"
    txtAsigFam = "0.00"
    txtContrato = Empty
    txtTipoCont = Empty
    txtRuc = "           "
    txtNumCtaMe = Empty
    txtNumCtaMn = Empty
    txtCTSCta = Empty
    txtcencos = Empty
    dtpCese.Value = Date
    dtpNacimiento.Value = Date
    dtpIngreso.Value = Date
    dtpFin.Value = Date
    dtpFecIns.Value = Date
    dtfechacese.Value = Date
    dtpCese.Value = Empty
    dtpNacimiento.Value = Empty
    dtpIngreso.Value = Empty
    dtpFin.Value = Empty
    dtfechacese.Value = Empty
    optComiAFP(1).Value = True
    TiposContratos cboContratos
    Horario cboHorLab
    SedesTrabajo cboEstTrabajo
    Categoria cboCategoria
    moneda cboMonBono
    moneda cboMonMovi
    moneda cboMonSueldo
    TiposCese Cbotipocese
    chkBono.Value = False
    chkver.Value = False
    lblDistrito = Empty
    lblEmpleado = Empty
    lblTipoDoc = Empty
    lblcargo = Empty
    lblAfp = Empty
    lblBanco = Empty
    lblTipoCont = Empty
    lblPrimero = Empty
    lblCuenta = Empty
    lblTotal = Empty
    lblSituacEmp = Empty
    lblcencos = Empty
    lbldivc = Empty
    lblnumsvl = ""
    lblnumsctr = ""
    imgFoto.Picture = LoadPicture("")
    strFoto = ""
    txtHCMEmpleado = Empty
    
End Sub

Public Sub BloqueoControles(valor As Boolean)
    txtEmpleado.Enabled = Not valor
    txtNombre1.Locked = valor
    txtNombre2.Locked = valor
    txtApePat.Locked = valor
    txtApeMat.Locked = valor
    txtNumDoc.Locked = valor
    txtDistrito.Locked = valor
    txtDpto.Locked = valor
    txtTipoDoc.Locked = valor
    txtDireccion.Locked = valor
    txtCarnetExt.Locked = valor
    txtPasaporte.Locked = valor
    txtmail.Locked = valor
    txtmailper.Locked = valor
    txtAfp.Locked = valor
    txtNumCtaMe.Locked = valor
    txtNumCtaMn.Locked = valor
    txtNumAfp.Locked = valor
    txtBanco.Locked = valor
    txtcargo.Locked = valor
    'txtTitulo.Locked = valor
    'txtGrado.Locked = valor
    txtDistrito.Locked = valor
    txtAsigFam.Locked = valor
    txtContrato.Locked = valor
    txtdivc.Locked = valor
    txtCTSCta.Locked = valor
    txtBancoCTS.Locked = valor
    cboTipCtaMe.Locked = valor
    cboTipCtaMn.Locked = valor
    CboMonCTS.Locked = valor
    cboNacion.Locked = valor
    Cbotipocese.Locked = valor
    'txtObs.Locked = valor
    chkBono.Locked = valor
    chkAsigFam.Locked = valor
    chkSctr.Locked = valor
    chkSvl.Locked = valor
    chkJubil.Locked = valor
    chkver.Locked = valor
    txtBrevete.Locked = valor
    txtFonoFijo.Locked = valor
    txtFonoMov.Locked = valor
    txtApoderado.Locked = valor
    txtDirApo.Locked = valor
    txtFijoApo.Locked = valor
    txtMovilApo.Locked = valor
    cboBreveteCat.Locked = valor
    cboTMame.Locked = valor
    cboGSanguineo.Locked = valor
    txtEstatura.Enabled = Not valor
    txtPeso.Enabled = Not valor
    txtCalzado.Enabled = Not valor
    txtRuc.Enabled = Not valor
    cboGenero.Locked = valor
    cboEstCivil.Locked = valor
    txtNumHijos.Enabled = Not valor
    cboPersonal.Locked = valor
    cboModal.Locked = valor
    cboCategoria.Locked = valor
    cboTipo.Locked = valor
    txtEdad.Enabled = Not valor
    txtBusqueda.Locked = Not valor
    txtcencos.Locked = valor
    cboBusqueda.Locked = Not valor
    dtpNacimiento.Enabled = Not valor
    dtpIngreso.Enabled = Not valor
    dtpFin.Enabled = Not valor
    dtpInicio.Enabled = Not valor
    dtpFecIns.Enabled = Not valor
    dtfechacese.Enabled = Not valor
    cboMonBono.Enabled = Not valor
    cboMonMovi.Enabled = Not valor
    cboMonSueldo.Enabled = Not valor
    cboContratos.Enabled = Not valor
    cboHorLab.Enabled = Not valor
    cboEstTrabajo.Enabled = Not valor
    CboTipIng.Enabled = Not valor
    cboDiv.Enabled = Not valor
    Cbotipocese.Enabled = Not valor
    txtHCMEmpleado.Enabled = Not valor
    txtHCMEmpleado.Locked = False
    'optComiAFP(0).Enabled = Not valor
    'optComiAFP(1).Enabled = Not valor
    If valor = True Then
        txtEmpleado.BackColor = ColorDeshabilitado
        txtNombre1.BackColor = ColorDeshabilitado
        txtNombre2.BackColor = ColorDeshabilitado
        txtApePat.BackColor = ColorDeshabilitado
        txtApeMat.BackColor = ColorDeshabilitado
        txtEdad.BackColor = ColorDeshabilitado
        txtTipoDoc.BackColor = ColorDeshabilitado
        txtCarnetExt.BackColor = ColorDeshabilitado
        txtPasaporte.BackColor = ColorDeshabilitado
        txtDireccion.BackColor = ColorDeshabilitado
        txtDpto.BackColor = ColorDeshabilitado
        txtFonoFijo.BackColor = ColorDeshabilitado
        txtFonoMov.BackColor = ColorDeshabilitado
        txtmail.BackColor = ColorDeshabilitado
        txtmailper.BackColor = ColorDeshabilitado
        txtNumDoc.BackColor = ColorDeshabilitado
        txtNumHijos.BackColor = ColorDeshabilitado
        'txtObs.BackColor = ColorDeshabilitado
        txtAfp.BackColor = ColorDeshabilitado
        txtNumAfp.BackColor = ColorDeshabilitado
        txtBanco.BackColor = ColorDeshabilitado
        txtMontoBono.BackColor = ColorDeshabilitado
        txtDistrito.BackColor = ColorDeshabilitado
        txtEstatura.BackColor = ColorDeshabilitado
        txtPeso.BackColor = ColorDeshabilitado
        txtCalzado.BackColor = ColorDeshabilitado
        txtBrevete.BackColor = ColorDeshabilitado
        txtcargo.BackColor = ColorDeshabilitado
        'txtTitulo.BackColor = ColorDeshabilitado
        'txtGrado.BackColor = ColorDeshabilitado
        txtApoderado.BackColor = ColorDeshabilitado
        txtDirApo.BackColor = ColorDeshabilitado
        txtdivc.BackColor = ColorDeshabilitado
        txtFijoApo.BackColor = ColorDeshabilitado
        txtMovilApo.BackColor = ColorDeshabilitado
        txtAsigFam.BackColor = ColorDeshabilitado
        txtNumCtaMe.BackColor = ColorDeshabilitado
        txtNumCtaMn.BackColor = ColorDeshabilitado
        cboTipCtaMe.BackColor = ColorDeshabilitado
        cboTipCtaMn.BackColor = ColorDeshabilitado
        CboMonCTS.BackColor = ColorDeshabilitado
        cboNacion.BackColor = ColorDeshabilitado
        cboGenero.BackColor = ColorDeshabilitado
        cboEstCivil.BackColor = ColorDeshabilitado
        cboPersonal.BackColor = ColorDeshabilitado
        cboModal.BackColor = ColorDeshabilitado
        cboCategoria.BackColor = ColorDeshabilitado
        cboTipo.BackColor = ColorDeshabilitado
        txtContrato.BackColor = ColorDeshabilitado
        txtRuc.BackColor = ColorDeshabilitado
        txtCTSCta.BackColor = ColorDeshabilitado
        txtBancoCTS.BackColor = ColorDeshabilitado
        cboGSanguineo.BackColor = ColorDeshabilitado
        cboTMame.BackColor = ColorDeshabilitado
        cboBreveteCat.BackColor = ColorDeshabilitado
        flxDependientes.BackColor = ColorDeshabilitado
        flxret.BackColor = ColorDeshabilitado
        flxsueldos.BackColor = ColorDeshabilitado
        msfseg.BackColor = ColorDeshabilitado
        txtBusqueda.BackColor = ColorHabilitado
        cboBusqueda.BackColor = ColorHabilitado
        Cbotipocese.BackColor = ColorDeshabilitado
        txtcencos.BackColor = ColorDeshabilitado
        txtHCMEmpleado.BackColor = ColorDeshabilitado
    Else
        txtEmpleado.BackColor = ColorHabilitado
        txtNombre1.BackColor = ColorHabilitado
        txtNombre2.BackColor = ColorHabilitado
        txtApePat.BackColor = ColorHabilitado
        txtApeMat.BackColor = ColorHabilitado
        txtEdad.BackColor = ColorHabilitado
        txtTipoDoc.BackColor = ColorHabilitado
        txtPasaporte.BackColor = ColorHabilitado
        txtCarnetExt.BackColor = ColorHabilitado
        txtDireccion.BackColor = ColorHabilitado
        txtDistrito.BackColor = ColorHabilitado
        txtDpto.BackColor = ColorHabilitado
        txtdivc.BackColor = ColorHabilitado
        txtFonoFijo.BackColor = ColorHabilitado
        txtFonoMov.BackColor = ColorHabilitado
        txtmail.BackColor = ColorHabilitado
        txtmailper.BackColor = ColorHabilitado
        txtNumDoc.BackColor = ColorHabilitado
        txtDistrito.BackColor = ColorHabilitado
        txtcargo.BackColor = ColorHabilitado
'        txtGrado.BackColor = ColorHabilitado
'        txtTitulo.BackColor = ColorHabilitado
        txtEstatura.BackColor = ColorHabilitado
        txtPeso.BackColor = ColorHabilitado
        txtCalzado.BackColor = ColorHabilitado
        txtBrevete.BackColor = ColorHabilitado
        cboBreveteCat.BackColor = ColorHabilitado
        cboTMame.BackColor = ColorHabilitado
        cboGSanguineo.BackColor = ColorHabilitado
        cboGenero.BackColor = ColorHabilitado
        cboEstCivil.BackColor = ColorHabilitado
        txtApoderado.BackColor = ColorHabilitado
        txtDirApo.BackColor = ColorHabilitado
        txtFijoApo.BackColor = ColorHabilitado
        txtMovilApo.BackColor = ColorHabilitado
        txtNumHijos.BackColor = ColorHabilitado
'        txtObs.BackColor = ColorHabilitado
        txtAfp.BackColor = ColorHabilitado
        txtNumAfp.BackColor = ColorHabilitado
        txtBanco.BackColor = ColorHabilitado
        txtRuc.BackColor = ColorHabilitado
        txtAsigFam.BackColor = ColorHabilitado
        txtNumCtaMe.BackColor = ColorHabilitado
        txtNumCtaMn.BackColor = ColorHabilitado
        txtCTSCta.BackColor = ColorHabilitado
        txtBancoCTS.BackColor = ColorHabilitado
        flxDependientes.BackColor = ColorHabilitado
        flxret.BackColor = ColorHabilitado
        flxsueldos.BackColor = ColorHabilitado
        msfseg.BackColor = ColorHabilitado
        CboMonCTS.BackColor = ColorHabilitado
        cboNacion.BackColor = ColorHabilitado
        cboPersonal.BackColor = ColorHabilitado
        cboModal.BackColor = ColorHabilitado
        cboCategoria.BackColor = ColorHabilitado
        cboTipo.BackColor = ColorHabilitado
        Cbotipocese.BackColor = ColorHabilitado
        cboTipCtaMe.BackColor = ColorHabilitado
        cboTipCtaMn.BackColor = ColorHabilitado
        txtContrato.BackColor = ColorHabilitado
        txtBusqueda.BackColor = ColorDeshabilitado
        cboBusqueda.BackColor = ColorDeshabilitado
        txtcencos.BackColor = ColorHabilitado
        txtHCMEmpleado.BackColor = ColorHabilitado
    End If
End Sub

Private Sub ComboBusq(cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem "Búsqueda por..."
    cbo.List(0, 1) = "0"
    cbo.AddItem "Apellido Paterno"
    cbo.List(1, 1) = "APEPAT"
    cbo.AddItem "Apellido Materno"
    cbo.List(2, 1) = "APEMAT"
    cbo.AddItem "Nombre 1"
    cbo.List(3, 1) = "NOMBRE1"
    cbo.AddItem "Nombre 2"
    cbo.List(4, 1) = "NOMBRE2"
    cbo.AddItem "Posición"
    cbo.List(5, 1) = "POS"
    cbo.AddItem "Dni"
    cbo.List(6, 1) = "NUMDOCIDE"
    cbo.AddItem "HCM"
    cbo.List(7, 1) = "CODIGOHCM"
    cbo.ListIndex = 1
End Sub

Private Sub Consulta()
    Dim SQL As String
    SQL = "Select e.*,a.* from empleado as e left join empleado_apoderado as a on e.codigo = a.codemp where e.codigo <> '00000000000' order by e.apepat,e.apemat,e.nombre1,e.nombre2"
    Set rsgral = oConexion.EjecutaSelectRS(SQL)
    rsgral.MoveFirst
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
             LimpiarDatos
             lblModo = "Acción"
             BloqueoControles True
             genero cboGenero
             Estado cboEstCivil
             TipoInstEduc cboTipInst
             EstudPrin cboEstPrin
             FormSup cboFormSup
             NomInst cboNomInst
             DameCarrera cboCarrera
             SituacionEmp cboSituacion
             TipCta cboTipCtaMe
             TipCta cboTipCtaMn
             Tipo cboTipo
             Personal cboPersonal
             Gsangre cboGSanguineo
             Nacionalidad cboNacion
             Categoria cboCategoria
             Tallas cboTMame
             TipoDoc cboTipoDocPar
             TipoParen cboParenContacto
             TBrevetes cboBreveteCat
             ComboBusq cboBusqueda
             TiposContratos cboContratos
             Horario cboHorLab
             SedesTrabajo cboEstTrabajo
             moneda cboMonSueldo
             moneda cboMonBono
             moneda cboMonMovi
             monedaCCI CboMonCTS
             Divisiones cboDiv
             ConfigurarBotones cfgCancelar
             BtnNuevo.Enabled = True
             btnContrato.Enabled = False
             ConfiguraGrilla
             txtEmpleado.Enabled = True
             txtEmpleado.BackColor = ColorHabilitado
             SSTab1.Tab = 0
             generocod = False
             txtEmpleado.SelStart = 0
             Consulta
             CargarDatos rsgral
             If rsgral.RecordCount > 0 Then
                ConfigBtnsBusq rsgral.AbsolutePosition, rsgral.RecordCount
             End If
             If BtnEliminar.tag <> "" Then BtnEliminar.Enabled = BtnEliminar.tag Else: BtnEliminar.Enabled = False
             btnContrato.Enabled = True
             btnContrato.Enabled = False
             btnRenovar.Enabled = False
             btnHabilitar.Enabled = False
             BtnSalir.Enabled = True
             txtcargo.Locked = True
             Exit Sub
        Case ModoForm.modNuevo
             LimpiarDatos
             txtBusqueda = Empty
             lblModo = "Nuevo"
             BloqueoControles False
             genero cboGenero
             Estado cboEstCivil
             SituacionEmp cboSituacion
             Categoria cboCategoria
             ComboBusq cboBusqueda
             TipCta cboTipCtaMe
             TipCta cboTipCtaMn
             TiposContratos cboContratos
             Horario cboHorLab
             SedesTrabajo cboEstTrabajo
             moneda cboMonSueldo
             moneda cboMonBono
             moneda cboMonMovi
             moneda CboMonCTS
             Gsangre cboGSanguineo
             Nacionalidad cboNacion
             Tallas cboTMame
             TipoDoc cboTipoDocPar
             TipoParen cboParenContacto
             TBrevetes cboBreveteCat
             Divisiones cboDiv
             ConfigurarBotones cfgNuevo
             ConfiguraGrilla
             ConfiguraGrillaSueldos
             ConfiguraGrillaSeguros
             TipCta cboTipCtaMe
             TipCta cboTipCtaMn
             TipodeCampo = cadena
             intNumHijos = 0
             txtEmpleado.SetFocus
             txtContrato.Locked = True
             SSTab1.Tab = 0
             frmRegEmpleado.Caption = "Mantenimiento de Personal - [ Registrar Nuevo Empleado ]"
             For I = 1 To msfseg.Rows - 1
                msfseg.row = I: msfseg.Col = 2
                msfseg.CellBackColor = ColorDeshabilitado
             Next
             txtcargo.Locked = True
             Call ConfiguraGrillaRet
             Exit Sub
             
             
        Case ModoForm.modConsulta
             lblModo = "Consulta"
             BloqueoControles True
             ConfigurarBotones cfgGrabar
             txtEmpleado.Enabled = True
             txtEmpleado.BackColor = ColorHabilitado
             generocod = False
             BtnCancelar.Enabled = False
             If BtnEliminar.tag <> "" Then BtnEliminar.Enabled = BtnEliminar.tag Else: BtnEliminar.Enabled = False
             btnCancelCont.Enabled = False
'             cboSituacion.Enabled = False
'             cboSituacion.Locked = True
'             cboSituacion.BackColor = ColorDeshabilitado
             txtcargo.Locked = True
             
             Exit Sub
        Case ModoForm.modEditar
             lblModo = "Modificar"
             BloqueoControles False
             ConfigurarBotones cfgModificar
             If txtContrato.tag = PENDIENTE Then btnContrato.Enabled = False
             If txtContrato.tag = "" Then btnContrato.Enabled = True
             txtEmpleado.Enabled = False
             txtEmpleado.BackColor = ColorDeshabilitado
             txtRuc.Enabled = False
             txtRuc.BackColor = ColorDeshabilitado
             generocod = False
             txtBusqueda = Empty
             BotonesContrato
             txtContrato.Locked = True
             txtContrato.BackColor = ColorDeshabilitado
             If chkAsigFam.Value = True Then
                txtAsigFam.Locked = False
                txtAsigFam.BackColor = ColorHabilitado
             Else
                txtAsigFam.Locked = True
                txtAsigFam.BackColor = ColorDeshabilitado
             End If
             If chkBono.Value = True Then
                txtMontoBono.Locked = False
                txtMontoBono.BackColor = ColorHabilitado
                cboMonBono.Locked = False
                cboMonBono.BackColor = ColorHabilitado
             Else
                txtMontoBono.Locked = True
                txtMontoBono.BackColor = ColorDeshabilitado
                cboMonBono.Locked = True
                cboMonBono.BackColor = ColorDeshabilitado
             End If
             
            txtMontoMov.Enabled = True
            txtMontoMov.Locked = False
            txtMontoMov.BackColor = ColorHabilitado
            cboMonMovi.Enabled = True
            cboMonMovi.Locked = False
            cboMonMovi.BackColor = ColorHabilitado
             
             For I = 1 To msfseg.Rows - 1
                msfseg.row = I: msfseg.Col = 2
                msfseg.CellBackColor = ColorDeshabilitado
             Next
             If cboTipo.List(cboTipo.ListIndex, 1) = 4 Then
                cboSituacion.Enabled = True
                cboSituacion.Locked = False
                cboSituacion.BackColor = ColorHabilitado
            Else
                cboSituacion.Enabled = False
                cboSituacion.Locked = True
                cboSituacion.BackColor = ColorDeshabilitado
             End If
             txtNombre1.SetFocus
             txtcargo.Locked = True
             Exit Sub
    End Select
End Sub

Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Dim RES As Integer
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            BtnNuevo.Enabled = False
            BtnModificar.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = True
            BtnCancelar.Enabled = True
            btnReporte.Enabled = False
            BtnSalir.Enabled = False
            btnContrato.Enabled = False
            btnRenovar.Enabled = False
            btnPrimero.Enabled = False
            btnPrevio.Enabled = False
            btnUltimo.Enabled = False
            btnSgt.Enabled = False
            btnContrato.Enabled = False
            btnRenovar.Enabled = False
            btnHabilitar.Enabled = False
            Publimensaje = "modificar"
            strFoto = ""
            Exit Sub
        Case ConfigBotones.cfgModificar
            BtnNuevo.Enabled = False
            BtnModificar.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = True
            BtnCancelar.Enabled = True
            btnContrato.Enabled = False
            btnRenovar.Enabled = False
            btnPrimero.Enabled = False
            btnPrevio.Enabled = False
            btnUltimo.Enabled = False
            btnSgt.Enabled = False
            btnContrato.Enabled = False
            btnRenovar.Enabled = False
            btnHabilitar.Enabled = False
            Publimensaje = "modificar"
            Exit Sub
        Case ConfigBotones.cfgEliminar
            BtnNuevo.Enabled = True
            BtnModificar.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = False
            btnReporte.Enabled = False
            BtnCancelar.Enabled = False
            BtnSalir.Enabled = False
            btnRenovar.Enabled = False
            btnContrato.Enabled = False
            btnRenovar.Enabled = False
            btnHabilitar.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgGrabar
            BtnNuevo.Enabled = True
            btnGrabar.Enabled = False
            BtnCancelar.Enabled = True
            BtnModificar.Enabled = True
            btnContrato.Enabled = True
            btnContrato.Enabled = False
            btnRenovar.Enabled = False
            btnHabilitar.Enabled = False
            Publimensaje = ""
        Case ConfigBotones.cfgCancelar
            Select Case lblModo.Caption
                Case "Nuevo"
                    RES = MsgBox("Desea Cancelar el Registro de Datos del empleado?", vbQuestion + vbYesNo, gsNomSW)
                    If RES = 6 Then
                        Publimensaje = ""
                        ModoFormulario modAccion
                        If txtContrato.tag = PENDIENTE Then btnContrato.Enabled = False Else: btnContrato.Enabled = True
                    End If
                Case "Consulta"
                     Publimensaje = ""
                     ModoFormulario modAccion
                     BtnNuevo.Enabled = True
                     btnGrabar.Enabled = False
                     btnReporte.Enabled = False
                     BtnCancelar.Enabled = False
                Case "Modificar"
                    Consulta
                    MoverPosRs Trim(txtEmpleado)
                    CargarDatos rsgral
                    If rsgral.State = MY_RS_OPEN Then
                        If rsgral.RecordCount > 0 Then
                           ConfigBtnsBusq rsgral.AbsolutePosition, rsgral.RecordCount
                        End If
                    End If
                    btnContrato.Enabled = False
                    btnCancelCont.Enabled = False
                    btnHabilitar.Enabled = False
                    btnRenovar.Enabled = False
                    If BtnEliminar.tag <> "" Then BtnEliminar.Enabled = BtnEliminar.tag Else: BtnEliminar.Enabled = False
            End Select
    End Select
End Sub

Private Sub txtNombre1_GotFocus()
    mark1 txtNombre1
End Sub

Private Sub txtNombre2_GotFocus()
    mark1 txtNombre2
End Sub

Private Sub txtNombre1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtNombre2.SetFocus
        Exit Sub
    End If
    If KeyCode <> vbKeySpace Then
        If KeyCode <> vbKeyDelete Then
            If KeyCode <> 8 Then
                If IsNumeric("0" & Chr(KeyCode)) Or (KeyCode >= 96 And KeyCode <= 109) Then
                    Beep
                    KeyCode = 0
                End If
            End If
        End If
    End If
  End If
    
End Sub

Private Sub txtNombre2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtApePat.SetFocus
        Exit Sub
    End If
    If KeyCode <> vbKeySpace Then
        If KeyCode <> vbKeyDelete Then
            If KeyCode <> 8 Then
                If IsNumeric("0" & Chr(KeyCode)) Or (KeyCode >= 96 And KeyCode <= 109) Then
                    Beep
                    KeyCode = 0
                End If
            End If
        End If
    End If
  End If
    
End Sub

Private Sub txtNumCtaMe_GotFocus()
    mark1 txtNumCtaMe
End Sub

Private Sub txtNumCtaMe_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtAfp.SetFocus
    End If
End Sub

Private Sub txtNumCtaMn_GotFocus()
    mark1 txtNumCtaMn
End Sub

Private Sub txtNumCtaMn_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cboTipCtaMe.SetFocus
    End If
End Sub

Private Sub txtNumDoc_GotFocus()
    mark1 txtNumDoc
End Sub

Private Sub txtNumDoc_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtDireccion.SetFocus
        Exit Sub
    End If
    If KeyCode <> vbKeySpace Then
        If KeyCode <> vbKeyDelete Then
            If KeyCode <> 8 Then
                If Not (KeyCode >= 96 And KeyCode <= 109) Then
                    If Not IsNumeric("0" & Chr(KeyCode)) Then
                        Beep
                        KeyCode = 0
                    End If
                End If
            End If
        End If
    End If
  End If
End Sub

Private Sub txtNumHijos_GotFocus()
    mark1 txtNumHijos
End Sub

Private Sub txtNumHijos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNumHijos = Right("0" & txtNumHijos, 1)
        txtApoderado.SetFocus
    End If
End Sub

'Private Sub TxtObs_GotFocus()
'    mark1 txtObs
'End Sub

Private Sub txtPasaporte_GotFocus()
    mark1 txtPasaporte
End Sub

Private Sub txtPasaporte_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
  If KeyCode = 13 Then
       cboBreveteCat.SetFocus
    End If
 End If
End Sub

Private Sub txtPeso_GotFocus()
    mark1 txtPeso
End Sub

Private Sub txtPeso_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 17 And Shift <> 2 Then
    If KeyCode = 13 Then
        txtCalzado.SetFocus
    End If
 End If
End Sub

Private Sub TxtRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEmpleado = txtRuc
        txtNombre1.SetFocus
    End If
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
Dim NewValue As Long
Dim Lstep As Single
On Error Resume Next
    With flxDependientes
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then
            Lstep = 10
        End If
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then
                NewValue = 1
            End If
        Else
            NewValue = .TopRow + Lstep
            If NewValue > .Rows - 1 Then
                NewValue = .Rows - 1
            End If
        End If
        .TopRow = NewValue
    End With
End Sub

Private Sub txtTipoDoc_Change()
    lblTipoDoc = DescripcionesdeCodigos("TIPODOCIDE", Right("00" & Trim(txtTipoDoc), 2))
End Sub

Private Sub txtTipoDoc_GotFocus()
    mark1 txtTipoDoc
End Sub

Private Sub txtTipoDoc_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 And txtTipoDoc.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1000
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Tipos de Doc. de Identidad"
            .pForm = FORM_REGEMP
            .pCaso = LABEL_DOC_IDEM
            .Show
        End With
    End If
End Sub

Private Sub txtTipoDoc_LostFocus()
 txtTipoDoc = Right("00" & Trim(txtTipoDoc), 2)
End Sub

'Private Sub txtTitulo_Change()
'    If txtTitulo = Empty Then
'        lblTitulo = Empty
'    End If
'End Sub

'Private Sub txtTitulo_GotFocus()
'    mark1 txtTitulo
'End Sub
'
'Private Sub txtTitulo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
' If KeyCode <> 17 And Shift <> 2 Then
'   If KeyCode = 13 Then
'        txtTitulo = Right("00" & txtTitulo, 2)
'        lblTitulo = DescripcionesdeCodigos("TITULOS", Trim(txtTitulo))
'        txtGrado.SetFocus
'        Exit Sub
'    End If
'    If KeyCode = vbKeyF1 And txtTitulo.BackColor = ColorHabilitado Then
'         With oConsulta
'            .pCols = 2
'            .pCol = 0: .pAnchoCol = 1500
'            .pCol = 1: .pAnchoCol = 4000
'            .pTitulo = "Grados"
'            .pForm = FORM_REGEMP
'            .pCaso = LABEL_TITULOS
'            .Show
'        End With
'    End If
' End If
'End Sub

Function DameTipoSuel(ByVal TextDoc As String) As String
   Select Case TextDoc
        Case "Fijo"
            DameTipoSuel = "F"
        Case "Variable"
            DameTipoSuel = "V"
        Case Else
            DameTipoSuel = "F"
    End Select

End Function


Sub ConfiguraGrillaFormEdu()
    With flxformEduEmp
        .Clear
        .Cols = 11
        .Rows = 2
        .FixedCols = 1
        .ColWidth(0) = 1800
        .TextMatrix(0, 0) = "NomEmpleado"
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = "Formación Educ"
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = "Estudió en Perú?"
        .ColWidth(3) = 1200
        .TextMatrix(0, 3) = "Regimen Educativo"
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "Tipo Institución"
        .ColWidth(5) = 3400
        .TextMatrix(0, 5) = "Nombre Institución"
        .ColWidth(6) = 2500
        .TextMatrix(0, 6) = "Carrera"
        .ColWidth(7) = 500
        .TextMatrix(0, 7) = "Año Egreso"
        .ColWidth(8) = 500
        .TextMatrix(0, 8) = "Visible"
        .ColWidth(9) = 5
        .TextMatrix(0, 9) = "Id"
        .ColWidth(10) = 500
        .TextMatrix(0, 10) = "Eliminar"
    End With
End Sub

Private Sub CargarFormacEduc(codempleado As String)
    Dim SQL As String
    Dim I As Integer
    Dim RQ As MYSQL_RS
    
    SQL = "Select v.id as Id,(select concat(e.nombre1,' ',e.apepat,' ',e.apemat) from empleado as e where e.codigo=v.codemp) as nombreemp," & _
          "(select p.descrip from pl_siteduc as p where p.codsit=v.codformsup) as nomformsup," & _
          "if(v.flgEstPeru='S','PERU','EXT') as nomEstPeru," & _
          "if(v.flgRegInst='P','PRIVADA','NACIONAL') as nomRegInst," & _
          "if(v.flgTipoInst='O','OTRO',if(v.flgTipoInst='C','COLEGIO',if(v.flgTipoInst='U','UNIVERSIDAD','INSTITUTO')))as nomTipoInst," & _
          "(select u.descrip from pl_universia as u where u.coduni=v.coduni) as nomuni," & _
          "ifnull((select f.descrip from pl_carrprof as f where f.codcarr=v.codcarr),'') as nomcarrprof," & _
          "v.anhoegr,v.Visible From pl_empformedu as v where v.codemp = '" & codempleado & "' order by v.anhoegr"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    ConfiguraGrillaFormEdu
    I = 1
    Do While Not RQ.EOF
        With flxformEduEmp
           .TextMatrix(I, 0) = RQ.Fields("nombreemp")
           .TextMatrix(I, 1) = RQ.Fields("nomformsup")
           .TextMatrix(I, 2) = RQ.Fields("nomEstPeru")
           .TextMatrix(I, 3) = RQ.Fields("nomRegInst")
           .TextMatrix(I, 4) = RQ.Fields("nomTipoInst")
           .TextMatrix(I, 5) = RQ.Fields("nomuni")
           .TextMatrix(I, 6) = RQ.Fields("nomcarrprof")
           .TextMatrix(I, 7) = RQ.Fields("anhoegr")
           .TextMatrix(I, 8) = RQ.Fields("visible")
           .TextMatrix(I, 9) = RQ.Fields("Id")
           .row = I
           .Col = 10
           .CellForeColor = vbRed
           .CellFontName = "Wingdings"
           .CellFontSize = 14
           .TextMatrix(I, 10) = strUnChecked
                  
           .Rows = .Rows + 1
           I = I + 1
            RQ.MoveNext
        End With
    Loop
   
    Set RQ = Nothing
End Sub


Private Sub flxformEduEmp_Click()
    Dim SCol As Integer
    With flxformEduEmp
        SCol = .Col
        If SCol = 10 Then
            If .TextMatrix(.row, 10) = strChecked Then
                .TextMatrix(.row, 10) = strUnChecked
            Else
                .TextMatrix(.row, 10) = strChecked
            End If
        End If
    End With
End Sub


Private Function ValidaFormacionEducativa(codigo As String) As Boolean
    Dim SQL As String
    Dim rscod As MYSQL_RS
    Dim RES As Integer
    ValidaFormacionEducativa = False
            
    SQL = "Select * from pl_empformedu where codemp = '" & Trim(txtEmpleado) & "' and codformsup = '" & cboFormSup.List(cboFormSup.ListIndex, 1) & "' and flgestperu = '" & OptPeruflg & "' " & _
          "and flgRegInst = '" & OptRegEduflg & "' and flgTipoInst = '" & cboTipInst.List(cboTipInst.ListIndex, 1) & "' and coduni = '" & cboNomInst.List(cboNomInst.ListIndex, 1) & "' " & _
          "and codcarr = '" & cboCarrera.List(cboCarrera.ListIndex, 1) & "' and  anhoegr= '" & txtAnhoEgr & "'  "
    Set rscod = oConexion.EjecutaSelectRS(SQL)
    
    If rscod.RecordCount >= 1 Then
        ValidaFormacionEducativa = True
    End If
    
End Function


Sub GenerarfolioPract(folioRef As String, folioPago As String, flagtipo As String)
    Dim SQL As String
    Dim v As String, lib As String, glo As String, Serdoc As String, Numdoc As String, fec As String, AnoMes As String, det As String
    Dim tc As Double, td As String, Div As String, mon As String, aux As String, caux As String, cencos As String, cenco As String, Cta As String, cto As String, dh As String
    Dim sol As Double, dol As Double, colv As String, correl As String, Clasf As String, At As String
    Dim Base As Double, Igv As Double, total As Double, TotalDolaresC As Double, TotalSolesC As Double
    Dim Xsol As Double, Xdol As Double, Tot As Integer, CantPB As Integer
    Dim AcumSol As Double, AcumDol As Double
    Dim RubRendicion As String
    Dim mensaje As String
    Dim FlagIGV As String
    Dim RSVOU As ADODB.Recordset
    Set RSVOU = New ADODB.Recordset
    Dim SQLVOU As String
    Dim flagtiempo As String
    Dim AcumDivCS_S As Double
    Dim AcumDivIP_S As Double
    Dim AcumDivFI_S As Double
    Dim AcumDivIT_S As Double
    Dim AcumDivCS_D As Double
    Dim AcumDivIP_D As Double
    Dim AcumDivFI_D As Double
    Dim AcumDivIT_D As Double
    
    AcumDivCS_S = 0
    AcumDivIP_S = 0
    AcumDivFI_S = 0
    AcumDivIT_S = 0
    AcumDivCS_D = 0
    AcumDivIP_D = 0
    AcumDivFI_D = 0
    AcumDivIT_D = 0
    
                    lib = "05"
                    mon = "N"
                    fec = Date
                    tc = BuscarTipoCambioPorLetra("V", fec) 'dblTipoCmbV
                    AnoMes = strAnoSistema & strMesSistema
                    v = MaxVoucher(AnoMes, lib)
                    glo = "PROVEEDORES"
                    det = "N"
                    correl = "0001"
                    AcumSol = 0
                    AcumDol = 0
                    SQL = "Call cn_Insert_Voucher('" & lib & "','" & v & "','" & glo & "','" & fec & _
                          " ','" & fec & "','V'," & tc & ",'" & mon & "','" & AnoMes & "','" & strUsuarioId & _
                          " ','CUADRADO','','','','','" & det & "','','')"
                    oConexionMYSQL.Execute (SQL)
                    
                    
                    SQLVOU = "Call CONT_Lista_Practicantes_nov('" & flagtipo & "');"
                    Set RSVOU = ADO_LlenaRs(SQLVOU)
 
                    If RSVOU Is Nothing = False Then
 
                    Do While Not RSVOU.EOF
                            td = "9"
                            Div = Trim(RSVOU.Fields("divi"))
                            total = Round(CDbl(RSVOU.Fields("cargos")), 2)
                            If mon = "N" Then
                                Xsol = Round(total, 2)
                                Xdol = Round(total / tc, 2)
                            Else
                                Xdol = Round(total, 2)
                                Xsol = Round(total * tc, 2)
                            End If
                            
                            Base = Round(total / 1.18, 2)
                            Igv = Round(total - Base, 2)
                            If flagtipo = "P" Then
                             Serdoc = "PRACT"
                            Else
                             Serdoc = "REJUD"
                            End If
                             
                            Numdoc = folioRef
                            Cta = "469911"
                            aux = "6"
                            caux = Trim(RSVOU.Fields("codaux"))
                            'CAMBIO A 11 DIGITOS CCLOCAL
                            cenco = "00000000000"
                            cencos = "0000"
                            dh = "H"
                            cto = UCase(DescripAuxiliar("6", Trim(RSVOU.Fields("codaux"))))
                            colv = "00"
                            
                            Ins_Movimiento lib, td, Div, Cta, colv, v, Serdoc, Numdoc, correl, mon, aux, caux, cencos, cenco, "N", cto, fec, dh, Xsol, Xdol, strUsuarioId, "000", ""
                            correl = Right("0000" & Trim(str(val(correl) + 1)), 4)
                            AcumSol = AcumSol + Xsol
                            AcumDol = AcumDol + Xdol
                            
                            If Div = "013100003836" Then
                             AcumDivCS_S = AcumDivCS_S + Xsol
                             AcumDivCS_D = AcumDivCS_D + Xdol
                            End If
                            
                            If Div = "013100003841" Then
                             AcumDivIP_S = AcumDivIP_S + Xsol
                             AcumDivIP_D = AcumDivIP_D + Xdol
                            End If
                            
                            If Div = "013100000150" Then
                             AcumDivFI_S = AcumDivFI_S + Xsol
                             AcumDivFI_D = AcumDivFI_D + Xdol
                            End If
                            
                            If Div = "013100000018" Then
                             AcumDivIT_S = AcumDivIT_S + Xsol
                             AcumDivIT_D = AcumDivIT_D + Xdol
                            End If
    
                    RSVOU.MoveNext
                  Loop
                    Set RSVOU = Nothing
                  End If
                  
              If flagtipo = "P" Then
                  td = "9"
                  Serdoc = ""
                  Numdoc = DameNombreMes(MesSistema)
                  cencos = "0370"
                  Cta = "939024"
                  aux = "0"
                  caux = "00000000000"
                  colv = "00"
                  dh = "D"
                  
                  flagtiempo = InputBox("Ingrese que quincena es  p.e:[P] si es Primera Quincena,[S] si es segunda ", "Voucher Terceros", "P")
                  
                  If flagtiempo = "P" Then
                    cto = "PAGO PRACT - 1RA QUINCENA " & DameNombreMes(MesSistema)
                  Else
                    cto = "PAGO PRACT - 2DA QUINCENA " & DameNombreMes(MesSistema)
                  End If
                            
                  If AcumDivCS_S > 0 Or AcumDivCS_D > 0 Then
                    Div = "013100003836" '0001
                    cenco = "13100003836"
                    Ins_Movimiento lib, td, Div, Cta, colv, v, Serdoc, Numdoc, correl, mon, aux, caux, cencos, cenco, "N", cto, fec, dh, AcumDivCS_S, AcumDivCS_D, strUsuarioId, "V01", ""
                    correl = Right("0000" & Trim(str(val(correl) + 1)), 4)
                  End If
                  
                  If AcumDivIP_S > 0 Or AcumDivIP_D > 0 Then
                    Div = "013100003841" '0002
                    cenco = "38410000006"
                    Ins_Movimiento lib, td, Div, Cta, colv, v, Serdoc, Numdoc, correl, mon, aux, caux, cencos, cenco, "N", cto, fec, dh, AcumDivIP_S, AcumDivIP_D, strUsuarioId, "V01", ""
                    correl = Right("0000" & Trim(str(val(correl) + 1)), 4)
                  End If
                  
                  If AcumDivFI_S > 0 Or AcumDivFI_D > 0 Then
                    Div = "013100000150" '0008
                    cenco = "13100000150"
                    Ins_Movimiento lib, td, Div, Cta, colv, v, Serdoc, Numdoc, correl, mon, aux, caux, cencos, cenco, "N", cto, fec, dh, AcumDivFI_S, AcumDivFI_D, strUsuarioId, "V01", ""
                    correl = Right("0000" & Trim(str(val(correl) + 1)), 4)
                  End If
                  
                  If AcumDivIT_S > 0 Or AcumDivIT_D > 0 Then
                    Div = "013100000018" '0018
                    cenco = "13100000018"
                    Ins_Movimiento lib, td, Div, Cta, colv, v, Serdoc, Numdoc, correl, mon, aux, caux, cencos, cenco, "N", cto, fec, dh, AcumDivIT_S, AcumDivIT_D, strUsuarioId, "V01", ""
                    correl = Right("0000" & Trim(str(val(correl) + 1)), 4)
                  End If
              Else
                  td = "9"
                  Div = "013100003836" '0001
                  Serdoc = "TW"
                  Numdoc = folioPago
                  Cta = "104101"
                  aux = "1"
                  caux = "00000000001"
                  'CAMBIO A 11 DIGITOS CCLOCAL
                  cenco = "00000000000"
                  cencos = "0000"
                  dh = "D"
                  cto = "RETENCION JUDICIAL PIVOT"
                  colv = "00"
                  Ins_Movimiento lib, td, Div, Cta, colv, v, Serdoc, Numdoc, correl, mon, aux, caux, cencos, cenco, "N", cto, fec, dh, AcumSol, AcumDol, strUsuarioId, "000", ""
              End If
                                    
Set Rs = Nothing

MsgBox "Se generó el voucher de Practicantes/Retenciones Judiciales satisfactoriamente." & vbNewLine & "Voucher: " & v & vbNewLine, vbInformation, "NOVCONT"

'fin
End Sub

Public Sub Ins_Movimiento(ilib As String, itdoc As String, idivi As String, icta As String, icolcv As String, ivou As String, iserdoc As String, idocum As String, icorr As String, imoneda As String, iaux As String, icaux As String, iCencos As String, iCenco As String, igen As String, icto As String, ifec As String, idh As String, isol As Double, idol As Double, iuser As String, imor As String, orden As String)
    Dim SQL As String
    Dim XCuenta As String
    Dim XDH As String
    Dim XSoles As Double
    Dim XDolares As Double
    Dim XGENERADA As String
    Dim I As Integer
    Dim RsCargos As ADODB.Recordset
    Set RsCargos = New ADODB.Recordset
    Dim RsPorc As ADODB.Recordset
    Set RsPorc = New ADODB.Recordset
    If TieneGenerada(icta) = True Then   'tiene generada ?
        
        SQL = "Call CONT_Rep_Proc_Genericos('CNMAYORCARGO','" & icta & "','','','','','','','','','' );"
        Set RsCargos = ADO_LlenaRs(SQL)
        
        SQL = "Call CONT_Rep_Proc_Genericos('CNMAYORPORC','" & icta & "','','','','','','','','','' );"
        Set RsPorc = ADO_LlenaRs(SQL)
        
        For I = 0 To 5
            XSoles = 0: XDolares = 0
            If I < 5 Then
                XDH = idh
                XSoles = isol * CDbl((RsPorc(I) / 100))
                XDolares = idol * CDbl((RsPorc(I) / 100))
            Else
                XDH = IIf(idh = "D", "H", "D")
                XSoles = isol
                XDolares = idol
            End If
            
            If CE(RsCargos(I)) <> Empty And Trim(RsCargos(I)) <> "" Then
                XCuenta = RsCargos(I)
                XGENERADA = "S"
                Ins_CnMovi ilib, itdoc, idivi, XCuenta, icolcv, ivou, iserdoc, idocum, icorr, imoneda, iaux, icaux, iCencos, iCenco, XGENERADA, icto, ifec, XDH, XSoles, XDolares, iuser, imor, ""
            End If
        Next
        RsCargos.Close
        Set RsCargos = Nothing
        RsPorc.Close
        Set RsPorc = Nothing
    End If
    Ins_CnMovi ilib, itdoc, idivi, icta, icolcv, ivou, iserdoc, idocum, icorr, imoneda, iaux, icaux, iCencos, iCenco, "N", icto, ifec, idh, isol, idol, iuser, imor, orden
End Sub


Public Sub Ins_CnMovi(lib As String, tdoc As String, Divi As String, Cta As String, colcv As String, vou As String, Serdoc As String, DOCUM As String, corr As String, moneda As String, aux As String, caux As String, cencos As String, cenco As String, gen As String, cto As String, fec As String, dh As String, sol As Double, dol As Double, User As String, mor As String, orden As String)
    On Error GoTo mensaje
  
      SQL = "call cn_Insert_Movi ('" & lib & "','" & tdoc & "','" & Divi & "','0000000000','" & _
            vou & "','" & Serdoc & "','" & DOCUM & "','" & corr & "','" & moneda & "','" & Trim(Cta) & "','" & _
            aux & "','" & caux & "','" & cencos & "','" & cenco & "','" & gen & "','" & _
            Replace(cto, "'", "") & "'," & _
            IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
            IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
            fec & "','" & strAnoSistema & strMesSistema & "','" & User & "','" & dh & "','" & _
            colcv & "','" & mor & "','" & orden & "')"
      oConexionMYSQL.BeginTrans
      
 oConexionMYSQL.Execute (SQL)
      oConexionMYSQL.CommitTrans
      Exit Sub
mensaje:
      MsgBox "No se grabo correctamente la linea " & corr & " modificada" & Chr(10) + Chr(13) & "Uno de los datos es incorrecto ", vbOKOnly + vbInformation, "Aviso"
      Resume
      ADOConexion.RollbackTrans
End Sub



Public Function BuscarTipoCambioPorLetra(TipoDeCambio As String, fecha) As String
Dim SQL As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset


  Select Case TipoDeCambio
    Case "C": SQL = "SELECT tipo_com FROM novperuhfm.cncambi WHERE fecha = '" & fecha & "'"
    Case "V": SQL = "SELECT tipo_ven FROM novperuhfm.cncambio WHERE fecha = '" & fecha & "'"
  End Select
  Set Rs = ADO_LlenaRs(SQL)
   If Rs Is Nothing Then
        BuscarTipoCambioPorLetra = "1.00"
   Else
        BuscarTipoCambioPorLetra = FormatNumber(CE(Rs(0)), 3)
   End If
  Set Rs = Nothing
End Function


Public Function DescripAuxiliar(cauxil As String, ccod As String) As String
Dim SQL As String
Dim rsaux As ADODB.Recordset
Set rsaux = New ADODB.Recordset
    SQL = "Call CONT_Rep_Proc_Genericos('cnauxilauxcod','" & cauxil & "','" & ccod & "','','','','','','','','' );"
    Set rsaux = ADO_LlenaRs(SQL)
    
    If rsaux Is Nothing Then
        DescripAuxiliar = ""
    Else
        DescripAuxiliar = rsaux(0)
    End If
    Set rsaux = Nothing
End Function

Function DameNombreMes(ByVal NroMes As String) As String
    Select Case NroMes
        Case "01"
            DameNombreMes = "ENERO"
        Case "02"
            DameNombreMes = "FEBRERO"
        Case "03"
            DameNombreMes = "MARZO"
        Case "04"
            DameNombreMes = "ABRIL"
        Case "05"
            DameNombreMes = "MAYO"
        Case "06"
            DameNombreMes = "JUNIO"
        Case "07"
            DameNombreMes = "JULIO"
        Case "08"
            DameNombreMes = "AGOSTO"
        Case "09"
            DameNombreMes = "SETIEMBRE"
        Case "10"
            DameNombreMes = "OCTUBRE"
        Case "11"
            DameNombreMes = "NOVIEMBRE"
        Case "12"
            DameNombreMes = "DICIEMBRE"
    End Select
End Function

Public Function TieneGenerada(cuenta As String, Optional acta As String) As Boolean
Dim SQL As String
Dim RsTG As ADODB.Recordset
Set RsTG = New ADODB.Recordset
      If IsNull(cuenta) Or cuenta = "" Then
          TieneGenerada = False
      Else
            SQL = "Call CONT_Rep_Proc_Genericos('cntcnmayor','" & cuenta & "','','','','','','','','','' );"
            Set RsTG = ADO_LlenaRs(SQL)
            
            If val(RsTG(0)) > 0 Then
                TieneGenerada = True
            Else
                TieneGenerada = False
            End If
            Set RsTG = Nothing
      End If
End Function
