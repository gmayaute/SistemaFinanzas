VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmTareasEmpleado 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11790
   ClientLeft      =   705
   ClientTop       =   10965
   ClientWidth     =   12930
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmTareasEmpleado.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11790
   ScaleWidth      =   12930
   WindowState     =   2  'Maximized
   Begin NOVAdmin.flxEditfac flxDetalle 
      Height          =   1905
      Left            =   7680
      TabIndex        =   63
      Top             =   2970
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   7560
      Top             =   1920
   End
   Begin VB.PictureBox picSocket 
      Height          =   3735
      Left            =   7470
      ScaleHeight     =   3675
      ScaleWidth      =   5070
      TabIndex        =   53
      Top             =   7560
      Visible         =   0   'False
      Width           =   5130
      Begin VB.TextBox txtLog 
         Height          =   1995
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   60
         Top             =   1020
         Width           =   4815
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1500
         TabIndex        =   59
         Text            =   "172.26.35.186"
         Top             =   60
         Width           =   1575
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1500
         TabIndex        =   58
         Text            =   "123"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox txtSend 
         Height          =   285
         Left            =   180
         TabIndex        =   57
         Top             =   3150
         Width           =   3735
      End
      Begin VB.CommandButton bntConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   3780
         TabIndex        =   56
         Tag             =   "Connect"
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton bntExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3780
         TabIndex        =   55
         Top             =   540
         Width           =   1095
      End
      Begin VB.CommandButton bntSend 
         Caption         =   "Send"
         Height          =   375
         Left            =   4020
         TabIndex        =   54
         Top             =   3150
         Width           =   855
      End
      Begin MSWinsockLib.Winsock sock1 
         Left            =   2580
         Top             =   540
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label5 
         Caption         =   "Remote Host IP :"
         Height          =   255
         Left            =   180
         TabIndex        =   62
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Remote Host Port :"
         Height          =   255
         Left            =   60
         TabIndex        =   61
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.PictureBox picWindow 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   16170
      Left            =   1470
      ScaleHeight     =   16170
      ScaleWidth      =   6495
      TabIndex        =   30
      Top             =   2580
      Width           =   6495
      Begin VB.ComboBox cboedit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   630
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   285
         Left            =   1260
         TabIndex        =   33
         Top             =   9810
         Width           =   4830
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   8805
         Left            =   6180
         TabIndex        =   32
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox picDatos 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   14220
         Left            =   480
         ScaleHeight     =   14220
         ScaleWidth      =   5370
         TabIndex        =   31
         Top             =   -2910
         Width           =   5370
         Begin NOVAdmin.flxEditfac flxtareasSemanal 
            Height          =   1725
            Left            =   60
            TabIndex        =   65
            Top             =   3480
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3043
         End
         Begin VB.PictureBox picFondo 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1950
            Index           =   6
            Left            =   60
            ScaleHeight     =   1950
            ScaleWidth      =   4965
            TabIndex        =   50
            Top             =   12030
            Width           =   4965
            Begin VB.PictureBox flxtareasEventual 
               FillColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1635
               Left            =   0
               ScaleHeight     =   1575
               ScaleWidth      =   4530
               TabIndex        =   51
               Top             =   270
               Width           =   4590
            End
            Begin VB.Label lblTipo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "EVENTUAL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   6
               Left            =   0
               TabIndex        =   52
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.PictureBox picFondo 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1950
            Index           =   5
            Left            =   45
            ScaleHeight     =   1950
            ScaleWidth      =   4965
            TabIndex        =   43
            Top             =   10035
            Width           =   4965
            Begin NOVAdmin.flxEditfac flxtareasAnual 
               Height          =   1545
               Left            =   -30
               TabIndex        =   69
               Top             =   180
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   2725
            End
            Begin VB.Label lblTipo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ANUAL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   5
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Width           =   630
            End
         End
         Begin VB.PictureBox picFondo 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1950
            Index           =   4
            Left            =   60
            ScaleHeight     =   1950
            ScaleWidth      =   4965
            TabIndex        =   42
            Top             =   8040
            Width           =   4965
            Begin NOVAdmin.flxEditfac flxtareasSemestral 
               Height          =   1725
               Left            =   390
               TabIndex        =   68
               Top             =   -1320
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   3043
            End
            Begin VB.Label lblTipo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "SEMESTRAL:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   4
               Left            =   0
               TabIndex        =   48
               Top             =   0
               Width           =   1185
            End
         End
         Begin VB.PictureBox picFondo 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1950
            Index           =   3
            Left            =   45
            ScaleHeight     =   1950
            ScaleWidth      =   4965
            TabIndex        =   41
            Top             =   6075
            Width           =   4965
            Begin NOVAdmin.flxEditfac flxtareasTrimestral 
               Height          =   1725
               Left            =   -360
               TabIndex        =   67
               Top             =   180
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   3043
            End
            Begin VB.Label lblTipo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "TRIMESTRAL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   47
               Top             =   0
               Width           =   1200
            End
         End
         Begin VB.PictureBox picFondo 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1950
            Index           =   2
            Left            =   120
            ScaleHeight     =   1950
            ScaleWidth      =   4965
            TabIndex        =   40
            Top             =   4290
            Width           =   4965
            Begin NOVAdmin.flxEditfac flxtareasMensual 
               Height          =   1755
               Left            =   270
               TabIndex        =   66
               Top             =   1860
               Width           =   4755
               _ExtentX        =   8387
               _ExtentY        =   3096
            End
            Begin VB.Label lblTipo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "MENSUAL:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picFondo 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1950
            Index           =   1
            Left            =   45
            ScaleHeight     =   1950
            ScaleWidth      =   4965
            TabIndex        =   39
            Top             =   2115
            Width           =   4965
            Begin VB.Label lblTipo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "SEMANAL:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   945
            End
         End
         Begin VB.PictureBox picFondo 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1950
            Index           =   0
            Left            =   45
            ScaleHeight     =   1950
            ScaleWidth      =   4965
            TabIndex        =   38
            Top             =   135
            Width           =   4965
            Begin NOVAdmin.flxEditfac flxtareas 
               Height          =   1785
               Left            =   0
               TabIndex        =   64
               Top             =   210
               Width           =   4635
               _ExtentX        =   8176
               _ExtentY        =   3149
            End
            Begin VB.Label lblTipo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "DIARIO:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   720
            End
         End
      End
   End
   Begin VB.Frame fraDesEmpleado 
      BackColor       =   &H009F5539&
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   7695
      TabIndex        =   19
      Top             =   4995
      Width           =   3525
      Begin VB.Label lblDescripcion 
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
         Height          =   555
         Left            =   60
         TabIndex        =   20
         Top             =   180
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H009F5539&
      Caption         =   "Funciones Laborales del Empleado "
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
      Height          =   1365
      Left            =   6000
      TabIndex        =   23
      Top             =   45
      Width           =   6825
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1035
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   6585
         Begin Proyecto1.chameleonButton cmdgrabar 
            Height          =   315
            Left            =   6000
            TabIndex        =   25
            ToolTipText     =   "Grabar datos del Empleado "
            Top             =   180
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   556
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
            MICON           =   "frmTareasEmpleado.frx":014A
            PICN            =   "frmTareasEmpleado.frx":0166
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSForms.TextBox txtfunciones 
            Height          =   495
            Left            =   1125
            TabIndex        =   29
            Top             =   480
            Width           =   4740
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            ScrollBars      =   2
            Size            =   "8361;873"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtposicion 
            Height          =   300
            Left            =   1125
            TabIndex        =   28
            Top             =   165
            Width           =   4740
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "8361;529"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Funciones:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   27
            Top             =   555
            Width           =   945
         End
         Begin VB.Label Lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Posición:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   26
            Top             =   180
            Width           =   795
         End
      End
   End
   Begin VB.Frame fraBotones 
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   7695
      TabIndex        =   16
      Top             =   5955
      Width           =   3525
      Begin Proyecto1.chameleonButton chBtnGrabar 
         Height          =   375
         Left            =   1245
         TabIndex        =   18
         ToolTipText     =   "Grabar Cambios en Tareas "
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         MICON           =   "frmTareasEmpleado.frx":0700
         PICN            =   "frmTareasEmpleado.frx":071C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton chBtnRefreshDet 
         Height          =   375
         Left            =   2340
         TabIndex        =   17
         ToolTipText     =   " Cargar lista de dias por Tarea "
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         MICON           =   "frmTareasEmpleado.frx":0CB6
         PICN            =   "frmTareasEmpleado.frx":0CD2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton chBtnEliminar 
         Height          =   375
         Left            =   135
         TabIndex        =   35
         ToolTipText     =   " Terminar Tareas Seleccionadas  "
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         MICON           =   "frmTareasEmpleado.frx":126C
         PICN            =   "frmTareasEmpleado.frx":1288
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
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Caption         =   "Filtros de Busqueda "
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
      Height          =   1350
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   5880
      Begin VB.Frame fr3 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   1170
         TabIndex        =   3
         Top             =   1530
         Visible         =   0   'False
         Width           =   2595
         Begin MSForms.ComboBox cbonumdia 
            Height          =   315
            Left            =   1230
            TabIndex        =   14
            Top             =   150
            Width           =   900
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "1587;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblFecha 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Día"
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
            Left            =   -645
            TabIndex        =   4
            Top             =   135
            Width           =   435
         End
      End
      Begin VB.Frame fr4 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6060
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   2625
         Begin MSForms.ComboBox cbomes 
            Height          =   315
            Left            =   360
            TabIndex        =   7
            Top             =   150
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   -165
            TabIndex        =   6
            Top             =   165
            Width           =   510
         End
      End
      Begin VB.Frame fr1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7410
         TabIndex        =   8
         Top             =   780
         Visible         =   0   'False
         Width           =   2565
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Prioridad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   15
            TabIndex        =   10
            Top             =   15
            Width           =   825
         End
         Begin MSForms.ComboBox cboprioridad 
            Height          =   315
            Left            =   855
            TabIndex        =   9
            Top             =   0
            Width           =   1650
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2910;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fr2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7650
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   2565
         Begin MSForms.ComboBox cbodia 
            Height          =   315
            Left            =   540
            TabIndex        =   13
            Top             =   0
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
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Días"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   15
            TabIndex        =   12
            Top             =   15
            Width           =   510
         End
      End
      Begin Proyecto1.chameleonButton chBtnReporte 
         Height          =   375
         Left            =   5355
         TabIndex        =   21
         ToolTipText     =   "Ver Reporte"
         Top             =   825
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
         MICON           =   "frmTareasEmpleado.frx":1822
         PICN            =   "frmTareasEmpleado.frx":183E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton chBtnRefresh 
         Height          =   375
         Left            =   4905
         TabIndex        =   22
         ToolTipText     =   " Cargar lista de Tareas "
         Top             =   825
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
         MICON           =   "frmTareasEmpleado.frx":1D80
         PICN            =   "frmTareasEmpleado.frx":1D9C
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
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes de Tarea"
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
         TabIndex        =   37
         Top             =   900
         Width           =   1185
      End
      Begin MSForms.ComboBox cboMesDet 
         Height          =   315
         Left            =   1710
         TabIndex        =   36
         Top             =   855
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
      Begin MSForms.ComboBox cboemp 
         Height          =   375
         Left            =   1710
         TabIndex        =   2
         Top             =   285
         Width           =   4035
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "7117;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre y Apellido"
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
         Index           =   4
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Width           =   1545
      End
   End
   Begin MSForms.CheckBox chk 
      Height          =   255
      Left            =   9360
      TabIndex        =   15
      Top             =   510
      Visible         =   0   'False
      Width           =   2070
      VariousPropertyBits=   746588179
      BackColor       =   10442041
      ForeColor       =   8421631
      DisplayStyle    =   4
      Size            =   "3651;450"
      Value           =   "0"
      Caption         =   "No volver a mostrar"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmTareasEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bAutorizado As Boolean
Dim Item As Integer
Dim bTermino As Boolean
Dim cNombreGridActivo As String
Dim nUbicacionesPicFondo(6) As Integer
Public FlgIngresar As Boolean
Public nFilaDet As Integer


Dim Par_UsuarioSO As String
Dim Par_UsuarioID As String
Dim Par_CodEmp As String
Dim Par_Host As String
Dim Par_IP As String
Dim cCodUsuariocliente As String

Private Function SeteaCadena() As String
    Par_UsuarioSO = UsuarioXSO
    Par_UsuarioID = strUsuarioId
    Par_CodEmp = strCodEmpleado
    Par_Host = sock1.LocalHostName
    Par_IP = sock1.LocalIP
    cCodUsuariocliente = cboemp.List(cboemp.ListIndex, 1)
    
    SeteaCadena = "P_USO=" & Par_UsuarioSO & " " & _
             "P_UID=" & Par_UsuarioID & " " & _
             "P_COD=" & Par_CodEmp & " " & _
             "P_HOS=" & Par_Host & " " & _
             "P_ADM=" & IIf(bAutorizado = True, 1, 0) & " " & _
             "P_CUC=" & cCodUsuariocliente & " " & _
             "P_NIP=" & Par_IP & " P_"
End Function


Private Sub bntConnect_Click()
    On Error GoTo T
    
    'sock1 is the name of our Winsock ActiveX Control
    
    sock1.Close 'we close it in case it was trying to connect
    
    'txtIP is the textbox holding the host IP
    'txtIP can contain both hostnames ( like www.google.com ) or IPs ( like 127.0.0.1 )
    sock1.RemoteHost = txtIP.Text 'set the remote host to the ip we wrote
                                'in the txtIP textbox
    
    
    'txtPort is the textbox holding the Port number
    sock1.RemotePort = txtPort  'set the port we want to connect to
                                '( the server must be listening on this port too)
                                
    
                                
    sock1.Connect               'try to connect
    
    
    Exit Sub
T:
    MsgBox "Error : " & err.Description, vbCritical
End Sub

Private Sub EnviarMensaje(Optional cMensaje As String = "")
    On Error GoTo T
    txtSend.Text = SeteaCadena & IIf(cMensaje <> "", " P_MSG=" & cMensaje & " P_", " P_MSG=XXX P_")
    sock1.SendData txtSend  'trasmits the string to host
    txtLog = txtLog & "Client : " & txtSend & vbCrLf
    txtSend = ""
    
    Exit Sub
T:
'   MsgBox "Error : " & err.Description
    sock1_Close

End Sub

Private Sub bntSend_Click()
    Call EnviarMensaje
End Sub


Private Sub cbodia_Change()
    Call LlenarGrid
End Sub


Private Sub HabilitarCeldas(ByRef oGrid As flxEditfac, Optional nRow As Integer = 0)
    Dim nCol As Integer
    
    With oGrid
        If nRow <> 0 Then
            oGrid.row = nRow
        End If
    
        nCol = oGrid.Col
        Select Case cNombreGridActivo
            Case "flxtareas":
                .Col = 4: .CellBackColor = ColorHabilitado
                .Col = 5: .CellBackColor = ColorDeshabilitado
                .Col = 6: .CellBackColor = ColorDeshabilitado
                .Col = 7: .CellBackColor = ColorDeshabilitado
                .TextMatrix(.row, 5) = "": .TextMatrix(.row, 6) = "": .TextMatrix(.row, 7) = ""
            
            Case "flxtareasEventual":
                .Col = 4: .CellBackColor = ColorHabilitado
                .Col = 5: .CellBackColor = ColorDeshabilitado
                .Col = 6: .CellBackColor = ColorDeshabilitado
                .Col = 7: .CellBackColor = ColorDeshabilitado
                .TextMatrix(.row, 5) = "": .TextMatrix(.row, 6) = "": .TextMatrix(.row, 7) = ""
            
            Case "flxtareasSemanal":
                .Col = 4: .CellBackColor = ColorHabilitado
                .Col = 5: .CellBackColor = ColorHabilitado
                .Col = 6: .CellBackColor = ColorDeshabilitado
                .Col = 7: .CellBackColor = ColorDeshabilitado
                .TextMatrix(.row, 6) = "": .TextMatrix(.row, 7) = ""
            
            Case "flxtareasMensual":
                .Col = 4: .CellBackColor = ColorHabilitado
                .Col = 5: .CellBackColor = ColorDeshabilitado
                .Col = 6: .CellBackColor = ColorHabilitado
                .Col = 7: .CellBackColor = ColorDeshabilitado
                .TextMatrix(.row, 5) = "": .TextMatrix(.row, 7) = ""
            
            Case Else
                .Col = 4: .CellBackColor = ColorHabilitado
                .Col = 5: .CellBackColor = ColorDeshabilitado
                .Col = 6: .CellBackColor = ColorDeshabilitado
                .Col = 7: .CellBackColor = ColorHabilitado
                .TextMatrix(.row, 5) = "": .TextMatrix(.row, 6) = ""
        End Select
        
        oGrid.Col = nCol
    End With

End Sub

Private Sub LlenarColumnasGrid(ByRef oGrid As flxEditfac)
    Dim MesFin As String, Mes1 As String
    Dim I As Long
    
    Call HabilitarCeldas(oGrid)
    
    With oGrid
        If cboedit.ListIndex > -1 Then
            If .Col = 1 Or .Col = 5 Or .Col = 6 Then
                .TextMatrix(.row, .Col) = Trim(cboedit.List(cboedit.ListIndex))
            Else
                .TextMatrix(.row, .Col) = CE(Mid(Trim(cboedit.List(cboedit.ListIndex)), 5, Len(Trim(cboedit.List(cboedit.ListIndex))) - 1))
            End If
            cboedit.Visible = False
            Select Case .Col

                    
                Case 4
                    .TextMatrix(.row, 10) = Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 1)
                    .Col = 5
                    If .CellBackColor = ColorDeshabilitado Then
                        .Col = 6
                        If .CellBackColor = ColorDeshabilitado Then
                            .Col = 7
                            If .CellBackColor = ColorDeshabilitado Then
                                GrabarFila .row, oGrid
                            End If
                        End If
                    End If
                Case 5
                    .Col = 6
                    If .CellBackColor = ColorDeshabilitado Then
                        .Col = 7
                        If .CellBackColor = ColorDeshabilitado Then
                            GrabarFila .row, oGrid
                        End If
                    End If
                Case 6
                    .Col = 7
                    If .CellBackColor = ColorDeshabilitado Then
                        GrabarFila .row, oGrid
                    End If
                    
                Case 7
                    Select Case Trim(.TextMatrix(.row, 12))
                        Case 4
                            For I = 0 To 3
                                MesFin = MesFin & Right("00" & Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 2) + (3 * I), 2) & "/"
                                Mes1 = Mes1 & Trim(NombreMes(Right("00" & Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 2) + (3 * I), 2), True)) & "/"
                            Next
                            MesFin = Mid(MesFin, 1, Len(MesFin) - 1)
                            Mes1 = Mid(Mes1, 1, Len(Mes1) - 1)
                            .TextMatrix(.row, 11) = MesFin
                            .TextMatrix(.row, .Col) = Mes1
                        Case 5
                            MesFin = Trim(NombreMes(Right("00" & IIf((val(Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 2)) + 5) > 12, 12, val(Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 2)) + 5), 2), True))
                            .TextMatrix(.row, .Col) = NombreMes(Left(cboedit.List(cboedit.ListIndex), 2), True) & "/" & MesFin
                            .TextMatrix(.row, 11) = Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 2) & "/" & Right("00" & IIf((val(Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 2)) + 5) > 12, 12, val(Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 2)) + 5), 2)
                        Case Else
                            .TextMatrix(.row, 11) = Mid(Trim(cboedit.List(cboedit.ListIndex)), 1, 2)
                    End Select
                    GrabarFila .row, oGrid
            End Select
        Else
            .TextMatrix(.row, .Col) = ""
        End If
    End With

End Sub

Private Sub cboedit_Click()
    Select Case cNombreGridActivo
        Case "flxtareas": Call LlenarColumnasGrid(flxtareas)
        Case "flxtareasSemanal": Call LlenarColumnasGrid(flxtareasSemanal)
        Case "flxtareasMensual": Call LlenarColumnasGrid(flxtareasMensual)
        Case "flxtareasTrimestral": Call LlenarColumnasGrid(flxtareasTrimestral)
        Case "flxtareasSemestral": Call LlenarColumnasGrid(flxtareasSemestral)
        Case "flxtareasAnual": Call LlenarColumnasGrid(flxtareasAnual)
        Case "flxtareasEventual": Call LlenarColumnasGrid(flxtareasEventual)
    End Select
    
End Sub

Private Function BuscaFechaInicial(cAnio As String, cMes As String, nNumSemana) As String
    Dim nCad As Double
    Dim cUltimaFecha As String
    Dim nNum As Integer
    Dim cDia As String
    
    cUltimaFecha = UltimoDiaMes(cMes, cAnio)
    
    If nNumSemana = 1 Then
        BuscaFechaInicial = "01/" & cMes & "/" & cAnio
        Exit Function
    End If
    
    nNum = 1
    
    For nCad = NE(cAnio & cMes & "01") To NE(cAnio & cMes & Left(cUltimaFecha, 2))

        If nNum = nNumSemana Then
            BuscaFechaInicial = Right(nCad, 2) & "/" & cMes & "/" & cAnio
            Exit For
        End If

        cDia = Left(UCase(WeekdayName(Weekday(Right(nCad, 2) & "/" & Mid(nCad, 5, 2) & "/" & Left(nCad, 4)))), 3)

        If cDia = "DOM" Then
            nNum = nNum + 1
        End If
        
        
    Next nCad
    
    
End Function

Sub GrabarFilaDetalle(fila As Integer, it As Integer, bTipoActualiza As Boolean, ByRef oGrid As flxEditfac)
    
    
    Dim SQL As String
    Dim Item As Integer
    Dim cFecha As String
    Dim nPorcentaje As Double
    Dim nInicio As String
    Dim nFin As String, J As String
    Dim cFechaAux As String
    Dim bInsertar As Boolean
    Dim nPos As Integer
    bInsertar = False
    
    nInicio = strAnoSistema & "0101"
    nFin = strAnoSistema & "1231"
    cFecha = Right(nInicio, 2) & "/" & Mid(nInicio, 5, 2) & "/" & Left(nInicio, 4)
    '------------------------------
    Item = 1
    nPos = 1
    
    'SI ES TRIMESTRAL, SEMESTRAL O ANUAL PEDIR DIA DE INICIO
    With oGrid
    
    
    If bTipoActualiza = True Then
        it = NE(.TextMatrix(.row, 8))
    End If
    Dim cDia As String
    Dim nDia As Integer
    Dim cFechaInicial As String
    
    ' BUSCA EL DIA SI ES TRIMESTRAL, SEMESTRAL, ANUAL
    If Trim(.TextMatrix(fila, 12)) >= "4" And Trim(.TextMatrix(fila, 12)) <= "6" Then
        nDia = CE(InputBox("Ingrese el número de semana que iniciara la tarea", "Numero de semana de inicio", 1))
        
        If NE(nDia) = 0 Then nDia = 1
        
        cFechaInicial = BuscaFechaInicial(strAnoSistema, strMesSistema, nDia)
        cDia = Left(cFechaInicial, 2)
    End If
    
    ' TAREAS EVENTUALES
    If Trim(.TextMatrix(fila, 12)) = "7" Then
        cFechaInicial = CE(InputBox("Ingrese la fecha de la tarea eventual", "Fecha de Tarea", Date))
        
        If IsDate(cFechaInicial) = False Then cFechaInicial = Date
        
        cDia = Left(cFechaInicial, 2)
    
        cFecha = cFechaInicial
        
        nInicio = Right(cFecha, 4) & Mid(cFecha, 4, 2) & cDia
        nFin = Right(cFecha, 4) & Mid(cFecha, 4, 2) & cDia
        
    
    End If
    
    nPorcentaje = 0
    
    
    '***********************************************
    SQL = "delete from rh_tareas_detalle where codemp = '" & .TextMatrix(.row, 9) & "' and item = " & it & ""
    oConexionSQL.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    
    
    '***********************************************

    
    Do While nInicio <= nFin

        ' SI ES DIARIO
        If Trim(.TextMatrix(fila, 12)) = "1" Then
            If Right(cFecha, 4) & Mid(cFecha, 4, 2) >= strAnoSistema & strMesSistema Then
                bInsertar = True
            End If
        End If

        ' SI ES EVENTUAL
        If Trim(.TextMatrix(fila, 12)) = "7" Then
            If Right(cFecha, 4) & Mid(cFecha, 4, 2) >= strAnoSistema & strMesSistema Then
                bInsertar = True
            End If
        End If

        
        ' SI ES SEMANAL
        If Trim(.TextMatrix(fila, 12)) = "2" Then
            If Right(cFecha, 4) & Mid(cFecha, 4, 2) >= strAnoSistema & strMesSistema Then
                If Left(UCase(WeekdayName(Weekday(cFecha))), 2) = Left(CE(.TextMatrix(fila, 5)), 2) Then
                    bInsertar = True
                End If
            End If
        End If
        
        'SI ES MENSUAL
        If Trim(.TextMatrix(fila, 12)) = "3" Then
            If Right(cFecha, 4) & Mid(cFecha, 4, 2) >= strAnoSistema & strMesSistema Then
                If Left(cFecha, 2) = CE(.TextMatrix(fila, 6)) Then
                    bInsertar = True
                End If
            End If
        End If

        'SI ES TRIMESTRAL
        If Trim(.TextMatrix(fila, 12)) = "4" Then
            cFechaAux = RetornaFechaPeriodos(CE(.TextMatrix(fila, 7)), nPos)
            DoEvents
            If cFechaAux <> "" Then
                If Left(cFecha, 5) = cDia & "/" & cFechaAux Then
                    If Right(cFecha, 4) & Mid(cFecha, 4, 2) >= strAnoSistema & strMesSistema Then
                    
                    
                        If Left(UCase(WeekdayName(Weekday(cFecha))), 1) = "S" Then
                            cFecha = DateAdd("d", 1, CDate(cFecha))
                        End If
                    
                        If Left(UCase(WeekdayName(Weekday(cFecha))), 1) = "D" Then
                            cFecha = DateAdd("d", 1, CDate(cFecha))
                        End If
                    
                        bInsertar = True
                    End If
                    nPos = nPos + 1
                End If
            End If
        End If

        'SI ES SEMESTRAL
        If Trim(.TextMatrix(fila, 12)) = "5" Then
            cFechaAux = RetornaFechaPeriodos(CE(.TextMatrix(fila, 7)), nPos)
            DoEvents
            If cFechaAux <> "" Then
                If Left(cFecha, 5) = cDia & "/" & cFechaAux Then
                    If Right(cFecha, 4) & Mid(cFecha, 4, 2) >= strAnoSistema & strMesSistema Then
                    
                    
                        If Left(UCase(WeekdayName(Weekday(cFecha))), 1) = "S" Then
                            cFecha = DateAdd("d", 1, CDate(cFecha))
                        End If
                    
                        If Left(UCase(WeekdayName(Weekday(cFecha))), 1) = "D" Then
                            cFecha = DateAdd("d", 1, CDate(cFecha))
                        End If
                                
                    
                        bInsertar = True
                    End If
                    nPos = nPos + 1
                End If
            End If
        End If

        'SI ES ANUAL
        If Trim(.TextMatrix(fila, 12)) = "6" Then
            cFechaAux = RetornaNumeroMes(Left(CE(.TextMatrix(fila, 7)), 3))
            DoEvents
            If cFechaAux <> "" Then
                If Left(cFecha, 5) = cDia & "/" & cFechaAux Then
                    'If Right(cFecha, 4) & Mid(cFecha, 4, 2) >= strAnoSistema & strMesSistema Then
                    
                        If Left(UCase(WeekdayName(Weekday(cFecha))), 1) = "S" Then
                            cFecha = DateAdd("d", 1, CDate(cFecha))
                        End If
                    
                        If Left(UCase(WeekdayName(Weekday(cFecha))), 1) = "D" Then
                            cFecha = DateAdd("d", 1, CDate(cFecha))
                        End If
                    
                        bInsertar = True
                    'End If
                    nPos = nPos + 1
                End If
            End If
        End If


        If bInsertar = True Then
            If Right(cFecha, 4) & Mid(cFecha, 4, 2) & Left(cFecha, 2) <= Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) Then
                nPorcentaje = 0
            Else
                nPorcentaje = 0
            End If
            
        
        
            SQL = "insert into rh_tareas_detalle(codemp,item,anio,mes,item_det,fecha, porcentaje) " & _
                  "values ( '" & CE(.TextMatrix(fila, 9)) & "'," & it & ",'" & Right(cFecha, 4) & "','" & Mid(cFecha, 4, 2) & "', " & _
                  Item & ",'" & cFecha & "'," & nPorcentaje & ")"
                  
            Call oConexionSQL.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, False)
            bInsertar = False
        End If
        
        Item = Item + 1
        cFecha = DateAdd("d", 1, CDate(cFecha))
        
        '****************
        nInicio = Right(cFecha, 4) & Mid(cFecha, 4, 2) & Left(cFecha, 2)
        '****************
        
    Loop
    
    Call FormatoGridPersonalizado(oGrid, False)
    Call LlenarGridObjeto(oGrid)

    Call AgregaUltimoRegistro(oGrid)

    End With
End Sub

Sub GrabarFila(fila As Integer, ByRef oGrid As flxEditfac)
Screen.MousePointer = vbHourglass
Dim cMeses As String
DoEvents
On Error GoTo CtrlError
    If Validar(fila, oGrid) Then
        Dim it As Integer, SQL As String
        
        With oGrid
            it = UltimoItem(oGrid, fila)
            
            cMeses = CE(.TextMatrix(fila, 11))
            cMeses = Replace(cMeses, "/13", "")
            cMeses = Replace(cMeses, "/14", "")
            cMeses = Replace(cMeses, "/15", "")
            cMeses = Replace(cMeses, "/16", "")
            
            If Trim(.TextMatrix(fila, 8)) = "" Then
                
                SQL = "insert into rh_tareasxperiodo(codemp,codtipotarea,item,tarea,prioridad,dia,numdia,mes,porcent) " & _
                      "values('" & CE(.TextMatrix(fila, 9)) & "','" & CE(.TextMatrix(fila, 12)) & "'," & it & ", " & _
                      "'" & CE(.TextMatrix(fila, 3)) & "'," & NE(.TextMatrix(fila, 10)) & ",'" & CE(.TextMatrix(fila, 5)) & "', " & _
                      "'" & CE(.TextMatrix(fila, 6)) & "','" & cMeses & "'," & NE(.TextMatrix(fila, 13)) & ")"
                
                If oConexionSQL.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, False) Then
                    Call GrabarFilaDetalle(fila, it, False, oGrid)
                    
                    Call EnviarMensaje("TAREA " & CE(.TextMatrix(fila, 2)) & " CREADA:" & Salto(1) & CE(.TextMatrix(fila, 3)) & Salto(2) & "PRIORIDAD:" & Salto(1) & CE(.TextMatrix(fila, 4)) & Salto(2) & "POR USUARIO:" & Par_UsuarioID)
                End If
            Else
                SQL = "update rh_tareasxperiodo set codtipotarea='" & CE(.TextMatrix(fila, 12)) & "', " & _
                      "dia='" & CE(.TextMatrix(fila, 5)) & "', numdia='" & CE(.TextMatrix(fila, 6)) & "', mes = '" & cMeses & "', " & _
                      "tarea='" & CE(.TextMatrix(fila, 3)) & "',prioridad=" & NE(.TextMatrix(fila, 10)) & " " & _
                      "where codemp = '" & CE(.TextMatrix(fila, 9)) & "' and item = " & CE(.TextMatrix(fila, 8)) & ""
                If oConexionSQL.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Modificar, False) Then
                    Call GrabarFilaDetalle(fila, it, True, oGrid)
                    
                    Call EnviarMensaje("TAREA " & CE(.TextMatrix(fila, 2)) & " MODIFICADA:" & Salto(1) & CE(.TextMatrix(fila, 3)) & Salto(2) & "PRIORIDAD:" & Salto(1) & CE(.TextMatrix(fila, 4)) & Salto(2) & "POR USUARIO:" & Par_UsuarioID)
               
                End If
            End If
        End With
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
CtrlError:
    MsgBox err.Description, vbCritical, "Error Registrando y/o actualizando datos."
    Screen.MousePointer = vbNormal
End Sub

Private Function RetornaNumeroMes(cDesMes As String) As String
    Select Case cDesMes
        Case "ENE": RetornaNumeroMes = "01"
        Case "FEB": RetornaNumeroMes = "02"
        Case "MAR": RetornaNumeroMes = "03"
        Case "ABR": RetornaNumeroMes = "04"
        Case "MAY": RetornaNumeroMes = "05"
        Case "JUN": RetornaNumeroMes = "06"
        Case "JUL": RetornaNumeroMes = "07"
        Case "AGO": RetornaNumeroMes = "08"
        Case "SET", "SEP": RetornaNumeroMes = "09"
        Case "OCT": RetornaNumeroMes = "10"
        Case "NOV": RetornaNumeroMes = "11"
        Case "DIC": RetornaNumeroMes = "12"
        Case Else: RetornaNumeroMes = ""
    End Select
End Function


Private Function RetornaFechaPeriodos(cCadena As String, nPos As Integer) As String
    On Error GoTo SERROR
    Dim I As Integer
    Dim cEncontro As String
    Dim nEncontro As Integer
    nEncontro = 0
    
    For I = 1 To Len(cCadena)
        If Mid(cCadena, I, 1) = "/" Then
            nEncontro = nEncontro + 1
            If nEncontro = nPos Then Exit For
        End If
    Next I
    
        
    cEncontro = Mid(cCadena, I - 3, 3)
    If InStr(1, cEncontro, "/") > 0 Then
        cEncontro = ""
    End If
    
    
    RetornaFechaPeriodos = RetornaNumeroMes(cEncontro)
        
    Exit Function
SERROR:
    RetornaFechaPeriodos = ""

End Function

Private Function UltimoItem(ByRef oGrid As flxEditfac, fila As Integer) As Integer
    Dim SQL As String, RQ As MYSQL_RS
    SQL = "select ifnull(max(item),0) as cant from rh_tareasxperiodo where codemp = '" & oGrid.TextMatrix(fila, 9) & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        UltimoItem = val(RQ.Fields("cant")) + 1
    End If
    Set RQ = Nothing
End Function

Function Validar(fila As Integer, ByRef oGrid As flxEditfac) As Boolean
    Validar = True
    With oGrid
        If Trim(.TextMatrix(fila, 1)) = "" Then
            MsgBox "Falta ingresar el Empleado", vbInformation, gsNomSW
            .Col = 1: .row = fila
            '.SetFocus
            Validar = False
            Exit Function
        End If
        If Trim(.TextMatrix(fila, 2)) <> "" Then
            Select Case .TextMatrix(fila, 12)
                Case 1
                    If .TextMatrix(fila, 4) = "" Then
                        MsgBox "Falta ingresar la prioridad de la tarea", vbInformation, gsNomSW
                        .Col = 4: .row = fila
                        '.SetFocus
                        Validar = False
                    End If
                Case 2
                    If .TextMatrix(fila, 5) = "" Then
                        MsgBox "Falta ingresar el día de la semana en la que se realizará la tarea", vbInformation, gsNomSW
                        .Col = 5: .row = fila
                        '.SetFocus
                        Validar = False
                    End If
                Case 3
                    If .TextMatrix(fila, 6) = "" Then
                        MsgBox "Falta ingresar el día de la semana en la que se realizará la tarea", vbInformation, gsNomSW
                        .Col = 6: .row = fila
                        '.SetFocus
                        Validar = False
                    End If
                Case 4, 5, 6
                    If .TextMatrix(fila, 7) = "" Then
                        MsgBox "Falta ingresar el Mes en el cual se realizará la tarea", vbInformation, gsNomSW
                        .Col = 7: .row = fila
                        '.SetFocus
                        Validar = False
                    End If
            End Select
            Exit Function
        End If
        If Trim(.TextMatrix(fila, 3)) = "" Then
            MsgBox "Falta ingresar la descripción de la tarea", vbInformation, gsNomSW
            .Col = 3: .row = fila
            '.SetFocus
            Validar = False
            Exit Function
        End If
    End With
End Function
Private Sub cboedit_LostFocus()
    cboedit.Visible = False
End Sub

Private Function RetornaCodigoTarea(ByRef oGrid As flxEditfac) As String
    Select Case oGrid.Nombre
        Case "flxtareas": RetornaCodigoTarea = "1"
        Case "flxtareasSemanal": RetornaCodigoTarea = "2"
        Case "flxtareasMensual": RetornaCodigoTarea = "3"
        Case "flxtareasTrimestral": RetornaCodigoTarea = "4"
        Case "flxtareasSemestral": RetornaCodigoTarea = "5"
        Case "flxtareasAnual": RetornaCodigoTarea = "6"
        Case "flxtareasEventual": RetornaCodigoTarea = "7"
    End Select

End Function

Private Function RetornaNombreTarea(ByRef oGrid As flxEditfac) As String
    Select Case oGrid.Nombre
        Case "flxtareas": RetornaNombreTarea = "DIARIO"
        Case "flxtareasSemanal": RetornaNombreTarea = "SEMANAL"
        Case "flxtareasMensual": RetornaNombreTarea = "MENSUAL"
        Case "flxtareasTrimestral": RetornaNombreTarea = "TRIMESTRAL"
        Case "flxtareasSemestral": RetornaNombreTarea = "SEMESTRAL"
        Case "flxtareasAnual": RetornaNombreTarea = "ANUAL"
        Case "flxtareasEventual": RetornaNombreTarea = "EVENTUAL"
    End Select
End Function


Private Function RetornaLeft(ByRef oGrid As flxEditfac) As Long
    Select Case oGrid.Nombre
        Case "flxtareas": RetornaLeft = flxtareas.Left
        Case "flxtareasSemanal": RetornaLeft = flxtareasSemanal.Left
        Case "flxtareasMensual": RetornaLeft = flxtareasMensual.Left
        Case "flxtareasTrimestral": RetornaLeft = flxtareasTrimestral.Left
        Case "flxtareasSemestral": RetornaLeft = flxtareasSemestral.Left
        Case "flxtareasAnual": RetornaLeft = flxtareasAnual.Left
        Case "flxtareasEventual": RetornaLeft = flxtareasEventual.Left
    End Select
End Function

Private Function RetornaTop(ByRef oGrid As flxEditfac) As Long
    Select Case oGrid.Nombre
        Case "flxtareas": RetornaTop = picFondo(0).Top - VScroll1.Value + 270
        Case "flxtareasSemanal": RetornaTop = picFondo(1).Top - VScroll1.Value + 270
        Case "flxtareasMensual": RetornaTop = picFondo(2).Top - VScroll1.Value + 270
        Case "flxtareasTrimestral": RetornaTop = picFondo(3).Top - VScroll1.Value + 270
        Case "flxtareasSemestral": RetornaTop = picFondo(4).Top - VScroll1.Value + 270
        Case "flxtareasAnual": RetornaTop = picFondo(5).Top - VScroll1.Value + 270
        Case "flxtareasEventual": RetornaTop = picFondo(6).Top - VScroll1.Value + 270
        
       ' RetornaTop = RetornaTop + 1950
    End Select
End Function

Private Sub AgregaUltimoRegistro(ByRef oGrid As flxEditfac)
    Dim I As Long
    With oGrid
        For I = 1 To oGrid.Rows - 1
            If cboemp.ListIndex > 0 Then
                .TextMatrix(I, 1) = cboemp.Text
                .TextMatrix(I, 9) = cboemp.List(cboemp.ListIndex, 1)
                
                .TextMatrix(I, 12) = RetornaCodigoTarea(oGrid)
                .TextMatrix(I, 2) = RetornaNombreTarea(oGrid)
                
            End If
            .Col = 1
            If Trim(.TextMatrix(.row, 1)) <> "" Then .CellBackColor = ColorDeshabilitado
        Next
        If cboemp.ListIndex > 0 Then
            cmdgrabar.Enabled = True
            DatosEmp Trim(.TextMatrix(1, 9))
            
        Else
            cmdgrabar.Enabled = False
            txtposicion = "": txtfunciones = ""
            
        End If
        ' If .Visible = True Then .Col = 2: .SetFocus
    End With

End Sub

Private Sub chBtnRefresh_Click()
    Call LlenarGrid
    
    Call AgregaUltimoRegistro(flxtareas)
    Call AgregaUltimoRegistro(flxtareasSemanal)
    Call AgregaUltimoRegistro(flxtareasMensual)
    Call AgregaUltimoRegistro(flxtareasTrimestral)
    Call AgregaUltimoRegistro(flxtareasSemestral)
    Call AgregaUltimoRegistro(flxtareasAnual)
    Call AgregaUltimoRegistro(flxtareasEventual)


End Sub

Private Sub cboemp_Change()
    Call chBtnRefresh_Click
End Sub
Private Sub cboMes_Change()
    Call LlenarGrid
End Sub

Private Sub cboMesDet_Change()
    Call chBtnRefresh_Click
    DoEvents
    On Error Resume Next
    cboMesDet.SetFocus
End Sub

Private Sub RefrescarxGrilla()
    Select Case cNombreGridActivo
        Case "flxtareas": Call LlenarGridDetalle(flxtareas)
        Case "flxtareasSemanal": Call LlenarGridDetalle(flxtareasSemanal)
        Case "flxtareasMensual": Call LlenarGridDetalle(flxtareasMensual)
        Case "flxtareasTrimestral": Call LlenarGridDetalle(flxtareasTrimestral)
        Case "flxtareasSemestral": Call LlenarGridDetalle(flxtareasSemestral)
        Case "flxtareasAnual": Call LlenarGridDetalle(flxtareasAnual)
        Case "flxtareasEventual": Call LlenarGridDetalle(flxtareasEventual)
    End Select

End Sub

Private Sub cbonumdia_Change()
    LlenarGrid
End Sub
Private Sub cboprioridad_Change()
    LlenarGrid
End Sub

Private Sub FormatoGridDetalle()

    Dim nColDia As Integer
    Dim nColFecha As Integer
    Dim nColSemana As Integer
    
    nColDia = 500
    nColFecha = 1000
    nColSemana = 700
    
    If cboMesDet.ListIndex > 0 Then
        Select Case cNombreGridActivo
            Case "flxtareas":
                nColSemana = 0
                
            Case "flxtareasSemanal":
                nColFecha = 0
            
            Case "flxtareasMensual":
                nColSemana = 0
            
            Case "flxtareasTrimestral":
                nColDia = 0
                nColFecha = 0
            
            Case "flxtareasSemestral":
                nColDia = 0
                nColFecha = 0
            
            Case "flxtareasAnual":
                'nColDia = 0
                'nColFecha = 0
                
            Case "flxtareasEventual":
                nColSemana = 0
        
        End Select
    End If
    
    With flxDetalle
        .Refresh
        .Clear
        .Rows = 2
        .Cols = 11
        .RowHeight(1) = 360
        .FixedCols = 7
        .FixedRows = 1
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = "Item"
        .ForeColor = vbBlack
        .ForeColorFixed = vbBlack
        
        
        .ColWidth(1) = nColDia
        .TextMatrix(0, 1) = "Dia"
        .ColType(1) = cadena
        .CaracteresValidos(1) = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz"
        
        
        .ColWidth(2) = 0 '3000
        .TextMatrix(0, 2) = Space(22) + "Empleado"
        .ColType(2) = cadena
        .CaracteresValidos(2) = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz"
        
        .ColWidth(3) = 0 '1000
        .TextMatrix(0, 3) = Space(7) + "Año"
        .ColType(3) = cadena
        .CaracteresValidos(3) = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz"
        
        
        .ColWidth(4) = 0 '6100
        .TextMatrix(0, 4) = Space(60) + "Mes"
        .ColType(4) = cadena
        .CaracteresValidos(4) = Chr(13) & "abcdefghijklmnñopqrstuvwxyz" & UCase("abcdefghijklmnñopqrstuvwxyz") & "1234567890.,()-;:{}\/#_*+´" & Chr(34) & Chr(167) & Chr(39) & Chr(32) & Chr(8)
        
        
        .ColWidth(5) = 0 '550
        .TextMatrix(0, 5) = Space(1) + "Item"
        .ColType(5) = cadena
        .CaracteresValidos(5) = Chr(13) & "abcdefghijklmnñopqrstuvwxyz" & UCase("abcdefghijklmnñopqrstuvwxyz") & "1234567890.,()-;:{}\/#_*+´" & Chr(34) & Chr(167) & Chr(39) & Chr(32) & Chr(8)
        
        .ColWidth(6) = nColFecha
        .TextMatrix(0, 6) = "Fecha"
        .ColType(6) = cadena
        .CaracteresValidos(6) = Chr(13) & "abcdefghijklmnñopqrstuvwxyz" & UCase("abcdefghijklmnñopqrstuvwxyz") & "1234567890.,()-;:{}\/#_*+´" & Chr(34) & Chr(167) & Chr(39) & Chr(32) & Chr(8)
        
        .ColWidth(7) = nColSemana
        .TextMatrix(0, 7) = "Semana"
        .ColType(7) = Numero
        .CaracteresValidos(7) = "0123456789"
        
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = " % "
        .ColType(8) = Numero
        .CaracteresValidos(8) = "0123456789"
        
        .ColWidth(9) = 500
        .TextMatrix(0, 9) = "ST."
        .ColType(9) = cadena
        .CaracteresValidos(9) = " "
        
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = " Item Det"
        .ColType(10) = Numero
        .CaracteresValidos(10) = "0123456789"
        
        
        flxtareas.SelectionMode = flexSelectionFree
    End With

End Sub

Private Sub ColorDetalle(nFila As Integer)
    With flxDetalle
        .row = nFila
        .Col = 9
        If NE(.TextMatrix(nFila, 8)) = 100 Then
            .CellBackColor = vbBlue
            .CellForeColor = vbBlue
        Else
            .CellBackColor = ColorHabilitado
            .CellForeColor = ColorHabilitado
        End If
        
    End With
        
End Sub

Private Sub FormatoGridPersonalizado(oFl As flxEditfac, Optional bVisible As Boolean = True)
    
    Dim nColDiaSem As Integer
    Dim nColDia As Integer
    Dim nColMes As Integer
    Dim nColTarea As Integer
    Dim nColItem As Integer
    Dim nColPrior As Integer
    
    nColTarea = 7000
    nColDiaSem = 650
    nColDia = 380
    nColMes = 1400
    nColItem = 400
    nColPrior = 550
    
    Select Case oFl.Nombre
        Case "flxtareas":
            nColDiaSem = 0
            nColDia = 0
            nColMes = 0
            nColTarea = flxtareas.Width - nColItem - nColPrior - nColDiaSem - nColDia - nColMes - 400
        
        Case "flxtareasSemanal":
            nColDia = 0
            nColMes = 0
            nColTarea = 7000
            nColTarea = flxtareasSemanal.Width - nColItem - nColPrior - nColDiaSem - nColDia - nColMes - 400
        
        Case "flxtareasMensual":
            nColDiaSem = 0
            nColMes = 0
            nColTarea = flxtareasMensual.Width - nColItem - nColPrior - nColDiaSem - nColDia - nColMes - 400
        
        Case "flxtareasTrimestral":
            nColDiaSem = 0
            nColDia = 0
            nColTarea = flxtareasTrimestral.Width - nColItem - nColPrior - nColDiaSem - nColDia - nColMes - 400
        
        Case "flxtareasSemestral":
            nColDiaSem = 0
            nColDia = 0
            nColTarea = flxtareasSemestral.Width - nColItem - nColPrior - nColDiaSem - nColDia - nColMes - 400
        
        Case "flxtareasAnual":
            nColDiaSem = 0
            nColDia = 0
            nColTarea = flxtareasAnual.Width - nColItem - nColPrior - nColDiaSem - nColDia - nColMes - 400
    
        Case "flxtareasEventual":
            nColDiaSem = 0
            nColDia = 0
            nColMes = 0
            nColTarea = flxtareasEventual.Width - nColItem - nColPrior - nColDiaSem - nColDia - nColMes - 400
        
    
    End Select
    
    With oFl
        '.Visible = bVisible
        .Refresh
        .Clear
        .Rows = 2
        .Cols = 14
        .RowHeight(1) = 360
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = nColItem
        .TextMatrix(0, 0) = "Item"
        .ForeColor = vbBlack
        .ForeColorFixed = vbBlack
        
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(22) + "Empleado"
        .ColType(1) = cadena
        .CaracteresValidos(1) = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz"
        .ColWidth(2) = 0
        .TextMatrix(0, 2) = Space(7) + "Tipo"
        .ColType(2) = cadena
        .CaracteresValidos(2) = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz"
        .ColWidth(3) = nColTarea
        .TextMatrix(0, 3) = Space(60) + "Tarea"
        .ColType(3) = cadena
        .CaracteresValidos(3) = Chr(13) & "abcdefghijklmnñopqrstuvwxyz" & UCase("abcdefghijklmnñopqrstuvwxyz") & "1234567890.,()-;:{}\/#_*+´" & Chr(34) & Chr(167) & Chr(39) & Chr(32) & Chr(8)
        .ColWidth(4) = nColPrior
        .TextMatrix(0, 4) = Space(1) + "Prior."
        .ColType(4) = cadena
        .CaracteresValidos(4) = "ALTAMEDIBJ"
        .ColWidth(5) = nColDiaSem
        .TextMatrix(0, 5) = "Día Sem"
        .ColType(5) = cadena
        .CaracteresValidos(5) = "LUNESMARTICOLJVBDG"
        .ColWidth(6) = nColDia
        .TextMatrix(0, 6) = "Día"
        .ColType(6) = Numero
        .CaracteresValidos(6) = "0123456789"
        .ColWidth(7) = nColMes
        .TextMatrix(0, 7) = Space(7) + "Mes"
        .ColType(7) = cadena
        .CaracteresValidos(7) = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz"
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = "Item"
        .ColWidth(9) = 0
        .TextMatrix(0, 9) = "codemp"
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = "codprioridad"
        .ColWidth(11) = 0
        .TextMatrix(0, 11) = "mes"
        .ColWidth(12) = 0
        .TextMatrix(0, 12) = "tipo"
        
        .TextMatrix(0, 13) = " % "
        .ColWidth(13) = 0
        .ColType(13) = Numero
        '.Visible = True
        
        
        
        .NoEditar = Not bAutorizado
        
    
        cmdgrabar.Visible = bAutorizado
        txtposicion.Locked = Not bAutorizado
        txtfunciones.Locked = Not bAutorizado
    
    End With
    

End Sub




Private Sub LlenarGridDetalle(ByRef oFlxGrid As flxEditfac)
    Dim RQ As MYSQL_RS
    Dim I As Integer
        
    If bTermino = False Then Exit Sub
    
    flxDetalle.Visible = False
    
    Call FormatoGridDetalle
    
    Dim cCadena As String
    Dim SQL  As String
     
    If cboMesDet.ListIndex > 0 Then
        cCadena = " mes='" & cboMesDet.List(cboMesDet.ListIndex, 1) & "' and "
    Else
        cCadena = ""
    End If
    
    SQL = "select * from rh_tareas_detalle where " & cCadena & " anio='" & strAnoSistema & "' and codemp='" & oFlxGrid.TextMatrix(oFlxGrid.row, 9) & "' and item=" & oFlxGrid.TextMatrix(oFlxGrid.row, 8)
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    I = 1
    

    
    Dim nPorcentaje As Double
    Dim cDia As String
    
    Do While Not RQ.EOF()
        With flxDetalle
            
            cDia = Left(UCase(WeekdayName(Weekday(Trim(RQ.Fields("fecha"))))), 3)
            cDia = Replace(cDia, "É", "E")
            cDia = Replace(cDia, "Á", "A")
            
            .TextMatrix(I, 0) = I
            .TextMatrix(I, 1) = cDia
            .TextMatrix(I, 2) = Trim(RQ.Fields("codemp"))
            .TextMatrix(I, 3) = Trim(RQ.Fields("item"))
            .TextMatrix(I, 4) = Trim(RQ.Fields("anio"))
            .TextMatrix(I, 5) = Trim(RQ.Fields("mes"))
            .TextMatrix(I, 6) = Trim(RQ.Fields("fecha"))
            .TextMatrix(I, 7) = CE(NumeroSemana(RQ.Fields("fecha"))) & "° SEM"
            .TextMatrix(I, 8) = NE(RQ.Fields("porcentaje"))
            
            .TextMatrix(I, 9) = IIf(NE(RQ.Fields("porcentaje")) = 100, "OK", "")
            .TextMatrix(I, 10) = Trim(RQ.Fields("item_det"))
            
            .row = I
            .Col = 1: .CellBackColor = ColorHabilitado
            .Col = 6: .CellBackColor = ColorHabilitado
            .Col = 7: .CellBackColor = ColorHabilitado
            
            Call ColorDetalle(I)
            
            I = I + 1
        End With
        
        RQ.MoveNext
        If Not (RQ.EOF) Then flxDetalle.Rows = flxDetalle.Rows + 1
    
    Loop
    
    flxDetalle.Visible = True
End Sub



Private Function NumeroSemana(cFecha As String) As Integer
    Dim cMes As String
    Dim cAnio As String
    Dim cDia As String
    Dim cNomDia As String
    Dim cCad As Double
    
    NumeroSemana = 1
    
    cDia = Left(cFecha, 2)
    cMes = Mid(cFecha, 4, 2)
    cAnio = Right(cFecha, 4)
    
    For cCad = NE(cAnio & cMes & "01") To NE(cAnio & cMes & cDia)

        cNomDia = Left(UCase(WeekdayName(Weekday(Right(cCad, 2) & "/" & Mid(cCad, 5, 2) & "/" & Left(cCad, 4)))), 3)

        If cNomDia = "DOM" Then
            NumeroSemana = NumeroSemana + 1
        End If
    Next cCad
End Function

Private Sub AsignaIndexCombo(nIndex)
    On Error Resume Next
    cboMesDet.ListIndex = nIndex
End Sub

Private Sub LlenarGridObjeto(ByRef oGrid As flxEditfac)
    Dim nTipodeTarea As Integer
    
    
    Select Case oGrid.Nombre
        Case "flxtareas": nTipodeTarea = 1
        Case "flxtareasSemanal": nTipodeTarea = 2
        Case "flxtareasMensual": nTipodeTarea = 3
        Case "flxtareasTrimestral": nTipodeTarea = 4
        Case "flxtareasSemestral": nTipodeTarea = 5
        Case "flxtareasAnual": nTipodeTarea = 6
        Case "flxtareasEventual": nTipodeTarea = 7
    End Select
    
    Dim SQL As String
    
    Dim RQ As New ADODB.Recordset
    
    Dim I As Integer, MesFin As String
    Dim Str1 As String, str2 As String, Str3 As String, Str4 As String, Str5 As String, Str6 As String
    
    
    Dim cMes As String
    
    cMes = Right("00" & cboMesDet.ListIndex, 2)
    
    If NE(cMes) <= 0 Then cMes = ""
    
    SQL = "call RH_SP_Tareas_Empleado('CUADRO_RESUMEN','" & strAnoSistema & "','" & cboemp.List(cboemp.ListIndex, 1) & "','" & nTipodeTarea & "','" & cboprioridad.List(cboprioridad.ListIndex, 1) & "','','','" & cMes & "','');"
    

    
    Set RQ = ADO_LlenaRs(SQL)
          
    I = 1
    
    Dim nPorcentaje As Double
    Dim J As Integer
    
    
    If RQ.State = adStateOpen Then
        If Not RQ Is Nothing And Not (RQ.EOF And RQ.BOF) Then
        
            Do While Not RQ.EOF()
                With oGrid
                    .TextMatrix(I, 1) = Trim(RQ.Fields("nombres"))
                    .TextMatrix(I, 2) = (Trim(RQ.Fields("tipotarea")))
                    .TextMatrix(I, 3) = Trim(RQ.Fields("tarea"))
                    .TextMatrix(I, 4) = Trim(RQ.Fields("desprioridad")) 'desprioridad
                    .TextMatrix(I, 5) = Trim(RQ.Fields("dia"))
                    .TextMatrix(I, 6) = IIf(val(RQ.Fields("numdia")) = 0, "", RQ.Fields("numdia"))
                    Select Case RQ.Fields("codtipotarea")
                        Case 4
                            If InStr(1, Trim(RQ.Fields("mes")), "/") > 0 Then
                                .TextMatrix(I, 7) = IIf(Trim(NombreMes(Mid(Trim(RQ.Fields("mes")), 1, 2), True)) = "", "", Trim(NombreMes(Mid(Trim(RQ.Fields("mes")), 1, 2), True)) & "/") & _
                                                    IIf(Trim(NombreMes(Mid(Trim(RQ.Fields("mes")), 4, 2), True)) = "", "", Trim(NombreMes(Mid(Trim(RQ.Fields("mes")), 4, 2), True)) & "/") & _
                                                    IIf(Trim(NombreMes(Mid(Trim(RQ.Fields("mes")), 7, 2), True)) = "", "", Trim(NombreMes(Mid(Trim(RQ.Fields("mes")), 7, 2), True)) & "/") & _
                                                    Trim(NombreMes(Mid(Trim(RQ.Fields("mes")), 10, 2), True))
                            Else
                                .TextMatrix(I, 7) = Trim(IIf(RQ.Fields("MES") = "", "", NombreMes(RQ.Fields("mes"), False)))
                            End If
                        Case 5
                            If InStr(1, Trim(RQ.Fields("mes")), "/") > 0 Then
                                MesFin = Trim(NombreMes(Mid(Trim(RQ.Fields("mes")), 4, 2), True))
                                .TextMatrix(I, 7) = Trim(IIf(RQ.Fields("MES") = "", "", NombreMes(Mid(RQ.Fields("mes"), 1, 2), True))) & "/" & MesFin
                            Else
                                .TextMatrix(I, 7) = Trim(IIf(RQ.Fields("MES") = "", "", NombreMes(RQ.Fields("mes"), False)))
                            End If
                        Case Else
                            .TextMatrix(I, 7) = Trim(IIf(RQ.Fields("MES") = "", "", NombreMes(RQ.Fields("mes"), False)))
                    End Select
                    
                    DoEvents
                    
                    .TextMatrix(I, 8) = NE(RQ.Fields("item"))
                    .TextMatrix(I, 9) = CE(RQ.Fields("codigo"))
                    .TextMatrix(I, 10) = CE(RQ.Fields("codprioridad")) 'codprioridad
                    .TextMatrix(I, 11) = CE(RQ.Fields("mes"))
                    .TextMatrix(I, 12) = CE(RQ.Fields("codtipotarea"))
                    
                    
                    nPorcentaje = 0 ' NE(Trim(RQ.Fields("porcentaje")))
                    
                    .TextMatrix(I, 13) = IIf(nPorcentaje > 100, 100, nPorcentaje)
                    
                    .Col = 1: .row = I: .CellBackColor = ColorDeshabilitado
                    
                    cNombreGridActivo = oGrid.Nombre
                    
                    Call HabilitarCeldas(oGrid, I)
                    
                    
                    .Col = 4
                    Select Case .TextMatrix(I, 4)
                        
                        Case "ALTA"
                            .CellBackColor = vbRed
                            .CellForeColor = vbBlack
                        Case "MEDIA"
                            .CellBackColor = vbYellow
                            .CellForeColor = vbBlack
                        Case "BAJA"
                            .CellBackColor = vbGreen
                            .CellForeColor = vbBlack
                        Case Else
                            .CellForeColor = vbBlack
                            .CellBackColor = vbWhite
                    End Select
                    
                    .Rows = .Rows + 1
                    I = I + 1
                    .RowHeight(I) = 300
                End With
                RQ.MoveNext
            Loop
            oGrid.Visible = True
            
            lblTipo(nTipodeTarea - 1).Visible = True
            
            nUbicacionesPicFondo(nTipodeTarea - 1) = 1
        Else
            oGrid.Visible = False
            lblTipo(nTipodeTarea - 1).Visible = False
            nUbicacionesPicFondo(nTipodeTarea - 1) = 0
            
        End If
    Else
    
        oGrid.Visible = bAutorizado
        
        oGrid.FormColor = vbBlack
        
        lblTipo(nTipodeTarea - 1).Visible = bAutorizado
        
        nUbicacionesPicFondo(nTipodeTarea - 1) = IIf(bAutorizado = True, 1, 0)
    End If
    
    Set RQ = Nothing
    
    If bAutorizado = False Then
        oGrid.Rows = oGrid.Rows - 1
    End If
    
    Call EnumerarItems2(oGrid)
    
    
End Sub

Private Sub PosicionPaneles()
    Dim nControles As Integer, nPos As Integer
    
    For nControles = 0 To 6
    
        If nUbicacionesPicFondo(nControles) = 0 Then
    
    
            For nPos = nControles To 6
                If nUbicacionesPicFondo(nPos) = 1 Then
                
                    picFondo(nPos).Top = nControles * 1980 + 135
                
                    nUbicacionesPicFondo(nControles) = 1
                    nUbicacionesPicFondo(nPos) = 0
                
                    
                    Exit For
                End If
            Next nPos
        
        
        End If
        
    Next nControles
    
End Sub


Private Sub LlenarGrid()

    Call FormatoGridPersonalizado(flxtareas, False)
    
    Call FormatoGridPersonalizado(flxtareasSemanal, False)
    Call FormatoGridPersonalizado(flxtareasMensual, False)
    Call FormatoGridPersonalizado(flxtareasSemestral, False)
    Call FormatoGridPersonalizado(flxtareasTrimestral, False)
    Call FormatoGridPersonalizado(flxtareasAnual, False)
    Call FormatoGridPersonalizado(flxtareasEventual, False)

    If cboemp.ListIndex = 0 Then
        Exit Sub
    End If
    

    bTermino = False
    
    Call LlenarGridObjeto(flxtareas)
    Call LlenarGridObjeto(flxtareasSemanal)
    Call LlenarGridObjeto(flxtareasMensual)
    Call LlenarGridObjeto(flxtareasTrimestral)
    Call LlenarGridObjeto(flxtareasSemestral)
    Call LlenarGridObjeto(flxtareasAnual)
    Call LlenarGridObjeto(flxtareasEventual)
    
    Call PosicionPaneles
    DoEvents
    bTermino = True
End Sub


Private Sub chBtnEliminar_Click()
    Dim fila As Integer
    Dim cFecha As String
    
    For fila = flxDetalle.Rows - 1 To flxDetalle.row Step -1
          If fila <= flxDetalle.Rowsel Then
          
            Call flxDetalle.RemoveItem(fila)
          
            
          End If
    Next fila
End Sub

Private Sub chBtnGrabar_Click()
    Screen.MousePointer = vbHourglass
    DoEvents
    Call GrabaDetalleGrid
    Screen.MousePointer = vbNormal
End Sub

Private Sub GrabaDetalleGrid()
    Select Case cNombreGridActivo
        Case "flxtareas": Call GrabarDetalle(flxtareas)
        Case "flxtareasSemanal": Call GrabarDetalle(flxtareasSemanal)
        Case "flxtareasMensual": Call GrabarDetalle(flxtareasMensual)
        Case "flxtareasTrimestral": Call GrabarDetalle(flxtareasTrimestral)
        Case "flxtareasSemestral": Call GrabarDetalle(flxtareasSemestral)
        Case "flxtareasAnual": Call GrabarDetalle(flxtareasAnual)
        Case "flxtareasEventual": Call GrabarDetalle(flxtareasEventual)
    End Select

End Sub

Private Sub ActualizaDetalleGrid()
    Select Case cNombreGridActivo
        Case "flxtareas":  Call ActualizaDetalle(flxtareas)
        Case "flxtareasSemanal": Call ActualizaDetalle(flxtareasSemanal)
        Case "flxtareasMensual": Call ActualizaDetalle(flxtareasMensual)
        Case "flxtareasTrimestral": Call ActualizaDetalle(flxtareasTrimestral)
        Case "flxtareasSemestral": Call ActualizaDetalle(flxtareasSemestral)
        Case "flxtareasAnual": Call ActualizaDetalle(flxtareasAnual)
        Case "flxtareasEventual": Call ActualizaDetalle(flxtareasEventual)
    End Select

End Sub

Private Sub ActualizaDetalle(ByRef oGrid As flxEditfac)
    Dim fila As Long
    Dim SQL As String, cCadena As String
    Dim nPorcentaje As Double
    
    
    With flxDetalle
    fila = .row
    nPorcentaje = NE(.TextMatrix(fila, 8))
    SQL = "update rh_tareas_detalle set porcentaje = " & nPorcentaje & " " & _
          "where codemp='" & CE(.TextMatrix(fila, 2)) & "' and item=" & CE(.TextMatrix(fila, 3)) & " and anio='" & CE(.TextMatrix(fila, 4)) & "' and mes='" & CE(.TextMatrix(fila, 5)) & "' and item_det=" & CE(.TextMatrix(fila, 10))
          
    Call oConexionSQL.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Modificar, False)
    
    
    End With

    Call RefrescarxGrilla
End Sub

Private Sub GrabarDetalle(ByRef oGrid As flxEditfac)
    Dim fila As Long
    Dim SQL As String, cCadena As String
    Dim nPorcentaje As Double
    
    With oGrid
        If CE(.TextMatrix(flxtareas.row, 12)) = "1" Or CE(.TextMatrix(flxtareas.row, 12)) = "2" Then
            cCadena = " mes='" & flxDetalle.TextMatrix(flxDetalle.row, 5) & "' and "
        Else
            cCadena = ""
        End If
        
        SQL = "delete from rh_tareas_detalle where " & cCadena & " anio='" & strAnoSistema & "' and codemp = '" & flxDetalle.TextMatrix(flxDetalle.row, 2) & "' and item = " & flxDetalle.TextMatrix(flxDetalle.row, 3) & ""
    End With
    
    With flxDetalle
   
        If oConexionSQL.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Eliminar, False) Then
            
                For fila = 1 To .Rows - 1
                    nPorcentaje = NE(.TextMatrix(fila, 8))
                    
                    If nPorcentaje < 0 Then nPorcentaje = 0
                    If nPorcentaje > 100 Then nPorcentaje = 100
                    
                    SQL = "insert into rh_tareas_detalle(codemp,item,anio,mes,item_det,fecha, porcentaje) " & _
                          "values ( '" & CE(.TextMatrix(fila, 2)) & "'," & CE(.TextMatrix(fila, 3)) & ",'" & CE(.TextMatrix(fila, 4)) & "','" & CE(.TextMatrix(fila, 5)) & "', " & _
                          CE(.TextMatrix(fila, 0)) & ",'" & CE(.TextMatrix(fila, 6)) & "'," & nPorcentaje & ")"
                          
                    If oConexionSQL.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, False) = True Then
                        Call EnviarMensaje
                        
                        
                        
                    End If
                    
    
                    
                Next fila
            
        End If
    End With

    Call RefrescarxGrilla
End Sub



Private Sub chBtnRefreshDet_Click()
    Screen.MousePointer = vbHourglass
    DoEvents
    Call RefrescarxGrilla
    Screen.MousePointer = vbNormal

    
End Sub



Private Sub chBtnReporte_Click()

    Dim Str1 As String, str2 As String, Str3 As String, Str4 As String, Str5 As String, Str6 As String
    If cboemp.ListIndex > 0 Then Str1 = cboemp.List(cboemp.ListIndex, 1)
'    If cbotarea.ListIndex > 0 Then str2 = cbotarea.List(cbotarea.ListIndex, 1)
    If cboprioridad.ListIndex > 0 Then Str3 = cboprioridad.List(cboprioridad.ListIndex, 1)
    If cbodia.ListIndex > 0 Then Str4 = cbodia.List(cbodia.ListIndex, 1)
    If cbonumdia.ListIndex > 0 Then Str5 = cbonumdia.List(cbonumdia.ListIndex, 1)
    If cbomes.ListIndex > 0 Then Str6 = cbomes.List(cbomes.ListIndex, 1)
    Set oReporte = New clsReporte
    oReporte.Reporte = "Rep_TareasxTrab.rpt"
    oReporte.empresa = strNombreEmpresa
    'If cbotarea.ListIndex > 0 Then
    '    oReporte.Titulo = "PROGRAMACIÓN DE TAREAS " & cbotarea.List(cbotarea.ListIndex, 2)
    'Else
        oReporte.Titulo = "PROGRAMACIÓN DE TAREAS"
    'End If
    oReporte.TareasProgramadas Str1, str2, val(Str3), Str4, Str5, Str6
End Sub


Private Sub chk_Click()
    If chk.Value = True Then FlgCerrar = True Else FlgCerrar = False
End Sub

Private Sub cmdGrabar_Click()

        
    Dim SQL As String
    SQL = "delete from rh_descripfun where codemp = '" & cboemp.List(cboemp.ListIndex, 1) & "'"
    oConexionMYSQL.Execute SQL
    SQL = "insert into rh_descripfun(codemp,posicion,descripcion) values('" & cboemp.List(cboemp.ListIndex, 1) & "', " & _
          "'" & Trim(txtposicion.Text) & "','" & Trim(txtfunciones.Text) & "')"
    oConexionMYSQL.Execute SQL
End Sub

Private Sub flxDetalle_DblClick()
    Screen.MousePointer = vbHourglass
    DoEvents
    
    With flxDetalle
        If .TextMatrix(.row, 0) <> "" Then
            If .TextMatrix(.row, 8) = "100" Then
                .TextMatrix(.row, 8) = "0"
            Else
                .TextMatrix(.row, 8) = "100"
            End If
        End If
        
        
        Call ColorDetalle(.row)
        nFilaDet = .row
        Call ActualizaDetalleGrid
    End With

    Screen.MousePointer = vbNormal
End Sub

Private Sub flxDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxDetalle
        Select Case .Col
            Case 8:
                    If NE(.TextMatrix(.row, .Col)) > 100 Then
                        Mensajes "El porcentaje no debe pasar de 100%"
                        .TextMatrix(.row, .Col) = "100"
                    End If
        End Select
    End With

End Sub

Private Sub flxDetalle_RowColChange()
    With flxDetalle
        Select Case .Col
            Case 8:
                If FlgIngresar = False Then
                    Publimensaje = "modificar"
                End If
        End Select
    End With

End Sub



Private Sub DobleClickGrid(ByRef oGrid As flxEditfac)
    With oGrid
        If bAutorizado = False Then
            Mensajes "El usuario " & strUsuarioId & " no esta autorizado para modificar las tareas registradas"
            .NoEditar = True
            Exit Sub
        Else
            .NoEditar = False
        End If
        
       
        If .CellBackColor <> ColorDeshabilitado Then
            Select Case .Col
                Case 1, 2, 4
                    Publimensaje = "sin-editar"
                    LlenarCombo .Col
                Case 5
                    Publimensaje = "sin-editar"
                    dias
                Case 6
                    Publimensaje = "sin-editar"
                    NumDias cboedit, False
                Case 7
                    Publimensaje = "sin-editar"
                    Meses
                Case Else
                    oGrid.NoEditar = False
                    
                    Call HabilitarCeldas(oGrid)
                    Exit Sub
            End Select
            
            cboedit.Top = .CellTop + RetornaTop(oGrid)
            cboedit.Left = .CellLeft + RetornaLeft(oGrid) + 100
            cboedit.Width = .CellWidth + 600
            cboedit.Visible = True
            cboedit.SetFocus
            
            DoEvents
            
            
        End If
    End With
End Sub

Private Sub flxtareasEventual_Click()
    cNombreGridActivo = "flxtareasEventual"
    cboedit.Visible = False
    Call VisibleDetalle(True)
    Call RefrescarxGrilla
End Sub

Private Sub flxtareas_Click()
    cNombreGridActivo = "flxtareas"
    cboedit.Visible = False
    Call VisibleDetalle(True)
    Call RefrescarxGrilla
End Sub

Private Sub flxtareasAnual_Click()
    cNombreGridActivo = "flxtareasAnual"
    cboedit.Visible = False
    Call VisibleDetalle(False)
    Call RefrescarxGrilla
End Sub

Private Sub flxtareasMensual_Click()
    cNombreGridActivo = "flxtareasMensual"
    cboedit.Visible = False
    Call VisibleDetalle(False)
    Call RefrescarxGrilla
End Sub

Private Sub flxtareasSemanal_Click()
    cNombreGridActivo = "flxtareasSemanal"
    cboedit.Visible = False
    Call VisibleDetalle(False)
    Call RefrescarxGrilla
End Sub

Private Sub flxtareasSemestral_Click()
    cNombreGridActivo = "flxtareasSemestral"
    cboedit.Visible = False
    Call VisibleDetalle(False)
    Call RefrescarxGrilla
End Sub

Private Sub flxtareasTrimestral_Click()
    cNombreGridActivo = "flxtareasTrimestral"
    cboedit.Visible = False
    Call VisibleDetalle(False)
    Call RefrescarxGrilla
End Sub

Private Sub flxtareasEventual_DblClick()
    Call DobleClickGrid(flxtareasEventual)
End Sub

Private Sub flxtareas_DblClick()
    Call DobleClickGrid(flxtareas)
End Sub

Private Sub flxtareasSemanal_DblClick()
    Call DobleClickGrid(flxtareasSemanal)
End Sub

Private Sub flxtareasMensual_DblClick()
    Call DobleClickGrid(flxtareasMensual)
End Sub


Private Sub flxtareasTrimestral_DblClick()
    Call DobleClickGrid(flxtareasTrimestral)
End Sub

Private Sub flxtareasSemestral_DblClick()
    Call DobleClickGrid(flxtareasSemestral)
End Sub

Private Sub flxtareasAnual_DblClick()
    Call DobleClickGrid(flxtareasAnual)
End Sub

Sub NumDias(cboedit As Control, op As Boolean)
    cboedit.Clear
    If op = True Then cboedit.AddItem "Seleccionar..."
    cboedit.AddItem "01": cboedit.AddItem "02": cboedit.AddItem "03"
    cboedit.AddItem "04": cboedit.AddItem "05": cboedit.AddItem "06"
    cboedit.AddItem "07": cboedit.AddItem "08": cboedit.AddItem "09"
    cboedit.AddItem "10": cboedit.AddItem "11": cboedit.AddItem "12"
    cboedit.AddItem "13": cboedit.AddItem "14": cboedit.AddItem "15"
    cboedit.AddItem "16": cboedit.AddItem "17": cboedit.AddItem "18"
    cboedit.AddItem "19": cboedit.AddItem "20": cboedit.AddItem "21"
    cboedit.AddItem "22": cboedit.AddItem "23": cboedit.AddItem "24"
    cboedit.AddItem "25": cboedit.AddItem "26": cboedit.AddItem "27"
    cboedit.AddItem "28": cboedit.AddItem "29": cboedit.AddItem "30"
    cboedit.AddItem "31"
    If op = True Then
        cboedit.List(0, 1) = "00": cboedit.List(1, 1) = "01": cboedit.List(2, 1) = "02"
        cboedit.List(3, 1) = "03": cboedit.List(4, 1) = "04": cboedit.List(5, 1) = "05"
        cboedit.List(6, 1) = "06": cboedit.List(7, 1) = "07": cboedit.List(8, 1) = "08"
        cboedit.List(9, 1) = "09": cboedit.List(10, 1) = "10": cboedit.List(11, 1) = "11"
        cboedit.List(12, 1) = "12": cboedit.List(13, 1) = "13": cboedit.List(14, 1) = "14"
        cboedit.List(15, 1) = "15": cboedit.List(16, 1) = "16": cboedit.List(17, 1) = "17"
        cboedit.List(18, 1) = "18": cboedit.List(19, 1) = "19": cboedit.List(20, 1) = "20"
        cboedit.List(21, 1) = "21": cboedit.List(22, 1) = "22": cboedit.List(23, 1) = "23"
        cboedit.List(24, 1) = "24": cboedit.List(25, 1) = "25": cboedit.List(26, 1) = "26"
        cboedit.List(27, 1) = "27": cboedit.List(28, 1) = "28": cboedit.List(29, 1) = "29"
        cboedit.List(30, 1) = "30": cboedit.List(31, 1) = "31"
    End If
    cboedit.ListIndex = -1
End Sub
Sub dias()
    cboedit.Clear
    cboedit.AddItem "LUN"
    cboedit.AddItem "MAR"
    cboedit.AddItem "MIE"
    cboedit.AddItem "JUE"
    cboedit.AddItem "VIE"
    cboedit.AddItem "SAB"
    cboedit.AddItem "DOM"
    cboedit.AddItem "LUN/MAR"
    cboedit.AddItem "LUN/MIE"
    cboedit.AddItem "LUN/JUE"
    cboedit.AddItem "LUN/VIE"
    cboedit.AddItem "MAR/MIE"
    cboedit.AddItem "MAR/JUE"
    cboedit.AddItem "MAR/VIE"
    cboedit.AddItem "MIE/JUE"
    cboedit.AddItem "MIE/VIE"
    cboedit.AddItem "JUE/VIE"
    cboedit.ListIndex = -1
End Sub
Sub Meses()
    Dim I As Integer
    I = 1
    cboedit.Clear
    For I = 1 To 12
        cboedit.AddItem Right("00" & Trim(str(I)), 2) & " - " & UCase(Trim(NombreMes(Right("00" & Trim(str(I)), 2), False)))
    Next
    cboedit.ListIndex = -1
End Sub

Sub LlenarCombo(Opc As Integer)

Dim SQL As String
Dim RQ As MYSQL_RS
    Select Case Opc
        Case 1:
            SQL = "select CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres " & _
                  "from CONTRATO C LEFT JOIN EMPLEADO E ON (C.CODEMP=E.CODIGO) where tipo <> 4 " & _
                  "and situacion = 1 and c.codigo = (select max(codigo) from contrato o where " & _
                  "o.codemp=e.codigo group by codemp) and c.division = '0008' order by nombres"
        Case 2: SQL = "select * from rh_tipotarea"
        Case 4: SQL = "select * from rh_tipoprioridad"
    End Select
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    cboedit.Clear
    Do While Not RQ.EOF
        If Opc = 1 Then
            cboedit.AddItem Trim(RQ.Fields("nombres"))
        Else
            cboedit.AddItem RQ.Fields("codigo") & " - " & Trim(RQ.Fields("descrip"))
        End If
        RQ.MoveNext
    Loop
    Set RQ = Nothing
End Sub


Private Sub KeyDownGrid(ByRef oGrid As flxEditfac, KeyCode As Integer, Shift As Integer)
    Dim cTipoTarea As String
    With oGrid
        If KeyCode = 13 Then
            Select Case .Col
                Case 3, 13:
                    If Trim(.TextMatrix(.row, 8)) <> "" Then GrabarFila .row, oGrid
            End Select
        End If
        If KeyCode = 46 Then
            Dim SQL As String
            Dim cCodigo As String
            cCodigo = .TextMatrix(.row, 9)
            cTipoTarea = .TextMatrix(.row, 2)
            
            If cCodigo = "" Then
                Call LlenarGrid
                Call AgregaUltimoRegistro(oGrid)
                Mensajes "Seleccione el trabajador a eliminar"
                
            Else
                If bAutorizado = False Then
                    KeyCode = 0
                    Mensajes "El usuario " & strUsuarioId & " no esta autorizado para eliminar tareas"
                    Exit Sub
                End If
                
                If MsgBox("Desea eliminar la tarea seleccionada", vbYesNo + vbQuestion) = vbYes Then
                    SQL = "delete from rh_tareasxperiodo where codemp = '" & cCodigo & "' and item = " & Item & ""
                    If oConexionSQL.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Eliminar, False) Then
                    
                        SQL = "delete from rh_tareas_detalle where codemp = '" & cCodigo & "' and item = " & Item & ""
                        If oConexionSQL.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Eliminar, False) Then
                            
                            Call FormatoGridPersonalizado(oGrid, False)
                            Call LlenarGridObjeto(oGrid)
                            Call AgregaUltimoRegistro(oGrid)
                            
                            
                            Call EnviarMensaje("TAREA " & cTipoTarea & " FUE ELIMINADA" & Salto(2) & "POR USUARIO: " & Par_UsuarioID)
                        End If
                    End If
                Else
                    Call FormatoGridPersonalizado(oGrid, False)
                    Call LlenarGridObjeto(oGrid)
                    Call AgregaUltimoRegistro(oGrid)

                End If
            End If
        End If
    End With

End Sub

Private Sub flxtareasEventual_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDownGrid(flxtareasEventual, KeyCode, Shift)
End Sub

Private Sub flxtareas_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDownGrid(flxtareas, KeyCode, Shift)
End Sub

Private Sub flxtareasSemanal_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDownGrid(flxtareasSemanal, KeyCode, Shift)
End Sub

Private Sub flxtareasMensual_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDownGrid(flxtareasMensual, KeyCode, Shift)
End Sub

Private Sub flxtareasTrimestral_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDownGrid(flxtareasTrimestral, KeyCode, Shift)
End Sub

Private Sub flxtareasSemestral_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDownGrid(flxtareasSemestral, KeyCode, Shift)
End Sub

Private Sub flxtareasAnual_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDownGrid(flxtareasAnual, KeyCode, Shift)
End Sub


Private Sub AccionGrilla(ByRef oGrid As flxEditfac)
    With oGrid
        Select Case .Col
            Case 1, 2, 4, 5, 6, 7:
                Publimensaje = "sin-editar"
            Case 3, 13:
                If FlgIngresar = False Then
                    Publimensaje = "modificar"
                End If
        End Select
        Item = val(.TextMatrix(.row, 8))
        If Trim(.TextMatrix(.row, 9)) <> "" Then
            If .Visible = True Then
                DatosEmp Trim(.TextMatrix(.row, 9))
                If FlgIngresar = False Then cmdgrabar.Enabled = True
            Else
                If cboemp.ListIndex <= 0 Then
                    txtfunciones.Text = "": txtposicion.Text = ""
                End If
                If FlgIngresar = False Then cmdgrabar.Enabled = False
            End If
        Else
            If cboemp.ListIndex <= 0 Then
                txtfunciones.Text = "": txtposicion.Text = ""
            End If
            If FlgIngresar = False Then cmdgrabar.Enabled = False
        End If
        
        lblDescripcion.Caption = .TextMatrix(.row, 1) & Salto(1) & " ITEM : " & .TextMatrix(.row, 0) & " , TIPO : " & UCase(.TextMatrix(.row, 2))
    End With

End Sub

Private Sub flxtareasEventual_RowColChange()
    If bTermino = False Then Exit Sub
    
    Call AccionGrilla(flxtareasEventual)
    flxtareasEventual.Nombre = "flxtareasEventual"
    Call LlenarGridDetalle(flxtareasEventual)
End Sub

Private Sub flxtareas_RowColChange()
    If bTermino = False Then Exit Sub
    
    Call AccionGrilla(flxtareas)
    flxtareas.Nombre = "flxtareas"
    Call LlenarGridDetalle(flxtareas)
End Sub

Private Sub flxtareasSemanal_RowColChange()
    If bTermino = False Then Exit Sub
    
    Call AccionGrilla(flxtareasSemanal)
    flxtareasSemanal.Nombre = "flxtareasSemanal"
    Call LlenarGridDetalle(flxtareasSemanal)
End Sub

Private Sub flxtareasMensual_RowColChange()
    If bTermino = False Then Exit Sub
    
    Call AccionGrilla(flxtareasMensual)
    flxtareasMensual.Nombre = "flxtareasMensual"
    Call LlenarGridDetalle(flxtareasMensual)
End Sub

Private Sub flxtareasTrimestral_RowColChange()
    If bTermino = False Then Exit Sub
    
    Call AccionGrilla(flxtareasTrimestral)
    flxtareasTrimestral.Nombre = "flxtareasTrimestral"
    Call LlenarGridDetalle(flxtareasTrimestral)
End Sub

Private Sub flxtareasSemestral_RowColChange()
    If bTermino = False Then Exit Sub
    
    Call AccionGrilla(flxtareasSemestral)
    flxtareasSemestral.Nombre = "flxtareasSemestral"
    Call LlenarGridDetalle(flxtareasSemestral)
End Sub

Private Sub flxtareasAnual_RowColChange()
    If bTermino = False Then Exit Sub
    
    Call AccionGrilla(flxtareasAnual)
    flxtareasAnual.Nombre = "flxtareasAnual"
    Call LlenarGridDetalle(flxtareasAnual)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cboedit.Visible = False
End Sub

Private Sub Form_Load()
    
    sock1.Protocol = sckTCPProtocol
    Me.txtIP.Text = gsServidorSocket
    DoEvents
    Call MostrarTareas
    DoEvents
    bAutorizado = AutorizadoParaTareas()

    flxtareas.Nombre = flxtareas.name
    flxtareasSemanal.Nombre = flxtareasSemanal.name
    flxtareasMensual.Nombre = flxtareasMensual.name
    flxtareasTrimestral.Nombre = flxtareasTrimestral.name
    flxtareasSemestral.Nombre = flxtareasSemestral.name
    flxtareasAnual.Nombre = flxtareasAnual.name
    flxtareasEventual.Nombre = flxtareasEventual.name

    Me.Visible = False
    bTermino = False
    Screen.MousePointer = vbHourglass
    DoEvents
    
    Me.Left = 0
    Me.Top = 0
    
    Me.KeyPreview = True
    
    Call LlenarEmpleados
    Call LlenarMes
    Call LlenarDia
    Call LLenarPrioridad
    Call NumDias(cbonumdia, True)
    Call LlenarGrid
    
    Publimensaje = "sin-editar"
    If cboemp.ListIndex < 1 Then
        txtposicion = "": txtfunciones = ""
    End If
    
    Dim I As Integer, J As Integer
    For I = 1 To cboemp.ListCount - 1
        If strCodEmpleado = cboemp.List(I, 1) Then
            cboemp.ListIndex = I
            Exit For
        End If
    Next
    chk.Visible = True
        
        
    
    If bAutorizado = True Then
        cboemp.Locked = False: cboemp.BackColor = ColorHabilitado
        txtposicion.Locked = False: txtposicion.BackColor = ColorHabilitado
        txtfunciones.Locked = False: txtfunciones.BackColor = ColorHabilitado
        
        cmdgrabar.Enabled = True
        
        FlgIngresar = False
    Else
        cboemp.Locked = True: cboemp.BackColor = ColorDeshabilitado
        txtposicion.Locked = True: txtposicion.BackColor = ColorDeshabilitado
        txtfunciones.Locked = True: txtfunciones.BackColor = ColorDeshabilitado
        
        cmdgrabar.Enabled = False
        
        FlgIngresar = True
    End If
        
    Me.WindowState = vbMaximized
    DoEvents
    Me.Visible = True
    Screen.MousePointer = vbNormal
    
    flxtareas.SelectionMode = flexSelectionFree
    fraBotones.Visible = bAutorizado
    DoEvents
    bTermino = True
    
    '--------------------------------
    Call bntConnect_Click
    DoEvents
    'Call EnviarMensaje
    
    Timer1.Enabled = True
End Sub

Sub DatosEmp(CodEmp As String)
Dim SQL As String, RQ As MYSQL_RS
SQL = "select ifnull(posicion,'') as posicion,ifnull(descripcion,'') as descripcion from rh_descripfun where codemp = '" & CodEmp & "'"
Set RQ = oConexion.EjecutaSelectRS(SQL)
If Not RQ.EOF() Then
    txtposicion.Text = Trim(RQ.Fields("posicion"))
    txtfunciones.Text = Trim(RQ.Fields("descripcion"))
Else
    txtposicion.Text = ""
    txtfunciones.Text = ""
End If
Set RQ = Nothing
End Sub

Sub LlenarEmpleados()
    Dim SQL As String, I As Integer
    Dim RQ As MYSQL_RS
    SQL = "select e.codigo,CONCAT_WS(' ',apepat,apemat,nombre1,nombre2) as nombres " & _
          "from CONTRATO C LEFT JOIN EMPLEADO E ON (C.CODEMP=E.CODIGO) where tipo <> 4 " & _
          "and situacion = 1 and c.codigo = (select max(codigo) from contrato o where " & _
          "o.codemp=e.codigo group by codemp) and ( c.division = '0008' or c.division = '0000'  ) order by nombres"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    cboemp.Clear
    cboemp.AddItem "Seleccionar..."
    cboemp.List(0, 1) = ""
    I = 1
    Do While Not RQ.EOF()
        cboemp.AddItem RQ.Fields("nombres")
        cboemp.List(I, 1) = RQ.Fields("codigo")
        RQ.MoveNext
        I = I + 1
    Loop
    cboemp.ListIndex = 0
    Set RQ = Nothing
End Sub

Sub LLenarTipoTarea(cbotarea As Control)
    Dim SQL As String, I As Integer
    Dim RQ As MYSQL_RS
    SQL = "select * from rh_tipotarea order by codigo"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    cbotarea.Clear
    cbotarea.AddItem "Seleccionar..."
    cbotarea.List(0, 1) = ""
    I = 1
    Do While Not RQ.EOF()
        cbotarea.AddItem Trim(RQ.Fields("descrip"))
        cbotarea.List(I, 1) = RQ.Fields("codigo")
        cbotarea.List(I, 2) = Trim(RQ.Fields("titulo"))
        RQ.MoveNext
        I = I + 1
    Loop
    cbotarea.ListIndex = 0
    Set RQ = Nothing
End Sub

Public Sub LlenarMes()
    Dim I As Integer, mesactual As Integer
    I = 1
    
    
    With cbomes
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "00"
        For I = 1 To 12
            .AddItem UCase(NombreMes(Right("00" & Trim(str(I)), 2), False))
            .List(I, 1) = Right("00" & Trim(str(I)), 2)
            If Month(Date) = I Then
                mesactual = I
            End If
        Next
        .ListIndex = 0 'mesactual
    End With

    I = 1
    
    ' MES DE LA TAREA
    With cboMesDet
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "00"
        For I = 1 To 12
            .AddItem UCase(NombreMes(Right("00" & Trim(str(I)), 2), False))
            .List(I, 1) = Right("00" & Trim(str(I)), 2)
            If Month(Date) = I Then
                mesactual = I
            End If
        Next
        .ListIndex = mesactual
    End With
    
End Sub

Public Sub LlenarDia()
    cbodia.Clear
    cbodia.AddItem "Seleccionar..."
    cbodia.List(0, 1) = "00"
    cbodia.AddItem "LUNES"
    cbodia.List(1, 1) = "LUN"
    cbodia.AddItem "MARTES"
    cbodia.List(2, 1) = "MAR"
    cbodia.AddItem "MIERCOLES"
    cbodia.List(3, 1) = "MIE"
    cbodia.AddItem "JUEVES"
    cbodia.List(4, 1) = "JUE"
    cbodia.AddItem "VIERNES"
    cbodia.List(5, 1) = "VIE"
    cbodia.AddItem "SABADO"
    cbodia.List(6, 1) = "SAB"
    cbodia.AddItem "DOMINGO"
    cbodia.List(7, 1) = "DOM"
    cbodia.AddItem "LUN/MAR"
    cbodia.List(8, 1) = "LUN/MAR"
    cbodia.AddItem "LUN/MIE"
    cbodia.List(9, 1) = "LUN/MIE"
    cbodia.AddItem "LUN/JUE"
    cbodia.List(10, 1) = "LUN/JUE"
    cbodia.AddItem "LUN/VIE"
    cbodia.List(11, 1) = "LUN/VIE"
    cbodia.AddItem "MAR/MIE"
    cbodia.List(12, 1) = "MAR/MIE"
    cbodia.AddItem "MAR/JUE"
    cbodia.List(13, 1) = "MAR/JUE"
    cbodia.AddItem "MAR/VIE"
    cbodia.List(14, 1) = "MAR/VIE"
    cbodia.AddItem "MIE/JUE"
    cbodia.List(15, 1) = "MIE/JUE"
    cbodia.AddItem "MIE/VIE"
    cbodia.List(16, 1) = "MIE/VIE"
    cbodia.AddItem "JUE/VIE"
    cbodia.List(17, 1) = "JUE/VIE"
    cbodia.ListIndex = 0
End Sub
Sub LLenarPrioridad()

    '
    Dim SQL As String, I As Integer
    Dim RQ As MYSQL_RS
    SQL = "select * from rh_tipoprioridad order by codigo"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    cboprioridad.Clear
    cboprioridad.AddItem "Seleccionar..."
    cboprioridad.List(0, 1) = ""
    I = 1
    Do While Not RQ.EOF()
        cboprioridad.AddItem RQ.Fields("descrip")
        cboprioridad.List(I, 1) = RQ.Fields("codigo")
        RQ.MoveNext
        I = I + 1
    Loop
    cboprioridad.ListIndex = 0
    Set RQ = Nothing

End Sub

Private Sub Form_Resize()
    On Error GoTo SERROR
    Me.Caption = ""
    If Me.WindowState <> vbMinimized Then
    
        If fraDesEmpleado.Visible = True Then
            picWindow.Width = Me.Width - fraDesEmpleado.Width - 250
        Else
            picWindow.Width = Me.Width - 250
        End If
    
        'picWindow.Width = Me.Width - fraDesEmpleado.Width - 250
        
        
        picWindow.Height = Me.Height - picWindow.Top - 500
        
        
        With VScroll1
            .Height = picWindow.Height - 250 - 10
            .Top = 0
            .Left = picWindow.Left + picWindow.Width - VScroll1.Width
            .Max = picDatos.Height - picWindow.Height
            .SmallChange = 100
            .LargeChange = picWindow.Height
        End With
        
        With HScroll1
            .Width = picWindow.Width - VScroll1.Width + 50
            .Left = 0
            .Top = picWindow.Top + picWindow.Height - HScroll1.Height - 1400 - 30
            .Max = Abs(picDatos.Width - picWindow.Width) - 3000
            .SmallChange = 100
            .LargeChange = picWindow.Width
        End With
        
        picDatos.Width = picWindow.Width - VScroll1.Width - 50
        
        '-----------------------
        
        flxtareas.Width = picDatos.Width - flxtareas.Left - 100
        flxtareas.Height = 1635
        
        
        flxtareasSemanal.Width = flxtareas.Width
        flxtareasMensual.Width = flxtareas.Width
        flxtareasSemestral.Width = flxtareas.Width
        flxtareasTrimestral.Width = flxtareas.Width
        flxtareasAnual.Width = flxtareas.Width
        flxtareasEventual.Width = flxtareas.Width
        
        flxtareasSemanal.Left = flxtareas.Left
        flxtareasMensual.Left = flxtareas.Left
        flxtareasSemestral.Left = flxtareas.Left
        flxtareasTrimestral.Left = flxtareas.Left
        flxtareasAnual.Left = flxtareas.Left
        flxtareasEventual.Left = flxtareasEventual.Left
        
        '-----------------------
        
        fraDesEmpleado.Top = picWindow.Top - 100
        fraDesEmpleado.Left = picWindow.Left + picWindow.Width + 50
        
        
        flxDetalle.Width = fraBotones.Width
        flxDetalle.Left = fraDesEmpleado.Left
        flxDetalle.Top = picWindow.Top + fraDesEmpleado.Height
        flxDetalle.Height = Me.Height - picWindow.Top - 500 - fraDesEmpleado.Height - 700
        
        
        fraBotones.Width = fraDesEmpleado.Width
        fraBotones.Left = flxDetalle.Left
        fraBotones.Top = flxDetalle.Top + flxDetalle.Height - fraBotones.Height + 700
        
        picFondo(0).Width = picDatos.Width
        picFondo(1).Width = picDatos.Width
        picFondo(2).Width = picDatos.Width
        picFondo(3).Width = picDatos.Width
        picFondo(4).Width = picDatos.Width
        picFondo(5).Width = picDatos.Width
        picFondo(6).Width = picDatos.Width
        
        
        
    End If
    
    Exit Sub
SERROR:
    
End Sub




Private Sub VisibleDetalle(bValor As Boolean)
    bValor = True

    fraDesEmpleado.Visible = bValor
    cboMesDet.Visible = bValor
    flxDetalle.Visible = bValor
    fraBotones.Visible = bValor
    
    
    Call Form_Resize
    
End Sub

Private Sub sock1_ConnectionRequest(ByVal requestID As Long)
requestID = requestID
End Sub

Private Sub Timer1_Timer()
    Call EnviarMensaje
    Timer1.Interval = 10000
    Timer1.Enabled = False
End Sub

Private Sub VScroll1_Change()
    picDatos.Top = -(VScroll1.Value)
End Sub


Private Sub HScroll1_Change()
    picDatos.Left = Abs(HScroll1.Value) * -1
End Sub






Private Sub sock1_Close()
    'handles the closing of the connection
    
    sock1.Close  'close connection
    
    txtLog = txtLog & "*** Disconnected" & vbCrLf

End Sub

Private Sub sock1_Connect()
    'txtLog is the textbox used as our
    'chat buffer.
    
    'sock1.RemoteHost returns the hostname( or ip ) of the host
    'sock1.RemoteHostIP returns the IP of the host
    
    txtLog = "Connected to " & sock1.RemoteHostIP & vbCrLf

End Sub

Private Sub sock1_DataArrival(ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

Dim dat As String     'where to put the data

sock1.GetData dat, vbString   'writes the new data in our string dat ( string format )

'add the new message to our chat buffer
txtLog = txtLog & dat & vbCrLf

If dat = "MAXIMIZAR" Then
    Call mPopRestore_Click
End If

End Sub

Private Sub sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this event is to handle any kind of errors
'happend while using winsock

'Number gives you the number code of that specific error
'Description gives you string with a simple explanation about the error

'append the error message in the chat buffer
txtLog = txtLog & "*** Error : " & Number & " ," & Description & vbCrLf

'and now we need to close the connection
sock1_Close

'you could also use sock1.close function but I
'prefer to call it within the Sock1_Close functions that
'handles the connection closing in general

End Sub


Private Sub mPopRestore_Click()
 'called when the user clicks the popup menu Restore command
 Dim Result As Long
 Me.WindowState = vbMaximized
 Result = SetForegroundWindow(Me.hwnd)
 Me.Show
End Sub

Private Sub mPopMinimized_Click()
 Dim Result As Long
 Me.WindowState = vbMinimized
 Result = SetForegroundWindow(Me.hwnd)
 Me.Show

End Sub

