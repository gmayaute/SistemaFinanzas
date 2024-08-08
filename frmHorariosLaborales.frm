VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmHorariosLaborales 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Horarios Laborales"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frmHorariosLaborales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox RegAleatorio 
      BackColor       =   &H009F5539&
      Caption         =   "Registro aleatorio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   2520
      TabIndex        =   70
      Top             =   3420
      Width           =   1875
   End
   Begin VB.CheckBox RegAutomatico 
      BackColor       =   &H009F5539&
      Caption         =   "Registro automático"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   270
      TabIndex        =   69
      Top             =   3420
      Width           =   2175
   End
   Begin VB.ComboBox cmbHorarios 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2460
      TabIndex        =   61
      Text            =   "Horarios"
      Top             =   30
      Width           =   3645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Caption         =   "Días                          Horario de Trabajo                                       Refrigerio"
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
      Height          =   4095
      Left            =   30
      TabIndex        =   21
      Top             =   420
      Width           =   8385
      Begin MSMask.MaskEdBox meTolEnt 
         Height          =   315
         Left            =   2400
         TabIndex        =   56
         Top             =   3330
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   2
         Format          =   "##"
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox RegRefigerio 
         BackColor       =   &H009F5539&
         Caption         =   "Registro automático de refrigerios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   4830
         TabIndex        =   15
         Top             =   2970
         Width           =   3225
      End
      Begin VB.CheckBox Refri 
         BackColor       =   &H009F5539&
         Height          =   255
         Index           =   6
         Left            =   4830
         TabIndex        =   14
         Top             =   2700
         Width           =   255
      End
      Begin VB.CheckBox Refri 
         BackColor       =   &H009F5539&
         Height          =   255
         Index           =   5
         Left            =   4830
         TabIndex        =   13
         Top             =   2340
         Width           =   255
      End
      Begin VB.CheckBox Refri 
         BackColor       =   &H009F5539&
         Height          =   255
         Index           =   4
         Left            =   4830
         TabIndex        =   12
         Top             =   1980
         Width           =   255
      End
      Begin VB.CheckBox Refri 
         BackColor       =   &H009F5539&
         Height          =   255
         Index           =   3
         Left            =   4830
         TabIndex        =   11
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox Refri 
         BackColor       =   &H009F5539&
         Height          =   255
         Index           =   2
         Left            =   4830
         TabIndex        =   10
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox Refri 
         BackColor       =   &H009F5539&
         Height          =   255
         Index           =   1
         Left            =   4830
         TabIndex        =   9
         Top             =   870
         Width           =   255
      End
      Begin VB.CheckBox Refri 
         BackColor       =   &H009F5539&
         Height          =   255
         Index           =   0
         Left            =   4830
         TabIndex        =   8
         Top             =   510
         Width           =   255
      End
      Begin MSComCtl2.DTPicker HorEntrada 
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   24
         Top             =   480
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   16777215
         Format          =   92471298
         CurrentDate     =   39225
      End
      Begin VB.CheckBox chkDias 
         BackColor       =   &H009F5539&
         Caption         =   "Martes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   900
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         BackColor       =   &H009F5539&
         Caption         =   "Domingo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   2730
         Width           =   1065
      End
      Begin VB.CheckBox chkDias 
         BackColor       =   &H009F5539&
         Caption         =   "Sábado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   2340
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         BackColor       =   &H009F5539&
         Caption         =   "Viernes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1980
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         BackColor       =   &H009F5539&
         Caption         =   "Jueves"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1620
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         BackColor       =   &H009F5539&
         Caption         =   "Miércoles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1260
         Width           =   1155
      End
      Begin VB.CheckBox chkDias 
         BackColor       =   &H009F5539&
         Caption         =   "Lunes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   975
      End
      Begin MSComCtl2.DTPicker HorEntrada 
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   25
         Top             =   870
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorEntrada 
         Height          =   285
         Index           =   2
         Left            =   1500
         TabIndex        =   26
         Top             =   1230
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorEntrada 
         Height          =   285
         Index           =   3
         Left            =   1500
         TabIndex        =   27
         Top             =   1590
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorEntrada 
         Height          =   285
         Index           =   4
         Left            =   1500
         TabIndex        =   28
         Top             =   1950
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorEntrada 
         Height          =   285
         Index           =   5
         Left            =   1500
         TabIndex        =   29
         Top             =   2310
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorEntrada 
         Height          =   285
         Index           =   6
         Left            =   1500
         TabIndex        =   30
         Top             =   2670
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorSalida 
         Height          =   285
         Index           =   0
         Left            =   2970
         TabIndex        =   31
         Top             =   510
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39224
      End
      Begin MSComCtl2.DTPicker HorSalida 
         Height          =   285
         Index           =   1
         Left            =   2970
         TabIndex        =   32
         Top             =   870
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorSalida 
         Height          =   285
         Index           =   2
         Left            =   2970
         TabIndex        =   33
         Top             =   1260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorSalida 
         Height          =   285
         Index           =   3
         Left            =   2970
         TabIndex        =   34
         Top             =   1620
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorSalida 
         Height          =   285
         Index           =   4
         Left            =   2970
         TabIndex        =   35
         Top             =   1950
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorSalida 
         Height          =   285
         Index           =   5
         Left            =   2970
         TabIndex        =   36
         Top             =   2310
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker HorSalida 
         Height          =   285
         Index           =   6
         Left            =   2970
         TabIndex        =   37
         Top             =   2670
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefIni 
         Height          =   285
         Index           =   0
         Left            =   5100
         TabIndex        =   38
         Top             =   510
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefIni 
         Height          =   285
         Index           =   1
         Left            =   5100
         TabIndex        =   39
         Top             =   870
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefIni 
         Height          =   285
         Index           =   2
         Left            =   5100
         TabIndex        =   40
         Top             =   1230
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefIni 
         Height          =   285
         Index           =   3
         Left            =   5100
         TabIndex        =   41
         Top             =   1590
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefIni 
         Height          =   285
         Index           =   4
         Left            =   5100
         TabIndex        =   42
         Top             =   1950
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefIni 
         Height          =   285
         Index           =   5
         Left            =   5100
         TabIndex        =   43
         Top             =   2310
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefIni 
         Height          =   285
         Index           =   6
         Left            =   5130
         TabIndex        =   44
         Top             =   2670
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefriFin 
         Height          =   285
         Index           =   0
         Left            =   6570
         TabIndex        =   45
         Top             =   510
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefriFin 
         Height          =   285
         Index           =   1
         Left            =   6570
         TabIndex        =   46
         Top             =   870
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefriFin 
         Height          =   285
         Index           =   2
         Left            =   6570
         TabIndex        =   47
         Top             =   1230
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefriFin 
         Height          =   285
         Index           =   3
         Left            =   6570
         TabIndex        =   48
         Top             =   1590
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefriFin 
         Height          =   285
         Index           =   4
         Left            =   6570
         TabIndex        =   49
         Top             =   1950
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefriFin 
         Height          =   285
         Index           =   5
         Left            =   6570
         TabIndex        =   50
         Top             =   2310
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin MSComCtl2.DTPicker RefriFin 
         Height          =   285
         Index           =   6
         Left            =   6570
         TabIndex        =   51
         Top             =   2670
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   92471298
         CurrentDate     =   39206
      End
      Begin Proyecto1.chameleonButton cmdCopiarH 
         Height          =   285
         Left            =   4380
         TabIndex        =   54
         ToolTipText     =   "Siguiente"
         Top             =   510
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHorariosLaborales.frx":0442
         PICN            =   "frmHorariosLaborales.frx":045E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdCopiarR 
         Height          =   285
         Left            =   7950
         TabIndex        =   55
         ToolTipText     =   "Siguiente"
         Top             =   510
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHorariosLaborales.frx":07C6
         PICN            =   "frmHorariosLaborales.frx":07E2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox meTolSal 
         Height          =   315
         Left            =   2400
         TabIndex        =   57
         Top             =   3690
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   2
         Format          =   "##"
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meTarde 
         Height          =   315
         Left            =   4290
         TabIndex        =   58
         Top             =   3330
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   2
         Format          =   "##"
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meInasistencia 
         Height          =   315
         Left            =   4290
         TabIndex        =   59
         Top             =   3720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   2
         Format          =   "##"
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meHrsxSem 
         Height          =   315
         Left            =   7830
         TabIndex        =   62
         Top             =   3330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Format          =   "##"
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meDT 
         Height          =   315
         Left            =   6150
         TabIndex        =   65
         Top             =   3360
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Format          =   "##"
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meDD 
         Height          =   315
         Left            =   6150
         TabIndex        =   67
         Top             =   3750
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Format          =   "##"
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Descanso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   4800
         TabIndex        =   66
         Top             =   3810
         Width           =   1365
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   4800
         TabIndex        =   64
         Top             =   3420
         Width           =   1395
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs x Semana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   6570
         TabIndex        =   63
         Top             =   3420
         Width           =   1245
      End
      Begin MSForms.CheckBox chkInasistencia 
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Top             =   3720
         Width           =   1365
         BackColor       =   10442041
         ForeColor       =   8438015
         DisplayStyle    =   4
         Size            =   "2408;503"
         Value           =   "0"
         Caption         =   "Inasistencia"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkTarde 
         Height          =   285
         Left            =   2880
         TabIndex        =   17
         Top             =   3360
         Width           =   1125
         BackColor       =   10442041
         ForeColor       =   8438015
         DisplayStyle    =   4
         Size            =   "1984;503"
         Value           =   "0"
         Caption         =   "Tardanza"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         DrawMode        =   15  'Merge Pen Not
         X1              =   750
         X2              =   7500
         Y1              =   3270
         Y2              =   3270
      End
      Begin MSForms.CheckBox chkTolSal 
         Height          =   285
         Left            =   210
         TabIndex        =   18
         Top             =   3720
         Width           =   2175
         BackColor       =   10442041
         ForeColor       =   8454016
         DisplayStyle    =   4
         Size            =   "3836;503"
         Value           =   "0"
         Caption         =   "Tolerancia de Salida"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkTolEnt 
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   3330
         Width           =   2055
         BackColor       =   10442041
         ForeColor       =   8454016
         DisplayStyle    =   4
         Size            =   "3625;503"
         Value           =   "0"
         Caption         =   "Tolerancia Entrada"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6960
         TabIndex        =   53
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5490
         TabIndex        =   52
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   23
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1860
         TabIndex        =   22
         Top             =   270
         Width           =   795
      End
   End
   Begin Proyecto1.chameleonButton cmdNuevo 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Nuevo"
      Top             =   30
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Nuevo"
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
      MICON           =   "frmHorariosLaborales.frx":0B4A
      PICN            =   "frmHorariosLaborales.frx":0B66
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdGrabar 
      Height          =   345
      Left            =   7230
      TabIndex        =   68
      ToolTipText     =   "Guardar"
      Top             =   60
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmHorariosLaborales.frx":0ED0
      PICN            =   "frmHorariosLaborales.frx":0EEC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox txtHorario 
      Height          =   285
      Left            =   2520
      TabIndex        =   60
      Top             =   60
      Visible         =   0   'False
      Width           =   3465
      VariousPropertyBits=   746604571
      Size            =   "6112;503"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      Left            =   1320
      TabIndex        =   20
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "frmHorariosLaborales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CantHorarios As Integer
Sub LimpiaControles()
    Dim I As Integer
    For I = 0 To 6
        HorSalida(I).MinDate = "01/01/1601"
        RefriFin(I).MinDate = "01/01/1601"
        HorEntrada(I).Value = "12:00:00"
        HorSalida(I).Value = "12:00:00"
        RefIni(I).Value = "12:00:00"
        RefriFin(I).Value = "12:00:00"
        chkDias(I).Value = 0
        Refri(I).Value = 0
    Next
    chkInasistencia.Value = 0
    chkTarde.Value = 0
    chkTolEnt.Value = 0
    chkTolSal.Value = 0
    RegRefigerio.Value = 0
    RegAleatorio.Value = 0
    RegAutomatico.Value = 0
    meDD.Text = "__"
    meDT.Text = "__"
    meHrsxSem.Text = "__"
End Sub
Function CargaHorarios() As Integer
    Dim SQL As String, I As Integer
    Dim rsHorarios As MYSQL_RS
    SQL = "Select * from rh_Horarios order by nombre"
    Set rsHorarios = oConexion.EjecutaSelectRS(SQL)
    I = 0
    cmbHorarios.Clear
    With rsHorarios
        Do While Not .EOF
            cmbHorarios.AddItem CE(.Fields("nombre"))
            I = I + 1
            .MoveNext
        Loop
        CargaHorarios = I
    End With
    Set rsHorarios = Nothing
End Function
Sub ActivaDesactivaDias(valor As Boolean)
    Dim I As Integer
    For I = 0 To 6
        chkDias(I).Enabled = valor
        Refri(I).Enabled = valor
    Next
End Sub

Private Sub chkDias_Click(Index As Integer)
    Dim I As Integer
    If chkDias(Index).Value = 1 Then
        HorEntrada(Index).Enabled = True
        HorSalida(Index).Enabled = True
    Else
        HorEntrada(Index).Enabled = False
        HorSalida(Index).Enabled = False
    End If
    cmdCopiarH.Enabled = False
    cmdGrabar.Enabled = False
    For I = 0 To 6
        If chkDias(I).Value = 1 Then
            cmdGrabar.Enabled = True
            cmdCopiarH.Enabled = True
        End If
    Next
End Sub

Private Sub chkInasistencia_Click()
    If chkInasistencia.Value = 0 Then
        meInasistencia.Text = "__"
        meInasistencia.Enabled = False
    Else
        meInasistencia.Enabled = True
        meInasistencia.SetFocus
    End If
End Sub

Private Sub chkTarde_Click()
    If chkTarde.Value = 0 Then
        meTarde.Text = "__"
        meTarde.Enabled = False
    Else
        meTarde.Enabled = True
        meTarde.SetFocus
    End If
End Sub

Private Sub chkTolEnt_Click()
    If chkTolEnt.Value = 0 Then
        meTolEnt.Text = "__"
        meTolEnt.Enabled = False
    Else
        meTolEnt.Enabled = True
        meTolEnt.SetFocus
    End If
End Sub

Private Sub chkTolSal_Click()
    If chkTolSal.Value = 0 Then
        meTolSal.Text = "__"
        meTolSal.Enabled = False
    Else
        meTolSal.Enabled = True
        meTolSal.SetFocus
    End If
End Sub

Private Sub cmbHorarios_Change()
    If txtHorario <> Empty Or cmbHorarios.Visible = True Then
        ActivaDesactivaDias True
        chkInasistencia.Enabled = True
        chkTarde.Enabled = True
        chkTolEnt.Enabled = True
        chkTolSal.Enabled = True
        RegRefigerio.Enabled = True
        RegAleatorio.Enabled = True
        RegAutomatico.Enabled = True
    Else
        ActivaDesactivaDias False
        chkInasistencia.Enabled = False
        chkTarde.Enabled = False
        chkTolEnt.Enabled = False
        chkTolSal.Enabled = False
        RegRefigerio.Enabled = False
        RegAleatorio.Enabled = False
        RegAutomatico.Enabled = False
    End If
    CargarHorarioDetalle cmbHorarios.Text
End Sub
Sub CargarHorarioDetalle(Nombre As String)
    Dim SQL As String, I As Integer
    Dim rsHorarios As MYSQL_RS
    SQL = "Select * from rh_Horarios where nombre='" & Nombre & "'"
    Set rsHorarios = oConexion.EjecutaSelectRS(SQL)
       
    HorEntrada(0).Value = rsHorarios.Fields("LuE")
    HorSalida(0).Value = rsHorarios.Fields("LuS")
    chkDias(0).Value = IIf(HorEntrada(0).Value = "12:00:00 a.m." And HorSalida(0).Value = "12:00:00 a.m.", 0, 1)
    chkDias_Click (0)
    HorEntrada(1).Value = rsHorarios.Fields("MaE")
    HorSalida(1).Value = rsHorarios.Fields("MaS")
    chkDias(1).Value = IIf(HorEntrada(1).Value = "12:00:00 a.m." And HorSalida(1).Value = "12:00:00 a.m.", 0, 1)
    chkDias_Click (1)
    HorEntrada(2).Value = rsHorarios.Fields("MiE")
    HorSalida(2).Value = rsHorarios.Fields("MiS")
    chkDias(2).Value = IIf(HorEntrada(2).Value = "12:00:00 a.m." And HorSalida(2).Value = "12:00:00 a.m.", 0, 1)
    chkDias_Click (2)
    HorEntrada(3).Value = rsHorarios.Fields("JuE")
    HorSalida(3).Value = rsHorarios.Fields("JuS")
    chkDias(3).Value = IIf(HorEntrada(3).Value = "12:00:00 a.m." And HorSalida(3).Value = "12:00:00 a.m.", 0, 1)
    chkDias_Click (3)
    HorEntrada(4).Value = rsHorarios.Fields("ViE")
    HorSalida(4).Value = rsHorarios.Fields("ViS")
    chkDias(4).Value = IIf(HorEntrada(4).Value = "12:00:00 a.m." And HorSalida(4).Value = "12:00:00 a.m.", 0, 1)
    chkDias_Click (4)
    HorEntrada(5).Value = rsHorarios.Fields("SaE")
    HorSalida(5).Value = rsHorarios.Fields("SaS")
    chkDias(5).Value = IIf(HorEntrada(5).Value = "12:00:00 a.m." And HorSalida(5).Value = "12:00:00 a.m.", 0, 1)
    chkDias_Click (5)
    HorEntrada(6).Value = rsHorarios.Fields("DoE")
    HorSalida(6).Value = rsHorarios.Fields("DoS")
    chkDias(6).Value = IIf(HorEntrada(6).Value = "12:00:00 a.m." And HorSalida(6).Value = "12:00:00 a.m.", 0, 1)
    chkDias_Click (6)
    
    RefIni(0).Value = rsHorarios.Fields("LuRE")
    RefriFin(0).Value = rsHorarios.Fields("LuRS")
    Refri(0).Value = IIf(RefIni(0).Value = "12:00:00 a.m." And RefriFin(0).Value = "12:00:00 a.m.", 0, 1)
    Refri_Click (0)
    RefIni(1).Value = rsHorarios.Fields("MaRE")
    RefriFin(1).Value = rsHorarios.Fields("MaRS")
    Refri(1).Value = IIf(RefIni(1).Value = "12:00:00 a.m." And RefriFin(1).Value = "12:00:00 a.m.", 0, 1)
    Refri_Click (1)
    RefIni(2).Value = rsHorarios.Fields("MiRE")
    RefriFin(2).Value = rsHorarios.Fields("MiRS")
    Refri(2).Value = IIf(RefIni(2).Value = "12:00:00 a.m." And RefriFin(2).Value = "12:00:00 a.m.", 0, 1)
    Refri_Click (2)
    RefIni(3).Value = rsHorarios.Fields("JuRE")
    RefriFin(3).Value = rsHorarios.Fields("JuRS")
    Refri(3).Value = IIf(RefIni(3).Value = "12:00:00 a.m." And RefriFin(3).Value = "12:00:00 a.m.", 0, 1)
    Refri_Click (3)
    RefIni(4).Value = rsHorarios.Fields("ViRE")
    RefriFin(4).Value = rsHorarios.Fields("ViRS")
    Refri(4).Value = IIf(RefIni(4).Value = "12:00:00 a.m." And RefriFin(4).Value = "12:00:00 a.m.", 0, 1)
    Refri_Click (4)
    RefIni(5).Value = rsHorarios.Fields("SaRE")
    RefriFin(5).Value = rsHorarios.Fields("SaRS")
    Refri(5).Value = IIf(RefIni(5).Value = "12:00:00 a.m." And RefriFin(5).Value = "12:00:00 a.m.", 0, 1)
    Refri_Click (5)
    RefIni(6).Value = rsHorarios.Fields("DoRE")
    RefriFin(6).Value = rsHorarios.Fields("DoRS")
    Refri(6).Value = IIf(RefIni(6).Value = "12:00:00 a.m." And RefriFin(6).Value = "12:00:00 a.m.", 0, 1)
    Refri_Click (6)
    
    chkTolEnt.Value = IIf(rsHorarios.Fields("TolEntrada") > 0, 1, 0)
    meTolEnt.Text = IIf(rsHorarios.Fields("TolEntrada") > 0, Right("00" & Trim(rsHorarios.Fields("TolEntrada")), 2), "__")
    chkTolSal.Value = IIf(rsHorarios.Fields("TolSalida") > 0, 1, 0)
    meTolSal.Text = IIf(rsHorarios.Fields("TolSalida") > 0, Right("00" & Trim(rsHorarios.Fields("TolSalida")), 2), "__")
    chkTarde.Value = IIf(rsHorarios.Fields("Tarde") > 0, 1, 0)
    meTarde.Text = IIf(rsHorarios.Fields("Tarde") > 0, rsHorarios.Fields("Tarde"), "__")
    chkInasistencia.Value = IIf(rsHorarios.Fields("Inasistencia") > 0, 1, 0)
    meInasistencia.Text = IIf(rsHorarios.Fields("Inasistencia") > 0, rsHorarios.Fields("Inasistencia"), "__")
    meDD.Text = IIf(rsHorarios.Fields("diast") > 0, Right("00" & Trim(str(rsHorarios.Fields("diast"))), 2), "__")
    meDT.Text = IIf(rsHorarios.Fields("diasd") > 0, Right("00" & Trim(str(rsHorarios.Fields("diasd"))), 2), "__")
    meHrsxSem.Text = IIf(rsHorarios.Fields("hrsxsem") > 0, Right("00" & Trim(str(rsHorarios.Fields("hrsxsem"))), 2), "__")
    RegRefigerio.Value = IIf(rsHorarios.Fields("AutoRefri") = "S", 1, 0)
    RegAleatorio.Value = IIf(rsHorarios.Fields("AsisAlea") = "S", 1, 0)
    RegAutomatico.Value = IIf(rsHorarios.Fields("AutoAsi") = "S", 1, 0)
    
End Sub

Function ValidaDatos() As Boolean
    ValidaDatos = True
    If chkInasistencia.Value = True And meInasistencia = "__" Then ValidaDatos = False
    If chkTarde.Value = True And meTarde = "__" Then ValidaDatos = False
    If chkTolEnt.Value = True And meTolEnt = "__" Then ValidaDatos = False
    If chkTolSal.Value = True And chkTolSal = "__" Then ValidaDatos = False
End Function

Function GrabaHorario(Nombre As String) As Boolean
    Dim SQL As String, I As Integer
    Dim rsHor As MYSQL_RS
    SQL = "Select * from rh_Horarios where nombre='" & Nombre & "'"
    Set rsHor = oConexion.EjecutaSelectRS(SQL)
    If rsHor.RecordCount > 0 Then
        SQL = " Update rH_Horarios set " & _
              " LuE='" & Format(HorEntrada(0).Value, "HH:MM:SS") & "',LuS='" & Format(HorSalida(0).Value, "HH:MM:SS") & "'," & _
              " MaE='" & Format(HorEntrada(1).Value, "HH:MM:SS") & "',MaS='" & Format(HorSalida(1).Value, "HH:MM:SS") & "'," & _
              " MiE='" & Format(HorEntrada(2).Value, "HH:MM:SS") & "',MiS='" & Format(HorSalida(2).Value, "HH:MM:SS") & "'," & _
              " JuE='" & Format(HorEntrada(3).Value, "HH:MM:SS") & "',JuS='" & Format(HorSalida(3).Value, "HH:MM:SS") & "'," & _
              " ViE='" & Format(HorEntrada(4).Value, "HH:MM:SS") & "',ViS='" & Format(HorSalida(4).Value, "HH:MM:SS") & "'," & _
              " SaE='" & Format(HorEntrada(5).Value, "HH:MM:SS") & "',SaS='" & Format(HorSalida(5).Value, "HH:MM:SS") & "'," & _
              " DoE='" & Format(HorEntrada(6).Value, "HH:MM:SS") & "',DoS='" & Format(HorSalida(6).Value, "HH:MM:SS") & "'," & _
              " LuRE='" & Format(RefIni(0).Value, "HH:MM:SS") & "',LuRS='" & Format(RefriFin(0).Value, "HH:MM:SS") & "'," & _
              " MaRE='" & Format(RefIni(1).Value, "HH:MM:SS") & "',MaRS='" & Format(RefriFin(1).Value, "HH:MM:SS") & "'," & _
              " MiRE='" & Format(RefIni(2).Value, "HH:MM:SS") & "',MiRS='" & Format(RefriFin(2).Value, "HH:MM:SS") & "'," & _
              " JuRE='" & Format(RefIni(3).Value, "HH:MM:SS") & "',JuRS='" & Format(RefriFin(3).Value, "HH:MM:SS") & "'," & _
              " ViRE='" & Format(RefIni(4).Value, "HH:MM:SS") & "',ViRS='" & Format(RefriFin(4).Value, "HH:MM:SS") & "'," & _
              " SaRE='" & Format(RefIni(5).Value, "HH:MM:SS") & "',SaRS='" & Format(RefriFin(5).Value, "HH:MM:SS") & "'," & _
              " DoRE='" & Format(RefIni(6).Value, "HH:MM:SS") & "',DoRS='" & Format(RefriFin(6).Value, "HH:MM:SS") & "'," & _
              " TolEntrada=" & IIf(chkTolEnt.Value = True, meTolEnt, 0) & ",TolSalida=" & IIf(chkTolSal.Value = True, meTolSal, 0) & "," & _
              " Tarde=" & IIf(chkTarde.Value = True, meTarde, 0) & ",Inasistencia=" & IIf(chkInasistencia.Value = True, meInasistencia, 0) & "," & _
              " AutoRefri='" & IIf(RegRefigerio.Value = 1, "S", "N") & "'," & _
              " diast=" & IIf(meDT.Text = "__", 0, val(meDT.Text)) & "," & _
              " diasd=" & IIf(meDD.Text = "__", 0, val(meDD.Text)) & "," & _
              " hrsxsem=" & IIf(meHrsxSem.Text = "__", 0, val(meHrsxSem.Text)) & "," & _
              " AutoAsi='" & IIf(RegAutomatico.Value = 1, "S", "N") & "'," & _
              " AsisAlea='" & IIf(RegAleatorio.Value = 1, "S", "N") & "'" & _
              " where nombre='" & Nombre & "'"
              
    Else
        SQL = "Insert into rh_Horarios (nombre,LuE,LuS,MaE,MaS,MiE,MiS,JuE,JuS,ViE," & _
              " ViS,SaE,SaS,DoE,DoS,LuRE,LuRS,MaRE,MaRS,MiRE,MiRS,JuRE,JuRS,ViRE,ViRS," & _
              " SaRE,SaRS,DoRE,DoRS,TolEntrada,TolSalida,Tarde,Inasistencia,AutoRefri,hrsxsem,diast,diasd,AutoAsi,AsisAlea)" & _
              " values ('" & Nombre & "',"
              For I = 0 To 6
                SQL = SQL & "'" & IIf(chkDias(I).Value = 1, Format(HorEntrada(I).Value, "HH:MM:SS"), "00:00:00") & "','" & IIf(chkDias(I).Value = 1, Format(HorSalida(I).Value, "HH:MM:SS"), "00:00:00") & "',"
              Next
              For I = 0 To 6
                SQL = SQL & "'" & IIf(Refri(I).Value = 1, Format(RefIni(I).Value, "HH:MM:SS"), "00:00:00") & "','" & IIf(Refri(I).Value = 1, Format(RefriFin(I).Value, "HH:MM:SS"), "00:00:00") & "',"
              Next
              SQL = SQL & IIf(chkTolEnt.Value = True, meTolEnt, 0) & ","
              SQL = SQL & IIf(chkTolSal.Value = True, meTolSal, 0) & ","
              SQL = SQL & IIf(chkTarde.Value = True, meTarde, 0) & ","
              SQL = SQL & IIf(chkInasistencia.Value = True, meInasistencia, 0) & ","
              SQL = SQL & "'" & IIf(RegRefigerio.Value = 1, "S", "N") & "',"
              SQL = SQL & IIf(meHrsxSem.Text = "__", 0, val(meHrsxSem.Text)) & ","
              SQL = SQL & IIf(meDT.Text = "__", 0, val(meDT.Text)) & ","
              SQL = SQL & IIf(meDD.Text = "__", 0, val(meDD.Text)) & ","
              SQL = SQL & "'" & IIf(RegRefigerio.Value = 1, "S", "N") & "',"
              SQL = SQL & "'" & IIf(RegRefigerio.Value = 1, "S", "N") & "')"
              
        
    End If
     oConexionMYSQL.Execute SQL
End Function

Private Sub cmbHorarios_Click()
    cmbHorarios_Change
End Sub

Private Sub cmdGrabar_Click()
    Dim NombreHorario As String
    If cmbHorarios.Visible = True Then
        NombreHorario = cmbHorarios.Text
    Else
        NombreHorario = txtHorario
    End If
    If ValidaDatos = True Then
        GrabaHorario NombreHorario
    Else
        MsgBox "Error al momento de grabar" & vbNewLine & "Revise los datos ingresados y vuelva intentarlo", vbOKOnly + vbExclamation, "NOVPeru"
    End If
End Sub

Private Sub cmdNuevo_Click()
    If CmdNuevo.Caption = "Nuevo" Then
        CmdNuevo.Caption = "Cancelar"
        cmbHorarios.Visible = False
        txtHorario.Visible = True
        txtHorario.Text = Empty
        txtHorario_Change
        txtHorario.SetFocus
    Else
        CmdNuevo.Caption = "Nuevo"
        cmbHorarios.Visible = True
        txtHorario.Visible = False
        LimpiaControles
        cmdGrabar.Enabled = False
        ActivaDesactivaDias False
        Form_Activate
    End If
    
End Sub

Private Sub Form_Activate()
    If CargaHorarios > 0 Then
        cmbHorarios.Enabled = True
        cmbHorarios.Enabled = True
        cmbHorarios.ListIndex = 0
        cmbHorarios_Change
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub HorEntrada_Change(Index As Integer)
    HorSalida(Index).MinDate = HorEntrada(Index).Value
    HorSalida(Index).Value = HorEntrada(Index).Value
End Sub

Private Sub meDD_GotFocus()
    meDD.SelStart = 0
    meDD.SelLength = Len(meDT)
End Sub

Private Sub meDT_GotFocus()
    meDT.SelStart = 0
    meDT.SelLength = Len(meDT)
End Sub

Private Sub meHrsxSem_Change()
    meHrsxSem.SelStart = 0
    meHrsxSem.SelLength = Len(meHrsxSem)
End Sub

Private Sub meInasistencia_GotFocus()
    meInasistencia.SelStart = 0
    meInasistencia.SelLength = Len(meInasistencia)
End Sub

Private Sub meTarde_GotFocus()
    meTarde.SelStart = 0
    meTarde.SelLength = Len(meTarde)
End Sub

Private Sub meTolEnt_GotFocus()
    meTolEnt.SelStart = 0
    meTolEnt.SelLength = Len(meTolEnt)
End Sub

Private Sub meTolSal_Change()
    meTolSal.SelStart = 0
    meTolSal.SelLength = Len(meTolSal)
End Sub

Private Sub RefIni_Change(Index As Integer)
    RefriFin(Index).MinDate = RefIni(Index).Value
    RefriFin(Index).Value = RefIni(Index).Value
End Sub

Private Sub Refri_Click(Index As Integer)
    Dim I As Integer
    If Refri(Index).Value = 1 Then
        RefIni(Index).Enabled = True
        RefriFin(Index).Enabled = True
    Else
        RefIni(Index).Enabled = False
        RefriFin(Index).Enabled = False
    End If
    cmdCopiarR.Enabled = False
    For I = 0 To 6
        If Refri(I).Value = 1 Then
            cmdCopiarR.Enabled = True
        End If
    Next
End Sub

Private Sub txtHorario_Change()
    If txtHorario.Text = Empty Then
        LimpiaControles
        ActivaDesactivaDias False
        chkInasistencia.Enabled = False
        chkTarde.Enabled = False
        chkTolEnt.Enabled = False
        chkTolSal.Enabled = False
        RegRefigerio.Enabled = False
    Else
        ActivaDesactivaDias True
        chkInasistencia.Enabled = True
        chkTarde.Enabled = True
        chkTolEnt.Enabled = True
        chkTolSal.Enabled = True
        RegRefigerio.Enabled = True
    End If
        
End Sub
