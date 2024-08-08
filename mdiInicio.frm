VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiInicio 
   BackColor       =   &H8000000C&
   Caption         =   "BRANDT - Sistema Integrado Administrativo"
   ClientHeight    =   10650
   ClientLeft      =   4470
   ClientTop       =   2235
   ClientWidth     =   13560
   Icon            =   "mdiInicio.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgsBarraHerramientas 
      Left            =   960
      Top             =   8820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   91
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":4730
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":4B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":4FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":5426
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":5878
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":5CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":611C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":B90E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":BD60
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":C1B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":C604
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":CA56
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":CEA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":D2FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":D74C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":DB9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":DFF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":E442
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":E894
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":ECE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":11498
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":117B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":11C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":12853
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":12D0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":13027
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":13909
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1549F
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":17FA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":183F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1870F
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":18A29
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":18D43
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":19065
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1937F
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":19699
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":199B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":19CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1A11F
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1A439
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1A753
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1AA6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1AD87
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1B0A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1B3BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1B6D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1B9EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1BD09
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1C023
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1C33D
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1C657
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1C971
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1CC8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1CFA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1D2BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":1D5D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2091E
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":26540
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2685A
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":26B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":26E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":271A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":274C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":277DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":27AF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":27FFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":28316
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":28630
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2894A
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":28C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":28F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":29298
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":295B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2B2BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2CFC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2D120
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2D27A
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2D3D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2D6EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2DA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2DB62
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2DCBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":2DE16
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":30198
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":30A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":3277C
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":328D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInicio.frx":32A30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbPrincipal 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgsBarraHerramientas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "mdiInicio.frx":32B8A
   End
   Begin MSComctlLib.StatusBar sbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   10335
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6174
            MinWidth        =   6174
            Picture         =   "mdiInicio.frx":33864
            Text            =   "NOV"
            TextSave        =   "NOV"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   2117
            MinWidth        =   2117
            Picture         =   "mdiInicio.frx":33DFE
            Text            =   "ADM"
            TextSave        =   "ADM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "mdiInicio.frx":34398
            Text            =   "Control Documentario"
            TextSave        =   "Control Documentario"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1411
            MinWidth        =   1411
            Picture         =   "mdiInicio.frx":34932
            Text            =   "2005"
            TextSave        =   "2005"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "DICIEMBRE"
            TextSave        =   "DICIEMBRE"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18/02/2020"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "mdiInicio.frx":34ECC
            Text            =   "EDUARDO"
            TextSave        =   "EDUARDO"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Servidor y BD"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10425
      Left            =   0
      Picture         =   "mdiInicio.frx":35466
      ScaleHeight     =   10365
      ScaleWidth      =   13500
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   13560
      Begin VB.PictureBox picprincipal 
         Height          =   600
         Left            =   540
         Picture         =   "mdiInicio.frx":B9DFA
         ScaleHeight     =   540
         ScaleWidth      =   1440
         TabIndex        =   21
         Top             =   1620
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picfac 
         Height          =   510
         Left            =   5760
         Picture         =   "mdiInicio.frx":13E78E
         ScaleHeight     =   450
         ScaleWidth      =   1665
         TabIndex        =   20
         Top             =   9135
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.PictureBox picpagar 
         Height          =   870
         Left            =   2700
         Picture         =   "mdiInicio.frx":1C0112
         ScaleHeight     =   810
         ScaleWidth      =   1440
         TabIndex        =   19
         Top             =   6345
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picpcobrar 
         Height          =   735
         Left            =   2310
         Picture         =   "mdiInicio.frx":2428E1
         ScaleHeight     =   675
         ScaleWidth      =   1395
         TabIndex        =   18
         Top             =   3960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox piccontrol 
         Height          =   465
         Left            =   11025
         Picture         =   "mdiInicio.frx":2C54A2
         ScaleHeight     =   405
         ScaleWidth      =   1260
         TabIndex        =   17
         Top             =   2655
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.PictureBox picplanilla 
         Height          =   645
         Left            =   8100
         Picture         =   "mdiInicio.frx":3489CC
         ScaleHeight     =   585
         ScaleWidth      =   1350
         TabIndex        =   16
         Top             =   3015
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.PictureBox picrechumanos 
         Height          =   690
         Left            =   4500
         Picture         =   "mdiInicio.frx":3C989B
         ScaleHeight     =   630
         ScaleWidth      =   1530
         TabIndex        =   15
         Top             =   2070
         Visible         =   0   'False
         Width           =   1590
      End
      Begin MSComDlg.CommonDialog CmD 
         Left            =   270
         Top             =   6720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsIpHost 
         Left            =   570
         Top             =   7740
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSForms.Label lblRecHum 
         Height          =   300
         Index           =   1
         Left            =   4950
         TabIndex        =   5
         Top             =   3000
         Width           =   1380
         ForeColor       =   8454143
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "R.R.H.H."
         Size            =   "2434;529"
         MousePointer    =   99
         MouseIcon       =   "mdiInicio.frx":44C3D0
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblPlanilla 
         Height          =   360
         Index           =   1
         Left            =   8550
         TabIndex        =   11
         Top             =   3840
         Width           =   1020
         ForeColor       =   8454143
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "Planilla"
         Size            =   "1799;635"
         MousePointer    =   99
         MouseIcon       =   "mdiInicio.frx":44D0AA
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblCtasxPagar 
         Height          =   360
         Index           =   1
         Left            =   3450
         TabIndex        =   9
         Top             =   7440
         Width           =   1860
         ForeColor       =   8454143
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "Ctas. x Pagar"
         Size            =   "3281;635"
         MousePointer    =   99
         MouseIcon       =   "mdiInicio.frx":44DD84
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblCtasxCobrar 
         Height          =   450
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Top             =   4920
         Width           =   2130
         ForeColor       =   8454143
         BackColor       =   -2147483636
         VariousPropertyBits=   8388627
         Caption         =   "Ctas. x Cobrar"
         Size            =   "3757;794"
         MousePointer    =   99
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblFacturacion 
         Height          =   360
         Index           =   1
         Left            =   5940
         TabIndex        =   3
         Top             =   8640
         Width           =   1665
         ForeColor       =   8454143
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "Facturación"
         Size            =   "2937;635"
         MousePointer    =   99
         MouseIcon       =   "mdiInicio.frx":44EA5E
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblFacturacion 
         Height          =   360
         Index           =   0
         Left            =   5970
         TabIndex        =   4
         Top             =   8670
         Width           =   1665
         ForeColor       =   8388608
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "Facturación"
         Size            =   "2937;635"
         MousePointer    =   99
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblControlDoc 
         Height          =   360
         Index           =   0
         Left            =   10710
         TabIndex        =   0
         Top             =   3300
         Width           =   3060
         ForeColor       =   8454143
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "Control Documentario"
         Size            =   "5397;635"
         MousePointer    =   99
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblControlDoc 
         Height          =   465
         Index           =   1
         Left            =   10740
         TabIndex        =   2
         Top             =   3330
         Width           =   3315
         ForeColor       =   8388608
         BackColor       =   -2147483636
         VariousPropertyBits=   8388627
         Caption         =   "Control Documentario"
         Size            =   "5847;820"
         MousePointer    =   99
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblCtasxCobrar 
         Height          =   450
         Index           =   0
         Left            =   2670
         TabIndex        =   8
         Top             =   4950
         Width           =   2130
         ForeColor       =   8388608
         BackColor       =   -2147483636
         VariousPropertyBits=   8388627
         Caption         =   "Ctas. x Cobrar"
         Size            =   "3757;794"
         MousePointer    =   99
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblCtasxPagar 
         Height          =   360
         Index           =   0
         Left            =   3450
         TabIndex        =   10
         Top             =   7470
         Width           =   1860
         ForeColor       =   8388608
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "Ctas. x Pagar"
         Size            =   "3281;635"
         MousePointer    =   99
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblPlanilla 
         Height          =   360
         Index           =   0
         Left            =   8580
         TabIndex        =   12
         Top             =   3870
         Width           =   1020
         ForeColor       =   8388608
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "Planilla"
         Size            =   "1799;635"
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblRecHum 
         Height          =   360
         Index           =   0
         Left            =   4980
         TabIndex        =   6
         Top             =   3030
         Width           =   1245
         ForeColor       =   8388608
         BackColor       =   -2147483636
         VariousPropertyBits=   276824083
         Caption         =   "R.R.H.H."
         Size            =   "2196;635"
         MousePointer    =   99
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Image imgRegresar 
         Height          =   600
         Left            =   12960
         MouseIcon       =   "mdiInicio.frx":44F738
         MousePointer    =   99  'Custom
         Picture         =   "mdiInicio.frx":45013A
         ToolTipText     =   " R E G R E S A R "
         Top             =   8970
         Width           =   645
      End
   End
   Begin VB.Menu m_01nivel01 
      Caption         =   "01Nivel01"
      Index           =   0
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   0
         Begin VB.Menu m_01nivel03_01 
            Caption         =   "01Nivel03_01"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_01 
            Caption         =   "01Nivel03_01"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_01 
            Caption         =   "01Nivel03_01"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_01 
            Caption         =   "01Nivel03_01"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_01 
            Caption         =   "01Nivel03_01"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_01 
            Caption         =   "01Nivel03_01"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_01 
            Caption         =   "01Nivel03_01"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_01 
            Caption         =   "01Nivel03_01"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   1
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   2
         Begin VB.Menu m_01nivel03_02 
            Caption         =   "01nivel03_02"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_02 
            Caption         =   "01nivel03_02"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_02 
            Caption         =   "01nivel03_02"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_02 
            Caption         =   "01nivel03_02"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_02 
            Caption         =   "01nivel03_02"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_02 
            Caption         =   "01nivel03_02"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_02 
            Caption         =   "01nivel03_02"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_02 
            Caption         =   "01nivel03_02"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   3
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   4
         Begin VB.Menu m_01nivel03_03 
            Caption         =   "01nivel03_03"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_03 
            Caption         =   "01nivel03_03"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_03 
            Caption         =   "01nivel03_03"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_03 
            Caption         =   "01nivel03_03"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_03 
            Caption         =   "01nivel03_03"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_03 
            Caption         =   "01nivel03_03"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_03 
            Caption         =   "01nivel03_03"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_03 
            Caption         =   "01nivel03_03"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   5
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   6
         Begin VB.Menu m_01nivel03_04 
            Caption         =   "01nivel03_04"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_04 
            Caption         =   "01nivel03_04"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_04 
            Caption         =   "01nivel03_04"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_04 
            Caption         =   "01nivel03_04"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_04 
            Caption         =   "01nivel03_04"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_04 
            Caption         =   "01nivel03_04"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_04 
            Caption         =   "01nivel03_04"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_04 
            Caption         =   "01nivel03_04"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   7
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   8
         Begin VB.Menu m_01nivel03_05 
            Caption         =   "01nivel03_05"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_05 
            Caption         =   "01nivel03_05"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_05 
            Caption         =   "01nivel03_05"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_05 
            Caption         =   "01nivel03_05"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_05 
            Caption         =   "01nivel03_05"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_05 
            Caption         =   "01nivel03_05"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_05 
            Caption         =   "01nivel03_05"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_05 
            Caption         =   "01nivel03_05"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   9
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   10
         Begin VB.Menu m_01nivel03_06 
            Caption         =   "01nivel03_06"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_06 
            Caption         =   "01nivel03_06"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_06 
            Caption         =   "01nivel03_06"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_06 
            Caption         =   "01nivel03_06"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_06 
            Caption         =   "01nivel03_06"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_06 
            Caption         =   "01nivel03_06"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_06 
            Caption         =   "01nivel03_06"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_06 
            Caption         =   "01nivel03_06"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   11
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   12
         Begin VB.Menu m_01nivel03_07 
            Caption         =   "01nivel03_07"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_07 
            Caption         =   "01nivel03_07"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_07 
            Caption         =   "01nivel03_07"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_07 
            Caption         =   "01nivel03_07"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_07 
            Caption         =   "01nivel03_07"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_07 
            Caption         =   "01nivel03_07"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_07 
            Caption         =   "01nivel03_07"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_07 
            Caption         =   "01nivel03_07"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   13
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   14
         Begin VB.Menu m_01nivel03_08 
            Caption         =   "01nivel03_08"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_08 
            Caption         =   "01nivel03_08"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_08 
            Caption         =   "01nivel03_08"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_08 
            Caption         =   "01nivel03_08"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_08 
            Caption         =   "01nivel03_08"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_08 
            Caption         =   "01nivel03_08"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_08 
            Caption         =   "01nivel03_08"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_08 
            Caption         =   "01nivel03_08"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   15
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   16
         Begin VB.Menu m_01nivel03_09 
            Caption         =   "01nivel03_09"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_09 
            Caption         =   "01nivel03_09"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_09 
            Caption         =   "01nivel03_09"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_09 
            Caption         =   "01nivel03_09"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_09 
            Caption         =   "01nivel03_09"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_09 
            Caption         =   "01nivel03_09"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_09 
            Caption         =   "01nivel03_09"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_09 
            Caption         =   "01nivel03_09"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   17
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   18
         Begin VB.Menu m_01nivel03_10 
            Caption         =   "01nivel03_10"
            Index           =   0
         End
         Begin VB.Menu m_01nivel03_10 
            Caption         =   "01nivel03_10"
            Index           =   1
         End
         Begin VB.Menu m_01nivel03_10 
            Caption         =   "01nivel03_10"
            Index           =   2
         End
         Begin VB.Menu m_01nivel03_10 
            Caption         =   "01nivel03_10"
            Index           =   3
         End
         Begin VB.Menu m_01nivel03_10 
            Caption         =   "01nivel03_10"
            Index           =   4
         End
         Begin VB.Menu m_01nivel03_10 
            Caption         =   "01nivel03_10"
            Index           =   5
         End
         Begin VB.Menu m_01nivel03_10 
            Caption         =   "01nivel03_10"
            Index           =   6
         End
         Begin VB.Menu m_01nivel03_10 
            Caption         =   "01nivel03_10"
            Index           =   7
         End
      End
      Begin VB.Menu m_01nivel02 
         Caption         =   "01Nivel02"
         Index           =   19
      End
   End
   Begin VB.Menu m_02nivel01 
      Caption         =   "02Nivel01"
      Index           =   0
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   0
         Begin VB.Menu m_02nivel03_01 
            Caption         =   "02Nivel03_01"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_01 
            Caption         =   "02Nivel03_01"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_01 
            Caption         =   "02Nivel03_01"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_01 
            Caption         =   "02Nivel03_01"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_01 
            Caption         =   "02Nivel03_01"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_01 
            Caption         =   "02Nivel03_01"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_01 
            Caption         =   "02Nivel03_01"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_01 
            Caption         =   "02Nivel03_01"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   1
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   2
         Begin VB.Menu m_02nivel03_02 
            Caption         =   "02nivel03_02"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_02 
            Caption         =   "02nivel03_02"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_02 
            Caption         =   "02nivel03_02"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_02 
            Caption         =   "02nivel03_02"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_02 
            Caption         =   "02nivel03_02"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_02 
            Caption         =   "02nivel03_02"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_02 
            Caption         =   "02nivel03_02"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_02 
            Caption         =   "02nivel03_02"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   3
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   4
         Begin VB.Menu m_02nivel03_03 
            Caption         =   "02nivel03_03"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_03 
            Caption         =   "02nivel03_03"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_03 
            Caption         =   "02nivel03_03"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_03 
            Caption         =   "02nivel03_03"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_03 
            Caption         =   "02nivel03_03"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_03 
            Caption         =   "02nivel03_03"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_03 
            Caption         =   "02nivel03_03"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_03 
            Caption         =   "02nivel03_03"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   5
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   6
         Begin VB.Menu m_02nivel03_04 
            Caption         =   "02nivel03_04"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_04 
            Caption         =   "02nivel03_04"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_04 
            Caption         =   "02nivel03_04"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_04 
            Caption         =   "02nivel03_04"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_04 
            Caption         =   "02nivel03_04"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_04 
            Caption         =   "02nivel03_04"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_04 
            Caption         =   "02nivel03_04"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_04 
            Caption         =   "02nivel03_04"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   7
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   8
         Begin VB.Menu m_02nivel03_05 
            Caption         =   "02nivel03_05"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_05 
            Caption         =   "02nivel03_05"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_05 
            Caption         =   "02nivel03_05"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_05 
            Caption         =   "02nivel03_05"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_05 
            Caption         =   "02nivel03_05"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_05 
            Caption         =   "02nivel03_05"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_05 
            Caption         =   "02nivel03_05"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_05 
            Caption         =   "02nivel03_05"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   9
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   10
         Begin VB.Menu m_02nivel03_06 
            Caption         =   "02nivel03_06"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_06 
            Caption         =   "02nivel03_06"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_06 
            Caption         =   "02nivel03_06"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_06 
            Caption         =   "02nivel03_06"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_06 
            Caption         =   "02nivel03_06"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_06 
            Caption         =   "02nivel03_06"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_06 
            Caption         =   "02nivel03_06"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_06 
            Caption         =   "02nivel03_06"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   11
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   12
         Begin VB.Menu m_02nivel03_07 
            Caption         =   "02nivel03_07"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_07 
            Caption         =   "02nivel03_07"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_07 
            Caption         =   "02nivel03_07"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_07 
            Caption         =   "02nivel03_07"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_07 
            Caption         =   "02nivel03_07"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_07 
            Caption         =   "02nivel03_07"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_07 
            Caption         =   "02nivel03_07"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_07 
            Caption         =   "02nivel03_07"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   13
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   14
         Begin VB.Menu m_02nivel03_08 
            Caption         =   "02nivel03_08"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_08 
            Caption         =   "02nivel03_08"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_08 
            Caption         =   "02nivel03_08"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_08 
            Caption         =   "02nivel03_08"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_08 
            Caption         =   "02nivel03_08"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_08 
            Caption         =   "02nivel03_08"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_08 
            Caption         =   "02nivel03_08"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_08 
            Caption         =   "02nivel03_08"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   15
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   16
         Begin VB.Menu m_02nivel03_09 
            Caption         =   "02nivel03_09"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_09 
            Caption         =   "02nivel03_09"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_09 
            Caption         =   "02nivel03_09"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_09 
            Caption         =   "02nivel03_09"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_09 
            Caption         =   "02nivel03_09"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_09 
            Caption         =   "02nivel03_09"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_09 
            Caption         =   "02nivel03_09"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_09 
            Caption         =   "02nivel03_09"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   17
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   18
         Begin VB.Menu m_02nivel03_10 
            Caption         =   "02nivel03_10"
            Index           =   0
         End
         Begin VB.Menu m_02nivel03_10 
            Caption         =   "02nivel03_10"
            Index           =   1
         End
         Begin VB.Menu m_02nivel03_10 
            Caption         =   "02nivel03_10"
            Index           =   2
         End
         Begin VB.Menu m_02nivel03_10 
            Caption         =   "02nivel03_10"
            Index           =   3
         End
         Begin VB.Menu m_02nivel03_10 
            Caption         =   "02nivel03_10"
            Index           =   4
         End
         Begin VB.Menu m_02nivel03_10 
            Caption         =   "02nivel03_10"
            Index           =   5
         End
         Begin VB.Menu m_02nivel03_10 
            Caption         =   "02nivel03_10"
            Index           =   6
         End
         Begin VB.Menu m_02nivel03_10 
            Caption         =   "02nivel03_10"
            Index           =   7
         End
      End
      Begin VB.Menu m_02nivel02 
         Caption         =   "02Nivel02"
         Index           =   19
      End
   End
   Begin VB.Menu m_03nivel01 
      Caption         =   "03Nivel01"
      Index           =   0
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   0
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_01 
               Caption         =   "03Nivel04_01_01"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_01 
               Caption         =   "03Nivel04_01_01"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_01 
               Caption         =   "03Nivel04_01_01"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_01 
               Caption         =   "03Nivel04_01_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_01 
               Caption         =   "03Nivel04_02_01"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_01 
               Caption         =   "03Nivel04_02_01"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_01 
               Caption         =   "03Nivel04_02_01"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_01 
               Caption         =   "03Nivel04_02_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_01 
               Caption         =   "03Nivel04_03_01"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_01 
               Caption         =   "03Nivel04_03_01"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_01 
               Caption         =   "03Nivel04_03_01"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_01 
               Caption         =   "03Nivel04_03_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_01 
               Caption         =   "03Nivel04_04_01"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_01 
               Caption         =   "03Nivel04_04_01"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_01 
               Caption         =   "03Nivel04_04_01"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_01 
               Caption         =   "03Nivel04_04_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_01 
               Caption         =   "03Nivel04_05_01"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_01 
               Caption         =   "03Nivel04_05_01"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_01 
               Caption         =   "03Nivel04_05_01"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_01 
               Caption         =   "03Nivel04_05_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_01 
            Caption         =   "03Nivel03_01"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   1
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   2
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_02 
               Caption         =   "03nivel04_01_02"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_02 
               Caption         =   "03nivel04_01_02"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_02 
               Caption         =   "03nivel04_01_02"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_02 
               Caption         =   "03nivel04_01_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_02 
               Caption         =   "03nivel04_02_02"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_02 
               Caption         =   "03nivel04_02_02"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_02 
               Caption         =   "03nivel04_02_02"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_02 
               Caption         =   "03nivel04_02_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_02 
               Caption         =   "03nivel04_03_02"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_02 
               Caption         =   "03nivel04_03_02"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_02 
               Caption         =   "03nivel04_03_02"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_02 
               Caption         =   "03nivel04_03_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_02 
               Caption         =   "03nivel04_04_02"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_02 
               Caption         =   "03nivel04_04_02"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_02 
               Caption         =   "03nivel04_04_02"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_02 
               Caption         =   "03nivel04_04_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_02 
               Caption         =   "03nivel04_05_02"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_02 
               Caption         =   "03nivel04_05_02"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_02 
               Caption         =   "03nivel04_05_02"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_02 
               Caption         =   "03nivel04_05_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_02 
            Caption         =   "03nivel03_02"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   3
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   4
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_03 
               Caption         =   "03nivel04_01_03"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_03 
               Caption         =   "03nivel04_01_03"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_03 
               Caption         =   "03nivel04_01_03"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_03 
               Caption         =   "03nivel04_01_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_03 
               Caption         =   "03nivel04_02_03"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_03 
               Caption         =   "03nivel04_02_03"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_03 
               Caption         =   "03nivel04_02_03"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_03 
               Caption         =   "03nivel04_02_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_03 
               Caption         =   "03nivel04_03_03"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_03 
               Caption         =   "03nivel04_03_03"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_03 
               Caption         =   "03nivel04_03_03"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_03 
               Caption         =   "03nivel04_03_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_03 
               Caption         =   "03nivel04_04_03"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_03 
               Caption         =   "03nivel04_04_03"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_03 
               Caption         =   "03nivel04_04_03"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_03 
               Caption         =   "03nivel04_04_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_03 
               Caption         =   "03nivel04_05_03"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_03 
               Caption         =   "03nivel04_05_03"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_03 
               Caption         =   "03nivel04_05_03"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_03 
               Caption         =   "03nivel04_05_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_03 
            Caption         =   "03nivel03_03"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   5
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   6
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_04 
               Caption         =   "03nivel04_01_04"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_04 
               Caption         =   "03nivel04_01_04"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_04 
               Caption         =   "03nivel04_01_04"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_04 
               Caption         =   "03nivel04_01_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_04 
               Caption         =   "03nivel04_02_04"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_04 
               Caption         =   "03nivel04_02_04"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_04 
               Caption         =   "03nivel04_02_04"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_04 
               Caption         =   "03nivel04_02_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_04 
               Caption         =   "03nivel04_03_04"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_04 
               Caption         =   "03nivel04_03_04"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_04 
               Caption         =   "03nivel04_03_04"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_04 
               Caption         =   "03nivel04_03_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_04 
               Caption         =   "03nivel04_04_04"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_04 
               Caption         =   "03nivel04_04_04"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_04 
               Caption         =   "03nivel04_04_04"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_04 
               Caption         =   "03nivel04_04_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_04 
               Caption         =   "03nivel04_05_04"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_04 
               Caption         =   "03nivel04_05_04"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_04 
               Caption         =   "03nivel04_05_04"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_04 
               Caption         =   "03nivel04_05_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_04 
            Caption         =   "03nivel03_04"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   7
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   8
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_05 
               Caption         =   "03nivel04_01_05"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_05 
               Caption         =   "03nivel04_01_05"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_05 
               Caption         =   "03nivel04_01_05"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_05 
               Caption         =   "03nivel04_01_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_05 
               Caption         =   "03nivel04_02_05"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_05 
               Caption         =   "03nivel04_02_05"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_05 
               Caption         =   "03nivel04_02_05"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_05 
               Caption         =   "03nivel04_02_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_05 
               Caption         =   "03nivel04_03_05"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_05 
               Caption         =   "03nivel04_03_05"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_05 
               Caption         =   "03nivel04_03_05"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_05 
               Caption         =   "03nivel04_03_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_05 
               Caption         =   "03nivel04_04_05"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_05 
               Caption         =   "03nivel04_04_05"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_05 
               Caption         =   "03nivel04_04_05"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_05 
               Caption         =   "03nivel04_04_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_05 
               Caption         =   "03nivel04_05_05"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_05 
               Caption         =   "03nivel04_05_05"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_05 
               Caption         =   "03nivel04_05_05"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_05 
               Caption         =   "03nivel04_05_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_05 
            Caption         =   "03nivel03_05"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   9
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   10
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_06 
               Caption         =   "03nivel04_01_06"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_06 
               Caption         =   "03nivel04_01_06"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_06 
               Caption         =   "03nivel04_01_06"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_06 
               Caption         =   "03nivel04_01_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_06 
               Caption         =   "03nivel04_02_06"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_06 
               Caption         =   "03nivel04_02_06"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_06 
               Caption         =   "03nivel04_02_06"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_06 
               Caption         =   "03nivel04_02_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_06 
               Caption         =   "03nivel04_03_06"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_06 
               Caption         =   "03nivel04_03_06"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_06 
               Caption         =   "03nivel04_03_06"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_06 
               Caption         =   "03nivel04_03_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_06 
               Caption         =   "03nivel04_04_06"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_06 
               Caption         =   "03nivel04_04_06"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_06 
               Caption         =   "03nivel04_04_06"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_06 
               Caption         =   "03nivel04_04_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_06 
               Caption         =   "03nivel04_05_06"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_06 
               Caption         =   "03nivel04_05_06"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_06 
               Caption         =   "03nivel04_05_06"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_06 
               Caption         =   "03nivel04_05_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_06 
            Caption         =   "03nivel03_06"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   11
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   12
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_07 
               Caption         =   "03nivel04_01_07"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_07 
               Caption         =   "03nivel04_01_07"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_07 
               Caption         =   "03nivel04_01_07"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_07 
               Caption         =   "03nivel04_01_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_07 
               Caption         =   "03nivel04_02_07"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_07 
               Caption         =   "03nivel04_02_07"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_07 
               Caption         =   "03nivel04_02_07"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_07 
               Caption         =   "03nivel04_02_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_07 
               Caption         =   "03nivel04_03_07"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_07 
               Caption         =   "03nivel04_03_07"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_07 
               Caption         =   "03nivel04_03_07"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_07 
               Caption         =   "03nivel04_03_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_07 
               Caption         =   "03nivel04_04_07"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_07 
               Caption         =   "03nivel04_04_07"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_07 
               Caption         =   "03nivel04_04_07"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_07 
               Caption         =   "03nivel04_04_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_07 
               Caption         =   "03nivel04_05_07"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_07 
               Caption         =   "03nivel04_05_07"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_07 
               Caption         =   "03nivel04_05_07"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_07 
               Caption         =   "03nivel04_05_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_07 
            Caption         =   "03nivel03_07"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   13
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   14
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_08 
               Caption         =   "03nivel04_01_08"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_08 
               Caption         =   "03nivel04_01_08"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_08 
               Caption         =   "03nivel04_01_08"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_08 
               Caption         =   "03nivel04_01_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_08 
               Caption         =   "03nivel04_02_08"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_08 
               Caption         =   "03nivel04_02_08"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_08 
               Caption         =   "03nivel04_02_08"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_08 
               Caption         =   "03nivel04_02_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_08 
               Caption         =   "03nivel04_03_08"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_08 
               Caption         =   "03nivel04_03_08"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_08 
               Caption         =   "03nivel04_03_08"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_08 
               Caption         =   "03nivel04_03_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_08 
               Caption         =   "03nivel04_04_08"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_08 
               Caption         =   "03nivel04_04_08"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_08 
               Caption         =   "03nivel04_04_08"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_08 
               Caption         =   "03nivel04_04_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_08 
               Caption         =   "03nivel04_05_08"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_08 
               Caption         =   "03nivel04_05_08"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_08 
               Caption         =   "03nivel04_05_08"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_08 
               Caption         =   "03nivel04_05_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_08 
            Caption         =   "03nivel03_08"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   15
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   16
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_09 
               Caption         =   "03nivel04_01_09"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_09 
               Caption         =   "03nivel04_01_09"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_09 
               Caption         =   "03nivel04_01_09"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_09 
               Caption         =   "03nivel04_01_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_09 
               Caption         =   "03nivel04_02_09"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_09 
               Caption         =   "03nivel04_02_09"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_09 
               Caption         =   "03nivel04_02_09"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_09 
               Caption         =   "03nivel04_02_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_09 
               Caption         =   "03nivel04_03_09"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_09 
               Caption         =   "03nivel04_03_09"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_09 
               Caption         =   "03nivel04_03_09"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_09 
               Caption         =   "03nivel04_03_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_09 
               Caption         =   "03nivel04_04_09"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_09 
               Caption         =   "03nivel04_04_09"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_09 
               Caption         =   "03nivel04_04_09"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_09 
               Caption         =   "03nivel04_04_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_09 
               Caption         =   "03nivel04_05_09"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_09 
               Caption         =   "03nivel04_05_09"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_09 
               Caption         =   "03nivel04_05_09"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_09 
               Caption         =   "03nivel04_05_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_09 
            Caption         =   "03nivel03_09"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   17
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   18
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   0
            Begin VB.Menu m_03nivel04_01_10 
               Caption         =   "03nivel04_01_10"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_01_10 
               Caption         =   "03nivel04_01_10"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_01_10 
               Caption         =   "03nivel04_01_10"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_01_10 
               Caption         =   "03nivel04_01_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   1
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   2
            Begin VB.Menu m_03nivel04_02_10 
               Caption         =   "03nivel04_02_10"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_02_10 
               Caption         =   "03nivel04_02_10"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_02_10 
               Caption         =   "03nivel04_02_10"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_02_10 
               Caption         =   "03nivel04_02_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   3
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   4
            Begin VB.Menu m_03nivel04_03_10 
               Caption         =   "03nivel04_03_10"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_03_10 
               Caption         =   "03nivel04_03_10"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_03_10 
               Caption         =   "03nivel04_03_10"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_03_10 
               Caption         =   "03nivel04_03_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   5
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   6
            Begin VB.Menu m_03nivel04_04_10 
               Caption         =   "03nivel04_04_10"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_04_10 
               Caption         =   "03nivel04_04_10"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_04_10 
               Caption         =   "03nivel04_04_10"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_04_10 
               Caption         =   "03nivel04_04_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   7
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   8
            Begin VB.Menu m_03nivel04_05_10 
               Caption         =   "03nivel04_05_10"
               Index           =   0
            End
            Begin VB.Menu m_03nivel04_05_10 
               Caption         =   "03nivel04_05_10"
               Index           =   1
            End
            Begin VB.Menu m_03nivel04_05_10 
               Caption         =   "03nivel04_05_10"
               Index           =   2
            End
            Begin VB.Menu m_03nivel04_05_10 
               Caption         =   "03nivel04_05_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   9
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   10
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   11
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   12
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   13
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   14
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   15
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   16
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   17
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   18
         End
         Begin VB.Menu m_03nivel03_10 
            Caption         =   "03nivel03_10"
            Index           =   19
         End
      End
      Begin VB.Menu m_03nivel02 
         Caption         =   "03Nivel02"
         Index           =   19
      End
   End
   Begin VB.Menu m_04nivel01 
      Caption         =   "04Nivel01"
      Index           =   0
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   0
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_01 
               Caption         =   "04Nivel04_01_01"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_01 
               Caption         =   "04Nivel04_01_01"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_01 
               Caption         =   "04Nivel04_01_01"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_01 
               Caption         =   "04Nivel04_01_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_01 
               Caption         =   "04Nivel04_02_01"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_01 
               Caption         =   "04Nivel04_02_01"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_01 
               Caption         =   "04Nivel04_02_01"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_01 
               Caption         =   "04Nivel04_02_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_01 
               Caption         =   "04Nivel04_03_01"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_01 
               Caption         =   "04Nivel04_03_01"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_01 
               Caption         =   "04Nivel04_03_01"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_01 
               Caption         =   "04Nivel04_03_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_01 
               Caption         =   "04Nivel04_04_01"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_01 
               Caption         =   "04Nivel04_04_01"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_01 
               Caption         =   "04Nivel04_04_01"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_01 
               Caption         =   "04Nivel04_04_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_01 
               Caption         =   "04Nivel04_05_01"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_01 
               Caption         =   "04Nivel04_05_01"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_01 
               Caption         =   "04Nivel04_05_01"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_01 
               Caption         =   "04Nivel04_05_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_01 
            Caption         =   "04Nivel03_01"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   1
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   2
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_02 
               Caption         =   "04nivel04_01_02"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_02 
               Caption         =   "04nivel04_01_02"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_02 
               Caption         =   "04nivel04_01_02"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_02 
               Caption         =   "04nivel04_01_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_02 
               Caption         =   "04nivel04_02_02"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_02 
               Caption         =   "04nivel04_02_02"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_02 
               Caption         =   "04nivel04_02_02"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_02 
               Caption         =   "04nivel04_02_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_02 
               Caption         =   "04nivel04_03_02"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_02 
               Caption         =   "04nivel04_03_02"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_02 
               Caption         =   "04nivel04_03_02"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_02 
               Caption         =   "04nivel04_03_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_02 
               Caption         =   "04nivel04_04_02"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_02 
               Caption         =   "04nivel04_04_02"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_02 
               Caption         =   "04nivel04_04_02"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_02 
               Caption         =   "04nivel04_04_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_02 
               Caption         =   "04nivel04_05_02"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_02 
               Caption         =   "04nivel04_05_02"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_02 
               Caption         =   "04nivel04_05_02"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_02 
               Caption         =   "04nivel04_05_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_02 
            Caption         =   "04nivel03_02"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   3
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   4
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_03 
               Caption         =   "04nivel04_01_03"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_03 
               Caption         =   "04nivel04_01_03"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_03 
               Caption         =   "04nivel04_01_03"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_03 
               Caption         =   "04nivel04_01_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_03 
               Caption         =   "04nivel04_02_03"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_03 
               Caption         =   "04nivel04_02_03"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_03 
               Caption         =   "04nivel04_02_03"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_03 
               Caption         =   "04nivel04_02_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_03 
               Caption         =   "04nivel04_03_03"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_03 
               Caption         =   "04nivel04_03_03"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_03 
               Caption         =   "04nivel04_03_03"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_03 
               Caption         =   "04nivel04_03_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_03 
               Caption         =   "04nivel04_04_03"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_03 
               Caption         =   "04nivel04_04_03"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_03 
               Caption         =   "04nivel04_04_03"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_03 
               Caption         =   "04nivel04_04_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_03 
               Caption         =   "04nivel04_05_03"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_03 
               Caption         =   "04nivel04_05_03"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_03 
               Caption         =   "04nivel04_05_03"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_03 
               Caption         =   "04nivel04_05_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_03 
            Caption         =   "04nivel03_03"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   5
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   6
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_04 
               Caption         =   "04nivel04_01_04"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_04 
               Caption         =   "04nivel04_01_04"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_04 
               Caption         =   "04nivel04_01_04"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_04 
               Caption         =   "04nivel04_01_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_04 
               Caption         =   "04nivel04_02_04"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_04 
               Caption         =   "04nivel04_02_04"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_04 
               Caption         =   "04nivel04_02_04"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_04 
               Caption         =   "04nivel04_02_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_04 
               Caption         =   "04nivel04_03_04"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_04 
               Caption         =   "04nivel04_03_04"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_04 
               Caption         =   "04nivel04_03_04"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_04 
               Caption         =   "04nivel04_03_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_04 
               Caption         =   "04nivel04_04_04"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_04 
               Caption         =   "04nivel04_04_04"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_04 
               Caption         =   "04nivel04_04_04"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_04 
               Caption         =   "04nivel04_04_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_04 
               Caption         =   "04nivel04_05_04"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_04 
               Caption         =   "04nivel04_05_04"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_04 
               Caption         =   "04nivel04_05_04"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_04 
               Caption         =   "04nivel04_05_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_04 
            Caption         =   "04nivel03_04"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   7
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   8
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_05 
               Caption         =   "04nivel04_01_05"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_05 
               Caption         =   "04nivel04_01_05"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_05 
               Caption         =   "04nivel04_01_05"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_05 
               Caption         =   "04nivel04_01_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_05 
               Caption         =   "04nivel04_02_05"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_05 
               Caption         =   "04nivel04_02_05"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_05 
               Caption         =   "04nivel04_02_05"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_05 
               Caption         =   "04nivel04_02_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_05 
               Caption         =   "04nivel04_03_05"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_05 
               Caption         =   "04nivel04_03_05"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_05 
               Caption         =   "04nivel04_03_05"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_05 
               Caption         =   "04nivel04_03_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_05 
               Caption         =   "04nivel04_04_05"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_05 
               Caption         =   "04nivel04_04_05"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_05 
               Caption         =   "04nivel04_04_05"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_05 
               Caption         =   "04nivel04_04_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_05 
               Caption         =   "04nivel04_05_05"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_05 
               Caption         =   "04nivel04_05_05"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_05 
               Caption         =   "04nivel04_05_05"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_05 
               Caption         =   "04nivel04_05_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_05 
            Caption         =   "04nivel03_05"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   9
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   10
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_06 
               Caption         =   "04nivel04_01_06"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_06 
               Caption         =   "04nivel04_01_06"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_06 
               Caption         =   "04nivel04_01_06"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_06 
               Caption         =   "04nivel04_01_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_06 
               Caption         =   "04nivel04_02_06"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_06 
               Caption         =   "04nivel04_02_06"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_06 
               Caption         =   "04nivel04_02_06"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_06 
               Caption         =   "04nivel04_02_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_06 
               Caption         =   "04nivel04_03_06"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_06 
               Caption         =   "04nivel04_03_06"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_06 
               Caption         =   "04nivel04_03_06"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_06 
               Caption         =   "04nivel04_03_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_06 
               Caption         =   "04nivel04_04_06"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_06 
               Caption         =   "04nivel04_04_06"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_06 
               Caption         =   "04nivel04_04_06"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_06 
               Caption         =   "04nivel04_04_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_06 
               Caption         =   "04nivel04_05_06"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_06 
               Caption         =   "04nivel04_05_06"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_06 
               Caption         =   "04nivel04_05_06"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_06 
               Caption         =   "04nivel04_05_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_06 
            Caption         =   "04nivel03_06"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   11
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   12
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_07 
               Caption         =   "04nivel04_01_07"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_07 
               Caption         =   "04nivel04_01_07"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_07 
               Caption         =   "04nivel04_01_07"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_07 
               Caption         =   "04nivel04_01_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_07 
               Caption         =   "04nivel04_02_07"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_07 
               Caption         =   "04nivel04_02_07"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_07 
               Caption         =   "04nivel04_02_07"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_07 
               Caption         =   "04nivel04_02_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_07 
               Caption         =   "04nivel04_03_07"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_07 
               Caption         =   "04nivel04_03_07"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_07 
               Caption         =   "04nivel04_03_07"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_07 
               Caption         =   "04nivel04_03_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_07 
               Caption         =   "04nivel04_04_07"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_07 
               Caption         =   "04nivel04_04_07"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_07 
               Caption         =   "04nivel04_04_07"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_07 
               Caption         =   "04nivel04_04_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_07 
               Caption         =   "04nivel04_05_07"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_07 
               Caption         =   "04nivel04_05_07"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_07 
               Caption         =   "04nivel04_05_07"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_07 
               Caption         =   "04nivel04_05_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_07 
            Caption         =   "04nivel03_07"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   13
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   14
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_08 
               Caption         =   "04nivel04_01_08"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_08 
               Caption         =   "04nivel04_01_08"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_08 
               Caption         =   "04nivel04_01_08"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_08 
               Caption         =   "04nivel04_01_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_08 
               Caption         =   "04nivel04_02_08"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_08 
               Caption         =   "04nivel04_02_08"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_08 
               Caption         =   "04nivel04_02_08"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_08 
               Caption         =   "04nivel04_02_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_08 
               Caption         =   "04nivel04_03_08"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_08 
               Caption         =   "04nivel04_03_08"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_08 
               Caption         =   "04nivel04_03_08"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_08 
               Caption         =   "04nivel04_03_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_08 
               Caption         =   "04nivel04_04_08"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_08 
               Caption         =   "04nivel04_04_08"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_08 
               Caption         =   "04nivel04_04_08"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_08 
               Caption         =   "04nivel04_04_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_08 
               Caption         =   "04nivel04_05_08"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_08 
               Caption         =   "04nivel04_05_08"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_08 
               Caption         =   "04nivel04_05_08"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_08 
               Caption         =   "04nivel04_05_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_08 
            Caption         =   "04nivel03_08"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   15
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   16
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_09 
               Caption         =   "04nivel04_01_09"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_09 
               Caption         =   "04nivel04_01_09"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_09 
               Caption         =   "04nivel04_01_09"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_09 
               Caption         =   "04nivel04_01_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_09 
               Caption         =   "04nivel04_02_09"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_09 
               Caption         =   "04nivel04_02_09"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_09 
               Caption         =   "04nivel04_02_09"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_09 
               Caption         =   "04nivel04_02_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_09 
               Caption         =   "04nivel04_03_09"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_09 
               Caption         =   "04nivel04_03_09"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_09 
               Caption         =   "04nivel04_03_09"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_09 
               Caption         =   "04nivel04_03_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_09 
               Caption         =   "04nivel04_04_09"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_09 
               Caption         =   "04nivel04_04_09"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_09 
               Caption         =   "04nivel04_04_09"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_09 
               Caption         =   "04nivel04_04_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_09 
               Caption         =   "04nivel04_05_09"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_09 
               Caption         =   "04nivel04_05_09"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_09 
               Caption         =   "04nivel04_05_09"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_09 
               Caption         =   "04nivel04_05_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_09 
            Caption         =   "04nivel03_09"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   17
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   18
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   0
            Begin VB.Menu m_04nivel04_01_10 
               Caption         =   "04nivel04_01_10"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_01_10 
               Caption         =   "04nivel04_01_10"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_01_10 
               Caption         =   "04nivel04_01_10"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_01_10 
               Caption         =   "04nivel04_01_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   1
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   2
            Begin VB.Menu m_04nivel04_02_10 
               Caption         =   "04nivel04_02_10"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_02_10 
               Caption         =   "04nivel04_02_10"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_02_10 
               Caption         =   "04nivel04_02_10"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_02_10 
               Caption         =   "04nivel04_02_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   3
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   4
            Begin VB.Menu m_04nivel04_03_10 
               Caption         =   "04nivel04_03_10"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_03_10 
               Caption         =   "04nivel04_03_10"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_03_10 
               Caption         =   "04nivel04_03_10"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_03_10 
               Caption         =   "04nivel04_03_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   5
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   6
            Begin VB.Menu m_04nivel04_04_10 
               Caption         =   "04nivel04_04_10"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_04_10 
               Caption         =   "04nivel04_04_10"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_04_10 
               Caption         =   "04nivel04_04_10"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_04_10 
               Caption         =   "04nivel04_04_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   7
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   8
            Begin VB.Menu m_04nivel04_05_10 
               Caption         =   "04nivel04_05_10"
               Index           =   0
            End
            Begin VB.Menu m_04nivel04_05_10 
               Caption         =   "04nivel04_05_10"
               Index           =   1
            End
            Begin VB.Menu m_04nivel04_05_10 
               Caption         =   "04nivel04_05_10"
               Index           =   2
            End
            Begin VB.Menu m_04nivel04_05_10 
               Caption         =   "04nivel04_05_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   9
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   10
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   11
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   12
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   13
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   14
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   15
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   16
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   17
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   18
         End
         Begin VB.Menu m_04nivel03_10 
            Caption         =   "04nivel03_10"
            Index           =   19
         End
      End
      Begin VB.Menu m_04nivel02 
         Caption         =   "04Nivel02"
         Index           =   19
      End
   End
   Begin VB.Menu m_05nivel01 
      Caption         =   "05Nivel01"
      Index           =   0
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   0
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_01 
               Caption         =   "05Nivel04_01_01"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_01 
               Caption         =   "05Nivel04_01_01"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_01 
               Caption         =   "05Nivel04_01_01"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_01 
               Caption         =   "05Nivel04_01_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_01 
               Caption         =   "05Nivel04_02_01"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_01 
               Caption         =   "05Nivel04_02_01"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_01 
               Caption         =   "05Nivel04_02_01"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_01 
               Caption         =   "05Nivel04_02_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_01 
               Caption         =   "05Nivel04_03_01"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_01 
               Caption         =   "05Nivel04_03_01"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_01 
               Caption         =   "05Nivel04_03_01"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_01 
               Caption         =   "05Nivel04_03_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_01 
               Caption         =   "05Nivel04_04_01"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_01 
               Caption         =   "05Nivel04_04_01"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_01 
               Caption         =   "05Nivel04_04_01"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_01 
               Caption         =   "05Nivel04_04_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_01 
               Caption         =   "05Nivel04_05_01"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_01 
               Caption         =   "05Nivel04_05_01"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_01 
               Caption         =   "05Nivel04_05_01"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_01 
               Caption         =   "05Nivel04_05_01"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_01 
            Caption         =   "05Nivel03_01"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   1
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   2
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_02 
               Caption         =   "05nivel04_01_02"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_02 
               Caption         =   "05nivel04_01_02"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_02 
               Caption         =   "05nivel04_01_02"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_02 
               Caption         =   "05nivel04_01_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_02 
               Caption         =   "05nivel04_02_02"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_02 
               Caption         =   "05nivel04_02_02"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_02 
               Caption         =   "05nivel04_02_02"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_02 
               Caption         =   "05nivel04_02_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_02 
               Caption         =   "05nivel04_03_02"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_02 
               Caption         =   "05nivel04_03_02"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_02 
               Caption         =   "05nivel04_03_02"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_02 
               Caption         =   "05nivel04_03_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_02 
               Caption         =   "05nivel04_04_02"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_02 
               Caption         =   "05nivel04_04_02"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_02 
               Caption         =   "05nivel04_04_02"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_02 
               Caption         =   "05nivel04_04_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_02 
               Caption         =   "05nivel04_05_02"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_02 
               Caption         =   "05nivel04_05_02"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_02 
               Caption         =   "05nivel04_05_02"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_02 
               Caption         =   "05nivel04_05_02"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_02 
            Caption         =   "05nivel03_02"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   3
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   4
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_03 
               Caption         =   "05nivel04_01_03"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_03 
               Caption         =   "05nivel04_01_03"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_03 
               Caption         =   "05nivel04_01_03"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_03 
               Caption         =   "05nivel04_01_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_03 
               Caption         =   "05nivel04_02_03"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_03 
               Caption         =   "05nivel04_02_03"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_03 
               Caption         =   "05nivel04_02_03"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_03 
               Caption         =   "05nivel04_02_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_03 
               Caption         =   "05nivel04_03_03"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_03 
               Caption         =   "05nivel04_03_03"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_03 
               Caption         =   "05nivel04_03_03"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_03 
               Caption         =   "05nivel04_03_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_03 
               Caption         =   "05nivel04_04_03"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_03 
               Caption         =   "05nivel04_04_03"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_03 
               Caption         =   "05nivel04_04_03"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_03 
               Caption         =   "05nivel04_04_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_03 
               Caption         =   "05nivel04_05_03"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_03 
               Caption         =   "05nivel04_05_03"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_03 
               Caption         =   "05nivel04_05_03"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_03 
               Caption         =   "05nivel04_05_03"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_03 
            Caption         =   "05nivel03_03"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   5
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   6
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_04 
               Caption         =   "05nivel04_01_04"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_04 
               Caption         =   "05nivel04_01_04"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_04 
               Caption         =   "05nivel04_01_04"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_04 
               Caption         =   "05nivel04_01_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_04 
               Caption         =   "05nivel04_02_04"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_04 
               Caption         =   "05nivel04_02_04"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_04 
               Caption         =   "05nivel04_02_04"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_04 
               Caption         =   "05nivel04_02_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_04 
               Caption         =   "05nivel04_03_04"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_04 
               Caption         =   "05nivel04_03_04"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_04 
               Caption         =   "05nivel04_03_04"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_04 
               Caption         =   "05nivel04_03_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_04 
               Caption         =   "05nivel04_04_04"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_04 
               Caption         =   "05nivel04_04_04"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_04 
               Caption         =   "05nivel04_04_04"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_04 
               Caption         =   "05nivel04_04_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_04 
               Caption         =   "05nivel04_05_04"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_04 
               Caption         =   "05nivel04_05_04"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_04 
               Caption         =   "05nivel04_05_04"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_04 
               Caption         =   "05nivel04_05_04"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_04 
            Caption         =   "05nivel03_04"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   7
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   8
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_05 
               Caption         =   "05nivel04_01_05"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_05 
               Caption         =   "05nivel04_01_05"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_05 
               Caption         =   "05nivel04_01_05"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_05 
               Caption         =   "05nivel04_01_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_05 
               Caption         =   "05nivel04_02_05"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_05 
               Caption         =   "05nivel04_02_05"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_05 
               Caption         =   "05nivel04_02_05"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_05 
               Caption         =   "05nivel04_02_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_05 
               Caption         =   "05nivel04_03_05"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_05 
               Caption         =   "05nivel04_03_05"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_05 
               Caption         =   "05nivel04_03_05"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_05 
               Caption         =   "05nivel04_03_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_05 
               Caption         =   "05nivel04_04_05"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_05 
               Caption         =   "05nivel04_04_05"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_05 
               Caption         =   "05nivel04_04_05"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_05 
               Caption         =   "05nivel04_04_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_05 
               Caption         =   "05nivel04_05_05"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_05 
               Caption         =   "05nivel04_05_05"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_05 
               Caption         =   "05nivel04_05_05"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_05 
               Caption         =   "05nivel04_05_05"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_05 
            Caption         =   "05nivel03_05"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   9
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   10
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_06 
               Caption         =   "05nivel04_01_06"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_06 
               Caption         =   "05nivel04_01_06"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_06 
               Caption         =   "05nivel04_01_06"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_06 
               Caption         =   "05nivel04_01_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_06 
               Caption         =   "05nivel04_02_06"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_06 
               Caption         =   "05nivel04_02_06"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_06 
               Caption         =   "05nivel04_02_06"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_06 
               Caption         =   "05nivel04_02_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_06 
               Caption         =   "05nivel04_03_06"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_06 
               Caption         =   "05nivel04_03_06"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_06 
               Caption         =   "05nivel04_03_06"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_06 
               Caption         =   "05nivel04_03_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_06 
               Caption         =   "05nivel04_04_06"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_06 
               Caption         =   "05nivel04_04_06"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_06 
               Caption         =   "05nivel04_04_06"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_06 
               Caption         =   "05nivel04_04_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_06 
               Caption         =   "05nivel04_05_06"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_06 
               Caption         =   "05nivel04_05_06"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_06 
               Caption         =   "05nivel04_05_06"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_06 
               Caption         =   "05nivel04_05_06"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_06 
            Caption         =   "05nivel03_06"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   11
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   12
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_07 
               Caption         =   "05nivel04_01_07"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_07 
               Caption         =   "05nivel04_01_07"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_07 
               Caption         =   "05nivel04_01_07"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_07 
               Caption         =   "05nivel04_01_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_07 
               Caption         =   "05nivel04_02_07"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_07 
               Caption         =   "05nivel04_02_07"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_07 
               Caption         =   "05nivel04_02_07"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_07 
               Caption         =   "05nivel04_02_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_07 
               Caption         =   "05nivel04_03_07"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_07 
               Caption         =   "05nivel04_03_07"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_07 
               Caption         =   "05nivel04_03_07"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_07 
               Caption         =   "05nivel04_03_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_07 
               Caption         =   "05nivel04_04_07"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_07 
               Caption         =   "05nivel04_04_07"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_07 
               Caption         =   "05nivel04_04_07"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_07 
               Caption         =   "05nivel04_04_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_07 
               Caption         =   "05nivel04_05_07"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_07 
               Caption         =   "05nivel04_05_07"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_07 
               Caption         =   "05nivel04_05_07"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_07 
               Caption         =   "05nivel04_05_07"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_07 
            Caption         =   "05nivel03_07"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   13
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   14
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_08 
               Caption         =   "05nivel04_01_08"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_08 
               Caption         =   "05nivel04_01_08"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_08 
               Caption         =   "05nivel04_01_08"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_08 
               Caption         =   "05nivel04_01_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_08 
               Caption         =   "05nivel04_02_08"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_08 
               Caption         =   "05nivel04_02_08"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_08 
               Caption         =   "05nivel04_02_08"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_08 
               Caption         =   "05nivel04_02_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_08 
               Caption         =   "05nivel04_03_08"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_08 
               Caption         =   "05nivel04_03_08"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_08 
               Caption         =   "05nivel04_03_08"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_08 
               Caption         =   "05nivel04_03_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_08 
               Caption         =   "05nivel04_04_08"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_08 
               Caption         =   "05nivel04_04_08"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_08 
               Caption         =   "05nivel04_04_08"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_08 
               Caption         =   "05nivel04_04_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_08 
               Caption         =   "05nivel04_05_08"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_08 
               Caption         =   "05nivel04_05_08"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_08 
               Caption         =   "05nivel04_05_08"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_08 
               Caption         =   "05nivel04_05_08"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_08 
            Caption         =   "05nivel03_08"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   15
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   16
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_09 
               Caption         =   "05nivel04_01_09"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_09 
               Caption         =   "05nivel04_01_09"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_09 
               Caption         =   "05nivel04_01_09"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_09 
               Caption         =   "05nivel04_01_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_09 
               Caption         =   "05nivel04_02_09"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_09 
               Caption         =   "05nivel04_02_09"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_09 
               Caption         =   "05nivel04_02_09"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_09 
               Caption         =   "05nivel04_02_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_09 
               Caption         =   "05nivel04_03_09"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_09 
               Caption         =   "05nivel04_03_09"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_09 
               Caption         =   "05nivel04_03_09"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_09 
               Caption         =   "05nivel04_03_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_09 
               Caption         =   "05nivel04_04_09"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_09 
               Caption         =   "05nivel04_04_09"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_09 
               Caption         =   "05nivel04_04_09"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_09 
               Caption         =   "05nivel04_04_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_09 
               Caption         =   "05nivel04_05_09"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_09 
               Caption         =   "05nivel04_05_09"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_09 
               Caption         =   "05nivel04_05_09"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_09 
               Caption         =   "05nivel04_05_09"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_09 
            Caption         =   "05nivel03_09"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   17
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   18
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   0
            Begin VB.Menu m_05nivel04_01_10 
               Caption         =   "05nivel04_01_10"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_01_10 
               Caption         =   "05nivel04_01_10"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_01_10 
               Caption         =   "05nivel04_01_10"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_01_10 
               Caption         =   "05nivel04_01_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   1
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   2
            Begin VB.Menu m_05nivel04_02_10 
               Caption         =   "05nivel04_02_10"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_02_10 
               Caption         =   "05nivel04_02_10"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_02_10 
               Caption         =   "05nivel04_02_10"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_02_10 
               Caption         =   "05nivel04_02_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   3
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   4
            Begin VB.Menu m_05nivel04_03_10 
               Caption         =   "05nivel04_03_10"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_03_10 
               Caption         =   "05nivel04_03_10"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_03_10 
               Caption         =   "05nivel04_03_10"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_03_10 
               Caption         =   "05nivel04_03_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   5
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   6
            Begin VB.Menu m_05nivel04_04_10 
               Caption         =   "05nivel04_04_10"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_04_10 
               Caption         =   "05nivel04_04_10"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_04_10 
               Caption         =   "05nivel04_04_10"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_04_10 
               Caption         =   "05nivel04_04_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   7
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   8
            Begin VB.Menu m_05nivel04_05_10 
               Caption         =   "05nivel04_05_10"
               Index           =   0
            End
            Begin VB.Menu m_05nivel04_05_10 
               Caption         =   "05nivel04_05_10"
               Index           =   1
            End
            Begin VB.Menu m_05nivel04_05_10 
               Caption         =   "05nivel04_05_10"
               Index           =   2
            End
            Begin VB.Menu m_05nivel04_05_10 
               Caption         =   "05nivel04_05_10"
               Index           =   3
            End
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   9
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   10
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   11
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   12
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   13
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   14
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   15
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   16
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   17
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   18
         End
         Begin VB.Menu m_05nivel03_10 
            Caption         =   "05nivel03_10"
            Index           =   19
         End
      End
      Begin VB.Menu m_05nivel02 
         Caption         =   "05Nivel02"
         Index           =   19
      End
   End
   Begin VB.Menu m_06nivel01 
      Caption         =   "06Nivel01"
      Index           =   0
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   0
         Begin VB.Menu m_06nivel03_01 
            Caption         =   "06Nivel03_01"
            Index           =   0
         End
         Begin VB.Menu m_06nivel03_01 
            Caption         =   "06Nivel03_01"
            Index           =   1
         End
         Begin VB.Menu m_06nivel03_01 
            Caption         =   "06Nivel03_01"
            Index           =   2
         End
         Begin VB.Menu m_06nivel03_01 
            Caption         =   "06Nivel03_01"
            Index           =   3
         End
         Begin VB.Menu m_06nivel03_01 
            Caption         =   "06Nivel03_01"
            Index           =   4
         End
         Begin VB.Menu m_06nivel03_01 
            Caption         =   "06Nivel03_01"
            Index           =   5
         End
         Begin VB.Menu m_06nivel03_01 
            Caption         =   "06Nivel03_01"
            Index           =   6
         End
         Begin VB.Menu m_06nivel03_01 
            Caption         =   "06Nivel03_01"
            Index           =   7
         End
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   1
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   2
         Begin VB.Menu m_06nivel03_02 
            Caption         =   "06nivel03_02"
            Index           =   0
         End
         Begin VB.Menu m_06nivel03_02 
            Caption         =   "06nivel03_02"
            Index           =   1
         End
         Begin VB.Menu m_06nivel03_02 
            Caption         =   "06nivel03_02"
            Index           =   2
         End
         Begin VB.Menu m_06nivel03_02 
            Caption         =   "06nivel03_02"
            Index           =   3
         End
         Begin VB.Menu m_06nivel03_02 
            Caption         =   "06nivel03_02"
            Index           =   4
         End
         Begin VB.Menu m_06nivel03_02 
            Caption         =   "06nivel03_02"
            Index           =   5
         End
         Begin VB.Menu m_06nivel03_02 
            Caption         =   "06nivel03_02"
            Index           =   6
         End
         Begin VB.Menu m_06nivel03_02 
            Caption         =   "06nivel03_02"
            Index           =   7
         End
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   3
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   4
         Begin VB.Menu m_06nivel03_03 
            Caption         =   "06nivel03_03"
            Index           =   0
         End
         Begin VB.Menu m_06nivel03_03 
            Caption         =   "06nivel03_03"
            Index           =   1
         End
         Begin VB.Menu m_06nivel03_03 
            Caption         =   "06nivel03_03"
            Index           =   2
         End
         Begin VB.Menu m_06nivel03_03 
            Caption         =   "06nivel03_03"
            Index           =   3
         End
         Begin VB.Menu m_06nivel03_03 
            Caption         =   "06nivel03_03"
            Index           =   4
         End
         Begin VB.Menu m_06nivel03_03 
            Caption         =   "06nivel03_03"
            Index           =   5
         End
         Begin VB.Menu m_06nivel03_03 
            Caption         =   "06nivel03_03"
            Index           =   6
         End
         Begin VB.Menu m_06nivel03_03 
            Caption         =   "06nivel03_03"
            Index           =   7
         End
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   5
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   6
         Begin VB.Menu m_06nivel03_04 
            Caption         =   "06nivel03_04"
            Index           =   0
         End
         Begin VB.Menu m_06nivel03_04 
            Caption         =   "06nivel03_04"
            Index           =   1
         End
         Begin VB.Menu m_06nivel03_04 
            Caption         =   "06nivel03_04"
            Index           =   2
         End
         Begin VB.Menu m_06nivel03_04 
            Caption         =   "06nivel03_04"
            Index           =   3
         End
         Begin VB.Menu m_06nivel03_04 
            Caption         =   "06nivel03_04"
            Index           =   4
         End
         Begin VB.Menu m_06nivel03_04 
            Caption         =   "06nivel03_04"
            Index           =   5
         End
         Begin VB.Menu m_06nivel03_04 
            Caption         =   "06nivel03_04"
            Index           =   6
         End
         Begin VB.Menu m_06nivel03_04 
            Caption         =   "06nivel03_04"
            Index           =   7
         End
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   7
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   8
         Begin VB.Menu m_06nivel03_05 
            Caption         =   "06nivel03_05"
            Index           =   0
         End
         Begin VB.Menu m_06nivel03_05 
            Caption         =   "06nivel03_05"
            Index           =   1
         End
         Begin VB.Menu m_06nivel03_05 
            Caption         =   "06nivel03_05"
            Index           =   2
         End
         Begin VB.Menu m_06nivel03_05 
            Caption         =   "06nivel03_05"
            Index           =   3
         End
         Begin VB.Menu m_06nivel03_05 
            Caption         =   "06nivel03_05"
            Index           =   4
         End
         Begin VB.Menu m_06nivel03_05 
            Caption         =   "06nivel03_05"
            Index           =   5
         End
         Begin VB.Menu m_06nivel03_05 
            Caption         =   "06nivel03_05"
            Index           =   6
         End
         Begin VB.Menu m_06nivel03_05 
            Caption         =   "06nivel03_05"
            Index           =   7
         End
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   9
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   10
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   11
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   12
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   13
      End
      Begin VB.Menu m_06nivel02 
         Caption         =   "06Nivel02"
         Index           =   14
      End
   End
   Begin VB.Menu m_07nivel01 
      Caption         =   "07Nivel01"
      Index           =   0
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   0
         Begin VB.Menu m_07nivel03_01 
            Caption         =   "07Nivel03_01"
            Index           =   0
         End
         Begin VB.Menu m_07nivel03_01 
            Caption         =   "07Nivel03_01"
            Index           =   1
         End
         Begin VB.Menu m_07nivel03_01 
            Caption         =   "07Nivel03_01"
            Index           =   2
         End
         Begin VB.Menu m_07nivel03_01 
            Caption         =   "07Nivel03_01"
            Index           =   3
         End
         Begin VB.Menu m_07nivel03_01 
            Caption         =   "07Nivel03_01"
            Index           =   4
         End
         Begin VB.Menu m_07nivel03_01 
            Caption         =   "07Nivel03_01"
            Index           =   5
         End
         Begin VB.Menu m_07nivel03_01 
            Caption         =   "07Nivel03_01"
            Index           =   6
         End
         Begin VB.Menu m_07nivel03_01 
            Caption         =   "07Nivel03_01"
            Index           =   7
         End
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   1
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   2
         Begin VB.Menu m_07nivel03_02 
            Caption         =   "07nivel03_02"
            Index           =   0
         End
         Begin VB.Menu m_07nivel03_02 
            Caption         =   "07nivel03_02"
            Index           =   1
         End
         Begin VB.Menu m_07nivel03_02 
            Caption         =   "07nivel03_02"
            Index           =   2
         End
         Begin VB.Menu m_07nivel03_02 
            Caption         =   "07nivel03_02"
            Index           =   3
         End
         Begin VB.Menu m_07nivel03_02 
            Caption         =   "07nivel03_02"
            Index           =   4
         End
         Begin VB.Menu m_07nivel03_02 
            Caption         =   "07nivel03_02"
            Index           =   5
         End
         Begin VB.Menu m_07nivel03_02 
            Caption         =   "07nivel03_02"
            Index           =   6
         End
         Begin VB.Menu m_07nivel03_02 
            Caption         =   "07nivel03_02"
            Index           =   7
         End
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   3
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   4
         Begin VB.Menu m_07nivel03_03 
            Caption         =   "07nivel03_03"
            Index           =   0
         End
         Begin VB.Menu m_07nivel03_03 
            Caption         =   "07nivel03_03"
            Index           =   1
         End
         Begin VB.Menu m_07nivel03_03 
            Caption         =   "07nivel03_03"
            Index           =   2
         End
         Begin VB.Menu m_07nivel03_03 
            Caption         =   "07nivel03_03"
            Index           =   3
         End
         Begin VB.Menu m_07nivel03_03 
            Caption         =   "07nivel03_03"
            Index           =   4
         End
         Begin VB.Menu m_07nivel03_03 
            Caption         =   "07nivel03_03"
            Index           =   5
         End
         Begin VB.Menu m_07nivel03_03 
            Caption         =   "07nivel03_03"
            Index           =   6
         End
         Begin VB.Menu m_07nivel03_03 
            Caption         =   "07nivel03_03"
            Index           =   7
         End
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   5
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   6
         Begin VB.Menu m_07nivel03_04 
            Caption         =   "07nivel03_04"
            Index           =   0
         End
         Begin VB.Menu m_07nivel03_04 
            Caption         =   "07nivel03_04"
            Index           =   1
         End
         Begin VB.Menu m_07nivel03_04 
            Caption         =   "07nivel03_04"
            Index           =   2
         End
         Begin VB.Menu m_07nivel03_04 
            Caption         =   "07nivel03_04"
            Index           =   3
         End
         Begin VB.Menu m_07nivel03_04 
            Caption         =   "07nivel03_04"
            Index           =   4
         End
         Begin VB.Menu m_07nivel03_04 
            Caption         =   "07nivel03_04"
            Index           =   5
         End
         Begin VB.Menu m_07nivel03_04 
            Caption         =   "07nivel03_04"
            Index           =   6
         End
         Begin VB.Menu m_07nivel03_04 
            Caption         =   "07nivel03_04"
            Index           =   7
         End
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   7
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   8
         Begin VB.Menu m_07nivel03_05 
            Caption         =   "07nivel03_05"
            Index           =   0
         End
         Begin VB.Menu m_07nivel03_05 
            Caption         =   "07nivel03_05"
            Index           =   1
         End
         Begin VB.Menu m_07nivel03_05 
            Caption         =   "07nivel03_05"
            Index           =   2
         End
         Begin VB.Menu m_07nivel03_05 
            Caption         =   "07nivel03_05"
            Index           =   3
         End
         Begin VB.Menu m_07nivel03_05 
            Caption         =   "07nivel03_05"
            Index           =   4
         End
         Begin VB.Menu m_07nivel03_05 
            Caption         =   "07nivel03_05"
            Index           =   5
         End
         Begin VB.Menu m_07nivel03_05 
            Caption         =   "07nivel03_05"
            Index           =   6
         End
         Begin VB.Menu m_07nivel03_05 
            Caption         =   "07nivel03_05"
            Index           =   7
         End
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   9
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   10
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   11
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   12
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   13
      End
      Begin VB.Menu m_07nivel02 
         Caption         =   "07Nivel02"
         Index           =   14
      End
   End
   Begin VB.Menu m_08nivel01 
      Caption         =   "08Nivel01"
      Index           =   0
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   0
         Begin VB.Menu m_08nivel03_01 
            Caption         =   "08Nivel03_01"
            Index           =   0
         End
         Begin VB.Menu m_08nivel03_01 
            Caption         =   "08Nivel03_01"
            Index           =   1
         End
         Begin VB.Menu m_08nivel03_01 
            Caption         =   "08Nivel03_01"
            Index           =   2
         End
         Begin VB.Menu m_08nivel03_01 
            Caption         =   "08Nivel03_01"
            Index           =   3
         End
         Begin VB.Menu m_08nivel03_01 
            Caption         =   "08Nivel03_01"
            Index           =   4
         End
         Begin VB.Menu m_08nivel03_01 
            Caption         =   "08Nivel03_01"
            Index           =   5
         End
         Begin VB.Menu m_08nivel03_01 
            Caption         =   "08Nivel03_01"
            Index           =   6
         End
         Begin VB.Menu m_08nivel03_01 
            Caption         =   "08Nivel03_01"
            Index           =   7
         End
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   1
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   2
         Begin VB.Menu m_08nivel03_02 
            Caption         =   "08nivel03_02"
            Index           =   0
         End
         Begin VB.Menu m_08nivel03_02 
            Caption         =   "08nivel03_02"
            Index           =   1
         End
         Begin VB.Menu m_08nivel03_02 
            Caption         =   "08nivel03_02"
            Index           =   2
         End
         Begin VB.Menu m_08nivel03_02 
            Caption         =   "08nivel03_02"
            Index           =   3
         End
         Begin VB.Menu m_08nivel03_02 
            Caption         =   "08nivel03_02"
            Index           =   4
         End
         Begin VB.Menu m_08nivel03_02 
            Caption         =   "08nivel03_02"
            Index           =   5
         End
         Begin VB.Menu m_08nivel03_02 
            Caption         =   "08nivel03_02"
            Index           =   6
         End
         Begin VB.Menu m_08nivel03_02 
            Caption         =   "08nivel03_02"
            Index           =   7
         End
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   3
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   4
         Begin VB.Menu m_08nivel03_03 
            Caption         =   "08nivel03_03"
            Index           =   0
         End
         Begin VB.Menu m_08nivel03_03 
            Caption         =   "08nivel03_03"
            Index           =   1
         End
         Begin VB.Menu m_08nivel03_03 
            Caption         =   "08nivel03_03"
            Index           =   2
         End
         Begin VB.Menu m_08nivel03_03 
            Caption         =   "08nivel03_03"
            Index           =   3
         End
         Begin VB.Menu m_08nivel03_03 
            Caption         =   "08nivel03_03"
            Index           =   4
         End
         Begin VB.Menu m_08nivel03_03 
            Caption         =   "08nivel03_03"
            Index           =   5
         End
         Begin VB.Menu m_08nivel03_03 
            Caption         =   "08nivel03_03"
            Index           =   6
         End
         Begin VB.Menu m_08nivel03_03 
            Caption         =   "08nivel03_03"
            Index           =   7
         End
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   5
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   6
         Begin VB.Menu m_08nivel03_04 
            Caption         =   "08nivel03_04"
            Index           =   0
         End
         Begin VB.Menu m_08nivel03_04 
            Caption         =   "08nivel03_04"
            Index           =   1
         End
         Begin VB.Menu m_08nivel03_04 
            Caption         =   "08nivel03_04"
            Index           =   2
         End
         Begin VB.Menu m_08nivel03_04 
            Caption         =   "08nivel03_04"
            Index           =   3
         End
         Begin VB.Menu m_08nivel03_04 
            Caption         =   "08nivel03_04"
            Index           =   4
         End
         Begin VB.Menu m_08nivel03_04 
            Caption         =   "08nivel03_04"
            Index           =   5
         End
         Begin VB.Menu m_08nivel03_04 
            Caption         =   "08nivel03_04"
            Index           =   6
         End
         Begin VB.Menu m_08nivel03_04 
            Caption         =   "08nivel03_04"
            Index           =   7
         End
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   7
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   8
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   9
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   10
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   11
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   12
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   13
      End
      Begin VB.Menu m_08nivel02 
         Caption         =   "08Nivel02"
         Index           =   14
      End
   End
   Begin VB.Menu m_Meses 
      Caption         =   "&Meses"
      Begin VB.Menu m_mes 
         Caption         =   "  A P E R T U R A "
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu m_mes 
         Caption         =   "  E N E R O"
         Index           =   1
      End
      Begin VB.Menu m_mes 
         Caption         =   "  F E B R E R O"
         Index           =   2
      End
      Begin VB.Menu m_mes 
         Caption         =   "  M A R Z O"
         Index           =   3
      End
      Begin VB.Menu m_mes 
         Caption         =   "  A B R I L"
         Index           =   4
      End
      Begin VB.Menu m_mes 
         Caption         =   "  M A Y O "
         Index           =   5
      End
      Begin VB.Menu m_mes 
         Caption         =   "  J U N I O"
         Index           =   6
      End
      Begin VB.Menu m_mes 
         Caption         =   "  J U L I O"
         Index           =   7
      End
      Begin VB.Menu m_mes 
         Caption         =   "  A G O S T O"
         Index           =   8
      End
      Begin VB.Menu m_mes 
         Caption         =   "  S E P T I E M B R E"
         Index           =   9
      End
      Begin VB.Menu m_mes 
         Caption         =   "  O C T U B R E"
         Index           =   10
      End
      Begin VB.Menu m_mes 
         Caption         =   "  N O V I E M B R E"
         Index           =   11
      End
      Begin VB.Menu m_mes 
         Caption         =   "  D I C I E M B R E"
         Index           =   12
      End
   End
End
Attribute VB_Name = "mdiInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Dim ActivaMenuEmergente As Boolean
Private Sub m_Maestro_Click(Index As Integer)
    Select Case Index
        Case 0: ConfigurarFormulario TIPO_FORMULARIO.FORM_MAESTRO_ESTADO
        Case 1: ConfigurarFormulario TIPO_FORMULARIO.FORM_MAESTRO_TIP_DOC
        Case 2: ConfigurarFormulario TIPO_FORMULARIO.FORM_MAESTRO_TIPO_CAMBIO
        Case 3: FrmAuxiliares.Show
    End Select
End Sub
Private Sub imgRegresar_Click()
    imgRegresar.Visible = False
    InicializarMenu True
    MuestraPrimeros
    InicializaCaptionMenu
    InicializarMenu False
    CargarMenu 0
    mdiInicio.ConfigBarraHerramientas 0
    mdiInicio.sbPrincipal.Panels(3) = "Sistema Integral-Principal"
    mdiInicio.sbPrincipal.Panels(3).Picture = mdiInicio.imgsBarraHerramientas.ListImages(27).Picture
    mdiInicio.tbPrincipal.Refresh
    OcultaPrimeros
    ModActivo = 0
End Sub
Private Sub lblControlDoc_Click(Index As Integer)
    
    EjecutaMenu 1, "01", "02", "00", "00"
End Sub
Private Sub lblCtasxCobrar_Click(Index As Integer)
    
    EjecutaMenu 7, "01", "02", "00", "00"
End Sub
Private Sub lblCtasxPagar_Click(Index As Integer)
     
     EjecutaMenu 9, "01", "02", "00", "00"
End Sub
Private Sub lblFacturacion_Click(Index As Integer)
    
    EjecutaMenu 5, "01", "02", "00", "00"
End Sub

Private Sub lblPlanilla_Click(Index As Integer)
    
    EjecutaMenu 11, "01", "02", "00", "00"
End Sub

Private Sub lblRecHum_Click(Index As Integer)
    EjecutaMenu 3, "01", "02", "00", "00"
End Sub
Private Sub m_01nivel02_Click(Index As Integer)
    EjecutaMenu Index, "01", "02", "00", "00"
End Sub
Private Sub m_01nivel03_01_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "01", "00"
End Sub
Private Sub m_01nivel03_02_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "02", "00"
End Sub
Private Sub m_01nivel03_03_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "03", "00"
End Sub
Private Sub m_01nivel03_04_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "04", "00"
End Sub
Private Sub m_01nivel03_06_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "06", "00"
End Sub
Private Sub m_01nivel03_07_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "07", "00"
End Sub
Private Sub m_01nivel03_08_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "08", "00"
End Sub
Private Sub m_01nivel03_09_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "09", "00"
End Sub
Private Sub m_01nivel03_10_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "10", "00"
End Sub
Private Sub m_01nivel03_11_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "11", "00"
End Sub
Private Sub m_01nivel03_12_Click(Index As Integer)
    EjecutaMenu Index, "01", "03", "12", "00"
End Sub

Private Sub m_02nivel02_Click(Index As Integer)
    EjecutaMenu Index, "02", "02", "00", "00"
End Sub
Private Sub m_02nivel03_01_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "01", "00"
End Sub
Private Sub m_02nivel03_02_Click(Index As Integer)
   EjecutaMenu Index, "02", "03", "02", "00"
End Sub
Private Sub m_02nivel03_03_Click(Index As Integer)
   EjecutaMenu Index, "02", "03", "03", "00"
End Sub
Private Sub m_02nivel03_04_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "04", "00"
End Sub
Private Sub m_02nivel03_05_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "05", "00"
End Sub
Private Sub m_02nivel03_06_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "06", "00"
End Sub
Private Sub m_02nivel03_07_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "07", "00"
End Sub
Private Sub m_02nivel03_08_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "08", "00"
End Sub
Private Sub m_02nivel03_09_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "09", "00"
End Sub
Private Sub m_02nivel03_10_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "10", "00"
End Sub
Private Sub m_02nivel03_11_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "11", "00"
End Sub
Private Sub m_02nivel03_12_Click(Index As Integer)
    EjecutaMenu Index, "02", "03", "12", "00"
End Sub

Private Sub m_03nivel02_Click(Index As Integer)
    EjecutaMenu Index, "03", "02", "00", "00"
End Sub
Private Sub m_03nivel03_01_Click(Index As Integer)
   EjecutaMenu Index, "03", "03", "01", "00"
End Sub
Private Sub m_03nivel03_02_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "02", "00"
End Sub
Private Sub m_03nivel03_03_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "03", "00"
End Sub
Private Sub m_03nivel03_04_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "04", "00"
End Sub
Private Sub m_03nivel03_05_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "05", "00"
End Sub
Private Sub m_03nivel03_06_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "06", "00"
End Sub
Private Sub m_03nivel03_07_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "07", "00"
End Sub
Private Sub m_03nivel03_08_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "08", "00"
End Sub
Private Sub m_03nivel03_09_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "09", "00"
End Sub
Private Sub m_03nivel03_10_Click(Index As Integer)
    EjecutaMenu Index, "03", "03", "10", "00"
End Sub
Private Sub m_03nivel04_01_01_Click(Index As Integer)
    EjecutaMenu Index, "03", "02", "01", "01"
End Sub
Private Sub m_03nivel04_01_02_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "01", "02"
End Sub
Private Sub m_03nivel04_01_03_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "01", "03"
End Sub
Private Sub m_03nivel04_02_01_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "02", "01"
End Sub
Private Sub m_03nivel04_02_02_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "02", "02"
End Sub
Private Sub m_03nivel04_02_06_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "02", "06"
End Sub
Private Sub m_03nivel04_03_01_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "03", "01"
End Sub
Private Sub m_03nivel04_03_02_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "03", "02"
End Sub
Private Sub m_03nivel04_04_01_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "04", "01"
End Sub
Private Sub m_03nivel04_04_02_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "04", "02"
End Sub
Private Sub m_03nivel04_04_03_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "04", "03"
End Sub
Private Sub m_03nivel04_05_01_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "05", "01"
End Sub
Private Sub m_03nivel04_05_02_Click(Index As Integer)
    EjecutaMenu Index, "03", "04", "05", "02"
End Sub
Private Sub m_04nivel02_Click(Index As Integer)
    EjecutaMenu Index, "04", "02", "00", "00"
End Sub
Private Sub m_04nivel03_01_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "01", "00"
End Sub
Private Sub m_04nivel03_02_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "02", "00"
End Sub
Private Sub m_04nivel03_03_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "03", "00"
End Sub
Private Sub m_04nivel03_05_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "05", "00"
End Sub
Private Sub m_04nivel03_06_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "06", "00"
End Sub
Private Sub m_04nivel03_07_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "07", "00"
End Sub
Private Sub m_04nivel03_08_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "08", "00"
End Sub
Private Sub m_04nivel03_09_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "09", "00"
End Sub
Private Sub m_04nivel03_10_Click(Index As Integer)
    EjecutaMenu Index, "04", "03", "10", "00"
End Sub
Private Sub m_04nivel04_01_01_Click(Index As Integer)
    EjecutaMenu Index, "04", "04", "01", "01"
End Sub
Private Sub m_04nivel04_01_07_Click(Index As Integer)
    EjecutaMenu Index, "04", "04", "01", "07"
End Sub
Private Sub m_05nivel02_Click(Index As Integer)
    EjecutaMenu Index, "05", "02", "00", "00"
End Sub
Private Sub m_05nivel03_01_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "01", "00"
End Sub
Private Sub m_05nivel03_02_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "02", "00"
End Sub
Private Sub m_05nivel03_03_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "03", "00"
End Sub
Private Sub m_05nivel03_04_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "04", "00"
End Sub
Private Sub m_05nivel03_05_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "05", "00"
End Sub
Private Sub m_05nivel03_06_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "06", "00"
End Sub
Private Sub m_05nivel03_07_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "07", "00"
End Sub
Private Sub m_05nivel03_08_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "08", "00"
End Sub
Private Sub m_05nivel03_09_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "09", "00"
End Sub
Private Sub m_05nivel03_10_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "10", "00"
End Sub

Private Sub m_05nivel03_11_Click(Index As Integer)
    EjecutaMenu Index, "05", "03", "11", "00"
End Sub

Private Sub m_05nivel04_01_01_Click(Index As Integer)
    EjecutaMenu Index, "05", "04", "01", "01"
End Sub
Private Sub m_06nivel02_Click(Index As Integer)
    EjecutaMenu Index, "06", "02", "00", "00"
End Sub
Private Sub m_06nivel03_01_Click(Index As Integer)
    EjecutaMenu Index, "06", "03", "01", "00"
End Sub
Private Sub m_06nivel03_02_Click(Index As Integer)
    EjecutaMenu Index, "06", "03", "02", "00"
End Sub
Private Sub m_06nivel03_03_Click(Index As Integer)
    EjecutaMenu Index, "06", "03", "03", "00"
End Sub
Private Sub m_06nivel03_04_Click(Index As Integer)
    EjecutaMenu Index, "06", "03", "04", "00"
End Sub
Private Sub m_06nivel03_05_Click(Index As Integer)
    EjecutaMenu Index, "06", "03", "05", "00"
End Sub
Private Sub m_07nivel02_Click(Index As Integer)
    EjecutaMenu Index, "07", "02", "00", "00"
End Sub
Private Sub m_07nivel03_01_Click(Index As Integer)
    EjecutaMenu Index, "07", "03", "01", "00"
End Sub
Private Sub m_08nivel02_Click(Index As Integer)
    EjecutaMenu Index, "08", "02", "00", "00"
End Sub
Private Sub m_08nivel03_01_Click(Index As Integer)
    EjecutaMenu Index, "08", "03", "01", "00"
End Sub
Private Sub m_mes_Click(Index As Integer)
    Select Case Index
        Case 1
            strMesSistema = "01"
            sbPrincipal.Panels(5).Text = NombreMes("01", False)
        Case 2
            strMesSistema = "02"
            sbPrincipal.Panels(5).Text = NombreMes("02", False)
        Case 3
            strMesSistema = "03"
            sbPrincipal.Panels(5).Text = NombreMes("03", False)
        Case 4
            strMesSistema = "04"
            sbPrincipal.Panels(5).Text = NombreMes("04", False)
        Case 5
            strMesSistema = "05"
            sbPrincipal.Panels(5).Text = NombreMes("05", False)
        Case 6
            strMesSistema = "06"
            sbPrincipal.Panels(5).Text = NombreMes("06", False)
        Case 7
            strMesSistema = "07"
            sbPrincipal.Panels(5).Text = NombreMes("07", False)
        Case 8
            strMesSistema = "08"
            sbPrincipal.Panels(5).Text = NombreMes("08", False)
        Case 9
            strMesSistema = "09"
            sbPrincipal.Panels(5).Text = NombreMes("09", False)
        Case 10
            strMesSistema = "10"
            sbPrincipal.Panels(5).Text = NombreMes("10", False)
        Case 11
            strMesSistema = "11"
            sbPrincipal.Panels(5).Text = NombreMes("11", False)
        Case 12
            strMesSistema = "12"
            sbPrincipal.Panels(5).Text = NombreMes("12", False)
    End Select
End Sub
Private Sub MDIForm_Activate()
    'Picture1.Visible = True
    InicializarMenu False  'Oculta todos los Menus
    ModActivo = 0
    ConfigBarraHerramientas ModActivo
    InicializarMenu True
    MuestraPrimeros
    InicializaCaptionMenu
    InicializarMenu False
    CargarMenu 0     'Carga menu del modulo 0
    OcultaPrimeros
    Dim I%
    Dim hMenu, hSubMenu, menuID, X
    ActivaMenuEmergente = False
    m_Meses.Visible = False
    hMenu = GetMenu(hwnd)
    hSubMenu = GetSubMenu(hMenu, 8)
    If strUsuarioId = "ADM" Then
        tbPrincipal.Enabled = True
    Else
        tbPrincipal.Enabled = False
    End If
    FoliosAutomaticos
    RegTotTrab
    Me.SetFocus
    
    
End Sub

Private Sub MDIForm_Load()
    Picture1.Visible = True
    Dim Resolucion As Double, imagen As String
    With Screen
        Resolucion = (.Width \ .TwipsPerPixelX)
    End With
    With sbPrincipal
        If Resolucion = 800 Or Resolucion = 600 Then
            'Picture1.Picture = Picture2.Picture
            'Me.Picture = Picture2.Picture
            .Panels(1).Width = 3500.22
            .Panels(2).Width = 1200.18
            .Panels(3).Width = 2500.15
            .Panels(4).Width = 799.93
            .Panels(5).Width = 1000.06
            .Panels(6).Width = 1000.06
            .Panels(7).Width = 1800
            .Font.Size = 6.75
            .Font.Bold = False
            imgRegresar.Left = 10890
            imgRegresar.Top = 7050
            lblPlanilla(0).Left = 7560
            lblPlanilla(0).Top = 6960
            lblPlanilla(1).Left = 7530
            lblPlanilla(1).Top = 6930
            lblCtasxPagar(0).Left = 3240
            lblCtasxPagar(0).Top = 6330
            lblCtasxPagar(1).Left = 3210
            lblCtasxPagar(1).Top = 6300
            lblCtasxCobrar(0).Left = 2670
            lblCtasxCobrar(0).Top = 4950
            lblCtasxCobrar(1).Left = 2640
            lblCtasxCobrar(1).Top = 4920
            lblRecHum(0).Left = 6060
            lblRecHum(0).Top = 4020
            lblRecHum(1).Left = 6030
            lblRecHum(1).Top = 3990
            lblControlDoc(0).Left = 8640
            lblControlDoc(0).Top = 3930
            lblControlDoc(1).Left = 8670
            lblControlDoc(1).Top = 3960
        Else
            'Picture1.Picture = Picture3.Picture
            'Me.Picture = Picture3.Picture
            .Panels(1).Width = 4500.28
            .Panels(2).Width = 1500.09
            .Panels(3).Width = 3000.18
            .Panels(4).Width = 1200.75
            .Panels(5).Width = 1500.09
            .Panels(6).Width = 1599.87
            .Panels(7).Width = 1800
            .Font.Size = 8.25
            .Font.Bold = True
            imgRegresar.Left = 13950
            imgRegresar.Top = 9150
            

        End If
    End With
    
    Me.Caption = gsNomSW & " - Sistema Administrativo " & gsVersion
    
End Sub


Private Sub MDIForm_Resize()
    On Error GoTo SERROR
    
        Picture1.Height = Me.Height
    Exit Sub
SERROR:
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim SQL As String
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    SQL = "Select * from 7cia_user where codcia='" & strCiaId & "' and usuario_id='" & strUsuarioId & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    If Rs.RecordCount <> 0 Then
        SQL = "Call Update_CiaUser ('" & strCiaId & "','" & _
                                         strUsuarioId & "','" & _
                                         strMesSistema & "','" & _
                                         strAnoSistema & "','" & _
                                         Rs.Fields("anomes_bloq") & "','" & _
                                         Rs.Fields("fec_conexion") & "','" & _
                                         Format(Date, "yyyy/mm/dd") & "','" & _
                                         Rs.Fields("hor_conexion") & "','" & _
                                         Format(Time, "hh:mm:ss") & "','" & _
                                         Rs.Fields("host_conexion") & "',0)"
        Rs.MoveNext
    End If
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
    Set oMenu = Nothing
    Rs.CloseRecordset
    Set Rs = Nothing
End Sub





Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbNormal
    If ModActivo = 0 Then
        lblControlDoc(0).Visible = True
        lblControlDoc(1).Visible = True
        lblFacturacion(0).Visible = True
        lblFacturacion(1).Visible = True
        lblRecHum(0).Visible = True
        lblRecHum(1).Visible = True
        lblCtasxCobrar(0).Visible = True
        lblCtasxCobrar(1).Visible = True
        lblCtasxPagar(0).Visible = True
        lblCtasxPagar(1).Visible = True
        imgRegresar.Visible = False
    End If
End Sub
Public Sub ConfigBarraHerramientas(Modulo As Integer)
    Dim I As Integer
    With tbPrincipal
        For I = 1 To 20
            .Buttons(I).Visible = False
            .Buttons(I).Style = tbrDefault
        Next
        Select Case Modulo
            Case 0
                .Buttons(1).Visible = True
                .Buttons(1).Image = 11
                .Buttons(1).ToolTipText = "Control Documentario"
                .Buttons(2).Visible = True
                .Buttons(2).Image = 4
                .Buttons(2).ToolTipText = "Recursos Humanos"
                .Buttons(3).Visible = True
                .Buttons(3).Image = 36
                .Buttons(3).ToolTipText = "Facturación"
                .Buttons(4).Visible = True
                .Buttons(4).Image = 2
                .Buttons(4).ToolTipText = "Cuentas por Cobrar"
                .Buttons(5).Visible = True
                .Buttons(5).Image = 3
                .Buttons(5).ToolTipText = "Cuentas por Pagar"
                .Buttons(6).Visible = True
                .Buttons(6).Image = 7
                .Buttons(6).ToolTipText = "Tesorería"
                .Buttons(7).Visible = True
                .Buttons(7).Image = 9
                .Buttons(7).ToolTipText = "Salir del Sistema"
                .Buttons(8).Style = tbrSeparator
                .Buttons(8).Visible = True
                .Buttons(9).Image = 10
                .Buttons(9).ToolTipText = "Mantenimiento de Usuarios"
                .Buttons(9).Visible = True
                .Buttons(10).Image = 13
                .Buttons(10).ToolTipText = "Mantenimiento de Empresas"
                .Buttons(10).Visible = True
                .Buttons(11).Style = tbrSeparator
                .Buttons(11).Visible = True
                .Buttons(12).Image = 12
                .Buttons(12).ToolTipText = "Ayuda"
                .Buttons(12).Visible = True
            Case 1
                .Buttons(1).Visible = True
                .Buttons(1).Image = 62
                .Buttons(1).ToolTipText = "Configurar Permisos"
                .Buttons(2).Visible = True
                .Buttons(2).Image = 14
                .Buttons(2).ToolTipText = "Bloqueo de Sistema"
                .Buttons(3).Visible = True
                .Buttons(3).Style = tbrSeparator
                .Buttons(4).Visible = True
                .Buttons(4).Image = 15
                .Buttons(4).ToolTipText = "Plan de Cuentas"
                .Buttons(5).Visible = True
                .Buttons(5).Image = 8
                .Buttons(5).ToolTipText = "Centros de Costo"
                .Buttons(6).Visible = True
                .Buttons(6).Image = 16
                .Buttons(6).ToolTipText = "Auxiliares"
                .Buttons(7).Visible = True
                .Buttons(7).Image = 18
                .Buttons(7).ToolTipText = "Tipos de Cambio"
                .Buttons(8).Visible = True
                .Buttons(8).Image = 51
                .Buttons(8).ToolTipText = "Tipos de Pago"
                .Buttons(9).Visible = True
                .Buttons(9).Style = tbrSeparator
                .Buttons(10).Visible = True
                .Buttons(10).Image = 17
                .Buttons(10).ToolTipText = "Tipos de Documento"
                .Buttons(11).Visible = True
                .Buttons(11).Image = 25
                .Buttons(11).ToolTipText = "Estados de Documento"
                .Buttons(12).Visible = True
                .Buttons(12).Image = 1
                .Buttons(12).ToolTipText = "Ciclo de Vida de Documento"
                .Buttons(13).Visible = True
                .Buttons(13).Style = tbrSeparator
                .Buttons(14).Visible = True
                .Buttons(14).Image = 22
                .Buttons(14).ToolTipText = "Registro de Documentos"
                .Buttons(15).Visible = True
                .Buttons(15).Image = 23
                .Buttons(15).ToolTipText = "Búsqueda de Documentos"
                .Buttons(16).Visible = True
                .Buttons(16).Image = 80
                .Buttons(16).ToolTipText = "Trámite"
                .Buttons(17).Visible = True
                .Buttons(17).Image = 24
                .Buttons(17).ToolTipText = "Importar Ordenes de Compra"
                .Buttons(18).Visible = True
                .Buttons(18).Style = tbrSeparator
                .Buttons(19).Visible = True
                .Buttons(19).Image = 9
                .Buttons(19).ToolTipText = "Salir"
                .Buttons(20).Visible = True
                .Buttons(20).Image = 12
                .Buttons(20).ToolTipText = "Ayuda"
            Case 2
                .Buttons(1).Visible = True
                .Buttons(1).Image = 62
                .Buttons(1).ToolTipText = "Configurar Permisos"
                .Buttons(2).Visible = True
                .Buttons(2).Style = tbrSeparator
                .Buttons(3).Visible = True
                .Buttons(3).Image = 65
                .Buttons(3).ToolTipText = "Tipos de Cargo"
                .Buttons(4).Visible = True
                .Buttons(4).Image = 83
                .Buttons(4).ToolTipText = "Tipos de Contrato"
                .Buttons(5).Visible = True
                .Buttons(5).Image = 64
                .Buttons(5).ToolTipText = "Tipos de Parentesco"
                .Buttons(6).Visible = True
                .Buttons(6).Style = tbrSeparator
                .Buttons(7).Visible = True
                .Buttons(7).Image = 91
                .Buttons(7).ToolTipText = "AFP's"
                .Buttons(8).Visible = True
                .Buttons(8).Image = 81
                .Buttons(8).ToolTipText = "Seguros Médicos"
                .Buttons(9).Visible = True
                .Buttons(9).Image = 82
                .Buttons(9).ToolTipText = "Bancos"
                .Buttons(10).Visible = True
                .Buttons(10).Style = tbrSeparator
                .Buttons(11).Visible = True
                .Buttons(11).Image = 84
                .Buttons(11).ToolTipText = "Registrar Empleado"
                .Buttons(12).Visible = True
                .Buttons(12).Image = 85
                .Buttons(12).ToolTipText = "Programación y Contratos"
                .Buttons(13).Visible = True
                .Buttons(13).Image = 89
                .Buttons(13).ToolTipText = "Bonos de Campo"
                .Buttons(14).Visible = True
                .Buttons(14).Style = tbrSeparator
                .Buttons(15).Visible = True
                .Buttons(15).Image = 83
                .Buttons(15).ToolTipText = "Historial de Contratos"
                .Buttons(16).Visible = True
                .Buttons(16).Image = 90
                .Buttons(16).ToolTipText = "Configuración de Reportes"
                .Buttons(17).Visible = True
                .Buttons(17).Style = tbrSeparator
                .Buttons(18).Visible = True
                .Buttons(18).Image = 9
                .Buttons(18).ToolTipText = "Salir"
                .Buttons(19).Visible = True
                .Buttons(19).Image = 12
                .Buttons(19).ToolTipText = "Ayuda"
            Case 3
                .Buttons(1).Visible = True
                .Buttons(1).Image = 62
                .Buttons(1).ToolTipText = "Configurar Permisos"
                .Buttons(2).Visible = True
                .Buttons(2).Image = 14
                .Buttons(2).ToolTipText = "Bloqueo de Sistema"
                .Buttons(3).Visible = True
                .Buttons(3).Style = tbrSeparator
                .Buttons(4).Visible = True
                .Buttons(4).Image = 59
                .Buttons(4).ToolTipText = "Formas de Pago"
                .Buttons(5).Visible = True
                .Buttons(5).Image = 57
                .Buttons(5).ToolTipText = "Servicios"
                .Buttons(6).Visible = True
                .Buttons(6).Image = 44
                .Buttons(6).ToolTipText = "Tarifas"
                .Buttons(7).Visible = True
                .Buttons(7).Image = 63
                .Buttons(7).ToolTipText = "Registrar Servicios"
                .Buttons(8).Visible = True
                .Buttons(8).Image = 79
                .Buttons(8).ToolTipText = "Clientes"
                .Buttons(9).Visible = True
                .Buttons(9).Image = 6
                .Buttons(9).ToolTipText = "Parámetros Facturación"
                .Buttons(10).Visible = True
                .Buttons(10).Image = 78
                .Buttons(10).ToolTipText = "Almacén"
                .Buttons(11).Visible = True
                .Buttons(11).Style = tbrSeparator
                .Buttons(12).Visible = True
                .Buttons(12).Image = 36
                .Buttons(12).ToolTipText = "Realizar_Facturacion"
                .Buttons(13).Visible = True
                .Buttons(13).Style = tbrSeparator
                .Buttons(14).Visible = True
                .Buttons(14).Image = 9
                .Buttons(14).ToolTipText = "Salir"
            Case 4
                .Buttons(1).Visible = True
                .Buttons(1).Image = 62
                .Buttons(1).ToolTipText = "Configurar Permisos"
                .Buttons(2).Visible = True
                .Buttons(2).Image = 14
                .Buttons(2).ToolTipText = "Bloqueo de Sistema"
                .Buttons(3).Visible = True
                .Buttons(3).Style = tbrSeparator
                .Buttons(4).Visible = True
                .Buttons(4).Image = 15
                .Buttons(4).ToolTipText = "Plan de Cuentas"
                .Buttons(5).Visible = True
                .Buttons(5).Image = 50
                .Buttons(5).ToolTipText = "Cuentas Corrientes"
                .Buttons(6).Visible = True
                .Buttons(6).Image = 16
                .Buttons(6).ToolTipText = "Auxiliares"
                .Buttons(7).Visible = True
                .Buttons(7).Image = 69
                .Buttons(7).ToolTipText = "Tipos de Crédito"
                .Buttons(8).Visible = True
                .Buttons(8).Image = 60
                .Buttons(8).ToolTipText = "Tipos de Operaciones de Pago"
                .Buttons(9).Visible = True
                .Buttons(9).Style = tbrSeparator
                .Buttons(10).Visible = True
                .Buttons(10).Image = 88
                .Buttons(10).ToolTipText = "Estado de Cobros"
                .Buttons(11).Visible = True
                .Buttons(11).Image = 75
                .Buttons(11).ToolTipText = "Flujo de Efectivo"
                .Buttons(12).Visible = True
                .Buttons(12).Image = 51
                .Buttons(12).ToolTipText = "Liquidacion de Cobranza"
                .Buttons(13).Visible = True
                .Buttons(13).Image = 9
                .Buttons(13).ToolTipText = "Salir"
                .Buttons(14).Visible = True
                .Buttons(14).Image = 12
                .Buttons(14).ToolTipText = "Ayuda"
            Case 5
                .Buttons(1).Visible = True
                .Buttons(1).Image = 14
                .Buttons(1).ToolTipText = "Bloqueo de Sistema"
                .Buttons(2).Visible = True
                .Buttons(2).Style = tbrSeparator
                .Buttons(3).Visible = True
                .Buttons(3).Image = 15
                .Buttons(3).ToolTipText = "Plan de Cuentas"
                .Buttons(4).Visible = True
                .Buttons(4).Image = 50
                .Buttons(4).ToolTipText = "Cuentas Corrientes"
                .Buttons(5).Visible = True
                .Buttons(5).Image = 16
                .Buttons(5).ToolTipText = "Auxiliares"
                .Buttons(6).Visible = True
                .Buttons(6).Image = 59
                .Buttons(6).ToolTipText = "Formas de Pago"
                .Buttons(7).Visible = True
                .Buttons(7).Image = 54
                .Buttons(7).ToolTipText = "Responsable Firmas"
                .Buttons(8).Visible = True
                .Buttons(8).Style = tbrSeparator
                .Buttons(9).Visible = True
                .Buttons(9).Image = 47
                .Buttons(9).ToolTipText = "Orden de Cheque"
                .Buttons(10).Visible = True
                .Buttons(10).Image = 87
                .Buttons(10).ToolTipText = "Orden de Transferencia"
                .Buttons(11).Visible = True
                .Buttons(11).Image = 86
                .Buttons(11).ToolTipText = "Interfaz Telewiese"
                .Buttons(12).Visible = True
                .Buttons(12).Style = tbrSeparator
                .Buttons(19).Visible = True
                .Buttons(19).Image = 9
                .Buttons(19).ToolTipText = "Salir"
            Case 6
                .Buttons(1).Visible = True
                .Buttons(1).Image = 62
                .Buttons(1).ToolTipText = "Configurar Permisos"
        End Select
        tbPrincipal.Refresh
    End With
End Sub
Private Sub sbPrincipal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ActivaMenuEmergente = True Then
        If Button = vbRightButton Then PopupMenu m_Meses, vbPopupMenuLeftAlign
    End If
End Sub
Private Sub sbPrincipal_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 5 Then
        ActivaMenuEmergente = True
    Else
        ActivaMenuEmergente = False
    End If
End Sub
Private Sub tbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .ToolTipText
            Case "Salir"
                ModActivo = 0
                mdiInicio.Picture1.Visible = True
                InicializarMenu True
                MuestraPrimeros
                InicializaCaptionMenu
                InicializarMenu False
                CargarMenu 0
                mdiInicio.ConfigBarraHerramientas 0
                mdiInicio.sbPrincipal.Panels(3) = "Sistema Integral-Principal"
                mdiInicio.sbPrincipal.Panels(3).Picture = mdiInicio.imgsBarraHerramientas.ListImages(27).Picture
                mdiInicio.tbPrincipal.Refresh
                OcultaPrimeros
                Exit Sub
            Case "Configurar Permisos"
                frmPermisos.Show
                Exit Sub
            Case "Control Documentario"
                EjecutaMenu 1, "01", "02", "00", "00"
                Exit Sub
            Case "Facturación"
                EjecutaMenu 5, "01", "02", "00", "00"
                Exit Sub
            Case "Cuentas por Pagar"
                EjecutaMenu 9, "01", "02", "00", "00"
                Exit Sub
            Case "Cuentas por Cobrar"
                EjecutaMenu 7, "01", "02", "00", "00"
                Exit Sub
            Case "Recursos Humanos"
                EjecutaMenu 3, "01", "02", "00", "00"
                Exit Sub
            Case "Mantenimiento de Usuarios"
                EjecutaMenu 3, "02", "02", "00", "00"
                Exit Sub
            Case "Mantenimiento de Empresas"
                EjecutaMenu 1, "02", "02", "00", "00"
                Exit Sub
                '***********MODULO CONTROL DOCUMENTARIO***********
            Case "Estados de Documento"
                frmAsignarEstados.Show
                Exit Sub
            Case "Centros de Costo"
                ConfigurarFormulario FORM_MAESTROS_CENCO
                Exit Sub
            Case "Auxiliares"
                FrmAuxiliares.Show
                Exit Sub
            Case "Tipos de Cambio"
                ConfigurarFormulario FORM_MAESTRO_TIPO_CAMBIO
                Exit Sub
            Case "Tipos de Pago"
                ConfigurarFormulario FORM_MAESTRO_TIPOPAGO
                Exit Sub
            Case "Tipos de Documento"
                ConfigurarFormulario FORM_MAESTRO_TIP_DOC
                Exit Sub
            Case "Estado de Documento"
                ConfigurarFormulario FORM_MAESTRO_ESTADO
                Exit Sub
            Case "Ciclo de Vida de Documento"
                frmCicloDocumento.Show
                Exit Sub
            Case "Registro de Documentos"
                frmIngresarDocumento.Show
                Exit Sub
            Case "Búsqueda de Documentos"
                frmBusquedaDocumentaria.Show
                Exit Sub
            Case "Trámite"
                frmTramiteGerencial.Show
                Exit Sub
            Case "Importar Ordenes de Compra"
                frmInterfaz.Show
                Exit Sub
            '***********RECURSOS HUMANOS *************
            Case "Tipos de Cargo"
                ConfigurarFormulario FORM_MAESTROS_CARGOS
                Exit Sub
            Case "Tipos de Parentesco"
                ConfigurarFormulario FORM_MAESTRO_PARENT
                Exit Sub
            Case "Seguros Médicos"
                ConfigurarFormulario FORM_MAESTRO_SEGUROS
                Exit Sub
            Case "Seguros Médicos"
                ConfigurarFormulario FORM_MAESTRO_AFP
                Exit Sub
            Case "Bancos"
                ConfigurarFormulario FORM_MAESTRO_BANCOS
                Exit Sub
            Case "Tipos de Contrato"
                ConfigurarFormulario FORM_MAESTRO_CONTRATO
                Exit Sub
            Case "Registrar Empleado"
                frmRegEmpleado.Show
                Exit Sub
            Case "Programación y Contratos"
                frmSalidasEmp.Show
                Exit Sub
            Case "Bonos de Campo"
                frmRegBonos.Show
                Exit Sub
            Case "Historial de Contratos"
                frmVerContratos.Show
                Exit Sub
            Case "Configuración de Reportes"
                frmReportesRRHH.Show
                Exit Sub
            '***********MODULO FACTURACION*************
            Case "Formas de Pago"
                ConfigurarFormulario FORM_MAESTRO_FORMPAGO
                Exit Sub
            Case "Servicios"
                frmServicios.Show
                Exit Sub
            Case "Tarifas"
                ConfigurarFormulario FORM_MAESTRO_TARIFAS
                Exit Sub
            Case "Registrar Servicios"
                frmRegServicio.Show
                Exit Sub
            Case "Clientes"
                strmenuelegido = "Clientes"
                FrmAuxiliares.Show
                Exit Sub
            Case "Almacén"
                ConfigurarFormulario FORM_MAESTRO_ALMACEN
                Exit Sub
            Case "Parámetros Facturación"
                frmOpciones.Show
                Exit Sub
            Case "Realizar_Facturacion"
                frmFacturacion.Show
                Exit Sub
            '***********MODULO CUENTAS X COBRAR***********
                Case "Tipos de Crédito"
                ConfigurarFormulario FORM_MAESTRO_FORMPAGO
                Exit Sub
            Case "Estado de Cobros"
                frmDocxCobrar.Show
                Exit Sub
            Case "Flujo de Efectivo"
                frmFlujoEfectivo.Show
                Exit Sub
            Case "Liquidacion de Cobranza"
                frmLiquidacionCobranzas.Show
                Exit Sub
            Case "Tipos de Operaciones de Pago"
                ConfigurarFormulario FORM_MAESTRO_OPERACIONES
                Exit Sub
                '************MODULO CUENTAS X PAGAR **********
            Case "Cuentas Corrientes"
                ConfigurarFormulario FORM_MAESTRO_CUENTAS
                Exit Sub
            Case "Responsable Firmas"
                frmFirmas.Show
                Exit Sub
            Case "Orden de Cheque"
                frmCheques.Show
                Exit Sub
            Case "Orden de Transferencia"
                frmMovTelewiese.Show
                Exit Sub
            Case "Interfaz Telewiese"
                frmInterfazTw.Show
                Exit Sub
            Case "Salir del Sistema"
                End
        End Select
    End With
End Sub
Sub FoliosAutomaticos()
    Dim SQL As String, Ident As String
    Dim RQ As MYSQL_RS, RQ1 As MYSQL_RS
    SQL = "select * from prog_folios where '" & Format(Date, "yyyy/mm/dd") & "' >= fec_ini and '" & Format(Date, "yyyy/mm/dd") & "' <= fec_fin " & _
          "and estado = 1 and (anomes = '' or anomes <> '" & Year(Date) & Right("00" & Month(Date), 2) & "')"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    Do While Not RQ.EOF
        If Day(Date) = val(RQ.Fields("dia")) Then
            Ident = GenFolio
            SQL = "insert into amarre_documento(identificador,cod_tipo_doc,fecha_registro,hora_registro,tipo_doc_ide, " & _
                  "ide_mensajero,nombre_mensajero,empresa,obs,cod_fam,usuario,anomes,flag) " & _
                  "select '" & Ident & "',cod_tipo_doc,'" & Format(Date, "yyyy/mm/dd") & "','" & Format(Time, "hh:mm") & "', " & _
                  "tipo_doc_ide,ide_mensajero,nombre_mensajero,empresa,'F.A.',cod_fam,usuario,'" & Left(Ident, 6) & "',flag from amarre_documento where " & _
                  "identificador = '" & RQ.Fields("identificador") & "'"
            oConexionMYSQL.Execute SQL
            SQL = "insert into movi_documento(identificador,fecha_movi,cod_estado,prioridad,usuario) " & _
                  "select '" & Ident & "','" & Format(Date, "yyyy/mm/dd") & "','RG',prioridad,usuario from movi_documento where " & _
                  "identificador = '" & RQ.Fields("identificador") & "'"
            oConexionMYSQL.Execute SQL
            SQL = "insert into historial_docs(identificador,cod_estado,id_area,fecha_movi,usuario) " & _
                  "select '" & Ident & "',cod_estado,id_area,'" & Format(Date, "yyyy/mm/dd") & "',usuario from historial_docs where " & _
                  "identificador = '" & RQ.Fields("identificador") & "'"
            oConexionMYSQL.Execute SQL
            SQL = "insert into documento_contables(identificador,serie,correl,orden,guia,auxiliar,codigo,cenco,encargado, " & _
                  "mon,dias_vcto,fec_vcto,fec_emision,fec_pago,subtotal,igv,otros_montos,total,revisado,obs,fec_contabilizada, " & _
                  "voucher,solicitado,division,impequi,ref,cod_tipo_ref,total_ref,cancelado,contrato,otrosimp,tipopago,codalmac) " & _
                  "select '" & Ident & "',serie,'000000000',orden,guia,auxiliar,codigo,cenco,encargado,mon,dias_vcto, " & _
                  "DATE_FORMAT(DATE_ADD('" & Format(Date, "yyyy/mm/dd") & "',INTERVAL dias_vcto DAY),'%Y/%m/%d'),'" & Format(Date, "yyyy/mm/dd") & "', " & _
                  "'" & Format(CalcularFechaPago(Date), "yyyy/mm/dd") & "',subtotal,igv,otros_montos,total,revisado,concat('F.A.',obs),' ',' ',solicitado,division,0,' ',cod_tipo_ref, " & _
                  "0,0,contrato,0,tipopago,codalmac from documento_contables where " & _
                  "identificador = '" & RQ.Fields("identificador") & "'"
            oConexionMYSQL.Execute SQL
            SQL = "update prog_folios set anomes = '" & Year(Date) & Right("00" & Month(Date), 2) & "' where identificador = '" & RQ.Fields("identificador") & "'"
            oConexionMYSQL.Execute SQL
'            SQL = "select monto_factura,d.total,o.identificador from orden_compra o inner join documento_contables d on (o.correl=d.orden) " & _
'                  "where d.identificador = '" & RQ.Fields("identificador") & "'"
'            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
'            If Not RQ1.EOF() Then
'                If RQ1.Fields("monto_factura") < RQ1.Fields("total") Then
'                    SQL = "update orden_compra set monto_factura = monto_factura + " & RQ1.Fields("total") & " where " & _
'                          "identificador = '" & RQ1.Fields("identificador") & "'"
'                    oConexionMYSQL.Execute SQL
'                End If
'            End If
        End If
        RQ.MoveNext
    Loop
    Set RQ = Nothing
End Sub
Private Function GenFolio() As String
    Dim rsfolio As MYSQL_RS
    Dim AnoMes As String
    Dim SQL As String
    AnoMes = Year(Date) & Right("00" & Month(Date), 2)
    SQL = "max_identificador where anomes = '" & AnoMes & "'"
    Set rsfolio = oConexion.EjecutaSelect(SQL)
    If Not rsfolio.EOF Then
        GenFolio = rsfolio.Fields("anomes") & Right("0000" & Trim(str(val(rsfolio.Fields("maximo")) + 1)), 4)
    End If
    If rsfolio.RecordCount = 0 Then
        GenFolio = AnoMes & "0001"
    End If
    rsfolio.CloseRecordset
    Set rsfolio = Nothing
End Function
Private Sub RegTotTrab()
    On Error GoTo SERROR
    Dim SQL As String
    Dim RQ As MYSQL_RS
    If strAnoSistema >= Year(Date) - 1 Then
        SQL = "select distinct fecha from TotTrabPorDia where fecha = '" & Format(Date, "yyyy/mm/dd") & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If RQ.EOF() Then
            SQL = "insert into TotTrabPorDia(fecha,codigo,divlocal,divhcm,fecingreso,tot,tipo) " & _
                  "select date_format(sysdate(),'%Y/%m/%d') as fecha,(select descriplocal from cnmdepar c " & _
                  "where c.coddep=o.division) as nombres,'' as dl,'' as dh,'L' as fec,count(*),'C' as tipo " & _
                  "from empleado e left join contrato o on(e.codigo=o.codemp) Where Situacion = 1 And Tipo <> 3 " & _
                  "and o.codigo = (select max(codigo) from contrato t where t.codemp=o.codemp group by t.codemp) group by nombres " & _
                  "Union select date_format(sysdate(),'%Y/%m/%d') as fecha,(select descrip from cnmdepar c " & _
                  "where c.coddep=o.divgas) as nombres,'' as dl,'' as dh,'H' as fec,count(*),'C' as tipo from empleado e " & _
                  "left join contrato o on(e.codigo=o.codemp) Where Situacion = 1 And Tipo <> 3 and o.codigo = " & _
                  "(select max(codigo) from contrato t where t.codemp=o.codemp group by t.codemp) group by nombres Union " & _
                  "select date_format(sysdate(),'%Y/%m/%d') as fecha,e.codigo, " & _
                  "o.division,o.divgas,e.fec_ingreso as fec,0,'D' as tipo from empleado e left join contrato o on(e.codigo=o.codemp) " & _
                  "Where Situacion = 1 And Tipo <> 3 and o.codigo = (select max(codigo) from contrato t where " & _
                  "t.codemp=o.codemp group by t.codemp) order by tipo,nombres"
            oConexionMYSQL.Execute SQL
        End If
        Set RQ = Nothing
    End If
    
    Exit Sub
SERROR:
    Set RQ = Nothing
End Sub
