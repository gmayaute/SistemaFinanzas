VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msintPos As String          'Posición del Menu
Private msstrNivel As String        'Nivel del Menu
Private msstrGrupo As String        'Grupo dentro del nivel 3
Private msstrSubGrupo As String     'Grupo dentro del nivel 4
Private msintIndice As Integer      'Indice del Menu
Private oMenu As Menu

Public Property Get Pos() As String
    Pos = msintPos
End Property
Public Property Let Pos(ByVal strPos As String)
    msintPos = strPos
End Property
Public Property Get Nivel() As String
    Nivel = msstrNivel
End Property
Public Property Let Nivel(ByVal strNivel As String)
    msstrNivel = strNivel
End Property
Public Property Get Grupo() As String
    Grupo = msstrGrupo
End Property
Public Property Let Grupo(ByVal strGrupo As String)
    msstrGrupo = strGrupo
End Property
Public Property Get SubGrupo() As String
    SubGrupo = msstrSubGrupo
End Property
Public Property Let SubGrupo(ByVal strSubGrupo As String)
    msstrSubGrupo = strSubGrupo
End Property
Public Property Get Indice() As Integer
    Indice = msintIndice
End Property
Public Property Let Indice(ByVal intIndice As Integer)
    msintIndice = intIndice
End Property
Public Property Get objMenu() As Menu
    Set objMenu = oMenu
End Property
Public Property Let objMenu(ByVal mMenu As Menu)
    Set oMenu = mMenu
End Property
