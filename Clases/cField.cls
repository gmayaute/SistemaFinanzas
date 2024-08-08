VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarValue As String
Private mvarLenght As Long
Public Property Let Lenght(ByVal vData As Long)
    mvarLenght = vData
End Property
Public Property Get Lenght() As Long
    Lenght = mvarLenght
End Property

Public Property Let Value(ByVal vData As String)
    mvarValue = vData
End Property
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
   Value = mvarValue
End Property



