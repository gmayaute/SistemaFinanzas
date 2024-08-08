VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Fields"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarLenght As Long
Private mvarValue As String
Private mvarFields As Fields

Public Property Get Fields() As Fields
    If mvarFields Is Nothing Then
        Set mvarFields = New Fields
    End If
    Set Fields = mvarFields
End Property
Public Property Set Fields(vData As Fields)
    Set mvarFields = vData
End Property
Private Sub Class_Terminate()
    Set mvarFields = Nothing
End Sub
Public Property Let Value(ByVal vData As String)
    mvarValue = vData
End Property
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = mvarValue
End Property
Public Property Let Lenght(ByVal vData As Long)
    mvarLenght = vData
End Property
Public Property Get Lenght() As Long
   Lenght = mvarLenght
End Property
