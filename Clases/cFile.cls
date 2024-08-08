VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Lines"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private caracteres As Long
Private mvarFileName As String
Private mvarSize As Long
Private mvarPath As String
Private mvarFileType As String
Private mvarFileText As String
Private mvarFieldSeparator As String
Private mvarPercentParsing As Byte
Private mvarPercentParsing2 As Byte
Private mvarDateCreation As Date
Private mvarDateModification As Date
Private mvarBOF As Boolean
Private mvarEOF As Boolean
Private mvarLines As Lines
Private FSO As FileSystemObject
Private File As File
Private Folder As Folder
Private txtStream As TextStream
Public Event Error(lErrNumber As Long, sErrDescription As String)
Public Event ParseComplete()
Public Event PercentParsing(bPercent As Byte)
Public Event ParseComplete2()
Public Event PercentParsing2(bPercent As Byte)
Public cuenta As Long

Public Property Get Lines() As Lines
    If mvarLines Is Nothing Then
        Set mvarLines = New Lines
    End If
    Set Lines = mvarLines
End Property
Private Sub Class_Initialize()
    Set FSO = New FileSystemObject
    mvarFieldSeparator = " "
    Set mvarLines = New Lines
End Sub
Private Sub Class_Terminate()
    Set mvarLines = Nothing
End Sub
Public Property Let EOF(ByVal vData As Boolean)
    mvarEOF = vData
End Property
Public Property Get EOF() As Boolean
    EOF = mvarEOF
End Property
Public Property Let BOF(ByVal vData As Boolean)
    mvarBOF = vData
End Property
Public Property Get BOF() As Boolean
    BOF = mvarBOF
End Property
Public Property Let DateModification(ByVal vData As Date)
    mvarDateModification = vData
End Property
Public Property Get DateModification() As Date
    DateModification = mvarDateModification
End Property
Public Property Let DateCreation(ByVal vData As Date)
    mvarDateCreation = vData
End Property
Public Property Get DateCreation() As Date
    DateCreation = mvarDateCreation
End Property
Public Property Let PercentParsing(ByVal vData As Byte)
    mvarPercentParsing = vData
End Property
Public Property Get PercentParsing() As Byte
    PercentParsing = mvarPercentParsing
End Property
Public Property Let PercentParsing2(ByVal vData As Byte)
    mvarPercentParsing2 = vData
End Property
Public Property Get PercentParsing2() As Byte
    PercentParsing2 = mvarPercentParsing2
End Property
Public Property Let FieldSeparator(ByVal vData As String)
    mvarFieldSeparator = vData
End Property
Public Property Get FieldSeparator() As String
    FieldSeparator = mvarFieldSeparator
End Property
Public Property Let FileText(ByVal vData As String)
    mvarFileText = vData
End Property
Public Property Get FileText() As String
    FileText = mvarFileText
End Property
Public Property Let FileType(ByVal vData As String)
    mvarFileType = vData
End Property
Public Property Get FileType() As String
    FileType = mvarFileType
End Property
Public Property Let Size(ByVal vData As Long)
    mvarSize = vData
End Property
Public Property Get Size() As Long
    Size = mvarSize
End Property
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property
Public Property Get FileName() As String
Attribute FileName.VB_UserMemId = 0
    FileName = mvarFileName
End Property
Public Function Parse()
    On Error GoTo err
    caracteres = 0
    If FSO.FileExists(mvarFileName) Then
        Set txtStream = FSO.OpenTextFile(mvarFileName, ForReading, False, TristateFalse)
        mvarBOF = True
        LoadFile
        Set txtStream = FSO.OpenTextFile(mvarFileName, ForReading, False, TristateFalse)
        mvarFileText = txtStream.ReadAll
        LoadSpecs
    Else
        RaiseEvent Error(60001, "File not found.")
    End If
    Exit Function
err:
    RaiseEvent Error(err.Number, err.Description)
End Function
Private Function LoadFile()
    Dim strLine As String, strFile As String, strTemp As String
    Dim lCont As Long
    mvarBOF = False
    Set mvarLines = New Lines
    strFile = txtStream.ReadAll
    Set txtStream = FSO.OpenTextFile(mvarFileName, ForReading, False, TristateFalse)
    Do Until txtStream.AtEndOfStream = True
        strLine = txtStream.ReadLine
        mvarLines.Add Len(strLine), strLine, LoadFields(strLine, Len(strFile))
        strTemp = strTemp & strLine
    Loop
    RaiseEvent PercentParsing(100)
    RaiseEvent ParseComplete
    mvarBOF = False
    mvarEOF = True
End Function
Private Function LoadFields(ByVal strLine As String, lenfile As Long) As Fields
    Dim oFields As Fields
    Dim strTemp As String
    Dim iCont As Long
    Set oFields = New Fields
    Dim Linea As String
    Dim cuentaCampos As Integer
    Dim i As Integer
    cuentaCampos = 0
    For iCont = 1 To Len(strLine)
        Select Case Mid(strLine, iCont, 1)
            Case Is <> mvarFieldSeparator, Chr(13)
                caracteres = caracteres + 1
                strTemp = strTemp & Replace(Mid(strLine, iCont, 1), Chr(34), Empty)
                RaiseEvent PercentParsing(Round(caracteres) * 100 / lenfile)
            Case Else
                oFields.Add strTemp, Len(strTemp)
                cuentaCampos = cuentaCampos + 1
                DoEvents
                strTemp = ""
        End Select
    Next
    If cuentaCampos < 22 Then
        For i = cuentaCampos To 21
            oFields.Add " ", 1
        Next i
    End If
    cuentaCampos = 0
    If Trim(strTemp) <> "" Then
        oFields.Add strTemp, Len(strTemp)
    End If
    strTemp = ""
    Set LoadFields = oFields
    Set oFields = Nothing
End Function
Private Function LoadSpecs()
    Set File = FSO.GetFile(mvarFileName)
    With File
        mvarDateCreation = .DateCreated
        mvarDateModification = .DateLastModified
        mvarFileType = .Type
        mvarSize = .Size
    End With
End Function
Public Sub CloseFile()
    Set txtStream = Nothing
End Sub
