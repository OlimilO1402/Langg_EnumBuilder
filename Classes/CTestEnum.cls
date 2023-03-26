VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnumName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dies ist eine Klasse, die eine Box um ein Enum herum bildet
'in Enum um Enum und um Enum herum, achso das war Ulm
'Diese Klasse wird automatisch mit dem PlugIn erstellt.
'folgende Programmiersschritte werden dardurch gespart:
'anlegen von InitSubs, schreiben des Enums, und der Strings
'anlegen der EEnum-Klasse, von dem diese Klasse erbt
'
'Name des Enums in der Klasse darf nicht
'so heissen wie die Klasse selber
'Und auch die Konstanten dürfen nicht
'so heissen wie Prozeduren der Klasse
'es sei denn, man deklariert das Enum in einem Modul
Public Enum EnumName
    EnumName_FirstConst = &H1
    EnumName_SecondConst = &H2
    EnumName_ThirdConst = &H4
    EnumName_FourthConst = &H8
End Enum

Private mBase  As EEnum
Private mValue As EnumName

Private Sub Class_Initialize()
    mBase.New_ 4
    InitEnum
End Sub
Private Sub Class_Terminate()
    Set mBase = Nothing
End Sub

Public Property Let Value(EVal As EnumName)
    mValue = EVal
End Property
Public Property Get Value() As EnumName
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "40"
    Value = mValue
End Property

Private Sub InitEnum()
    mBase.SetEnumItem 1, "FirstConst", &H1
    mBase.SetEnumItem 2, "SecondConst", &H2
    mBase.SetEnumItem 3, "ThirdConst", &H4
    mBase.SetEnumItem 4, "FourthConst", &H8
End Sub

Public Function ToString() As String
    ToString = mBase.GetNameForVal(mValue)
End Function
