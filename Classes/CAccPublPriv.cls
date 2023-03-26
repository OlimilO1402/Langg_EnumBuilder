VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccPublPriv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'####################'  Enum Class from EnumBuilder  '####################'
'##################'  Name of this Class: CAccPublPriv  '##################'

'Public Enum AccPublPriv
'    AccPublic 'Public
'    AccPrivate 'Private
'End Enum

Private mBase  As EEnum
Private mValue As AccPublPriv

Private Sub Class_Initialize()
    Set mBase = New EEnum
    mBase.New_ 2
    InitEnum
End Sub
Private Sub Class_Terminate()
    Set mBase = Nothing
End Sub

Public Property Let Value(EVal As AccPublPriv)
    mValue = EVal
End Property
Public Property Get Value() As AccPublPriv
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "40"
    Value = mValue
End Property

Private Sub InitEnum()
    mBase.SetEnumItem 1, "Public", 1
    mBase.SetEnumItem 2, "Private", 2
End Sub

Public Property Get AccPublic() As AccPublPriv
    AccPublic = AccPublPriv.AccPublic
End Property
Public Property Get AccPrivate() As AccPublPriv
    AccPrivate = AccPublPriv.AccPrivate
End Property

Public Function ToString() As String
    ToString = mBase.GetNameForVal(mValue)
End Function

Public Sub ToListBox(ComboBoxOrListBox)
    mBase.ToListBox ComboBoxOrListBox
End Sub
