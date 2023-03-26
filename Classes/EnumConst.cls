VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICanUndo
Private Enum ChangeMode
    cm1Name = &H1
    cm2Value = &H2
    cm3HxVal = &H4
    cm4Descr = &H8
    cm5LFRes = &H10
End Enum
Public Name     As String      '
Public Value    As String     'As Long 'in Value soll nur der Long als String gespeichert werden
Public AsHexVal As Boolean '
Public AddDescr As String  'additional description
Public LoadFRes As Boolean 'if (LoadFromRessource) then the above description is the ResID
'Private cm As ChangeMode   '

Friend Sub New_(Optional aName As String, Optional aValue As String, Optional bAsHexVal As Boolean, Optional sAddDescr As String, Optional bLoadRes As Boolean)
    Name = aName:  Value = aValue:  AsHexVal = bAsHexVal:  AddDescr = sAddDescr: LoadFRes = bLoadRes
End Sub
'Achtung wenn der User eine Änderung macht, dann nicht einfach New_ verwenden
'sondern SetAllVals, da hier ein Undoschritt gespeichert wird
Public Sub SetAllVals(aName As String, aValue As String, bAsHexVal As Boolean, sAddDescr As String, bLoadRes As Boolean)
    Dim cm As ChangeMode: cm = Getcm(aName, aValue, bAsHexVal, sAddDescr, bLoadRes)
    OnChanging cm1Name
    New_ aName, aValue, bAsHexVal, sAddDescr, bLoadRes
    OnChanged
End Sub
Private Function Getcm(aName As String, aValue As String, bAsHexVal As Boolean, sAddDescr As String, bLoadRes As Boolean) As ChangeMode
  If (Name <> aName) Then Getcm = Getcm Or cm1Name
  If (Value <> aValue) Then Getcm = Getcm Or cm2Value
  If (AsHexVal <> bAsHexVal) Then Getcm = Getcm Or cm3HxVal
  If (AddDescr <> sAddDescr) Then Getcm = Getcm Or cm4Descr
  If (LoadFRes <> bLoadRes) Then Getcm = Getcm Or cm5LFRes
End Function
Public Function StrToVal(StrVal As String) As String
Dim v As Long
  If Not Len(StrVal) = 0 Then
    v = CLng(StrVal)
    StrToVal = CStr(v)
  End If
End Function
Public Function ValToStr() As String
Dim v As Long
  'schreibt nur eine Wert, wenn angegeben
  If Not Len(Value) = 0 Then
    v = CLng(Value)
    If AsHexVal Then ValToStr = "&H" & Hex$(v) Else ValToStr = CStr(v)
  End If
End Function
Public Sub Assign(SrcEC As EnumConst)
    With SrcEC
        New_ .Name, .Value, .AsHexVal, .AddDescr, .LoadFRes
    End With
End Sub
Public Function ToString() As String
Dim v As String:  v = ValToStr
Dim a As String: a = AddDescr
  If Len(v) > 0 Then v = " = " & v
  If Len(a) > 0 Then a = " '" & a
  ToString = Name & v & a
End Function
Public Sub InitToString(sb As StringBuilder, BN As String, mi As String, i As String)
  Call sb.Append(mi).Append("Call ").Append(BN).Append(".SetEnumItem(").Append(i).Append(", """).Append(Name).Append(""", ").Append(ValToStr).Append(")").Append(vbCrLf)
End Sub
Public Sub PropGetToString(sb As StringBuilder, EN As String, mi As String)
  Call sb.Append("Public Property Get ").Append(Name).Append("()").Append(" As ").Append(EN).Append(vbCrLf)
  Call sb.Append(mi).Append(Name).Append(" = ").Append(ValToStr).Append(vbCrLf)
  Call sb.Append("End Property").Append(vbCrLf)
End Sub



'##############################' v   OnChange   v '##############################'
Private Sub OnChanging(acm As ChangeMode)
  Call EnumBApp.UndoRedo.SaveUndo(Me, acm)
End Sub
Private Sub OnChanged()
  Call EnumBApp.CurrentDoc.UpdateView(vmEditEnum Or vmClassEnum)
End Sub
'##############################' ^   OnChange   ^ '##############################'



'##############################' Implements ICanUndo '##############################'
Public Function GetSimpleUndo(aFlag As Long) As SimpleUndo
    Dim UStr As String: Call WriteToURStr(UStr, aFlag)
    'Dim acm As ChangeMode: cm = cmAll 'aFlag braucht man hier nicht
    Set GetSimpleUndo = MNew.SimpleUndo(aFlag, UStr, Me)
End Function
Private Function ICanUndo_GetSimpleUndo(aFlag As Long) As SimpleUndo
    Set ICanUndo_GetSimpleUndo = GetSimpleUndo(aFlag)
End Function

Public Sub UndoRedo(SU As SimpleUndo)
    Dim cm As ChangeMode: cm = SU.ChgFlag
    Dim TmpStr As String: TmpStr = SU.URStr
    WriteToURStr SU.URStr, cm
    ReadFromURStr TmpStr, cm
    OnChanged
End Sub
Private Sub ICanUndo_Redo(SU As SimpleUndo)
    UndoRedo SU
End Sub
Private Sub ICanUndo_Undo(SU As SimpleUndo)
    UndoRedo SU
End Sub

Private Sub ReadFromURStr(StrVal As String, acm As ChangeMode)
Dim sep As String * 1: sep = "\"
Dim mSA() As String
Dim i As Long: i = 0
  If Len(StrVal) > 0 Then
    mSA = Split(StrVal, sep)
    If (acm And cm1Name) Then Name = mSA(i): i = i + 1
    If (acm And cm2Value) Then Value = mSA(i): i = i + 1
    If (acm And cm3HxVal) Then AsHexVal = mSA(i): i = i + 1
    If (acm And cm4Descr) Then AddDescr = mSA(i): i = i + 1
    If (acm And cm5LFRes) Then LoadFRes = mSA(i)
  End If
'  If UBound(mSA) = 4 Then '0 to 4 = 5 elems
'    Call New_(mSA(0), mSA(1), CBool(mSA(2)), mSA(3), CBool(mSA(4)))
'  End If
  'MsgBox StrVal
End Sub
Private Sub WriteToURStr(URStr As String, acm As ChangeMode)
Dim sep As String * 1: sep = "\"
  If (acm And cm1Name) Then Call AddToURStr(URStr, sep, Name)
  If (acm And cm2Value) Then Call AddToURStr(URStr, sep, ValToStr)
  If (acm And cm3HxVal) Then Call AddToURStr(URStr, sep, CbStr(AsHexVal))
  If (acm And cm4Descr) Then Call AddToURStr(URStr, sep, AddDescr)
  If (acm And cm5LFRes) Then Call AddToURStr(URStr, sep, CbStr(LoadFRes))
  'StrVal = Name & sep & ValToStr & sep & CbStr(AsHexVal) & sep & AddDescr & sep & CbStr(LoadFRes)
  'MsgBox StrVal
End Sub
Private Sub AddToURStr(URStr As String, asep As String, addStr As String)
  If Len(URStr) > 0 Then URStr = URStr & asep
  URStr = URStr & addStr
End Sub

