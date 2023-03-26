VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumConsts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICanUndo
Private Enum ChangeMode
  cmAddNew
  cmInsert
  cmMove
  cmDelete
  cmBitHex
End Enum
Private mEnumConstCol As Collection
Private mMvFr As Long 'MoveFrom zwischenspeichern, für UndoRedo
Private mMvTo As Long 'MoveTo   zwischenspeichern, für UndoRedo
'Private mIndex As Long 'Index von Remove zwischenspeichern für UndoRedo

Private Sub Class_Initialize()
  Clear
End Sub
Private Sub Class_Terminate()
  Set mEnumConstCol = Nothing
End Sub
Public Property Get Count() As Long:  Count = mEnumConstCol.Count: End Property
Public Property Get Item(index As Long) As EnumConst:   Set Item = mEnumConstCol.Item(index): End Property
Public Property Get BaseCol() As Collection: Set BaseCol = mEnumConstCol: End Property
Public Sub Clear()
  Set mEnumConstCol = New Collection
End Sub

'##########################' v   Special Functions   v '###########################'
Public Sub MakeBitRow()
Dim aEC As EnumConst, i As Long
  Call OnChanging(cmBitHex)
  For i = 1 To mEnumConstCol.Count
    Set aEC = mEnumConstCol.Item(i)
    aEC.Value = GetBitHexVal(i)
    aEC.AsHexVal = True
  Next
  Call OnChanged
End Sub
Private Function GetBitHexVal(i As Long) As String
  GetBitHexVal = "&H" & Hex$(2 ^ (i - 1))
End Function
'##########################' ^   Special Functions   ^ '###########################'

'#############################' v   ColFunctions   v '#############################'
'Things to undo
Public Sub Remove(index As Long)
  mMvFr = index
  Call OnChanging(cmDelete)
  Call mEnumConstCol.Remove(index)
  Call OnChanged
End Sub
Public Sub Add(aEC As EnumConst, Optional Key, Optional Before, Optional After)
  Call OnChanging(cmAddNew)
  Call mEnumConstCol.Add(aEC, Key, Before, After)
  Call OnChanged
End Sub
Public Sub Move(FromIndex As Long, ToIndex As Long)
Dim aEC As EnumConst
  mMvFr = FromIndex
  mMvTo = ToIndex
  Call OnChanging(cmMove)
  Set aEC = mEnumConstCol.Item(FromIndex)
  'an der Stelle FromIndex entfernen, und
  Call mEnumConstCol.Remove(FromIndex)
  'wieder hinzufügen, je nachdem nach oben oder nach unten
  If FromIndex > ToIndex Then 'MoveUp
    Call mEnumConstCol.Add(aEC, , ToIndex)
  ElseIf FromIndex < ToIndex Then
    Call mEnumConstCol.Add(aEC, , , FromIndex)
  End If
  Call OnChanged
End Sub
Public Sub Insert(aConst As EnumConst, AtPos As Long)
  If AtPos <= mEnumConstCol.Count Then
    mMvFr = AtPos
    Call OnChanging(cmInsert)
    Call mEnumConstCol.Add(aConst, , AtPos)
    Call OnChanged
  Else
    Call mEnumConstCol.Add(aConst)
  End If
End Sub
'#############################' ^   ColFunctions   ^ '#############################'


'##############################' v   OnChange   v '##############################'
Private Sub OnChanging(acm As ChangeMode)
  Call EnumBApp.UndoRedo.SaveUndo(Me, acm)
End Sub
Private Sub OnChanged()
  Call EnumBApp.CurrentDoc.UpdateView(vmEditEnum Or vmClassEnum)
End Sub
'##############################' ^   OnChange   ^ '##############################'


Public Sub Assign(SrcECC As EnumConsts)
Dim i As Long, n As Long: n = SrcECC.Count
Dim aEC As EnumConst, nEC As EnumConst
  For i = 1 To n
    Set aEC = SrcECC.Item(i)
    Set nEC = New EnumConst: Call nEC.Assign(aEC)
    Call Add(nEC)
  Next
End Sub
Public Sub InitToString(sb As StringBuilder, BN As String, mi As String)
Dim aEC As EnumConst, i As Long
  For i = 1 To Count
    Set aEC = mEnumConstCol.Item(i)
    Call aEC.InitToString(sb, BN, mi, CStr(i))
  Next
End Sub
Public Sub PropGetToString(sb As StringBuilder, EN As String, mi As String)
Dim aEC As EnumConst
  For Each aEC In mEnumConstCol
    Call aEC.PropGetToString(sb, EN, mi)
  Next
End Sub

'#############################' v   ImportFuncs   v '#############################'
Public Sub Import(StrVal As String)
'extrahiert aus einem String Alle Enum-Constanten
Dim StrArr() As String, i As Long
Dim aEC As EnumConst
  StrArr = Split(StrVal, vbCrLf)
  For i = LBound(StrArr) To UBound(StrArr)
    'MsgBox StrArr(i)
    Set aEC = GetConst(StrArr(i))
    If Not aEC Is Nothing Then Call mEnumConstCol.Add(aEC)
  Next
End Sub
Private Function GetConst(StrVal As String) As EnumConst
Dim ConstNam As String, CVal As String, bHx As Boolean, SAdD As String
Dim ComPos As Long, EquPos As Long, Pos As Long
Dim rEC As EnumConst
'so also was kann da alles drinstecken?
'nur kommentar
'ein Const ohne Val ohne Komment
'ein const ohne Val mit Komment
'ein Const mit Val ohne Komment
'ein Const mit Val mit Komment
'ist der Val als Hex oder als Dec
'hat der Const den Enumnamen drin oder nicht
'wenn ja, gibt es ein SepChar zwischen Enumnamen und ConstNamen und wenn ja welches
  StrVal = Trim$(StrVal)
  If Len(StrVal) > 0 Then
    If Not (Left$(StrVal, 1) = "'") Then
      'ok kein kommentar ;) (also nicht nur)
      ComPos = InStr(1, StrVal, "'")
      If ComPos > 0 Then
        SAdD = Right$(StrVal, Len(StrVal) - ComPos)
      End If
      EquPos = InStr(1, StrVal, "=")
      If EquPos > 0 Then
        If ComPos > 0 Then
          If EquPos < ComPos Then
            Pos = Min(ComPos - 1, Len(StrVal))
          End If
        Else
          Pos = Len(StrVal)
        End If
        CVal = Trim$(Mid$(StrVal, EquPos + 1, Pos - EquPos))
        If Len(CVal) > 2 Then
          If (Left$(CVal, 2) = "&H") Then bHx = True
        End If
        'alles vor dem Wert ist jetzt der Name
        Pos = EquPos - 1
      Else
        'es gibt keinen wert, nur noch der Name vor dem Kommentar
        Pos = ComPos - 1
      End If
      ConstNam = Trim$(Left$(StrVal, Pos))
      'so jetz hamma schon ziemlich was
      Set rEC = MNew.EnumConst(ConstNam, CVal, bHx, SAdD, False)
    End If
  End If
  Set GetConst = rEC
End Function

Public Function GetPreFix() As String
Dim i As Long, j As Long, n As Long: n = mEnumConstCol.Count
Dim aEC As EnumConst, sN As String, sN2 As String
Dim spf As String, c As String, bBrk As Boolean
  If n > 0 Then
    Set aEC = mEnumConstCol.Item(1)
    sN = aEC.Name
    For j = 1 To Len(sN)
      c = Mid$(sN, j, 1)
      For i = 2 To n
        Set aEC = mEnumConstCol.Item(i)
        sN2 = aEC.Name
        If (c <> Mid$(sN2, j, 1)) Then
          bBrk = True
          Exit For
        End If
      Next
      If bBrk Then
        spf = Left$(sN, j - 1)
        Exit For
      End If
    Next
  End If
  If Len(spf) > 0 Then
    GetPreFix = spf
  Else
    GetPreFix = vbNullString
  End If
End Function

Public Sub RemoveFromName(StrVal As String, Optional Idx As Long = 0)
Dim aEC As EnumConst
Dim sN As String, d As Long
  If Idx <= 0 Then Idx = 1
  For Each aEC In mEnumConstCol
    sN = aEC.Name
    d = Len(sN) - Len(StrVal)
    If d < 0 Then Exit Sub
    If Mid$(sN, Idx, Len(StrVal)) = StrVal Then
      sN = Right$(sN, d)
    End If
    aEC.Name = sN
  Next
End Sub
'#############################' ^   ImportFuncs   ^ '#############################'



'##############################' Implements ICanUndo '##############################'
Public Function GetSimpleUndo(aFlag As Long) As SimpleUndo
Dim URStr As String, acm As ChangeMode: acm = aFlag
  Select Case aFlag
  Case cmAddNew
  Case cmBitHex
  Case cmDelete
  Case cmInsert
  Case cmMove
  End Select
  Set GetSimpleUndo = MNew.SimpleUndo(aFlag, URStr, Me)
End Function
Private Function ICanUndo_GetSimpleUndo(aFlag As Long) As SimpleUndo
  Set ICanUndo_GetSimpleUndo = GetSimpleUndo(aFlag)
End Function
Private Sub UndoRedo(SU As SimpleUndo)
  'bei Redo von hinzufügen eines Elementes dieses Element im SU.URStr speichern
  
End Sub
Private Sub ICanUndo_Redo(SU As SimpleUndo)
  Call UndoRedo(SU)
End Sub
Private Sub ICanUndo_Undo(SU As SimpleUndo)
  Call UndoRedo(SU)
End Sub
