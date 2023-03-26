VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UndoRedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ECBUndoRedo 'ein allgemeines UndoRedo
Private mURCol As Collection
Private mIxLastItem As Long
Private mBolUndoing As Boolean

Private Sub Class_Initialize()
  Call Clear
End Sub
Private Sub Class_Terminate()
  Set mURCol = Nothing
End Sub

Public Sub Clear()
  Set mURCol = New Collection
End Sub
Public Property Get Count() As Long
  Count = mURCol.Count
End Property
Public Sub SaveUndo(ICU As ICanUndo, aFlag As Long)
'  If aFlag = 0 Then
'    Call Add(New_SimpleUndo(aFlag, vbNullString, ICU))
'  Else
    Call Add(ICU.GetSimpleUndo(aFlag))
'  End If
End Sub
Private Sub Add(SU As SimpleUndo)
  If mIxLastItem = mURCol.Count Then
    Call mURCol.Add(SU)
    mIxLastItem = mURCol.Count
  Else
    'oder hier zuerst alle vorausgegangenen Redos löschen
    'und dann ans Ende anfügen
    If mIxLastItem < 1 Then
      Call mURCol.Add(SU)
    Else
      Call mURCol.Add(SU, , , mIxLastItem)
    End If
    mIxLastItem = mIxLastItem + 1
  End If
End Sub
Public Sub Undo()
Dim SU As SimpleUndo
Dim fSU As SimpleUndo
Dim aICU As ICanUndo
  'kann der erste sein, und darf auch der letzte sein
  If (0 < mIxLastItem) And (mIxLastItem <= mURCol.Count) Then
    Set SU = mURCol.Item(mIxLastItem)
    Set aICU = SU.ICU
    mIxLastItem = mIxLastItem - 1
    mBolUndoing = True
    Set fSU = aICU.GetSimpleUndo(SU.ChgFlag)
    Call aICU.Undo(SU)
    Set SU = fSU
    mBolUndoing = False
  End If
End Sub
'Private Sub FlipStr(StrVal1 As String, StrVal2 As String)
'Dim TmpStr As String:  TmpStr = StrVal1:  StrVal1 = StrVal2:  StrVal2 = TmpStr
'End Sub
Public Sub Redo()
Dim SU As SimpleUndo
Dim fSU As SimpleUndo
Dim aICU As ICanUndo
  'kann der erste sein, aber kann nicht der letzte sein
  If (0 <= mIxLastItem) And (mIxLastItem < mURCol.Count) Then
    mIxLastItem = mIxLastItem + 1
    Set SU = mURCol.Item(mIxLastItem)
    Set aICU = SU.ICU
    mBolUndoing = True
    Set fSU = aICU.GetSimpleUndo(SU.ChgFlag)
    Call aICU.Redo(SU)
    Set SU = fSU
    mBolUndoing = False
  End If

'Dim SU As SimpleUndo, TmpStr As String
'Dim TECB As EnumBuilderDoc 'TemporaryVal
'  If (0 < mIxLastItem) And (mIxLastEdit < mURCol.count) Then 'darf nicht der letzte sein
'    mBolUndoing = True
'    mIxLastEdit = mIxLastEdit + 1
'    Set SU = mURCol.Item(mIxLastEdit)
'    TmpStr = SU.URStr
'    Select Case SU.um
'    Case umAccess: SU.URStr = CStr(aECB.Access):        aECB.Access = CLng(TmpStr)
'    Case umEnumNm: SU.URStr = aECB.EnumName:            aECB.EnumName = TmpStr
'    Case umIndent: SU.URStr = CStr(aECB.Indent):        aECB.Indent = CLng(TmpStr)
'    Case umCnsInd: SU.URStr = CStr(aECB.ConstIndent):   aECB.ConstIndent = CLng(TmpStr)
'    Case umIcPrFx: SU.URStr = CStr(aECB.IncludePreFix): aECB.Indent = CLng(TmpStr)
'    Case umPreFix: SU.URStr = aECB.PreFix:              aECB.PreFix = CLng(TmpStr)
'    Case umPGCnst: SU.URStr = CStr(aECB.PropGetConst):  aECB.PropGetConst = CLng(TmpStr)
'    Case umRetCls: SU.URStr = CStr(aECB.PreFix):        aECB.PreFix = CBool(TmpStr)
'    Case umAll: Set TECB = aECB: aECB = SU.ECB: SU.ECB = TECB
'    End Select
'    mBolUndoing = False
'  End If
End Sub

Public Property Get Undoing() As Boolean
  Undoing = mBolUndoing
End Property
Public Property Get UndoEnabled() As Boolean
  If (0 < mIxLastItem) And (mIxLastItem <= mURCol.Count) Then UndoEnabled = True Else UndoEnabled = False
End Property
Public Property Get RedoEnabled() As Boolean
  If (-1 < mIxLastItem) And (mIxLastItem < mURCol.Count) Then RedoEnabled = True Else RedoEnabled = False
End Property

Public Function GetURStr() As String
  GetURStr = CStr(mIxLastItem) & " / " & CStr(mURCol.Count - mIxLastItem)
End Function
