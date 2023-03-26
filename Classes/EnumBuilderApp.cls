VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumBuilderApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'EnumBuilderApp als Objekt EnumBApp deklariert in ModSubMain
Implements ICanUndo
'Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private mFMain        As FMain 'm_pMainWnd
Private mFEnumConst   As FEnumConst
Private mUndoRedo     As UndoRedo
Private mDocTemplate  As SingleDocTemplate
Private Const AppName As String = "EnumBuilder"
'Private WithEvents mDocument As EnumBuilderDoc

Private Sub Class_Initialize()
    '
End Sub
Public Sub EnableVisualStyles()
    'InitCommonControls
End Sub
Public Sub InitRun()
    Set mFMain = New FMain
    Set mFEnumConst = New FEnumConst
    Set mDocTemplate = MNew.SingleDocTemplate(MNew.EnumBuilderDoc, mFMain, MNew.EnumBuilderView)
    Call mDocTemplate.Init
    'Load mFMain
    mFMain.Show
    mDocTemplate.CurrentDoc.UpdateAllViews (vmAll)
    Set mUndoRedo = New UndoRedo
    mFMain.EnDisUndoRedoBtn
End Sub
Private Sub Class_Terminate()
    Set mDocTemplate = Nothing
    Set mFMain = Nothing
    Set mFEnumConst = Nothing
    Set mUndoRedo = Nothing
    'Unload FMain
    'Unload FEnumConst
    Dim Frm As Form
    For Each Frm In VB.Forms
        Unload Frm
    Next
End Sub
Public Sub OnFileNew(bolmess As Boolean)
    Dim nUR As Long: nUR = mUndoRedo.Count
    Dim mess As String
    Dim mr As VbMsgBoxResult
    If nUR > 0 Then
        mess = "Wollen Sie ein neues Enum anlegen?" & vbCrLf & _
               "Alle " & CStr(nUR) & " Änderungen werden gelöscht"
        mr = MsgBox(mess, vbOKCancel Or vbInformation)
        If mr = vbOK Then
            'erstellt ein neues EnumBuilderDoc-Objekt
            Call mDocTemplate.NewDocument
            Set mUndoRedo = New UndoRedo
            Call mFMain.EnDisUndoRedoBtn
        End If
    Else
        mess = "Es wurden noch keine Änderungen gemacht, es wird kein neues Enum angelegt."
        MsgBox mess
    End If
End Sub
Public Sub OnFileExit()
    'Set Me = Nothing
    Call DeleteObj(EnumBApp)
End Sub

Public Sub OnAppAbout()
    FAbout.Show 0, mFMain
End Sub

Public Property Get UndoRedo() As UndoRedo
    Set UndoRedo = mUndoRedo
End Property
'Public Property Get DocTemplate() As SingleDocTemplate
'  Set DocTemplate = mDocTemplate
'End Property

Public Property Get CurrentDoc() As EnumBuilderDoc
    Set CurrentDoc = mDocTemplate.CurrentDoc
End Property
Public Property Get CurrentView() As EnumBuilderView
    Set CurrentView = mDocTemplate.CurrentView
End Property
Public Property Get MainWnd() As FMain
    Set MainWnd = mFMain
End Property
Public Property Get EnumConstantWnd() As FEnumConst
    Set EnumConstantWnd = mFEnumConst
End Property


















Public Function GetSimpleUndo(aFlag As Long) As SimpleUndo
  '
End Function
Private Function ICanUndo_GetSimpleUndo(aFlag As Long) As SimpleUndo
    Set ICanUndo_GetSimpleUndo = GetSimpleUndo(aFlag)
End Function

Public Sub Redo(SU As SimpleUndo)
  '
End Sub
Private Sub ICanUndo_Redo(SU As SimpleUndo)
    Call Redo(SU)
End Sub

Public Sub Undo(SU As SimpleUndo)
  '
End Sub
Private Sub ICanUndo_Undo(SU As SimpleUndo)
    Call Undo(SU)
End Sub



