VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumBuilderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This is not more views of the same view, but more of (some) different views
'of different parts, resp. aspects of the document
Private mViewMode As ViewMode
Private WithEvents mDocument As EnumBuilderDoc
Attribute mDocument.VB_VarHelpID = -1
'mForm und mFrm ist natürlich zweimal dasselbe Objekt, aber
'nur über die Schnittstelle Form bekommt man auch die Events, und
'nur über die Schnittstelle FMain bekommt man Zugriff auf die
'einzelnen Controls auf der Form
Private WithEvents mFrm As Form
Attribute mFrm.VB_VarHelpID = -1
Private mForm           As FMain
'die einzelenn Views:
Private mPbViewEnumOption As PictureBox
Private WithEvents mPbViewEnumEdit As PictureBox
Attribute mPbViewEnumEdit.VB_VarHelpID = -1
Private mTBViewEnumClass  As TextBox

Private Sub Class_Initialize()
  '
End Sub
Private Sub Class_Terminate()
    Set mDocument = Nothing
    Set mFrm = Nothing
    Set mForm = Nothing
End Sub

Friend Sub New_(aEBDoc As EnumBuilderDoc)
    Set mDocument = aEBDoc
End Sub
Public Sub Init(aFrm As FMain)
    Set mForm = aFrm
    Set mFrm = aFrm
    Set mPbViewEnumOption = mForm.PbViewEnumOption
    Set mPbViewEnumEdit = mForm.PbViewEnumEdit
    Set mTBViewEnumClass = mForm.TbViewEnumClass
    InitEnumIndent
    InitConstIndent
    mViewMode = vmEditEnum
End Sub
Private Sub InitEnumIndent()
    Call InitCmb(mForm.CmbIndent) ', mDocument.Indent)
End Sub
Private Sub InitConstIndent()
    Call InitCmb(mForm.CmbConstIndent) ', mDocument.ConstIndent)
End Sub
Private Sub InitCmb(CB As ComboBox) ', ind As Long)
    Dim i As Long
    For i = 0 To 9
        Call CB.AddItem(CStr(i))
    Next
    'CB.ListIndex = ind
End Sub

Public Property Let ViewMode(vmVal As ViewMode)
    mViewMode = vmVal
    Select Case mViewMode
    Case vmAll
        UpdateEnumOptionsView
        UpdateEnumEditView
        UpdateEnumClassView
    Case vmClassEnum
        mTBViewEnumClass.ZOrder 0
        UpdateEnumClassView
    Case vmEditEnum
        mPbViewEnumEdit.ZOrder 0
        UpdateEnumEditView
    Case vmOptionEnum
        mPbViewEnumOption.ZOrder 0
        UpdateEnumOptionsView
    Case Else
    End Select
End Property
Public Property Get ViewMode() As ViewMode: ViewMode = mViewMode: End Property

'Private Sub mpDocument_Changed()
'  Select Case mViewMode
'  Case vmEnumEdit
'    Call EnumBApp.MainWnd.UpdateEditWnd(mEBDoc)
'  Case vmEnumClass
'    Call EnumBApp.MainWnd.UpdateClassWnd(mEBDoc)
'  End Select
'End Sub
'a more likely version of UpdateAllViews is UpdateView
Private Sub mDocument_UpdateView(vmView As ViewMode)
  If Not mForm Is Nothing Then
    If vmView And vmClassEnum Then Call UpdateEnumClassView
    If vmView And vmEditEnum Then Call UpdateEnumEditView
    If vmView And vmOptionEnum Then Call UpdateEnumOptionsView
  End If
End Sub

Private Sub mDocument_UpdateAllViews(vmExcludedView As ViewMode)
'maybe a view itself calls this function, and wants to exclude itself
  If Not mForm Is Nothing Then
    If (Not (vmExcludedView = vmOptionEnum)) Then Call UpdateEnumOptionsView
    If (Not (vmExcludedView = vmEditEnum)) Then Call UpdateEnumEditView
    If (Not (vmExcludedView = vmClassEnum)) And (mViewMode = vmClassEnum) Then Call UpdateEnumClassView
  End If
End Sub

Private Sub UpdateEnumOptionsView()
  With mForm
    .CmbIndent.Text = CStr(mDocument.Indent)
    .CmbConstIndent.Text = CStr(mDocument.ConstIndent)

    .TxtPreFix.Text = mDocument.PreFix 'zuerst die TextBox setzen!!
    .ChkInclPreFix.Value = Bol2CheckBoxValue(mDocument.IncludePreFix)
    .TxtPreFix.Enabled = mDocument.IncludePreFix
    .ChkConstPropGet.Value = Bol2CheckBoxValue(mDocument.PropGetConst)
    .OptRetClass.Value = (mDocument.ReturnClass = True)
    .OptRetEnum.Value = (mDocument.ReturnClass = False)
  End With
End Sub
Private Sub UpdateEnumEditView()
  With mForm
    .LblEnumIndent.Caption = mDocument.IndentToString '(1)
    If Len(.LblEnumIndent.Caption) = 0 Then .LblEnumIndent.Width = 0
    .LblPublPriv.Caption = mDocument.AccToStr
  
    .LblPublPriv.Move .LblEnumIndent.Left + .LblEnumIndent.Width, .LblEnumIndent.Top
    .LstPublPriv.Move .LblPublPriv.Left, .LblPublPriv.Top + .LblPublPriv.Height
    .LblEnum.Move .LblPublPriv.Left + .LblPublPriv.Width, .LblPublPriv.Top

    .LblEnumName.Caption = mDocument.EnumName
    .LblEnumName.Move .LblEnum.Left + .LblEnum.Width, .LblEnum.Top
    .TxtEnumName.Text = mDocument.EnumName
    .TxtEnumName.Move .LblEnumName.Left, .LblEnumName.Top, mPbViewEnumEdit.ScaleWidth - .LblEnumName.Left

    .LblEnumConst.Caption = RepAByAA(mDocument.ConstsToVBLabel)
    .LblEnumConst.Move .LblEnumIndent.Left + .LblEnumIndent.Width, .LblPublPriv.Top + .LblPublPriv.Height

    .ChkConstPropGet.Value = Bol2CheckBoxValue(mDocument.PropGetConst)
    Call mDocument.ConstsToVBListBox(.LstEnumConst)
    .LstEnumConst.Move .LblEnumIndent.Left + .LblEnumIndent.Width, .LblPublPriv.Top + .LblPublPriv.Height, mPbViewEnumEdit.ScaleWidth - .LblPublPriv.Left, .LblEnumConst.Height + 90
    .LblEndEnum.Move .LblEnumIndent.Left + .LblEnumIndent.Width, .LblEnumConst.Top + .LblEnumConst.Height
    .EnDisUndoRedoBtn
  End With
End Sub
Private Sub UpdateEnumClassView()
  mTBViewEnumClass.Text = mDocument.EnumClassToString
End Sub

Private Sub mFrm_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    Dim Brdr As Single ': Brdr = 120
    L = 0 'mPbViewEnumOption.Left
    T = mPbViewEnumOption.Top + mPbViewEnumOption.Height + 1 * Brdr
    W = Max(mForm.ScaleWidth - 2 * Brdr, 0)
    H = Max(mForm.ScaleHeight - T - 1 * Brdr, 0)
    If W > 0 And H > 0 Then mTBViewEnumClass.Move L, T, W, H
    If W > 0 And H > 0 Then mPbViewEnumEdit.Move L, T, W, H
End Sub

Private Sub mPbViewEnumEdit_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    With mForm
        L = .LblEnumIndent.Left + .LblEnumIndent.Width
        T = .LblPublPriv.Top + .LblPublPriv.Height
        W = Max(mPbViewEnumEdit.ScaleWidth - .LblPublPriv.Left, 15)
        H = Max(.LblEnumConst.Height + 90, 15)
        If W > 0 And H > 0 Then .LstEnumConst.Move L, T, W, H
    End With
End Sub
