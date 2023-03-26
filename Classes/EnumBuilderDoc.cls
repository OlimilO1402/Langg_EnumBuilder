VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumBuilderDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse schreibt die Klasse <EnumName>
Implements ICanUndo
Public Enum ChangeMode
    'cmAll = 0 'As EnumConsts 'das ganze Objekt soll in Undo gespeichert werden
    cmAccess  'as AccPublPriv 'Long
    cmEnumNm  'As String
    cmIndent  'As Long
    cmCnsInd  'As Long
    cmIcPrFx  'As Boolean
    cmPreFix  'As String
    cmPGCnst  'As Boolean
    cmRetCls  'As Boolean
    'cmSetCns  'eine ganze Konstante rückgängig machen
End Enum

Private mAccess       As AccPublPriv
Private mEnumName     As String
Private mIndent       As Long    'Variable für Texteinrückung, nach Vorliebe des Users
Private mConstIndent  As Long
Private mInclPreFix   As Boolean
Private mPreFix       As String  'das Zeichen das im Namen der Konstanten den EnumNamen und den Konstantennamen voneinander abtrennt
Private mPropGetConst As Boolean
Private mRetClass     As Boolean 'Class = True Enum = False

Private mEnumConsts   As EnumConsts
Private mSelected     As Long
'Private mEBView       As EnumBuilderView 'nur 1 ein View! also kein Liste von Views
'Public Event Changing(cm As ChangeMode)  'noch vor der Veränderung
'Public Event Changed()
Public Event UpdateView(vmView As ViewMode)
Public Event UpdateAllViews(vmExcludedView As ViewMode)

Private Sub Class_Initialize()
    Set mEnumConsts = New EnumConsts
End Sub
Private Sub Class_Terminate()
    Set mEnumConsts = Nothing
End Sub

Friend Sub New_(Optional aAcc As AccPublPriv = DefAccess, Optional aEnumName As String = DefEnumName, Optional aIndent As Long = DefIndent, Optional aConstIndent As Long = DefCnstInd, Optional bIncPrFx As Boolean = DefIncPrFx, Optional aPreFix As String = DefPreFix, Optional bPropGet As Boolean = DefPropGet, Optional bRetClass As Boolean = DefRetCls)
    mAccess = aAcc: mEnumName = aEnumName: mIndent = aIndent: mConstIndent = aConstIndent: mInclPreFix = bIncPrFx: mPreFix = aPreFix: mPropGetConst = bPropGet: mRetClass = bRetClass
End Sub
'Public Sub AddView(aView As EnumBuilderView)
'  Set mEBView = aView
'End Sub
Friend Property Get Consts() As EnumConsts
    Set Consts = mEnumConsts
End Property

Public Sub Clear()
    mAccess = DefAccess 'AccPublic
    mEnumName = DefEnumName 'vbNullString
    'mIndent 'so lassen
    'mConstIndent 'so lassen
    mInclPreFix = DefIncPrFx 'False
    mPreFix = DefPreFix 'vbNullString
    mEnumConsts.Clear
    mSelected = 0
End Sub


'##############################' v   Properties   v '##############################'
Public Property Let Access(aAcc As AccPublPriv)
    If Not (mAccess = aAcc) Then Call OnChanging(cmAccess): mAccess = aAcc: OnChanged
End Property
Public Property Get Access() As AccPublPriv: Access = mAccess: End Property
Public Function AccToStr() As String
    If mAccess = AccPrivate Then AccToStr = "Private" Else AccToStr = "Public"
End Function

Public Property Let EnumName(StrVal As String)
    If Not (mEnumName = StrVal) Then Call OnChanging(cmEnumNm): mEnumName = StrVal: OnChanged
End Property
Public Property Get EnumName() As String: EnumName = mEnumName:   End Property

Public Property Let Indent(LngVal As Long)
    If Not (mIndent = LngVal) Then Call OnChanging(cmIndent): mIndent = LngVal: OnChanged
End Property
Public Property Get Indent() As Long: Indent = mIndent: End Property

Public Property Let ConstIndent(LngVal As Long)
    If Not (mConstIndent = LngVal) Then Call OnChanging(cmCnsInd): mConstIndent = LngVal: OnChanged
End Property
Public Property Get ConstIndent() As Long: ConstIndent = mConstIndent: End Property

Public Property Let IncludePreFix(BolVal As Boolean)
    If Not (mInclPreFix = BolVal) Then Call OnChanging(cmIcPrFx): mInclPreFix = BolVal: OnChanged
End Property
Public Property Get IncludePreFix() As Boolean:  IncludePreFix = mInclPreFix: End Property

Public Property Let PreFix(StrVal As String)
    If Not (mPreFix = StrVal) Then Call OnChanging(cmPreFix): mPreFix = StrVal: OnChanged
End Property
Public Property Get PreFix() As String:   PreFix = mPreFix: End Property

Public Property Let PropGetConst(BolVal As Boolean)
        If Not (mPropGetConst = BolVal) Then Call OnChanging(cmPGCnst): mPropGetConst = BolVal: OnChanged
End Property
Public Property Get PropGetConst() As Boolean:   PropGetConst = mPropGetConst: End Property

Public Property Let ReturnClass(BolVal As Boolean)
    If Not (mRetClass = BolVal) Then Call OnChanging(cmRetCls): mRetClass = BolVal: OnChanged
End Property
Public Property Get ReturnClass() As Boolean:   ReturnClass = mRetClass: End Property

Public Property Let SelectedIndex(index As Long):  mSelected = index:  End Property
Public Property Get SelectedIndex() As Long:  SelectedIndex = mSelected: End Property
'##############################' ^   Properties   ^ '##############################'




'##############################' v   EnumConsts   v '##############################'
Public Property Get SelectedItem() As EnumConst
    Set SelectedItem = Item(mSelected)
End Property
Public Sub SetSelectedItem(Optional aName As String, Optional aValue As String, Optional bAsHexVal As Boolean, Optional sAddDescription As String, Optional bLoadRes As Boolean)
    Call Item(mSelected).SetAllVals(aName, aValue, bAsHexVal, sAddDescription, bLoadRes)
End Sub
Public Property Get Item(index As Long) As EnumConst
    Set Item = mEnumConsts.Item(index)
End Property
Public Property Get Count() As Long
    Count = mEnumConsts.Count
End Property
Public Sub AddConst(aNewConst As EnumConst)
    Call mEnumConsts.Add(aNewConst)
    mSelected = mEnumConsts.Count
End Sub
Public Sub DeleteSelectedConst()
    If mSelected <= mEnumConsts.Count Then
        'Call OnChanging(cmAll)
        mEnumConsts.Remove (mSelected)
        If mSelected > 1 Then mSelected = mSelected - 1
        'Call OnChanged
    End If
End Sub
Public Sub MoveSelectedUp()
    Call MoveItem(True)
End Sub
Public Sub MoveSelectedDown()
    Call MoveItem(False)
End Sub
Private Sub MoveItem(Up As Boolean)
    Dim ToIx As Long, L As Long, u As Long
    If Up Then
        L = 1: u = mEnumConsts.Count:     ToIx = mSelected - 1
    Else
        L = 0: u = mEnumConsts.Count - 1: ToIx = mSelected + 1
    End If
    If (L < mSelected) And (mSelected < u) Then
        'Call OnChanging(cmAll)
        Call mEnumConsts.Move(mSelected, ToIx) 'Reihenfolge beachten!
        mSelected = ToIx
        'Call OnChanged
    End If
End Sub
Public Sub InsertConst(aNewConst As EnumConst)
    Dim ix As Long
Try: On Error GoTo Catch
    If mSelected = 0 Then ix = 1 Else ix = mSelected
    Call mEnumConsts.Insert(aNewConst, ix)
    Exit Sub
Catch:
    MsgBox "Fehler: " & CStr(Err.Number) & " in EnumBuilderdoc::InsertConst " & Err.Description
End Sub
'##############################' ^   EnumConsts   ^ '##############################'



'##############################' v    OnChange    v '##############################'
Private Sub OnChanging(cm As ChangeMode)
    'RaiseEvent Changing(cm)
    EnumBApp.UndoRedo.SaveUndo Me, cm
End Sub
Private Sub OnChanged()
    'RaiseEvent Changed
    UpdateAllViews vmAll
End Sub
'##############################' ^    OnChange    ^ '##############################'

Public Sub UpdateView(vmView As ViewMode)
    RaiseEvent UpdateView(vmView)
End Sub
Public Sub UpdateAllViews(vmExcludedView As ViewMode)
    RaiseEvent UpdateAllViews(vmExcludedView)
End Sub

Public Sub Assign(SrcECB As EnumBuilderDoc)
    'kopiert Alles vom source SrcECB in diese
    With SrcECB
        New_ .Access, .EnumName, .Indent, .ConstIndent, .IncludePreFix, .PreFix
        mEnumConsts.Assign .Consts
    End With
End Sub

Public Function IndentToString(Optional nLoops As Long = -1) As String
    Dim i As Long, n As Long
    If nLoops = -1 Then
        n = mEnumConsts.Count
    Else
        n = nLoops
    End If
    For i = 1 To n
        IndentToString = IndentToString & Space$(mIndent)
        If i < n Then IndentToString = IndentToString & vbCrLf
    Next
End Function
Public Function ConstsToVBLabel() As String
    Dim StrRet As String
    GetConstStr vbNullString, True, Nothing, StrRet
    ConstsToVBLabel = StrRet
End Function
Public Sub ConstsToVBListBox(LB As ListBox)
    LB.Clear
    GetConstStr vbNullString, False, LB, vbNullString
End Sub

Public Function EnumToString() As String
    Dim StrRet As String
    Dim sp1 As String: sp1 = IndentToString(1)
    StrRet = StrRet & sp1 & AccToStr & " Enum " & mEnumName & vbCrLf
    Call GetConstStr(sp1, False, Nothing, StrRet)
    StrRet = StrRet & vbCrLf & sp1 & "End Enum"
    EnumToString = StrRet
End Function

'##############################' Special functions '##############################'
Public Sub MakeBitRow()
    mEnumConsts.MakeBitRow
End Sub

Private Sub GetConstStr(sp1 As String, IsLabel As Boolean, LB As ListBox, StrRet As String)
    Dim i As Long, n As Long: n = mEnumConsts.Count
    Dim eC As EnumConst
    Dim sp2 As String: sp2 = sp1 & Space$(mConstIndent)
    Dim cnst As String
    For i = 1 To n
        Set eC = mEnumConsts.Item(i)
        cnst = eC.ToString
        If mInclPreFix Then
            cnst = mPreFix & cnst
        End If
        If LB Is Nothing Then
            'keine Ahnung warum, aber im Label muß ab dem zweiten ein Space mehr rein
            If IsLabel Then If i > 1 Then sp2 = sp1 & Space$(mConstIndent + 1)
            StrRet = StrRet & sp2 & cnst
            If i < n Then StrRet = StrRet & vbCrLf
        Else
            Call LB.AddItem(sp2 & cnst)
        End If
    Next
End Sub
Public Sub Import(StrVal As String)
    Dim nPos As Long, eePos As Long '= Pos of "End Enum"
    Dim i As Long
    Call ParseAccBeginEndPos(StrVal, nPos, eePos)
    If nPos > 0 Then
        If eePos > 0 Then
            If nPos < eePos Then
                'Alle vorhandenen Konstanten und alles andere löschen
                Clear
                'Ok scheinbar vollständig enthalten
                'den Namen rauslesen
                mEnumName = GetTheName(StrVal, nPos)
                'so jetzt vom Anfang bis nPos alles wegschneiden
                'und von eepos bis Ende alles wegschneiden
                StrVal = Mid$(StrVal, nPos, eePos - nPos)
                
                mEnumConsts.Import StrVal
                mPreFix = mEnumConsts.GetPreFix
                If Len(mPreFix) > 0 Then
                    mInclPreFix = True
                    Call mEnumConsts.RemoveFromName(mPreFix, 1)
                End If
            Else
                MsgBox "Achtung! 'End Enum' kommt vor 'Public Enum': " & StrVal
            End If
        Else
            MsgBox "Enum im String nicht vollständig: " & StrVal
        End If
    Else
        MsgBox "Kein Enum im String gefunden: " & StrVal
    End If
    Call OnChanged
End Sub
Private Function ParseAccBeginEndPos(StrVal As String, BeginPos As Long, EndPos As Long) As Boolean
    'liefert Private/Public, und Beginn und Ende des Enums
    Dim s As String
    s = "Public Enum "
    BeginPos = InStr(1, StrVal, s, vbTextCompare)
    If BeginPos = 0 Then
        s = "Private Enum "
        BeginPos = InStr(1, StrVal, s, vbTextCompare)
        If BeginPos > 0 Then mAccess = AccPrivate
    Else
        mAccess = AccPublic
    End If
    If BeginPos > 0 Then BeginPos = BeginPos + Len(s)
    EndPos = InStr(1, StrVal, "End Enum", vbTextCompare)
End Function
Private Function GetTheName(StrVal As String, nPos As Long) As String
    Dim spcPos As Long, clPos As Long, Pos As Long
    spcPos = InStr(nPos + 1, StrVal, " ")
    clPos = InStr(nPos, StrVal, vbCrLf, vbBinaryCompare)
    Pos = Min(spcPos, clPos)
    If Pos > 0 Then
        GetTheName = Mid$(StrVal, nPos, Pos - nPos)
    End If
    nPos = Pos
End Function

Public Function EnumClassToString() As String
    'Dieser Code is komplett .NET
    Const BaseName As String = "mBase"
    Dim ClassName As String: ClassName = "C" & mEnumName
    Dim RetName As String
    Dim i1 As String: i1 = "'" & String$((75 - 23 - 5 - Len(mEnumName)) \ 2, "#") & "'"
    Dim MyPreferredIndent As Long: MyPreferredIndent = mConstIndent '2 'select another Indent here
    Dim mi As String: mi = String$(MyPreferredIndent, " ")
    Dim sb As StringBuilder: Set sb = MNew.StringBuilder("Option Explicit"): sb.Append vbCrLf
    
    sb.Append("'####################'  Enum Class from EnumBuilder  '####################'").Append vbCrLf
    sb.Append(i1 & "  " & "Name of this Class: ").Append(ClassName & "  ").Append(i1).Append(vbCrLf).Append vbCrLf
    sb.Append(EnumToString).Append (vbCrLf) ' acctostring
    sb.Append (vbCrLf)
    sb.Append("Private ").Append(BaseName).Append(" As EEnum").Append vbCrLf
    sb.Append("Private mValue As ").Append(mEnumName).Append(vbCrLf).Append vbCrLf
    
    sb.Append("Private Sub Class_Initialize()").Append (vbCrLf)
    sb.Append(mi).Append("Call ").Append(BaseName).Append(".NewC(").Append(CStr(mEnumConsts.Count)).Append(")").Append vbCrLf
    sb.Append(mi).Append("Call InitEnum").Append vbCrLf
    sb.Append("End Sub").Append (vbCrLf)
    'SB.Append(vbCrLf)
    sb.Append("Private Sub Class_Terminate()").Append (vbCrLf)
    sb.Append(mi).Append("Set ").Append(BaseName).Append(" = Nothing").Append (vbCrLf)
    sb.Append("End Sub").Append(vbCrLf).Append (vbCrLf)
    
    If mRetClass Then RetName = ClassName Else RetName = mEnumName
    sb.Append("Public Property Let Value(EVal As ").Append(RetName).Append(")").Append vbCrLf
    sb.Append(mi).Append("mValue = EVal").Append vbCrLf
    sb.Append("End Property").Append vbCrLf
    sb.Append("Public Property Get Value() As ").Append(RetName).Append vbCrLf
    sb.Append(mi).Append("value = mValue").Append vbCrLf
    sb.Append("End Property").Append(vbCrLf).Append vbCrLf
    
    sb.Append("Private Sub InitEnum()").Append vbCrLf
    mEnumConsts.InitToString sb, BaseName, mi
    sb.Append("End Sub").Append(vbCrLf).Append vbCrLf
    
    If mPropGetConst Then
        Dim ASEC As String
        If mRetClass Then ASEC = ClassName Else ASEC = mEnumName
        Call mEnumConsts.PropGetToString(sb, ASEC, mi)
        Call sb.Append(vbCrLf)
    End If
    sb.Append("Public Function ToString() As String").Append vbCrLf
    sb.Append("  ToString = mBase.GetNameForVal(mValue)").Append vbCrLf
    sb.Append("End Function").Append vbCrLf
    
    EnumClassToString = sb.ToString
End Function

Public Function EEnumClassToString() As String
    Dim bA() As Byte: bA = LoadResData(1001, "Custom")
    EEnumClassToString = StrConv(bA, vbUnicode) 'LoadResData(101, 6)
End Function
'##############################' Implements ICanUndo '##############################'
Public Function GetSimpleUndo(aFlag As Long) As SimpleUndo
    Dim cm As ChangeMode: cm = aFlag 'easy Typumwandlung ;)
    Select Case cm
    Case cmAccess: Set GetSimpleUndo = MNew.SimpleUndo(cm, CStr(mAccess), Me)
    Case cmEnumNm: Set GetSimpleUndo = MNew.SimpleUndo(cm, mEnumName, Me)
    Case cmIndent: Set GetSimpleUndo = MNew.SimpleUndo(cm, CStr(mIndent), Me)
    Case cmCnsInd: Set GetSimpleUndo = MNew.SimpleUndo(cm, CStr(mConstIndent), Me)
    Case cmIcPrFx: Set GetSimpleUndo = MNew.SimpleUndo(cm, CbStr(mInclPreFix), Me)
    Case cmPreFix: Set GetSimpleUndo = MNew.SimpleUndo(cm, mPreFix, Me)
    Case cmPGCnst: Set GetSimpleUndo = MNew.SimpleUndo(cm, CbStr(mPropGetConst), Me)
    Case cmRetCls: Set GetSimpleUndo = MNew.SimpleUndo(cm, CbStr(mRetClass), Me)
    End Select
End Function
Private Function ICanUndo_GetSimpleUndo(aFlag As Long) As SimpleUndo
    Set ICanUndo_GetSimpleUndo = GetSimpleUndo(aFlag)
End Function
Public Sub UndoRedo(SU As SimpleUndo)
    Dim TmpStr As String, cm As ChangeMode: cm = SU.ChgFlag
    TmpStr = SU.URStr
    Select Case cm
    Case cmAccess: SU.URStr = CStr(mAccess):        mAccess = CLng(TmpStr)
    Case cmEnumNm: SU.URStr = mEnumName:            mEnumName = TmpStr
    Case cmIndent: SU.URStr = CStr(mIndent):        mIndent = CLng(TmpStr)
    Case cmCnsInd: SU.URStr = CStr(mConstIndent):   mConstIndent = CLng(TmpStr)
    Case cmIcPrFx: SU.URStr = CbStr(mInclPreFix):   mInclPreFix = CBool(TmpStr)
    Case cmPreFix: SU.URStr = mPreFix:              mPreFix = TmpStr
    Case cmPGCnst: SU.URStr = CbStr(mPropGetConst): mPropGetConst = CBool(TmpStr)
    Case cmRetCls: SU.URStr = CbStr(mRetClass):     mRetClass = CBool(TmpStr)
    End Select
    Call OnChanged
End Sub
Private Sub ICanUndo_Undo(SU As SimpleUndo)
    UndoRedo SU
End Sub
Private Sub ICanUndo_Redo(SU As SimpleUndo)
    UndoRedo SU
End Sub
