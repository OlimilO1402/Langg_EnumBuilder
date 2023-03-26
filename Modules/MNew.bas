Attribute VB_Name = "MNew"
Option Explicit
Public Enum AccPublPriv 'mit einem Private Enum ist diese Klasse nur eingeschränkt sinnvoll
    AccPublic  'Public
    AccPrivate 'Private
End Enum
Public Enum ViewMode
    vmNone = &H0
    vmOptionEnum = &H1
    vmEditEnum = &H2
    vmClassEnum = &H4
    vmAll = &H7
End Enum

'Modify default consts here, if you like others
Public Const DefAccess   As Long = 0 'AccPublic 'Achtung hier später durch EnumClass ersetzen!!
Public Const DefEnumName As String = "EnumName"
Public Const DefIndent   As Long = 0 '4
Public Const DefCnstInd  As Long = 4
Public Const DefShowHex  As Boolean = True
Public Const DefIncPrFx  As Boolean = True 'False
Public Const DefPreFix   As String = DefEnumName & "_" '"EnumName_"
Public Const DefPropGet  As Boolean = True
Public Const DefRetCls   As Boolean = True
Public Const DefConstN   As String = "Const1"
Public Const DefConstV   As String = "1"
Public Const DefConstD   As String = "a description"
'ein paar Enums zum Parsenüben für Import-Funktion
Private Enum EEE 'noch ein Kommentar
    E1 'noch ein Kommentar
    'nur ein Kommentar
    E2 = 1
    E3 = &H2     'mit Komemntar
    E4 = &H3
    E5 = 27 'EE5
End Enum
Private Enum E 'noch ein Kommentar
    
    egt_A1 'noch ein Kommentar
    'nur ein Kommentar
    egt_B2 = 1
    egt_C3 = &H2     'mit Komemntar
    egt_D4 = &H3
    egt_E5 = 27 'EE5
    
End Enum

Public Function SingleDocTemplate(aDoc As EnumBuilderDoc, aForm As FMain, aView As EnumBuilderView) As SingleDocTemplate
    Set SingleDocTemplate = New SingleDocTemplate: SingleDocTemplate.New_ aDoc, aForm, aView
End Function

Public Function EnumBuilderView(Optional aDoc As EnumBuilderDoc) As EnumBuilderView
    Set EnumBuilderView = New EnumBuilderView: EnumBuilderView.New_ aDoc
End Function

Public Function EnumBuilderDoc(Optional aAcc As AccPublPriv = DefAccess, Optional aEnumName As String = DefEnumName, Optional aIndent As Long = DefIndent, Optional aConstIndent As Long = DefCnstInd, Optional bIncPrFx As Boolean = DefIncPrFx, Optional aPreFix As String = DefPreFix, Optional bPropGet As Boolean = DefPropGet, Optional bRetClass As Boolean = DefRetCls) As EnumBuilderDoc
  Set EnumBuilderDoc = New EnumBuilderDoc: EnumBuilderDoc.New_ aAcc, aEnumName, aIndent, aConstIndent, bIncPrFx, aPreFix, bPropGet
End Function

Public Function EnumBuilderDocA(SrcECB As EnumBuilderDoc) As EnumBuilderDoc
    Set EnumBuilderDocA = New EnumBuilderDoc: EnumBuilderDocA.Assign SrcECB
End Function

Public Function EnumBuilderDocI(ImpSrcStr As String) As EnumBuilderDoc
    Set EnumBuilderDocI = New EnumBuilderDoc: EnumBuilderDocI.Import ImpSrcStr
End Function

Public Function EnumConst(Optional aName As String, Optional aValue As String, Optional bAsHexVal As Boolean, Optional sAddDescription As String, Optional bLoadRes As Boolean) As EnumConst
    Set EnumConst = New EnumConst:  EnumConst.New_ aName, aValue, bAsHexVal, sAddDescription, bLoadRes
End Function

Public Function StringBuilder(Optional ByVal value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilder
    Set StringBuilder = New StringBuilder: StringBuilder.New_ value, startIndex, Length, Capacity, maxCapacity
End Function

'Public Function New_SimpleUndo(aFlag As Long, aCount As Long, aURStr As String, aICU As ICanUndo) As SimpleUndo
Public Function SimpleUndo(aFlag As Long, aURStr As String, aICU As ICanUndo) As SimpleUndo
    Set SimpleUndo = New SimpleUndo: SimpleUndo.ChgFlag = aFlag: SimpleUndo.URStr = aURStr: Set SimpleUndo.ICU = aICU
    'New_SimpleUndo.nCount = aCount
End Function

'Public Function Min(LngVal1 As Long, LngVal2 As Long) As Long
'    If LngVal1 < LngVal2 Then Min = LngVal1 Else Min = LngVal2
'End Function
'Public Function Max(LngVal1 As Long, LngVal2 As Long) As Long
'    If LngVal1 > LngVal2 Then Max = LngVal1 Else Max = LngVal2
'End Function
'Public Function maxs(SngVal1 As Single, SngVal2 As Single) As Single
'  If SngVal1 > SngVal2 Then maxs = SngVal1 Else maxs = SngVal2
'End Function

Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

Public Function CbStr(BolVal As Boolean) As String
    If BolVal Then CbStr = "1" Else CbStr = "0"
End Function

Public Function Bol2CheckBoxValue(BolVal As Boolean) As CheckBoxConstants
    If BolVal Then Bol2CheckBoxValue = vbChecked Else Bol2CheckBoxValue = vbUnchecked
End Function

Public Sub DeleteObj(aObj As Object)
    Set aObj = Nothing
End Sub

Public Function RepAByAA(aStrVal As String) As String
    'ersetzt ein AmpersAnd = "&" durch ein doppeltes "&&"
    'das ist nötig, damit in einem Label ein einzelnes "&" als solches
    'darstellt, ansonsten wird ein unterstrichener Buchstabe angezeigt
    Dim AAPos As Long 'AA = AmpersAnd
    AAPos = 1
    While AAPos > 0
        AAPos = InStr(AAPos, aStrVal, "&", vbTextCompare)
        If AAPos > 0 Then
            aStrVal = Left$(aStrVal, AAPos) & "&" & Right$(aStrVal, Len(aStrVal) - AAPos)
            AAPos = AAPos + 2
        End If
    Wend
    RepAByAA = aStrVal
End Function
'Private Function RepAAByA(aStrVal As String) As String
'Dim AAPos As Long
'  AAPos = 1
'  While AAPos > 0
'    AAPos = InStr(AAPos, aStrVal, "&&", vbTextCompare)
'    If AAPos > 0 Then
'      aStrVal = Left$(aStrVal, AAPos) & Right$(aStrVal, Len(aStrVal) - AAPos - 1)
'      AAPos = AAPos + 1
'    End If
'  Wend
'  RepAAByA = aStrVal
'End Function

