VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'diese Klasse muﬂ nur einmal im Projekt vorhanden sein
Private Type UDTEnum
    StrBez    As String
    LngVal    As Long 'oder Byte oder Integer...
    StrLggBez As String 'eine Alternative Bezeichnung, die in ListBox oder ComboBox angezeigt wird
End Type

Private mEnumArr() As UDTEnum

Friend Sub New_(aCount As Long)
    ReDim mEnumArr(0 To aCount - 1)
End Sub

Friend Sub SetEnumItem(ByVal index As Long, aStrBez As String, aLngVal As Long, Optional aStrLggBez As String)
    Dim n As Long: n = UBound(mEnumArr)
    If index > n Then ReDim Preserve mEnumArr(n + 1) 'bloﬂ um eins mehr OK
    mEnumArr(index) = NewUDTEnum(aStrBez, aLngVal)
End Sub

Friend Function NewUDTEnum(aStrBez As String, aLngVal As Long) As UDTEnum
    NewUDTEnum.StrBez = aStrBez: NewUDTEnum.LngVal = aLngVal
End Function

Public Function GetNameForVal(aVal As Long) As String
    Dim i As Long
    For i = 0 To UBound(mEnumArr)
        If mEnumArr(i).LngVal = aVal Then GetNameForVal = mEnumArr(i).StrBez: Exit Function
    Next
End Function

Public Function GetValForName(aBez As String) As Long
    Dim i As Long
    For i = 0 To UBound(mEnumArr)
        If mEnumArr(i).StrBez = aBez Then GetValForName = mEnumArr(i).LngVal: Exit Function
    Next
End Function
'
'Public Sub ToComboBox(VBComboBox As ComboBox)
'    ToBox VBComboBox
'End Sub

Public Sub ToListBox(CombBoxOrListBox) ' As ListBox)
'    ToBox CombBoxOrListBox
    Dim i As Long, StrVal As String
    For i = 0 To UBound(mEnumArr)
        StrVal = mEnumArr(i).StrLggBez
        If Len(StrVal) = 0 Then StrVal = mEnumArr(i).StrBez
        CombBoxOrListBox.AddItem StrVal
    Next
End Sub

'Private Sub ToBox(CBLB As Variant)
'    Dim i As Long, StrVal As String
'    For i = 0 To UBound(mEnumArr)
'        StrVal = mEnumArr(i).StrLggBez
'        If Len(StrVal) = 0 Then StrVal = mEnumArr(i).StrBez
'        Call CBLB.AddItem(StrVal)
'    Next
'End Sub
