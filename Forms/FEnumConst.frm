VERSION 5.00
Begin VB.Form FEnumConst 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Enum Constant"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3735
   Icon            =   "FEnumConst.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox ChkLoadFromRessource 
      Caption         =   "Load from ressource"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3495
   End
   Begin VB.CheckBox ChkShowHexVal 
      Caption         =   "As hex value"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton BtnTake 
      Caption         =   "&Take"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox TxtConstDescription 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox TxtConstValue 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox TxtConstName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Additional description to the constant:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label LblConstValue 
      Caption         =   "Numeric Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label LblConstName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "FEnumConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mECB As EnumBuilderDoc 'eine Verbindung zum EnumBuilderDoc
'bei "OK" wird das Neue Element aktualisiert und der Dialog geschlossen
'bzw bei "Übernehmen", das aktuell ausgewählte Element aktualisiert.
'Public Enum DoWhat
'  DoAdd
'  DoSet
'End Enum

Private mBolAdd As Boolean 'was wird hier getan?
'True: ein neues hin zufügen
'False: ein bestehendes editieren

'Private Sub Form_Initialize()
'  Call Init(vbNullString, vbNullString, False, vbNullString, False)
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'  Set mEB = Nothing
'End Sub
'Private Sub Form_Terminate()
'  Set mEB = Nothing
'End Sub
'1. Das Formular wird neu geladen, um eine neue Konstante anzulegen
'   -> New_
'2. das Formular war bereits geladen, es wird aber noch eine Konstante angelegt
'   -> SetSelected
'3. eine vorhandene Konstante wird geändert
'   -> SetSelected
Friend Sub New_(aECB As EnumBuilderDoc, BolAdd As Boolean)
    Set mECB = aECB
    mBolAdd = BolAdd
    If mBolAdd Then
        Init vbNullString, vbNullString, False, vbNullString, False
    Else
        SetSelected
    End If
End Sub

Private Sub SetSelected()
    With mECB.SelectedItem
        Init .Name, .ValToStr, .AsHexVal, .AddDescr, .LoadFRes
    End With
End Sub

Private Sub Init(aName As String, aVal As String, ashex As Boolean, aDesc As String, LFRes As Boolean)
    TxtConstName.Text = aName
    TxtConstValue.Text = aVal
    ChkShowHexVal.Value = Bol2CheckBoxValue(ashex)
    TxtConstDescription.Text = aDesc
    ChkLoadFromRessource.Value = Bol2CheckBoxValue(LFRes)
End Sub

Private Sub SetEnumConst()
    If Len(TxtConstName.Text) > 0 Then
        If mBolAdd Then
            mECB.AddConst MNew.EnumConst
        End If
        'Else
        mECB.SetSelectedItem TxtConstName.Text, TxtConstValue.Text, (ChkShowHexVal.Value = vbChecked), TxtConstDescription.Text, (ChkLoadFromRessource.Value = vbChecked)
        'End If
    End If
End Sub

Private Sub BtnOK_Click()
    SetEnumConst
    Unload Me
End Sub

Private Sub BtnTake_Click()
    SetEnumConst
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

