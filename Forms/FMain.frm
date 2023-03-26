VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "EnumBuilder"
   ClientHeight    =   5355
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7245
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   5355
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnShowEEnum 
      Caption         =   "Show EEnum Class"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton BtnShowEnumClass 
      Caption         =   "Show Enum Class"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton BtnShowEnumModule 
      Caption         =   "Show Enum Module"
      Height          =   375
      Left            =   1800
      TabIndex        =   38
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton BtnShowEditMode 
      Caption         =   "Edit Mode"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.PictureBox PbViewEnumEdit 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   5715
      TabIndex        =   22
      Top             =   2400
      Width           =   5775
      Begin VB.ListBox LstPublPriv 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         ItemData        =   "FMain.frx":000C
         Left            =   120
         List            =   "FMain.frx":000E
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtEnumName 
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox LstEnumConst 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FMain.frx":0010
         Left            =   120
         List            =   "FMain.frx":0012
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label LblEnumIndent 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   15
      End
      Begin VB.Label LblEnumName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "EName"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1920
         TabIndex        =   30
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label LblPublPriv 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Public"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   720
      End
      Begin VB.Label LblEnum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " Enum "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1200
         TabIndex        =   28
         Top             =   120
         Width           =   720
      End
      Begin VB.Label LblEnumConst 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DiesesIstDerErsteSTreichDerZweiteFolgtSogleich"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   5520
      End
      Begin VB.Label LblEndEnum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "End Enum"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   960
      End
   End
   Begin VB.PictureBox PbViewEnumOption 
      BorderStyle     =   0  'Kein
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1815
      ScaleWidth      =   5895
      TabIndex        =   4
      Top             =   480
      Width           =   5895
      Begin VB.CommandButton BtnRedo 
         Caption         =   "R"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         ToolTipText     =   "Redo"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BtnUndo 
         Caption         =   "U"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Undo"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BtnBitRow 
         Caption         =   "BitRow"
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         ToolTipText     =   "Set values of constants to &H1, &H2, &H4, &H8, &H10, &H20 ..."
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton BtnImportClipBoard 
         Caption         =   "Import"
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         ToolTipText     =   "Import From ClipBoard a String that cotains a Visual Basic Enum"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton BtnInsertBefore 
         Caption         =   "Insert"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton BtnDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox CmbIndent 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox CmbConstIndent 
         Height          =   315
         ItemData        =   "FMain.frx":0014
         Left            =   3000
         List            =   "FMain.frx":0016
         TabIndex        =   18
         Top             =   0
         Width           =   735
      End
      Begin VB.CheckBox ChkInclPreFix 
         Caption         =   "Const Prefix:"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TxtPreFix 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox ChkConstPropGet 
         Caption         =   "Put <Property Get> of All Const in Enum Class"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptRetEnum 
         Caption         =   "Return As Enum"
         Enabled         =   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton OptRetClass 
         Caption         =   "Return As Class"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Frame FraMove 
         Caption         =   "Move Selected Const"
         Height          =   1335
         Left            =   3840
         TabIndex        =   12
         Top             =   0
         Width           =   1935
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'Kein
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   1455
            TabIndex        =   33
            Top             =   360
            Width           =   1455
            Begin VB.CommandButton BtnMoveUp 
               Caption         =   "^"
               Height          =   375
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton BtnMoveDown 
               Caption         =   "v"
               Height          =   375
               Left            =   0
               TabIndex        =   34
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label3 
               Caption         =   "Move Up"
               Height          =   255
               Left            =   480
               TabIndex        =   37
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Move Down"
               Height          =   255
               Left            =   480
               TabIndex        =   36
               Top             =   480
               Width           =   975
            End
         End
      End
      Begin VB.CommandButton BtnNew 
         Caption         =   "New"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label LblUR 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3840
         TabIndex        =   32
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label LblIndent 
         Caption         =   "Enum Indent:"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LblConstIndent 
         Caption         =   "Const Indent:"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox TbViewEnumClass 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New VB Enum"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import VB Enum From Clipboard"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewEnumEdit 
         Caption         =   "Visual Enum Editor"
      End
      Begin VB.Menu mnuViewEnumClass 
         Caption         =   "Switch To Enum Class"
      End
   End
   Begin VB.Menu mnuEnumConsts 
      Caption         =   "Enum&Consts"
      Begin VB.Menu mnuEnumConstsEditSel 
         Caption         =   "Edit Selected Constant"
      End
      Begin VB.Menu mnuEnumConstsAddNew 
         Caption         =   "Add New Enum Constant"
      End
      Begin VB.Menu mnuEnumConstsDeleteSel 
         Caption         =   "Delete Selected Enum Constant"
      End
      Begin VB.Menu mnuEnumConstsInsertBefSel 
         Caption         =   "Insert Before Selected"
      End
      Begin VB.Menu mnuEnumConstsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnumConstsMoveUp 
         Caption         =   "Move ^ Selected Up"
      End
      Begin VB.Menu mnuEnumConstsMoveDown 
         Caption         =   "Move v Selected Down"
      End
   End
   Begin VB.Menu mnuExtr 
      Caption         =   "E&xtras"
      Begin VB.Menu mnuExtrBitRow 
         Caption         =   "Make Bit Row"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &? "
      Begin VB.Menu mnuHelpThemes 
         Caption         =   "Help Themes"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info About..."
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aPP As CAccPublPriv

Private Sub Form_Initialize()
    Set aPP = New CAccPublPriv
    aPP.ToListBox Me.LstPublPriv
    'EnumBApp.EnableVisualStyles
    Me.Icon = LoadResPicture("0", vbResIcon)
End Sub

'##############################'    FMain    '##############################'
'Private Sub Form_Load()
'  'Call EnumBApp.CurrentDoc.UpdateAllViews(vmAll)
'  'Call InitPublPriv
'  'Call InitEnumName
'  Call InitIndent
'  Call InitConstIndent
'  'Call InitEnumCnst
'  'Call InitOpt
'  'Call InitUndoRedo
'  'Call Updateenumwnd
'End Sub

'Private Sub Form_Resize()
'Dim Brdr As Long: Brdr = 120
'  PbViewEdit.Move PbViewEdit.Left, PbViewEdit.Top, Me.ScaleWidth - 2 * Brdr, Me.ScaleHeight - PbViewEdit.Top - Brdr
'  'PbEnum.Move PbEnum.Left, PbEnum.Top, Me.ScaleWidth - 2 * Brdr, Me.ScaleHeight - PbEnum.Top - Brdr
'  TxtEnumClass.Move PbViewEdit.Left, PbViewEdit.Top, Me.ScaleWidth - 2 * Brdr, Me.ScaleHeight - PbViewEdit.Top - Brdr
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Abfrage ob gespeichert werden soll entfällt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not EnumBApp Is Nothing Then EnumBApp.OnFileExit
End Sub

'##############################'    Init    '##############################'
'Private Sub InitEnumCnst()
'  Call mEBuilder.AddConst(New_EnumConst("LstEnumConst", CLng(&H1), True, "LstEnumConst Auf Deutsch", False))
'End Sub
'Private Sub InitIndent()
'  Call InitCmb(CmbIndent, mEBuilder.Indent)
'End Sub
'Private Sub InitConstIndent()
'  Call InitCmb(CmbConstIndent, mEBuilder.ConstIndent)
'End Sub
'Private Sub InitCmb(CB As ComboBox, ind As Long)
'Dim i As Long
'  For i = 0 To 9
'    Call CB.AddItem(CStr(i))
'  Next
'  CB.ListIndex = ind
'End Sub
'Private Sub InitPublPriv()
'  LstPublPriv.Visible = False
'  LstPublPriv.ZOrder 1
'  LstPublPriv.ListIndex = 0
'End Sub
'Private Sub InitEnumName()
'  LblEnumName.Caption = mEBuilder.EnumName
'  TxtEnumName.Text = LblEnumName.Caption
'  TxtEnumName.Visible = False
'  TxtEnumName.ZOrder 1
'End Sub
'Private Sub InitOpt()
'  OptRetEnum.Enabled = True 'False
'  OptRetClass.Enabled = True 'False
'End Sub

Private Sub ChkConstPropGet_Click()
    EnumBApp.CurrentDoc.PropGetConst = (ChkConstPropGet.Value = vbChecked)
    OptRetEnum.Enabled = Not OptRetEnum.Enabled 'True:
    OptRetClass.Enabled = Not OptRetClass.Enabled 'True
    If EnumBApp.CurrentDoc.PropGetConst Then OptRetClass.Value = True
End Sub

'##############################'   PreFix   '##############################'
Private Sub ChkInclPreFix_Click()
    EnumBApp.CurrentDoc.IncludePreFix = (ChkInclPreFix.Value = vbChecked)
    'Call SetPreFix
End Sub

Private Sub TxtPreFix_KeyDown(KeyCode As Integer, Shift As Integer)
  'If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyTab) Then
'  If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
'    Call SetPreFix
'  End If
End Sub
Private Sub TxtPreFix_KeyPress(KeyAscii As Integer)
    '
End Sub
Private Sub TxtPreFix_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        SetPreFix
    End If
End Sub
Private Sub SetPreFix()
    If Len(TxtPreFix.Text) > 0 Then
        EnumBApp.CurrentDoc.PreFix = TxtPreFix.Text
    End If
End Sub

'##############################'  PublPriv  '##############################'
Private Sub LblPublPriv_Click()
    LstPublPriv.Visible = True
    LstPublPriv.ZOrder 0
    If LblPublPriv.Caption = "Public" Then
        LstPublPriv.ListIndex = 0
    Else
        LstPublPriv.ListIndex = 1
    End If
End Sub
Private Sub LstPublPriv_Click()
    Call SetPublPriv
End Sub
Private Sub SetPublPriv()
    Dim a As AccPublPriv
    If (LstPublPriv.ListIndex <= 1) Then a = AccPublic Else a = AccPrivate
    If Not EnumBApp.CurrentDoc Is Nothing Then EnumBApp.CurrentDoc.Access = a
    LstPublPriv.Visible = False
End Sub

'##############################'  EnumBuilder_Changed  '##############################'
'Private Sub mEBuilder_Changing(cm As ChangeMode)
'  If Not UndoRedo Is Nothing Then
'    If Not UndoRedo.Undoing Then
'      If Not EnumBApp.EnumBDoc Is Nothing Then
'        Call UndoRedo.SaveUndo(mEBuilder, cm)
'      End If
'    End If
'  End If
'End Sub

'##############################'    UndoRedoBtn    '##############################'
Public Sub EnDisUndoRedoBtn()
    If Not EnumBApp.UndoRedo Is Nothing Then
        BtnUndo.Enabled = EnumBApp.UndoRedo.UndoEnabled
        BtnRedo.Enabled = EnumBApp.UndoRedo.RedoEnabled
        mnuEditUndo.Enabled = BtnUndo.Enabled
        mnuEditRedo.Enabled = BtnRedo.Enabled
        LblUR.Caption = EnumBApp.UndoRedo.GetURStr
    End If
End Sub
'##############################'  EnumName  '##############################'
Private Sub TxtEnumName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SetEnumName
    End If
End Sub
Private Sub LblEnumName_Click()
    TxtEnumName.Visible = True
    TxtEnumName.ZOrder 0
    TxtEnumName.Text = LblEnumName.Caption
    Call TxtEnumName.Move(LblEnumName.Left, LblEnumName.Top)
End Sub
Private Sub SetEnumName()
    EnumBApp.CurrentDoc.EnumName = TxtEnumName.Text
    TxtEnumName.Visible = False
End Sub

'###########################'  PbViewEnumEdit  '###########################'
Private Sub PbViewEnumEdit_Click()
    'Call SetPublPriv
    Call SetEnumName
    Call SetEnumConsts
    'Call SetPreFix
    EnumBApp.CurrentDoc.SelectedIndex = 0
End Sub

Private Sub LblEndEnum_Click()
    PbViewEnumEdit_Click
End Sub

'##############################'  EnumConst  '##############################'
Private Sub LblEnumConst_Click()
    LstEnumConst.Visible = True
    LstEnumConst.ZOrder 0
End Sub
Private Sub LstEnumConst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        LblEnumConst.Visible = True
        LstEnumConst.Visible = False
        LstEnumConst.ZOrder 1
        LblEnumConst.ZOrder 0
    End If
End Sub
Private Sub SetEnumConsts()
    LstEnumConst.Visible = False
    LstEnumConst.ZOrder 1
End Sub

'Private Sub LblToLst(aLbl As Label, aLst As ListBox)
''schreibt den Text aus einem Label in eine ListBox
'Dim CLPos As Long, OCLPos As Long
'Dim StrVal As String, StrLin As String
'  StrVal = aLbl.Caption
'  aLst.Clear
'  If Len(StrVal) > 0 Then
'    'der erste Eintrag:
'    CLPos = 1
'    OCLPos = 1
'    CLPos = InStr(CLPos, StrVal, vbCrLf, vbBinaryCompare)
'    If CLPos = 0 Then
'      'einfach den ganzen string
'      Call aLst.AddItem(RepAAByA(StrVal))
'    Else
'      StrLin = Mid$(StrVal, OCLPos, CLPos - OCLPos)
'      Call aLst.AddItem(RepAAByA(StrLin))
'      OCLPos = CLPos
'      OCLPos = CLPos
'      While CLPos > 0
'        CLPos = InStr(CLPos, StrVal, vbCrLf, vbBinaryCompare)
'        If CLPos > 0 Then
'          StrLin = Mid$(StrVal, OCLPos, CLPos - OCLPos)
'          Call aLst.AddItem(RepAAByA(StrLin))
'          OCLPos = CLPos
'        End If
'      Wend
'    End If
'  End If
'End Sub
'Private Sub LstToLbl(aLst As ListBox, aLbl As Label)
''schreibt den Text aus einer ListBox in einen Label
''is bissl einfacher als der umgekehrte Weg
'Dim i As Long, StrVal As String
'  For i = 0 To aLst.ListCount - 1
'    StrVal = StrVal & RepAByAA(aLst.List(i)) 'AmpersAnd verdoppeln
'    If i < aLst.ListCount - 1 Then StrVal = StrVal & vbCrLf
'  Next
'  aLbl.Caption = StrVal
'End Sub

'##############################'   CmbIndent   '##############################'
Private Sub CmbIndent_Click()
    EnumBApp.CurrentDoc.Indent = CLng(CmbIndent.Text)
End Sub
Private Sub CmbConstIndent_Click()
    EnumBApp.CurrentDoc.ConstIndent = CLng(CmbConstIndent.Text) ' = CmbConstIndent.ListIndex '- 1
End Sub

'##############################'     OptRet    '##############################'
Private Sub OptRetClass_Click()
    EnumBApp.CurrentDoc.ReturnClass = True
End Sub
Private Sub OptRetEnum_Click()
    EnumBApp.CurrentDoc.ReturnClass = False
End Sub


'###########################' Menu + Button Handler ##########################'
'##############################'    mnuFile    '##############################'
'---------------------------------
Private Sub mnuFileNew_Click()
    EnumBApp.OnFileNew True
End Sub
Private Sub mnuFileImport_Click()
    EnumBApp.CurrentDoc.Import Clipboard.GetText
End Sub
Private Sub BtnImportClipBoard_Click()
    EnumBApp.CurrentDoc.Import Clipboard.GetText
End Sub
'---------------------------------
Private Sub mnuFileExit_Click()
    EnumBApp.OnFileExit
End Sub
'---------------------------------

'##############################'    mnuEdit    '##############################'
'---------------------------------
Private Sub mnuEditUndo_Click()
    EnumBApp.UndoRedo.Undo
End Sub
Private Sub BtnUndo_Click()
    EnumBApp.UndoRedo.Undo
End Sub
Private Sub mnuEditRedo_Click()
    EnumBApp.UndoRedo.Redo
End Sub
Private Sub BtnRedo_Click()
    EnumBApp.UndoRedo.Redo
End Sub
'---------------------------------
Private Sub mnuEditCut_Click()
    nyi
End Sub
Private Sub mnuEditCopy_Click()
    nyi
End Sub
Private Sub mnuEditPaste_Click()
    'nyi
    mnuFileImport_Click
End Sub

Private Sub nyi()
    MsgBox "Not Yet Implemented"
End Sub
'##############################'    mnuView    '##############################'
'---------------------------------
Private Sub mnuViewEnumEdit_Click()
    EnumBApp.CurrentView.ViewMode = vmEditEnum
End Sub
Private Sub BtnShowEditMode_Click()
    EnumBApp.CurrentView.ViewMode = vmEditEnum
End Sub
Private Sub mnuViewEnumClass_Click()
    EnumBApp.CurrentView.ViewMode = vmClassEnum
End Sub
Private Sub BtnShowEnumClass_Click()
    EnumBApp.CurrentView.ViewMode = vmClassEnum
End Sub
Private Sub BtnShowEEnum_Click()
    EnumBApp.CurrentView.ViewMode = vmClassEnum
    TbViewEnumClass.Text = EnumBApp.CurrentDoc.EEnumClassToString
End Sub
'---------------------------------

'##############################'    mnuEnum    '##############################'
'---------------------------------
Private Sub mnuEnumConstsAddNew_Click()
    AddNewConst
End Sub
Private Sub BtnNew_Click()
    AddNewConst
End Sub
Private Sub AddNewConst()
    EnumBApp.EnumConstantWnd.New_ EnumBApp.CurrentDoc, True
    EnumBApp.EnumConstantWnd.Show vbModeless, Me
End Sub
Private Sub mnuEnumConstsEditSel_Click()
    EditSelectedConst
End Sub
Private Sub EditSelectedConst()
    Call SetSelected
    EnumBApp.EnumConstantWnd.New_ EnumBApp.CurrentDoc, False
    EnumBApp.EnumConstantWnd.Show vbModeless, Me
End Sub
Private Sub SetSelected()
    If LstEnumConst.ListIndex < 0 Then LstEnumConst.ListIndex = 0
    EnumBApp.CurrentDoc.SelectedIndex = LstEnumConst.ListIndex + 1
End Sub
Private Sub LstEnumConst_Click()
    SetSelected
End Sub
Private Sub LstEnumConst_DblClick()
    EditSelectedConst
End Sub
Private Sub mnuEnumConstsDeleteSel_Click()
    EnumBApp.CurrentDoc.DeleteSelectedConst
End Sub
Private Sub BtnDelete_Click()
    EnumBApp.CurrentDoc.DeleteSelectedConst
End Sub
Private Sub mnuEnumConstsInsertBefSel_Click()
    EnumBApp.CurrentDoc.InsertConst MNew.EnumConst()
End Sub
Private Sub BtnInsertBefore_Click()
    EnumBApp.CurrentDoc.InsertConst MNew.EnumConst()
End Sub
'---------------------------------
Private Sub mnuEnumConstsMoveUp_Click()
    EnumBApp.CurrentDoc.MoveSelectedUp
    LstEnumConst.ListIndex = EnumBApp.CurrentDoc.SelectedIndex - 1
End Sub
Private Sub BtnMoveUp_Click()
    EnumBApp.CurrentDoc.MoveSelectedUp
    LstEnumConst.ListIndex = EnumBApp.CurrentDoc.SelectedIndex - 1
End Sub
Private Sub mnuEnumConstsMoveDown_Click()
    EnumBApp.CurrentDoc.MoveSelectedDown
    LstEnumConst.ListIndex = EnumBApp.CurrentDoc.SelectedIndex - 1
End Sub
Private Sub BtnMoveDown_Click()
    EnumBApp.CurrentDoc.MoveSelectedDown
    LstEnumConst.ListIndex = EnumBApp.CurrentDoc.SelectedIndex - 1
End Sub
'---------------------------------

'##############################'    mnuExtr    '##############################'
'---------------------------------
Private Sub mnuExtrBitRow_Click()
    EnumBApp.CurrentDoc.MakeBitRow
End Sub
Private Sub BtnBitRow_Click()
    EnumBApp.CurrentDoc.MakeBitRow
End Sub
'---------------------------------

'##############################'    mnuInfo    '##############################'
'---------------------------------
Private Sub mnuHelpInfo_Click()
    EnumBApp.OnAppAbout
End Sub

