VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SingleDocTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mDocument As EnumBuilderDoc
Private mForm     As FMain
Private mView     As EnumBuilderView

Friend Sub New_(aDoc As EnumBuilderDoc, aForm As FMain, aView As EnumBuilderView)
    Set mDocument = aDoc
    Set mView = aView
    Set mForm = aForm
    mDocument.Consts.BaseCol.Add MNew.EnumConst(DefConstN, DefConstV, True, DefConstD, False)
    mView.New_ mDocument ', aForm)
    'Call mDocument.AddView(mView)
End Sub
Public Sub Init()
    Call mView.Init(mForm)
End Sub

Public Sub NewDocument()
    New_ MNew.EnumBuilderDoc, mForm, MNew.EnumBuilderView
    Init
    mDocument.UpdateAllViews vmNone
End Sub

Public Property Get CurrentDoc() As EnumBuilderDoc
    Set CurrentDoc = mDocument
End Property
Public Property Get CurrentView() As EnumBuilderView
    Set CurrentView = mView
End Property

