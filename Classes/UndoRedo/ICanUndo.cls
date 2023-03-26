VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICanUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Klasse ICanUndo ' aha - I Can Boogie :)
'Klasse ist Schnittstelle, ohne eigene Funktionalität

Public Function GetSimpleUndo(aFlag As Long) As SimpleUndo
  '
End Function

Public Sub Undo(SU As SimpleUndo)
  '
End Sub

Public Sub Redo(SU As SimpleUndo)
  '
End Sub
