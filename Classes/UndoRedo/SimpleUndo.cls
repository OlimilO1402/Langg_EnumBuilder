VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Klasse stellt eine simple atomare �nderung eines
'Objektes dar, das ICanUndo implementert.
'der ge�nderte wert wird in URStr gespeichert.
'mit ChgFlag kann ein Wert (z.B. aus einer Enum-Konstanten)
'gespeichert werden das eine Information �ber die Art der
'�nderung gibt.
Public ChgFlag As Long
'Public nCount  As Long
Public URStr   As String
Public ICU     As ICanUndo

