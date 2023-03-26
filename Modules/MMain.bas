Attribute VB_Name = "MMain"
Option Explicit
'// Das einzige EnumBuilderApp-Objekt
Public EnumBApp As EnumBuilderApp 'theApp DooBAppDooWApp a DooBAppDooWApp

Sub Main()
    Set EnumBApp = New EnumBuilderApp
    EnumBApp.InitRun
End Sub
