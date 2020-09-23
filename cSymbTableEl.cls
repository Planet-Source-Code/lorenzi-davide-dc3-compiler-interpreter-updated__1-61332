VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSymbTableEl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Part of DC3 Compiler - Interpreter
' Author: Lorenzi Davide (http://www.hexagora.com)
' See the file 'license.txt' for informations
'================================================================================
'
' Rappresenta un elemento all'interno della symbol table
'
'================================================================================

Option Explicit

Public iType As eSymbTableElType
Public iAccMod As eAccessModifier
Public sName As String
Public oFuncInfos As cFuncInfos

Private Sub Class_Initialize()
    iType = estet_null
    iAccMod = eam_private
    sName = ""
    Set oFuncInfos = Nothing
End Sub
