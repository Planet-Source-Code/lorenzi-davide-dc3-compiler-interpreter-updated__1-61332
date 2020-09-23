VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFuncInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Part of DC3 Compiler - Interpreter
' Author: Lorenzi Davide (http://www.hexagora.com)
' See the file 'license.txt' for informations
'================================================================================

Option Explicit

'Simbol Table Locale
Public oSymbTable As cSymbTable
'Collection di cFuncParam
Public oCollParams As Collection
'Puntatore all'istruzione in cui si trova la funzione
Public lInstrPointer As Long

Private Sub Class_Initialize()
    Set oSymbTable = New cSymbTable
    Set oCollParams = New Collection
End Sub

Public Function Add(oParam As cFuncParam) As cFuncParam
    oCollParams.Add oParam
    
    Set Add = oParam
End Function

Public Function AddParam(ByVal iParmMod As eParmMod, sParmName As String) As cFuncParam
    Dim oParm As New cFuncParam
    
    oParm.iParmMod = iParmMod
    oParm.sParmName = sParmName

    Set AddParam = Add(oParm)
End Function
