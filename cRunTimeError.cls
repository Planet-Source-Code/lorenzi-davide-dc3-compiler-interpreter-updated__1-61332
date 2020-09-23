VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRunTimeError"
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
'Classe di gestione degli errori di compilazione
'
'================================================================================

Option Explicit

Public Enum eRuntimeError
    ere_null = 0
    ere_TypeMismatch
End Enum

'Solleva l'errore
Public Sub RaiseError(ByVal iErr As eRuntimeError, _
    ByVal lRow As Long)
    
    Err.Raise C_ERR_RUNTIME, , "(" & Hex(iErr) & ") " & ErrToString(iErr) & " Line: " & lRow
End Sub

Public Function ErrToString(ByVal iErr As eRuntimeError) As String
    Select Case iErr
        Case ere_TypeMismatch
            ErrToString = "Type Mismatch"
    End Select
End Function



