VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResSymbols"
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
'Classe per la gestione delle funzioni riservate chiamate dall'istro PCALL
'
'Fornisce la traduzione della funzione riservata al codice associato
'
'================================================================================

Option Explicit

'Write symbols in lower-case
Private Const C_S_Print = "print"
Private Const C_S_MsgBox = "msgbox"
Private Const C_F_CLng = "clng"
Private Const C_F_CStr = "cstr"
Private Const C_F_InputBox = "inputbox"
Private Const C_F_Rnd = "rnd"

Public Enum ePFunc
    eps_Print = 0
    eps_MsgBox
    epf_CLng
    epf_CStr
    epf_InputBox
    epf_Rnd
End Enum

Public Function GetPFuncCode(oSymbEl As cSymbTableEl) As ePFunc
    If oSymbEl.iAccMod <> eam_reserved Then
        Err.Raise vbObjectError + 1, "cPFunc.GetPFuncCode", "This is not a Reserved Symbol"
    End If

    Select Case LCase(oSymbEl.sName)
        Case C_S_Print: GetPFuncCode = eps_Print
        Case C_S_MsgBox: GetPFuncCode = eps_MsgBox
        Case C_F_CLng: GetPFuncCode = epf_CLng
        Case C_F_CStr: GetPFuncCode = epf_CStr
        Case C_F_InputBox: GetPFuncCode = epf_InputBox
        Case C_F_Rnd: GetPFuncCode = epf_Rnd
        Case Else
            Err.Raise vbObjectError + 1, "cPFunc.GetPFuncCode", "Unknown Reserved Symbol"
    End Select
End Function

