VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCompilerError"
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

Public Enum eCompilerError
    ece_null = 0
    ece_AssignmentToConstantNotPermitted
    ece_AssignmentToFunctionNotPermitted
    ece_ConstantMustHaveAValue
    ece_DuplicateDeclarationInCurrentScope
    ece_AmbiguousNameDetected
    ece_VariableNotDefined
    ece_NotImplementedYet
    ece_SubCannotReturnAValue
    ece_CannotUseReservedWord
    ece_AccessModifiersNotAllowedHere
    ece_CannotRedefineReservedWords
    ece_WrongNumbersOfParameters
End Enum

'Solleva l'errore
Public Sub RaiseError(ByVal iErr As eCompilerError, _
    ByVal lRow As Long, ByVal lCol As Long)
    
    Err.Raise C_ERR_COMP, , "(" & Hex(iErr) & ") " & ErrToString(iErr) & " Line: " & lRow & " Col: " & lCol
End Sub

Public Function ErrToString(ByVal iErr As eCompilerError) As String
    Select Case iErr
        Case ece_AssignmentToConstantNotPermitted
            ErrToString = "Assignment To Constant Not Permitted"
        Case ece_AssignmentToFunctionNotPermitted
            ErrToString = "Assignment To Function/Sub Not Permitted"
        Case ece_ConstantMustHaveAValue
            ErrToString = "Constant Must Have a Value"
        Case ece_DuplicateDeclarationInCurrentScope
            ErrToString = "Variable/Function Already Defined"
        Case ece_AmbiguousNameDetected
            ErrToString = "Ambiguous Name Detected"
        Case ece_VariableNotDefined
            ErrToString = "Variable/Function Not Defined"
        Case ece_NotImplementedYet
            ErrToString = "Not Implemented Yet"
        Case ece_SubCannotReturnAValue
            ErrToString = "Sub Cannot Return a Value"
        Case ece_CannotUseReservedWord
            ErrToString = "Cannot Use 'Reserved' Word"
        Case ece_AccessModifiersNotAllowedHere
            ErrToString = "Access Modifiers Not Allower Here"
        Case ece_CannotRedefineReservedWords
            ErrToString = "Cannot Redefine Reserved Words"
        Case ece_WrongNumbersOfParameters
            ErrToString = "Wrong Number of Parameters"
    End Select
End Function

