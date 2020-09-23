VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStack"
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
'Implementazione di uno stack per variabili complesse
'
'================================================================================

Option Explicit

Private Const C_START_EL = 100

Private aStack() As tStackEl    'Array Elementi
Private sp As Long              'Stack Pointer
Private smax As Long            'Stack Max

Private Sub Class_Initialize()
    smax = C_START_EL
    ReDim aStack(0 To smax) As tStackEl
    sp = -1
End Sub

Private Sub Class_Terminate()
    Erase aStack
End Sub

'Rende il numero di elementi presenti nello stack
Public Function Count() As Long
    Count = sp + 1
End Function

'Aggiunge un elemento in testa allo stack
Public Sub Push(iType As eVarType, vValue As Variant)
    sp = sp + 1
    
    If sp > smax Then
        smax = smax + C_START_EL
        ReDim Preserve aStack(0 To smax) As tStackEl
    End If
    
    With aStack(sp)
        .iType = iType
        .vValue = vValue
    End With
End Sub

'Toglie l'elemento in testa allo stack
Public Sub Pop(iType As eVarType, vValue As Variant)
    With aStack(sp)
        iType = .iType
        vValue = .vValue
    End With
    sp = sp - 1
End Sub

'Setta il valore di un certo elemento
Public Sub SetElement(ByVal lIdx As Long, _
                iType As eVarType, vValue As Variant)

    If lIdx < 0 Or lIdx > sp Then
        Stop
    End If

    With aStack(lIdx)
        .iType = iType
        .vValue = vValue
    End With
End Sub

'Prende il valore di un certo elemento
Public Sub GetElement(ByVal lIdx As Long, _
                iType As eVarType, vValue As Variant)

    If lIdx < 0 Or lIdx > sp Then
        Stop
    End If
    
    With aStack(lIdx)
        iType = .iType
        vValue = .vValue
    End With
End Sub

'Prende l'elemento in testa
Public Sub GetTopElement(iType As eVarType, vValue As Variant)
    GetElement sp, iType, vValue
End Sub

Public Sub Output(lstBox As ListBox)
    lstBox.Clear
    
    Dim i As Long
    For i = 0 To sp
        Log Format(i, "00000") & " Type:" & CStr(aStack(i).iType) & ", Val:" & CStr(aStack(i).vValue), lstBox
    Next
End Sub
