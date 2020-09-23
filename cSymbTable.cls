VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSymbTable"
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
' Classe per la gestione della Symbol Table per il compilatore
'
'================================================================================
Option Explicit

Private oColl As Collection

Private Sub Class_Initialize()
    Set oColl = New Collection
End Sub

'Rende il numero di elementi interni
Public Function Count() As Integer
    Count = oColl.Count
End Function

'Conta tutte gli elementi di un certo tipo definiti nella symbol table
Public Function CountType(iType As eSymbTableElType) As Integer
    Dim i As Integer
    Dim c As Integer
    Dim oEl As cSymbTableEl
    
    c = 0
    For i = 1 To oColl.Count
        Set oEl = oColl(i)
        
        If oEl.iType = iType Then
            c = c + 1
        End If
    Next
    CountType = c
End Function

'Rende l'indice di una certa var. all'interno della SymbTable ignorando
'gli altri simboli. Questo serve xche' nello stack sono presenti solo le var.
Public Function GetVarIndex(sSymbName As String) As Integer
    Dim i As Integer
    Dim oSE As cSymbTableEl
    
    Dim lIdx As Long
    lIdx = -1
    
    For i = 1 To oColl.Count
        Set oSE = oColl.Item(i)
        If oSE.iType = estet_const Or oSE.iType = estet_var Then
            lIdx = lIdx + 1
            If LCase(oSE.sName) = LCase(sSymbName) Then
                Exit For
            End If
        End If
    Next
    
    GetVarIndex = lIdx
End Function


'
'Aggiunge un nuovo simbolo alla tabella
'Rende l'indice della Symbol Table in cui e' stato aggiunto
'
Public Function Add(oSymbEl As cSymbTableEl) As cSymbTableEl
    If Not oSymbEl Is Nothing Then
        If Search(oSymbEl.sName) Is Nothing And (oSymbEl.sName <> "") Then
            oColl.Add oSymbEl
            Set Add = oSymbEl
            Exit Function
        End If
    End If
    Set Add = Nothing
End Function

'Aggiunge una variabile
Public Function AddVar(sName As String, _
                        iType As eSymbTableElType, _
                        iAccMod As eAccessModifier) As cSymbTableEl
    
    Dim oSymbEl  As New cSymbTableEl
    
    'Crea la classe
    With oSymbEl
        .sName = sName
        .iType = iType
        .iAccMod = iAccMod
    End With
    
    'Aggiunge
    Set AddVar = Add(oSymbEl)
End Function

'Aggiunge una funzione
Public Function AddFunc(sName As String, _
                iType As eSymbTableElType, _
                iAccMod As eAccessModifier) As cSymbTableEl
    
    Dim oSymbEl As New cSymbTableEl
    
    'Crea la classe
    With oSymbEl
        .sName = sName
        .iType = iType
        .iAccMod = iAccMod
    End With
    Set oSymbEl.oFuncInfos = New cFuncInfos
    
    'Aggiunge
    Set AddFunc = Add(oSymbEl)
End Function


'
'Rende l'indice di esistenza di un certo elemento all'interno della tabella
'
Public Function Search(sName As String, Optional ByVal bMatchCase As Boolean = False) As cSymbTableEl
    Dim i As Integer
    Dim oSE As cSymbTableEl
    
    For i = 1 To oColl.Count
        Set oSE = oColl.Item(i)
        If bMatchCase Then
            If oSE.sName = sName Then
                Set Search = oSE
                Exit Function
            End If
        Else
            If LCase(oSE.sName) = LCase(sName) Then
                Set Search = oSE
                Exit Function
            End If
        End If
    Next
    
    Set Search = Nothing
End Function

Public Sub Output(ByVal iIndent As Integer, lstOut As ListBox)
    Dim i As Integer
    Dim o As cSymbTableEl
    Dim sLog As String
    
    If iIndent = 0 Then lstOut.Clear
    
    For i = 1 To oColl.Count
        Set o = oColl(i)
        
        sLog = "+" & String(iIndent, "-") & " " & Format(i - 1, "00000") & " "
        
        Select Case o.iAccMod
            Case eAccessModifier.eam_private: sLog = sLog & "[PRIVATE] "
            Case eAccessModifier.eam_public: sLog = sLog & "[PUBLIC] "
            Case eAccessModifier.eam_reserved: sLog = sLog & "[RESERVED] "
        End Select
        
        Select Case o.iType
            Case estet_const: sLog = sLog & "Const"
            Case estet_var: sLog = sLog & "Var"
            Case estet_func: sLog = sLog & "Func"
            Case estet_sub: sLog = sLog & "Sub"
        End Select
        sLog = sLog & " " & o.sName
        If Not o.oFuncInfos Is Nothing Then
            sLog = sLog & " addr:" & o.oFuncInfos.lInstrPointer
        End If
        
        Log sLog, lstOut
        
        If Not o.oFuncInfos Is Nothing Then
            o.oFuncInfos.oSymbTable.Output iIndent + 5, lstOut
        End If
    Next
End Sub
