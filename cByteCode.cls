VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cByteCode"
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
'Classe per la gestione del ByteCode per la virtual machine
'
'================================================================================

Option Explicit

Private Const C_START_EL = 100

Public Enum eByteCodeIstro
    ebci_INop    '..............................................
    ebci_IConst  'type,val......................................
    ebci_IFrame  '..............................................
    ebci_IFCall  'addr,npars....................................
    ebci_IPCall  'id pfunc......................................
    ebci_INData  'n. var to alloc...............................
    ebci_IStore  'offset, bLocal................................
    ebci_IFetch  'offset, bLocal................................
    ebci_IChop   '..............................................
    ebci_IReturn '..............................................
    ebci_IJump   'addr..........................................
    ebci_IEnd    '..............................................
    ebci_IBinOp  'opcode........................................
    ebci_IUnOp   'opcode........................................
    ebci_IRow    'row (only for debug purpose)..................
    ebci_IPopJF  'addr..........................................
End Enum

Public Enum eBinOpCodes
    eboc_plus = 0       '+
    eboc_minus          '-
    eboc_mod            'mod
    eboc_intdiv         '\
    eboc_mul            '*
    eboc_div            '/
    eboc_caret          '^
    eboc_imp            'imp
    eboc_eqv            'eqv
    eboc_xor            'xor
    eboc_or             'or
    eboc_and            'and
    eboc_be             '>=
    eboc_le             '<=
    eboc_bt             '>
    eboc_lt             '<
    eboc_neq            '<>
    eboc_eq             '=
    eboc_amp            '&
End Enum

Public Enum eUnOpCodes
    euoc_plus = 0       '+
    euoc_minus          '-
    euoc_not            'not
End Enum

Private aByteCode() As Variant 'E' un array di variant
Private ip As Long 'Puntatore al prossimo elemento
Private ip_max As Long 'Massimo numero di elementi
Private aIstroWeight As Variant

'Riga e colonna corrispondenti all'interno del programma
Private mlCurRow As Long

Private Sub Class_Initialize()
    ip_max = C_START_EL
    ReDim aByteCode(0 To ip_max) As Variant
    ip = -1
    mlCurRow = 1
    aIstroWeight = Array(1, 3, 1, 3, 2, 2, 3, 3, 1, 1, 2, 1, 2, 2, 2, 2)
End Sub

Private Sub Class_Terminate()
    Erase aByteCode
End Sub

'Rende il numero di istruzioni presenti
Public Function Count() As Long
    Count = ip + 1
End Function

'Rende l'istruzione presente in una certa posizione
Public Function GetData(idx As Long) As Variant
    If idx > -1 And idx < ip_max Then
        GetData = aByteCode(idx)
    Else
        Err.Raise vbObjectError + 1, "cByteCode.GetData", "Instruction Not Present"
    End If
End Function

'Setta il dato ad un certo indirizzo
Public Sub SetData(vData As Variant, idx As Long)
    If idx > -1 And idx < ip_max Then
        aByteCode(idx) = vData
    Else
        Err.Raise vbObjectError + 1, "cByteCode.SetData", "Instruction Not Present"
    End If
End Sub

'Setta un dato ad un certo indice ma se trova delle istruzioni per il debug avanza il contatore
Public Sub SetIstroParam(vData As Variant, ByVal idx As Long, ByVal iOffsetParam As Integer)
    If idx > -1 And idx < ip_max Then
        'Ignora le istruzioni di debug
        While aByteCode(idx) = eByteCodeIstro.ebci_IRow
            idx = idx + aIstroWeight(aByteCode(idx))
        Wend
        aByteCode(idx + iOffsetParam) = vData
    Else
        Err.Raise vbObjectError + 1, "cByteCode.SetDataWithDBInfo", "Instruction Not Present"
    End If
End Sub

'Aggiunge un dato al ByteCode
Private Sub AddData(vData As Variant)
    ip = ip + 1
    If ip > ip_max Then
        ip_max = ip_max + C_START_EL
        ReDim Preserve aByteCode(0 To ip_max) As Variant
    End If
    aByteCode(ip) = vData
End Sub

'Aggiunge un'instruzione al programma
Private Sub AddIstro(iIstro As eByteCodeIstro, lRow As Long)
    If mlCurRow <> lRow Then
        mlCurRow = lRow
        'Aggiunge l'istruzione IRow
        AddData eByteCodeIstro.ebci_IRow
        AddData lRow
    End If
    
    AddData iIstro
End Sub


'Mette un valore costante in testa allo stack
Public Sub Add_IConst(iType As eVarType, vValue As Variant, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IConst, lRow
    AddData iType
    AddData vValue
End Sub

'Salva i puntatori per una chiamata a funzione
Public Sub Add_IFrame( _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IFrame, lRow
End Sub

'Esegue una chiamata a funzione
Public Sub Add_IFCall(lAddr As Long, lNPars As Integer, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IFCall, lRow
    AddData lAddr
    AddData lNPars
End Sub

'Esegue una chiamata ad una funzione riservata
Public Sub Add_IPCall(iIdFunc As ePFunc, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IPCall, lRow
    AddData iIdFunc
End Sub

'Aggiunge lo spazio sullo stack per n variabili
Public Sub Add_INData(lNumAlloc As Long, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_INData, lRow
    AddData lNumAlloc
End Sub

'Prende l'elemento in testa allo stack e lo mette in una variabile
Public Sub Add_IStore(lOffset As Long, bLocal As Boolean, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IStore, lRow
    AddData lOffset
    AddData bLocal
End Sub

'Prende l'elemento dalle var. e lo mette in testa allo stack
Public Sub Add_IFetch(lOffset As Long, bLocal As Boolean, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IFetch, lRow
    AddData lOffset
    AddData bLocal
End Sub

'Toglie il valore in testa allo stack
Public Sub Add_IChop( _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IChop, lRow
End Sub

'Ripristina il pc e lo sp per il ritorno da una funzione
Public Sub Add_IReturn( _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IReturn, lRow
End Sub

'Sposta il pc a addr
Public Sub Add_IJump(lAddr As Long, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IJump, lRow
    AddData lAddr
End Sub

'Termina il programma
Public Sub Add_IEnd( _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IEnd, lRow
End Sub

'NOP, non ha riga, mi serve da marcatore
Public Sub Add_INOp()
    AddData eByteCodeIstro.ebci_INop
End Sub

'Esegue un'operazione tra due valori
Public Sub Add_IBinOp(iOpCode As eBinOpCodes, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IBinOp, lRow
    AddData iOpCode
End Sub

'Esegue un'operazione con un valore
Public Sub Add_IUnOp(iOpCode As eBinOpCodes, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IUnOp, lRow
    AddData iOpCode
End Sub

'Toglie l'ultimo elemento dallo stack e salta all'indirizzo se il valore e' vero
Public Sub Add_IPopJF(lAddr As Long, _
    ByVal lRow As Long)
    
    AddIstro eByteCodeIstro.ebci_IPopJF, lRow
    AddData lAddr
End Sub

Public Function OutputIndex(ByVal i As Long) As String
    Dim sLog As String
    sLog = Format(i, "00000:") & " "
    
    Select Case aByteCode(i)
        Case ebci_IChop
            sLog = sLog & "ICHOP"
        Case ebci_IConst
            sLog = sLog & "ICONST type:" & CStr(aByteCode(i + 1)) & " value:" & CStr(aByteCode(i + 2))
        Case ebci_IFCall
            sLog = sLog & "IFCALL addr:" & CStr(aByteCode(i + 1)) & " pars:" & CStr(aByteCode(i + 2))
        Case ebci_IFetch
            sLog = sLog & "IFETCH offset:" & CStr(aByteCode(i + 1)) & " local:" & IIf(aByteCode(i + 2), "1", "0")
        Case ebci_IFrame
            sLog = sLog & "IFRAME"
        Case ebci_IJump
            sLog = sLog & "IJUMP " & CStr(aByteCode(i + 1))
        Case ebci_INData
            sLog = sLog & "INDATA " & CStr(aByteCode(i + 1))
        Case ebci_IPCall
            sLog = sLog & "IPCALL id:" & CStr(aByteCode(i + 1))
        Case ebci_IReturn
            sLog = sLog & "IRETURN"
        Case ebci_IStore
            sLog = sLog & "ISTORE offset:" & CStr(aByteCode(i + 1)) & " local:" & IIf(aByteCode(i + 2), "1", "0")
        Case ebci_IEnd
            sLog = sLog & "IEND"
        Case ebci_IBinOp
            sLog = sLog & "IBINOP opcode:" & CStr(aByteCode(i + 1))
        Case ebci_IUnOp
            sLog = sLog & "IUNOP opcode:" & CStr(aByteCode(i + 1))
        Case ebci_IRow
            sLog = sLog & "IROW row:" & CStr(aByteCode(i + 1))
        Case ebci_IPopJF
            sLog = sLog & "IPOPJF addr:" & CStr(aByteCode(i + 1))
        Case ebci_INop
            sLog = sLog & "INOP"
    End Select

    OutputIndex = sLog
End Function

Public Sub Output(lstLog As ListBox)
    On Error GoTo Errore
    
    lstLog.Clear
    
    Dim i As Long
    Dim sLog As String
    
    i = 0
    While i <= ip
        Log OutputIndex(i), lstLog
        i = i + aIstroWeight(aByteCode(i))
    Wend
    
    Exit Sub
Errore:
    Err.Raise vbObjectError + 1, "cByteCode.Output", "Unknown Instruction"
End Sub
