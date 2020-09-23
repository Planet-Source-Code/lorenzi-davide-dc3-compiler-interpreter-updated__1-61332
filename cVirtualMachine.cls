VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVirtualMachine"
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
'Virtual Machine per l'esecuzione del codice generato
'
'================================================================================
Option Explicit

Private moByteCode As cByteCode
Private moStackVar As cStack
Private moStackAddr As cStack
Private moRuntimeError As cRunTimeError

Private mlPC As Long 'Program Counter
Private mlBP As Long 'Base Pointer (serve per il reperimento/salvataggio su stack)

'Riga corrente all'interno del sorgente
Private mlCurRow As Long

Private mTxtOut As TextBox

'Get the program counter index
Public Function GetPC() As Long
    GetPC = mlPC
End Function

'Get the stack base pointer
Public Function GetBP() As Long
    GetBP = mlBP
End Function

'Get the address stack
Public Function GetStackAddr() As cStack
    Set GetStackAddr = moStackAddr
End Function

'Get the var stack
Public Function GetStackVar() As cStack
    Set GetStackVar = moStackVar
End Function

Private Sub InitVM(oByteCode As cByteCode, txtOut As TextBox)
    Set moByteCode = oByteCode
    
    Set moStackVar = New cStack
    Set moStackAddr = New cStack
    Set moRuntimeError = New cRunTimeError
    mlPC = 0
    mlBP = 0
    mlCurRow = 1
    Randomize
    
    Set mTxtOut = txtOut
    mTxtOut.Text = ""
End Sub

'Inizializza la Virtual Machine
Public Sub Execute(oByteCode As cByteCode, txtOut As TextBox)
    'Inizializza la VM
    InitVM oByteCode, txtOut
    
    mTxtOut.Locked = False
    
    'Esegue il programma
    While ExecuteNextIstro()
    Wend
    
    'Programma terminato
    mTxtOut.Locked = True
End Sub

'Inizializza la Virtual Machine ed esegue la prima istruzione
Public Sub StepExecute(oByteCode As cByteCode, txtOut As TextBox)
    'Inizializza la VM
    InitVM oByteCode, txtOut
End Sub

'Esegue la prossima istruzione e ne rende il codice
Public Function ExecuteNextIstro() As Boolean
    Dim iIstro As eByteCodeIstro
    iIstro = moByteCode.GetData(mlPC)
    
    Select Case iIstro
        Case eByteCodeIstro.ebci_IConst:    doIConst
        Case eByteCodeIstro.ebci_IFrame:    doIFrame
        Case eByteCodeIstro.ebci_IFCall:    doIFCall
        Case eByteCodeIstro.ebci_IPCall:    doIPCall
        Case eByteCodeIstro.ebci_INData:    doINData
        Case eByteCodeIstro.ebci_IStore:    doIStore
        Case eByteCodeIstro.ebci_IFetch:    doIFetch
        Case eByteCodeIstro.ebci_IChop:     doIChop
        Case eByteCodeIstro.ebci_IReturn:   doIReturn
        Case eByteCodeIstro.ebci_IJump:     doIJump
        Case eByteCodeIstro.ebci_IEnd
            'Program is terminated
            ExecuteNextIstro = False
            Exit Function
        Case eByteCodeIstro.ebci_IBinOp:    doIBinOp
        Case eByteCodeIstro.ebci_IUnOp:     doIUnOp
        Case eByteCodeIstro.ebci_IRow:      doIRow
        Case eByteCodeIstro.ebci_IPopJF:    doIPopJF
        Case eByteCodeIstro.ebci_INop:
            mlPC = mlPC + 1
        Case Else
            Err.Raise vbObjectError + 1, "cVirtualMachine.Execute", "Unknown Instruction"
    End Select
    
    ExecuteNextIstro = True
End Function

'For debug purpose only
Private Sub doIRow()
    mlCurRow = moByteCode.GetData(mlPC + 1)
    mlPC = mlPC + 2
End Sub

'Esegue l'istruzione const
Private Sub doIConst()
    'Legge il tipo di variabile
    Dim iVarType As eVarType
    iVarType = moByteCode.GetData(mlPC + 1)
    
    'Legge il suo valore
    Dim vVarVal As Variant
    vVarVal = moByteCode.GetData(mlPC + 2)
    
    'Aggiunge allo stack
    moStackVar.Push iVarType, vVarVal
    
    'Avanza il program counter
    mlPC = mlPC + 3
End Sub

Private Sub doIFrame()
    'Salva il BP attuale
    moStackAddr.Push evt_long, mlBP
    
    'Salva il nuovo BP per le var. locali alla funzione
    moStackAddr.Push evt_long, moStackVar.Count
    
    'Avanza
    mlPC = mlPC + 1
End Sub

Private Sub doIFCall()
    'Legge il nuovo BP (puntatore alle var. locali) salvato dalla IFRAME
    moStackAddr.GetTopElement 0, mlBP
    
    'Salva il Program Counter alla prox istruzione
    moStackAddr.Push evt_long, mlPC + 3

    'Salta!
    mlPC = moByteCode.GetData(mlPC + 1)

End Sub

Private Sub doINData()
    Dim n As Long
    n = moByteCode.GetData(mlPC + 1)

    Dim i As Long
    For i = 1 To n
        moStackVar.Push evt_null, 0
    Next
    
    mlPC = mlPC + 2
End Sub

Private Sub doIStore()
    Dim lOffset As Long
    Dim bLocal As Boolean
    Dim iVarType As eVarType
    Dim vVarVal As Variant
    
    lOffset = moByteCode.GetData(mlPC + 1)
    bLocal = moByteCode.GetData(mlPC + 2)
    
    'Se locale
    If bLocal Then lOffset = mlBP + lOffset
    
    'Prende l'elemento in testa
    moStackVar.Pop iVarType, vVarVal

    'E lo salva
    moStackVar.SetElement lOffset, iVarType, vVarVal

    mlPC = mlPC + 3
End Sub

Private Sub doIFetch()
    Dim lOffset As Long
    Dim bLocal As Boolean
    Dim iVarType As eVarType
    Dim vVarVal As Variant
    
    lOffset = moByteCode.GetData(mlPC + 1)
    bLocal = moByteCode.GetData(mlPC + 2)
    
    'Se locale
    If bLocal Then lOffset = mlBP + lOffset
    
    'Prende l'elemento in testa
    moStackVar.GetElement lOffset, iVarType, vVarVal

    'E lo salva
    moStackVar.Push iVarType, vVarVal

    mlPC = mlPC + 3
End Sub

'Toglie l'elemento in testa allo stack
Private Sub doIChop()
    moStackVar.Pop 0, 0
    
    mlPC = mlPC + 1
End Sub

Private Sub doIReturn()
    'Legge la posizione alla quale tornare
    moStackAddr.Pop 0, mlPC
    
    'Ignora questa info che rappresenta il BP attuale
    moStackAddr.Pop 0, mlBP
    'Adesso deve rimuovere fino ad arrivare alla pos - 1
    While moStackVar.Count() > mlBP + 1
        moStackVar.Pop 0, 0
    Wend
    
    'Ripristina il Base Pointer originale
    moStackAddr.Pop 0, mlBP
    
End Sub

'Salta ad un certo indirizzo
Private Sub doIJump()
    Dim lAddr As Long
    lAddr = moByteCode.GetData(mlPC + 1)
    mlPC = lAddr
End Sub

'Salta ad un certo indirizzo
Private Sub doIPopJF()
    Dim lAddr As Long
    lAddr = moByteCode.GetData(mlPC + 1)
    
    
    Dim iType As eVarType
    Dim vVarVal As Variant
    moStackVar.Pop iType, vVarVal
    
    If iType = evt_bool Then
        If vVarVal = False Then
            mlPC = lAddr
            Exit Sub
        End If
    Else
        moRuntimeError.RaiseError ere_TypeMismatch, mlCurRow
    End If

    mlPC = mlPC + 2
End Sub


Private Sub doIUnOp()
    On Error GoTo Errore
    
    Dim iop As eBinOpCodes
    iop = moByteCode.GetData(mlPC + 1)
    
    Dim vVal1 As Variant
    Dim iType1 As eVarType
    Dim vRes As Variant
    
    moStackVar.Pop iType1, vVal1
    
    Select Case iop
        Case eUnOpCodes.euoc_plus
            vRes = vVal1
        Case eUnOpCodes.euoc_minus
            vRes = -vVal1
        Case eUnOpCodes.euoc_not
            vRes = Not vVal1
    End Select
    
    moStackVar.Push iType1, vRes
    mlPC = mlPC + 2
    
    Exit Sub

Errore:
    moRuntimeError.RaiseError ere_TypeMismatch, mlCurRow

End Sub

'Esegue un'operazione tra due valori in testa allo stack e pusha il risultato
Private Sub doIBinOp()
    On Error GoTo Errore
    
    Dim iop As eBinOpCodes
    iop = moByteCode.GetData(mlPC + 1)
    
    Dim vVal1 As Variant, vVal2 As Variant
    Dim iType1 As eVarType, iType2 As eVarType
    Dim vRes As Variant
    
    moStackVar.Pop iType1, vVal1
    moStackVar.Pop iType2, vVal2
    
    'Verifica corrispondenza tipi
    If iType1 <> iType2 Then
        moRuntimeError.RaiseError ere_TypeMismatch, mlCurRow
    End If
    
    Select Case iop
        Case eBinOpCodes.eboc_amp '&
            vRes = vVal2 & vVal1
        Case eBinOpCodes.eboc_and 'and
            vRes = vVal2 And vVal1
        Case eBinOpCodes.eboc_be '>=
            vRes = vVal2 >= vVal1
            iType1 = evt_bool
        Case eBinOpCodes.eboc_bt '>
            vRes = vVal2 > vVal1
            iType1 = evt_bool
        Case eBinOpCodes.eboc_caret '^
            vRes = vVal2 ^ vVal1
        Case eBinOpCodes.eboc_div '/
            vRes = vVal2 / vVal1
        Case eBinOpCodes.eboc_eq '=
            vRes = (vVal2 = vVal1)
            iType1 = evt_bool
        Case eBinOpCodes.eboc_eqv 'eqv
            vRes = vVal2 Eqv vVal1
        Case eBinOpCodes.eboc_imp 'imp
            vRes = vVal2 Imp vVal1
        Case eBinOpCodes.eboc_intdiv '\
            vRes = vVal2 \ vVal1
        Case eBinOpCodes.eboc_le '<=
            vRes = vVal2 <= vVal1
            iType1 = evt_bool
        Case eBinOpCodes.eboc_lt '<
            vRes = vVal2 < vVal1
            iType1 = evt_bool
        Case eBinOpCodes.eboc_minus '-
            vRes = vVal2 - vVal1
        Case eBinOpCodes.eboc_mod 'mod
            vRes = vVal2 Mod vVal1
        Case eBinOpCodes.eboc_mul '*
            vRes = vVal2 * vVal1
        Case eBinOpCodes.eboc_neq '<>
            vRes = (vVal2 <> vVal1)
            iType1 = evt_bool
        Case eBinOpCodes.eboc_or 'or
            vRes = vVal2 Or vVal1
        Case eBinOpCodes.eboc_plus '+
            vRes = vVal2 + vVal1
        Case eBinOpCodes.eboc_xor 'xor
            vRes = vVal2 Xor vVal1
    End Select
    
    moStackVar.Push iType1, vRes
    
    mlPC = mlPC + 2
    
    Exit Sub

Errore:
    moRuntimeError.RaiseError ere_TypeMismatch, mlCurRow

End Sub

Private Sub doIPCall()
    Dim iCode As ePFunc
    iCode = moByteCode.GetData(mlPC + 1)
    
    'Legge il nuovo puntatore alle var. locali salvato dalla IFRAME
    moStackAddr.GetTopElement 0, mlBP
    
    Select Case iCode
        Case ePFunc.eps_Print:      pcall_doPrint
        Case ePFunc.eps_MsgBox:     pcall_doMsgBox
        Case ePFunc.epf_CLng:       pcall_doCLng
        Case ePFunc.epf_CStr:       pcall_doCStr
        Case ePFunc.epf_InputBox:   pcall_doInputBox
        Case ePFunc.epf_Rnd:        pcall_doRnd
    End Select
    
    'Ignora questa info che rappresenta il BP attuale
    moStackAddr.Pop 0, mlBP
    'Adesso deve rimuovere fino ad arrivare alla pos - 1
    While moStackVar.Count() > mlBP + 1
        moStackVar.Pop 0, 0
    Wend
    
    'Ripristina il Base Pointer
    moStackAddr.Pop 0, mlBP
    
    mlPC = mlPC + 2
End Sub

Private Sub pcall_doPrint()
    Dim iVarType As eVarType
    Dim vVarVal As Variant
    
    moStackVar.GetTopElement iVarType, vVarVal
    mTxtOut.SelText = CStr(vVarVal) & vbCrLf
End Sub

Private Sub pcall_doMsgBox()
    Dim iVarType As eVarType
    Dim vVarVal As Variant
    
    moStackVar.GetTopElement iVarType, vVarVal
    MsgBox CStr(vVarVal)
End Sub

Private Sub pcall_doCLng()
    Dim iVarType As eVarType
    Dim vVarVal As Variant
    
    moStackVar.GetTopElement iVarType, vVarVal
    
    'E lo salva nello stack alla pos. di ritorno della funzione
    moStackVar.SetElement mlBP + 0, eVarType.evt_long, CLng(vVarVal)
End Sub

'lo salva nello stack alla pos. di ritorno della funzione
Private Sub pcall_doRnd()
    moStackVar.SetElement mlBP + 0, eVarType.evt_long, Rnd
End Sub


Private Sub pcall_doCStr()
    Dim iVarType As eVarType
    Dim vVarVal As Variant
    
    moStackVar.GetTopElement iVarType, vVarVal
    
    'E lo salva nello stack alla pos. di ritorno della funzione
    moStackVar.SetElement mlBP + 0, eVarType.evt_string, CStr(vVarVal)
End Sub

Private Sub pcall_doInputBox()
    Dim iVarType As eVarType
    Dim vVarVal As Variant
    
    moStackVar.GetTopElement iVarType, vVarVal
        
    vVarVal = InputBox(vVarVal)
    
    'E lo salva nello stack alla pos. di ritorno della funzione
    moStackVar.SetElement mlBP + 0, eVarType.evt_string, CStr(vVarVal)
End Sub

