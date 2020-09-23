Attribute VB_Name = "modMain"
'================================================================================
' Part of DC3 Compiler - Interpreter
' Author: Lorenzi Davide (http://www.hexagora.com)
' See the file 'license.txt' for informations
'================================================================================
Option Explicit


'//////////////////////////////////////////////////////////////////
'API Declarations
'
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_ALT = &H12

'//////////////////////////////////////////////////////////////////


Public Const C_ERR_COMP = vbObjectError + 1
Public Const C_ERR_RUNTIME = vbObjectError + 2

'Tipi di simboli per la SymbTable
Public Enum eSymbTableElType
    estet_null = 0
    estet_var
    estet_const
    estet_func
    estet_sub
End Enum

'Tipi di modificatori per le variabili
Public Enum eParmMod
    epm_NotSpecified = 0
    epm_ByVal
    epm_ByRef
End Enum

'Tipi di variabile validi
Public Enum eVarType
    evt_null = 0
    evt_long
    evt_string
    evt_bool
    evt_double
End Enum

'Tipo di modificatore d'accesso per funzioni e variabili
Public Enum eAccessModifier
    eam_notused = 0
    eam_public
    eam_private
    eam_reserved
End Enum

'Per la gestione dello stack
Public Type tStackEl
    iType As eVarType 'Tipo di variabile
    vValue As Variant 'Valore
End Type


Sub Main()
    InitDialogs
    moParser.LoadCompiledGrammar (App.Path & "\dc3.cgt")
    frmDavComp.Show
End Sub

Public Sub Log(ByVal s As String, lstOut As ListBox)
    lstOut.AddItem s
End Sub

'Legge un file di testo e lo rende nella stringa
Public Function FileToString(ByVal sNomeFile As String) As String
    On Error GoTo Errore
    Dim nf As Integer
    Dim sTmp As String, sMain As String
    Dim riga As String
    
    sTmp = ""
    sMain = ""
    
    nf = FreeFile
    Open sNomeFile For Input As #nf
    While Not EOF(nf)
        Line Input #nf, riga
        sTmp = sTmp & riga & vbCrLf
        'Cosi' velocizza il discorso nel caso di file molto lunghi
        If Len(sTmp) > 15000 Then
            sMain = sMain & sTmp
            sTmp = ""
        End If
    Wend
    Close #nf

    sMain = sMain & sTmp

    FileToString = sMain
    Exit Function

Errore:
    'Sparo fuori l'errore con una descrizione diversa
    On Error GoTo 0
    Err.Raise vbObjectError + 1, , Err.Description & " " & sNomeFile
End Function

'Da stringa a file
Public Sub StringToFile(ByVal vsPathFile As String, ByVal vsText As String)
    On Error GoTo Errore
    Dim iFile As Integer
    
    iFile = FreeFile
    Open vsPathFile For Output As #iFile
        Print #iFile, vsText
    Close #iFile
    Exit Sub

Errore:
    'Sparo fuori l'errore con una descrizione diversa
    On Error GoTo 0
    Err.Raise vbObjectError + 1, , Err.Description & " " & vsPathFile
End Sub

Public Function isKeyDown(KeyCode As Integer) As Boolean
    Dim i1 As Long
    
    'Adesso devo testare solo il bit piu' significativo per lo stato
    i1 = GetAsyncKeyState(KeyCode)
    i1 = i1 And &H8000
    
    isKeyDown = (i1 <> 0)
End Function

'Rende vero se uno dei due shift e' premuto
Public Function isShiftDown() As Boolean
    isShiftDown = isKeyDown(VK_SHIFT)
End Function

'Rende vero se uno dei due control e' premuto
Public Function isControlDown() As Boolean
    isControlDown = isKeyDown(VK_CONTROL)
End Function

'Rende vero se e' premuto ALT
Public Function isAltDown() As Boolean
    isAltDown = isKeyDown(VK_ALT)
End Function

'Visualizza una msgbox di errore
Public Sub ShowError(sError As String)
    MsgBox sError, vbCritical, "Attention"
End Sub

'Visualizza una msgbox di info
Public Sub ShowInfo(sInfo As String)
    MsgBox sInfo, vbInformation, "Attention"
End Sub

'Rende vero se l'utente ha risposto si
Public Function ShowYesNo(ByVal sQuestion As String, Optional ByVal sTitle As String = "") As Boolean
    If Trim(sTitle) = "" Then sTitle = "Attention"
    ShowYesNo = (MsgBox(sQuestion, vbQuestion Or vbYesNo, sTitle) = vbYes)
End Function

'In debug mode doesn't activate some features
Public Function IsDebug() As Boolean
    IsDebug = (InStr(1, Command$, "DEBUG") > 0)
End Function
