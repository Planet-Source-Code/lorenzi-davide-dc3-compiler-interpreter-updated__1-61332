VERSION 5.00
Begin VB.Form frmDavComp 
   Caption         =   "Dav Compiler 3"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   7650
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstLog 
      Height          =   1035
      Left            =   30
      TabIndex        =   1
      Top             =   4440
      Width           =   7545
   End
   Begin VB.TextBox txtProgram 
      Height          =   4335
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   7545
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "Debug"
      Begin VB.Menu mnuRun 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmDavComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================================
' Part of DC3 Compiler - Interpreter
' Author: Lorenzi Davide (http://www.hexagora.com)
' See the file 'license.txt' for informations
'================================================================================
Option Explicit

'Used for TAB management in Vb Textbox
Private cHook As cHook
Implements WinSubHook2.iHook
Private mbHasFocus As Boolean

Private msCurFile As String
Private Const C_DEF_FILE = "NewFile"
Private bChanged As Boolean


Private Sub Form_Unload(Cancel As Integer)
    Dim r As VbMsgBoxResult
    If bChanged Then
        r = MsgBox("Current file has changed, save it now?", vbYesNoCancel)
        If r = vbYes Then
            mnuSave_Click
            'See if I've really saved the file
            If bChanged Then Exit Sub
        ElseIf r = vbCancel Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    cHook.UnHook
End Sub

Private Sub iHook_Proc(ByVal bBefore As Boolean, bHandled As Boolean, _
    lReturn As Long, nCode As WinSubHook2.eHookCode, wParam As Long, lParam As Long)

    If nCode = HC_ACTION Then
        If mbHasFocus Then
            If (lParam And &H80000000) Or (lParam And &H40000000) Then
                Exit Sub
            End If
            
            Select Case wParam
                Case 9: 'VK_TAB
                    If isShiftDown() Then
                        txtProgram.SelText = Replace(txtProgram.SelText, vbCrLf & vbTab, vbCrLf)
                    Else
                        txtProgram.SelText = vbTab & Replace(txtProgram.SelText, vbCrLf, vbCrLf & vbTab)
                    End If
                    
                    lReturn = 1 'Cosi' la fermo e non spedisce 2 tab
                    bHandled = True
                    Exit Sub
            End Select
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.ScaleMode = vbPixels
    
    'Attach the hook so I can take the TAB character
    '(only if the program is compiled in .exe)
    Set cHook = New cHook
    If Not IsDebug() Then
        cHook.Hook Me, WH_KEYBOARD
    End If
    
    'Setup the name of the file
    msCurFile = C_DEF_FILE
    
    mnuNew_Click

    OpenFile App.Path & "\start.txt"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lstLog.Left = 0
    lstLog.Top = Me.ScaleHeight - lstLog.Height
    lstLog.Width = Me.ScaleWidth - 1
    
    txtProgram.Left = 0
    txtProgram.Width = Me.ScaleWidth - 1
    txtProgram.Height = Me.ScaleHeight - txtProgram.Top - lstLog.Height
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Copyright 2004-2005" & vbCrLf & "by Lorenzi Davide" & vbCrLf & "http://www.hexagora.com"
End Sub

Private Sub mnuExit_Click()
    Unload frmDebug
    Unload Me
End Sub

Private Sub mnuNew_Click()

    Dim r As VbMsgBoxResult
    If bChanged Then
        r = MsgBox("Current file has changed, save it now?", vbYesNoCancel)
        If r = vbYes Then
            mnuSave_Click
            'See if I've really saved the file
            If bChanged Then Exit Sub
        ElseIf r = vbCancel Then
            Exit Sub
        End If
    End If

    txtProgram.Text = ""
    bChanged = False
    msCurFile = C_DEF_FILE
    Me.Caption = msCurFile
End Sub

Private Sub mnuOpen_Click()
    Dim r As VbMsgBoxResult
    If bChanged Then
        r = MsgBox("Current file has changed, save it now?", vbYesNoCancel)
        If r = vbYes Then
            mnuSave_Click
            'See if I've really saved the file
            If bChanged Then Exit Sub
        ElseIf r = vbCancel Then
            Exit Sub
        End If
    End If
    
    OpenFileWithDialog
End Sub

Private Sub mnuRun_Click()
    Dim oProgram As New cProgram
    If oProgram.Compile(txtProgram.Text, lstLog) Then
        Log "Compilation was ok!", lstLog
        
        With frmDebug
            .Show
            Set .moByteCode = oProgram.GetByteCode()
            Set .moSymbTable = oProgram.GetSymbTable()
            .Execute
        End With
    End If
End Sub

Private Sub mnuSave_Click()
    On Error GoTo Errore

    If msCurFile = C_DEF_FILE Then
        mnuSaveAs_Click
    Else
        'Salva
        StringToFile msCurFile, txtProgram.Text
        bChanged = False
    End If
    
    Exit Sub
Errore:
    ShowError Err.Description
End Sub

Private Sub mnuSaveAs_Click()
    On Error GoTo Errore
    
    Dim sFile As String
    Dim sFilter As String
    Dim iFilterIndex As Integer
    
    'Filtro
    sFilter = "DC3 File (*.vb)" + Chr$(0) + "*.vb" + Chr$(0) & _
              "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    sFile = ShowSave(Me, sFilter, "Save File...", OFN_OVERWRITEPROMPT, "", "", iFilterIndex)
    If sFile <> "" Then
        'Prima assegna l'estensione giusta in base al filtro selezionato
        Dim sExt As String
        Select Case iFilterIndex
            Case 1: sExt = ".vb"
            Case 2: sExt = ""
        End Select
        
        If Right(LCase(sFile), Len(sExt)) <> sExt Then
            sFile = sFile & sExt
        End If
        
        'Salva
        StringToFile sFile, txtProgram.Text
        Me.Caption = sFile
        msCurFile = sFile
        bChanged = False
    End If
    
    Exit Sub
Errore:
    ShowError Err.Description
End Sub

Private Sub mnuSelectAll_Click()
    txtProgram.SelStart = 0
    txtProgram.SelLength = Len(txtProgram.Text)
End Sub

Private Sub txtProgram_Change()
    bChanged = True
End Sub

Private Sub txtProgram_GotFocus()
    mbHasFocus = True
End Sub

Private Sub txtProgram_LostFocus()
    mbHasFocus = False
End Sub

Private Sub OpenFileWithDialog()
    Dim s As String
    Dim sFilter As String
    
    'Filtro
    sFilter = "DC3 Files (*.vb)" + Chr$(0) + "*.vb" + Chr$(0) + _
            "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    s = ShowOpen(frmDavComp, sFilter, "Open File...", _
            OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or _
            OFN_EXPLORER, "")
            
    
    Dim v As Variant
    Dim sDir As String
    Dim sFileName As String
    Dim i As Integer
    
    If Len(s) > 0 Then
        v = Split(s, vbNullChar)
        sDir = v(0)
        
        sFileName = sDir
        OpenFile sFileName
    End If
End Sub

Private Sub OpenFile(sFile As String)
    On Error GoTo Errore
    
    frmDavComp.txtProgram.Text = FileToString(sFile)
    Me.Caption = sFile
    msCurFile = sFile
    bChanged = False
    
    Exit Sub
Errore:
    ShowError Err.Description
    
End Sub


