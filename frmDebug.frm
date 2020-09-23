VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Execution Window"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRun 
      Caption         =   "Re-Run"
      Height          =   345
      Left            =   30
      TabIndex        =   10
      Top             =   30
      Width           =   945
   End
   Begin VB.ListBox lstVar 
      Height          =   1815
      Left            =   5850
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3450
      Width           =   2445
   End
   Begin VB.ListBox lstAddr 
      Height          =   2400
      Left            =   5850
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   2445
   End
   Begin VB.ListBox lstSymbolTable 
      Height          =   645
      Left            =   2850
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4620
      Width           =   2955
   End
   Begin VB.ListBox lstByteCode 
      Height          =   4545
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   2850
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   720
      Width           =   2955
   End
   Begin VB.Label lblSymbTable 
      AutoSize        =   -1  'True
      Caption         =   "Symbol Table"
      Height          =   195
      Left            =   2910
      TabIndex        =   9
      Top             =   4320
      Width           =   960
   End
   Begin VB.Label lblStackVar 
      AutoSize        =   -1  'True
      Caption         =   "Stack Variables"
      Height          =   195
      Left            =   5880
      TabIndex        =   8
      Top             =   3180
      Width           =   1110
   End
   Begin VB.Label lblStackAddr 
      AutoSize        =   -1  'True
      Caption         =   "Stack Addresses"
      Height          =   195
      Left            =   5850
      TabIndex        =   7
      Top             =   510
      Width           =   1200
   End
   Begin VB.Label lblProgramOut 
      AutoSize        =   -1  'True
      Caption         =   "Program Output"
      Height          =   195
      Left            =   2850
      TabIndex        =   3
      Top             =   480
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Byte Code"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmDebug"
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

Public moByteCode As cByteCode
Public moSymbTable As cSymbTable

'Execute the program!
Public Sub Execute()
    On Error GoTo Errore

    moSymbTable.Output 0, lstSymbolTable
    moByteCode.Output lstByteCode
    
    Dim oVM As New cVirtualMachine
    oVM.Execute moByteCode, txtOutput
    
    oVM.GetStackAddr().Output lstAddr
    oVM.GetStackVar().Output lstVar
    
    Exit Sub
Errore:
    ShowError "Runtime Error: " & Err.Description
End Sub

Private Sub btnRun_Click()
    Execute
End Sub

Private Sub Form_Load()
    Me.ScaleMode = vbPixels
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lstByteCode.Left = 0
    lstByteCode.Height = Me.ScaleHeight - lstByteCode.Top
    
    lstAddr.Left = Me.ScaleWidth - lstAddr.Width - 1
    lblStackAddr.Left = lstAddr.Left
    lstVar.Left = lstAddr.Left
    lstVar.Height = Me.ScaleHeight - lstVar.Top
    lblStackVar.Left = lstVar.Left
    
    txtOutput.Left = lstByteCode.Left + lstByteCode.Width + 1
    txtOutput.Width = lstAddr.Left - txtOutput.Left
    lblProgramOut.Left = txtOutput.Left
    
    lstSymbolTable.Left = txtOutput.Left
    lstSymbolTable.Width = txtOutput.Width
    lstSymbolTable.Height = Me.ScaleHeight - lstSymbolTable.Top
    lblSymbTable.Left = lstSymbolTable.Left
    
End Sub

