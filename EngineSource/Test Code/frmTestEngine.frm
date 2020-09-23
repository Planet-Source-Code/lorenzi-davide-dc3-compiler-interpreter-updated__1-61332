VERSION 5.00
Begin VB.Form frmTestEngine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test The Grammar"
   ClientHeight    =   7350
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   8925
   Icon            =   "frmTestEngine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTrimReductions 
      Caption         =   "Trim Reductions"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Value           =   1  'Checked
      Width           =   2595
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parse Tree"
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   2700
      Width           =   8655
      Begin VB.TextBox txtParseTree 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "frmTestEngine.frx":0442
         Top             =   300
         Width           =   8415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GOLD Parser Input"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8655
      Begin VB.TextBox txtTestInput 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   780
         Width           =   7170
      End
      Begin VB.TextBox txtCGTFilePath 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   7155
      End
      Begin VB.Label Label1 
         Caption         =   "CGT File"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Test Input"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   6900
      Width           =   1215
   End
End
Attribute VB_Name = "frmTestEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'                 The GOLD Parser Freeware License Agreement
'                 ==========================================
'
'this software Is provided 'as-is', without any expressed or implied warranty.
'In no event will the authors be held liable for any damages arising from the
'use of this software.
'
'Permission is granted to anyone to use this software for any purpose. If you
'use this software in a product, an acknowledgment in the product documentation
'would be deeply appreciated but is not required.
'
'In the case of the GOLD Parser Engine source code, permission is granted to
'anyone to alter it and redistribute it freely, subject to the following
'restrictions:
'
'   1. The origin of this software must not be misrepresented; you must not
'      claim that you wrote the original software.
'
'   2. Altered source versions must be plainly marked as such, and must not
'      be misrepresented as being the original software.
'
'   3. This notice may not be removed or altered from any source distribution


Option Explicit


Private Sub AddToReport(ByVal Action As String, ByVal Description As String, ByVal Value As String, ByVal Index As String, ByVal LineNumber As Long)
     'This simple procedure is used to log the actions performed by the parser.
     'For now, it just prints the information to the Debug window.
     
     Debug.Print Action & ", " & Description & ", " & Value & ", " & Index & ", " & LineNumber
     
End Sub




Private Sub DoParse()
   'This procedure starts the GOLD Parser Engine and handles each of the
   'messages it returns. In all cases, the "AddToReport" procedure is called
   'the relevant information - exactly what happens in the GOLD Parser
   'Builder.
   
   'Once the parsing is complete and the text is accepted, this procedure calls
   'the DrawReductionTree procedure which creates an ASCII version of the parse
   'tree in the Builder Test Window.
      
   Dim Response As GPMessageConstants
   Dim Parser   As New GOLDParser
   Dim Done As Boolean                                    'Controls when we leave the loop
   
   Dim ReductionNumber As Integer                         'Just for information
   Dim n As Integer, Text As String
      
   'The Compiled Grammar Table file is loaded each time this procedure is called.
   'It is recommended that you put the LoadCompiledGrammar method call in another
   'procedure that is called when your program is starting.
      
   If Parser.LoadCompiledGrammar(txtCGTFilePath.Text) Then
       Parser.OpenTextString txtTestInput.Text
       Parser.TrimReductions = (chkTrimReductions.Value = vbChecked)
                 
       Done = False
       Do Until Done
           Response = Parser.Parse()
              
           Select Case Response
           Case gpMsgLexicalError
              AddToReport "Lexical Error", "Cannot recognize token", Parser.CurrentToken.Data, "", Parser.CurrentLineNumber
              txtParseTree.Text = "Line " & Parser.CurrentLineNumber & ": Lexical Error: Cannot recognize token: " & Parser.CurrentToken.Data
              Done = True
                 
           Case gpMsgSyntaxError
              Text = ""
              For n = 0 To Parser.TokenCount - 1
                  Text = Text & " " & Parser.Tokens(n).Name
              Next
              AddToReport "Syntax Error", "Expecting the following tokens", LTrim(Text), "", Parser.CurrentLineNumber
              txtParseTree.Text = "Line " & Parser.CurrentLineNumber & ": Syntax Error: Expecting the following tokens: " & LTrim(Text)
              Done = True
              
           Case gpMsgReduction
              ReductionNumber = ReductionNumber + 1
              Parser.CurrentReduction.Tag = ReductionNumber   'Mark the reduction
              AddToReport "Reduce", Parser.CurrentReduction.ParentRule.Text, ReductionNumber, Parser.CurrentReduction.ParentRule.TableIndex, Parser.CurrentLineNumber
                                  
           Case gpMsgAccept
              '=== Success!
              AddToReport "Accept", Parser.CurrentReduction.ParentRule.Text, "", Parser.CurrentReduction.ParentRule.TableIndex, Parser.CurrentLineNumber
              DrawReductionTree Parser.CurrentReduction
              Done = True
              
           Case gpMsgTokenRead
              AddToReport "Token Read", Parser.CurrentToken.Name, Parser.CurrentToken.Text, Parser.CurrentToken.TableIndex, Parser.CurrentLineNumber
              
           Case gpMsgInternalError
              AddToReport "Internal Error", "Something is horribly wrong", "", "", Parser.CurrentLineNumber
              Done = True
              
           Case gpMsgNotLoadedError
              '=== Due to the if-statement above, this case statement should never be true
              AddToReport "Not Loaded Error", "Compiled Gramar Table not loaded", "", "", 0
              Done = True
              
           Case gpMsgCommentError
              AddToReport "Comment Error", "Unexpected end of file", "", "", Parser.CurrentLineNumber
              Done = True
              
           Case gpMsgCommentBlockRead
               AddToReport "Block Comment Read", Parser.CurrentComment, "", "", Parser.CurrentLineNumber
               
           Case gpMsgCommentLineRead
               AddToReport "Line Comment Read", Parser.CurrentComment, "", "", Parser.CurrentLineNumber
               
           End Select
           
        Loop
    Else
        MsgBox "Could not load the CGT file", vbCritical
    End If
End Sub
Private Sub DrawReductionTree(TheReduction As Reduction)
    'This procedure starts the recursion that draws the parse tree.
    
    txtParseTree.Visible = False   'Keep the system from updating it until we are done
    txtParseTree.Text = ""
    
    DrawReduction TheReduction, 0
    
    txtParseTree.Visible = True
End Sub

Private Sub DrawReduction(TheReduction As Reduction, Indent As Integer)
   'This is a simple recursive procedure that draws an ASCII version of the parse
   'tree
      
   Const kIndentText = "|  "
   Dim n As Integer, IndentText As String
   
   IndentText = ""
   For n = 1 To Indent
       IndentText = IndentText & kIndentText
   Next
   
   '==== Display Reduction
   PrintParseTree IndentText & "+--" & TheReduction.ParentRule.Text
   
   '=== Display the children of the reduction
   For n = 0 To TheReduction.TokenCount - 1
       Select Case TheReduction.Tokens(n).Kind
       Case SymbolTypeNonterminal
          DrawReduction TheReduction.Tokens(n).Data, (Indent + 1)
       Case Else
          PrintParseTree IndentText & kIndentText & "+--" & TheReduction.Tokens(n).Data
       End Select
   Next

End Sub


Private Sub PrintParseTree(Text As String)
   'This sub just appends the Text to the end of the txtParseTree textbox.
    
    txtParseTree.Text = txtParseTree.Text & Text & vbNewLine
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdParse_Click()

    DoParse

End Sub












Private Sub Form_Load()
    Dim Text As String
    
    txtCGTFilePath.Text = App.Path & "\Test Code\Simple 2.cgt"
    
    '==== Enter a Simple Program
    Text = Text & "DISPLAY 'Enter a number' READ Num" & vbNewLine
    Text = Text & "ASSIGN Num = Num * 2" & vbNewLine
    Text = Text & "DISPLAY 'This the square of the number' & Num" & vbNewLine
        
    txtTestInput.Text = Text
End Sub


