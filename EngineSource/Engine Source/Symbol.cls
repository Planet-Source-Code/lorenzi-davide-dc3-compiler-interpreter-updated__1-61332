VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Symbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================
' Class Name:
'      Symbol
'
' Instancing:
'      Public; Non-creatable  (VB Setting: 2- PublicNotCreatable)
'
' Purpose:
'       This class is used to store of the nonterminals used by the Deterministic
'       Finite Automata (DFA) and LALR Parser. Symbols can be either
'       terminals (which represent a class of tokens - such as identifiers) or
'       nonterminals (which represent the rules and structures of the grammar).
'       Terminal symbols fall into several catagories for use by the GOLD Parser
'       Engine which are enumerated below.
'
' Author(s):
'      Devin Cook
'      GOLDParser@DevinCook.com
'
' Dependacies:
'      (None)
'
'====================================================================


Option Explicit

Private pName As String
Private pKind As SymbolTypeConstants

Private pPattern As String
Private pVariableLength As Boolean

Public Enum SymbolTypeConstants
   SymbolTypeNonterminal = 0     'Normal nonterminal
   SymbolTypeTerminal = 1        'Normal terminal
   SymbolTypeWhitespace = 2      'Type of terminal
   SymbolTypeEnd = 3             'End character (EOF)
   SymbolTypeCommentStart = 4    'Comment start
   SymbolTypeCommentEnd = 5      'Comment end
   SymbolTypeCommentLine = 6     'Comment line
   SymbolTypeError = 7           'Error symbol
End Enum

Private pTableIndex As Integer

Private Const kQuotedChars = "|-+*?()[]{}<>!"


Friend Property Let Kind(TheType As SymbolTypeConstants)

   pKind = TheType

End Property

Public Property Get Kind() As SymbolTypeConstants

   Kind = pKind

End Property

Friend Property Let TableIndex(Index As Integer)

   pTableIndex = Index

End Property

Public Property Get TableIndex() As Integer

   TableIndex = pTableIndex

End Property

Friend Property Let Name(NewName As String)

   pName = NewName

End Property

Public Property Get Name() As String

   Name = pName

End Property


Public Property Get Text() As String
   Dim str As String
   
   Select Case Kind
   Case SymbolTypeNonterminal
      str = "<" & Name & ">"
   Case SymbolTypeTerminal
      str = PatternFormat(Name)
   Case Else
      str = "(" & Name & ")"
   End Select
   
   Text = str
End Property

Private Function PatternFormat(Source As String) As String

   '=== Create a valid Regular Expression for a source string
   '=== Put all special characters in single quotes
   
   Dim c As Integer, Result As String, ch As String
   
   For c = 1 To Len(Source)
      ch = Mid(Source, c, 1)
      If ch = "'" Then
         ch = "''"
      ElseIf InStr(kQuotedChars, ch) <> 0 Or ch = Chr(34) Then
         ch = "'" & ch & "'"
      End If
      
      Result = Result & ch
   Next
   
   PatternFormat = Result
End Function

Private Sub Class_Initialize()
    pTableIndex = -1
End Sub


