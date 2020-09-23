VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================
' Class Name:
'      Rule
'
' Instancing:
'      Public; Non-creatable  (VB Setting: 2- PublicNotCreatable)
'
' Purpose:
'      The Rule class is used to represent the logical structures of the grammar.
'      Rules consist of a head containing a nonterminal followed by a series of
'      both nonterminals and terminals.
'
' Author(s):
'      Devin Cook
'      GOLDParser@DevinCook.com
'
' Dependacies:
'      Symbol Class, SymbolList Class
'
'========================================================================
Option Explicit

Private pRuleNonterminal As Symbol
Private pRuleSymbols As New ObjectArray   'This consist the body of the rule
Private pTableIndex As Integer

Friend Function ContainsOneNonTerminal() As Boolean
    'New 12/2001 - used by the GOLDParser object to TrimReductions
    
    Dim Result As Boolean
    
    Result = False

    If pRuleSymbols.Count = 1 Then
        If pRuleSymbols.Member(0).Kind = SymbolTypeNonterminal Then
            Result = True
        End If
    End If
    
    ContainsOneNonTerminal = Result
End Function


Friend Function Definition() As String
   Dim n As Integer, str As String
   
   For n = 0 To pRuleSymbols.Count - 1
      str = str & pRuleSymbols.Member(n).Text & " "
   Next
   
   Definition = RTrim(str)
End Function



Public Property Get SymbolCount() As Integer

   SymbolCount = pRuleSymbols.Count

End Property

Friend Function Name() As String

   Name = "<" & pRuleNonterminal.Name & ">"

End Function



Friend Sub SetRuleNonterminal(Nonterminal As Symbol)

   Set pRuleNonterminal = Nonterminal

End Sub


Public Property Get RuleNonterminal() As Symbol

   Set RuleNonterminal = pRuleNonterminal

End Property

Public Property Get Symbols(Index As Integer) As Symbol

   If Index >= 0 And Index < pRuleSymbols.Count Then
      Set Symbols = pRuleSymbols.Member(Index)
   End If

End Property



Friend Sub AddItem(Item As Symbol)

   pRuleSymbols.Add Item

End Sub
Friend Property Let TableIndex(Index As Integer)

   pTableIndex = Index

End Property

Public Property Get TableIndex() As Integer

   TableIndex = pTableIndex

End Property
Public Function Text() As String
   
   Text = Name() & " ::= " & Definition()

End Function


Private Sub Class_Initialize()
    pTableIndex = -1
End Sub

