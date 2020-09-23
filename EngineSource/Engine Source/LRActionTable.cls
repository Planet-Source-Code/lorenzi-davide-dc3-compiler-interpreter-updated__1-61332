VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LRActionTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Class Name:
'      LRActionTable
'
' Instancing:
'      Private; Internal  (VB Setting: 1 - Private)
'
' Purpose:
'      This class contains the actions (reduce/shift) and goto information
'      for a STATE in a LR parser. Essentially, this is just a row of actions in
'      the LR state transition table. The only data structure is a list of
'      LR Actions.
'
' Author(s):
'      Devin Cook
'      GOLDParser@DevinCook.com
'
' Dependacies:
'      LRAction Class
'
'================================================================================

Option Explicit

Private MemberList() As LRAction
Private MemberCount As Long


Public Function ActionIndexForSymbol(SymbolIndex As Integer) As Integer
   'Returns the index of SymbolIndex in the table, -1 if not found
   Dim n As Integer, Found As Boolean, Index As Integer
   
   n = 0
   Found = False
   Do While Not Found And n < MemberCount
      If MemberList(n).Symbol.TableIndex = SymbolIndex Then
         Index = n
         Found = True
      End If
      n = n + 1
   Loop
   
   If Found Then
      ActionIndexForSymbol = Index
   Else
      ActionIndexForSymbol = -1
   End If

End Function


Public Sub AddItem(TheSymbol As Symbol, Action As ActionConstants, Value As Integer)
   Dim TableEntry As New LRAction
   
   Set TableEntry.Symbol = TheSymbol
   TableEntry.Action = Action
   TableEntry.Value = Value
   
   MemberCount = MemberCount + 1
   ReDim Preserve MemberList(0 To MemberCount - 1)     'Change the size of the array
   Set MemberList(MemberCount - 1) = TableEntry

End Sub




Public Function Count() As Integer

   Count = MemberCount

End Function

Public Function Item(ByVal n As Integer) As LRAction

   If n >= 0 And n < MemberCount Then
      Set Item = MemberList(n)
   End If

End Function


