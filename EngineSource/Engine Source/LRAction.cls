VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LRAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Class Name:
'      LRAction
'
' Instancing:
'      Private; Internal  (VB Setting: 1 - Private)
'
' Purpose:
'      This class represents an action in a LALR State. There is one and only one
'      action for any given symbol.
'
' Author(s):
'      Devin Cook
'      GOLDParser@DevinCook.com
'
' Dependacies:
'      (None)
'
'================================================================================
Option Explicit

Public Enum ActionConstants
   ActionShift = 1       'Shift a symbol and goto a state
   ActionReduce = 2      'Reduce by a specified rule
   ActionGoto = 3        'Goto to a state on reduction
   ActionAccept = 4      'Input successfully parsed
   ActionError = 5       'Programmars see this often!
End Enum

Private pSymbol As Symbol
Public Action As ActionConstants
Public Value As Integer      'shift to state, reduce rule, goto state




Public Property Set Symbol(Sym As Symbol)

   Set pSymbol = Sym

End Property


Public Property Get Symbol() As Symbol

   Set Symbol = pSymbol

End Property


Public Function SymbolIndex() As Integer

   SymbolIndex = pSymbol.TableIndex

End Function





