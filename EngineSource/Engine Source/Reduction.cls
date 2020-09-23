VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Reduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================================================
' Class Name:
'      Reduction
'
' Instancing:
'      Public; Creatable  (VB Setting: 5 - MultiUse)
'
' Purpose:
'      This class is used by the engine to hold a reduced rule. Rather the contain
'      a list of Symbols, a reduction contains a list of Tokens corresponding to the
'      the rule it represents. This class is important since it is used to store the
'      actual source program parsed by the Engine.
'
' Author(s):
'      Devin Cook
'      GOLDParser@DevinCook.com
'
' Dependacies:
'      Token Class, Rule Class
'
'==========================================================================

Option Explicit



Private pTokens() As Token
Private pTokenCount As Integer

Private pParentRule As Rule

Private pTag As Integer            'General purpose

Friend Property Let TokenCount(Value As Integer)

   If Value < 1 Then
      Erase pTokens
      pTokenCount = 0
   Else
      pTokenCount = Value
      ReDim Preserve pTokens(0 To pTokenCount - 1)     'Change the size of the array
   End If

End Property


Public Property Get ParentRule() As Rule

   Set ParentRule = pParentRule

End Property

Friend Property Set ParentRule(Value As Rule)

   Set pParentRule = Value

End Property

Property Let Tag(Value As Integer)
    pTag = Value
End Property

Property Get Tag() As Integer
    Tag = pTag
End Property

Public Property Get TokenCount() As Integer
    TokenCount = pTokenCount
End Property

Public Property Get Tokens(Index As Integer) As Token

   If Index >= 0 And Index < pTokenCount Then
      Set Tokens = pTokens(Index)
   Else
      Set Tokens = Nothing
   End If

End Property


Friend Property Set Tokens(Index As Integer, Value As Token)

   If Index >= 0 And Index < pTokenCount Then
      Set pTokens(Index) = Value
   'Else
   '  MsgBox "Error"
   End If

End Property




