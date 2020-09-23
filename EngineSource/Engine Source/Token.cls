VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' Class Name:
'      Token
'
' Instancing:
'      Public; Creatable  (VB Setting: 5 - MultiUse)
'
' Purpose:
'       While the Symbol represents a class of terminals and nonterminals, 
'       the Token represents an individual piece of information.
'       Ideally, the token would inherit directly from the Symbol Class, 
'       but do to the fact that Visual Basic 5/6 does not support this aspect 
'       of Object Oriented Programming, a Symbol is created as a member and 
'       its methods are mimicked.
'
' Author(s):
'      Devin Cook
'      GOLDParser@DevinCook.com
'
' Dependecies:
'      Symbol class
'
'====================================================================

Option Explicit

Private pState As Integer
Private pData As Variant
Private pParentSymbol As Symbol




Public Property Get Kind() As SymbolTypeConstants

   Kind = ParentSymbol.Kind

End Property


Public Property Get Name() As String
   Name = ParentSymbol.Name
End Property
Public Property Get ParentSymbol() As Symbol
    Set ParentSymbol = pParentSymbol
End Property

Public Property Get Data() As Variant
    'The ugliness of this is caused by that fact that Visual Basic 5/6 treats
    'object different from scalar data types. This will be resolved in VB .NET
        
    If VarType(pData) = vbObject Then
        Set Data = pData
    Else
        Data = pData
    End If
End Property

Public Property Let Data(Value As Variant)
    pData = Value
End Property

Public Property Set Data(Value As Variant)
    Set pData = Value
End Property

Public Property Set ParentSymbol(TheSymbol As Symbol)
    Set pParentSymbol = TheSymbol
End Property


Public Property Get TableIndex() As Integer

   TableIndex = ParentSymbol.TableIndex

End Property

Public Property Get Text() As String
   Text = ParentSymbol.Text
End Property

Friend Property Let State(Value As Integer)

   pState = Value

End Property

Friend Property Get State() As Integer

   State = pState

End Property


Private Sub Class_Initialize()

Data = Empty

End Sub


