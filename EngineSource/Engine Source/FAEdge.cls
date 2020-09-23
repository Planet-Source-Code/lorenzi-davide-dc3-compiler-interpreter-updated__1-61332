VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FAEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Class Name:
'      FAEdge
'
' Instancing:
'      Private; Internal  (VB Setting: 1 - Private)
'
' Purpose:
'      Each state in the Determinstic Finite Automata contains multiple edges which
'      link to other states in the automata.
'
'      This class is used to represent an edge.
'
' Author(s):
'      Devin Cook
'      GOLDParser@DevinCook.com
'
' Dependacies:
'      (None)
'
'================================================================================
   
Public Characters  As NumberSet       'Characters to advance on
   
Public SetIndex As Integer          'Index in the Character Set table
Public Target As Integer 

