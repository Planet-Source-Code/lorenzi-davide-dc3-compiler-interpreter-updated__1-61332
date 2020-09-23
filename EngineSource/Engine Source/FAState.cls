VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FAState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Class Name:
'      FAState
'
' Instancing:
'      Private; Internal  (VB Setting: 1 - Private)
'
' Purpose:
'      Represents a state in the Deterministic Finite Automata which is used by
'      the tokenizer.
'
' Author(s):
'      Devin Cook
'      GOLDParser@DevinCook.com
'
' Dependacies:
'      FAEdge
'
'================================================================================

Option Explicit

Private Edges As New ObjectArray

Public AcceptList As New NumberSet
Public TableIndex As Integer

Public Property Get AcceptSymbol() As Integer

   If AcceptList.Count >= 1 Then
      AcceptSymbol = AcceptList.Member(0)
   Else
      AcceptSymbol = -1
   End If

End Property


Public Property Let AcceptSymbol(SymbolIndex As Integer)

   AcceptList.Clear
   If SymbolIndex >= 0 Then
      AcceptList.Add SymbolIndex
   End If

End Property



Public Sub AddEdge(Characters As NumberSet, ByVal TheTarget As Integer, Optional SetIndex As Integer)
    Dim Edge As New FAEdge
   
    Set Edge.Characters = Characters
    Edge.Target = TheTarget
    Edge.SetIndex = SetIndex
            
    Edges.Add Edge
End Sub
Public Function Edge(ByVal Index As Long) As FAEdge

   If Index >= 0 And Index < Edges.Count Then
       Set Edge = Edges.Member(Index)
   End If

End Function

Public Function EdgeCount() As Integer

   EdgeCount = Edges.Count

End Function


Private Sub Class_Initialize()

'VariableLength = False

End Sub


Private Sub Class_Terminate()
    Set Edges = Nothing
    Set AcceptList = Nothing
End Sub


