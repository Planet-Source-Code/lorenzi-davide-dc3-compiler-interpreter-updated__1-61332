VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TokenStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' Class Name:
'      TokenStack
'
' Instancing:
'      Private; Internal  (VB Setting: 1 - Private)
'
' Purpose:
'      This class is used by the GOLDParser class to store tokens
'      during parsing. In particular, this class is used the the LALR(1)
'      state machine.
'
' Author(s):
'      Devin Cook
'
' Dependecies:
'      Token Class
'
'======================================================================
Option Explicit

Private MemberList() As Token
Private MemberCount As Long

Friend Property Let Count(Value As Long)

If Value < 1 Then
      Erase MemberList
      MemberCount = 0
   Else
      ReDim Preserve MemberList(0 To Value - 1)     'Change the size of the array
      MemberCount = Value
   End If

End Property

Public Sub Clear()

   Erase MemberList
   MemberCount = 0

End Sub

Property Get Count() As Long

   Count = MemberCount

End Property



Property Get Member(ByVal Index As Long) As Token

   If Index >= 0 And Index < MemberCount Then
      Set Member = MemberList(Index)
   Else
      Set Member = Nothing
   '   MsgBox "Nothing!"
   End If
End Property

Property Set Member(ByVal Index As Long, TheToken As Token)
   
   If Index >= 0 And Index < MemberCount Then
      Set MemberList(Index) = TheToken
   End If

End Property

Public Sub Push(TheToken As Token)

   MemberCount = MemberCount + 1
   ReDim Preserve MemberList(0 To MemberCount - 1)     'Change the size of the array
   Set MemberList(MemberCount - 1) = TheToken

End Sub






Public Function Pop() As Token

    'Modified 12/11/2001
    If MemberCount >= 1 Then
        Set Pop = MemberList(MemberCount - 1)       'Indexing from 0
        Set MemberList(MemberCount - 1) = Nothing   'List no longer points to the token
        MemberCount = MemberCount - 1
    Else
        Set Pop = Nothing
    End If

End Function

Public Function Top() As Token

   If MemberCount >= 1 Then
      Set Top = MemberList(MemberCount - 1)
   End If

End Function
