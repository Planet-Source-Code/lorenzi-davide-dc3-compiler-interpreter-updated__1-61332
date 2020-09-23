VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumberSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MemberList() As Integer
Private MemberCount As Long








Public Sub Clear()

   Erase MemberList
   MemberCount = 0

End Sub
Public Function Count() As Long

   Count = MemberCount

End Function

Public Function HasMember(Number As Integer) As Boolean
'   Dim n As Long, bFound As Boolean, bFail As Boolean
'
'   'Optimization possible since the array is sorted. Loop exits if a larger
'   'item is found
'
'   bFound = False
'   bFail = False
'   n = 0
'   Do While n < MemberCount And Not bFound And Not bFail
'       If Number < MemberList(n) Then
'           bFail = True
'       Else
'           bFound = (Number = MemberList(n))
'       End If
'       n = n + 1
'   Loop
'
'   HasMember = bFound

    HasMember = (MemberIndex(Number) <> -1)

End Function


Private Function MemberIndex(ByVal Number As Integer) As Long
    Dim Done As Boolean
    Dim Upper As Long, Lower As Long, Middle As Long, Index As Long
   
    If MemberCount = 0 Then
        Index = -1
    ElseIf Number < MemberList(0) Or Number > MemberList(MemberCount - 1) Then
        Index = -1
    Else
        'THIS IS A BINARY SEARCH - since the list is sorted!
        Upper = MemberCount - 1
        Lower = 0
        Index = -1
        Done = False
          
        Do
            Middle = (Lower + Upper) / 2
            
            If Lower > Upper Then  'Two ends passed each other, fail - member not found
                Done = True
            ElseIf MemberList(Middle) = Number Then
                Index = Middle
                Done = True
            ElseIf MemberList(Middle) < Number Then
                Lower = Middle + 1
            ElseIf MemberList(Middle) > Number Then
                Upper = Middle - 1
            End If
        Loop Until Done
    End If
   
    MemberIndex = Index
End Function

Property Get Member(ByVal Index As Long) As Long
Attribute Member.VB_UserMemId = 0

   If Index >= 0 And Index < MemberCount Then
      Member = MemberList(Index)
   End If

End Property

Property Let Member(ByVal Index As Long, var As Long)
   
   If Index >= 0 And Index < MemberCount Then
      MemberList(Index) = var
   End If

End Property



Public Sub Add(ByVal Number As Integer)
   
   'Insert the Number into the list at the correct position - maintaining a sorted
   'list
   
    Dim Index As Integer, n As Integer, Found As Boolean
          
    '=== Find point of insertion, before Index.
    n = 0
    Index = -1
    Found = False
    Do Until n > MemberCount - 1 Or Index <> -1 Or Found
        If MemberList(n) = Number Then
            Found = True
        ElseIf MemberList(n) > Number Then
            Index = n
        End If
        n = n + 1
    Loop
   
    If Not Found Then
        '=== Increase array
        MemberCount = MemberCount + 1
        ReDim Preserve MemberList(0 To MemberCount - 1)
           
        If Index = -1 Then                         'Not found Add last
            MemberList(MemberCount - 1) = Number
        Else                                       'Insert,  Move items
            For n = MemberCount - 1 To Index + 1 Step -1
                MemberList(n) = MemberList(n - 1)
            Next
            MemberList(Index) = Number
        End If
    End If
End Sub

