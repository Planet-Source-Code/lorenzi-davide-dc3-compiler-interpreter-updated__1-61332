VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Class Name:
'      ObjectArray
'
' Instancing:
'      Private; Internal  (VB Setting: 1 - Private)
'
' Purpose:
'      This class is used to store a list of objects. In Visual Basic, an object
'      can contain any class. This class would have been used universally in the
'      code, but limitations imposed by the Visual Basic programming language
'      required the creation of specialized list classes.
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


Private MemberList() As Object
Private MemberCount As Long

Public Sub Clear()

   Erase MemberList
   MemberCount = 0

End Sub

Public Function Count() As Long

   Count = MemberCount

End Function

Property Get Member(ByVal Index As Long) As Object

   If Index >= 0 And Index < MemberCount Then
       Set Member = MemberList(Index)
   Else
       Set Member = Nothing
   '   MsgBox "Nothing!"
   End If
End Property

Property Let Member(ByVal Index As Long, Obj As Object)
   
   If Index >= 0 And Index < MemberCount Then
      MemberList(Index) = Obj
   End If

End Property

Property Set Member(ByVal Index As Long, Obj As Object)
   
   If Index >= 0 And Index < MemberCount Then
      Set MemberList(Index) = Obj
   End If

End Property

Public Sub Add(Obj As Object)

   MemberCount = MemberCount + 1
   ReDim Preserve MemberList(0 To MemberCount - 1)
   Set MemberList(MemberCount - 1) = Obj

End Sub

Public Function IsEqual(Obj As ObjectArray) As Boolean
   Dim n As Integer
   Dim Equal As Boolean
   
   If MemberCount = Obj.Count Then
     Equal = True
     n = 0
     Do While Equal And n = Count
       Equal = MemberList(n) = Obj.Member(n)
       n = n + 1
     Loop
   Else
     Equal = False
   End If
   
   IsEqual = Equal

End Function



Public Function Pop() As Object

   MsgBox "Pop on ObjectArray"
   
   'Set Pop = MemberList(MemberCount)
   
   'MemberCount = MemberCount - 1
   'ReDim Preserve MemberList(1 To MemberCount)

End Function
Public Sub ReDimension(ByVal NewCount As Long)
   
   MemberCount = NewCount
   If NewCount = 0 Then
       ReDim MemberList(0 To 0)
       Set MemberList(0) = Nothing
   Else
       ReDim Preserve MemberList(0 To MemberCount - 1)
   End If

End Sub


Private Sub Class_Terminate()
    Dim n As Integer
    For n = 0 To MemberCount - 1
        Set MemberList(n) = Nothing
    Next
End Sub


