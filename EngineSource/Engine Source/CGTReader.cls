VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGTReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'================================================================================
' Class Name:
'      CGT Reader
'
' Instancing:
'      Private; Internal  (VB Setting: 1 - Private)
'
' Purpose:
'      This class is used to read information stored in the very simple file
'      structure used by the Compiled Grammar Table file.
'
'      Record Structure
'         |1 byte         | 2 byte      | Variable size
'         |Character 'M'  | Entry Count | Entries (0 to 32k)
'
'      For more information, please consult http://www.DevinCook.com/GOLDParser/doc/cgt/
'
'      Version 2.0
'         * The class was adapted from SimpleDataBase
'         * The dependancy on the Stream class was removed.
'
'      Version 1.1
'         * Now uses the Lookahead Stream class to read and write data
'
' Author(s):
'      Devin Cook
'
' Dependacies:
'      (None)
'
'================================================================================
Private Const kRecordContentMulti = 77                 'M
Private Const kFileHeader = "GOLD Parser Tables/v1.0"

Public Enum CGTEntryConstants
    CGTEntryEmpty = 69        'E
    CGTEntryInt16 = 73        'I - Signed, 2 byte
    CGTEntryString = 83       'S - Unicode format
    CGTEntryBoolean = 66      'B - 1 Byte, Value is 0 or 1
    'CGTEntryInt32 = 76       'L
    CGTEntryByte = 98         'b
End Enum

'============================================================
Public Enum CGTRecTypeConstants
    CGTRecTypeParameters = 80     'P
    CGTRecTypeCounts = 84         'T
    CGTRecTypeInitial = 73        'I
    CGTRecTypeSymbol = 83         'S
    CGTRecTypeCharSet = 67        'C
    CGTRecTypeRule = 82           'R
    CGTRecTypeDFAState = 68       'D
    CGTRecTypeLRState = 76        'L
    CGTRecTypeComment = 33        '!
End Enum

'============================================================

Private EntryList()       As Variant
Private EntryListCount    As Integer
Private EntryReadPosition As Integer
Private pRecordType       As CGTRecTypeConstants

Private pFileMode    As String * 1
Private pFileNumber  As Integer
Private pFileIsOpen  As Boolean

Private pErrorDescription As String

Public Function ErrorDescription() As String
    ErrorDescription = pErrorDescription
End Function

Public Function RetrieveDone() As Boolean

   RetrieveDone = Not (EntryReadPosition < EntryListCount)

End Function


Public Function RetrieveNext() As Variant

   If Not RetrieveDone() Then
      RetrieveNext = EntryList(EntryReadPosition)
      EntryReadPosition = EntryReadPosition + 1
   Else
      RetrieveNext = Empty
   End If

End Function

Public Sub Store(ParamArray Variables() As Variant)

   Dim Lower As Integer, Upper As Integer, Length As Integer
   Dim n As Integer
   
   Lower = LBound(Variables)
   Upper = UBound(Variables)
   Length = Upper - Lower + 1
   
   ReDim Preserve EntryList(0 To EntryListCount + Length - 1)
   
   For n = Lower To Upper
      EntryList(EntryListCount) = Variables(n)
      EntryListCount = EntryListCount + 1
   Next
   DoEvents

End Sub

Public Sub Clear()
   
   EntryListCount = 0
   EntryReadPosition = 0
   Erase EntryList

End Sub

Public Sub CloseFile()

   If pFileIsOpen Then
      Close pFileNumber
      pFileIsOpen = False
   End If

End Sub

Public Function EntryCount() As Integer
   EntryCount = EntryListCount
End Function

Public Function EOF() As Boolean
    EOF = VBA.EOF(pFileNumber)
End Function


Public Function Entry(ByVal Index As Integer) As Variant
   If Index >= 0 And Index < EntryListCount Then
      Entry = EntryList(Index)
   Else
      Entry = Null
   End If
End Function




Public Function OpenFile(FileName As String, Mode As String) As Boolean
   On Error GoTo Problem
 
   Dim Success As Boolean, ErrorDesc As String
   Dim Header As String
      
   pFileMode = IIf(UCase(Mode) = "W", "W", "R")
   Success = True
   
   '======== Close the file if necessary
   If pFileIsOpen Then
      Close pFileNumber
   End If
      
   '======== Open the file to read or write
   If pFileMode = "W" Then
      'DeleteFile FileName     'Delete the file
      
      pFileNumber = FreeFile()
      Open FileName For Binary Access Write As pFileNumber
      
      WriteString kFileHeader
   Else
      pFileNumber = FreeFile()
      Open FileName For Binary Access Read As pFileNumber
      
      Header = ReadString()
           
      If Header <> kFileHeader Then
          ErrorDesc = "Invalid file"
          Success = False
      End If
   End If
      
      
   pFileIsOpen = Success
   pErrorDescription = ErrorDesc
   OpenFile = Success
   
   Exit Function
   
Problem:
   OpenFile = False
   pErrorDescription = Err.Description
   
   Err.Clear

End Function


Private Sub EntryWrite(Value As Variant)
    Dim b As Byte, i As Long, char As Integer
   
    Select Case VarType(Value)
        Case vbEmpty
           WriteByte CGTEntryEmpty
           
        Case vbBoolean
           b = IIf(Value, 1, 0)
           WriteByte CGTEntryBoolean
           WriteBoolean Value
        
        Case vbInteger, vbLong
           WriteByte CGTEntryInt16
           WriteInt16 Value
           
        Case vbByte
           WriteByte CGTEntryByte
           WriteByte Value
           
        Case vbString
           WriteByte CGTEntryString
           WriteString Value
    End Select
End Sub

Private Function EntryRead() As Variant
    Dim ID As Byte, Value As Variant, Result As Variant
    
    ID = ReadByte()
      
    Select Case ID
        Case CGTEntryEmpty
            Result = Empty
            
        Case CGTEntryBoolean
            Result = ReadBoolean()
            
        Case CGTEntryInt16
            Result = ReadInt16()
            
        Case CGTEntryString
            Result = ReadString()
            
        Case CGTEntryByte
            Result = ReadByte()
            
        Case Else
            MsgBox "Error on CGT read"
            Result = Empty
    End Select
   
    EntryRead = Result
End Function




Private Function ReadString() As String
   Dim Done As Boolean, char16 As Integer
   Dim Result As String
     
   Done = False
   
   Do Until Done
       Get #pFileNumber, , char16
       
       If char16 = 0 Then
           Done = True
       Else
           Result = Result & ChrW(char16)
       End If
   Loop
   
   ReadString = Result
End Function

Private Function ReadInt16() As Integer
   Dim Value As Integer
          
   Get #pFileNumber, , Value
   
   ReadInt16 = Value
End Function

Private Function ReadBoolean() As Boolean
   Dim Value As Byte
          
   Get #pFileNumber, , Value
   
   ReadBoolean = (Value = 1)
End Function

Private Function ReadByte() As Byte
   Dim Value As Byte
          
   Get #pFileNumber, , Value
   
   ReadByte = Value
End Function

Private Sub WriteString(ByVal Text As String)
   Dim n As Integer, char16 As Integer
     
   For n = 1 To Len(Text)
       char16 = AscW(Mid(Text, n, 1))
       Put #pFileNumber, , char16
   Next
   
   Put #pFileNumber, , CInt(0)
End Sub

Private Sub WriteInt16(ByVal Value As Integer)
   Put #pFileNumber, , Value
End Sub

Private Sub WriteByte(ByVal Value As Byte)
   Put #pFileNumber, , Value
End Sub

Private Sub WriteBoolean(ByVal Value As Boolean)
   Put #pFileNumber, , CByte(IIf(Value, 1, 0))
End Sub

Public Function EndRecord() As Boolean
   Dim n As Integer
   
   If pFileMode = "W" And pFileIsOpen Then
      WriteByte kRecordContentMulti
      WriteInt16 (EntryListCount + 1)
      
      EntryWrite CByte(pRecordType)   'Write first entry!
      
      For n = 0 To EntryListCount - 1
         EntryWrite EntryList(n)
      Next
      
      EndRecord = True
      Clear
   Else
      EndRecord = False
   End If

End Function


Public Function GetNextRecord() As Boolean

   Dim n As Integer, Count As Integer, ID As Byte, Success As Boolean
   Dim Value As Variant
   
   If pFileMode = "R" And (Not EOF()) Then
      Clear
      ID = ReadByte()
      
      Select Case ID
      Case kRecordContentMulti
          Count = ReadInt16()
          pRecordType = EntryRead()
          
          For n = 1 To Count - 1
             Store EntryRead()
          Next
          EntryReadPosition = 0
          Success = True
   
      Case Else
          Success = False
      End Select
        
   Else
      Success = False
   End If
   
   GetNextRecord = Success

End Function


Private Sub Class_Initialize()
    EntryReadPosition = 1
    pFileIsOpen = False
End Sub

Private Sub Class_Terminate()
    CloseFile
End Sub

Public Property Get RecordType() As CGTRecTypeConstants
    RecordType = pRecordType
End Property
Public Property Let RecordType(Value As CGTRecTypeConstants)
    pRecordType = Value
End Property
