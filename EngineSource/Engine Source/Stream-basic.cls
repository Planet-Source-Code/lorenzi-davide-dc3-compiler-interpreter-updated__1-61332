VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Class Name:
'      Stream (Basic version)
'
' Instancing:
'      Private; Internal  (VB Setting: 1 - Private)
'
' Purpose:
'       This class is used to create a very basic version of the stream object that
'       exists in the C++ programming language. Using streams, the programmer can read
'       data from any number of sources while masking this from the code using the
'       stream. Visual Basic .NET will add this functionality.
' Author(s):
'      Devin Cook
'
' Dependacies:
'      (None)
'
'================================================================================
Option Explicit

Public Enum StreamTransferConstants
   StreamTransferBinary = 1
   StreamTransferText = 2
End Enum

'==== For deciding whether to write strings in Binary mode
'==== as C++ (null end) or with a 2-byte predicessor field
Public Enum StreamStringFormConstants
   StreamStringFormNull = 1           '== Null character delimited
   StreamStringFormNone = 2           '== Raw data
   StreamStringFormField = 3          '== 2-Byte prefix integer
End Enum

Public Enum StreamStringTypeConstants
   StreamStringTypeASCII = 1
   StreamStringTypeUnicode = 2
End Enum

Public Enum StreamEncodingConstants
   StreamEncodingLittleEndian = 1     '== For those x86 guys
   StreamEncodingBigEndian = 2        '== For those Motorola guys
End Enum

Private Enum StreamTargetConstants
   StreamTargetString = 1
   StreamTargetFile = 2
End Enum


'========= File Streaming
Private m_FileNumber As Integer
Private m_IsFileOpen As Boolean
Private m_StreamTransfer As StreamTransferConstants

'========= The main data type and control
Private m_Buffer As String             'Up to 2 billion chars

Private m_StreamTarget As StreamTargetConstants
Private m_StreamCanRead As Boolean

'========= Public data types
Private m_StringForm  As StreamStringFormConstants
Private m_StringType  As StreamStringTypeConstants
Private m_Encoding    As StreamEncodingConstants

'========= This is a very useful ADT that organizes very complex binary data
Private Type ByteSequence
    Length As Integer             'In bytes, not bits
    Bytes(0 To 15) As Byte        '0 is LSB
End Type



Public Property Get Encoding() As StreamEncodingConstants
    Encoding = m_Encoding
End Property

Public Property Let Encoding(ByVal Value As StreamEncodingConstants)
    m_Encoding = Value
End Property

Private Function ReadChar(Optional DiscardData As Boolean = True) As String
    '======= This function reads a ASCII or Unicode character
    Dim Result As String
    
    Select Case m_StreamTarget
    Case StreamTargetFile
        Select Case m_StringType
        Case StreamStringTypeASCII
            Result = StreamRead(1, DiscardData)
        Case StreamStringTypeUnicode
            Result = ChrW(DecodeInt16(ReadByteSequence(2, DiscardData)))
        End Select
    
    Case StreamTargetString
        Result = StreamRead(1, DiscardData)
        
    End Select

    ReadChar = Result
End Function

Public Property Get StringType() As StreamStringTypeConstants
    StringType = m_StringType
End Property

Public Property Let StringType(ByVal Value As StreamStringTypeConstants)
    m_StringType = Value
End Property
Public Property Get StringForm() As StreamStringFormConstants
    StringForm = m_StringForm
End Property

Public Property Let StringForm(ByVal Value As StreamStringFormConstants)
    m_StringForm = Value
End Property

Private Sub CloseStream()
 
   If m_StreamTarget = StreamTargetFile And m_IsFileOpen Then
       Close m_FileNumber
       m_IsFileOpen = False
   End If
   
End Sub

Public Function EOF() As Boolean
    EOF = Done()
End Function

Public Function IsFileOpen() As Boolean
   IsFileOpen = m_IsFileOpen

End Function


Private Function ReadByteSequence(Length As Integer, Optional DiscardData As Boolean = True) As ByteSequence

   'This function is where I will put the Big/Little endian code
   
   Dim str As String, Result As ByteSequence, n As Integer
   str = StreamRead(Length, DiscardData)
   
   Result.Length = Length
   
   If m_Encoding = StreamEncodingBigEndian Then
      For n = 1 To Length
         Result.Bytes(Length - n) = Asc(Mid(str, n, 1))
      Next
   Else
      For n = Length To 1 Step -1
         Result.Bytes(n - 1) = Asc(Mid(str, n, 1))
      Next
   End If
   
   ReadByteSequence = Result

End Function

Public Function ReadValue(DataType As VbVarType) As Variant

    'This function reads the appropiate number of characters from the buffer
    'and reconstructs the data type. It is then returned.
   
    Dim Variable As Variant, Length As Long, n As Long, Char As String
    Dim t As String
   
    If m_StreamTransfer = StreamTransferBinary Then
        Select Case DataType
        Case vbInteger
            Variable = DecodeInt16(ReadByteSequence(2))
        Case vbLong
            Variable = DecodeInt32(ReadByteSequence(4))
        Case vbByte
            t = StreamRead(1)
            If t = "" Then
                Variable = 0
            Else
                Variable = Asc(t)
            End If
        Case vbString
            Variable = ""
            Select Case m_StringForm
            Case StreamStringFormField
                Length = DecodeInt16(ReadByteSequence(2))
                For n = 1 To Length
                    Variable = Variable & ReadChar()
                Next
                
            Case StreamStringFormNull
                Char = ReadChar()
                Do Until Asc(Char) = 0
                    Variable = Variable & Char
                    Char = ReadChar()
                Loop
            End Select
        End Select
    Else
       Variable = ReadUntil(Chr(0))
    End If
   
    ReadValue = Variable
End Function

Private Sub CopyBits(Source As ByteSequence, Start As Integer, Finish As Integer, Target As ByteSequence, TargetStart As Integer)

   Dim n As Integer, OffSet As Integer
   'A Complex procedure to be sure
   OffSet = Start - TargetStart
   
   For n = Start To Finish
      SetBit Target, n - OffSet, GetBit(Source, n)
   Next n

End Sub



Private Function DecodeFP32(Source As ByteSequence) As Single
   Dim FPExp As ByteSequence, Fraction As ByteSequence
   Dim Exponent As Integer, Result As Single
   
   ' S EEEEEEEE FFFFFFFFFFFFFFFFFFFFFFF
   '31 30    23 22                    0
   '===================================
   'Bias=127, Denominator=2^23
   
   Fraction.Length = 3
   CopyBits Source, 0, 22, Fraction, 0
   
   FPExp.Length = 1
   CopyBits Source, 23, 30, FPExp, 0
   
   Exponent = CNumber(FPExp)
   
   If Exponent = 0 Then                               'Unnormalized
     'Result = 2 ^ (-126) * (CNumber(Fraction) / (2 ^ 23))
      Result = 2 ^ (-149) * CNumber(Fraction)
   Else                                               'Normalized
      Result = 2 ^ (Exponent - 127) * (1 + CNumber(Fraction) / (2 ^ 23))
   End If
   
   If GetBit(Source, 31) = 1 Then Result = -Result
   
   DecodeFP32 = Result

End Function


Private Function EncodeInt16(ByVal Value As Integer) As ByteSequence
   Dim Temp As ByteSequence
   
   If Value < 0 Then                   'Negative - two's compliment: -1 & negate
      Temp = CByteSequence(Abs(Value + 1), 2)   'Note: Abs(Value+1) == Abs(Value)-1 when Value<0
      EncodeInt16 = Negate(Temp)
   Else                                'Positive or zero
      EncodeInt16 = CByteSequence(Value, 2)
   End If

End Function

Private Function DecodeInt16(Source As ByteSequence) As Integer

   Dim Temp As ByteSequence
   
   If GetBit(Source, 15) = 1 Then      'Negative
      Temp = Negate(Source)
      DecodeInt16 = -(CNumber(Temp) + 1)
   Else                                'Positive or zero
      DecodeInt16 = CNumber(Source)
   End If

End Function

Private Function DecodeInt32(Source As ByteSequence) As Long

   Dim Temp As ByteSequence
   
   If GetBit(Source, 31) = 1 Then      'Negative
      Temp = Negate(Source)
      DecodeInt32 = -(CNumber(Temp) + 1)
   Else                                'Positive or zero
      DecodeInt32 = CNumber(Source)
   End If

End Function
Private Function EncodeInt32(ByVal Value As Double) As ByteSequence
   Dim Temp As ByteSequence
   
   If Value < 0 Then                   'Negative - two's compliment: -1 & negate
      Temp = CByteSequence(Abs(Value + 1), 4)   'Note: Abs(Value+1) == Abs(Value)-1 when Value<0
      EncodeInt32 = Negate(Temp)
   Else                                'Positive or zero
      EncodeInt32 = CByteSequence(Value, 4)
   End If

End Function


Private Function CByteSequence(ByVal Value As Variant, ByteCount As Integer) As ByteSequence

   Dim TheByte As Byte, i As Integer, Result As ByteSequence
   
   '===Copy whole bytes
   For i = 0 To ByteCount - 1
      TheByte = Int(Value) - Int(Value / 256) * 256
      Result.Bytes(i) = TheByte
      Value = Fix(Value / 256)
   Next
   
   Result.Length = ByteCount
   
   CByteSequence = Result

End Function


Private Function CNumber(Sequence As ByteSequence) As Double

   Dim TheByte As Byte, i As Integer, Result As Double
   
   Result = 0
   '===Copy whole bytes
   For i = Sequence.Length - 1 To 0 Step -1
      Result = Result * 256 + Sequence.Bytes(i)          'Move
   Next
   
   CNumber = Result

End Function
Private Function Negate(Source As ByteSequence) As ByteSequence

   Dim Result As ByteSequence, n  As Integer
   
   Result = Source
   For n = 0 To Result.Length - 1
      Result.Bytes(n) = Not Result.Bytes(n)
   Next
   
   Negate = Result

End Function

Private Sub SetBits(Target As ByteSequence, Start As Integer, Finish As Integer, Value As Integer)
   
   Dim n As Integer
   For n = Start To Finish
   SetBit Target, n, Value
   Next n

End Sub



Private Function GetBit(Source As ByteSequence, BitNumber As Integer) As Byte
   
   'Returns the bit found in the BinaryBuffer starting at index
   '0 in last byte and moving to the left
   
   Dim ByteIndex As Integer, BitIndex As Integer
   
   ByteIndex = BitNumber \ 8        'Whole bytes
   BitIndex = BitNumber Mod 8
   
   If ByteIndex < Source.Length Then
     GetBit = IIf((Source.Bytes(ByteIndex) And (2 ^ BitIndex)) = 0, 0, 1)
   Else                        'Outside buffer = assume 0
     GetBit = 0
   End If

End Function

Private Sub SetBit(Source As ByteSequence, BitNumber As Integer, BitValue As Integer)

   'If the BitValue = 0 then the bit is cleared. Else, the bit is set
   
   Dim ByteIndex As Integer, BitIndex As Integer
   
   ByteIndex = BitNumber \ 8       'Whole bytes
   BitIndex = BitNumber Mod 8
   
   If ByteIndex < Source.Length Then
     If BitValue = 0 Then                  'Clear
        Source.Bytes(ByteIndex) = Source.Bytes(ByteIndex) And (255 - 2 ^ BitIndex)
     Else                                  'Set
        Source.Bytes(ByteIndex) = Source.Bytes(ByteIndex) Or (2 ^ BitIndex)
     End If
   End If

End Sub




Public Sub CloseFile()

   If m_IsFileOpen Then
       CloseStream
   End If

End Sub

Public Function Done() As Boolean

   Select Case m_StreamTarget
   Case StreamTargetString
       Done = Len(m_Buffer) = 0
   Case StreamTargetFile
       If Not m_IsFileOpen Then
           Done = True
       Else
           Done = VBA.EOF(m_FileNumber)
       End If
   End Select

End Function


Public Function OpenFile( _
    ByVal FileName As String, _
    Optional ByVal FileType As StreamTransferConstants = StreamTransferText, _
    Optional ByVal DetectEncodingFromByteOrderMarks As Boolean = True) As Boolean
    
    On Error Resume Next
    Dim Success As Boolean, n As Long
    Dim BOM1 As Byte, BOM2 As Byte
    
      
    '=== BOM: FE FF for big-endian, FF FE for little-endian
   
    If m_IsFileOpen Then
        Close m_FileNumber
    End If
   
    m_StreamCanRead = False
    m_StreamTransfer = FileType
   
    Success = True             'Unless determined otherwise
    
    m_StreamCanRead = True
    m_FileNumber = FreeFile()
    Open FileName For Binary Access Read As m_FileNumber
   
    '========================= React to error
    If Err.Number = 0 And Success = True Then
        m_IsFileOpen = True
        m_StreamTarget = StreamTargetFile
        m_StringForm = StreamStringFormNone
    Else
        Err.Clear
        m_IsFileOpen = False
        Success = False
    End If


    '========================= Setup endianess, etc.. information
    If Success And _
       DetectEncodingFromByteOrderMarks And _
       m_StreamTransfer = StreamTransferText And _
       LOF(m_FileNumber) >= 2 Then
        
        '=== Read BOM - determine endianess
        BOM1 = Asc(Input(1, m_FileNumber))
        BOM2 = Asc(Input(1, m_FileNumber))
        
        If BOM1 = &HFE And BOM2 = &HFF Then
            m_Encoding = StreamEncodingBigEndian
            m_StringType = StreamStringTypeUnicode
        ElseIf BOM1 = &HFF And BOM2 = &HFE Then
            m_Encoding = StreamEncodingLittleEndian
            m_StringType = StreamStringTypeUnicode
        Else
            Seek m_FileNumber, 1                     'Rewind!
            m_StringType = StreamStringTypeASCII
        End If
    End If
    
    OpenFile = Success
End Function

Public Function ReadIn(ParamArray Values() As Variant) As Boolean

   'Used in both binary and text modes
   'I wish I could use "print", but VB won't let me - dammit!
   'Returns whether the datatype if Variant could be stored
   
   Dim n As Integer
   
   If m_StreamCanRead Then
      For n = LBound(Values) To UBound(Values)
         Values(n) = ReadValue(VarType(Values(n)))
      Next
   End If

End Function



Public Function Read(Optional ByVal Length As Integer = 1) As String

   'This function reads the appropiate amount of characters from the
   'buffer and constucts the requested datatype
   
   If m_StreamCanRead Then
      Read = StreamRead(Length)
   Else
      Read = ""
   End If

End Function
Public Function ReadUntil(ByVal EndChar As String) As String

    Dim EndReached As Boolean, TextSegment As String, ch As String
   
    If m_StreamCanRead Then
        If EndChar = "" Then EndChar = Left(vbNewLine, 1)
        If Len(EndChar) > 1 Then EndChar = Left(EndChar, 1)
   
        EndReached = False
        Do Until EndReached Or Done()
            ch = Read(1)
            If StrComp(ch, EndChar, vbBinaryCompare) <> 0 Then
                TextSegment = TextSegment & ch
            Else
                EndReached = True
            End If
            DoEvents
        Loop
        ReadUntil = TextSegment
   
    Else
        ReadUntil = ""
    End If

End Function

Public Sub DiscardWhiteSpace()
   Dim Char As String, Complete As Boolean

   If m_StreamCanRead Then
       Do Until Complete Or Done()
           Char = NextCharacter()
           If Char <> "" Then
               If Char = " " Or (AscW(Char) >= 9 And AscW(Char) <= 13) Or AscW(Char) = 160 Then
                   Read                    'Remove it
               Else
                   Complete = True
               End If
           Else
               Complete = True
           End If
       Loop
   End If

End Sub




Private Function StreamRead(Optional ByVal Size As Integer = 1, Optional DiscardData As Boolean = True) As String

    'This function takes data from the buffer and creates a string of the
    'appropiate size.
      
    Dim StreamLeft As Long, Result As String, CurrentPos As Long
          
    If Size < 0 Then
        Size = 0
    End If
          
    Select Case m_StreamTarget
    Case StreamTargetFile
        If m_IsFileOpen Then       'Read from file
            '=== The total bytes left
            If VBA.EOF(m_FileNumber) Then
                StreamLeft = 0
            Else
                StreamLeft = LOF(m_FileNumber) - Seek(m_FileNumber) + 1
            End If
                       
            If DiscardData Then
                If Size >= StreamLeft Then                 'This read will finish the rest of the file
                    Result = Input(StreamLeft, m_FileNumber)
                    CloseFile
                Else
                    Result = Input(Size, m_FileNumber)
                End If
            Else
                CurrentPos = Seek(m_FileNumber)
                If Size >= StreamLeft Then                 'This read will finish the rest of the file
                    Result = Input(StreamLeft, m_FileNumber)
                Else
                    Result = Input(Size, m_FileNumber)
                End If
                                
                Seek m_FileNumber, CurrentPos        'Rewind 1 character
            End If
        Else
            Result = ""        'File closed!!!!
        End If
        
    Case StreamTargetString
        StreamLeft = Len(m_Buffer)
        
        If DiscardData Then
            If Size > StreamLeft Then
                Result = m_Buffer    'Rest of Buffer
                m_Buffer = ""         'Kill rest of buffer
            Else
                Result = Left(m_Buffer, Size)      'Rest of Buffer is beginning
                m_Buffer = Mid(m_Buffer, Size + 1)
            End If
         End If
    End Select
    
    StreamRead = Result
End Function


Public Function NextCharacter() As String
    'Get the next character in the stream, but DO NOT READ IT!
'    Dim Result As String
'
'    Select Case m_StreamTarget
'    Case StreamTargetFile
'        If m_IsFileOpen Then       'Read from file
'            If VBA.EOF(m_FileNumber) Then
'                Result = ""
'            Else
'                Result = ReadChar(False)
'            End If
'        Else
'            Result = ""        'File closed!!!!
'        End If
'
'    Case StreamTargetString
'        Result = Left(m_Buffer, 1)
'
'    End Select
'
'    NextCharacter = Result

    NextCharacter = ReadChar(False)
End Function


Property Get Text() As String
Attribute Text.VB_UserMemId = 0
   
   If m_StreamTarget = StreamTargetString Then
      Text = m_Buffer
   Else
      Text = ""
   End If

End Property

Property Let Text(NewString As String)

   If m_StreamTarget = StreamTargetString Then
       m_Buffer = NewString
   End If

End Property

Private Sub Class_Initialize()

    m_Buffer = ""
    m_StreamTransfer = StreamTransferText
    m_StreamCanRead = True
    m_StreamTarget = StreamTargetString
    m_StringType = StreamStringTypeASCII
    m_StringForm = StreamStringFormNull
    m_Encoding = StreamEncodingLittleEndian
End Sub
Public Function ReadLine() As String

    Dim EndReached As Boolean, Text As String, ch As String
   
    If m_StreamCanRead Then
        EndReached = False
        Do Until EndReached Or Done()
            ch = ReadChar()
            If ch = Chr(10) Or ch = Chr(13) Then         'End char
                ch = NextCharacter()
                If ch = Chr(10) Or ch = Chr(13) Then     'Discard second of line-feed, carriage return pair
                    ReadChar
                End If
                EndReached = True
            Else
                Text = Text & ch
            End If
        Loop
        ReadLine = Text
    Else
        ReadLine = ""
    End If

End Function

Private Sub Class_Terminate()

    If m_IsFileOpen Then
       CloseFile
    End If

End Sub


