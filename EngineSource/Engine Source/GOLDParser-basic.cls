VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GOLDParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================
' Class Name:
'    GOLDParser (basic version)
'
' Instancing:
'      Public; Creatable  (VB Setting: 5 - MultiUse)
'
' Purpose:
'   This is the main class in the GOLD Parser Engine and is used to
'   perform all duties required to the parsing of a source text
'   string. This class contains the LALR(1) State Machine code,
'   the DFA State Machine code, character table (used by the DFA
'   algorithm) and all other structures and methods needed to
'   interact with the developer.
'
'Author(s):
'   Devin Cook
'
'Public Dependencies:
'   Token, Rule, Symbol, Reduction
'
'Private Dependencies:
'   ObjectArray, SimpleDatabase, SymbolList, TokenStack
'
'Revision History:
'   June 9, 2001:
'      Added the ReductionMode property and modified the Reduction object (which was
'      used only for internal use). In addition the Reduction property was renamed to
'      CurrentReduction to avoid possible name conflicts in different programming languages
'      (which this VB source will be converted to eventually)
'
'   Sept 5, 2001:
'      I was alerted to an error in the engine logic by Szczepan Holyszewski [rulatir@poczta.arena.pl].
'      When reading tokens inside a block quote, the line-comment token would still eliminate the rest
'      of a line - possibly eliminating the block quote end.
'
'   Nov 28, 2001:
'      Fixed several errors.
'
'   December 2001:
'      Added the TrimReductions property and required logic
'
'   March 7, 2004:
'      Changed the logic in the DFA to use a local Lookahead buffer. Before, the system was using
'      a property of the Stream class. This change is designed to simply the logic and
'      allow a normal stream class to be used.
'
'   June 6, 2004:
'      Fixed an Engine bug concerning line comments. When the Engine encounters a line
'      comment, text is discarded until either a new line or the end of file is encountered.
'      The new line characters are not discarded. The Engine was discarding the first
'      character of the carriage return / line feed sequence which caused problems with
'      line-based grammars. Special thanks to Dean McNamee for finding this bug.
'
'
'   July 8, 2004:
'      Added the ability to return comments to the developer. Two new gpMessageConstants
'      were added that indicate that either a line comment or block comment was read.
'      The first, gpMsgCommentBlockRead, is returned anytime a block comment is complete.
'      The second, gpMsgCommentLineRead, is returned for line comments. When the system
'      is reading a block comment, line comments are are completely ignored.
'      When either of these two messages are returned, the content of the CurrentComment
'      property is set to the text.
'
'===================================================================

'===================================================================
'
'             The GOLD Parser Freeware License Agreement
'             ==========================================
'
'this software Is provided 'as-is', without any expressed or
'implied warranty. In no event will the authors be held liable for any
'damages arising from the use of this software.
'
'Permission is granted to anyone to use this software for any
'purpose. If you use this software in a product, an acknowledgment
'in the product documentation would be deeply appreciated but is
'not required.
'
'In the case of the GOLD Parser Engine source code, permission is
'granted to anyone to alter it and redistribute it freely, subject
'to the following restrictions:
'
'   1. The origin of this software must not be misrepresented; you
'      must not claim that you wrote the original software.
'
'   2. Altered source versions must be plainly marked as such, and
'      must not be misrepresented as being the original software.
'
'   3. This notice may not be removed or altered from any source
'      distribution
'
'====================================================================
Option Explicit

'================================== Symbols recognized by the system
Private pSymbolTable As New ObjectArray

'================================== DFA. Contains FAStates.
Private pDFA As New ObjectArray                       'FAState

Private pCharacterSetTable As New ObjectArray

'================================== Rules. Contains Rule Objects.
Private pRuleTable As New ObjectArray

'================================== LALR(1) action table. Contains LRActionTables.
'This is different from LALR in the Builder which contains LRStates
Private pActionTable As New ObjectArray


'================================== Parsing messages
Public Enum GPMessageConstants
   gpMsgTokenRead = 1               'A new token is read
   gpMsgReduction = 2               'A rule is reduced
   gpMsgAccept = 3                  'Grammar complete
   gpMsgNotLoadedError = 4          'Now grammar is loaded
   gpMsgLexicalError = 5            'Token not recognized
   gpMsgSyntaxError = 6             'Token is not expected
   gpMsgCommentError = 7            'Reached the end of the file - mostly due to being stuck in comment mode
   gpMsgInternalError = 8           'Something is wrong, very wrong
   gpMsgCommentBlockRead = 9        'Added July 8, 2004
   gpMsgCommentLineRead = 10        'Added July 8, 2004
End Enum
  
'========================================= DFA runtime constants
Private kErrorSymbol As Symbol
Private kEndSymbol As Symbol

'========================================= DFA runtime variables
Private pInitialDFAState As Integer
Private pLookaheadBuffer As String        'Added 3/7/04

'========================================= LALR runtime variables
Private pInitialLALRState As Integer
Private pStartSymbol As Long
Private CurrentLALR As Long
Private Stack As New TokenStack

'===================== Used for Reductions & Errors
'The set of tokens for 1. Expecting during error, 2. Reduction
Private pTokens         As New TokenStack
Private pHaveReduction  As Boolean
Private pTrimReductions   As Boolean

'===================== Private control variables
Private pTablesLoaded     As Boolean
Private pInputTokens      As New TokenStack  'Stack of tokens to be analyzed
Private pSource           As New Stream

Private pCommentLevel     As Integer         'Current level of block comments (1+)
Private pCurrentComment   As String          'Added June 06/29/04

'=== Internal column and line number. This is not reported to the user
Private pSysLineNumber        As Long                  'Incremented by tokenizer
Private pSysColumnNumber      As Long                  'Set by tokenizer

'=== Current line and column information. This is stored for the benefit of the developer
Private pCurrentLineNumber       As Long
Private pCurrentColumnNumber     As Long

Private Enum ParseResultConstants
    ParseResultAccept = 1
    ParseResultShift = 2
    ParseResultReduceNormal = 3
    ParseResultReduceEliminated = 4
    ParseResultSyntaxError = 5
    ParseResultInternalError = 6
End Enum


'============================================================
Private Const RecordIdParameters  As Byte = 80   'P
Private Const RecordIdTableCounts As Byte = 84   'T
Private Const RecordIdInitial     As Byte = 73   'I
Private Const RecordIdSymbols     As Byte = 83   'S
Private Const RecordIdCharSets    As Byte = 67   'C
Private Const RecordIdRules       As Byte = 82   'R
Private Const RecordIdDFAStates   As Byte = 68   'D
Private Const RecordIdLRTables    As Byte = 76   'L
Private Const RecordIdComment     As Byte = 33   '!
Private Const FileHeader = "GOLD Parser Tables/v1.0"


'================================== Parameters
Private pParameterName          As String
Private pParameterAuthor        As String
Private pParameterVersion       As String
Private pParameterAbout         As String
Private pParameterStartSymbol   As String
Private pParameterCaseSensitive As String



  
Public Property Get CurrentComment() As String
    CurrentComment = pCurrentComment
End Property

Public Function CurrentLineNumber() As Long

   CurrentLineNumber = pCurrentLineNumber

End Function

Public Function CurrentColumnNumber() As Long

   CurrentColumnNumber = pCurrentColumnNumber

End Function


Public Sub CloseFile()

   pSource.CloseFile

End Sub

Public Function CurrentToken() As Token

   Set CurrentToken = pInputTokens.Top

End Function


Private Function DiscardRestOfLine() As String
    Dim EndReached As Boolean, Position As Long
      
    'Lookahead in the stream (we have buffer and we have to use it)
    'and find the next chr(10) or chr(13). Then discard the lookahead
    'until that point.
   
    Position = 1
    EndReached = False
    
    Do Until EndReached Or pSource.Done()
        Select Case Lookahead(Position)
        Case Chr(10), Chr(13), ""
            EndReached = True
        Case Else
            Position = Position + 1
        End Select
    Loop

    'Updated 6/1/04: The value read from the buffer is now (position - 1). It was incorrect before.
    'The Position variable cannot be less than 1
    
    DiscardRestOfLine = ReadBuffer(Position - 1)  'The value returned from this function is disgarded
End Function

Public Function PopInputToken() As Token

   Set PopInputToken = pInputTokens.Pop

End Function

Private Sub PrepareToParse()
   Dim Start As New Token
   
   Start.State = pInitialLALRState
   Set Start.ParentSymbol = pSymbolTable.Member(pStartSymbol)
   
   Stack.Push Start

End Sub

Public Sub PushInputToken(TheToken As Token)

   pInputTokens.Push TheToken

End Sub

Public Property Get CurrentReduction() As Object
    If pHaveReduction Then
        Set CurrentReduction = Stack.Top.Data
    Else
        Set CurrentReduction = Nothing
    End If
End Property


Public Property Set CurrentReduction(Value As Object)
    If pHaveReduction Then
        Set Stack.Top.Data = Value
    End If
End Property

Private Function ReadBuffer(ByVal CharCount As Long) As String
    If CharCount <= Len(pLookaheadBuffer) Then
        '=== Remove the characters from the front of the buffer. This code will
        '=== be very different in other programming languages
        
        ReadBuffer = Left(pLookaheadBuffer, CharCount)
        pLookaheadBuffer = Mid(pLookaheadBuffer, CharCount + 1)  'Remove chars
        
    Else
        'ERROR - DFA LOGIC DOES NOT ALLOW THIS!
        ReadBuffer = ""
    End If
End Function

Private Function Lookahead(ByVal CharIndex As Long) As String
    Dim NewChars As String, ReadCount As Long
    
    If CharIndex > Len(pLookaheadBuffer) Then
        '=== We must read characters from the Stream
        ReadCount = CharIndex - Len(pLookaheadBuffer)
                               
        pLookaheadBuffer = pLookaheadBuffer & pSource.Read(ReadCount)
    End If
           
           
    '=== If the buffer is still smaller than the index, we have reached
    '=== the end of the text. In this case, return a null string - the DFA
    '=== code will understand
    
    If CharIndex >= Len(pLookaheadBuffer) Then
        Lookahead = Mid(pLookaheadBuffer, CharIndex, 1)
    Else
        Lookahead = ""
    End If
End Function


Public Sub ShowAboutWindow()

   MsgBox "GOLD Parser Engine" & vbNewLine & App.Major & "." & App.Minor & "." & App.Revision

End Sub


Public Sub Clear()
   pSymbolTable.Clear
   pRuleTable.Clear
   pCharacterSetTable.Clear
   pTokens.Clear
   pInputTokens.Clear
   pActionTable.Clear
   
   pParameterName = ""
   pParameterVersion = ""
   pParameterAuthor = ""
   pParameterAbout = ""
   pParameterStartSymbol = ""
   
   Reset
End Sub


Public Property Let TrimReductions(Value As Boolean)
    pTrimReductions = Value
End Property

Public Property Get TrimReductions() As Boolean
    TrimReductions = pTrimReductions
End Property

Public Property Get Parameter(ByVal Name As String) As String
   
   Select Case UCase(Name)
   Case "NAME"
       Parameter = pParameterName
   Case "VERSION"
       Parameter = pParameterVersion
   Case "AUTHOR"
       Parameter = pParameterAuthor
   Case "ABOUT"
       Parameter = pParameterAbout
   Case "START SYMBOL"
       Parameter = pParameterStartSymbol
   Case "CASE SENSITIVE"
       Parameter = pParameterCaseSensitive
   Case Else
       Parameter = ""
   End Select
   
End Property




Private Function LoadTables(FileName As String) As Boolean
   On Error GoTo Problem
      
   Dim File As New CGTReader, ID As Integer, bAccept As Boolean
   Dim n As Integer, SetIndex As Integer, Target As Integer
   Dim ReadSymbol As Symbol, ReadRule As Rule
   Dim ReadDFA As FAState, ReadLALR As LRActionTable
   Dim Success As Boolean, Text As String
   
   If File.OpenFile(FileName, "R") Then
       Success = True
       File.GetNextRecord
       
       Do Until File.EOF() Or Success = False
               
           Select Case File.RecordType
           Case CGTRecTypeParameters
               'Name, Version, Author, About, Case-Sensitive
              
               pParameterName = File.RetrieveNext
               pParameterVersion = File.RetrieveNext
               pParameterAuthor = File.RetrieveNext
               pParameterAbout = File.RetrieveNext
               pParameterCaseSensitive = CStr(File.RetrieveNext)
               pStartSymbol = Val(File.RetrieveNext)
              
           Case CGTRecTypeCounts
               'Symbol, CharacterSet, Rule, DFA, LALR
           
               pSymbolTable.ReDimension Val(File.RetrieveNext)
               pCharacterSetTable.ReDimension Val(File.RetrieveNext)
               pRuleTable.ReDimension Val(File.RetrieveNext)
               pDFA.ReDimension Val(File.RetrieveNext)
               pActionTable.ReDimension Val(File.RetrieveNext)
            
           Case CGTRecTypeInitial
               'DFA, LALR
        
               pInitialDFAState = File.RetrieveNext
               pInitialLALRState = File.RetrieveNext
                             
           Case CGTRecTypeSymbol
               '#, Name, Kind
                              
               Set ReadSymbol = New Symbol
               n = File.RetrieveNext
               ReadSymbol.Name = File.RetrieveNext
               ReadSymbol.Kind = File.RetrieveNext
               File.RetrieveNext                      'Empty
                                       
               ReadSymbol.TableIndex = n
               Set pSymbolTable.Member(n) = ReadSymbol
                                  
           Case CGTRecTypeCharSet
               '#, Characters
                              
               n = File.RetrieveNext
               Set pCharacterSetTable.Member(n) = ToNumberSet(File.RetrieveNext)
                             
           Case CGTRecTypeRule
               '#, ID#, Reserved, (Symbol#,  ...)
                    
               Set ReadRule = New Rule
               n = File.RetrieveNext
               ReadRule.TableIndex = n
               ReadRule.SetRuleNonterminal pSymbolTable.Member(Val(File.RetrieveNext))
               File.RetrieveNext                                                'Reserved
               Do Until File.RetrieveDone
                   ReadRule.AddItem pSymbolTable.Member(Val(File.RetrieveNext()))
               Loop
               Set pRuleTable.Member(n) = ReadRule
                                          
           Case CGTRecTypeDFAState
               '#, Accept?, Accept#, Reserved (CharSet#, Target#, Reserved)...
                           
               Set ReadDFA = New FAState
               n = File.RetrieveNext
               bAccept = File.RetrieveNext
                             
               If bAccept Then
                   ReadDFA.AcceptSymbol = File.RetrieveNext
               Else
                   ReadDFA.AcceptSymbol = -1
                   File.RetrieveNext                'Discard value
               End If
               File.RetrieveNext         'Reserved                                       'Reserved
            
               '(Edge chars, Target#, Reserved)...
               Do Until File.RetrieveDone
                   SetIndex = File.RetrieveNext   'Char table index
                   Target = File.RetrieveNext   'Target
                   ReadDFA.AddEdge pCharacterSetTable.Member(SetIndex), Target, SetIndex
                   File.RetrieveNext                                            'Reserved
               Loop
               Set pDFA.Member(n) = ReadDFA
                     
           Case CGTRecTypeLRState
               '#, Reserved (Symbol#, Action, Target#, Reserved)...
                          
               Set ReadLALR = New LRActionTable
               n = File.RetrieveNext
               File.RetrieveNext                                                'Reserved
               Do Until File.RetrieveDone
                   ReadLALR.AddItem pSymbolTable.Member(File.RetrieveNext), File.RetrieveNext, File.RetrieveNext
                   File.RetrieveNext                                            'Reserved
               Loop
               Set pActionTable.Member(n) = ReadLALR
                                       
           Case Else       'RecordIDComment
               Success = False
           End Select
           
           DoEvents
           File.GetNextRecord
       Loop
             
       pParameterStartSymbol = pSymbolTable.Member(pStartSymbol).Name
       File.CloseFile
       LoadTables = Success
   Else
       LoadTables = False
   End If

   Exit Function
   
Problem:
   LoadTables = False
End Function

Private Function ToNumberSet(ByVal Text As String) As NumberSet
    Dim Result As New NumberSet
    Dim n As Long
    
    For n = 1 To Len(Text)
        Result.Add AscW(Mid(Text, n, 1))
    Next
            
    Set ToNumberSet = Result
End Function

Public Function SymbolTableCount() As Integer

   SymbolTableCount = pSymbolTable.Count

End Function

Public Function RuleTableCount() As Integer

   RuleTableCount = pRuleTable.Count

End Function

Public Function SymbolTableEntry(ByVal Index As Integer) As Symbol

   If Index >= 0 And Index < pSymbolTable.Count Then
      Set SymbolTableEntry = pSymbolTable.Member(Index)
   End If

End Function

Public Function RuleTableEntry(ByVal Index As Integer) As Rule

   If Index >= 0 And Index < pRuleTable.Count Then
      Set RuleTableEntry = pRuleTable.Member(Index)
   End If

End Function


Public Function TokenCount() As Integer

   TokenCount = pTokens.Count

End Function


Public Function Tokens(ByVal Index As Integer) As Token

   If Index >= 0 And Index < pTokens.Count Then
      Set Tokens = pTokens.Member(Index)
   Else
      Set Tokens = Nothing
   End If

End Function

Public Function LoadCompiledGrammar(FileName As String) As Boolean
   Reset
   LoadCompiledGrammar = LoadTables(FileName)
End Function
Public Function OpenTextString(Text As String) As Boolean
   
   Reset
   pSource.Text = Text
   PrepareToParse
   OpenTextString = True

End Function
Public Function Parse() As GPMessageConstants
    '1. If the tables are not setup then report GPM_NotLoadedError
    '2. If parser is in comment mode then read tokens until a
    '   recognized one is found and report it
    '3. Otherwise, parser normal
    '   a. If there are no tokens on the stack
    '      1) Read one and trap error
    '      2) End function with GPM_TokenRead
    '   b. Otherwise, call ParseToken with the top of the stack.
    '      1) If success, then Pop the value
    '      2) Loop if the token was shifted (nothing to report)

    Dim Result As GPMessageConstants, Done As Boolean
    Dim ReadToken As Token, ParseResult As ParseResultConstants
       
    If pActionTable.Count < 1 Or pDFA.Count < 1 Then
        Result = gpMsgNotLoadedError
    Else
        Done = False
        Do Until Done
            If pInputTokens.Count = 0 Then                     'We must read a token
                Set ReadToken = RetrieveToken(pSource)
                If ReadToken Is Nothing Then
                    Result = gpMsgInternalError
                    Done = True
                Else
                    pInputTokens.Push ReadToken
                    If pCommentLevel = 0 And ReadToken.Kind <> SymbolTypeCommentLine And ReadToken.Kind <> SymbolTypeCommentStart And ReadToken.Kind <> SymbolTypeWhitespace Then
                        Result = gpMsgTokenRead           'Report that a token was read
                        Done = True
                    End If
                End If

            ElseIf pCommentLevel > 0 Then           'We are in a block comment
                Set ReadToken = pInputTokens.Pop()
                
                pCurrentComment = pCurrentComment & ReadToken.Data   'Build comment text. Added 06/29/04

                Select Case ReadToken.Kind
                Case SymbolTypeCommentStart
                    pCommentLevel = pCommentLevel + 1
                Case SymbolTypeCommentEnd
                    pCommentLevel = pCommentLevel - 1
                    If pCommentLevel = 0 Then    'Done with comment - report it. Added 06/29/04
                        Result = gpMsgCommentBlockRead
                        Done = True
                    End If
                    
                Case SymbolTypeEnd
                    Result = gpMsgCommentError
                    Done = True
                Case Else
                    'Do nothing, ignore
                    'The 'comment line' symbol is ignored as well
                End Select
            Else
                '==== Normal parse mode - we have a token and we are not in comment mode
                Set ReadToken = pInputTokens.Top

                Select Case ReadToken.Kind
                Case SymbolTypeWhitespace
                    pInputTokens.Pop                           'Disgard whitespace
                Case SymbolTypeCommentStart
                    pCommentLevel = 1
                    pCurrentComment = ReadToken.Data              'Start with the Block Start token
                    pInputTokens.Pop                           'Remove it
                Case SymbolTypeCommentLine
                    pInputTokens.Pop                           'Remove it and rest of line
                    pCurrentComment = ReadToken.Data & DiscardRestOfLine()          'Procedure also increments the line number
                    Result = gpMsgCommentLineRead
                    Done = True
                Case SymbolTypeError
                    Result = gpMsgLexicalError
                    Done = True
                Case Else                                      'FINALLY, we can parse the token
                    ParseResult = ParseToken(ReadToken)
                    'NEW 12/2001: Now we are using the internal enumerated constant
                    Select Case ParseResult
                    Case ParseResultAccept
                        Result = gpMsgAccept
                        Done = True
                    Case ParseResultInternalError
                        Result = gpMsgInternalError
                        Done = True
                    Case ParseResultReduceNormal
                        Result = gpMsgReduction
                        Done = True
                    Case ParseResultShift                      'A simple shift, we must continue
                        pInputTokens.Pop                       'Okay, remove the top token, it is on the stack
                    Case ParseResultSyntaxError
                        Result = gpMsgSyntaxError
                        Done = True
                    Case Else
                        'Do nothing
                    End Select

                End Select
            End If
        Loop
        
    End If

    
    Parse = Result

End Function
Private Function ParseToken(NextToken As Token) As ParseResultConstants
   'This function analyzes a token and either:
   '  1. Makes a SINGLE reduction and pushes a complete Reduction
   '     object on the stack
   '  2. Accepts the token and shifts
   '  3. Errors and places the expected symbol indexes in the
   '     Tokens list. The Token is assumed to be valid and WILL
   '     be checked. If an action is performed that requires control
   '     to be returned to the user, the function returns true.
   '     The Message parameter is then set to the type of action.
   
   Dim n As Integer, Found As Boolean, Index As Integer, RuleIndex As Integer, CurrentRule As Rule
   Dim str As String, Head As Token, NewReduction As Reduction
   Dim Result As ParseResultConstants
   
      
   Index = pActionTable.Member(CurrentLALR).ActionIndexForSymbol(NextToken.ParentSymbol.TableIndex)
         
   If Index <> -1 Then              'Work - shift or reduce
       pHaveReduction = False       'Will be set true if a reduction is made
       pTokens.Count = 0
   
       Select Case pActionTable.Member(CurrentLALR).Item(Index).Action
       Case ActionAccept
           pHaveReduction = True
           Result = ParseResultAccept
         
       Case ActionShift
           CurrentLALR = pActionTable.Member(CurrentLALR).Item(Index).Value
           NextToken.State = CurrentLALR
           Stack.Push NextToken
           Result = ParseResultShift
   
       Case ActionReduce
           'Produce a reduction - remove as many tokens as members in the rule & push a nonterminal token
           
           RuleIndex = pActionTable.Member(CurrentLALR).Item(Index).Value
           Set CurrentRule = pRuleTable.Member(RuleIndex)
           
           '======== Create Reduction
           If pTrimReductions And CurrentRule.ContainsOneNonTerminal Then
               'The current rule only consists of a single nonterminal
               'and can be trimmed from the parse tree. Usually we
               'create a new Reduction, assign it to the Data property
               'of Head and push it on the stack. However, in this case,
               'the Data property of the Head will be assigned the Data
               'property of the reduced token (i.e. the only one'on the
               'stack). In this case, to save code, the value popped of
               'the stack is changed into the head.
               
               Set Head = Stack.Pop()
               Set Head.ParentSymbol = CurrentRule.RuleNonterminal
               
               Result = ParseResultReduceEliminated
           Else                                           'Build a Reduction
               pHaveReduction = True
               Set NewReduction = New Reduction
               With NewReduction
                   Set .ParentRule = CurrentRule
                   .TokenCount = CurrentRule.SymbolCount
                   For n = .TokenCount - 1 To 0 Step -1
                       Set .Tokens(n) = Stack.Pop()
                   Next
               End With
                    
               Set Head = New Token
               Set Head.Data = NewReduction
               Set Head.ParentSymbol = CurrentRule.RuleNonterminal
               
               Result = ParseResultReduceNormal
           End If
                    
           'Goto
           Index = Stack.Top().State
                
           'If n is -1 here, then we have an Internal Table Error!!!!
           n = pActionTable.Member(Index).ActionIndexForSymbol(CurrentRule.RuleNonterminal.TableIndex)
           If n <> -1 Then
               CurrentLALR = pActionTable.Member(Index).Item(n).Value

               Head.State = CurrentLALR
               Stack.Push Head
           Else
               Result = ParseResultInternalError
           End If
       End Select
   
   Else
       '=== Syntax Error! Fill Expected Tokens
       pTokens.Clear
       For n = 0 To pActionTable.Member(CurrentLALR).Count - 1
           '01/26/2002: Fixed bug. EOF was not being added to the expected tokens
           Select Case pActionTable.Member(CurrentLALR).Item(n).Symbol.Kind
           Case SymbolTypeTerminal, SymbolTypeEnd
               Set Head = New Token
               Head.Data = ""
               Set Head.ParentSymbol = pActionTable.Member(CurrentLALR).Item(n).Symbol
               pTokens.Push Head
           End Select
       Next
       'If pTokens.Count = 0 Then Stop
       Result = ParseResultSyntaxError
   End If
      
   ParseToken = Result     'Very important
      
End Function

Public Function OpenFile(ByVal FileName As String, Optional ByVal DetectEncodingFromByteOrderMarks = True) As Boolean
    Dim Success As Boolean
    Reset
                
    Success = pSource.OpenFile(FileName, StreamTransferText, DetectEncodingFromByteOrderMarks) ',  StreamType)
        
    PrepareToParse
    OpenFile = Success
End Function

Public Sub Reset()
   On Error GoTo Problem
   
   Dim n As Integer
   
   '================= Setup global varibles
   For n = 0 To pSymbolTable.Count - 1
      Select Case pSymbolTable.Member(n).Kind
      Case SymbolTypeError
         Set kErrorSymbol = pSymbolTable.Member(n)
      Case SymbolTypeEnd
         Set kEndSymbol = pSymbolTable.Member(n)
      End Select
   Next
   
   CurrentLALR = pInitialLALRState
   
   pSysLineNumber = 1
   pSysColumnNumber = 1
   
   pCurrentLineNumber = 1
   pCurrentColumnNumber = 1
      
   pSource.CloseFile
   pCommentLevel = 0
   pHaveReduction = False
   
   pTokens.Clear
   pInputTokens.Clear
   Stack.Clear
   pLookaheadBuffer = ""
   
   Exit Sub

Problem:
   Err.Clear
End Sub

Private Function RetrieveToken(Source As Stream) As Token
    'THIS IS THE TOKENIZER!
    '
    'This function implements the DFA algorithm and returns
    'a token to the LALR state machine
    
    Dim ch As String, n As Integer, Found As Boolean, Done As Boolean, Target As Integer
    Dim CharSetIndex As Integer, CurrentDFA As Integer, CurrentPosition As Long
    Dim LastAcceptState As Integer, LastAcceptPosition As Integer
    Dim Result As New Token
    
    Done = False
    CurrentDFA = pInitialDFAState  'The first state is almost always #1.
    CurrentPosition = 1    'Next byte in the input Stream
    LastAcceptState = -1   'We have not yet accepted a character string
    LastAcceptPosition = -1
    
    If Lookahead(1) <> "" Then     'NO MORE DATA
        Do Until Done
            ' This code searches all the branches of the current DFA
            ' state for the next character in the input Stream. If
            ' found the target state is returned. The InStr()
            ' function searches the string
            ' pCharacterSetTable.Member(CharSetIndex) starting at
            ' position 1 for ch.  The pCompareMode variable determines
            ' whether the search is case sensitive.
          
            ch = Lookahead(CurrentPosition)
            If ch = "" Then         'End reached, do not match
                Found = False
            Else
                n = 0
                Found = False
                Do While n < pDFA.Member(CurrentDFA).EdgeCount And Not Found
                    CharSetIndex = pDFA.Member(CurrentDFA).Edge(n).SetIndex
                    If pCharacterSetTable.Member(CharSetIndex).HasMember(AscW(ch)) Then
                        Found = True
                        Target = pDFA.Member(CurrentDFA).Edge(n).Target '.TableIndex
                    End If
                    n = n + 1
                Loop
            End If
          
            ' This block-if statement checks whether an edge was found
            ' from the current state. If so, the state and current
            ' position advance. Otherwise it is time to exit the main
            ' loop and report the token found (if there was it fact one).
            ' If the LastAcceptState is -1, then we never found a match
            ' and the Error Token is created. Otherwise, a new token
            ' is created using the Symbol in the Accept State and all
            ' the characters that comprise it.
          
            If Found Then
                ' This code checks whether the target state accepts a
                ' token. If so, it sets the appropiate variables so when
                ' the algorithm in done, it can return the proper
                ' token and number of characters.
              
                If pDFA.Member(Target).AcceptSymbol <> -1 Then
                    LastAcceptState = Target
                    LastAcceptPosition = CurrentPosition
                End If
                   
                CurrentDFA = Target
                CurrentPosition = CurrentPosition + 1
            Else                                           'No edge found
                Done = True
                If LastAcceptState = -1 Then
                    'Tokenizer cannot recognize symbol
                    Set Result.ParentSymbol = kErrorSymbol
                    Result.Data = ReadBuffer(1)
                Else
                    'Create Token, read characters
                    'The data contains the total number of accept characters
                 
                    Set Result.ParentSymbol = pSymbolTable.Member(pDFA.Member(LastAcceptState).AcceptSymbol)
                    Result.Data = ReadBuffer(LastAcceptPosition)
                End If
            End If
            DoEvents
        Loop
    Else
        Result.Data = ""  'End of file reached, create End Token
        Set Result.ParentSymbol = kEndSymbol
    End If
   
   
    ' Count Carriage Returns and increment the Line Number. This is
    ' done for the Developer and is not necessary for the
    ' DFA algorithm
      
    pCurrentColumnNumber = pSysColumnNumber
    pCurrentLineNumber = pSysLineNumber
    
    ' Advance system counters
    For n = 1 To Len(Result.Data)
        Select Case Mid(Result.Data, n, 1)
        Case vbCr
            pSysLineNumber = pSysLineNumber + 1
            pSysColumnNumber = 1
        Case vbLf
            'Ignore
        Case Else
            pSysColumnNumber = pSysColumnNumber + 1
        End Select
    Next

   Set RetrieveToken = Result
End Function

Private Sub Class_Initialize()
    Reset
    pTablesLoaded = False
    
    pTrimReductions = True
End Sub


Private Sub Class_Terminate()
    Set pSymbolTable = Nothing
    Set pDFA = Nothing
    Set pCharacterSetTable = Nothing
    Set pRuleTable = Nothing
    Set pActionTable = Nothing
    
    Set kErrorSymbol = Nothing
    Set kEndSymbol = Nothing
    
    Set Stack = Nothing
    Set pTokens = Nothing
    
    Set pInputTokens = Nothing
    Set pSource = Nothing
End Sub


