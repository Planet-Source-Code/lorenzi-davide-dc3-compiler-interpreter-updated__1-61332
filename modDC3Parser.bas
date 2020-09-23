Attribute VB_Name = "modDC3Parser"
Option Explicit

Public moParser As New GOLDParser

'Parserizza il file
Public Function DoParseFile(sSourceFile As String, lstLog As ListBox) As cReduction
    Dim sTmp As String
    sTmp = FileToString(sSourceFile)

    Set DoParseFile = DoParse(sTmp, lstLog)
End Function

'Rende l'albero parserizzato
Public Function DoParse(ByVal sSource As String, lstLog As ListBox) As cReduction
    
    'Cosi' evito il problema del NewLine finale
    sSource = sSource & vbCrLf
    
    'This procedure starts the GOLD Parser Engine and handles each of the
    'messages it returns. Each time a reduction is made, a new custom object
    'can be created and used to store the rule. Otherwise, the system will use
    'the Reduction object that was returned.
    '
    'The resulting tree will be a pure representation of the language
    'and will be ready to implement.
   
    Dim Response As GPMessageConstants
    Dim Done As Boolean, Success As Boolean       'Controls when we leave the loop
               
    Success = False    'Unless the program is accepted by the parser
          
    With moParser
      .OpenTextString sSource
      .TrimReductions = False  'Please read about this feature before enabling
                     
      Done = False
      Do Until Done
         Response = .Parse()
                  
         Select Case Response
         Case gpMsgLexicalError
            'Cannot recognize token
            Log "LEXICAL ERROR. Line " & .CurrentLineNumber & ". Cannot recognize token: " & moParser.CurrentToken.Data, lstLog
            Done = True
                     
         Case gpMsgSyntaxError
            'Expecting a different token
            Log "SYNTAX ERROR. Line:" & .CurrentLineNumber & ", Column:" & .CurrentColumnNumber & ".", lstLog
            
            On Error Resume Next
            Log "Expecting expression like: " & moParser.CurrentReduction.oData.ParentRule.Text, lstLog
            On Error GoTo 0

            Dim sText As String
            Dim n As Integer
            For n = 0 To .TokenCount - 1
                sText = sText & moParser.Tokens(n).Name
                If n < .TokenCount - 1 Then sText = sText & " "
            Next
            
            Log "Expecting word like: " & LTrim(sText) & ".", lstLog
            Done = True
             
         Case gpMsgReduction
            'Non fa nulla, i controlli vengono fatti nella classe cProgram
            Set .CurrentReduction = NewReductionObject(.CurrentReduction)
            
            
         Case gpMsgAccept
            'Success!
            Set DoParse = .CurrentReduction  'The root node!
            Done = True
            Success = True
                  
         Case gpMsgTokenRead
            'You don't have to do anything here.
                   
         Case gpMsgInternalError
            'INTERNAL ERROR! Something is horribly wrong.
            Log "INTERNAL ERROR! Something is horribly wrong", lstLog
            Done = True
                   
         Case gpMsgNotLoadedError
            'Due to the if-statement above, this case statement should never be true
            Log "NOT LOADED ERROR! Compiled Grammar Table not loaded", lstLog
            Done = True
                 
         Case gpMsgCommentError
            'COMMENT ERROR! Unexpected end of file
            Log "COMMENT ERROR! Unexpected end of file", lstLog
            Done = True
         End Select
      Loop
    End With
    
End Function

'Crea la nuova struttura con le info aggiuntive
Private Function NewReductionObject(oParserRed As Reduction) As cReduction
    Dim oReduction As New cReduction
    
    With oReduction
        .lRow = moParser.CurrentLineNumber
        .lCol = moParser.CurrentColumnNumber
        Set .oData = oParserRed
    End With
    
    Set NewReductionObject = oReduction
End Function

