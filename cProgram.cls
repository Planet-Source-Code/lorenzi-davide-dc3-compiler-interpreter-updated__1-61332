VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================
' Part of DC3 Compiler - Interpreter
' Author: Lorenzi Davide (http://www.hexagora.com)
' See the file 'license.txt' for informations
'================================================================================
'
'Esegue la generazione del programma
'
'================================================================================

Option Explicit

'
'Variabili di gestione
'
Private moSymbTable As cSymbTable
Private moCurrSymbTable As cSymbTable 'Puntatore alla SymbTable Corrente
Private moCompError As cCompilerError
Private mbAllowResAccMod As Boolean
Private moByteCode As cByteCode
Private moResSymb As cResSymbols

'
'//////////////////////////////////////////////////////////////////////////////
'

'Return Byte Code
Public Function GetByteCode() As cByteCode
    Set GetByteCode = moByteCode
End Function

'Return Symbol Table
Public Function GetSymbTable() As cSymbTable
    Set GetSymbTable = moSymbTable
End Function

'
'//////////////////////////////////////////////////////////////////////////////
'
Public Function Compile(sSource As String, lstLog As ListBox) As Boolean
    On Error GoTo Errore
    
    'Crea la tabella dei simboli per il programma
    Set moSymbTable = New cSymbTable
    Set moCurrSymbTable = moSymbTable
    Set moCompError = New cCompilerError
    Set moByteCode = New cByteCode
    Set moResSymb = New cResSymbols
    
    lstLog.Clear
    
    'La prima istruzione del programma e' sempre NDATA xche' mi permette di
    'allocare lo spazio per le variabili
    'Questa istruzione verra' poi modificata alla fine della compilazione
    Dim lBCIndexVars As Long
    lBCIndexVars = moByteCode.Count
    moByteCode.Add_INData 0, 1
    
    'Prima compila le parole riservate
    Dim oRes As cReduction
    Set oRes = DoParseFile(App.Path & "\dc3res.vb", lstLog)
    
    If Not oRes Is Nothing Then
        'Solo io posso dichiarare cose riservate ;-)
        mbAllowResAccMod = True
        pCompile oRes
        mbAllowResAccMod = False
    
        'Poi il programma vero e proprio
        Dim oParsed As cReduction
        Set oParsed = DoParse(sSource, lstLog)
        
        If Not oParsed Is Nothing Then
            'Poi compila il programma vero e proprio
            pCompile oParsed
            
            'Istruzione di termine del programma
            moByteCode.Add_IEnd 1
            
            'Adesso cambio il valore dell'NDATA
            moByteCode.SetIstroParam moCurrSymbTable.CountType(estet_const) + _
                                    moCurrSymbTable.CountType(estet_var), _
                                    lBCIndexVars, 1
        
            Compile = True
        Else
            Compile = False
        End If
    Else
        Compile = False
    End If
        
    Exit Function
    
Errore:
    Compile = False
    Log "Compiler Error: " & Err.Description, lstLog

End Function

'
'//////////////////////////////////////////////////////////////////////////////
'


'Rende il tipo di modificatore per una certa variabile o procedura
Private Function getAccessModifierOpt(oReduction As cReduction) As eAccessModifier
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Accessmodifieropt
               '<AccessModifierOpt> ::= <AccessModifier>
                getAccessModifierOpt = getAccessModifierOpt(.Tokens(0).Data)
                
                'Se in pratica sono all'interno di una funzione allora sbrocco
                If Not moCurrSymbTable Is moSymbTable Then
                    moCompError.RaiseError ece_AccessModifiersNotAllowedHere, oReduction.lRow, oReduction.lCol
                End If
                
            Case Rule_Accessmodifieropt2
               '<AccessModifierOpt> ::=
                getAccessModifierOpt = eam_private
            Case Rule_Accessmodifier_Public
               '<AccessModifier> ::= Public
                  getAccessModifierOpt = eam_public
            Case Rule_Accessmodifier_Private
               '<AccessModifier> ::= Private
                  getAccessModifierOpt = eam_private
            Case Rule_Accessmodifier_Reserved
               '<AccessModifier> ::= Reserved
                getAccessModifierOpt = eam_reserved
                
                'Solo io posso usare il modificatore d'accesso e lo abilito solo per la parte delle funzioni riservate
                If Not mbAllowResAccMod Then
                    moCompError.RaiseError ece_CannotUseReservedWord, oReduction.lRow, oReduction.lCol
                End If
        End Select
    End With
End Function


'<VarDecl> ::= <AccessModifierOpt> Const <VarDeclList> <NL>
'<VarDecl> ::= <AccessModifierOpt> Dim <VarDeclList> <NL>
'<VarDecl> ::= <AccessModifierOpt> Var <VarDeclList> <NL>
Private Sub doVarDecl(oReduction As cReduction)
    Dim iAccessMod As eAccessModifier
    
    With oReduction.oData
        'Leggo il modificatore d'accesso
        iAccessMod = getAccessModifierOpt(.Tokens(0).Data)
        
        Select Case .ParentRule.TableIndex
            Case Rule_Vardecl_Dim, Rule_Vardecl_Var
                doVarDeclList .Tokens(2).Data, iAccessMod, estet_var
            
            Case Rule_Vardecl_Const
                doVarDeclList .Tokens(2).Data, iAccessMod, estet_const
        End Select
    End With
End Sub

'<VarDeclList> ::= ID '=' <ConstExpr> ',' <VarDeclList>
'<VarDeclList> ::= ID ',' <VarDeclList>
'<VarDeclList> ::= ID '=' <ConstExpr>
'<VarDeclList> ::= ID
Private Sub doVarDeclList(oReduction As cReduction, _
    ByVal iAccMod As eAccessModifier, ByVal iElType As eSymbTableElType)
    
    Dim sVarName As String
    Dim oSymbEl As cSymbTableEl
    
    'Per la lettura delle costanti
    Dim vVarVal As Variant
    Dim iVarType As eVarType
    
    With oReduction.oData
        sVarName = .Tokens(0).Data
    
        'Aggiunge la variabile alla simbol table
        Set oSymbEl = moCurrSymbTable.AddVar(sVarName, iElType, iAccMod)
        If oSymbEl Is Nothing Then
            'Allora lo cerca e vede se e' riservata
            Set oSymbEl = moCurrSymbTable.Search(sVarName)
            If oSymbEl.iAccMod = eam_reserved Then
                moCompError.RaiseError ece_CannotRedefineReservedWords, oReduction.lRow, oReduction.lCol
            Else
                moCompError.RaiseError ece_DuplicateDeclarationInCurrentScope, oReduction.lRow, oReduction.lCol
            End If
        End If
        
        Select Case .ParentRule.TableIndex
          Case Rule_Vardecllist_Id_Eq_Comma
             '<VarDeclList> ::= ID '=' <ConstExpr> ',' <VarDeclList>
            
            'Legge il valore costante
            doConstExpr .Tokens(2).Data, iVarType, vVarVal
            'Crea l'istruzione per la sua valorizzazione
            moByteCode.Add_IConst iVarType, vVarVal, oReduction.lRow
            moByteCode.Add_IStore moCurrSymbTable.GetVarIndex(sVarName), Not moCurrSymbTable Is moSymbTable, oReduction.lRow
            
            'E si richiama sulle altre
            doVarDeclList .Tokens(4).Data, iAccMod, iElType
          
          Case Rule_Vardecllist_Id_Comma
             '<VarDeclList> ::= ID ',' <VarDeclList>
                        
            'Le costanti devono avere un valore
            If iElType = estet_const Then
                moCompError.RaiseError ece_ConstantMustHaveAValue, oReduction.lRow, oReduction.lCol
            End If
            
            'Si richiama sulle altre
            doVarDeclList .Tokens(2).Data, iAccMod, iElType
            
          Case Rule_Vardecllist_Id_Eq
             '<VarDeclList> ::= ID '=' <ConstExpr>
             
            'Legge il valore costante
            doConstExpr .Tokens(2).Data, iVarType, vVarVal
            'Crea l'istruzione per la sua creazione
            moByteCode.Add_IConst iVarType, vVarVal, oReduction.lRow
            moByteCode.Add_IStore moCurrSymbTable.GetVarIndex(sVarName), Not moCurrSymbTable Is moSymbTable, oReduction.lRow
          
          Case Rule_Vardecllist_Id
             '<VarDeclList> ::= ID
             
            'Le costanti devono avere un valore
            If iElType = estet_const Then
                moCompError.RaiseError ece_ConstantMustHaveAValue, oReduction.lRow, oReduction.lCol
            End If
        End Select
    End With
End Sub

'Rende il carattere separatore dei decimali
Private Function GetDecSep() As String
    GetDecSep = Mid(CStr(0.1), 2, 1)
End Function

'<ConstExpr> ::= <BoolLiteral>
'<ConstExpr> ::= <IntLiteral>
'<ConstExpr> ::= FloatLiteral
'<ConstExpr> ::= StringLiteral
'<ConstExpr> ::= <Nothing>
Private Sub doConstExpr(oReduction As cReduction, _
                    ByRef iType As eVarType, _
                    ByRef vVal As Variant)
    
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Constexpr
               '<ConstExpr> ::= <BoolLiteral>
                vVal = getBoolLiteral(.Tokens(0).Data)
                iType = evt_bool
                
            Case Rule_Constexpr2
               '<ConstExpr> ::= <IntLiteral>
                vVal = getIntLiteral(.Tokens(0).Data)
                iType = evt_long
                  
            Case Rule_Constexpr_Floatliteral
               '<ConstExpr> ::= FloatLiteral
               
               'Rimpiazzo il carattere . con quello che serve in base alla lingua
                vVal = CDbl(Replace(.Tokens(0).Data, ".", GetDecSep()))
                iType = evt_double
                  
            Case Rule_Constexpr_Stringliteral
               '<ConstExpr> ::= StringLiteral
                vVal = CStr(.Tokens(0).Data)
                'Toglie le parentesi
                vVal = Mid(vVal, 2, Len(vVal) - 2)
                'Toglie il doppio apice e lo fa diventare singolo
                vVal = Replace(vVal, """""", """")
                'Setta il tipo stringa
                iType = evt_string
                  
            Case Rule_Constexpr3
               '<ConstExpr> ::= <Nothing>
                vVal = 0
                iType = evt_null
        End Select
    End With
End Sub

'<BoolLiteral> ::= True
'<BoolLiteral> ::= False
Private Function getBoolLiteral(oReduction As cReduction) As Boolean
    With oReduction.oData
        Select Case .ParentRule.TableIndex
        
            Case Rule_Boolliteral_True
               '<BoolLiteral> ::= True
                getBoolLiteral = True
            
            Case Rule_Boolliteral_False
               '<BoolLiteral> ::= False
                getBoolLiteral = False
        
        End Select
    End With
End Function

'<IntLiteral> ::= IntLiteral
'<IntLiteral> ::= HexLiteral
'<IntLiteral> ::= OctLiteral
Private Function getIntLiteral(oReduction As cReduction) As Long
    With oReduction.oData
        getIntLiteral = CLng(.Tokens(0).Data)
        
'        Select Case .ParentRule.TableIndex
'            Case Rule_Intliteral_Intliteral
'               '<IntLiteral> ::= IntLiteral
'
'            Case Rule_Intliteral_Hexliteral
'               '<IntLiteral> ::= HexLiteral
'
'            Case Rule_Intliteral_Octliteral
'               '<IntLiteral> ::= OctLiteral
'        End Select
    End With
End Function

'<SubDecl> ::= <AccessModifierOpt> Sub ID <MethodArgList> <NL> <MethodStmtList> End Sub <NL>
'<SubDecl> ::= <AccessModifierOpt> Function ID <MethodArgList> <NL> <MethodStmtList> End Function <NL>
Private Sub doSubDef(oReduction As cReduction)

    Const C_IDX_ACCMOD = 0
    Const C_IDX_SUBNAME = 2
    Const C_IDX_ARGLIST = 3
    Const C_IDX_BLOCK = 5

    Dim sSubName As String
    Dim oSymbEl As cSymbTableEl
    Dim iVarType As eSymbTableElType
    Dim iAccMod As eAccessModifier
    Dim lBCIndexJump As Long

    With oReduction.oData
        sSubName = .Tokens(C_IDX_SUBNAME).Data
        iAccMod = getAccessModifierOpt(.Tokens(C_IDX_ACCMOD).Data)
        
        Select Case .ParentRule.TableIndex
            Case Rule_Subdecl_Sub_Id_End_Sub
                '<SubDecl> ::= <AccessModifierOpt> Sub ID <MethodArgList> <NL> <MethodStmtList> End Sub <NL>
                iVarType = estet_sub
            
            Case Rule_Subdecl_Function_Id_End_Function
                '<SubDecl> ::= <AccessModifierOpt> Function ID <MethodArgList> <NL> <MethodStmtList> End Function <NL>
                iVarType = estet_func
        End Select
    
        'Aggiunge la variabile alla simbol table
        Set oSymbEl = moCurrSymbTable.AddFunc(sSubName, iVarType, iAccMod)
        If oSymbEl Is Nothing Then
            'Allora lo cerca
            Set oSymbEl = moCurrSymbTable.Search(sSubName)
            If oSymbEl.iAccMod = eam_reserved Then
                moCompError.RaiseError ece_CannotRedefineReservedWords, oReduction.lRow, oReduction.lCol
            Else
                moCompError.RaiseError ece_AmbiguousNameDetected, oReduction.lRow, oReduction.lCol
            End If
        Else
            'Setta come symbol table quella corrente
            Set moCurrSymbTable = oSymbEl.oFuncInfos.oSymbTable
        End If
        
        'Dopo aver creato la Symbol Table locale mette come prima variabile locale
        'il nome della funzione/sub
        moCurrSymbTable.AddVar sSubName, estet_var, eam_notused
        
        'Ovviamente non devo creare il codice per le funzioni/sub riservate
        If oSymbEl.iAccMod <> eam_reserved Then
            'Inserisce un'istruzione di salto oltre la funzione in modo
            'che passi qui dentro solo se viene chiamata
            'Il valore lo modifico alla fine
            
            lBCIndexJump = moByteCode.Count
            moByteCode.Add_IJump -1, oReduction.lRow
        End If
    
        'Processa i parametri della sub/function
        doMethodArgList oSymbEl.oFuncInfos, .Tokens(C_IDX_ARGLIST).Data
    
        'Ovviamente non devo creare il codice per le funzioni/sub riservate
        If oSymbEl.iAccMod <> eam_reserved Then
            'Assegna il puntatore all'indirizzo in memoria in cui e' la funzione
            oSymbEl.oFuncInfos.lInstrPointer = moByteCode.Count()
            
            'Istruzione per l'allocazione delle n. variabili locali
            'Da modificare alla fine quando ho tutta la Symb Table completa
            Dim lBCIndexVars As Long
            lBCIndexVars = moByteCode.Count
            moByteCode.Add_INData 0, oReduction.lRow
            
            'Adesso processa le istruzioni interne
            pCompile .Tokens(C_IDX_BLOCK).Data
        
            'Alla fine inserisce l'istruzione di ritorno per l'uscita dalla funzione
            moByteCode.Add_IReturn oReduction.lRow
        
            'Adesso cambia le coordinate del salto
            moByteCode.SetIstroParam moByteCode.Count(), lBCIndexJump, 1
            
            'Adesso cambia il numero di var. da allocare
            moByteCode.SetIstroParam moCurrSymbTable.CountType(estet_const) + _
                                moCurrSymbTable.CountType(estet_var) - _
                                oSymbEl.oFuncInfos.oCollParams.Count - 1, _
                                lBCIndexVars, 1
        Else
            'Se e' una funzione riservta allora tolgo il puntatore al programma per pulizia
            oSymbEl.oFuncInfos.lInstrPointer = -1
        End If
        
    End With
    
    'Risetta come symboltable quella globale
    Set moCurrSymbTable = moSymbTable
End Sub

'<MethodArgList> ::= '(' <ArgList> ')'
'<MethodArgList> ::= '(' ')'
'<ArgList> ::= <Arg> ',' <ArgList>
'<ArgList> ::= <Arg>
'<Arg> ::= <ArgModifier> ID
'
'Si richiama varie volte fino a raggiungere l'argomento
'
Private Sub doMethodArgList(oFuncInfos As cFuncInfos, oReduction As cReduction)
    
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Methodarglist_Lparan_Rparan
               '<MethodArgList> ::= '(' <ArgList> ')'
                doMethodArgList oFuncInfos, .Tokens(1).Data
            
            Case Rule_Methodarglist_Lparan_Rparan2
               '<MethodArgList> ::= '(' ')'
            
            Case Rule_Arglist_Comma
               '<ArgList> ::= <Arg> ',' <ArgList>
               
               '<Arg>
               doMethodArgList oFuncInfos, .Tokens(0).Data
                  
               '<ArgList>
               doMethodArgList oFuncInfos, .Tokens(2).Data
                  
            Case Rule_Arglist
               '<ArgList> ::= <Arg>
                
                doMethodArgList oFuncInfos, .Tokens(0).Data
                
            Case Rule_Arg_Id
               '<Arg> ::= <ArgModifier> ID
               
               'Aggiunge il parametro alla funzione corrente
                oFuncInfos.AddParam getArgModifier(.Tokens(0).Data), .Tokens(1).Data
               
               'e lo aggiunge anche alle variabili locali
                oFuncInfos.oSymbTable.AddVar .Tokens(1).Data, estet_var, eAccessModifier.eam_notused
               
        End Select
    End With
End Sub

'<ArgModifier> ::= ByVal
'<ArgModifier> ::= ByRef
'<ArgModifier> ::=
Private Function getArgModifier(oReduction As cReduction) As eParmMod
    'Modificatore non specificato
    getArgModifier = epm_NotSpecified
    
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Argmodifier_Byval
               '<ArgModifier> ::= ByVal
               getArgModifier = epm_ByVal
                  
            Case Rule_Argmodifier_Byref
               '<ArgModifier> ::= ByRef
               getArgModifier = epm_ByRef
        End Select
    End With
End Function

'<AssignStmt> ::= <QualifiedID> '=' <Expr>
'<AssignStmt> ::= <QualifiedID> '+=' <Expr>
'<AssignStmt> ::= <QualifiedID> '-=' <Expr>
'<AssignStmt> ::= <QualifiedID> '++'
'<AssignStmt> ::= <QualifiedID> '--'
Private Sub doAssignStmt(oReduction As cReduction)
    Dim oSymbEl As cSymbTableEl
    Dim bIsLocal As Boolean
    Dim poSymbTable As cSymbTable

    With oReduction.oData
        'Prende il nome della variabile
        getQualifiedId .Tokens(0).Data, oSymbEl, bIsLocal
        
        'Non posso cambiare il valore di una costante
        If oSymbEl.iType = estet_const Then
            moCompError.RaiseError ece_AssignmentToConstantNotPermitted, oReduction.lRow, oReduction.lCol
        End If
        
        If bIsLocal Then
            Set poSymbTable = moCurrSymbTable
        Else
            Set poSymbTable = moSymbTable
            
            If oSymbEl.iType = estet_func Or oSymbEl.iType = estet_sub Then
                moCompError.RaiseError ece_AssignmentToFunctionNotPermitted, oReduction.lRow, oReduction.lCol
            End If
            
        End If
    
        Select Case .ParentRule.TableIndex
            Case Rule_Assignstmt_Eq
                '<AssignStmt> ::= <QualifiedID> '=' <Expr>
                
                'Adesso esegue l'espressione
                doExpr .Tokens(2).Data
                
                'Adesso crea l'istro di assegnazione
                moByteCode.Add_IStore poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
        
            Case Rule_Assignstmt_Pluseq
             '<AssignStmt> ::= <QualifiedID> '+=' <Expr>
                
                'Aggiunge il suo valore attuale in testa allo stack
                moByteCode.Add_IFetch poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
                
                'Adesso esegue l'espressione
                doExpr .Tokens(2).Data
                
                'Aggiunge l'istruzione BinOp
                moByteCode.Add_IBinOp eboc_plus, oReduction.lRow
             
                'Adesso crea l'istro di assegnazione
                moByteCode.Add_IStore poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
                
            Case Rule_Assignstmt_Minuseq
             '<AssignStmt> ::= <QualifiedID> '-=' <Expr>
             
                'Aggiunge il suo valore attuale in testa allo stack
                moByteCode.Add_IFetch poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
                
                'Adesso esegue l'espressione
                doExpr .Tokens(2).Data
                
                'Aggiunge l'istruzione BinOp
                moByteCode.Add_IBinOp eboc_minus, oReduction.lRow
             
                'Adesso crea l'istro di assegnazione
                moByteCode.Add_IStore poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
            
            Case Rule_Assignstmt_Plusplus
             '<AssignStmt> ::= <QualifiedID> '++'
                
                'Aggiunge il suo valore attuale in testa allo stack
                moByteCode.Add_IFetch poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
                
                'Aggiunge uno
                moByteCode.Add_IConst eVarType.evt_long, 1, oReduction.lRow
                
                'Aggiunge l'istruzione BinOp
                moByteCode.Add_IBinOp eboc_plus, oReduction.lRow
             
                'Adesso crea l'istro di assegnazione
                moByteCode.Add_IStore poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
             
            Case Rule_Assignstmt_Minusminus
             '<AssignStmt> ::= <QualifiedID> '--'
             
                'Aggiunge il suo valore attuale in testa allo stack
                moByteCode.Add_IFetch poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
                
                'Aggiunge uno
                moByteCode.Add_IConst eVarType.evt_long, 1, oReduction.lRow
                
                'Aggiunge l'istruzione BinOp
                moByteCode.Add_IBinOp eboc_minus, oReduction.lRow
             
                'Adesso crea l'istro di assegnazione
                moByteCode.Add_IStore poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
        
        End Select
    End With
End Sub

'Genera le istruzioni per una chiamata a funzione/sub
Private Sub makeSubCall(oExprList As cReduction, oSymbEl As cSymbTableEl, bIsStatement As Boolean)
    'Aggiunge l'IFrame
    moByteCode.Add_IFrame oExprList.lRow
    
    'Crea lo spazio per il valore di ritorno (la var. con lo stesso nome della sub/func)
    moByteCode.Add_INData 1, oExprList.lRow
    
    'Adesso vede i vari parametri da passare, li esegue e conta
    doExprList oExprList, oSymbEl.oFuncInfos.oCollParams.Count, 0
    
    'Adesso aggiunge la chiamata alla funzione
    If oSymbEl.iAccMod = eam_reserved Then
        moByteCode.Add_IPCall moResSymb.GetPFuncCode(oSymbEl), oExprList.lRow
    Else
        moByteCode.Add_IFCall oSymbEl.oFuncInfos.lInstrPointer, oSymbEl.oFuncInfos.oCollParams.Count, oExprList.lRow
    End If
    
    'Se e' una riga non mi serve il valore ritornato dalla funzione
    If bIsStatement Then
        moByteCode.Add_IChop oExprList.lRow
    End If
    
End Sub

'<SubCallStmt> ::= <QualifiedID> '(' <ExprList> ')'
Private Sub doSubCallStmt(oReduction As cReduction)
    Dim oSymbEl As cSymbTableEl
    Dim poSymbTable As cSymbTable
    
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Subcallstmt_Lparan_Rparan
               '<SubCallStmt> ::= <QualifiedID> '(' <ExprList> ')'
                
                'Controllo solo sul globale per funzioni
                Set poSymbTable = moCurrSymbTable
                Set moCurrSymbTable = moSymbTable
                
                'Prende il nome della variabile se definita
                getQualifiedId .Tokens(0).Data, oSymbEl, False
                
                'Risetto com'era prima del treppo
                Set moCurrSymbTable = poSymbTable
                
                'Se e' sub o function
                If oSymbEl.iType = estet_func Or oSymbEl.iType = estet_sub Then
                    'Crea le istruzioni per la chiamata
                    makeSubCall .Tokens(2).Data, oSymbEl, True
                Else
                    'Vede se per caso e' una funzione che ha lo stesso nome
                    Set oSymbEl = moSymbTable.Search(oSymbEl.sName)

                    'Non ancora implementato
                    moCompError.RaiseError ece_NotImplementedYet, oReduction.lRow, oReduction.lCol
                End If
        End Select
    End With
End Sub

'<QualifiedID> ::= ID
Private Sub getQualifiedId(oReduction As cReduction, _
            ByRef oSymbEl As cSymbTableEl, ByRef bIsLocal As Boolean)
    
    
    Dim sId As String
    
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Qualifiedid_Id
               '<QualifiedID> ::= ID
                sId = .Tokens(0).Data
        End Select
    End With

    'Lo metto fuori cosi' se arriva un altro caso non gestito segnala errore
    'Controlla sulla symbol table locale
    Set oSymbEl = moCurrSymbTable.Search(sId)
    If Not oSymbEl Is Nothing Then
        bIsLocal = Not (moCurrSymbTable Is moSymbTable)
    Else
        'Controlla sulla symbol table globale
        Set oSymbEl = moSymbTable.Search(sId)
        If Not oSymbEl Is Nothing Then
            bIsLocal = False
        Else
            'Variabile/Funzione non definita
            moCompError.RaiseError ece_VariableNotDefined, oReduction.lRow, oReduction.lCol
        End If
    End If
End Sub


Private Sub doExpr(oReduction As cReduction)
    With oReduction.oData
        Select Case .ParentRule.TableIndex
        
            Case Rule_Expr
               '<Expr> ::= <ImpExpr>
                doExpr .Tokens(0).Data
                  
            Case Rule_Impexpr_Imp
               '<ImpExpr> ::= <ImpExpr> Imp <EqvExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_imp, oReduction.lRow
                
            Case Rule_Impexpr
               '<ImpExpr> ::= <EqvExpr>
                doExpr .Tokens(0).Data
                  
            Case Rule_Eqvexpr_Eqv
               '<EqvExpr> ::= <EqvExpr> Eqv <XorExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_eqv, oReduction.lRow
            
            Case Rule_Eqvexpr
               '<EqvExpr> ::= <XorExpr>
                doExpr .Tokens(0).Data
                  
            Case Rule_Xorexpr_Xor
               '<XorExpr> ::= <XorExpr> Xor <OrExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_xor, oReduction.lRow
            
            Case Rule_Xorexpr
               '<XorExpr> ::= <OrExpr>
                doExpr .Tokens(0).Data
                  
            Case Rule_Orexpr_Or
               '<OrExpr> ::= <OrExpr> Or <AndExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_or, oReduction.lRow
            
            Case Rule_Orexpr
               '<OrExpr> ::= <AndExpr>
                doExpr .Tokens(0).Data
                  
            Case Rule_Andexpr_And
               '<AndExpr> ::= <AndExpr> And <NotExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_and, oReduction.lRow
            
            Case Rule_Andexpr
               '<AndExpr> ::= <NotExpr>
                doExpr .Tokens(0).Data
                  
            Case Rule_Notexpr_Not
               '<NotExpr> ::= Not <NotExpr>
                doExpr .Tokens(1).Data
                'operazione
                moByteCode.Add_IUnOp eUnOpCodes.euoc_not, oReduction.lRow
                  
            Case Rule_Notexpr
               '<NotExpr> ::= <CompareExpr>
                doExpr .Tokens(0).Data
                  
            Case Rule_Compareexpr_Gteq
               '<CompareExpr> ::= <CompareExpr> '>=' <ConcatExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_be, oReduction.lRow
                
            Case Rule_Compareexpr_Lteq
               '<CompareExpr> ::= <CompareExpr> '<=' <ConcatExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_le, oReduction.lRow
                  
            Case Rule_Compareexpr_Gt
               '<CompareExpr> ::= <CompareExpr> '>' <ConcatExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_bt, oReduction.lRow
                  
            Case Rule_Compareexpr_Lt
               '<CompareExpr> ::= <CompareExpr> '<' <ConcatExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_lt, oReduction.lRow
                  
            Case Rule_Compareexpr_Ltgt
               '<CompareExpr> ::= <CompareExpr> '<>' <ConcatExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_neq, oReduction.lRow
                  
            Case Rule_Compareexpr_Eq
               '<CompareExpr> ::= <CompareExpr> '=' <ConcatExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_eq, oReduction.lRow
                  
            Case Rule_Compareexpr
               '<CompareExpr> ::= <ConcatExpr>
               doExpr .Tokens(0).Data
                  
            Case Rule_Concatexpr_Amp
               '<ConcatExpr> ::= <ConcatExpr> '&' <AddExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_amp, oReduction.lRow
            
            Case Rule_Concatexpr
               '<ConcatExpr> ::= <AddExpr>
                doExpr .Tokens(0).Data
                  
            Case Rule_Addexpr_Plus
               '<AddExpr> ::= <AddExpr> '+' <ModExpr>
               
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_plus, oReduction.lRow
            
            Case Rule_Addexpr_Minus
               '<AddExpr> ::= <AddExpr> '-' <ModExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_minus, oReduction.lRow
                  
            Case Rule_Addexpr
               '<AddExpr> ::= <ModExpr>
               doExpr .Tokens(0).Data
                  
            Case Rule_Modexpr_Mod
               '<ModExpr> ::= <ModExpr> Mod <IntDivExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_mod, oReduction.lRow
            
            Case Rule_Modexpr
               '<ModExpr> ::= <IntDivExpr>
               doExpr .Tokens(0).Data
                  
            Case Rule_Intdivexpr_Backslash
               '<IntDivExpr> ::= <IntDivExpr> '\' <MultExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_intdiv, oReduction.lRow
            
            Case Rule_Intdivexpr
               '<IntDivExpr> ::= <MultExpr>
               doExpr .Tokens(0).Data
                  
            Case Rule_Multexpr_Times
               '<MultExpr> ::= <MultExpr> '*' <UnaryExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_mul, oReduction.lRow
                  
            Case Rule_Multexpr_Div
               '<MultExpr> ::= <MultExpr> '/' <UnaryExpr>
                doExpr .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_div, oReduction.lRow
                  
            Case Rule_Multexpr
               '<MultExpr> ::= <UnaryExpr>
               doExpr .Tokens(0).Data
                  
            Case Rule_Unaryexpr_Minus
               '<UnaryExpr> ::= '-' <UnaryExpr>
                doExpr .Tokens(1).Data
                'operazione
                moByteCode.Add_IUnOp eUnOpCodes.euoc_minus, oReduction.lRow
                  
            Case Rule_Unaryexpr_Plus
               '<UnaryExpr> ::= '+' <UnaryExpr>
                'Non fa nulla
                
            Case Rule_Unaryexpr
               '<UnaryExpr> ::= <ExpExpr>
               doExpr .Tokens(0).Data
                  
            Case Rule_Expexpr_Caret
               '<ExpExpr> ::= <Value> '^' <ExpExpr>
                doValue .Tokens(0).Data
                doExpr .Tokens(2).Data
                'operazione
                moByteCode.Add_IBinOp eBinOpCodes.eboc_caret, oReduction.lRow
            
            Case Rule_Expexpr
               '<ExpExpr> ::= <Value>
               doValue .Tokens(0).Data
        
        End Select
    End With
End Sub

'Chiamata nella doValue
'<LeftExpr> ::= <QualifiedID>
'<LeftExpr> ::= <QualifiedID> '(' <ExprList> ')'
Private Sub doLeftExpr(oReduction As cReduction)
    Dim oSymbEl As cSymbTableEl
    Dim bIsLocal As Boolean
    Dim poSymbTable As cSymbTable
    
    With oReduction.oData
        'Prende l'ID e controlla la sua esistenza
        getQualifiedId .Tokens(0).Data, oSymbEl, bIsLocal
        
        If bIsLocal Then
            Set poSymbTable = moCurrSymbTable
        Else
            Set poSymbTable = moSymbTable
        End If
        
        Select Case .ParentRule.TableIndex
            Case Rule_Leftexpr
               '<LeftExpr> ::= <QualifiedID>
                
                'Devo metterne il valore in cima allo stack
                moByteCode.Add_IFetch poSymbTable.GetVarIndex(oSymbEl.sName), bIsLocal, oReduction.lRow
            
                'Parentesi obbligatorie altrimenti si spacca tutto
                If Not bIsLocal Then
                    If oSymbEl.iType = estet_func Or oSymbEl.iType = estet_sub Then
                        moCompError.RaiseError ece_VariableNotDefined, oReduction.lRow, oReduction.lCol
                    End If
                End If
            
            Case Rule_Leftexpr_Lparan_Rparan
               '<LeftExpr> ::= <QualifiedID> '(' <ExprList> ')'
               
               'Devo vedere se e' una funzione o un'array
                If oSymbEl.iType = estet_func Then
                    
                    'Crea le istruzioni per la chiamata
                    makeSubCall .Tokens(2).Data, oSymbEl, False
                    'In questo caso lascia il valore di ritorno in cima allo stack
                    
                ElseIf oSymbEl.iType = estet_sub Then
                    moCompError.RaiseError ece_SubCannotReturnAValue, oReduction.lRow, oReduction.lCol
                
                    'Non posso chiamare una sub qua
                Else
                    'Esegue la lettura di un valore all'interno di un'array
                    '!!!Per ora bloccato!!!
                    moCompError.RaiseError ece_NotImplementedYet, oReduction.lRow, oReduction.lCol
                End If

        End Select
    End With
End Sub

'Parametri da passare ad una chiamata di funzione/sub
Private Sub doExprList(oReduction As cReduction, _
    ByVal iNumOfParsDefined As Integer, _
    ByRef iNumOfParsCounted As Integer)
    
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Exprlist
               '<ExprList> ::= <Expr> <CommaExprList>
                iNumOfParsCounted = iNumOfParsCounted + 1
                doExpr .Tokens(0).Data
                doExprList .Tokens(1).Data, iNumOfParsDefined, iNumOfParsCounted
                Exit Sub
                
            Case Rule_Exprlist2
               '<ExprList> ::= <Expr>
                iNumOfParsCounted = iNumOfParsCounted + 1
                doExpr .Tokens(0).Data
            
            Case Rule_Exprlist3
               '<ExprList> ::=
                  
            Case Rule_Commaexprlist_Comma
               '<CommaExprList> ::= ',' <Expr> <CommaExprList>
                iNumOfParsCounted = iNumOfParsCounted + 1
                doExpr .Tokens(1).Data
                doExprList .Tokens(2).Data, iNumOfParsDefined, iNumOfParsCounted
                Exit Sub
                  
            Case Rule_Commaexprlist_Comma2
               '<CommaExprList> ::= ',' <Expr>
                iNumOfParsCounted = iNumOfParsCounted + 1
                doExpr .Tokens(1).Data
        End Select
    End With

    'Controllo sul numero di parametri passati
    If iNumOfParsDefined > -1 Then
        If iNumOfParsCounted <> iNumOfParsDefined Then
            moCompError.RaiseError ece_WrongNumbersOfParameters, oReduction.lRow, oReduction.lCol
        End If
    End If

End Sub

Private Sub doValue(oReduction As cReduction)
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Value
               '<Value> ::= <ConstExpr>
               
                Dim iVarType As eVarType
                Dim vVarValue As Variant
                'Legge il valore della costante
                doConstExpr .Tokens(0).Data, iVarType, vVarValue
                'Aggiunge l'istruzione di gestione
                moByteCode.Add_IConst iVarType, vVarValue, oReduction.lRow
                  
            Case Rule_Value2
               '<Value> ::= <LeftExpr>
                doLeftExpr .Tokens(0).Data
                  
            Case Rule_Value_Lparan_Rparan
               '<Value> ::= '(' <Expr> ')'
                doExpr .Tokens(1).Data

        End Select
    End With
End Sub

Private Sub doInlineStm(oReduction As cReduction)
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Inlinestmt
               '<InlineStmt> ::= <AssignStmt>
                doAssignStmt .Tokens(0).Data
            Case Rule_Inlinestmt2
               '<InlineStmt> ::= <SubCallStmt>
                doSubCallStmt .Tokens(0).Data
            Case Rule_Inlinestmt_Exit_Lparan_Rparan
                moByteCode.Add_IEnd oReduction.lRow
        End Select
    End With
End Sub

Private Sub doLoopStm(oReduction As cReduction)
    Dim lWhileAddr As Long
    Dim lIfAddr As Long

    With oReduction.oData
        Select Case .ParentRule.TableIndex
            ' <LoopStmt> ::= While <Expr> <NL> <BlockStmtList> WEnd <NL>
            Case Rule_Loopstmt_While_Wend
                lWhileAddr = moByteCode.Count
                
                doExpr .Tokens(1).Data
                
                'Adesso controllo se l'espressione e' vera
                lIfAddr = moByteCode.Count
                moByteCode.Add_IPopJF -1, oReduction.lRow
                
                'Eseguo le espressioni interne
                pCompile .Tokens(3).Data
                
                'Salta in alto
                moByteCode.Add_IJump lWhileAddr, oReduction.lRow
                
                'Adesso metto a posto il jump all'istruzione seguente
                moByteCode.SetIstroParam moByteCode.Count, lIfAddr, 1
        End Select
    End With
End Sub

Private Sub doIfStm(oReduction As cReduction)

    Dim lIfAddr As Long
    Dim lElseAddr As Long
    
    With oReduction.oData
        Select Case .ParentRule.TableIndex
            Case Rule_Ifstmt_If_Then_End_If
               '<IfStmt> ::= If <Expr> Then <NL> <BlockStmtList> <ElseStmtList> End If <NL>
                doExpr .Tokens(1).Data
                
                'Adesso controllo se l'espressione e' vera
                lIfAddr = moByteCode.Count
                moByteCode.Add_IPopJF -1, oReduction.lRow
                
                'Eseguo le espressioni interne
                pCompile .Tokens(4).Data
                
                'Adesso devo saltare oltre l'else
                lElseAddr = moByteCode.Count
                moByteCode.Add_IJump -1, oReduction.lRow
                'Adesso metto a posto il jump all'istruzione seguente
                moByteCode.SetIstroParam moByteCode.Count, lIfAddr, 1
                
                'Eseguo l'else se esiste
                doIfStm .Tokens(5).Data
                
                'E qui metto a posto il jump
                moByteCode.SetIstroParam moByteCode.Count, lElseAddr, 1
            
            Case Rule_Ifstmt_If_Then
               '<IfStmt> ::= If <Expr> Then <InlineStmt> <ElseOpt> <EndIfOpt> <NL>
                doExpr .Tokens(1).Data
                  
                'Adesso controllo se l'espressione e' vera
                lIfAddr = moByteCode.Count
                moByteCode.Add_IPopJF -1, oReduction.lRow
                
                'Esegue le funzioni
                doInlineStm .Tokens(3).Data
                  
                'Adesso devo saltare oltre l'else
                lElseAddr = moByteCode.Count
                moByteCode.Add_IJump -1, oReduction.lRow
                'Adesso metto a posto il jump all'istruzione seguente
                moByteCode.SetIstroParam moByteCode.Count, lIfAddr, 1
            
                'Eseguo l'else se esiste
                doIfStm .Tokens(4).Data
            
                'E qui metto a posto il jump
                moByteCode.SetIstroParam moByteCode.Count, lElseAddr, 1
            
            Case Rule_Elsestmtlist_Elseif_Then
               '<ElseStmtList> ::= ElseIf <Expr> Then <NL> <BlockStmtList> <ElseStmtList>
                doExpr .Tokens(1).Data
                  
                'Adesso controllo se l'espressione e' vera
                lIfAddr = moByteCode.Count
                moByteCode.Add_IPopJF -1, oReduction.lRow
                  
                'Eseguo le espressioni interne
                pCompile .Tokens(4).Data
                
                'Adesso devo saltare oltre l'else
                lElseAddr = moByteCode.Count
                moByteCode.Add_IJump -1, oReduction.lRow
                'Adesso metto a posto il jump all'istruzione seguente
                moByteCode.SetIstroParam moByteCode.Count, lIfAddr, 1
                
                'Eseguo l'else se esiste
                doIfStm .Tokens(5).Data
                  
                'E qui metto a posto il jump
                moByteCode.SetIstroParam moByteCode.Count, lElseAddr, 1
            
            Case Rule_Elsestmtlist_Else
               '<ElseStmtList> ::= Else <NL> <BlockStmtList>
                  pCompile .Tokens(2).Data
                  
            Case Rule_Elsestmtlist
               '<ElseStmtList> ::=
                  
            Case Rule_Elseopt_Else
               '<ElseOpt> ::= Else <InlineStmt>
                doInlineStm .Tokens(1).Data
                  
            Case Rule_Elseopt
               '<ElseOpt> ::=
                  
            Case Rule_Endifopt_End_If
               '<EndIfOpt> ::= End If
                  
            Case Rule_Endifopt
               '<EndIfOpt> ::=
        End Select
    End With
End Sub


Private Sub pCompile(oReduction As cReduction)
    Dim i As Integer
    Dim o As Token
    
    With oReduction
        Select Case .oData.ParentRule.TableIndex
            
            Case Rule_Vardecl_Const, Rule_Vardecl_Dim, Rule_Vardecl_Var
                '<VarDecl> ::= <AccessModifierOpt> Const <VarDeclList> <NL>
                '<VarDecl> ::= <AccessModifierOpt> Dim <VarDeclList> <NL>
                '<VarDecl> ::= <AccessModifierOpt> Var <VarDeclList> <NL>
                doVarDecl oReduction
                Exit Sub
            
            Case Rule_Subdecl_Sub_Id_End_Sub, _
                Rule_Subdecl_Function_Id_End_Function
               '<SubDecl> ::= <AccessModifierOpt> Sub ID <MethodArgList> <NL> <MethodStmtList> End Sub <NL>
               '<SubDecl> ::= <AccessModifierOpt> Function ID <MethodArgList> <NL> <MethodStmtList> End Function <NL>
               doSubDef oReduction
               Exit Sub
      
            Case Rule_Blockstmt
               '<BlockStmt> ::= <InlineStmt> <NL>
               doInlineStm .oData.Tokens(0).Data
               Exit Sub
               
           Case Rule_Blockstmt2
               '<BlockStmt> ::= <IfStmt>
               doIfStm .oData.Tokens(0).Data
               Exit Sub
        
            Case Rule_Blockstmt3
               '<BlockStmt> ::= <LoopStmt>
               doLoopStm .oData.Tokens(0).Data
               Exit Sub
        
        End Select
        
        'Adesso va in ricorsione sui simboli non terminali
        For i = 0 To .oData.TokenCount - 1
            Set o = .oData.Tokens(i)
            
            If o.Kind = SymbolTypeNonterminal Then
                pCompile o.Data
            End If
        Next
    End With

End Sub

