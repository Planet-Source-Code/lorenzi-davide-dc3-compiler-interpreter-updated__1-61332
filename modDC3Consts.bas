Public Enum SymbolConstants
   Symbol_Eof               = 0   ' (EOF)
   Symbol_Error             = 1   ' (Error)
   Symbol_Whitespace        = 2   ' (Whitespace)
   Symbol_Commentline       = 3   ' (Comment Line)
   Symbol_Minus             = 4   ' '-'
   Symbol_Minusminus        = 5   ' '--'
   Symbol_Amp               = 6   ' '&'
   Symbol_Lparan            = 7   ' '('
   Symbol_Rparan            = 8   ' ')'
   Symbol_Times             = 9   ' '*'
   Symbol_Comma             = 10  ' ','
   Symbol_Div               = 11  ' '/'
   Symbol_Backslash         = 12  ' '\'
   Symbol_Caret             = 13  ' '^'
   Symbol_Plus              = 14  ' '+'
   Symbol_Plusplus          = 15  ' '++'
   Symbol_Pluseq            = 16  ' '+='
   Symbol_Lt                = 17  ' '<'
   Symbol_Lteq              = 18  ' '<='
   Symbol_Ltgt              = 19  ' '<>'
   Symbol_Eq                = 20  ' '='
   Symbol_Minuseq           = 21  ' '-='
   Symbol_Gt                = 22  ' '>'
   Symbol_Gteq              = 23  ' '>='
   Symbol_And               = 24  ' And
   Symbol_Byref             = 25  ' ByRef
   Symbol_Byval             = 26  ' ByVal
   Symbol_Const             = 27  ' Const
   Symbol_Dim               = 28  ' Dim
   Symbol_Else              = 29  ' Else
   Symbol_Elseif            = 30  ' ElseIf
   Symbol_Empty             = 31  ' Empty
   Symbol_End               = 32  ' End
   Symbol_Eqv               = 33  ' Eqv
   Symbol_Exit              = 34  ' Exit
   Symbol_False             = 35  ' False
   Symbol_Floatliteral      = 36  ' FloatLiteral
   Symbol_Function          = 37  ' Function
   Symbol_Hexliteral        = 38  ' HexLiteral
   Symbol_Id                = 39  ' ID
   Symbol_If                = 40  ' If
   Symbol_Imp               = 41  ' Imp
   Symbol_Intliteral        = 42  ' IntLiteral
   Symbol_Mod               = 43  ' Mod
   Symbol_Newline           = 44  ' NewLine
   Symbol_Not               = 45  ' Not
   Symbol_Nothing           = 46  ' Nothing
   Symbol_Null              = 47  ' Null
   Symbol_Octliteral        = 48  ' OctLiteral
   Symbol_Or                = 49  ' Or
   Symbol_Private           = 50  ' Private
   Symbol_Public            = 51  ' Public
   Symbol_Reserved          = 52  ' Reserved
   Symbol_Stringliteral     = 53  ' StringLiteral
   Symbol_Sub               = 54  ' Sub
   Symbol_Then              = 55  ' Then
   Symbol_True              = 56  ' True
   Symbol_Var               = 57  ' Var
   Symbol_Wend              = 58  ' WEnd
   Symbol_While             = 59  ' While
   Symbol_Xor               = 60  ' Xor
   Symbol_Accessmodifier    = 61  ' <AccessModifier>
   Symbol_Accessmodifieropt = 62  ' <AccessModifierOpt>
   Symbol_Addexpr           = 63  ' <AddExpr>
   Symbol_Andexpr           = 64  ' <AndExpr>
   Symbol_Arg               = 65  ' <Arg>
   Symbol_Arglist           = 66  ' <ArgList>
   Symbol_Argmodifier       = 67  ' <ArgModifier>
   Symbol_Assignstmt        = 68  ' <AssignStmt>
   Symbol_Blockstmt         = 69  ' <BlockStmt>
   Symbol_Blockstmtlist     = 70  ' <BlockStmtList>
   Symbol_Boolliteral       = 71  ' <BoolLiteral>
   Symbol_Commaexprlist     = 72  ' <CommaExprList>
   Symbol_Compareexpr       = 73  ' <CompareExpr>
   Symbol_Concatexpr        = 74  ' <ConcatExpr>
   Symbol_Constexpr         = 75  ' <ConstExpr>
   Symbol_Elseopt           = 76  ' <ElseOpt>
   Symbol_Elsestmtlist      = 77  ' <ElseStmtList>
   Symbol_Endifopt          = 78  ' <EndIfOpt>
   Symbol_Eqvexpr           = 79  ' <EqvExpr>
   Symbol_Expexpr           = 80  ' <ExpExpr>
   Symbol_Expr              = 81  ' <Expr>
   Symbol_Exprlist          = 82  ' <ExprList>
   Symbol_Globalstmt        = 83  ' <GlobalStmt>
   Symbol_Globalstmtlist    = 84  ' <GlobalStmtList>
   Symbol_Ifstmt            = 85  ' <IfStmt>
   Symbol_Impexpr           = 86  ' <ImpExpr>
   Symbol_Inlinestmt        = 87  ' <InlineStmt>
   Symbol_Intdivexpr        = 88  ' <IntDivExpr>
   Symbol_Intliteral2       = 89  ' <IntLiteral>
   Symbol_Leftexpr          = 90  ' <LeftExpr>
   Symbol_Loopstmt          = 91  ' <LoopStmt>
   Symbol_Methodarglist     = 92  ' <MethodArgList>
   Symbol_Methodstmt        = 93  ' <MethodStmt>
   Symbol_Methodstmtlist    = 94  ' <MethodStmtList>
   Symbol_Modexpr           = 95  ' <ModExpr>
   Symbol_Multexpr          = 96  ' <MultExpr>
   Symbol_Nl                = 97  ' <NL>
   Symbol_Nlopt             = 98  ' <NLOpt>
   Symbol_Notexpr           = 99  ' <NotExpr>
   Symbol_Nothing2          = 100 ' <Nothing>
   Symbol_Orexpr            = 101 ' <OrExpr>
   Symbol_Program           = 102 ' <Program>
   Symbol_Qualifiedid       = 103 ' <QualifiedID>
   Symbol_Subcallstmt       = 104 ' <SubCallStmt>
   Symbol_Subdecl           = 105 ' <SubDecl>
   Symbol_Unaryexpr         = 106 ' <UnaryExpr>
   Symbol_Value             = 107 ' <Value>
   Symbol_Vardecl           = 108 ' <VarDecl>
   Symbol_Vardecllist       = 109 ' <VarDeclList>
   Symbol_Xorexpr           = 110 ' <XorExpr>
End Enum

Public Enum RuleConstants
   Rule_Nl_Newline                       = 0   ' <NL> ::= NewLine <NL>
   Rule_Nl_Newline2                      = 1   ' <NL> ::= NewLine
   Rule_Nlopt                            = 2   ' <NLOpt> ::= <NL>
   Rule_Nlopt2                           = 3   ' <NLOpt> ::= 
   Rule_Program                          = 4   ' <Program> ::= <NLOpt> <GlobalStmtList>
   Rule_Globalstmtlist                   = 5   ' <GlobalStmtList> ::= <GlobalStmt> <GlobalStmtList>
   Rule_Globalstmtlist2                  = 6   ' <GlobalStmtList> ::= 
   Rule_Globalstmt                       = 7   ' <GlobalStmt> ::= <VarDecl>
   Rule_Globalstmt2                      = 8   ' <GlobalStmt> ::= <SubDecl>
   Rule_Globalstmt3                      = 9   ' <GlobalStmt> ::= <BlockStmt>
   Rule_Vardecl_Const                    = 10  ' <VarDecl> ::= <AccessModifierOpt> Const <VarDeclList> <NL>
   Rule_Vardecl_Dim                      = 11  ' <VarDecl> ::= <AccessModifierOpt> Dim <VarDeclList> <NL>
   Rule_Vardecl_Var                      = 12  ' <VarDecl> ::= <AccessModifierOpt> Var <VarDeclList> <NL>
   Rule_Accessmodifieropt                = 13  ' <AccessModifierOpt> ::= <AccessModifier>
   Rule_Accessmodifieropt2               = 14  ' <AccessModifierOpt> ::= 
   Rule_Accessmodifier_Public            = 15  ' <AccessModifier> ::= Public
   Rule_Accessmodifier_Private           = 16  ' <AccessModifier> ::= Private
   Rule_Accessmodifier_Reserved          = 17  ' <AccessModifier> ::= Reserved
   Rule_Vardecllist_Id_Eq_Comma          = 18  ' <VarDeclList> ::= ID '=' <ConstExpr> ',' <VarDeclList>
   Rule_Vardecllist_Id_Comma             = 19  ' <VarDeclList> ::= ID ',' <VarDeclList>
   Rule_Vardecllist_Id_Eq                = 20  ' <VarDeclList> ::= ID '=' <ConstExpr>
   Rule_Vardecllist_Id                   = 21  ' <VarDeclList> ::= ID
   Rule_Constexpr                        = 22  ' <ConstExpr> ::= <BoolLiteral>
   Rule_Constexpr2                       = 23  ' <ConstExpr> ::= <IntLiteral>
   Rule_Constexpr_Floatliteral           = 24  ' <ConstExpr> ::= FloatLiteral
   Rule_Constexpr_Stringliteral          = 25  ' <ConstExpr> ::= StringLiteral
   Rule_Constexpr3                       = 26  ' <ConstExpr> ::= <Nothing>
   Rule_Boolliteral_True                 = 27  ' <BoolLiteral> ::= True
   Rule_Boolliteral_False                = 28  ' <BoolLiteral> ::= False
   Rule_Intliteral_Intliteral            = 29  ' <IntLiteral> ::= IntLiteral
   Rule_Intliteral_Hexliteral            = 30  ' <IntLiteral> ::= HexLiteral
   Rule_Intliteral_Octliteral            = 31  ' <IntLiteral> ::= OctLiteral
   Rule_Nothing_Nothing                  = 32  ' <Nothing> ::= Nothing
   Rule_Nothing_Null                     = 33  ' <Nothing> ::= Null
   Rule_Nothing_Empty                    = 34  ' <Nothing> ::= Empty
   Rule_Subdecl_Sub_Id_End_Sub           = 35  ' <SubDecl> ::= <AccessModifierOpt> Sub ID <MethodArgList> <NL> <MethodStmtList> End Sub <NL>
   Rule_Subdecl_Function_Id_End_Function = 36  ' <SubDecl> ::= <AccessModifierOpt> Function ID <MethodArgList> <NL> <MethodStmtList> End Function <NL>
   Rule_Methodarglist_Lparan_Rparan      = 37  ' <MethodArgList> ::= '(' <ArgList> ')'
   Rule_Methodarglist_Lparan_Rparan2     = 38  ' <MethodArgList> ::= '(' ')'
   Rule_Arglist_Comma                    = 39  ' <ArgList> ::= <Arg> ',' <ArgList>
   Rule_Arglist                          = 40  ' <ArgList> ::= <Arg>
   Rule_Arg_Id                           = 41  ' <Arg> ::= <ArgModifier> ID
   Rule_Argmodifier_Byval                = 42  ' <ArgModifier> ::= ByVal
   Rule_Argmodifier_Byref                = 43  ' <ArgModifier> ::= ByRef
   Rule_Argmodifier                      = 44  ' <ArgModifier> ::= 
   Rule_Blockstmt                        = 45  ' <BlockStmt> ::= <InlineStmt> <NL>
   Rule_Blockstmt2                       = 46  ' <BlockStmt> ::= <IfStmt>
   Rule_Blockstmt3                       = 47  ' <BlockStmt> ::= <LoopStmt>
   Rule_Blockstmtlist                    = 48  ' <BlockStmtList> ::= <BlockStmt> <BlockStmtList>
   Rule_Blockstmtlist2                   = 49  ' <BlockStmtList> ::= 
   Rule_Methodstmtlist                   = 50  ' <MethodStmtList> ::= <MethodStmt> <MethodStmtList>
   Rule_Methodstmtlist2                  = 51  ' <MethodStmtList> ::= 
   Rule_Methodstmt                       = 52  ' <MethodStmt> ::= <VarDecl>
   Rule_Methodstmt2                      = 53  ' <MethodStmt> ::= <BlockStmt>
   Rule_Inlinestmt                       = 54  ' <InlineStmt> ::= <AssignStmt>
   Rule_Inlinestmt2                      = 55  ' <InlineStmt> ::= <SubCallStmt>
   Rule_Inlinestmt_Exit_Lparan_Rparan    = 56  ' <InlineStmt> ::= Exit '(' ')'
   Rule_Assignstmt_Eq                    = 57  ' <AssignStmt> ::= <QualifiedID> '=' <Expr>
   Rule_Assignstmt_Pluseq                = 58  ' <AssignStmt> ::= <QualifiedID> '+=' <Expr>
   Rule_Assignstmt_Minuseq               = 59  ' <AssignStmt> ::= <QualifiedID> '-=' <Expr>
   Rule_Assignstmt_Plusplus              = 60  ' <AssignStmt> ::= <QualifiedID> '++'
   Rule_Assignstmt_Minusminus            = 61  ' <AssignStmt> ::= <QualifiedID> '--'
   Rule_Subcallstmt_Lparan_Rparan        = 62  ' <SubCallStmt> ::= <QualifiedID> '(' <ExprList> ')'
   Rule_Qualifiedid_Id                   = 63  ' <QualifiedID> ::= ID
   Rule_Leftexpr                         = 64  ' <LeftExpr> ::= <QualifiedID>
   Rule_Leftexpr_Lparan_Rparan           = 65  ' <LeftExpr> ::= <QualifiedID> '(' <ExprList> ')'
   Rule_Exprlist                         = 66  ' <ExprList> ::= <Expr> <CommaExprList>
   Rule_Exprlist2                        = 67  ' <ExprList> ::= <Expr>
   Rule_Exprlist3                        = 68  ' <ExprList> ::= 
   Rule_Commaexprlist_Comma              = 69  ' <CommaExprList> ::= ',' <Expr> <CommaExprList>
   Rule_Commaexprlist_Comma2             = 70  ' <CommaExprList> ::= ',' <Expr>
   Rule_Expr                             = 71  ' <Expr> ::= <ImpExpr>
   Rule_Impexpr_Imp                      = 72  ' <ImpExpr> ::= <ImpExpr> Imp <EqvExpr>
   Rule_Impexpr                          = 73  ' <ImpExpr> ::= <EqvExpr>
   Rule_Eqvexpr_Eqv                      = 74  ' <EqvExpr> ::= <EqvExpr> Eqv <XorExpr>
   Rule_Eqvexpr                          = 75  ' <EqvExpr> ::= <XorExpr>
   Rule_Xorexpr_Xor                      = 76  ' <XorExpr> ::= <XorExpr> Xor <OrExpr>
   Rule_Xorexpr                          = 77  ' <XorExpr> ::= <OrExpr>
   Rule_Orexpr_Or                        = 78  ' <OrExpr> ::= <OrExpr> Or <AndExpr>
   Rule_Orexpr                           = 79  ' <OrExpr> ::= <AndExpr>
   Rule_Andexpr_And                      = 80  ' <AndExpr> ::= <AndExpr> And <NotExpr>
   Rule_Andexpr                          = 81  ' <AndExpr> ::= <NotExpr>
   Rule_Notexpr_Not                      = 82  ' <NotExpr> ::= Not <NotExpr>
   Rule_Notexpr                          = 83  ' <NotExpr> ::= <CompareExpr>
   Rule_Compareexpr_Gteq                 = 84  ' <CompareExpr> ::= <CompareExpr> '>=' <ConcatExpr>
   Rule_Compareexpr_Lteq                 = 85  ' <CompareExpr> ::= <CompareExpr> '<=' <ConcatExpr>
   Rule_Compareexpr_Gt                   = 86  ' <CompareExpr> ::= <CompareExpr> '>' <ConcatExpr>
   Rule_Compareexpr_Lt                   = 87  ' <CompareExpr> ::= <CompareExpr> '<' <ConcatExpr>
   Rule_Compareexpr_Ltgt                 = 88  ' <CompareExpr> ::= <CompareExpr> '<>' <ConcatExpr>
   Rule_Compareexpr_Eq                   = 89  ' <CompareExpr> ::= <CompareExpr> '=' <ConcatExpr>
   Rule_Compareexpr                      = 90  ' <CompareExpr> ::= <ConcatExpr>
   Rule_Concatexpr_Amp                   = 91  ' <ConcatExpr> ::= <ConcatExpr> '&' <AddExpr>
   Rule_Concatexpr                       = 92  ' <ConcatExpr> ::= <AddExpr>
   Rule_Addexpr_Plus                     = 93  ' <AddExpr> ::= <AddExpr> '+' <ModExpr>
   Rule_Addexpr_Minus                    = 94  ' <AddExpr> ::= <AddExpr> '-' <ModExpr>
   Rule_Addexpr                          = 95  ' <AddExpr> ::= <ModExpr>
   Rule_Modexpr_Mod                      = 96  ' <ModExpr> ::= <ModExpr> Mod <IntDivExpr>
   Rule_Modexpr                          = 97  ' <ModExpr> ::= <IntDivExpr>
   Rule_Intdivexpr_Backslash             = 98  ' <IntDivExpr> ::= <IntDivExpr> '\' <MultExpr>
   Rule_Intdivexpr                       = 99  ' <IntDivExpr> ::= <MultExpr>
   Rule_Multexpr_Times                   = 100 ' <MultExpr> ::= <MultExpr> '*' <UnaryExpr>
   Rule_Multexpr_Div                     = 101 ' <MultExpr> ::= <MultExpr> '/' <UnaryExpr>
   Rule_Multexpr                         = 102 ' <MultExpr> ::= <UnaryExpr>
   Rule_Unaryexpr_Minus                  = 103 ' <UnaryExpr> ::= '-' <UnaryExpr>
   Rule_Unaryexpr_Plus                   = 104 ' <UnaryExpr> ::= '+' <UnaryExpr>
   Rule_Unaryexpr                        = 105 ' <UnaryExpr> ::= <ExpExpr>
   Rule_Expexpr_Caret                    = 106 ' <ExpExpr> ::= <Value> '^' <ExpExpr>
   Rule_Expexpr                          = 107 ' <ExpExpr> ::= <Value>
   Rule_Value                            = 108 ' <Value> ::= <ConstExpr>
   Rule_Value2                           = 109 ' <Value> ::= <LeftExpr>
   Rule_Value_Lparan_Rparan              = 110 ' <Value> ::= '(' <Expr> ')'
   Rule_Ifstmt_If_Then_End_If            = 111 ' <IfStmt> ::= If <Expr> Then <NL> <BlockStmtList> <ElseStmtList> End If <NL>
   Rule_Ifstmt_If_Then                   = 112 ' <IfStmt> ::= If <Expr> Then <InlineStmt> <ElseOpt> <EndIfOpt> <NL>
   Rule_Elsestmtlist_Elseif_Then         = 113 ' <ElseStmtList> ::= ElseIf <Expr> Then <NL> <BlockStmtList> <ElseStmtList>
   Rule_Elsestmtlist_Else                = 114 ' <ElseStmtList> ::= Else <NL> <BlockStmtList>
   Rule_Elsestmtlist                     = 115 ' <ElseStmtList> ::= 
   Rule_Elseopt_Else                     = 116 ' <ElseOpt> ::= Else <InlineStmt>
   Rule_Elseopt                          = 117 ' <ElseOpt> ::= 
   Rule_Endifopt_End_If                  = 118 ' <EndIfOpt> ::= End If
   Rule_Endifopt                         = 119 ' <EndIfOpt> ::= 
   Rule_Loopstmt_While_Wend              = 120 ' <LoopStmt> ::= While <Expr> <NL> <BlockStmtList> WEnd <NL>
End Enum
