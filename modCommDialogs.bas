Attribute VB_Name = "modCommDialogs"
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Type TCHOOSEFONT
    lStructSize As Long         ' Filled with UDT size
    hWndOwner As Long           ' Caller's window handle
    Hdc As Long                 ' Printer DC/IC or NULL
    lpLogFont As Long           ' Pointer to LOGFONT
    iPointSize As Long          ' 10 * size in points of font
    Flags As Long               ' Type flags
    rgbColors As Long           ' Returned text color
    lCustData As Long           ' Data passed to hook function
    lpfnHook As Long            ' Pointer to hook function
    lpTemplateName As Long      ' Custom template name
    hInstance As Long           ' Instance handle for template
    lpszStyle As String         ' Return style field
    nFontType As Integer        ' Font type bits
    iAlign As Integer           ' Filler
    nSizeMin As Long            ' Minimum point size allowed
    nSizeMax As Long            ' Maximum point size allowed
End Type

Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Enum EChooseFont
    CF_SCREENFONTS = &H1
    CF_PRINTERFONTS = &H2
    CF_BOTH = &H3
    CF_FontShowHelp = &H4
    CF_UseStyle = &H80
    CF_EFFECTS = &H100
    CF_AnsiOnly = &H400
    CF_NoVectorFonts = &H800
    CF_NoOemFonts = CF_NoVectorFonts
    CF_NoSimulations = &H1000
    CF_LIMITSIZE = &H2000
    CF_FixedPitchOnly = &H4000
    CF_WYSIWYG = &H8000  ' Must also have ScreenFonts And PrinterFonts
    CF_FORCEFONTEXIST = &H10000
    CF_ScalableOnly = &H20000
    CF_TTOnly = &H40000
    CF_NoFaceSel = &H80000
    CF_NoStyleSel = &H100000
    CF_NoSizeSel = &H200000
    ' Win95 only
    CF_SelectScript = &H400000
    CF_NoScriptSel = &H800000
    CF_NoVertFonts = &H1000000

    CF_INITTOLOGFONTSTRUCT = &H40
    CF_Apply = &H200
    CF_EnableHook = &H8
    CF_EnableTemplate = &H10
    CF_EnableTemplateHandle = &H20
    CF_FontNotSupported = &H238
End Enum

' These are extra nFontType bits that are added to what is returned to the
' EnumFonts callback routine
Private Enum EFontType
    Simulated_FontType = &H8000
    Printer_FontType = &H4000
    Screen_FontType = &H2000
    Bold_FontType = &H100
    Italic_FontType = &H200
    REGULAR_FONTTYPE = &H400
End Enum

Private Declare Function ChooseFont Lib "COMDLG32" Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_EDITBOX = &H10
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)

Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function OleInitialize Lib "ole32.dll" (ByVal lRes As Long) As Long
Private Declare Sub OleUninitialize Lib "ole32.dll" ()

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private CustomColors() As Byte
Private Declare Sub InitCommonControls Lib "COMCTL32" ()

Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXPLORER = &H80000                         '  new look commdlg

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'Per la rimozione della x di chiusura della finestra
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Sub RemoveSysMenus(ByVal frm As Form, _
    ByVal remove_restore As Boolean, _
    ByVal remove_move As Boolean, _
    ByVal remove_size As Boolean, _
    ByVal remove_minimize As Boolean, _
    ByVal remove_maximize As Boolean, _
    ByVal remove_separator As Boolean, _
    ByVal remove_close As Boolean)
    Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(frm.hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_separator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

Public Sub InitDialogs()
    InitCommonControls

    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    Dim i As Integer
    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
    Next i
End Sub

Public Function ShowColor(ofrm As Form) As Long
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hWndOwner = ofrm.hwnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.Flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function

Public Function ShowOpen(ofrm As Form, sFilter As String, _
sTitle As String, lFlags As Long, sStartDir As String) As String

    Dim OFName As OPENFILENAME
    
    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hWndOwner = ofrm.hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    'OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFilter = sFilter
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = sStartDir
    'Set the dialog title
    OFName.lpstrTitle = sTitle
    'no extra flags
    OFName.Flags = lFlags

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        ShowOpen = Trim(OFName.lpstrFile)
        ShowOpen = Left(ShowOpen, Len(ShowOpen) - 1)
    Else
        ShowOpen = ""
    End If
End Function

Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function
Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        CopyMemoryStr ab(LBound(ab)), s, cab
    End If
End Sub
Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

' ChooseFont wrapper
Private Function VBChooseFont(CurFont As Font, _
                      Optional PrinterDC As Long = -1, _
                      Optional Owner As Long = -1, _
                      Optional Color As Long = vbBlack, _
                      Optional MinSize As Long = 0, _
                      Optional MaxSize As Long = 0, _
                      Optional Flags As Long = 0 _
                    ) As Boolean
    Dim lr As Long

    ' Unwanted Flags bits
    Const CF_FontNotSupported = CF_Apply Or CF_EnableHook Or CF_EnableTemplate
    
    ' Flags can get reference variable or constant with bit flags
    ' PrinterDC can take printer DC
    If PrinterDC = -1 Then
        PrinterDC = 0
        If Flags And CF_PRINTERFONTS Then PrinterDC = Printer.Hdc
    Else
        Flags = Flags Or CF_PRINTERFONTS
    End If
    ' Must have some fonts
    If (Flags And CF_PRINTERFONTS) = 0 Then Flags = Flags Or CF_SCREENFONTS
    ' Color can take initial color, receive chosen color
    If Color <> vbBlack Then Flags = Flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If MinSize Then Flags = Flags Or CF_LIMITSIZE
    ' MaxSize can be maximum size accepted
    If MaxSize Then Flags = Flags Or CF_LIMITSIZE
    
    ' Put in required internal flags and remove unsupported
    Flags = (Flags Or CF_INITTOLOGFONTSTRUCT) And Not CF_FontNotSupported
    
    ' Initialize LOGFONT variable
    Dim fnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    fnt.lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = CurFont.Weight
    fnt.lfItalic = CurFont.Italic
    fnt.lfUnderline = CurFont.Underline
    fnt.lfStrikeOut = CurFont.Strikethrough
    ' Other fields zero
    StrToBytes fnt.lfFaceName, CurFont.Name
    
    ' Initialize TCHOOSEFONT variable
    Dim cf As TCHOOSEFONT
    cf.lStructSize = Len(cf)
    If Owner <> -1 Then cf.hWndOwner = Owner
    cf.Hdc = PrinterDC
    cf.lpLogFont = VarPtr(fnt)
    cf.iPointSize = CurFont.Size * 10
    cf.Flags = Flags
    cf.rgbColors = Color
    cf.nSizeMin = MinSize
    cf.nSizeMax = MaxSize
        
    ' All other fields zero
    lr = ChooseFont(cf)
    Select Case lr
    Case 1
        ' Success
        VBChooseFont = True
        Flags = cf.Flags
        Color = cf.rgbColors
        CurFont.Bold = cf.nFontType And Bold_FontType
        'CurFont.Italic = cf.nFontType And Italic_FontType
        CurFont.Italic = fnt.lfItalic
        CurFont.Strikethrough = fnt.lfStrikeOut
        CurFont.Underline = fnt.lfUnderline
        CurFont.Weight = fnt.lfWeight
        CurFont.Size = cf.iPointSize / 10
        CurFont.Name = BytesToStr(fnt.lfFaceName)
    Case 0
        ' Cancelled
        VBChooseFont = False
    Case Else
        ' Extended error
        VBChooseFont = False
    End Select
        
End Function

Public Function ShowFont(ofrm As Form, _
ByRef sFontName As String, ByRef lFontColor As Long, _
ByRef bFontBold As Boolean, ByRef bFontItalic As Boolean, _
ByRef bFontUnderline As Boolean, ByRef bFontStrikeout As Boolean, _
ByRef iSize As Integer) As Boolean

    Dim oFnt As StdFont
    Dim bRet As Boolean
    
    Set oFnt = New StdFont
    If Trim(sFontName) <> "" Then oFnt.Name = sFontName
    oFnt.Bold = bFontBold
    oFnt.Italic = bFontItalic
    If iSize > 0 Then oFnt.Size = iSize
    oFnt.Strikethrough = bFontStrikeout
    oFnt.Underline = bFontUnderline
    
    bRet = VBChooseFont(oFnt, , , lFontColor)
    If bRet Then
        sFontName = oFnt.Name
        bFontBold = oFnt.Bold
        bFontItalic = oFnt.Italic
        bFontUnderline = oFnt.Underline
        bFontStrikeout = oFnt.Strikethrough
        iSize = oFnt.Size
    End If
    
    ShowFont = bRet
End Function

Public Function ShowSave(ofrm As Form, sFilter As String, _
sTitle As String, lFlags As Long, sStartDir As String, _
sDefaultFile As String, ByRef iFilterIndex As Integer) As String
    
    Dim OFName As OPENFILENAME
    
    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hWndOwner = ofrm.hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    'OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFilter = sFilter
    'Create a buffer
    OFName.lpstrFile = sDefaultFile & Space(254 - Len(sDefaultFile))
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = sStartDir
    'Set the dialog title
    OFName.lpstrTitle = sTitle
    'no extra flags
    OFName.Flags = lFlags
    If iFilterIndex <> 0 Then
        'Gli do' l'indice iniziale
        OFName.nFilterIndex = iFilterIndex
    End If

    'Show the 'Save File'-dialog
    If GetSaveFileName(OFName) Then
        ShowSave = Trim(OFName.lpstrFile)
        ShowSave = Left(ShowSave, Len(ShowSave) - 1)
        iFilterIndex = OFName.nFilterIndex
    Else
        ShowSave = ""
        iFilterIndex = 0
    End If
End Function

Public Function ShowDir(hwndParent As Long)
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    
    ShowDir = ""

    With udtBI
        'Set the owner window
        .hWndOwner = hwndParent
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With

    'Obbligatoria se uso BIF_USENEWUI
    OleInitialize 0

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(256, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
    'Chiudo
    OleUninitialize
    
    If Len(sPath) > 0 Then
        ShowDir = sPath
    End If
End Function

Public Function GetCustomColorsString() As String
    Dim s As String
    Dim i As Integer
    
    s = ""
    For i = LBound(CustomColors) To UBound(CustomColors)
        s = s & Hex(CustomColors(i)) & "*"
    Next
    
    GetCustomColorsString = s
End Function

Public Sub SetCustomColorsString(sColors As String)
    On Error Resume Next
    
    Dim v As Variant
    Dim i As Integer
        
    v = Split(sColors, "*")
    If UBound(v) > 0 Then
        For i = 0 To UBound(v) - 1
            CustomColors(i) = CByte("&h" & v(i))
        Next
    End If
End Sub

