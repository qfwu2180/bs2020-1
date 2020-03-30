Attribute VB_Name = "模块1"
Option Explicit

Public Const LF_FACESIZE = 32

Private fontNameCollection As New Collection

Type LOGFONT
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

Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" _
(ByVal hDC As Long, ByVal lpszFamily As String, _
ByVal lpEnumFontFamProc As Long, LParam As Any) As Long

Private Declare Function GetFocus Lib "User32" () As Long

Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Function OutputFontName(fontName As String)

    fontNameCollection.Add fontName, fontName
    
End Function

Private Function EnumFontFamProc(lpNLF As LOGFONT, _
lpNTM As NEWTEXTMETRIC, _
ByVal FontType As Long, _
LParam As Long) As Long

    On Error GoTo errorcode

    Dim FaceName As String

    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)

    OutputFontName Left$(FaceName, InStr(FaceName, vbNullChar) - 1)

    EnumFontFamProc = 1

Exit Function

errorcode:

    EnumFontFamProc = 1

End Function

Public Function ListAllFonts(Optional hWndTarget As Variant) As Collection
    
    Dim hDC As Long
    
    On Error GoTo Error_H

    If IsMissing(hWndTarget) Then hWndTarget = GetFocus

    hDC = GetDC(hWndTarget)

    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamProc, ByVal 0&

Finish:

    On Error Resume Next

    ReleaseDC hWndTarget, hDC

    Set ListAllFonts = fontNameCollection

Exit Function

Error_H:

    Resume Finish

End Function

Sub test()

    Dim fontList As Collection

    Set fontList = New Collection

    On Error Resume Next

    Set fontList = ListAllFonts()

    If fontList.Count > 0 Then

    Dim fontName

    For Each fontName In fontList

    Debug.Print fontName

    Next fontName

End If
On Error GoTo 0
End Sub

Sub GetFonts()

    Dim bar As CommandBar
    
    Set bar = CommandBars.Add(Name:="测试工具栏")
    
    Dim cbo As CommandBarComboBox
    
    Set cbo = bar.Controls.Add(Type:=msoControlComboBox)
    
    Dim fontList As Collection
    
    Set fontList = New Collection
    
    On Error Resume Next
    
    Set fontList = ListAllFonts()
    
    If fontList.Count > 0 Then
    
        Dim fontName
    
        With cbo
    
            .BeginGroup = True
        
            .Caption = "字体"
        
            .Style = msoComboLabel
        
            For Each fontName In fontList
        
                .AddItem fontName
            
            Next fontName
        
        End With
        
    End If
    
    bar.Visible = True
    

End Sub

Sub DelBar()
    
    DelToolBar "测试工具栏"
    
End Sub
Sub DelToolBar(ByVal barName As String)

    CommandBars(barName).Delete
    
End Sub
