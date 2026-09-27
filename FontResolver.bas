Attribute VB_Name = "FontResolver"
Option Explicit

Private m_FontCache As Dictionary
Public g_DefaultFontName As String '暂存系统默认字体名，避免每次查询

Private Type LogFontW
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
    lfFaceName(0 To 63) As Byte
End Type

Private Declare Sub copyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LogFontW) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetFontData Lib "gdi32" (ByVal hdc As Long, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetObjectW Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Const DEFAULT_GUI_FONT = 17
Private Const LF_FACESIZE = 32

Private Const NAME_TABLE_TAG As Long = &H656D616E ' 'name' 表标签

'获取系统默认字体名
Public Function GetDefaultFontName() As String
    If Len(g_DefaultFontName) > 0 Then
        GetDefaultFontName = g_DefaultFontName
        Exit Function
    End If

    Dim hFont As Long
    hFont = GetStockObject(DEFAULT_GUI_FONT)
    If hFont = 0 Then Exit Function

    Dim lf As LogFontW
    ' 使用 GetObjectW 填充 Unicode 结构体
    If GetObjectW(hFont, LenB(lf), lf) <> 0 Then
        ' 计算以 Null 结尾的 Unicode 字符长度
        Dim cch As Long
        cch = lstrlenW(VarPtr(lf.lfFaceName(0)))
        
        If cch > 0 Then
            ' 将字节直接作为 UTF-16 字符串读取，无需 StrConv
            g_DefaultFontName = GetEnglishFontName(Left$(CStr(lf.lfFaceName), cch))
        End If
    End If
    
    ' 注：StockObject 无需 DeleteObject
    GetDefaultFontName = g_DefaultFontName
End Function

Public Function GetEnglishFontName(localName As String) As String
    Dim logFont As LogFontW
    Dim nameLen As Long
    Dim hdc As Long, hFont As Long, hOldFont As Long
    
    Dim sKey As String
    sKey = Trim$(localName)
    If Len(sKey) = 0 Then
        GetEnglishFontName = ""
        Exit Function
    End If
    
    InitFontCache
    
    ' 查询缓存：若存在则直接返回
    If m_FontCache.Exists(sKey) Then
        GetEnglishFontName = m_FontCache.Item(sKey)
        Exit Function
    End If

    nameLen = Len(sKey)
    If nameLen > 31 Then nameLen = 31 ' LF_FACESIZE-1，留一个WCHAR给null terminator
    copyMemory logFont.lfFaceName(0), ByVal StrPtr(sKey), nameLen * 2
    logFont.lfCharSet = 1 ' DEFAULT_CHARSET

    hdc = GetDC(0)
    hFont = CreateFontIndirectW(logFont)
    hOldFont = SelectObject(hdc, hFont)

    Dim tableSize As Long
    tableSize = GetFontData(hdc, NAME_TABLE_TAG, 0, ByVal 0&, 0)
    If tableSize = -1 Or tableSize = 0 Then
        SelectObject hdc, hOldFont
        DeleteObject hFont
        ReleaseDC 0, hdc
        GetEnglishFontName = sKey
        Exit Function
    End If
    
    Dim resolvedName As String
    Dim buf() As Byte
    ReDim buf(tableSize - 1)
    GetFontData hdc, NAME_TABLE_TAG, 0, buf(0), tableSize

    SelectObject hdc, hOldFont
    DeleteObject hFont
    ReleaseDC 0, hdc

    resolvedName = parseNameTable(buf)
    If resolvedName = "" Then
        resolvedName = sKey
    End If
    
    m_FontCache.Add sKey, resolvedName
    GetEnglishFontName = resolvedName
End Function

' 初始化缓存实例
Private Sub InitFontCache()
    If m_FontCache Is Nothing Then
        Set m_FontCache = New Dictionary
    End If
End Sub

Private Function readUInt16(buf() As Byte, offset As Long) As Long
    readUInt16 = buf(offset) * 256 + buf(offset + 1) ' name表内为big-endian
End Function

Private Function parseNameTable(buf() As Byte) As String
    Dim recordCount As Long, stringOffset As Long
    recordCount = readUInt16(buf, 2)
    stringOffset = readUInt16(buf, 4)

    Dim i As Long, recOffset As Long
    Dim platformId As Long, languageId As Long, nameId As Long, strLen As Long, strOfs As Long
    Dim bestOffset As Long, bestLen As Long, found As Boolean

    For i = 0 To recordCount - 1
        recOffset = 6 + i * 12
        platformId = readUInt16(buf, recOffset)
        languageId = readUInt16(buf, recOffset + 4)
        nameId = readUInt16(buf, recOffset + 6)
        strLen = readUInt16(buf, recOffset + 8)
        strOfs = readUInt16(buf, recOffset + 10)

        If nameId = 1 And platformId = 3 And languageId = &H409 Then
            bestOffset = stringOffset + strOfs
            bestLen = strLen
            found = True
            Exit For
        End If
    Next i

    If Not found Then Exit Function

    Dim s As String, ch As Integer, j As Long
    For j = 0 To bestLen - 1 Step 2
        ch = buf(bestOffset + j) * 256 + buf(bestOffset + j + 1)
        s = s & ChrW(ch)
    Next j
    parseNameTable = s
End Function

