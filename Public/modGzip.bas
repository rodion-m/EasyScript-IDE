Attribute VB_Name = "modGzip"
Option Explicit

Private Declare Function InitDecompression Lib "E:\Programming\Заказы\shironet.mako.co.il Parser\gzip.dll" () As Long
Private Declare Function CreateDecompression Lib "E:\Programming\Заказы\shironet.mako.co.il Parser\gzip.dll" (ByRef context As Long, ByVal Flags As Long) As Long
Private Declare Function Decompress Lib "E:\Programming\Заказы\shironet.mako.co.il Parser\gzip.dll" (ByVal context As Long, inBytes As Any, ByVal input_size As Long, outBytes As Any, ByVal output_size As Long, ByRef input_used As Long, ByRef output_used As Long) As Long
Private Declare Function DestroyDecompression Lib "E:\Programming\Заказы\shironet.mako.co.il Parser\gzip.dll" (ByRef context As Long) As Long
Private Declare Function ResetDecompression Lib "E:\Programming\Заказы\shironet.mako.co.il Parser\gzip.dll" (ByVal context As Long) As Long

Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)

Public Property Get DecompressGZip(ByVal ByteString As String) As String
    Dim lngBuffer As Long, strBuffer As String
    Dim lngContext As Long
    Dim lngPos As Long, lngLen As Long
    Dim lngInUsed As Long, lngOutUsed As Long
    ' create a buffer of 64 kB = this is much faster than Space$(32768)
    lngBuffer = SysAllocStringByteLen(0, 65536)
    PutMem4 VarPtr(strBuffer), lngBuffer
    ' initialize GZIP decompression & get handle
    InitDecompression
    CreateDecompression lngContext, 1
    ' start position & original length
    lngPos = StrPtr(ByteString)
    lngLen = LenB(ByteString)
    ' do decompression while success
    Do While 0 = Decompress(lngContext, ByVal lngPos, lngLen, ByVal lngBuffer, LenB(strBuffer), lngInUsed, lngOutUsed)
        ' did we get any data?
        If lngOutUsed Then
            ' create final output string (note: String = String & String = performance bottleneck)
            DecompressGZip = DecompressGZip & StrConv(LeftB$(strBuffer, lngOutUsed), vbUnicode)
        End If
        ' reduce amount of data processed
        lngLen = lngLen - lngInUsed
        ' exit loop if nothing more to do
        If lngLen < 1 Then Exit Do
        ' move pointer
        lngPos = lngPos + lngInUsed
    Loop
    DecompressGZip = LeftB(DecompressGZip, LenB(DecompressGZip) - 1) ' Функция всегда добавляет Chr(0) в конец -- убираю его.
    ' we are done, close decompression handle
    ResetDecompression lngContext
End Property
