Attribute VB_Name = "Base64"
Option Explicit

Private aDecTab(255) As Integer
Private Const sEncTab As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Public Function EncodeStr64(ByRef sInput As String) As String
    Dim sOutput As String, sLast As String
    Dim b(2)    As Byte
    Dim j       As Integer
    Dim i       As Long, nLen As Long, nQuants As Long
    Dim iIndex  As Long
    
    nLen = Len(sInput)
    nQuants = nLen \ 3
    sOutput = String(nQuants * 4, " ")

    For i = 0 To nQuants - vbNull
        For j = 0 To 2
            b(j) = VBA.Asc(VBA.Mid$(sInput, (i * 3) + j + vbNull, vbNull))
        Next

        Mid(sOutput, iIndex + vbNull, 4) = EncodeQuantum(b)
        iIndex = iIndex + 4
        DoEvents
    Next

    Select Case nLen Mod 3
        Case 0
            sLast = vbNullString

        Case 1
            b(0) = VBA.Asc(VBA.Mid$(sInput, nLen, vbNull))
            b(1) = 0&
            b(2) = 0&
            sLast = EncodeQuantum(b)
            sLast = VBA.Left$(sLast, 2) & "=="

        Case 2
            b(0) = VBA.Asc(VBA.Mid$(sInput, nLen - vbNull, vbNull))
            b(1) = VBA.Asc(VBA.Mid$(sInput, nLen, vbNull))
            b(2) = 0&
            sLast = EncodeQuantum(b)
            sLast = VBA.Left(sLast, 3) & "="
    End Select

    EncodeStr64 = sOutput & sLast
End Function

Public Function DecodeStr64(ByRef sEncoded As String) As String
    Dim d(3)        As Byte
    Dim c           As Byte
    Dim di          As Integer
    Dim i           As Long
    Dim nLen        As Long
    Dim iIndex      As Long
    
    nLen = Len(sEncoded)
    DecodeStr64 = String((nLen \ 4) * 3, " ")
    Call MakeDecTab

    For i = vbNull To Len(sEncoded)
        c = VBA.CByte(VBA.Asc(VBA.Mid$(sEncoded, i, vbNull)))
        c = aDecTab(c)

        If c >= 0& Then
            d(di) = c
            di = di + vbNull

            If di = 4 Then
                Mid$(DecodeStr64, iIndex + vbNull, 3) = DecodeQuantum(d)
                iIndex = iIndex + 3

                If d(3) = 64 Then
                    DecodeStr64 = VBA.Left(DecodeStr64, VBA.Len(DecodeStr64) - vbNull)
                    iIndex = iIndex - vbNull
                End If

                If d(2) = 64 Then
                    DecodeStr64 = VBA.Left(DecodeStr64, VBA.Len(DecodeStr64) - vbNull)
                    iIndex = iIndex - vbNull
                End If
                di = 0&
            End If
        End If
    DoEvents
    Next
End Function

Private Function EncodeQuantum(ByRef b() As Byte) As String
    Dim c As Integer

    c = SHR2(b(0)) And &H3F
    EncodeQuantum = EncodeQuantum & VBA.Mid$(sEncTab, c + vbNull, vbNull)

    c = SHL4(b(0) And &H3) Or (SHR4(b(1)) And &HF)
    EncodeQuantum = EncodeQuantum & VBA.Mid$(sEncTab, c + vbNull, vbNull)

    c = SHL2(b(1) And &HF) Or (SHR6(b(2)) And &H3)
    EncodeQuantum = EncodeQuantum & VBA.Mid$(sEncTab, c + vbNull, vbNull)

    c = b(2) And &H3F
    EncodeQuantum = EncodeQuantum & VBA.Mid$(sEncTab, c + vbNull, vbNull)
End Function

Private Function DecodeQuantum(ByRef d() As Byte) As String
    Dim c As Long

    c = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
    DecodeQuantum = DecodeQuantum & VBA.Chr$(c)

    c = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
    DecodeQuantum = DecodeQuantum & VBA.Chr$(c)

    c = SHL6(d(2) And &H3) Or d(3)
    DecodeQuantum = DecodeQuantum & VBA.Chr$(c)
End Function

Private Function MakeDecTab()
    Dim t As Integer
    Dim c As Integer

    For c = 0 To 255
        aDecTab(c) = &HFFF
    Next
    For c = VBA.Asc("A") To VBA.Asc("Z")
        aDecTab(c) = t
        t = t + vbNull
    Next
    For c = VBA.Asc("a") To VBA.Asc("z")
        aDecTab(c) = t
        t = t + vbNull
    Next
    For c = VBA.Asc("0") To VBA.Asc("9")
        aDecTab(c) = t
        t = t + vbNull
    Next

    c = Asc("+")
    aDecTab(c) = t
    t = t + vbNull

    c = Asc("/")
    aDecTab(c) = t
    t = t + vbNull
    
    c = Asc("=")
    aDecTab(c) = t
End Function

Private Function SHL2(ByVal bytValue As Byte) As Byte
    SHL2 = (bytValue * &H4) And &HFF
End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
    SHL4 = (bytValue * &H10) And &HFF
End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
    SHL6 = (bytValue * &H40) And &HFF
End Function

Private Function SHR2(ByVal bytValue As Byte) As Byte
    SHR2 = bytValue \ &H4
End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
    SHR4 = bytValue \ &H10
End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
    SHR6 = bytValue \ &H40
End Function

'Text2 = EncodeStr64(Text1)
'Text1 = DecodeStr64(Text2)



