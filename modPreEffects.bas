Attribute VB_Name = "modPreEffects"
'*************************************************************************
'BrightNess Contrast Saturation

Private Type THistSingle
    a(0 To 255)    As Single
End Type

Private histR(0 To 255) As Long
Private histG(0 To 255) As Long
Private histB(0 To 255) As Long

Private OutPUT()   As Byte

'*************************************************************************

Public Function MagneKleverBCS(ByVal BRIGHT As Long, ByVal CONTRAST As Long, ByVal SATURATION As Long)
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre

    Dim CR         As Long
    Dim Cg         As Long
    Dim cB         As Long
    Dim X          As Long
    Dim Y          As Long

    Dim i          As Long
    Dim k          As Long
    Dim V          As Long
    Dim ci1        As Long
    Dim ci2        As Long
    Dim ci3        As Long

    Dim Alpha      As Long
    Dim a          As Single



    Dim ContrastLut(0 To 255) As Long
    Dim BCLut(0 To 255) As Long
    Dim SATGrays(0 To 767) As Long
    Dim SATAlpha(0 To 255) As Long


    Dim ProgX      As Long
    Dim ProgXStep  As Long


    ReDim OutPUT(0 To 2, 0 To pW, 0 To pH) As Byte


    If CONTRAST = 100 Then CONTRAST = 99
    If CONTRAST > 0 Then
        a = 1 / Cos(CONTRAST * (PI / 200))
    Else
        a = 1 * Cos(CONTRAST * (PI / 200))
    End If

    For i = 0 To 255
        V = Round(a * (i - 170) + 170)
        If V > 255 Then V = 255 Else If V < 0 Then V = 0
        ContrastLut(i) = V
    Next


    Alpha = BRIGHT
    For i = 0 To 255
        k = 256 - Alpha
        V = (k + Alpha * i) \ 256
        If V < 0 Then V = 0 Else If V > 255 Then V = 255
        BCLut(i) = ContrastLut(V)
    Next

    For i = 0 To 255
        SATAlpha(i) = (((i + 1) * SATURATION) \ 256)
    Next i

    X = 0
    For i = 0 To 255
        Y = i - SATAlpha(i)
        SATGrays(X) = Y
        X = X + 1
        SATGrays(X) = Y
        X = X + 1
        SATGrays(X) = Y
        X = X + 1
    Next


    For X = 0 To pW

        For Y = 0 To pH

            CR = SOURCEbyte(2, X, Y)
            Cg = SOURCEbyte(1, X, Y)
            cB = SOURCEbyte(0, X, Y)

            V = CR + Cg + cB

            ci1 = SATGrays(V) + SATAlpha(cB)
            ci2 = SATGrays(V) + SATAlpha(Cg)
            ci3 = SATGrays(V) + SATAlpha(CR)
            If ci1 < 0 Then ci1 = 0 Else If ci1 > 255 Then ci1 = 255
            If ci2 < 0 Then ci2 = 0 Else If ci2 > 255 Then ci2 = 255
            If ci3 < 0 Then ci3 = 0 Else If ci3 > 255 Then ci3 = 255
            OutPUT(0, X, Y) = BCLut(ci1)
            OutPUT(1, X, Y) = BCLut(ci2)
            OutPUT(2, X, Y) = BCLut(ci3)

        Next
        If X > ProgX Then
            FX_PercDONE "BCS", X / pW, 0
            ProgX = ProgX + ProgXStep
        End If


    Next

    SOURCEbyte = OutPUT

    Erase OutPUT
    FX_PercDONE "BCS", 1, 0
End Function


Public Function MagneKleverExposure(k As Single)
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre

    Dim i          As Long
    Dim X          As Long
    Dim Y          As Long
    Dim LUT(0 To 255) As Long

    Dim ProgX      As Long
    Dim ProgXStep  As Long


    ReDim OutPUT(0 To 2, 0 To pW, 0 To pH) As Byte



    For i = 0 To 255
        If k < 0 Then
            LUT(i) = i - ((-Round((1 - Exp((i / -128) * (k / 128))) * 256) * (i Xor 255)) \ 256)
        Else
            LUT(i) = i + ((Round((1 - Exp((i / -128) * (k / 128))) * 256) * (i Xor 255)) \ 256)
        End If

        If LUT(i) < 0 Then LUT(i) = 0 Else If LUT(i) > 255 Then LUT(i) = 255
    Next

    ProgXStep = Round(3 * pW / 100)
    ProgX = 0


    For X = 0 To pW
        For Y = 0 To pH
            OutPUT(2, X, Y) = LUT(SOURCEbyte(2, X, Y))
            OutPUT(1, X, Y) = LUT(SOURCEbyte(1, X, Y))
            OutPUT(0, X, Y) = LUT(SOURCEbyte(0, X, Y))
        Next
        If X > ProgX Then
            FX_PercDONE "Exposure", X / pW, 0
            ProgX = ProgX + ProgXStep
        End If
    Next

    SOURCEbyte = OutPUT
    Erase OutPUT

    FX_PercDONE "Exposure", 1, 0

End Function


Public Sub MagneKleverfxHistCalc()
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre

    Dim X          As Long
    Dim Y          As Long
    For X = 0 To 255
        histR(X) = 0
        histG(X) = 0
        histB(X) = 0
    Next
    For Y = 0 To pH
        For X = 0 To pW
            histR(SOURCEbyte(2, X, Y)) = histR(SOURCEbyte(2, X, Y)) + 1
            histG(SOURCEbyte(1, X, Y)) = histG(SOURCEbyte(1, X, Y)) + 1
            histB(SOURCEbyte(0, X, Y)) = histB(SOURCEbyte(0, X, Y)) + 1
        Next

    Next


End Sub


Private Function MagneKleverCumSum(Hist As THistSingle) As THistSingle
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre
    Dim X          As Long
    Dim temp       As THistSingle

    temp.a(0) = Hist.a(0)
    For X = 1 To 255
        temp.a(X) = temp.a(X - 1) + Hist.a(X)
    Next

    MagneKleverCumSum = temp
End Function
Public Sub MagneKleverHistogramEQU(Z As Single)
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre

    Dim X          As Long
    Dim Y          As Long

    Dim Q1         As Single
    Dim Q2         As Single
    Dim Q3         As Single

    Dim Hist       As THistSingle
    Dim VCumSumR   As THistSingle
    Dim VCumSumG   As THistSingle
    Dim VCumSumB   As THistSingle

    Dim ProgX      As Long
    Dim ProgXStep  As Long


    ReDim OutPUT(0 To 2, 0 To pW, 0 To pH) As Byte

    MagneKleverfxHistCalc

    Q1 = 0                        '// RED Channel
    For X = 0 To 255
        Hist.a(X) = histR(X) ^ Z
        Q1 = Q1 + Hist.a(X)
    Next
    VCumSumR = MagneKleverCumSum(Hist)

    Q2 = 0
    For X = 0 To 255
        Hist.a(X) = histG(X) ^ Z
        Q2 = Q2 + Hist.a(X)
    Next
    VCumSumG = MagneKleverCumSum(Hist)

    Q3 = 0
    For X = 0 To 255
        Hist.a(X) = histB(X) ^ Z
        Q3 = Q3 + Hist.a(X)
    Next
    VCumSumB = MagneKleverCumSum(Hist)


    ProgXStep = Round(3 * pW / 100)
    ProgX = 0



    For X = 0 To pW
        For Y = 0 To pH


            OutPUT(2, X, Y) = Fix((255 / Q1) * VCumSumR.a(SOURCEbyte(2, X, Y)))
            OutPUT(1, X, Y) = Fix((255 / Q2) * VCumSumG.a(SOURCEbyte(1, X, Y)))
            OutPUT(0, X, Y) = Fix((255 / Q3) * VCumSumB.a(SOURCEbyte(0, X, Y)))

            '            RGB[x].B := Trunc((255 / q3) * vcumsumB[RGB[x].B]);
        Next
        If X > ProgX Then
            'RaiseEvent PercDONE("Equalize", x / pW, 0)
            ProgX = ProgX + ProgXStep
        End If
    Next
    SOURCEbyte = OutPUT
    Erase OutPUT
    'RaiseEvent PercDONE("Equalize", 1, 0)

End Sub






