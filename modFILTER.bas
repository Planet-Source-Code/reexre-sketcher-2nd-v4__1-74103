Attribute VB_Name = "modFILTER"
Option Explicit

Private Type tBK
    X              As Single
    Y              As Single
    a              As Single
    M              As Single
End Type


Private Type Bitmap
    bmType         As Long
    bmWidth        As Long
    bmHeight       As Long
    bmWidthBytes   As Long
    bmPlanes       As Integer
    bmBitsPixel    As Integer
    bmBits         As Long
End Type



Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long




Private hBmp       As Bitmap

Public pW          As Long
Public pH          As Long
Public pB          As Long


Public HUE()       As Single
Public HUE2()      As Single

Public SOURCEbyte() As Byte
Public SOURCEbyte2() As Byte
Public SOURCEbyte4() As Byte

Public OUTbyte()   As Byte
Public OUTsingle() As Single

Public EdgeAngle() As Single
Public EdgeMAG()   As Single


Public Back()      As tBK
Public BackBuff()  As tBK


Public MaxMag      As Single



Public gaborMatrix() As Single

Private Const SmoothPow As Single = 1    ' 4



'From WIKI:
'function gb=gabor_fn(sigma,theta,lambda,psi,gamma)
'
'sigma_x = sigma;
'sigma_y = sigma/gamma;
'
'% Bounding box
'nstds = 3;
'xmax = max(abs(nstds*sigma_x*cos(theta)),abs(nstds*sigma_y*sin(theta)));
'xmax = ceil(max(1,xmax));
'ymax = max(abs(nstds*sigma_x*sin(theta)),abs(nstds*sigma_y*cos(theta)));
'ymax = ceil(max(1,ymax));
'xmin = -xmax; ymin = -ymax;
'[x,y] = meshgrid(xmin:xmax,ymin:ymax);
'
'% Rotation
'x_theta=x*cos(theta)+y*sin(theta);
'y_theta=-x*sin(theta)+y*cos(theta);
'
'gb=exp(-.5*(x_theta.^2/sigma_x^2+y_theta.^2/sigma_y^2)).*cos(2*pi/lambda*x_theta+psi);
Public Sub InitGaborFilter(Sigma, Lambda, PSI, Gamma, RR)
    Dim SigmaX     As Single
    Dim SigmaY     As Single
    Dim Xmax       As Single
    Dim Ymax       As Single
    Dim Xmin       As Single
    Dim Ymin       As Single
    Dim Xtheta     As Single
    Dim Ytheta     As Single
    Dim nstds      As Long
    Dim X          As Long
    Dim Y          As Long
    Dim GB         As Single
    Dim Theta      As Single
    Dim CC         As Single
    Dim a          As Long
    Dim Avg        As Single
    Dim min        As Single
    Dim max        As Single


    Dim s          As Single

    SigmaX = Sigma
    SigmaY = Sigma / Gamma
    'Bounding box
    EdgeRadius = RR


    a = 0


    'A go from 0 to 15
    'Form 0 to PI ( - PI/16) RAD

    For Theta = 0 To PI Step (PI / 16)
        'For theta = 0 To PI * 2 Step (2 * PI / 16)
        'For theta = 0 To PI / 2 Step (2 * PI / 64)

        For X = -EdgeRadius To EdgeRadius    ' nstds
            For Y = -EdgeRadius To EdgeRadius    'nstds
                Xtheta = X * Cos(Theta) + Y * Sin(Theta)
                Ytheta = -X * Sin(Theta) + Y * Cos(Theta)
                '
                GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + PSI)

                gaborMatrix(X, Y, a) = GB    ' * 255
            Next
        Next

        a = a + 1
    Next



    '----------------------------------------------------------------------------
    ' Now find max value
    ' and divide Gabor filter values by max value
    For a = 0 To 15
        Avg = 0
        s = 0
        For X = -EdgeRadius To EdgeRadius    'nstds
            For Y = -EdgeRadius To EdgeRadius    'nstds
                Avg = Avg + gaborMatrix(X, Y, a)
                s = s + 1
                If gaborMatrix(X, Y, a) > max Then max = gaborMatrix(X, Y, a)
            Next
        Next

        Avg = Avg / s
        For X = -EdgeRadius To EdgeRadius
            For Y = -EdgeRadius To EdgeRadius
                '                GaborMATRIX(X, Y, A) = GaborMATRIX(X, Y, A) - Avg
                gaborMatrix(X, Y, a) = gaborMatrix(X, Y, a) / max
            Next
        Next
    Next
    '----------------------------------------------------------------------

    'Draw gabor at 0 degrees
    For X = -EdgeRadius To EdgeRadius
        For Y = -EdgeRadius To EdgeRadius
            GB = gaborMatrix(X, Y, 0)
            CC = 0
            If GB < 0 Then CC = -GB: GB = 0

            frmMAIN.PIC2.Line ((X + EdgeRadius) * 4, (Y + EdgeRadius) * 4)-((X + EdgeRadius + 1) * 4, (Y + EdgeRadius + 1) * 4), RGB(GB * 255, CC * 255, 0), BF
        Next
    Next

End Sub



Public Sub SetSource(pboxImageHandle As Long)
    'Public Sub GetBits(pBoxPicHand As Long)
    Dim iRet       As Long
    'Get the bitmap header
    iRet = GetObject(pboxImageHandle, Len(hBmp), hBmp)
    '   iRet = GetObject(pBoxPicHand, Len(hBmp), hBmp)

    'Resize to hold image data
    ReDim SOURCEbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into SOURCEbyte array
    'iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, SOURCEbyte(0, 0, 0))
    iRet = GetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, SOURCEbyte(0, 0, 0))


    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1
    pB = (hBmp.bmBitsPixel \ 8) - 1



    ReDim OUTbyte(0 To pB, 0 To pW, 0 To pH)
    ReDim OUTsingle(0 To pB, 0 To pW, 0 To pH)

    ReDim EdgeAngle(0 To pW, 0 To pH)
    ReDim EdgeMAG(0 To pW, 0 To pH)
    ReDim Back(0 To pW, 0 To pH)
    ReDim BackBuff(0 To pW, 0 To pH)

    ReDim HUE(0 To pW, 0 To pH)
    ReDim HUE2(0 To pW, 0 To pH)
End Sub
Public Sub SetSource2(pboxImageHandle As Long)
    'For PIC2
    'MUST be called Before SetSource(1)
    'Public Sub GetBits(pBoxPicHand As Long)
    Dim iRet       As Long
    'Get the bitmap header
    iRet = GetObject(pboxImageHandle, Len(hBmp), hBmp)
    '   iRet = GetObject(pBoxPicHand, Len(hBmp), hBmp)

    'Resize to hold image data
    ReDim SOURCEbyte2(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into SOURCEbyte array
    'iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, SOURCEbyte(0, 0, 0))
    iRet = GetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, SOURCEbyte2(0, 0, 0))
End Sub
Public Sub GetEffect(pboxImageHandle As Long)
    Dim iRet       As Long

    iRet = SetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, OUTbyte(0, 0, 0))

End Sub




Public Sub FX_PercDONE(FT As String, Value As Single, CI As Long)
    frmMAIN.ShapeProg.Width = frmMAIN.ShapeBG.Width * Value
    frmMAIN.LabProg = FT & " " & Int(Value * 100) & "%  Iter. " & CI & ""
    DoEvents
End Sub



Public Sub Apply3(StrokeLen, Limit, BackDarkness, BKGmode As Long)
    Dim X          As Long
    Dim Y          As Long
    Dim Xp         As Long
    Dim Yp         As Long
    Dim Rx         As Single
    Dim Gx         As Single
    Dim Bx         As Single
    Dim Ry         As Single
    Dim Gy         As Single
    Dim By         As Single

    Dim XX         As Single
    Dim YY         As Single


    Dim Area       As Single
    Dim InvArea    As Single

    Dim a          As Long

    Dim RR         As Single
    Dim GG         As Single
    Dim BB         As Single

    Dim XpXP       As Long
    Dim YpYP       As Long

    Dim V          As Single


    Dim dR         As Single
    Dim dG         As Single
    Dim dB         As Single

    Dim BRIGHT     As Single

    Dim StrokePressure As Single

    Dim HH         As Single
    Dim SS         As Single
    Dim PP         As Single

    Dim LLL        As Single
    Dim AAA        As Single
    Dim BBB        As Single


    Dim ProgX      As Long
    Dim ProgXStep  As Long
    Dim KI         As Single

    Dim P          As Long
    Dim ToP        As Long

    Dim MinHUE     As Single
    Dim MaxHUE     As Single
    Dim MaxHUEdiff As Single

    Dim PrevNot0Hue As Single

    Dim i          As Long

    Const DEGtoRAD As Single = PI / 180
    Const RADtoDEG As Single = 180 / PI

    Const kR       As Single = 0.299    '0.3
    Const kG       As Single = 0.587    '.59
    Const kB       As Single = 0.114    '.11


    ProgXStep = pW / 50


    If BackDarkness <> 0 Then

        GoTo SkipHUE

        MinHUE = 3600
        MaxHUE = -360

        For X = 0 To pW \ 2
            Xp = X * 2
            For Y = 0 To pH \ 2
                Yp = Y * 2

                ' OLD WAY   ---------------------------------------------
                RgbToHLS SOURCEbyte2(2, X, Y), SOURCEbyte2(1, X, Y), SOURCEbyte2(0, X, Y), _
                         HH, SS, PP

                HUE(Xp, Yp) = 45 + HH * 3.618
                HUE(Xp, Yp) = HUE(Xp, Yp) * 2 Mod 360

                'If HUE(Xp, Yp) > 10 Then
                '    PrevNot0Hue = HUE(Xp, Yp)
                'Else
                '    HUE(Xp, Yp) = PrevNot0Hue
                'End If




                'HUE(Xp, Yp) = HUE(Xp, Yp) + Cos(Xp * 0.05) * 20 + Sin(Yp * 0.05) * 20

                HUE(Xp + 1, Yp) = HUE(Xp, Yp)
                HUE(Xp, Yp + 1) = HUE(Xp, Yp)
                HUE(Xp + 1, Yp + 1) = HUE(Xp, Yp)

                If HUE(Xp, Yp) > MaxHUE Then MaxHUE = HUE(Xp, Yp)
                If HUE(Xp, Yp) < MinHUE Then MinHUE = HUE(Xp, Yp)

            Next
        Next
SkipHUE:
        '        MaxHUEdiff = MaxHUE - MinHUE
        '        For X = 0 To pW
        '        For Y = 0 To pH
        '        HUE(X, Y) = 180 * (HUE(X, Y) - MinHUE) / (MaxHUEdiff)
        '        Next
        '        Next
        GoTo skipBlurHUE

        For i = 1 To 20
            For X = 2 To pW - 2 Step 2
                For Y = 2 To pH - 2 Step 2
                    XX = Cos(DEGtoRAD * HUE(X, Y))
                    YY = Sin(DEGtoRAD * HUE(X, Y))
                    XX = XX + Cos(DEGtoRAD * HUE(X - 2, Y))
                    YY = YY + Sin(DEGtoRAD * HUE(X - 2, Y))
                    XX = XX + Cos(DEGtoRAD * HUE(X + 2, Y))
                    YY = YY + Sin(DEGtoRAD * HUE(X + 2, Y))
                    XX = XX + Cos(DEGtoRAD * HUE(X, Y - 2))
                    YY = YY + Sin(DEGtoRAD * HUE(X, Y - 2))
                    XX = XX + Cos(DEGtoRAD * HUE(X, Y + 2))
                    YY = YY + Sin(DEGtoRAD * HUE(X, Y + 2))
                    XX = XX + Cos(DEGtoRAD * HUE(X - 2, Y - 2))
                    YY = YY + Sin(DEGtoRAD * HUE(X - 2, Y - 2))
                    XX = XX + Cos(DEGtoRAD * HUE(X - 2, Y + 2))
                    YY = YY + Sin(DEGtoRAD * HUE(X - 2, Y + 2))
                    XX = XX + Cos(DEGtoRAD * HUE(X + 2, Y - 2))
                    YY = YY + Sin(DEGtoRAD * HUE(X + 2, Y - 2))
                    XX = XX + Cos(DEGtoRAD * HUE(X + 2, Y + 2))
                    YY = YY + Sin(DEGtoRAD * HUE(X + 2, Y + 2))

                    HUE2(X, Y) = RADtoDEG * Atan2(XX, YY)
                Next
            Next
            FX_PercDONE "Blurring HUE ", i / 20, 0
            HUE = HUE2
        Next
skipBlurHUE:
        '-----------------------------------------------------------------
    End If

    'COMPUTE EDGES
    Area = (2 * EdgeRadius + 1) ^ 2
    InvArea = 1 / Area

    MaxMag = 0
    For X = StrokeR To pW - StrokeR
        For Y = StrokeR To pH - StrokeR
            Rx = 0: Gx = 0: Bx = 0
            Ry = 0: Gy = 0: By = 0

            For Xp = -EdgeRadius To EdgeRadius
                XpXP = X + Xp
                For Yp = -EdgeRadius To EdgeRadius
                    YpYP = Y + Yp
                    Rx = Rx + EdgeFilterMatrixX(Xp, Yp) * (SOURCEbyte(2, XpXP, YpYP) \ 1)
                    Gx = Gx + EdgeFilterMatrixX(Xp, Yp) * (SOURCEbyte(1, XpXP, YpYP) \ 1)
                    Bx = Bx + EdgeFilterMatrixX(Xp, Yp) * (SOURCEbyte(0, XpXP, YpYP) \ 1)
                    Ry = Ry + EdgeFilterMatrixY(Xp, Yp) * (SOURCEbyte(2, XpXP, YpYP) \ 1)
                    Gy = Gy + EdgeFilterMatrixY(Xp, Yp) * (SOURCEbyte(1, XpXP, YpYP) \ 1)
                    By = By + EdgeFilterMatrixY(Xp, Yp) * (SOURCEbyte(0, XpXP, YpYP) \ 1)

                Next
            Next

            Rx = 0.333333 * (Rx * kR + Gx * kG + Bx * kB) * InvArea
            Ry = 0.333333 * (Ry * kR + Gy * kG + By * kB) * InvArea

            EdgeMAG(X, Y) = Sqr(Rx * Rx + Ry * Ry)
            'EdgeMAG(X, Y) = (Rx * Rx + Ry * Ry)

            If EdgeMAG(X, Y) > MaxMag Then MaxMag = EdgeMAG(X, Y)

            EdgeAngle(X, Y) = Atan2(Rx, Ry)

        Next
        If X Mod 20 = 0 Then FX_PercDONE "Detecting Edges ", X / pW, 0
    Next

    If MaxMag = 0 Then MaxMag = 1
    For X = 0 To pW
        For Y = 0 To pH
            EdgeMAG(X, Y) = EdgeMAG(X, Y) / MaxMag
        Next
    Next
    FX_PercDONE "Detecting Edges ", 1, 0


    '---------------------------------------------------------
    'For X = 0 To pW
    '    For Y = 0 To pH
    '        HUE(X, Y) = HUE(X, Y) * 0.5 + 180 * (EdgeAngle(X, Y) / PI)
    '        HUE(X, Y) = HUE(X, Y) Mod 360
    '    Next
    'Next
    '--------------------------------------------------------


    GoTo skipBACK

    If frmMAIN.sBackDarkness = 0 Then GoTo skipBACK

    'COMPUTE BACKGROUND
    Area = (2 * BACKRadius + 1) ^ 2
    InvArea = 1 / Area

    MaxMag = 0
    For X = StrokeR + BACKRadius To pW - StrokeR - BACKRadius Step 2
        For Y = StrokeR + BACKRadius To pH - StrokeR - BACKRadius Step 2
            Rx = 0: Gx = 0: Bx = 0
            Ry = 0: Gy = 0: By = 0

            For Xp = -BACKRadius To BACKRadius
                XpXP = X \ 2 + Xp
                For Yp = -BACKRadius To BACKRadius
                    YpYP = Y \ 2 + Yp

                    Rx = Rx + 0.1 * BackFilterMatrixX(Xp, Yp) * (SOURCEbyte2(2, XpXP, YpYP) \ 1)
                    Gx = Gx + 0.1 * BackFilterMatrixX(Xp, Yp) * (SOURCEbyte2(1, XpXP, YpYP) \ 1)
                    Bx = Bx + 0.1 * BackFilterMatrixX(Xp, Yp) * (SOURCEbyte2(0, XpXP, YpYP) \ 1)
                    Ry = Ry + 0.1 * BackFilterMatrixY(Xp, Yp) * (SOURCEbyte2(2, XpXP, YpYP) \ 1)
                    Gy = Gy + 0.1 * BackFilterMatrixY(Xp, Yp) * (SOURCEbyte2(1, XpXP, YpYP) \ 1)
                    By = By + 0.1 * BackFilterMatrixY(Xp, Yp) * (SOURCEbyte2(0, XpXP, YpYP) \ 1)

                Next
            Next

            'Rx = 0.333333 * (Rx + Gx + Bx) * InvArea '/ Area
            'Ry = 0.333333 * (Ry + Gy + By) * InvArea  '/ Area
            'Back(X, Y).M = Sqr(Rx * Rx + Ry * Ry)
            'If Back(X, Y).M > MaxMag Then MaxMag = Back(X, Y).M
            'Back(X, Y).a = Atan2(Rx, Ry)
            'Back(X, Y).X = Rx
            'Back(X, Y).Y = Ry

            Rx = (Rx * kR + Gx * kG + Bx * kB)
            Ry = (Ry * kR + Gy * kG + By * kB)


            Back(X, Y).a = Atan2(Rx, Ry)
            Back(X, Y).M = 0.1 + Sqr(Rx * Rx + Ry * Ry) ^ SmoothPow
            Back(X, Y).X = Cos(Back(X, Y).a) * (Back(X, Y).M)
            Back(X, Y).Y = Sin(Back(X, Y).a) * (Back(X, Y).M)
            If Back(X, Y).M > MaxMag Then MaxMag = Back(X, Y).M

            Back(X + 1, Y) = Back(X, Y)
            Back(X + 1, Y + 1) = Back(X, Y)
            Back(X, Y + 1) = Back(X, Y)

        Next
        If X Mod 20 = 0 Then FX_PercDONE "Detecting Background", X / pW, 0
        DoEvents
        'If X > ProgX Then
        '    FX_PercDONE "Detecting Background", X / pW, 0
        '    ProgX = ProgX + ProgXStep
        'End If
    Next


    If MaxMag = 0 Then MaxMag = 1
    For X = 0 To pW
        For Y = 0 To pH
            Back(X, Y).M = Back(X, Y).M / MaxMag
            Back(X, Y).X = Back(X, Y).X / MaxMag
            Back(X, Y).Y = Back(X, Y).Y / MaxMag
        Next
    Next

    SmoothBackGround 2            '3 'EdgeRadius

    For X = 0 To pW
        For Y = 0 To pH
            'HUE(X, Y) = 180 * Atan2(Back(X, Y).X, Back(X, Y).Y) / PI
            HUE(X, Y) = RADtoDEG * Back(X, Y).a
            If HUE(X, Y) > PI Then HUE(X, Y) = HUE(X, Y) - PI
        Next
    Next



    '********************************************************************************
    '********************************************************************************
    '********************************************************************************
    '********************************************************************************
    '********************************************************************************

skipBACK:

    ProgX = 0

    '---------DARW STROKE
    FX_PercDONE "strokes", 0, 0
    For X = StrokeR To pW - StrokeR    'Step 2
        For Y = StrokeR To pH - StrokeR    'Step 2
            '********************************

            If EdgeMAG(X, Y) > Limit Then

                Rx = Not (SOURCEbyte(2, X, Y))
                Gx = Not (SOURCEbyte(1, X, Y))
                Bx = Not (SOURCEbyte(0, X, Y))

                a = RADtoDEG * EdgeAngle(X, Y)

                'a = (a \ 10) * 10
                'StrokePressure = ((EdgeMAG(X, Y)) ^ 0.25) / 21
                StrokePressure = ((EdgeMAG(X, Y)) ^ 0.3) / 21    '.4
                StrokeLen = 2 * StrokePressure * 21 * 21
            Else
                StrokePressure = 0
                If frmMAIN.sBackDarkness <> 0 Then

                    If X Mod 2 = 0 Then
                        If Y Mod 2 = 0 Then

                            'Rx = Not (SOURCEbyte2(2, X \ 2, Y \ 2))
                            'Gx = Not (SOURCEbyte2(1, X \ 2, Y \ 2))
                            'Bx = Not (SOURCEbyte2(0, X \ 2, Y \ 2))

                            Rx = Not (SOURCEbyte4(2, X \ 4, Y \ 4))
                            Gx = Not (SOURCEbyte4(1, X \ 4, Y \ 4))
                            Bx = Not (SOURCEbyte4(0, X \ 4, Y \ 4))

                            'a = HUE(X, Y)
                            Select Case BKGmode
                                Case 0

                                    a = RADtoDEG * FindAbyEDGEV3(X, Y, Limit * 0.8) '0.4  in V3
                                Case 1

                                    a = RADtoDEG * FindAbyEDGEV4(X, Y, Limit * 0.8)
                            End Select

                            'a = 5 + (a \ 30) * 30
                            'StrokePressure = 10 * BackDarkness * ((EdgeMAG(X, Y)) ^ 0.25) / 21

                            StrokePressure = 0.5 * BackDarkness
                            StrokeLen = 7 + 7 * (EdgeMAG(X, Y) ^ 0.5)

                        End If
                    End If
                End If
            End If
            If StrokePressure <> 0 Then

                If StrokeLen > 21 Then StrokeLen = 21

                With Stroke3(a, StrokeLen)
                    ToP = .NofPoints
                    For P = 1 To ToP
                        Xp = .StrokePoint(P).Dx
                        Yp = .StrokePoint(P).Dy
                        XpXP = X + Xp
                        YpYP = Y + Yp
                        KI = .StrokePoint(P).Intens * StrokePressure

                        dR = Rx * KI
                        dG = Gx * KI
                        dB = Bx * KI

                        OUTsingle(2, XpXP, YpYP) = OUTsingle(2, XpXP, YpYP) + dR
                        OUTsingle(1, XpXP, YpYP) = OUTsingle(1, XpXP, YpYP) + dG
                        OUTsingle(0, XpXP, YpYP) = OUTsingle(0, XpXP, YpYP) + dB

                    Next
                End With
            End If
        Next
        If X Mod 20 = 0 Then FX_PercDONE "strokes", X / pW, 0
    Next

    FX_PercDONE "strokes", 1, 0

    For X = 0 To pW
        For Y = 0 To pH

            If OUTsingle(2, X, Y) < 0 Then OUTsingle(2, X, Y) = 0
            If OUTsingle(1, X, Y) < 0 Then OUTsingle(1, X, Y) = 0
            If OUTsingle(0, X, Y) < 0 Then OUTsingle(0, X, Y) = 0
            If OUTsingle(2, X, Y) > 255 Then OUTsingle(2, X, Y) = 255
            If OUTsingle(1, X, Y) > 255 Then OUTsingle(1, X, Y) = 255
            If OUTsingle(0, X, Y) > 255 Then OUTsingle(0, X, Y) = 255

            OUTbyte(2, X, Y) = 255 - OUTsingle(2, X, Y)
            OUTbyte(1, X, Y) = 255 - OUTsingle(1, X, Y)
            OUTbyte(0, X, Y) = 255 - OUTsingle(0, X, Y)

        Next
    Next



End Sub

Public Sub ReducePicBy2()

    Dim X          As Long
    Dim Y          As Long
    Dim Xp         As Long
    Dim Yp         As Long
    Dim Xto        As Long
    Dim Yto        As Long

    Dim Xfrom      As Long
    Dim Yfrom      As Long

    Dim pW2        As Long
    Dim pH2        As Long

    Dim R          As Single
    Dim G          As Single
    Dim B          As Single

    Dim F          As Long        'single

    Dim Sum        As Single

    Dim KKK(-1 To 2, -1 To 2) As Long    'single
    'http://www.velocityreviews.com/forums/t426518-magic-kernel-for-image-zoom-resampling.html

    ReDim SOURCEbyte2(0 To 2, 0 To pW \ 2, 0 To pH \ 2)

    KKK(-1, -1) = 1
    KKK(0, -1) = 3
    KKK(1, -1) = 3
    KKK(2, -1) = 1

    KKK(-1, 0) = 3
    KKK(0, 0) = 9                 '13
    KKK(1, 0) = 9                 '13
    KKK(2, 0) = 3

    KKK(-1, 1) = 3
    KKK(0, 1) = 9                 '13
    KKK(1, 1) = 9                 '13
    KKK(2, 1) = 3

    KKK(-1, 2) = 1
    KKK(0, 2) = 3
    KKK(1, 2) = 3
    KKK(2, 2) = 1

    pW2 = pW - 2
    pH2 = pH - 2
    For X = 1 To pW2 Step 2
        Xto = X \ 2 + 1
        For Y = 1 To pH2 Step 2
            Yto = Y \ 2 + 1
            R = 0
            G = 0
            B = 0
            Sum = 0
            For Xp = -1 To 2
                Xfrom = X + Xp
                For Yp = -1 To 2
                    Yfrom = Y + Yp
                    F = KKK(Xp, Yp)
                    R = R + SOURCEbyte(2, Xfrom, Yfrom) * F
                    G = G + SOURCEbyte(1, Xfrom, Yfrom) * F
                    B = B + SOURCEbyte(0, Xfrom, Yfrom) * F
                    Sum = Sum + F
                Next
            Next

            Sum = 1 / Sum
            R = R * Sum           '* 0.015625      ' 1 / 64
            G = G * Sum           '* 0.015625      ' 1 / 64
            B = B * Sum           '* 0.015625      ' 1 / 64
            SOURCEbyte2(2, Xto, Yto) = R
            SOURCEbyte2(1, Xto, Yto) = G
            SOURCEbyte2(0, Xto, Yto) = B
        Next
        If X Mod 21 = 0 Then
            FX_PercDONE "Piramid 1 ", X / pW, 0
        End If
    Next

End Sub

Public Sub ReducePicBy4()

    Dim X          As Long
    Dim Y          As Long
    Dim Xp         As Long
    Dim Yp         As Long
    Dim Xto        As Long
    Dim Yto        As Long

    Dim Xfrom      As Long
    Dim Yfrom      As Long

    Dim pW2        As Long
    Dim pH2        As Long

    Dim R          As Single
    Dim G          As Single
    Dim B          As Single

    Dim F          As Long        'single

    Dim Sum        As Single

    Dim KKK(-1 To 2, -1 To 2) As Long    'single
    'http://www.velocityreviews.com/forums/t426518-magic-kernel-for-image-zoom-resampling.html

    ReDim SOURCEbyte4(0 To 2, 0 To pW \ 4, 0 To pH \ 4)

    KKK(-1, -1) = 1
    KKK(0, -1) = 3
    KKK(1, -1) = 3
    KKK(2, -1) = 1

    KKK(-1, 0) = 3
    KKK(0, 0) = 9                 '13
    KKK(1, 0) = 9                 '13
    KKK(2, 0) = 3

    KKK(-1, 1) = 3
    KKK(0, 1) = 9                 '13
    KKK(1, 1) = 9                 '13
    KKK(2, 1) = 3

    KKK(-1, 2) = 1
    KKK(0, 2) = 3
    KKK(1, 2) = 3
    KKK(2, 2) = 1

    pW2 = pW \ 2 - 2
    pH2 = pH \ 2 - 2
    For X = 1 To pW2 Step 2
        Xto = X \ 2 + 1
        For Y = 1 To pH2 Step 2
            Yto = Y \ 2 + 1
            R = 0
            G = 0
            B = 0
            Sum = 0
            For Xp = -1 To 2
                Xfrom = X + Xp
                For Yp = -1 To 2
                    Yfrom = Y + Yp
                    F = KKK(Xp, Yp)
                    R = R + SOURCEbyte2(2, Xfrom, Yfrom) * F
                    G = G + SOURCEbyte2(1, Xfrom, Yfrom) * F
                    B = B + SOURCEbyte2(0, Xfrom, Yfrom) * F
                    Sum = Sum + F
                Next
            Next

            Sum = 1 / Sum
            R = R * Sum           '* 0.015625      ' 1 / 64
            G = G * Sum           '* 0.015625      ' 1 / 64
            B = B * Sum           '* 0.015625      ' 1 / 64
            SOURCEbyte4(2, Xto, Yto) = R
            SOURCEbyte4(1, Xto, Yto) = G
            SOURCEbyte4(0, Xto, Yto) = B
        Next
        If X Mod 21 = 0 Then
            FX_PercDONE "Piramid 2 ", X / pW2, 0
        End If
    Next

End Sub



Private Sub SmoothBackGround(N As Long)



    Dim X          As Long
    Dim Y          As Long
    Dim Xp         As Long
    Dim Yp         As Long

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Dim XmN        As Long
    Dim XpN        As Long
    Dim YmN        As Long
    Dim YpN        As Long
    Dim XPmX       As Long


    Dim DirW       As Single
    Dim sign       As Single

    Dim TmpM       As Single


    Dim Wmax       As Single

    Dim Area       As Long
    Dim InvArea    As Single

    Dim IT         As Long
    Dim MaxMag     As Single

    Dim Vx         As Single
    Dim Vy         As Single
    Dim Vm         As Single


    Dim Segno      As Single


    Xfrom = 0 + N
    Xto = pW - N
    Yfrom = 0 + N
    Yto = pH - N

    Area = (2 * N + 1) ^ 2
    InvArea = 1 / Area

    For IT = 1 To 2               '3               '3

        MaxMag = 0
        For X = Xfrom To Xto

            XmN = X - N
            XpN = X + N
            For Y = Yfrom To Yto

                YmN = Y - N
                YpN = Y + N

                Vm = 0
                Vx = 0
                Vy = 0            '0.0001

                For Xp = XmN To XpN
                    XPmX = Xp - X
                    For Yp = YmN To YpN
                        If Abs(AngleDIFF(Atan2(Vx, Vy), Back(Xp, Yp).a)) >= PIh Then
                            ''If Abs(AngleDIFF(EdgeFlow(x, Y).a, EdgeFlow(XP, YP).a)) > PIH Then
                            Segno = -1
                        Else
                            Segno = 1
                        End If

                        Vx = Vx + Back(Xp, Yp).X * Segno
                        Vy = Vy + Back(Xp, Yp).Y * Segno
                    Next
                Next
                '                Stop

                With BackBuff(X, Y)
                    '.X = Vx * InvArea
                    '.Y = Vy * InvArea
                    '.a = Atan2(Vx, Vy)
                    '' .M = Sqr(.x * .x + .Y * .Y)
                    ''If .M > 0 Then .x = .x / .M:                    .Y = .Y / .M


                    Vx = Vx * InvArea
                    Vy = Vy * InvArea
                    'If Abs(Vx) > 1 Then Stop

                    .a = Atan2(Vx, Vy)
                    .M = 0.1 + Sqr(Vx * Vx + Vy * Vy) ^ SmoothPow
                    .X = Cos(.a) * .M
                    .Y = Sin(.a) * .M
                    If .M > MaxMag Then MaxMag = .M
                End With

            Next
        Next
        Back = BackBuff

        'Normalize
        For X = 0 To pW
            For Y = 0 To pH
                With Back(X, Y)
                    .M = .M / MaxMag
                    'If .M > 0 Then
                    .X = .X / MaxMag
                    .Y = .Y / MaxMag    '??????? don't know way, but with "-" it works!
                    '.a = Atan2(.X, .Y)
                    'End If
                End With
            Next
        Next

        FX_PercDONE "Smoothing Back Ground", IT / 2, 0

    Next







    'Debug
    For X = 0 To pW Step 10       '6
        For Y = 0 To pH Step 10   '6
            With Back(X, Y)
                Xfrom = X + Cos(.a + PIh) * (1 + 3)    ' * .M)
                Yfrom = Y + Sin(.a + PIh) * (1 + 3)    ' * .M)
                Xto = X - Cos(.a + PIh) * (1 + 3)    ' * .M)
                Yto = Y - Sin(.a + PIh) * (1 + 3)    ' * .M)
            End With

            frmMAIN.PIC1.Line (Xfrom, Yfrom)-(Xto, Yto), vbYellow
            'frmMAIN.PIC1.Line (Xfrom, Yfrom)-(x, Y), vbYellow
            'frmMAIN.PIC1.Line (x, Y)-(Xto, Yto), vbRed

        Next
        frmMAIN.PIC1.Refresh

    Next
    frmMAIN.PIC1.Refresh
    '  MsgBox "Smoothed..."


End Sub


Private Function FindAbyEDGEV3(X As Long, Y As Long, LIM As Single)
    Dim a          As Single
    Dim XX         As Long
    Dim YY         As Long
    Dim R          As Single

    R = 2

    FindAbyEDGEV3 = -1

    Do

        XX = X + Cos(a) * R

        If XX > 0 Then
            If XX < pW Then

                YY = Y + Sin(a) * R
                
                If YY > 0 Then
                    If YY < pH Then
                        If EdgeMAG(XX, YY) > LIM Then
                            FindAbyEDGEV3 = EdgeAngle(XX, YY)    ': Exit For
                        End If
                    End If
                End If
            End If
        End If

        a = a + PI2 * 0.618
        R = R + 0.1               ' 0.3

        If R Mod 50 = 0 Then
            If R > pW * 3 Then FindAbyEDGEV3 = PIh * 0.5
        End If

    Loop While FindAbyEDGEV3 < 0
    Debug.Print R & vbTab & a


End Function
Private Function FindAbyEDGEV4(X As Long, Y As Long, LIM As Single)
    Dim a          As Single
    Dim XX         As Long
    Dim YY         As Long
    Dim R          As Single
    Dim sStep      As Single
    Dim MaxV       As Single

    R = 2

    FindAbyEDGEV4 = -1
    MaxV = -1
    Do
        'sStep = PI2 * 0.618 / R
        sStep = PI / R
        For a = 0 To PI2 Step sStep
            XX = X + Cos(a) * R

            If XX > 0 Then
                If XX < pW Then

                    YY = Y + Sin(a) * R

                    If YY > 0 Then
                        If YY < pH Then
                            If EdgeMAG(XX, YY) > LIM Then
                                If EdgeMAG(XX, YY) > MaxV Then
                                    FindAbyEDGEV4 = EdgeAngle(XX, YY)

                                    MaxV = FindAbyEDGEV4
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Next

        R = R + 1
        If R Mod 50 = 0 Then
            If R > pW * 3 Then FindAbyEDGEV4 = PIh * 0.5
        End If

    Loop While FindAbyEDGEV4 < 0
    Debug.Print R & vbTab & a

    'If FindAbyEDGE2 > PI2 Then FindAbyEDGE2 = FindAbyEDGE2 - PI2
    'If FindAbyEDGE2 < 0 Then FindAbyEDGE2 = FindAbyEDGE2 = PI2
End Function

