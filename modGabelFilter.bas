Attribute VB_Name = "modGabelFilter"
Option Explicit


'Public vSigma      As Single
'Public vLambda     As Single
'Public vPsi        As Single
'Public vGamma      As Single


Public EdgeFilterMatrixX(-100 To 100, -100 To 100) As Single
Public EdgeFilterMatrixY(-100 To 100, -100 To 100) As Single
Public BackFilterMatrixX(-100 To 100, -100 To 100) As Single
Public BackFilterMatrixY(-100 To 100, -100 To 100) As Single


Public EdgeRadius  As Long
Public BACKRadius  As Long

Public Sub InitEDGEFilter(RR)
    Dim Sigma      As Single
    Dim Lambda     As Single
    Dim PSI        As Single
    Dim Gamma      As Single

    Dim SigmaX     As Single
    Dim SigmaY     As Single
    Dim Xtheta     As Single
    Dim Ytheta     As Single
    Dim X          As Single
    Dim Y          As Single
    Dim GB         As Single
    Dim CC         As Single
    Dim MaxX       As Single
    Dim MaxY       As Single

    Dim Theta      As Single

    EdgeRadius = RR

    '*************************************
    Sigma = RR * 0.5
    Lambda = RR * 3
    PSI = PI / 2
    Gamma = 1
    '*************************************

    If Sigma = 0 Then Sigma = 0.000001

    SigmaX = Sigma

    SigmaY = Sigma / Gamma


    MaxX = 0
    Theta = 0
    For X = -EdgeRadius To EdgeRadius
        For Y = -EdgeRadius To EdgeRadius
            Xtheta = X * Cos(Theta) + Y * Sin(Theta)
            Ytheta = -X * Sin(Theta) + Y * Cos(Theta)
            '
            GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + PSI)
            'GB = GB ^ 3
            If GB > MaxX Then MaxX = GB
            EdgeFilterMatrixX(X, Y) = GB
        Next
    Next

    MaxY = 0
    Theta = PI / 2
    For X = -EdgeRadius To EdgeRadius
        For Y = -EdgeRadius To EdgeRadius
            Xtheta = X * Cos(Theta) + Y * Sin(Theta)
            Ytheta = -X * Sin(Theta) + Y * Cos(Theta)
            '
            GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + PSI)
            'GB = GB ^ 3
            If GB > MaxY Then MaxY = GB
            EdgeFilterMatrixY(X, Y) = GB
        Next
    Next

    For X = -EdgeRadius To EdgeRadius
        For Y = -EdgeRadius To EdgeRadius
            EdgeFilterMatrixX(X, Y) = EdgeFilterMatrixX(X, Y) / MaxX
            EdgeFilterMatrixY(X, Y) = EdgeFilterMatrixY(X, Y) / MaxY
        Next
    Next



    'Draw gabor at 0 degrees
    For X = -EdgeRadius To EdgeRadius
        For Y = -EdgeRadius To EdgeRadius
            GB = EdgeFilterMatrixX(X, Y)
            CC = 0
            If GB < 0 Then CC = -GB: GB = 0

            frmMAIN.PIC2.Line ((X + EdgeRadius) * 2, (Y + EdgeRadius) * 2)-((X + EdgeRadius + 1) * 2, (Y + EdgeRadius + 1) * 2), RGB(GB * 255, CC * 255, 0), BF
        Next
    Next


    'Changed to SOBEL:
    'Sobel Jähne et al. [1999] variation
    EdgeFilterMatrixX(-1, -1) = 0.183
    EdgeFilterMatrixX(0, -1) = 0
    EdgeFilterMatrixX(1, -1) = -0.183
    EdgeFilterMatrixX(-1, 0) = 0.634
    EdgeFilterMatrixX(0, 0) = 0
    EdgeFilterMatrixX(1, 0) = -0.634
    EdgeFilterMatrixX(-1, 1) = 0.183
    EdgeFilterMatrixX(0, 1) = 0
    EdgeFilterMatrixX(1, 1) = -0.183

    EdgeFilterMatrixY(-1, -1) = 0.183
    EdgeFilterMatrixY(0, -1) = 0.634
    EdgeFilterMatrixY(1, -1) = 0.183
    EdgeFilterMatrixY(-1, 0) = 0
    EdgeFilterMatrixY(0, 0) = 0
    EdgeFilterMatrixY(1, 0) = 0
    EdgeFilterMatrixY(-1, 1) = -0.183
    EdgeFilterMatrixY(0, 1) = -0.634
    EdgeFilterMatrixY(1, 1) = -0.183

End Sub

Public Sub InitBACKFilter(RR)
    Dim Sigma      As Single
    Dim Lambda     As Single
    Dim PSI        As Single
    Dim Gamma      As Single

    Dim SigmaX     As Single
    Dim SigmaY     As Single
    Dim Xtheta     As Single
    Dim Ytheta     As Single
    Dim X          As Single
    Dim Y          As Single
    Dim GB         As Single
    Dim CC         As Single
    Dim MaxX       As Single
    Dim MaxY       As Single

    Dim Theta      As Single


    BACKRadius = RR

    '*************************************
    Sigma = RR * 0.5
    Lambda = RR * 3
    PSI = PI / 2
    Gamma = 1
    '*************************************

    If Sigma = 0 Then Sigma = 0.000001

    SigmaX = Sigma

    SigmaY = Sigma / Gamma

    MaxX = 0
    Theta = 0
    For X = -BACKRadius To BACKRadius
        For Y = -BACKRadius To BACKRadius
            Xtheta = X * Cos(Theta) + Y * Sin(Theta)
            Ytheta = -X * Sin(Theta) + Y * Cos(Theta)
            '
            GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + PSI)
            If GB > MaxX Then MaxX = GB
            BackFilterMatrixX(X, Y) = GB
        Next
    Next

    MaxY = 0
    Theta = PI / 2
    For X = -BACKRadius To BACKRadius
        For Y = -BACKRadius To BACKRadius
            Xtheta = X * Cos(Theta) + Y * Sin(Theta)
            Ytheta = -X * Sin(Theta) + Y * Cos(Theta)
            '
            GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + PSI)
            If GB > MaxY Then MaxY = GB
            BackFilterMatrixY(X, Y) = GB
        Next
    Next

    For X = -BACKRadius To BACKRadius
        For Y = -BACKRadius To BACKRadius
            BackFilterMatrixX(X, Y) = BackFilterMatrixX(X, Y) / MaxX
            BackFilterMatrixY(X, Y) = BackFilterMatrixY(X, Y) / MaxY
        Next
    Next



    'Draw gabor at 0 degrees
    For X = -BACKRadius To BACKRadius
        For Y = -BACKRadius To BACKRadius
            GB = BackFilterMatrixX(X, Y)
            CC = 0
            If GB < 0 Then CC = -GB: GB = 0

            frmMAIN.PIC2.Line ((X + BACKRadius) * 2, (Y + BACKRadius) * 2)-((X + BACKRadius + 1) * 2, (Y + BACKRadius + 1) * 2), RGB(GB * 255, CC * 255, 0), BF
        Next
    Next

End Sub
