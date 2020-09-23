Attribute VB_Name = "modStroke"
Option Explicit

Public Type tStrokePoint3
    Dx             As Long
    Dy             As Long
    Intens         As Single
End Type

Public Type tStroke3
    NofPoints      As Long
    StrokePoint()  As tStrokePoint3
End Type


Public Stroke3(0 To 360, 0 To 21) As tStroke3

Public StrokeR     As Long



Public Sub SetupStroke3(RR)
    Dim SigmaX     As Single
    Dim SigmaY     As Single
    Dim Xtheta     As Single
    Dim Ytheta     As Single
    Dim X          As Single
    Dim Y          As Single
    Dim GB         As Single
    Dim CC         As Single
    Dim max(0 To 360) As Single
    Dim Theta      As Single
    Dim Sigma      As Single
    Dim Lambda     As Single
    Dim PSI        As Single
    Dim Gamma      As Single
    Dim a          As Long
    Dim P          As Long


    StrokeR = RR
    Sigma = RR

    Lambda = 5
    PSI = 0.0001
    Gamma = PI / 2

    Dim NP         As Long

    For StrokeR = 0 To 21
        Sigma = StrokeR
        If Sigma = 0 Then Sigma = 0.000001
        SigmaX = Sigma
        SigmaY = Sigma / Gamma

        For a = 0 To 360
            max(a) = 0
            Theta = 2 * PI * a / 360
            NP = 0
            For X = -StrokeR To StrokeR
                For Y = -StrokeR To StrokeR

                    Xtheta = X * Cos(Theta) + Y * Sin(Theta)
                    Ytheta = -X * Sin(Theta) + Y * Cos(Theta)

                    If Xtheta > -2 And Xtheta < 2 Then

                        GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + PSI)
                        If GB < 0 Then GB = 0
                        If GB > max(a) Then max(a) = GB

                        If GB <> 0 Then
                            NP = NP + 1
                            ReDim Preserve Stroke3(a, StrokeR).StrokePoint(NP)
                            Stroke3(a, StrokeR).NofPoints = NP
                            With Stroke3(a, StrokeR).StrokePoint(NP)
                                .Dx = X
                                .Dy = Y
                                .Intens = GB
                            End With

                        End If

                        'StrokeMatrix(x, Y, a, StrokeR) = GB
                    Else
                        'StrokeMatrix(x, Y, a, StrokeR) = 0
                    End If
                Next
            Next

        Next

        For a = 0 To 360

            For P = 1 To Stroke3(a, StrokeR).NofPoints
                Stroke3(a, StrokeR).StrokePoint(P).Intens = _
                Stroke3(a, StrokeR).StrokePoint(P).Intens / max(a)
            Next

        Next

    Next


End Sub


