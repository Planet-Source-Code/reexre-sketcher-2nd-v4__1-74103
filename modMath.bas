Attribute VB_Name = "modMath"
Option Explicit

Public Const PI    As Single = 3.14159265358979
Public Const PIh   As Single = 3.14159265358979 * 0.5
Public Const PI2   As Single = 3.14159265358979 * 2

Public Function Atan2(X As Single, Y As Single) As Single
    If X Then
        Atan2 = -PI + Atn(Y / X) - (X > 0) * PI
    Else
        Atan2 = -PIh - (Y > 0) * PI
    End If
    If Atan2 < 0 Then Atan2 = Atan2 + PI2

End Function

Public Function Atan2Fast2(ByRef X As Single, ByRef Y As Single) As Single
    'http://lists.apple.com/archives/perfoptimization-dev/2005/Jan/msg00051.html

    Dim Z          As Single

    If X = 0 Then
        If (Y > 0) Then Atan2Fast2 = PIh: Exit Function
        If (Y = 0) Then Atan2Fast2 = 0: Exit Function
        Atan2Fast2 = -PIh: Exit Function
    End If

    Z = Y / X
    If (Abs(Z) < 1) Then
        Atan2Fast2 = Z / (1 + 0.28 * Z * Z)
        If (X < 0) Then
            If (Y < 0) Then Atan2Fast2 = Atan2Fast2 + PI: Exit Function
            Atan2Fast2 = Atan2Fast2 + PI: Exit Function
        End If
    Else
        Atan2Fast2 = PIh - Z / (Z * Z + 0.28)
        If (Y < 0) Then Atan2Fast2 = Atan2Fast2 + PI: Exit Function
    End If

    If Atan2Fast2 < 0 Then Atan2Fast2 = Atan2Fast2 + PI2

End Function

Public Function AngleDIFF(A1 As Single, A2 As Single) As Single
    'double difference = secondAngle - firstAngle;
    'while (difference < -180) difference += 360;
    'while (difference > 180) difference -= 360;
    'return difference;

    AngleDIFF = A2 - A1
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend

    '''' this is to have values between 0 and 1
    'AngleDIFF = AngleDIFF + PI
    'AngleDIFF = AngleDIFF / (PI * 2)


End Function
