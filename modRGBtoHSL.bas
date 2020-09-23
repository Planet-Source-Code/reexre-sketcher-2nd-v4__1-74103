Attribute VB_Name = "modRGBtoHSL"
Public Sub RgbToHLS(ByVal R As Single, ByVal G As Single, _
                    ByVal B As Single, ByRef H As Single, ByRef L As _
                                                          Single, ByRef s As Single)
    Dim max        As Single
    Dim min        As Single
    Dim diff       As Single
    Dim InvDiff    As Single
    Dim r_dist     As Single
    Dim g_dist     As Single
    Dim b_dist     As Single

    ' Get the maximum and minimum RGB components.
    If R > G Then
        max = R: min = G
    Else
        max = G: min = R
    End If
    If B > max Then
        max = B
    ElseIf B < min Then
        min = B
    End If


    diff = max - min

    L = (max + min) / 2
    If Abs(diff) < 0.00001 Then
        s = 0
        H = 0                     ' H is really undefined.
    Else

        If L <= 0.5 Then
            s = diff / (max + min)
        Else
            s = diff / (2 - max - min + 0.001)
        End If
        InvDiff = 1 / diff

        r_dist = (max - R) * InvDiff
        g_dist = (max - G) * InvDiff
        b_dist = (max - B) * InvDiff

        If R = max Then
            H = b_dist - g_dist
        ElseIf G = max Then
            H = 2 + r_dist - b_dist
        Else
            H = 4 + g_dist - r_dist
        End If

        H = H * 60
        If H < 0 Then H = H + 360
    End If
End Sub

