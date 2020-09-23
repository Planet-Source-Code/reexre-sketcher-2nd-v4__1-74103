Attribute VB_Name = "modHSP"
'©2006 Darel Rex Finley
'
'http://www.alienryderflex.com/hsp.html
'
'#define  Pr  .241
'#define  Pg  .691
'#define  Pb  .068
'
'
'
'//  public domain function by Darel Rex Finley, 2006
'//
'//  This function expects the passed-in values to be on a scale
'//  of 0 to 1, and uses that same scale for the return values.
'//
'//  See description/examples at alienryderflex.com/hsp.html
'
'void RGBtoHSP(
'double  R, double  G, double  B,
'double *H, double *S, double *P) {
'
'  //  Calculate the Perceived brightness.
'  *P=sqrt(R*R*Pr+G*G*Pg+B*B*Pb);
'
'  //  Calculate the Hue and Saturation.  (This part works
'  //  the same way as in the HSV/B and HSL systems???.)
'  if      (R==G && R==B) {
'    *H=0.; *S=0.; return; }
'  if      (R>=G && R>=B) {   //  R is largest
'    if    (B>=G) {
'      *H=6./6.-1./6.*(B-G)/(R-G); *S=1.-G/R; }
'    else         {
'      *H=0./6.+1./6.*(G-B)/(R-B); *S=1.-B/R; }}
'  else if (G>=R && G>=B) {   //  G is largest
'    if    (R>=B) {
'      *H=2./6.-1./6.*(R-B)/(G-B); *S=1.-B/G; }
'    else         {
'      *H=2./6.+1./6.*(B-R)/(G-R); *S=1.-R/G; }}
''  else                   {   //  B is largest
'    if    (G>=R) {
'      *H=4./6.-1./6.*(G-R)/(B-R); *S=1.-R/B; }
'    else         {
'      *H=4./6.+1./6.*(R-G)/(B-G); *S=1.-G/B; }}}
'
'
'
'//  public domain function by Darel Rex Finley, 2006
'//
'//  This function expects the passed-in values to be on a scale
'//  of 0 to 1, and uses that same scale for the return values.
'//
'//  Note that some combinations of HSP, even if in the scale
'//  0-1, may return RGB values that exceed a value of 1.  For
'//  example, if you pass in the HSP color 0,1,1, the result
'//  will be the RGB color 2.037,0,0.
'//
'//  See description/examples at alienryderflex.com/hsp.html'
'
'void HSPtoRGB(
'double  H, double  S, double  P,
'double *R, double *G, double *B) {
'
'  double  part, minOverMax=1.-S ;
'
'  if (minOverMax>0.) {
'    if      ( H<1./6.) {   //  R>G>B
'      H= 6.*( H-0./6.); part=1.+H*(1./minOverMax-1.);
'      *B=P/sqrt(Pr/minOverMax/minOverMax+Pg*part*part+Pb);
'      *R=(*B)/minOverMax; *G=(*B)+H*((*R)-(*B)); }
'    else if ( H<2./6.) {   //  G>R>B
'      H= 6.*(-H+2./6.); part=1.+H*(1./minOverMax-1.);
'      *B=P/sqrt(Pg/minOverMax/minOverMax+Pr*part*part+Pb);
'      *G=(*B)/minOverMax; *R=(*B)+H*((*G)-(*B)); }
'    else if ( H<3./6.) {   //  G>B>R
'      H= 6.*( H-2./6.); part=1.+H*(1./minOverMax-1.);
'      *R=P/sqrt(Pg/minOverMax/minOverMax+Pb*part*part+Pr);
'      *G=(*R)/minOverMax; *B=(*R)+H*((*G)-(*R)); }
'    else if ( H<4./6.) {   //  B>G>R
'      H= 6.*(-H+4./6.); part=1.+H*(1./minOverMax-1.);
'      *R=P/sqrt(Pb/minOverMax/minOverMax+Pg*part*part+Pr);
'      *B=(*R)/minOverMax; *G=(*R)+H*((*B)-(*R)); }
'    else if ( H<5./6.) {   //  B>R>G
'      H= 6.*( H-4./6.); part=1.+H*(1./minOverMax-1.);
'      *G=P/sqrt(Pb/minOverMax/minOverMax+Pr*part*part+Pg);
'      *B=(*G)/minOverMax; *R=(*G)+H*((*B)-(*G)); }
'    else               {   //  R>B>G
'      H= 6.*(-H+6./6.); part=1.+H*(1./minOverMax-1.);
'      *G=P/sqrt(Pr/minOverMax/minOverMax+Pb*part*part+Pg);
'      *R=(*G)/minOverMax; *B=(*G)+H*((*R)-(*G)); }}
'  else {
'    if      ( H<1./6.) {   //  R>G>B
'      H= 6.*( H-0./6.); *R=sqrt(P*P/(Pr+Pg*H*H)); *G=(*R)*H; *B=0.; }
'    else if ( H<2./6.) {   //  G>R>B
'      H= 6.*(-H+2./6.); *G=sqrt(P*P/(Pg+Pr*H*H)); *R=(*G)*H; *B=0.; }
'    else if ( H<3./6.) {   //  G>B>R
'      H= 6.*( H-2./6.); *G=sqrt(P*P/(Pg+Pb*H*H)); *B=(*G)*H; *R=0.; }
'    else if ( H<4./6.) {   //  B>G>R
'      H= 6.*(-H+4./6.); *B=sqrt(P*P/(Pb+Pg*H*H)); *G=(*B)*H; *R=0.; }
'    else if ( H<5./6.) {   //  B>R>G
'      H= 6.*( H-4./6.); *B=sqrt(P*P/(Pb+Pr*H*H)); *R=(*B)*H; *G=0.; }
'    else               {   //  R>B>G
'      H= 6.*(-H+6./6.); *R=sqrt(P*P/(Pr+Pb*H*H)); *B=(*R)*H; *G=0.; }}}''
'
'
Option Explicit

Private Const Pr   As Double = 0.299    ' 0.241
Private Const Pg   As Double = 0.587    '0.691
Private Const pB   As Double = 0.114    '0.068


Private Const D12  As Double = 0.5
Private Const D13  As Double = 0.333333333333333
Private Const D14  As Double = 0.25
Private Const D15  As Double = 0.2
Private Const D16  As Double = 0.166666666666667
Private Const D23  As Double = 0.666666666666667




Public Sub RGBtoHSP(ByVal R, ByVal G, ByVal B, ByRef rH, ByRef rS, ByRef rP)
    Dim P          As Double
    Dim H          As Double
    Dim s          As Double


    P = Sqr(R * R * Pr + G * G * Pg + B * B * pB)


    '  //  Calculate the Hue and Saturation.  (This part works
    '  //  the same way as in the HSV/B and HSL systems???.)
    If (R = G And R = B) Then H = 0: s = 0: GoTo RET

    If (R >= G And R >= B) Then
        '   //  R is largest
        If (B >= G) Then
            H = 1 - D16 * (B - G) / (R - G): s = 1 - G / R
        Else
            H = D16 * (G - B) / (R - B): s = 1 - B / R
        End If
    Else
        If (G >= R And G >= B) Then
            '//  G is largest
            If (R >= B) Then
                H = D13 - D16 * (R - B) / (G - B): s = 1 - B / G
            Else
                H = D13 + D16 * (B - R) / (G - R): s = 1 - R / G
            End If

        Else
            '//  B is largest
            If (G >= R) Then
                H = D23 - D16 * (G - R) / (B - R): s = 1 - R / B
            Else
                H = D23 + D16 * (R - G) / (B - G): s = 1 - G / B
            End If
        End If
    End If

RET:

    'Stop

    rP = P
    rH = H * 255
    rS = s * 255

End Sub


Public Sub HSPtoRGB(ByVal H, ByVal s, ByVal P, ByRef R, ByRef G, ByRef B)

    Dim Part       As Double
    Dim minOverMax As Double


    H = H / 255
    s = s / 255

    minOverMax = 1 - s
    'Stop

    If (minOverMax > 0) Then

        If (H < 1 / 6) Then       '   //  R>G>B
            H = 6 * (H - 0 / 6): Part = 1 + H * (1 / minOverMax - 1)
            B = P / Sqr(Pr / minOverMax / minOverMax + Pg * Part * Part + pB)
            R = (B) / minOverMax: G = (B) + H * (R - B)
        Else

            If (H < 2 / 6) Then   '   //  G>R>B
                H = 6 * (-H + 2 / 6): Part = 1 + H * (1 / minOverMax - 1)
                B = P / Sqr(Pg / minOverMax / minOverMax + Pr * Part * Part + pB)
                G = (B) / minOverMax: R = (B) + H * (G - B)
            Else
                If (H < 3 / 6) Then    '   //  G>B>R
                    H = 6 * (H - 2 / 6): Part = 1 + H * (1 / minOverMax - 1)
                    R = P / Sqr(Pg / minOverMax / minOverMax + pB * Part * Part + Pr)
                    G = (R) / minOverMax: B = (R) + H * (G - R)
                Else
                    If (H < 4 / 6) Then    '   //  B>G>R
                        H = 6 * (-H + 4 / 6): Part = 1 + H * (1 / minOverMax - 1)
                        R = P / Sqr(pB / minOverMax / minOverMax + Pg * Part * Part + Pr)
                        B = (R) / minOverMax: G = (R) + H * (B - R)
                    Else


                        If (H < 5 / 6) Then    '  //  B>R>G
                            H = 6 * (H - 4 / 6): Part = 1 + H * (1 / minOverMax - 1)
                            G = P / Sqr(pB / minOverMax / minOverMax + Pr * Part * Part + Pg)
                            B = (G) / minOverMax: R = (G) + H * (B - G)
                        Else      '   //  R>B>G
                            H = 6 * (-H + 6 / 6): Part = 1 + H * (1 / minOverMax - 1)
                            G = P / Sqr(Pr / minOverMax / minOverMax + pB * Part * Part + Pg)
                            R = (G) / minOverMax: B = (G) + H * (R - G)
                        End If
                    End If
                End If
            End If
        End If

    Else

        If (H < 1 / 6) Then       '   //  R>G>B
            H = 6# * (H - 0# / 6#): R = Sqr(P * P / (Pr + Pg * H * H)): G = (R) * H: B = 0
        Else
            If (H < 2 / 6) Then   '   //  G>R>B
                H = 6# * (-H + 2# / 6#): G = Sqr(P * P / (Pg + Pr * H * H)): R = (G) * H: B = 0
            Else
                If (H < 3 / 6) Then    '   //  G>B>R
                    H = 6# * (H - 2# / 6#): G = Sqr(P * P / (Pg + pB * H * H)): B = (G) * H: R = 0
                Else
                    If (H < 4 / 6) Then    '   //  B>G>R
                        H = 6# * (-H + 4# / 6#): B = Sqr(P * P / (pB + Pg * H * H)): G = (B) * H: R = 0
                    Else
                        If (H < 5 / 6) Then    '  //  B>R>G
                            H = 6# * (H - 4# / 6#): B = Sqr(P * P / (pB + Pr * H * H)): R = (B) * H: G = 0
                        Else      '   //  R>B>G
                            H = 6# * (-H + 6# / 6#): R = Sqr(P * P / (Pr + pB * H * H)): B = (R) * H: G = 0
                        End If
                    End If
                End If
            End If
        End If
    End If

    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255

End Sub


Public Function HUEDifference(H1, H2) As Single
    Dim d          As Single

    'D = Abs(H2 - H1)
    'While D < -127.5
    '    D = D + 255
    'Wend
    'While D > 127.5
    '    D = D - 255
    'wend
    'D = Abs(D)
    'If D > 128 Then D = 0

    d = H2 - (-H1)
    If d < -127.5 Then d = d + 255
    If d > 127.5 Then d = d - 255
    HUEDifference = Abs(d)


End Function
