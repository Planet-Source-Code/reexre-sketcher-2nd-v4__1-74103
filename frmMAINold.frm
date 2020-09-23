VERSION 5.00
Begin VB.Form frmMAINold 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar sLIMIT 
      Height          =   255
      Left            =   12360
      Max             =   950
      TabIndex        =   17
      Top             =   6480
      Value           =   125
      Width           =   1695
   End
   Begin VB.HScrollBar sANG 
      Height          =   255
      Left            =   12960
      Max             =   360
      TabIndex        =   16
      Top             =   3000
      Value           =   180
      Width           =   1695
   End
   Begin VB.CommandButton CommandTestStroke 
      Caption         =   "Stroke"
      Height          =   615
      Left            =   12240
      TabIndex        =   15
      Top             =   3720
      Width           =   855
   End
   Begin VB.HScrollBar sRadius 
      Height          =   255
      Left            =   12960
      Max             =   48
      Min             =   2
      TabIndex        =   13
      Top             =   480
      Value           =   4
      Width           =   1695
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   1215
      Left            =   13200
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.HScrollBar sGAMMA 
      Height          =   255
      Left            =   12960
      Max             =   500
      Min             =   1
      TabIndex        =   6
      Top             =   2280
      Value           =   100
      Width           =   1695
   End
   Begin VB.HScrollBar sPSI 
      Height          =   255
      Left            =   12960
      Max             =   200
      Min             =   1
      TabIndex        =   5
      Top             =   1800
      Value           =   1
      Width           =   1695
   End
   Begin VB.HScrollBar sLAMBDA 
      Height          =   255
      Left            =   12960
      Max             =   4000
      Min             =   1
      TabIndex        =   4
      Top             =   1320
      Value           =   2000
      Width           =   1695
   End
   Begin VB.HScrollBar sSIGMA 
      Height          =   255
      Left            =   12960
      Max             =   3000
      TabIndex        =   3
      Top             =   840
      Value           =   500
      Width           =   1695
   End
   Begin VB.CommandButton Command 
      Caption         =   "Init Gabor Filter"
      Height          =   615
      Left            =   13200
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   287
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   1
      Top             =   5640
      Width           =   6495
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   120
      Picture         =   "frmMAINold.frx":0000
      ScaleHeight     =   343
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   607
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label lRadius 
      Height          =   255
      Left            =   12240
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lSigma 
      Height          =   255
      Left            =   11640
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lLambda 
      Height          =   255
      Left            =   11640
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lgamma 
      Height          =   255
      Left            =   11640
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "gabor filter parameters"
      Height          =   255
      Left            =   12960
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lPSI 
      Height          =   255
      Left            =   11640
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmMAINold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()


    'SetupStroke2 16
    
    InitEDGEFilter sRadius
    InitBACKFilter 16
    
    SetSource PIC1.Image.Handle
    Apply 22, 100 - scrollLIMIT, sBackDarkness / 100
    
    GetEffect PIC2.Image.Handle
    PIC2.Refresh

    '    Pic2.Refresh
    '    Pic1.PaintPicture Pic2.Image, 0, 0, Pic1.Width, Pic1.Height, 0, 0, Pic2.Width, Pic2.Height
    '    Pic1.Refresh

End Sub

Private Sub CommandTestStroke_Click()
    SetupStroke2 sRadius

End Sub

Private Sub Form_Load()
    PIC2.Cls
    PIC2.Width = PIC1.Width
    PIC2.Height = PIC1.Height
SetupStroke2 0


End Sub

Private Sub Command_Click()



    'InitEDGEFilter , sRadius

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub sANG_Change()
    CommandTestStroke_Click
End Sub

Private Sub sANG_Scroll()
    CommandTestStroke_Click
End Sub

Private Sub sGAMMA_Change()
    Command_Click
    lgamma = sGAMMA / 100
End Sub

Private Sub sGAMMA_Scroll()
    Command_Click
    lgamma = sGAMMA / 100
End Sub

Private Sub sLAMBDA_Change()
    Command_Click
    lLambda = sLAMBDA / 100
End Sub

Private Sub sLAMBDA_Scroll()
    Command_Click
    lLambda = sLAMBDA / 100
End Sub

Private Sub sPSI_Change()
    Command_Click
    lPSI = sPSI / 100
End Sub

Private Sub sPSI_Scroll()
    Command_Click
    lPSI = sPSI / 100
End Sub

Private Sub sRadius_Change()
    Command_Click
    lRadius = sRadius

End Sub

Private Sub sRadius_Scroll()
    Command_Click
    lRadius = sRadius
End Sub

Private Sub sSIGMA_Change()
    Command_Click
    lSigma = sSIGMA / 100
End Sub

Private Sub sSIGMA_Scroll()
    Command_Click
    lSigma = sSIGMA / 100
End Sub
