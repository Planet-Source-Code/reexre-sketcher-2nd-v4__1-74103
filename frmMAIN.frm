VERSION 5.00
Begin VB.Form frmMAIN 
   BackColor       =   &H00808080&
   Caption         =   "abstraction    SKETCHER   2ND"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   673
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMOVE 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12960
      MousePointer    =   5  'Size
      ScaleHeight     =   15
      ScaleMode       =   0  'User
      ScaleWidth      =   36
      TabIndex        =   4
      Top             =   240
      Width           =   570
   End
   Begin VB.PictureBox MAINframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   6960
      ScaleHeight     =   9225
      ScaleWidth      =   8265
      TabIndex        =   3
      Top             =   240
      Width           =   8295
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "about"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   45
         Top             =   8880
         Width           =   1455
      End
      Begin VB.TextBox tYCrop 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         TabIndex        =   43
         Text            =   "0"
         ToolTipText     =   "Crop Input picture Top and Bottom by this N of pixels"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbResizeMode 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   40
         ToolTipText     =   "Input Resize"
         Top             =   512
         Width           =   2535
      End
      Begin VB.TextBox tMAXWH 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         TabIndex        =   39
         Text            =   "360"
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox PicFolder 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   240
         ScaleHeight     =   5385
         ScaleWidth      =   3465
         TabIndex        =   34
         Top             =   1440
         Width           =   3495
         Begin VB.CheckBox chStart 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Starting from this picture"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   44
            ToolTipText     =   "if Checked Elaborate all Pictures in this Folder. Uncheck to stop this process."
            Top             =   5040
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.CheckBox chALL 
            BackColor       =   &H00C0C0C0&
            Caption         =   "All Pictures in this Folder"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   38
            ToolTipText     =   "if Checked Elaborate all Pictures in this Folder. Uncheck to stop this process."
            Top             =   4680
            Width           =   3255
         End
         Begin VB.DriveListBox Drive1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   3495
         End
         Begin VB.DirListBox Dir1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2250
            Left            =   0
            TabIndex        =   36
            ToolTipText     =   "Select Folder"
            Top             =   360
            Width           =   3495
         End
         Begin VB.FileListBox File1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1920
            Left            =   0
            Pattern         =   "*.jpg;*.bmp"
            TabIndex        =   35
            ToolTipText     =   "Click to Load input Picture"
            Top             =   2640
            Width           =   3495
         End
      End
      Begin VB.PictureBox PICpar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8295
         Left            =   4680
         Picture         =   "frmMAIN.frx":08CA
         ScaleHeight     =   553
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   233
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
         Begin VB.PictureBox pPARAM 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3975
            Index           =   1
            Left            =   120
            ScaleHeight     =   263
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   191
            TabIndex        =   14
            Top             =   1680
            Width           =   2895
            Begin VB.ComboBox cBKGmode 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmMAIN.frx":43BA8
               Left            =   120
               List            =   "frmMAIN.frx":43BAA
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   1560
               Width           =   1695
            End
            Begin VB.HScrollBar sBackDarkness 
               Height          =   255
               Left            =   120
               Max             =   50
               TabIndex        =   48
               Top             =   1200
               Value           =   25
               Width           =   1695
            End
            Begin VB.HScrollBar sRadius 
               Height          =   255
               Left            =   120
               Max             =   32
               Min             =   2
               TabIndex        =   46
               Top             =   2880
               Value           =   2
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.HScrollBar scrollLIMIT 
               Height          =   255
               Left            =   120
               Max             =   100
               Min             =   25
               TabIndex        =   32
               Top             =   600
               Value           =   85
               Width           =   1695
            End
            Begin VB.Label Label2 
               BackColor       =   &H00C0C0C0&
               Caption         =   "BackGround Darkness"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   960
               Width           =   2655
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Detect Radius"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   2640
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.Label lLIMIT 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Limit"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   360
               Width           =   2655
            End
            Begin VB.Label lParam 
               BackColor       =   &H00E0E0E0&
               Caption         =   "* SKETCH"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   30
               ToolTipText     =   "Hide/Show ""Sketch Settings"""
               Top             =   0
               Width           =   2895
            End
         End
         Begin VB.PictureBox pPARAM 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2175
            Index           =   0
            Left            =   120
            ScaleHeight     =   143
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   191
            TabIndex        =   13
            Top             =   120
            Width           =   2895
            Begin VB.ComboBox cmbPREeffect 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   29
               ToolTipText     =   "Pre EFFECT"
               Top             =   270
               Width           =   1695
            End
            Begin VB.PictureBox pManual 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1575
               Left            =   120
               ScaleHeight     =   105
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   105
               TabIndex        =   19
               Top             =   600
               Width           =   1575
               Begin VB.HScrollBar sBRIGHT 
                  Height          =   255
                  Left            =   0
                  Max             =   512
                  TabIndex        =   25
                  Top             =   240
                  Value           =   256
                  Width           =   1335
               End
               Begin VB.HScrollBar sCONTRA 
                  Height          =   255
                  Left            =   0
                  Max             =   100
                  Min             =   -100
                  TabIndex        =   24
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.HScrollBar sSATUR 
                  Height          =   255
                  Left            =   0
                  Max             =   512
                  TabIndex        =   23
                  Top             =   1200
                  Value           =   256
                  Width           =   1335
               End
               Begin VB.CommandButton resetB 
                  Caption         =   "B"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   22
                  ToolTipText     =   "Reset Brightness"
                  Top             =   240
                  Width           =   255
               End
               Begin VB.CommandButton restC 
                  Caption         =   "C"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   21
                  ToolTipText     =   "Reset Contrast"
                  Top             =   720
                  Width           =   255
               End
               Begin VB.CommandButton resetS 
                  Caption         =   "S"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   20
                  ToolTipText     =   "Reset Saturation"
                  Top             =   1200
                  Width           =   255
               End
               Begin VB.Label lBRIGHT 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bright."
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   28
                  Top             =   0
                  Width           =   1575
               End
               Begin VB.Label lCONTRA 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contra : 0%"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   27
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.Label lSATUR 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Satur."
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   26
                  Top             =   960
                  Width           =   1575
               End
            End
            Begin VB.PictureBox pExposure 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   120
               ScaleHeight     =   41
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   105
               TabIndex        =   15
               Top             =   600
               Width           =   1575
               Begin VB.CommandButton ResetExpo 
                  Caption         =   "E"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   17
                  ToolTipText     =   "Reset Exposure"
                  Top             =   240
                  Width           =   255
               End
               Begin VB.HScrollBar sEXPO 
                  Height          =   255
                  Left            =   0
                  Max             =   256
                  Min             =   -127
                  TabIndex        =   16
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label lEXPO 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exposure"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   18
                  Top             =   0
                  Width           =   1575
               End
            End
            Begin VB.Label lParam 
               BackColor       =   &H00E0E0E0&
               Caption         =   "* pre EFFECT"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   31
               ToolTipText     =   "Hide/Show ""Pre EFFECT"""
               Top             =   0
               Width           =   2895
            End
         End
         Begin VB.VScrollBar ScrollPar 
            Height          =   5775
            Left            =   3120
            Max             =   1
            TabIndex        =   12
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CheckBox chPrintParams 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print Parameters"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   2400
         TabIndex        =   8
         ToolTipText     =   "Print Parameters to Output Picture(s)"
         Top             =   6960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "S K E T C H"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         TabIndex        =   7
         Top             =   7680
         Width           =   2055
      End
      Begin VB.CheckBox chSelFold 
         BackColor       =   &H00C0C0C0&
         Caption         =   "File <-> Parameters"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Swap ""Folder/File Selection"" <-> ""Parameters"""
         Top             =   6960
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chMakeCompare 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Make Compare Picture"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Create a compare picture with both input and output in the same picture"
         Top             =   8040
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Source Crop Y by"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   42
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Input Resize"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   265
         Width           =   1455
      End
      Begin VB.Label MainFrameLabel 
         BackColor       =   &H0009C009&
         Caption         =   "  Panel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Click to Hide/show"
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label LabProg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   8520
         Width           =   3495
      End
      Begin VB.Shape ShapeBG 
         BorderWidth     =   2
         Height          =   255
         Left            =   240
         Top             =   8520
         Width           =   3615
      End
      Begin VB.Shape ShapeProg 
         BorderWidth     =   2
         FillColor       =   &H0070B070&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   240
         Top             =   8520
         Width           =   3375
      End
   End
   Begin VB.PictureBox PIC2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.PictureBox PIC1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3795
      Left            =   120
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   1
      Top             =   120
      Width           =   5820
      Begin VB.Shape sP 
         BorderWidth     =   2
         Height          =   1095
         Left            =   480
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox PicIN 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   960
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   2
      Top             =   -4320
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox PicLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4920
      Picture         =   "frmMAIN.frx":43BAC
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   50
      Top             =   4560
      Visible         =   0   'False
      Width           =   1530
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private MaxWH      As Integer


Private pX1        As Integer
Private pY1        As Integer
Private pX2        As Integer
Private pY2        As Integer
Private Rect       As Boolean



Private Const JpgQuality As Byte = 99    ' 95

Private PaH(0 To 4) As Long
Private ParH       As Long


Private tSIGMA     As Single
Private tRad       As Long
Private tSigmaSpatial As Single
Private ITER       As Long
Private oRGB       As Boolean
Private tCONT      As Single
Private tLUMHUE    As Single

Private ExtendSetting As String





Private Sub chALL_Click()
    chStart.Visible = chALL

End Sub





Private Sub chSelFold_Click()

    If chSelFold.Value = Checked Then
        PicFolder.Visible = True

        PICpar.Visible = False

    Else
        PicFolder.Visible = False

        PICpar.Visible = True
    End If


End Sub






Private Sub cmbPREeffect_Change()
    pManual.Visible = IIf(cmbPREeffect.ListIndex = 3, True, False)
    pExposure.Visible = IIf(cmbPREeffect.ListIndex = 2, True, False)
End Sub

Private Sub cmbPREeffect_Click()
    pManual.Visible = IIf(cmbPREeffect.ListIndex = 3, True, False)
    pExposure.Visible = IIf(cmbPREeffect.ListIndex = 2, True, False)

End Sub

Private Sub cmbResizeMode_Click()
    If cmbResizeMode.ListIndex = 0 Then
        tMAXWH.Enabled = False
    Else
        tMAXWH.Enabled = True
    End If


End Sub






Private Sub Command1_Click()
    frmAbout.Show

End Sub

Private Sub Command2_Click()

    Dim s          As String
    Dim S2         As String


    Dim SPath      As String


    Me.MousePointer = 13

    S2 = UpDateSetString
    'SaveSetting "LastSettings.txt"


    SPath = Dir1 & "\"

    If File1 = "" Then MsgBox "Select a Folder/File", vbCritical: Exit Sub

    If chALL.Value = Checked Then
        If chStart.Value = Unchecked Then
            s = Dir(SPath & "*.jpg")
        Else
            s = Dir(SPath & "*.jpg")
            While s <> File1: s = Dir: Wend

        End If
    Else
        s = File1
    End If




    Do
        Me.Caption = "Filering... " & s & " (Wait)"


        PicIN.Cls
        PicIN.Picture = LoadPicture(SPath & s)
        PicIN.Refresh

        INPUTresize Val(tYCrop)

        '-------------------ReducePicBy2
        PIC2.Cls
        'PIC2.Width = PIC1.Width \ 2
        'PIC2.Height = PIC1.Height \ 2
        'PIC2.Refresh
        PIC2.Visible = False
        'SetStretchBltMode PIC2.Hdc, STRETCHMODE
        'StretchBlt PIC2.Hdc, 0, 0, PIC2.Width, PIC2.Height, PIC1.Hdc, 0, 0, PIC1.Width, PIC1.Height, vbSrcCopy
        ''PIC2.Visible = True
        '--------------------------------------------------------------------



        'SetSource2 PIC2.Image.Handle
        SetSource PIC1.Image.Handle
        ReducePicBy2
        ReducePicBy4


        DoPreEFFECT
        'Apply 21, (100 - scrollLIMIT) / 100, sBackDarkness / 100


        Apply3 21, (100 - scrollLIMIT) / 100, sBackDarkness / 100, cBKGmode.ListIndex
        GetEffect PIC1.Image.Handle
        PIC1.Refresh

        BitBlt PIC1.Hdc, 0, 0, PicLogo.Width, PicLogo.Height, PicLogo.Hdc, 0, 0, vbSrcCopy
                
        '    Pic2.Refresh
        '    Pic1.PaintPicture Pic2.Image, 0, 0, Pic1.Width, Pic1.Height, 0, 0, Pic2.Width, Pic2.Height
        '    Pic1.Refresh
        '--------------------------------------------------------------------

        '        SaveJPG PIC1.Image, App.Path & "\OUT\Sketch" & s, JpgQuality
        SaveJPG PIC1.Image, App.Path & "\OUT\" & s, JpgQuality
        '***********************************************************

        'If (Not (PREVIEWmode)) And chALL.Value = Checked Then
        If chALL.Value = Checked Then
            s = Dir
        Else
            s = ""
        End If

    Loop While s <> ""

    Me.Caption = "Filering Done."

    Me.MousePointer = 0

End Sub

Private Sub Dir1_Change()
    'File1 = Dir1 & "\*.jpg"
    File1 = Dir1                  '& "\*.*"
End Sub



Private Sub Drive1_Change()
    Dir1.Path = Drive1

End Sub

Private Sub File1_Click()
    PicIN.Cls

    On Error Resume Next

    PicIN.Picture = LoadPicture(Dir1 & "\" & File1)
    PicIN.Refresh

    INPUTresize Val(tYCrop)


End Sub

Private Sub Form_Activate()

    MAINframe.Width = ShapeBG.Width / Screen.TwipsPerPixelX + 30

    picMOVE.Left = Me.Width / Screen.TwipsPerPixelX - picMOVE.Width - 20


    picMOVE_MouseMove 1, 0, 1, 1


End Sub

Private Sub Form_Initialize()


    InitCommonControls
    'XPStyle False

End Sub
Private Function REPOSParams()
    Dim i          As Long
    Dim H          As Long

    H = H + pPARAM(0).Height
    pPARAM(0).ToP = 10 - ScrollPar.Value * 20
    'pPARAM(0).Left = 0


    For i = 1 To pPARAM.Count - 1
        H = H + pPARAM(i - 1).Height
        pPARAM(i).ToP = pPARAM(i - 1).Height + pPARAM(i - 1).ToP + 5
        'pPARAM(I).Left = 0
    Next


    ParH = 20 + pPARAM(pPARAM.Count - 1).ToP + pPARAM(pPARAM.Count - 1).Height - pPARAM(0).ToP


End Function

Private Sub Form_Load()
    Dim i          As Long

    PaH(0) = 145
    PaH(1) = 265 - 100
    PaH(2) = 129                  '112
    PaH(3) = 95

    If Dir(App.Path & "\OUT", vbDirectory) = "" Then MkDir App.Path & "\OUT"

    File1 = Dir1                  '& "\*.*"

    '    tMAXWH = 520
    MaxWH = Val(tMAXWH)

    scrollLIMIT_Change

    cmbPREeffect.AddItem "0 None"
    cmbPREeffect.AddItem "1 Auto Equalize"
    cmbPREeffect.AddItem "2 Exposure"
    cmbPREeffect.AddItem "3 BCS Manual"

    cmbPREeffect.ListIndex = 0

    cmbResizeMode.AddItem "No Resize"
    cmbResizeMode.AddItem "by Longest Late ="
    cmbResizeMode.AddItem "by Shortest Late ="
    cmbResizeMode.AddItem "by Area (Kpixels)="

    cmbResizeMode.ListIndex = 2


    cBKGmode.AddItem "mode V3 (faster)"
    cBKGmode.AddItem "mode V4 (slower) "
    cBKGmode.ListIndex = 1
    

    'LoadSetting "LastSettings.txt"

    If App.LogMode = 0 Then
        MsgBox "Compile me!", vbInformation
    Else
        ProcessPrioritySet 0, 0, ppbelownormal
    End If

    PICpar.ToP = PicFolder.ToP
    PICpar.Left = PicFolder.Left
    PICpar.Height = PicFolder.Height

    ScrollPar.Height = PICpar.ScaleHeight


    For i = 0 To 1
        lParam_Click (i)
        'lParam_Click (0)
    Next


    'GRAD.Angle = -75
    'GRAD.Color1 = RGB(120, 200, 120)
    'GRAD.Color2 = RGB(80, 120, 80)
    'GRAD.Draw frmMAIN.PICpar
    'SavePicture frmMAIN.PICpar.Image, App.Path & "\Gradgreen.bmp"

    '    SetupStroke2 0
    '    Stop

    SetupStroke3 0

    InitEDGEFilter 1              ' 2
    'InitBACKFilter 16


    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End


End Sub






'Private Sub HHH_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'frmHelp.Show
'frmHelp.ShowHelp Index
'
'End Sub

Private Sub lParam_Click(Index As Integer)


    If pPARAM(Index).Height <> 20 Then
        pPARAM(Index).Height = 20
    Else
        pPARAM(Index).Height = PaH(Index)
    End If
    DoEvents

    REPOSParams

    ScrollPar.max = 0.05 * (ParH - PICpar.ScaleHeight)


    If ScrollPar.max > 0 Then
        ScrollPar.Visible = True
    Else
        ScrollPar.Visible = False
    End If

End Sub

Private Sub MainFrameLabel_Click()
    MAINframe.Height = IIf(MAINframe.Height > 18, 18, 625)


End Sub





Private Sub PIC1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If PREVIEWmode Then
    '    PIC2.Visible = False
    '    If Rect = False Then
    '        pX1 = X
    '        pY1 = Y
    '    End If
    '    Rect = Not Rect
    'End If

End Sub

Private Sub PIC1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    'If PREVIEWmode Then
    '    If Rect Then
    '        pX2 = X
    '        pY2 = Y
    '
    '            If pX2 < pX1 Then
    '                sP.Left = pX2
    '                sP.Width = pX1 - pX2
    '            Else
    '                sP.Left = pX1
    '                sP.Width = pX2 - pX1
    '            End If
    '
    '            If pY2 < pY1 Then
    '                sP.ToP = pY2
    '                sP.Height = pY1 - pY2
    '            Else
    '                sP.ToP = pY1
    '                sP.Height = pY2 - pY1
    '            End If
    '
    '        End If
    '    End If
End Sub


Private Sub picMOVE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picMOVE.Left = picMOVE.Left + X - picMOVE.Width \ 2
        picMOVE.ToP = picMOVE.ToP + Y - picMOVE.Height \ 2
        MAINframe.Left = picMOVE.Left - MAINframe.Width + picMOVE.Width
        'MAINframe.Left = picMOVE.Left
        MAINframe.ToP = picMOVE.ToP    '+ picMOVE.Height \ 2
    End If

End Sub

Private Sub resetB_Click()
    sBRIGHT = (sBRIGHT.max + sBRIGHT.min) * 0.5
End Sub

Private Sub ResetExpo_Click()
    sEXPO = 0
End Sub

Private Sub resetS_Click()
    sSATUR = (sSATUR.max + sSATUR.min) * 0.5
End Sub

Private Sub restC_Click()
    sCONTRA = 0
End Sub

Private Sub sBackDarkness_Change()
    If sBackDarkness <> 0 Then
        Label2 = "BackGround Darkness " & sBackDarkness
    Else
        Label2 = "No BackGround"
    End If

End Sub

Private Sub sBackDarkness_Scroll()
    If sBackDarkness <> 0 Then
        Label2 = "BackGround Darkness " & sBackDarkness
    Else
        Label2 = "No BackGround"
    End If
End Sub

Private Sub sBRIGHT_Change()
    lBRIGHT = "Bright : " & Int(200 * (sBRIGHT.Value / sBRIGHT.max)) & "%"
End Sub

Private Sub sCONTRA_Change()
    lCONTRA = "Contra : " & sCONTRA & "%"
End Sub






Private Sub ScrollPar_Change()
    REPOSParams
End Sub



Private Sub sEXPO_Change()
    lEXPO = "Exposure " & sEXPO
End Sub

Private Sub sEXPO_Scroll()
    lEXPO = "Exposure " & sEXPO
End Sub

Private Sub sRadius_Change()
    Label1 = "Edge Detect Radius " & sRadius

End Sub

Private Sub sRadius_Scroll()
    Label1 = "Edge Detect Radius " & sRadius

End Sub

Private Sub sSATUR_Change()
    lSATUR = "Satura : " & Int(200 * (sSATUR.Value / sSATUR.max)) & "%"

End Sub
Private Sub sBRIGHT_Scroll()
    lBRIGHT = "Bright : " & Int(200 * (sBRIGHT.Value / sBRIGHT.max)) & "%"
End Sub

Private Sub sCONTRA_Scroll()
    lCONTRA = "Contra : " & sCONTRA & "%"
End Sub
Private Sub sSATUR_Scroll()
    lSATUR = "Satura : " & Int(200 * (sSATUR.Value / sSATUR.max)) & "%"

End Sub



Private Sub scrollLIMIT_Change()

    lLIMIT = "Limit " & 100 - scrollLIMIT
End Sub

Private Sub scrollLIMIT_Scroll()

    lLIMIT = "Limit " & 100 - scrollLIMIT
End Sub







Private Sub SK_PercDONE(Value As Single, CurrIteration As Long)
    ShapeProg.Width = ShapeBG.Width * Value
    LabProg = Int(Value * 100) & "%  Iteration " & CurrIteration & ""
    DoEvents
End Sub







Private Sub tMAXWH_Change()
    MaxWH = Val(tMAXWH)
End Sub


Public Sub SaveSetting(ByVal F As String)
    Open App.Path & "\" & F For Output As 1

    Print #1, cmbPREeffect
    Print #1, sEXPO
    Print #1, sBRIGHT
    Print #1, sCONTRA
    Print #1, sSATUR


    Print #1, Dir1



    Close 1

    F = Left$(F, Len(F) - 4) & "EX.txt"
    Open App.Path & "\" & F For Output As 1
    Print #1, ExtendSetting
    Close 1

End Sub
Public Sub LoadSetting(F As String)
    Dim s          As String
    Dim N          As Single

    Open App.Path & "\" & F For Input As 1



    If Not EOF(1) Then
        Input #1, s: cmbPREeffect.ListIndex = Val(Left$(s, 1))
        Input #1, N: sEXPO = N
        Input #1, N: sBRIGHT = N
        Input #1, N: sCONTRA = N
        Input #1, N: sSATUR = N

    End If

    If Not EOF(1) Then
        Input #1, s
        If Dir(s, vbDirectory) <> "" Then Drive1 = Left$(s, 2): Dir1 = s


    End If
    Close 1
    UpDateSetString
End Sub

Private Sub PrintTextToPic(txt As String, ByRef Pic As PictureBox)
    Pic.CurrentX = 6              '+ 1
    Pic.CurrentY = Pic.Height - 33 + 1
    Pic.ForeColor = vbBlack
    Pic.Print txt

    Pic.CurrentX = 6              '- 1
    Pic.CurrentY = Pic.Height - 33 - 1
    Pic.ForeColor = vbWhite
    Pic.Print txt

    Pic.CurrentX = 6
    Pic.CurrentY = Pic.Height - 33
    Pic.ForeColor = RGB(127, 127, 127)
    Pic.Print txt
End Sub


Public Sub DoPreEFFECT()

    Select Case cmbPREeffect.ListIndex    '
        Case 0
        Case 1
            MagneKleverHistogramEQU 0.3
        Case 2
            MagneKleverExposure sEXPO
        Case 3
            MagneKleverBCS sBRIGHT, sCONTRA, sSATUR
    End Select

End Sub



Public Function UpDateSetString() As String


    UpDateSetString = ""
    ExtendSetting = ""




    Select Case cmbPREeffect.ListIndex

        Case 0
            UpDateSetString = UpDateSetString & " pFX:none "
            ExtendSetting = ExtendSetting & "PreEFFECT: None" & vbCrLf & vbCrLf
        Case 1
            UpDateSetString = UpDateSetString & " pFX:AutoEqu "
            ExtendSetting = ExtendSetting & "PreEFFECT: Auto Equalize" & vbCrLf & vbCrLf

        Case 2
            UpDateSetString = UpDateSetString & " pFX:Exposure " & sEXPO
            ExtendSetting = ExtendSetting & "PreEFFECT:" & vbCrLf
            ExtendSetting = ExtendSetting & vbTab & "Exposure:" & sEXPO & vbCrLf & vbCrLf

        Case 3
            UpDateSetString = UpDateSetString & " pFX:BCS " & Int(200 * (sBRIGHT.Value / sBRIGHT.max)) & _
                              " " & sCONTRA & _
                              " " & Int(200 * (sSATUR.Value / sSATUR.max))

            ExtendSetting = ExtendSetting & "PreEFFECT:" & vbCrLf
            ExtendSetting = ExtendSetting & vbTab & "Brightness:" & Int(200 * (sBRIGHT.Value / sBRIGHT.max)) & vbCrLf
            ExtendSetting = ExtendSetting & vbTab & "Contrast  :" & sCONTRA & vbCrLf
            ExtendSetting = ExtendSetting & vbTab & "Saturation:" & Int(200 * (sSATUR.Value / sSATUR.max)) & vbCrLf & vbCrLf


    End Select

    ExtendSetting = ExtendSetting & "BILATERAL FILTER: " & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Color Space: " & IIf(oRGB, "RGB", "CieLAB") & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Radius: " & tRad & vbCrLf

    ExtendSetting = ExtendSetting & vbTab & "Intensity Sigma: " & tSIGMA & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Spatial   Sigma: " & tSigmaSpatial & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Iterations: " & ITER & vbCrLf & vbCrLf

    ExtendSetting = ExtendSetting & "CONTOUR: " & vbCrLf

    ExtendSetting = ExtendSetting & vbTab & "Amount : " & tCONT & vbCrLf
    'If cmbContourMode.ListIndex = 0 Then
    ExtendSetting = ExtendSetting & vbTab & "Lum/(A&B): " & tLUMHUE & vbCrLf
    'Else
    '    ExtendSetting = ExtendSetting & vbTab & "Threshold: " & tLUMHUE & vbCrLf
    'End If

    'MsgBox ExtendSetting


End Function

Private Sub INPUTresize(YCrop)
    Dim KArea      As Double

    PIC1.Cls

    Select Case cmbResizeMode.ListIndex


        Case 0
            PIC1.Width = PicIN.Width
            PIC1.Height = (PicIN.Height - YCrop * 2)

        Case 1
            If PicIN.Width > (PicIN.Height - YCrop * 2) Then
                PIC1.Width = MaxWH
                PIC1.Height = Fix((PicIN.Height - YCrop * 2) / PicIN.Width * PIC1.Width)
            Else
                PIC1.Height = MaxWH
                PIC1.Width = Fix(PicIN.Width / (PicIN.Height - YCrop * 2) * PIC1.Height)
            End If

        Case 2
            If PicIN.Width < (PicIN.Height - YCrop * 2) Then
                PIC1.Width = MaxWH
                PIC1.Height = Fix((PicIN.Height - YCrop * 2) / PicIN.Width * PIC1.Width)
            Else
                PIC1.Height = MaxWH
                PIC1.Width = Fix(PicIN.Width / (PicIN.Height - YCrop * 2) * PIC1.Height)
            End If

        Case 3
            KArea = (CDbl(PicIN.Width) * CDbl((PicIN.Height - YCrop * 2))) / (CDbl(MaxWH) * 1024)
            KArea = Sqr(KArea)
            PIC1.Width = (PicIN.Width / KArea) \ 1
            PIC1.Height = ((PicIN.Height - YCrop * 2) / KArea) \ 1
            '           MsgBox PIC1.Width * PIC1.Height

    End Select

    PIC1.Width = PIC1.Width \ 1
    PIC1.Height = PIC1.Height \ 1

    'While PIC1.Width Mod 4 <> 0: PIC1.Width = PIC1.Width - 1: Wend
    'While PIC1.Height Mod 4 <> 0: PIC1.Height = PIC1.Height - 1: Wend
    PIC1.Width = PIC1.Width - (PIC1.Width Mod 4)
    PIC1.Height = PIC1.Height - (PIC1.Height Mod 4)

    SetStretchBltMode PIC1.Hdc, vbPaletteModeNone
    StretchBlt PIC1.Hdc, 0, 0, PIC1.Width, PIC1.Height, PicIN.Hdc, 0, YCrop, PicIN.Width - 1, (PicIN.Height - YCrop * 2) - 1, vbSrcCopy
    PIC1.Refresh

End Sub
