VERSION 5.00
Begin VB.Form frmMAIN 
   BackColor       =   &H00808080&
   Caption         =   "Multipass Bilateral Filter - Cartoonizer (by Roberto Mior) [08Gen2011]"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
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
         TabIndex        =   69
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   60
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
            TabIndex        =   73
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
            TabIndex        =   64
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   61
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
         Picture         =   "frmMAIN.frx":0000
         ScaleHeight     =   553
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   233
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
         Begin VB.PictureBox pPARAM 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1935
            Index           =   2
            Left            =   120
            ScaleHeight     =   127
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   191
            TabIndex        =   35
            Top             =   4560
            Width           =   2895
            Begin VB.ComboBox cmbContourMode 
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
               TabIndex        =   75
               ToolTipText     =   "Intensity Curve"
               Top             =   360
               Width           =   1695
            End
            Begin VB.HScrollBar scrLUMHUE 
               Height          =   255
               LargeChange     =   5
               Left            =   120
               Max             =   100
               SmallChange     =   5
               TabIndex        =   72
               Top             =   1440
               Value           =   100
               Width           =   1695
            End
            Begin VB.HScrollBar scrCONT 
               Height          =   255
               LargeChange     =   5
               Left            =   120
               Max             =   200
               SmallChange     =   5
               TabIndex        =   71
               Top             =   960
               Value           =   100
               Width           =   1695
            End
            Begin VB.Label lLumHue 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cont Lum-Hue (0-1)"
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
               TabIndex        =   38
               Top             =   1200
               Width           =   2415
            End
            Begin VB.Label lContour 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Contour Amount 0-100"
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
               TabIndex        =   37
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label lParam 
               BackColor       =   &H00E0E0E0&
               Caption         =   "* CONTOUR"
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
               Index           =   2
               Left            =   0
               TabIndex        =   36
               ToolTipText     =   "Hide/Show ""Contour"""
               Top             =   0
               Width           =   2895
            End
         End
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
            TabIndex        =   17
            Top             =   1320
            Width           =   2895
            Begin VB.CheckBox chDirectional 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Directional"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2040
               TabIndex        =   74
               Top             =   480
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.ComboBox CmbColorSpace 
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
               TabIndex        =   59
               Top             =   360
               Width           =   1695
            End
            Begin VB.ComboBox cmbIntensityMode 
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
               TabIndex        =   51
               ToolTipText     =   "Intensity Curve"
               Top             =   1680
               Width           =   1695
            End
            Begin VB.HScrollBar scrSpatial 
               Height          =   255
               Left            =   120
               Max             =   10000
               TabIndex        =   50
               Top             =   3000
               Value           =   200
               Width           =   1695
            End
            Begin VB.HScrollBar scrITER 
               Height          =   255
               Left            =   120
               Max             =   25
               Min             =   1
               TabIndex        =   49
               Top             =   3600
               Value           =   4
               Width           =   1695
            End
            Begin VB.HScrollBar scrRAD 
               Height          =   255
               Left            =   120
               Max             =   30
               Min             =   1
               TabIndex        =   48
               Top             =   1080
               Value           =   3
               Width           =   1695
            End
            Begin VB.HScrollBar scrIntensitySigma 
               Height          =   255
               Left            =   120
               Max             =   1000
               TabIndex        =   47
               Top             =   2400
               Value           =   350
               Width           =   1695
            End
            Begin VB.PictureBox PicSpatial 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   735
               Left            =   2040
               ScaleHeight     =   47
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   47
               TabIndex        =   46
               ToolTipText     =   "Spatial Domain"
               Top             =   3000
               Width           =   735
            End
            Begin VB.PictureBox SigmaPrev 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   735
               Left            =   2040
               ScaleHeight     =   47
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   47
               TabIndex        =   45
               ToolTipText     =   "X = Delta Intensity from Central Point. (displaied range 0-32 of 255) .  Y = % Used (0-100)"
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label lIter 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Iterations"
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
               TabIndex        =   58
               Top             =   3360
               Width           =   2655
            End
            Begin VB.Label Label7 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Intensity MODE"
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
               TabIndex        =   57
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label Label5 
               BackColor       =   &H00C0C0C0&
               Caption         =   "SPATIAL Sigma"
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
               TabIndex        =   56
               Top             =   2760
               Width           =   1935
            End
            Begin VB.Label Label3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Radius (HalfLate)"
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
               TabIndex        =   55
               Top             =   840
               Width           =   2655
            End
            Begin VB.Label Label2 
               BackColor       =   &H00C0C0C0&
               Caption         =   "INTENSITY Sigma 0-1"
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
               TabIndex        =   54
               Top             =   2160
               Width           =   1695
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Spatial"
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
               Left            =   2040
               TabIndex        =   53
               Top             =   2760
               Width           =   1215
            End
            Begin VB.Label Label 
               BackStyle       =   0  'Transparent
               Caption         =   "Intensity"
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
               Left            =   1920
               TabIndex        =   52
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label lParam 
               BackColor       =   &H00E0E0E0&
               Caption         =   "* BILATERAL Filter"
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
               TabIndex        =   33
               ToolTipText     =   "Hide/Show ""Bilateral Filter"""
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
            TabIndex        =   16
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
               TabIndex        =   32
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
               TabIndex        =   22
               Top             =   600
               Width           =   1575
               Begin VB.HScrollBar sBRIGHT 
                  Height          =   255
                  Left            =   0
                  Max             =   512
                  TabIndex        =   28
                  Top             =   240
                  Value           =   256
                  Width           =   1335
               End
               Begin VB.HScrollBar sCONTRA 
                  Height          =   255
                  Left            =   0
                  Max             =   100
                  Min             =   -100
                  TabIndex        =   27
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.HScrollBar sSATUR 
                  Height          =   255
                  Left            =   0
                  Max             =   512
                  TabIndex        =   26
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
                  TabIndex        =   25
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
                  TabIndex        =   24
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
                  TabIndex        =   23
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
                  TabIndex        =   31
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
                  TabIndex        =   30
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
                  TabIndex        =   29
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
               TabIndex        =   18
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
                  TabIndex        =   20
                  ToolTipText     =   "Reset Exposure"
                  Top             =   240
                  Width           =   255
               End
               Begin VB.HScrollBar sEXPO 
                  Height          =   255
                  Left            =   0
                  Max             =   256
                  Min             =   -127
                  TabIndex        =   19
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
                  TabIndex        =   21
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
               TabIndex        =   34
               ToolTipText     =   "Hide/Show ""Pre EFFECT"""
               Top             =   0
               Width           =   2895
            End
         End
         Begin VB.PictureBox pPARAM 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1455
            Index           =   3
            Left            =   120
            ScaleHeight     =   95
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   191
            TabIndex        =   39
            Top             =   6600
            Width           =   2895
            Begin VB.CheckBox chIsVideo 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Video Mode"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   1920
               TabIndex        =   70
               ToolTipText     =   "Check this if you are Computing large Number of Video Frames."
               Top             =   600
               Width           =   855
            End
            Begin VB.HScrollBar scrPRES 
               Height          =   255
               LargeChange     =   5
               Left            =   120
               Max             =   100
               SmallChange     =   5
               TabIndex        =   43
               Top             =   1080
               Value           =   4
               Width           =   1695
            End
            Begin VB.HScrollBar scrNseg 
               Height          =   255
               Left            =   120
               Max             =   10
               Min             =   2
               TabIndex        =   41
               Top             =   600
               Value           =   5
               Width           =   1695
            End
            Begin VB.Label lPres 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Pres"
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
               TabIndex        =   44
               Top             =   840
               Width           =   2415
            End
            Begin VB.Label lNSeg 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nseg"
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
               TabIndex        =   42
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label lParam 
               BackColor       =   &H00E0E0E0&
               Caption         =   "* Lum. SEGMENTATION"
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
               Index           =   3
               Left            =   0
               TabIndex        =   40
               ToolTipText     =   "Hide/Show ""Luminance Segmentation"""
               Top             =   0
               Width           =   2895
            End
         End
         Begin VB.VScrollBar ScrollPar 
            Height          =   5775
            Left            =   3120
            Max             =   1
            TabIndex        =   15
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdTEST 
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   8880
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cndSKETCH 
         BackColor       =   &H00808080&
         Caption         =   "SKETCH"
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
         Left            =   2400
         TabIndex        =   12
         Top             =   8880
         Visible         =   0   'False
         Width           =   1455
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
         TabIndex        =   9
         ToolTipText     =   "Print Parameters to Output Picture(s)"
         Top             =   6960
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "BILATERAL FILTER Cartoonizer"
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
         TabIndex        =   8
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
         TabIndex        =   7
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
      Begin VB.CheckBox chPreviewMode 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Preview Mode"
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
         Height          =   735
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Elaborate only a portion of the image. (With PreEffect AutoEqualize gives wrong result)"
         Top             =   7440
         Width           =   1095
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         Width           =   3615
      End
   End
   Begin VB.PictureBox PIC2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
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
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Cartoonizer - Convert Photos into Cartoon Like Images
'    Copyright (c) 2011 - Roberto Mior
'
'    This file is part of "Bilateral Filter - Cartoonizer".
'
'    "Bilateral Filter - Cartoonizer" is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    "Bilateral Filter - Cartoonizer" is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with "Bilateral Filter - Cartoonizer".  If not, see <http://www.gnu.org/licenses/>.


Option Explicit


Private WithEvents FX As myEffects
Attribute FX.VB_VarHelpID = -1
Private WithEvents SK As clsSketch
Attribute SK.VB_VarHelpID = -1

'Private GRAD As New clsGradient


Private MaxWH      As Integer


Private N          As Long
Private SigmaIntensity As Single
Private SigmaSpatial As Single
Private IntensityMode As Integer


Private pX1        As Integer
Private pY1        As Integer
Private pX2        As Integer
Private pY2        As Integer
Private Rect       As Boolean
Private PREVIEWmode As Boolean


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

Private Sub chIsVideo_Click()
    If chIsVideo Then FX.SetUpHistoChache scrNseg
End Sub

Private Sub chPreviewMode_Click()
    PREVIEWmode = IIf(chPreviewMode.Value = Checked, True, False)
    sP.Visible = PREVIEWmode
    chMakeCompare.Visible = PREVIEWmode
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

Private Sub CmbColorSpace_Click()
    oRGB = IIf(CmbColorSpace.ListIndex = 0, True, False)

End Sub

Private Sub cmbContourMode_Click()
    scrLUMHUE_Change
End Sub

Private Sub cmbIntensityMode_Click()
    IntensityMode = cmbIntensityMode.ListIndex
    FX.zPreview_Intensity SigmaPrev, SigmaIntensity, IntensityMode

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

Private Sub cmdTEST_Click()

    PIC2.Cls

    PIC2.Width = (PIC1.Width) * 3
    PIC2.Height = (PIC1.Height) * 3
    PIC2.Visible = True

    PIC2.Top = PIC1.Top + PIC1.Height / 2 + 10
    PIC2.Left = PIC1.Left

    '    FX.IWH PIC1.Image.Handle, PIC2.Image.Handle

    MsgBox "   FX.TEST2 PIC1.Image.Handle"
End Sub

Private Sub cndSKETCH_Click()
    Dim S          As String
    Dim S2         As String


    Dim SPath      As String

    If File1 = "" Then MsgBox "Select a Folder/File", vbCritical: Exit Sub

    If chALL.Value = Checked Then
        If chStart.Value = Unchecked Then
            S = Dir(SPath & "*.jpg")
        Else
            S = Dir(SPath & "*.jpg")
            While S <> File1: S = Dir: Wend

        End If
    Else
        S = File1
    End If


    PIC2.Visible = False

    Do
        Me.Caption = "Sketching... " & S & " (Wait)"

        ' S2 = "CARTOON " & "R=" & tRad & "  I=" & tSIGMA & "  S=" & tSigmaSpatial & "  IT=" & ITER & "  C=" & tCONT & "|" & tLumHue & "  iMode=" & IntensityMode & "  " & IIf(oRGB, "RGB", "CieLAB") & _
          "  BCS " & Int(200 * (sBRIGHT.Value / sBRIGHT.max)) & _
          " " & sCONTRA & _
          " " & Int(200 * (sSATUR.Value / sSATUR.max))



        PicIN.Cls
        PicIN.Picture = LoadPicture(SPath & S)
        PicIN.Refresh


        INPUTresize Val(tYCrop)


        SetStretchBltMode PIC1.Hdc, vbPaletteModeNone
        StretchBlt PIC1.Hdc, 0, 0, PIC1.Width, PIC1.Height, PicIN.Hdc, 0, 0, PicIN.Width - 1, PicIN.Height - 1, vbSrcCopy
        PIC1.Refresh

        FX.zSet_Source PIC1.Image.Handle
        'SK.SetSource PIC1.Image.Handle


        'FX.zEFF_MedianFilter 1, 1
        If chDirectional Then
            FX.zEFF_BilateralFilterDirectional N, ITER, oRGB
        Else
            FX.zEFF_BilateralFilter N, ITER, oRGB
        End If


        FX.zGet_Effect PIC1.Image.Handle

        SK.SetSource PIC1.Image.Handle

        SK.SetSourceToMIX PIC1.Image.Handle
        SK.Sketch
        SK.MIX Val(tCONT)
        SK.GetEffect PIC1.Image.Handle



        If chPrintParams.Value = Checked Then
            PrintTextToPic S2, PIC1
        End If

        SaveJPG PIC1.Image, App.Path & "\OUT\Sketch" & S, JpgQuality
        '***********************************************************



        If (Not (PREVIEWmode)) And chALL.Value = Checked Then
            S = Dir
        Else
            S = ""
        End If

    Loop While S <> ""

    Me.Caption = "Sketching Done."
End Sub

Private Sub Command2_Click()

    Dim S          As String
    Dim S2         As String


    Dim SPath      As String

    Me.MousePointer = 13

    S2 = UpDateSetString
    SaveSetting "LastSettings.txt"


    SPath = Dir1 & "\"

    If File1 = "" Then MsgBox "Select a Folder/File", vbCritical: Exit Sub

    If chALL.Value = Checked Then
        If chStart.Value = Unchecked Then
            S = Dir(SPath & "*.jpg")
        Else
            S = Dir(SPath & "*.jpg")
            While S <> File1: S = Dir: Wend

        End If
    Else
        S = File1
    End If


    PIC2.Cls
    PIC2.Width = sP.Width - 1
    PIC2.Height = sP.Height - 1
    PIC2.Refresh
    PIC2.Visible = PREVIEWmode
    PIC2.Top = PIC1.Top + sP.Top
    PIC2.Left = PIC1.Left + sP.Left



    Do
        Me.Caption = "Filering... " & S & " (Wait)"


        PicIN.Cls
        PicIN.Picture = LoadPicture(SPath & S)
        PicIN.Refresh

        INPUTresize Val(tYCrop)



        If PREVIEWmode Then
            If chMakeCompare.Value = Checked Then
                PicIN.Cls
                PicIN.Width = PIC2.Width * 2 - 1
                PicIN.Height = PIC2.Height
                PicIN.Refresh
                SetStretchBltMode PicIN.Hdc, vbPaletteModeNone
                StretchBlt PicIN.Hdc, 0, 0, PIC2.Width - 1, PIC2.Height - 1, PIC1.Hdc, sP.Left, sP.Top, PIC2.Width - 1, PIC2.Height - 1, vbSrcCopy
                PicIN.Refresh

            End If


            SetStretchBltMode PIC2.Hdc, vbPaletteModeNone
            StretchBlt PIC2.Hdc, 0, 0, PIC2.Width, PIC2.Height, PIC1.Hdc, sP.Left, sP.Top, PIC2.Width, PIC2.Height, vbSrcCopy
            PIC2.Refresh
            SaveJPG PIC2.Image, App.Path & "\OUT\EFForig" & S, JpgQuality
            FX.zSet_Source PIC2.Image.Handle

            'FX.PreBrightNessAndContrast source, -sBRIGHT / 100, -sCONTRA / 100
            DoPreEFFECT



            'FX.zEFF_BLUR 1, 1
            If chDirectional Then
                FX.zEFF_BilateralFilterDirectional N, ITER, oRGB
            Else
                FX.zEFF_BilateralFilter N, ITER, oRGB
            End If

            If tCONT > 0 Then
                If cmbContourMode.ListIndex = 0 Then
                    FX.zEFF_ContourBySobel tCONT, tLUMHUE
                Else
                    FX.zEFF_ContourBySobel2 tCONT, tLUMHUE
                End If

                FX.zEFF_QuantizeLuminance scrNseg, scrPRES / 100, chIsVideo
                FX.zEFF_Contour_Apply
            Else
                FX.zEFF_QuantizeLuminance scrNseg, scrPRES / 100, chIsVideo
            End If

            'FX.zGet_Effect PIC2.Image.Handle
            'FX.zSet_Source PIC2.Image.Handle
            'FX.zEFF_MedianFilter 1, 5

            FX.zGet_Effect PIC2.Image.Handle
            SaveJPG PIC2.Image, App.Path & "\OUT\EFFprev" & S, JpgQuality
            If chMakeCompare.Value = Checked Then
                StretchBlt PicIN.Hdc, PIC2.Width - 1, 0, PIC2.Width - 1, PIC2.Height - 1, PIC2.Hdc, 0, 0, PIC2.Width - 1, PIC2.Height - 1, vbSrcCopy
                PicIN.Refresh
                If chPrintParams.Value = Checked Then
                    PrintTextToPic S2, PicIN
                End If
                SaveJPG PicIN.Image, App.Path & "\OUT\Compare_" & S, JpgQuality
            End If
        Else

            '**************************************************
            FX.zSet_Source PIC1.Image.Handle

            'FX.PreBrightNessAndContrast source, -sBRIGHT / 100, -sCONTRA / 100
            DoPreEFFECT


            'FX.zEFF_BLUR 1, 1
            If chDirectional Then
                FX.zEFF_BilateralFilterDirectional N, ITER, oRGB
            Else
                FX.zEFF_BilateralFilter N, ITER, oRGB
            End If
            If tCONT > 0 Then
                If cmbContourMode.ListIndex = 0 Then
                    FX.zEFF_ContourBySobel tCONT, tLUMHUE
                Else
                    FX.zEFF_ContourBySobel2 tCONT, tLUMHUE
                End If
                FX.zEFF_QuantizeLuminance scrNseg, scrPRES / 100, chIsVideo
                FX.zEFF_Contour_Apply
            Else
                FX.zEFF_QuantizeLuminance scrNseg, scrPRES / 100, chIsVideo
            End If

            'FX.zGet_Effect PIC1.Image.Handle
            'FX.zSet_Source PIC1.Image.Handle
            'FX.zEFF_MedianFilter 1, 5

            FX.zGet_Effect PIC1.Image.Handle
            If chPrintParams.Value = Checked Then
                PrintTextToPic S2, PIC1
            End If

            SaveJPG PIC1.Image, App.Path & "\OUT\" & S, JpgQuality
            '***********************************************************

        End If

        If (Not (PREVIEWmode)) And chALL.Value = Checked Then
            S = Dir
        Else
            S = ""
        End If

    Loop While S <> ""

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
'XPStyle False

End Sub
Private Function REPOSParams()
    Dim I          As Long
    Dim H          As Long

    H = H + pPARAM(0).Height
    pPARAM(0).Top = 10 - ScrollPar.Value * 20
    'pPARAM(0).Left = 0


    For I = 1 To pPARAM.Count - 1
        H = H + pPARAM(I - 1).Height
        pPARAM(I).Top = pPARAM(I - 1).Height + pPARAM(I - 1).Top + 5
        'pPARAM(I).Left = 0
    Next


    ParH = 20 + pPARAM(pPARAM.Count - 1).Top + pPARAM(pPARAM.Count - 1).Height - pPARAM(0).Top


End Function

Private Sub Form_Load()
    Dim I          As Long


    PaH(0) = 145
    PaH(1) = 265
    PaH(2) = 129                  '112
    PaH(3) = 95


    ProcessPrioritySet 0, 0, ppbelownormal

    'Me.Caption = Me.Caption & " V" & App.Major & "." & App.Minor

    Set FX = New myEffects
    Set SK = New clsSketch


    If Dir(App.Path & "\OUT", vbDirectory) = "" Then MkDir App.Path & "\OUT"

    'File1 = Dir1 & "\*.jpg"
    File1 = Dir1                  '& "\*.*"



    '    tMAXWH = 520
    MaxWH = Val(tMAXWH)


    FX.zInit_SpatialDomain 10

    scrRAD_Change
    scrIntensitySigma_Change
    scrSpatial_Change
    scrITER_Change


    cmbIntensityMode.AddItem "0 Gaussian"
    cmbIntensityMode.AddItem "1 Gaussian2"
    cmbIntensityMode.AddItem "2 InvProp"
    cmbIntensityMode.AddItem "3 Linear"
    '  cmbIntensityMode.AddItem "4 Cosine"
    cmbIntensityMode.ListIndex = 0

    cmbPREeffect.AddItem "0 None"
    cmbPREeffect.AddItem "1 Auto Equalize"
    cmbPREeffect.AddItem "2 Exposure"
    cmbPREeffect.AddItem "3 BCS Manual"


    cmbPREeffect.ListIndex = 0


    CmbColorSpace.AddItem "RGB"
    CmbColorSpace.AddItem "Cie LAB"
    CmbColorSpace.ListIndex = 0

    cmbResizeMode.AddItem "No Resize"
    cmbResizeMode.AddItem "by Longest Late ="
    cmbResizeMode.AddItem "by Shortest Late ="
    cmbResizeMode.AddItem "by Area (Kpixels)="

    cmbResizeMode.ListIndex = 2


    cmbContourMode.AddItem "Sobel"
    cmbContourMode.AddItem "Sobel BOLD"
    cmbContourMode.ListIndex = 0


    LoadSetting "LastSettings.txt"


    FX.zInit_IntensityDomain SigmaIntensity, IntensityMode

    If App.LogMode = 0 Then MsgBox "Compile me!", vbInformation

    PICpar.Top = PicFolder.Top
    PICpar.Left = PicFolder.Left
    PICpar.Height = PicFolder.Height


    ScrollPar.Height = PICpar.ScaleHeight



    For I = 0 To 3
        lParam_Click (I)
        'lParam_Click (0)
    Next


    'GRAD.Angle = -75
    'GRAD.Color1 = RGB(120, 200, 120)
    'GRAD.Color2 = RGB(80, 120, 80)
    'GRAD.Draw frmMAIN.PICpar
    'SavePicture frmMAIN.PICpar.Image, App.Path & "\Gradgreen.bmp"


End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "LastSettings.txt"
    End
End Sub


Private Sub FX_PercDONE(FT As String, Value As Single, CI As Long)
    ShapeProg.Width = ShapeBG.Width * Value
    LabProg = FT & " " & Int(Value * 100) & "%  Iter. " & CI & ""
    DoEvents
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
    If PREVIEWmode Then
        PIC2.Visible = False
        If Rect = False Then
            pX1 = X
            pY1 = Y
        End If
        Rect = Not Rect
    End If

End Sub

Private Sub PIC1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    If PREVIEWmode Then
        If Rect Then
            pX2 = X
            pY2 = Y

            If pX2 < pX1 Then
                sP.Left = pX2
                sP.Width = pX1 - pX2
            Else
                sP.Left = pX1
                sP.Width = pX2 - pX1
            End If

            If pY2 < pY1 Then
                sP.Top = pY2
                sP.Height = pY1 - pY2
            Else
                sP.Top = pY1
                sP.Height = pY2 - pY1
            End If

        End If
    End If
End Sub


Private Sub picMOVE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picMOVE.Left = picMOVE.Left + X - picMOVE.Width \ 2
        picMOVE.Top = picMOVE.Top + Y - picMOVE.Height \ 2
        MAINframe.Left = picMOVE.Left - MAINframe.Width + picMOVE.Width
        'MAINframe.Left = picMOVE.Left
        MAINframe.Top = picMOVE.Top    '+ picMOVE.Height \ 2
    End If

End Sub

Private Sub resetB_Click()
    sBRIGHT = (sBRIGHT.max + sBRIGHT.Min) * 0.5
End Sub

Private Sub ResetExpo_Click()
    sEXPO = 0
End Sub

Private Sub resetS_Click()
    sSATUR = (sSATUR.max + sSATUR.Min) * 0.5
End Sub

Private Sub restC_Click()
    sCONTRA = 0
End Sub

Private Sub sBRIGHT_Change()
    lBRIGHT = "Bright : " & Int(200 * (sBRIGHT.Value / sBRIGHT.max)) & "%"
End Sub

Private Sub sCONTRA_Change()
    lCONTRA = "Contra : " & sCONTRA & "%"
End Sub



Private Sub scrCONT_Change()
    tCONT = scrCONT
    lContour = "Contour Amount " & tCONT

End Sub

Private Sub scrCONT_Scroll()
    tCONT = scrCONT
    lContour = "Contour Amount " & tCONT

End Sub

Private Sub scrLUMHUE_Change()
    tLUMHUE = scrLUMHUE / 100
    'If cmbContourMode.ListIndex = 0 Then
    lLumHue = "Based on " & (100 - scrLUMHUE) & "%Lum " & scrLUMHUE & "%AB"
    'Else
    '    lLumHue = "Threshold " & tLUMHUE
    'End If

End Sub

Private Sub scrLUMHUE_Scroll()
    tLUMHUE = scrLUMHUE / 100
    'If cmbContourMode.ListIndex = 0 Then
    lLumHue = "Based on " & (100 - scrLUMHUE) & "%Lum " & scrLUMHUE & "%AB"
    'Else
    '    lLumHue = "Threshold " & tLUMHUE
    'End If

End Sub

Private Sub scrNseg_Change()
    lNSeg = "Segments " & scrNseg
    If chIsVideo Then FX.SetUpHistoChache scrNseg

End Sub

Private Sub scrNseg_Scroll()
    lNSeg = "Segments " & scrNseg
    If chIsVideo Then FX.SetUpHistoChache scrNseg

End Sub

Private Sub ScrollPar_Change()
    REPOSParams
End Sub

Private Sub scrPRES_Change()
    lPres = "Presence " & scrPRES & "%"
End Sub

Private Sub scrPRES_Scroll()
    lPres = "Presence " & scrPRES & "%"
End Sub

Private Sub sEXPO_Change()
    lEXPO = "Exposure " & sEXPO
End Sub

Private Sub sEXPO_Scroll()
    lEXPO = "Exposure " & sEXPO
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
Private Sub scrIntensitySigma_Change()
    tSIGMA = scrIntensitySigma / 1000

    SigmaIntensity = Val(Replace(tSIGMA, ",", ".")) * 0.1
    FX.zPreview_Intensity SigmaPrev, SigmaIntensity, IntensityMode

    Label2 = "INTENS S " & tSIGMA
End Sub

Private Sub scrIntensitySigma_Scroll()
    tSIGMA = scrIntensitySigma / 1000
    SigmaIntensity = Val(Replace(tSIGMA, ",", ".")) * 0.1
    FX.zPreview_Intensity SigmaPrev, SigmaIntensity, IntensityMode
    Label2 = "INTENS S " & tSIGMA
End Sub

Private Sub scrITER_Change()
    ITER = scrITER
    lIter = "Iterations " & ITER
End Sub

Private Sub scrITER_Scroll()
    ITER = scrITER
    lIter = "Iterations " & ITER
End Sub
Private Sub scrRAD_Change()
    tRad = scrRAD
    N = Val(tRad)
    If N < 1 Then N = 1: tRad = N
    FX.zPreview_Spatial PicSpatial, N, SigmaSpatial
    Label3 = "Radius (HalfLate) = " & N

End Sub

Private Sub scrRAD_Scroll()
    tRad = scrRAD
    N = Val(tRad)
    If N < 1 Then N = 1: tRad = N
    FX.zPreview_Spatial PicSpatial, N, SigmaSpatial
    Label3 = "Radius (HalfLate) = " & N

End Sub

Private Sub scrSpatial_Change()
    tSigmaSpatial = scrSpatial / 10
    SigmaSpatial = Val(Replace(tSigmaSpatial, ",", "."))
    FX.zPreview_Spatial PicSpatial, N, SigmaSpatial

    Label5.Caption = "SPATIAL S " & SigmaSpatial


End Sub

Private Sub scrSpatial_Scroll()
    tSigmaSpatial = scrSpatial / 10
    SigmaSpatial = Val(Replace(tSigmaSpatial, ",", "."))
    FX.zPreview_Spatial PicSpatial, N, SigmaSpatial
    Label5.Caption = "SPATIAL S " & SigmaSpatial


End Sub

Private Sub SK_PercDONE(Value As Single, CurrIteration As Long)
    ShapeProg.Width = ShapeBG.Width * Value
    LabProg = Int(Value * 100) & "%  Iteration " & CurrIteration & ""
    DoEvents
End Sub





Private Sub tLumHue_Change()
    If Val(tLUMHUE) < 0 Then tLUMHUE = 0
    If Val(tLUMHUE) > 1 Then tLUMHUE = 1

End Sub

Private Sub tMAXWH_Change()
    MaxWH = Val(tMAXWH)
End Sub

Private Sub tSigmaSpatial_Change()
    SigmaSpatial = Val(Replace(tSigmaSpatial, ",", "."))
    FX.zPreview_Spatial PicSpatial, N, SigmaSpatial

End Sub


Public Sub SaveSetting(ByVal F As String)
    Open App.Path & "\" & F For Output As 1
    Print #1, scrRAD
    Print #1, scrIntensitySigma
    Print #1, scrSpatial
    Print #1, scrITER
    Print #1, scrCONT
    Print #1, scrLUMHUE
    Print #1, cmbIntensityMode
    Print #1, IIf(oRGB, 1, 0)

    Print #1, cmbPREeffect
    Print #1, sEXPO
    Print #1, sBRIGHT
    Print #1, sCONTRA
    Print #1, sSATUR

    Print #1, scrNseg
    Print #1, scrPRES

    Print #1, Dir1

    Print #1, cmbContourMode.ListIndex


    Close 1

    F = Left$(F, Len(F) - 4) & "EX.txt"
    Open App.Path & "\" & F For Output As 1
    Print #1, ExtendSetting
    Close 1

End Sub
Public Sub LoadSetting(F As String)
    Dim S          As String
    Dim N          As Single

    Open App.Path & "\" & F For Input As 1
    Input #1, N: scrRAD.Value = N
    Input #1, N: scrIntensitySigma.Value = N
    Input #1, N: scrSpatial = N
    Input #1, N: scrITER.Value = N

    Input #1, N: scrCONT = N
    Input #1, N: scrLUMHUE = N
    Input #1, S
    cmbIntensityMode.ListIndex = Val(Left$(S, 1))
    Input #1, N
    If N = 1 Then
        oRGB = True
        CmbColorSpace.ListIndex = 0
    Else
        oRGB = False
        CmbColorSpace.ListIndex = 1
    End If

    If Not EOF(1) Then
        Input #1, S: cmbPREeffect.ListIndex = Val(Left$(S, 1))
        Input #1, N: sEXPO = N
        Input #1, N: sBRIGHT = N
        Input #1, N: sCONTRA = N
        Input #1, N: sSATUR = N
        Input #1, N: scrNseg = N
        Input #1, N: scrPRES = N
    End If

    If Not EOF(1) Then
        Input #1, S
        If Dir(S, vbDirectory) <> "" Then Drive1 = Left$(S, 2): Dir1 = S

        Input #1, N: cmbContourMode.ListIndex = N
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

    Select Case cmbPREeffect.ListIndex

        Case 0
        Case 1
            FX.MagneKleverHistogramEQU 0.3
        Case 2
            FX.MagneKleverExposure sEXPO
        Case 3
            FX.MagneKleverBCS sBRIGHT, sCONTRA, sSATUR
    End Select

End Sub



Public Function UpDateSetString() As String


    UpDateSetString = "R:" & tRad & "  I:" & tSIGMA & "  S:" & tSigmaSpatial & "  IT:" & ITER & "  C:" & tCONT & "|" & tLUMHUE & cmbContourMode & "  iMode:" & IntensityMode & "  " & IIf(oRGB, "RGB", "CieLAB") & vbCrLf
    ExtendSetting = ""


    UpDateSetString = UpDateSetString & " Seg:" & scrNseg & " %" & scrPRES

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
    ExtendSetting = ExtendSetting & vbTab & "Intensity Mode : " & cmbIntensityMode & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Intensity Sigma: " & tSIGMA & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Spatial   Sigma: " & tSigmaSpatial & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Iterations: " & ITER & vbCrLf & vbCrLf

    ExtendSetting = ExtendSetting & "CONTOUR: " & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & cmbContourMode & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Amount : " & tCONT & vbCrLf
    'If cmbContourMode.ListIndex = 0 Then
    ExtendSetting = ExtendSetting & vbTab & "Lum/(A&B): " & tLUMHUE & vbCrLf
    'Else
    '    ExtendSetting = ExtendSetting & vbTab & "Threshold: " & tLUMHUE & vbCrLf
    'End If

    ExtendSetting = ExtendSetting & vbCrLf & "LUMINANCE SEGMENTATION: " & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Segments: " & scrNseg & vbCrLf
    ExtendSetting = ExtendSetting & vbTab & "Presence: " & scrPRES & "%" & vbCrLf & vbCrLf

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
