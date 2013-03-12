VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Ed 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font Creator by Dnawrkshp & )(oDu$ - v1.0"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   2400
   ClientWidth     =   14400
   DrawStyle       =   2  'Dot
   Icon            =   "grid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   14400
   Begin VB.PictureBox ExportC 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   5400
      Picture         =   "grid.frx":08CA
      ScaleHeight     =   1530
      ScaleWidth      =   3390
      TabIndex        =   55
      Top             =   3360
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.PictureBox About 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   4100
      Picture         =   "grid.frx":117FC
      ScaleHeight     =   3795
      ScaleWidth      =   6135
      TabIndex        =   49
      Top             =   2000
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Image Up 
         Height          =   405
         Index           =   4
         Left            =   5580
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":5D5DA
         Top             =   3195
         Width           =   405
      End
      Begin VB.Shape border_about 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   3720
         Left            =   45
         Top             =   0
         Width           =   6090
      End
   End
   Begin VB.Timer Timer2 
      Left            =   1215
      Top             =   10485
   End
   Begin VB.HScrollBar FntSiz 
      Height          =   195
      Left            =   1710
      Max             =   500
      Min             =   1
      TabIndex        =   25
      Top             =   10485
      Value           =   150
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   765
      Top             =   10485
   End
   Begin VB.TextBox Ybox 
      Height          =   285
      Left            =   2610
      TabIndex        =   14
      Text            =   "Ybox"
      Top             =   10710
      Width           =   825
   End
   Begin VB.TextBox Xbox 
      Height          =   285
      Left            =   1710
      TabIndex        =   13
      Text            =   "Xbox"
      Top             =   10710
      Width           =   825
   End
   Begin VB.PictureBox FontPict 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00BCA93F&
      ForeColor       =   &H80000008&
      Height          =   7050
      Left            =   2790
      MousePointer    =   2  'Cross
      ScaleHeight     =   17.552
      ScaleMode       =   0  'User
      ScaleWidth      =   21.357
      TabIndex        =   7
      Top             =   900
      Width           =   8680
      Begin VB.Line StrightLine 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   30
         Visible         =   0   'False
         X1              =   17.152
         X2              =   15.16
         Y1              =   2.913
         Y2              =   6.274
      End
      Begin VB.Shape Rectangle 
         BorderColor     =   &H00BCA93F&
         BorderStyle     =   3  'Dot
         BorderWidth     =   30
         FillColor       =   &H00FFFFC0&
         Height          =   1095
         Left            =   1575
         Top             =   1260
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Shape Circles 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   30
         FillColor       =   &H00FFFFC0&
         Height          =   1140
         Left            =   3735
         Shape           =   3  'Circle
         Top             =   1260
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Line MinLine 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   3  'Dot
         X1              =   0.885
         X2              =   0.885
         Y1              =   0.336
         Y2              =   18.336
      End
      Begin VB.Line maxLine 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   20.029
         X2              =   20.029
         Y1              =   0.448
         Y2              =   18.448
      End
      Begin VB.Line TopGuide 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   22.001
         Y1              =   1.001
         Y2              =   1.001
      End
      Begin VB.Line BottomGuide 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   22.001
         Y1              =   16.469
         Y2              =   16.469
      End
      Begin VB.Image TOPSLIDER 
         DragMode        =   1  'Automatic
         Height          =   165
         Index           =   0
         Left            =   8415
         MousePointer    =   7  'Size N S
         Picture         =   "grid.frx":5DEF8
         Top             =   315
         Width           =   180
      End
      Begin VB.Image TOPSLIDER 
         DragMode        =   1  'Automatic
         Height          =   165
         Index           =   1
         Left            =   8415
         MousePointer    =   7  'Size N S
         Picture         =   "grid.frx":5E0C6
         Top             =   6525
         Width           =   180
      End
   End
   Begin VB.PictureBox Sliderpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1700
      Left            =   2790
      ScaleHeight     =   4.131
      ScaleMode       =   0  'User
      ScaleWidth      =   17.113
      TabIndex        =   3
      Top             =   7965
      Width           =   8680
      Begin VB.CheckBox KeyCheck 
         BackColor       =   &H00000000&
         Caption         =   "Key Shortcuts"
         ForeColor       =   &H00BCA93F&
         Height          =   255
         Left            =   5040
         TabIndex        =   54
         Top             =   1366
         Width           =   1455
      End
      Begin VB.CheckBox AutoSize 
         BackColor       =   &H00000000&
         Caption         =   "Auto Size"
         ForeColor       =   &H00BCA93F&
         Height          =   195
         Left            =   3645
         TabIndex        =   50
         Top             =   1395
         Width           =   1500
      End
      Begin VB.CheckBox AutoUpdate 
         BackColor       =   &H00000000&
         Caption         =   "Auto Update"
         ForeColor       =   &H00BCA93F&
         Height          =   195
         Left            =   315
         TabIndex        =   48
         Top             =   1395
         Width           =   1500
      End
      Begin VB.CheckBox Showbyte 
         BackColor       =   &H00000000&
         Caption         =   "Show Hex"
         ForeColor       =   &H00BCA93F&
         Height          =   195
         Left            =   2025
         TabIndex        =   47
         Top             =   1395
         Width           =   1230
      End
      Begin VB.PictureBox PVUE 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   7785
         Picture         =   "grid.frx":5E294
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   28
         Top             =   540
         Width           =   720
         Begin VB.PictureBox prevue 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   225
            ScaleHeight     =   18
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   18
            TabIndex        =   29
            Top             =   270
            Width           =   300
         End
      End
      Begin VB.PictureBox RightTick 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8010
         MousePointer    =   9  'Size W E
         Picture         =   "grid.frx":5FDD6
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   5
         Top             =   45
         Width           =   285
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00BCA93F&
         Height          =   285
         Left            =   3690
         TabIndex        =   4
         Top             =   135
         Width           =   390
      End
      Begin VB.Shape border_options 
         BorderColor     =   &H00404000&
         BorderWidth     =   2
         Height          =   300
         Left            =   225
         Shape           =   4  'Rounded Rectangle
         Top             =   1350
         Width           =   8295
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   8
         Left            =   6435
         Picture         =   "grid.frx":60020
         Tag             =   "25"
         Top             =   450
         Width           =   300
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   7
         Left            =   6030
         Picture         =   "grid.frx":60512
         Tag             =   "24"
         Top             =   720
         Width           =   300
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   6
         Left            =   6435
         Picture         =   "grid.frx":60A04
         Tag             =   "23"
         Top             =   990
         Width           =   300
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   5
         Left            =   6885
         Picture         =   "grid.frx":60EF6
         Tag             =   "22"
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label_Shift 
         BackStyle       =   0  'Transparent
         Caption         =   "SHIFT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BCA93F&
         Height          =   240
         Left            =   6390
         TabIndex        =   43
         Top             =   765
         Width           =   495
      End
      Begin VB.Image Up 
         Height          =   720
         Index           =   19
         Left            =   270
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":613E8
         Tag             =   "27"
         Top             =   540
         Width           =   720
      End
      Begin VB.Shape SizeFx 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   360
         Left            =   360
         Top             =   45
         Width           =   7800
      End
      Begin VB.Label label_charwidth 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Character width"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BCA93F&
         Height          =   240
         Left            =   0
         TabIndex        =   6
         Top             =   540
         Width           =   8175
      End
      Begin VB.Shape border_circlepan 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   570
         Left            =   6210
         Shape           =   2  'Oval
         Top             =   585
         Width           =   780
      End
   End
   Begin VB.PictureBox PicExporting 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   810
      Picture         =   "grid.frx":62F2A
      ScaleHeight     =   328.571
      ScaleMode       =   0  'User
      ScaleWidth      =   460
      TabIndex        =   1
      Top             =   10935
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   225
      Top             =   10440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox CharList1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BCA93F&
      Height          =   7455
      ItemData        =   "grid.frx":BE844
      Left            =   270
      List            =   "grid.frx":BE846
      TabIndex        =   24
      Top             =   855
      Width           =   2040
   End
   Begin VB.PictureBox UndoPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   7170
      Left            =   11700
      ScaleHeight     =   7140
      ScaleWidth      =   2460
      TabIndex        =   30
      Top             =   765
      Visible         =   0   'False
      Width           =   2490
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   11
         Left            =   1260
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   42
         Top             =   5940
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   10
         Left            =   90
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   41
         Top             =   5940
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   9
         Left            =   1260
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   40
         Top             =   4770
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   8
         Left            =   90
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   39
         Top             =   4770
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   7
         Left            =   1260
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   38
         Top             =   3600
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   6
         Left            =   90
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   37
         Top             =   3600
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   5
         Left            =   1260
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   36
         Top             =   2430
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   4
         Left            =   90
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   35
         Top             =   2430
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   0
         Left            =   90
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   34
         Top             =   90
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   1
         Left            =   1260
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   33
         Top             =   90
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   2
         Left            =   90
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   32
         Top             =   1260
         Width           =   1100
      End
      Begin VB.PictureBox undo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1100
         Index           =   3
         Left            =   1260
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   31
         Top             =   1260
         Width           =   1100
      End
   End
   Begin VB.PictureBox Toolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   14370
      TabIndex        =   45
      Top             =   0
      Width           =   14400
      Begin VB.Image Up 
         Height          =   405
         Index           =   27
         Left            =   12420
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":BE848
         Tag             =   "28"
         Top             =   90
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   30
         Left            =   16290
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":BF166
         Tag             =   "17"
         Top             =   45
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   29
         Left            =   12870
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":BFA84
         Tag             =   "21"
         Top             =   90
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   28
         Left            =   13320
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C03A2
         Tag             =   "20"
         Top             =   90
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   26
         Left            =   13770
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C0CC0
         Tag             =   "18"
         Top             =   90
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   25
         Left            =   2475
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C15DE
         Tag             =   "13"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   24
         Left            =   225
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C1EFC
         Tag             =   "17"
         Top             =   90
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   23
         Left            =   2025
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C281A
         Tag             =   "26"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   22
         Left            =   1575
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C3138
         Tag             =   "14"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   21
         Left            =   1125
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C3A56
         Tag             =   "15"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   20
         Left            =   675
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C4374
         Tag             =   "16"
         Top             =   105
         Width           =   405
      End
      Begin VB.Shape border_toolbar 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   520
         Left            =   90
         Top             =   45
         Width           =   14190
      End
      Begin VB.Image ClearIt 
         Height          =   390
         Index           =   0
         Left            =   11205
         Picture         =   "grid.frx":C4C92
         Top             =   90
         Width           =   405
      End
      Begin VB.Image ClearIt 
         Height          =   390
         Index           =   2
         Left            =   10755
         Picture         =   "grid.frx":C555C
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   15
         Left            =   6840
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C5E26
         Tag             =   "12"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   14
         Left            =   7290
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C6744
         Tag             =   "11"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   17
         Left            =   8190
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C7062
         Tag             =   "9"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   16
         Left            =   7740
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C7980
         Tag             =   "10"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image Up 
         Height          =   405
         Index           =   18
         Left            =   8640
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":C829E
         Tag             =   "19"
         Top             =   105
         Width           =   405
      End
      Begin VB.Image opt 
         Height          =   390
         Index           =   3
         Left            =   4410
         Picture         =   "grid.frx":C8BBC
         Top             =   105
         Width           =   405
      End
      Begin VB.Image opt 
         Height          =   390
         Index           =   4
         Left            =   4860
         Picture         =   "grid.frx":C9486
         Top             =   105
         Width           =   405
      End
      Begin VB.Image opt 
         Height          =   390
         Index           =   5
         Left            =   5310
         Picture         =   "grid.frx":C9D50
         Top             =   105
         Width           =   405
      End
      Begin VB.Image opt 
         Height          =   390
         Index           =   6
         Left            =   5760
         Picture         =   "grid.frx":CA61A
         Top             =   105
         Width           =   405
      End
      Begin VB.Image opt 
         Height          =   390
         Index           =   0
         Left            =   3060
         Picture         =   "grid.frx":CAEE4
         Top             =   105
         Width           =   405
      End
      Begin VB.Image opt 
         Height          =   390
         Index           =   1
         Left            =   3510
         Picture         =   "grid.frx":CB7AE
         Top             =   105
         Width           =   405
      End
      Begin VB.Image opt 
         Height          =   390
         Index           =   2
         Left            =   3960
         Picture         =   "grid.frx":CC078
         Top             =   105
         Width           =   405
      End
   End
   Begin VB.PictureBox Hexbox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   11700
      ScaleHeight     =   9375
      ScaleWidth      =   2535
      TabIndex        =   51
      Top             =   730
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox Hextext 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BCA93F&
         Height          =   8430
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   53
         Top             =   225
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "HEX View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BCA93F&
         Height          =   195
         Left            =   720
         TabIndex        =   52
         Top             =   0
         Width           =   855
      End
      Begin VB.Shape border_hexbox 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   9255
         Left            =   45
         Top             =   75
         Width           =   2400
      End
   End
   Begin VB.PictureBox FontFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9350
      Left            =   11700
      ScaleHeight     =   9345
      ScaleWidth      =   2490
      TabIndex        =   9
      Top             =   765
      Width           =   2490
      Begin PS2FontCreator.UserControl1 fscale 
         Height          =   300
         Left            =   225
         TabIndex        =   46
         Top             =   4635
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
      End
      Begin VB.PictureBox IMP_PROG 
         BackColor       =   &H00000000&
         Height          =   150
         Left            =   135
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   255
         TabIndex        =   44
         Top             =   8685
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.ListBox winfontlist 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BCA93F&
         Height          =   2760
         ItemData        =   "grid.frx":CC942
         Left            =   240
         List            =   "grid.frx":CC949
         MouseIcon       =   "grid.frx":CC95A
         MousePointer    =   14  'Arrow and Question
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   300
         Width           =   1995
      End
      Begin VB.PictureBox SmallFontpic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00BCA93F&
         Height          =   1485
         Left            =   405
         ScaleHeight     =   18
         ScaleMode       =   0  'User
         ScaleWidth      =   18
         TabIndex        =   10
         Top             =   5355
         Width           =   1665
      End
      Begin VB.TextBox FontChar 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BCA93F&
         Height          =   960
         Left            =   225
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3375
         Visible         =   0   'False
         Width           =   1275
      End
      Begin PS2FontCreator.UserControl2 DitherPick 
         Height          =   300
         Left            =   180
         TabIndex        =   26
         Top             =   7380
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
      End
      Begin VB.Shape border_dither 
         BorderColor     =   &H00404000&
         BorderWidth     =   2
         Height          =   585
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   7200
         Width           =   2220
      End
      Begin VB.Label Dith 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dither Tolerance 30%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   225
         TabIndex        =   27
         Top             =   7200
         Width           =   2010
      End
      Begin VB.Shape border_winfontlist 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   3045
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   225
         Width           =   2220
      End
      Begin VB.Label fontsiz 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "150"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   1440
         TabIndex        =   21
         Top             =   4455
         Width           =   435
      End
      Begin VB.Image Up 
         Height          =   720
         Index           =   13
         Left            =   1635
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":CD224
         Top             =   7875
         Width           =   720
      End
      Begin VB.Image Up 
         Height          =   720
         Index           =   12
         Left            =   885
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":CED66
         Top             =   7875
         Width           =   720
      End
      Begin VB.Image Up 
         Height          =   720
         Index           =   11
         Left            =   135
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":D08A8
         Top             =   7875
         Width           =   720
      End
      Begin VB.Label Label_ascii 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "asc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1890
         TabIndex        =   20
         Top             =   3375
         Width           =   315
      End
      Begin VB.Label Label_hex 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "hex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1890
         TabIndex        =   19
         Top             =   3915
         Width           =   315
      End
      Begin VB.Label hexchar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "hex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   1890
         TabIndex        =   18
         Top             =   4095
         Width           =   315
      End
      Begin VB.Label ascchar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "asc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   1890
         TabIndex        =   17
         Top             =   3600
         Width           =   315
      End
      Begin VB.Label Fontbox 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BCA93F&
         Height          =   1005
         Left            =   225
         TabIndex        =   15
         Top             =   3330
         Width           =   1290
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   10
         Left            =   1530
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":D23EA
         Tag             =   "23"
         Top             =   4005
         Width           =   300
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   9
         Left            =   1530
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":D28DC
         Tag             =   "25"
         Top             =   3420
         Width           =   300
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   3
         Left            =   2070
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":D2DCE
         Tag             =   "22"
         Top             =   5985
         Width           =   300
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   2
         Left            =   1080
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":D32C0
         Tag             =   "23"
         Top             =   6840
         Width           =   300
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   1
         Left            =   90
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":D37B2
         Tag             =   "24"
         Top             =   5985
         Width           =   300
      End
      Begin VB.Image Up 
         Height          =   300
         Index           =   0
         Left            =   1080
         MousePointer    =   99  'Custom
         Picture         =   "grid.frx":D3CA4
         Tag             =   "25"
         Top             =   5040
         Width           =   300
      End
      Begin VB.Shape border_winfontcharacter 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   1065
         Index           =   8
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   3330
         Width           =   2220
      End
      Begin VB.Label Label_importwinfont 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Import Win Font"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BCA93F&
         Height          =   195
         Left            =   540
         TabIndex        =   11
         Top             =   -30
         Width           =   1365
      End
      Begin VB.Label Label_scale 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Scale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   540
         TabIndex        =   12
         Top             =   4455
         Width           =   480
      End
      Begin VB.Shape border_fontpan 
         BorderColor     =   &H00404000&
         BorderWidth     =   2
         Height          =   1830
         Left            =   225
         Shape           =   4  'Rounded Rectangle
         Top             =   5175
         Width           =   2040
      End
      Begin VB.Shape border_winfontsize 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   540
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   4455
         Width           =   2220
      End
      Begin VB.Shape border_importwinfont 
         BorderColor     =   &H00BCA93F&
         BorderWidth     =   2
         Height          =   9255
         Left            =   45
         Top             =   45
         Width           =   2400
      End
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   28
      Left            =   16065
      Picture         =   "grid.frx":D4196
      Top             =   9810
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   28
      Left            =   15615
      Picture         =   "grid.frx":D4AB4
      Top             =   9810
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   720
      Index           =   27
      Left            =   15075
      Picture         =   "grid.frx":D53D2
      Top             =   9000
      Width           =   720
   End
   Begin VB.Image tool_down 
      Height          =   720
      Index           =   27
      Left            =   15840
      Picture         =   "grid.frx":D6F14
      Top             =   9000
      Width           =   720
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   26
      Left            =   15660
      Picture         =   "grid.frx":D8A56
      Top             =   2520
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   26
      Left            =   16110
      Picture         =   "grid.frx":D9374
      Top             =   2520
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   300
      Index           =   25
      Left            =   15660
      Picture         =   "grid.frx":D9C92
      Top             =   7155
      Width           =   300
   End
   Begin VB.Image tool_down 
      Height          =   300
      Index           =   25
      Left            =   16110
      Picture         =   "grid.frx":DA184
      Top             =   7155
      Width           =   300
   End
   Begin VB.Image tool_still 
      Height          =   300
      Index           =   24
      Left            =   15660
      Picture         =   "grid.frx":DA676
      Top             =   7605
      Width           =   300
   End
   Begin VB.Image tool_down 
      Height          =   300
      Index           =   24
      Left            =   16110
      Picture         =   "grid.frx":DAB68
      Top             =   7605
      Width           =   300
   End
   Begin VB.Image tool_still 
      Height          =   300
      Index           =   23
      Left            =   15660
      Picture         =   "grid.frx":DB05A
      Top             =   8055
      Width           =   300
   End
   Begin VB.Image tool_down 
      Height          =   300
      Index           =   23
      Left            =   16110
      Picture         =   "grid.frx":DB54C
      Top             =   8055
      Width           =   300
   End
   Begin VB.Image tool_still 
      Height          =   300
      Index           =   22
      Left            =   15660
      Picture         =   "grid.frx":DBA3E
      Top             =   8505
      Width           =   300
   End
   Begin VB.Image tool_down 
      Height          =   300
      Index           =   22
      Left            =   16110
      Picture         =   "grid.frx":DBF30
      Top             =   8505
      Width           =   300
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   21
      Left            =   15660
      Picture         =   "grid.frx":DC422
      Top             =   5220
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   21
      Left            =   16110
      Picture         =   "grid.frx":DCD40
      Top             =   5220
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   20
      Left            =   15660
      Picture         =   "grid.frx":DD65E
      Top             =   5670
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   20
      Left            =   16110
      Picture         =   "grid.frx":DDF7C
      Top             =   5670
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   19
      Left            =   15660
      Picture         =   "grid.frx":DE89A
      Top             =   6120
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   19
      Left            =   16110
      Picture         =   "grid.frx":DF1B8
      Top             =   6120
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   18
      Left            =   15660
      Picture         =   "grid.frx":DFAD6
      Top             =   6570
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   18
      Left            =   16110
      Picture         =   "grid.frx":E03F4
      Top             =   6570
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   17
      Left            =   15660
      Picture         =   "grid.frx":E0D12
      Top             =   720
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   17
      Left            =   16110
      Picture         =   "grid.frx":E1630
      Top             =   720
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   16
      Left            =   15660
      Picture         =   "grid.frx":E1F4E
      Top             =   1170
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   16
      Left            =   16110
      Picture         =   "grid.frx":E286C
      Top             =   1170
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   15
      Left            =   15660
      Picture         =   "grid.frx":E318A
      Top             =   1620
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   15
      Left            =   16110
      Picture         =   "grid.frx":E3AA8
      Top             =   1620
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   14
      Left            =   15660
      Picture         =   "grid.frx":E43C6
      Top             =   2070
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   14
      Left            =   16110
      Picture         =   "grid.frx":E4CE4
      Top             =   2070
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   13
      Left            =   15660
      Picture         =   "grid.frx":E5602
      Top             =   2970
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   13
      Left            =   16110
      Picture         =   "grid.frx":E5F20
      Top             =   2970
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   12
      Left            =   15660
      Picture         =   "grid.frx":E683E
      Top             =   3420
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   12
      Left            =   16110
      Picture         =   "grid.frx":E715C
      Top             =   3420
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   11
      Left            =   15660
      Picture         =   "grid.frx":E7A7A
      Top             =   3870
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   11
      Left            =   16110
      Picture         =   "grid.frx":E8398
      Top             =   3870
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   10
      Left            =   15660
      Picture         =   "grid.frx":E8CB6
      Top             =   4320
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   10
      Left            =   16110
      Picture         =   "grid.frx":E95D4
      Top             =   4320
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   405
      Index           =   9
      Left            =   15660
      Picture         =   "grid.frx":E9EF2
      Top             =   4770
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   405
      Index           =   9
      Left            =   16110
      Picture         =   "grid.frx":EA810
      Top             =   4770
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   8
      Left            =   15165
      Picture         =   "grid.frx":EB12E
      Top             =   4320
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   8
      Left            =   14715
      Picture         =   "grid.frx":EB9F8
      Top             =   4320
      Width           =   405
   End
   Begin VB.Label Label_characterlist 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Character List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BCA93F&
      Height          =   195
      Left            =   585
      TabIndex        =   23
      Top             =   630
      Width           =   1200
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   7
      Left            =   15165
      Picture         =   "grid.frx":EC2C2
      Top             =   3870
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   7
      Left            =   14715
      Picture         =   "grid.frx":ECB8C
      Top             =   3870
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   6
      Left            =   15165
      Picture         =   "grid.frx":ED456
      Top             =   3420
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   6
      Left            =   14715
      Picture         =   "grid.frx":EDD20
      Top             =   3420
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   5
      Left            =   15165
      Picture         =   "grid.frx":EE5EA
      Top             =   2970
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   5
      Left            =   14715
      Picture         =   "grid.frx":EEEB4
      Top             =   2970
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   4
      Left            =   15165
      Picture         =   "grid.frx":EF77E
      Top             =   2520
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   4
      Left            =   14715
      Picture         =   "grid.frx":F0048
      Top             =   2520
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   3
      Left            =   15165
      Picture         =   "grid.frx":F0912
      Top             =   2070
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   3
      Left            =   14715
      Picture         =   "grid.frx":F11DC
      Top             =   2070
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   2
      Left            =   15165
      Picture         =   "grid.frx":F1AA6
      Top             =   1620
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   2
      Left            =   14715
      Picture         =   "grid.frx":F2370
      Top             =   1620
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   1
      Left            =   15165
      Picture         =   "grid.frx":F2C3A
      Top             =   1170
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   1
      Left            =   14715
      Picture         =   "grid.frx":F3504
      Top             =   1170
      Width           =   405
   End
   Begin VB.Image tool_down 
      Height          =   390
      Index           =   0
      Left            =   15165
      Picture         =   "grid.frx":F3DCE
      Top             =   720
      Width           =   405
   End
   Begin VB.Image tool_still 
      Height          =   390
      Index           =   0
      Left            =   14715
      Picture         =   "grid.frx":F4698
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label_MainEditor 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Character Editor (18x18)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BCA93F&
      Height          =   195
      Left            =   2925
      TabIndex        =   2
      Top             =   585
      Width           =   2085
   End
   Begin VB.Shape border_helptext 
      BorderColor     =   &H00BCA93F&
      BorderWidth     =   2
      Height          =   300
      Left            =   2700
      Top             =   9765
      Width           =   8925
   End
   Begin VB.Label HelpText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2745
      TabIndex        =   8
      Top             =   9765
      Width           =   8820
   End
   Begin VB.Shape border_editor 
      BorderColor     =   &H00BCA93F&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   8895
      Left            =   2700
      Top             =   810
      Width           =   8925
   End
   Begin VB.Image LOGO 
      Height          =   1650
      Left            =   90
      Picture         =   "grid.frx":F4F62
      Top             =   8505
      Width           =   2460
   End
   Begin VB.Label SPaceMarker 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   10
      Left            =   2.45745e5
      TabIndex        =   0
      Top             =   6000
      Width           =   330
   End
   Begin VB.Shape border_characterlist 
      BorderColor     =   &H00BCA93F&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   7725
      Left            =   90
      Top             =   720
      Width           =   2445
   End
   Begin VB.Shape border_charcterEditor 
      BorderColor     =   &H00BCA93F&
      BorderWidth     =   2
      FillColor       =   &H00BCA93F&
      Height          =   9435
      Left            =   2610
      Top             =   720
      Width           =   11670
   End
End
Attribute VB_Name = "Ed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _______________________________________________________________________________________________________________
'|#######################################[        DECLARATIONS       ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Option Explicit
Public StartX
Public StartY
Public drawtype
Public drawing
Public CurX
Public CurY
Const darkgrey = &H202020
Public Import As Boolean
Public isdoing
Public Undos

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim A
    If KeyShort = False Then: GoTo needed
    Select Case KeyCode
        Case 70: opt_Click 6: opt_MouseDown 6, 1, 1, 1, 1                                               '[F]ONT TOOL
        Case 68: opt_Click 5: opt_MouseDown 5, 1, 1, 1, 1                                               '[D]RAW TOOL
        Case 80: opt_Click 4: opt_MouseDown 4, 1, 1, 1, 1                                               '[P]EN TOOL
        Case 66: opt_Click 3: opt_MouseDown 3, 1, 1, 1, 1                                               'FILLED [B]OX
        Case 67: opt_Click 2: opt_MouseDown 2, 1, 1, 1, 1                                               '[C]IRCLE
        Case 83: opt_Click 1: opt_MouseDown 1, 1, 1, 1, 1                                               '[S]QUARE
        Case 76: opt_Click 0: opt_MouseDown 0, 1, 1, 1, 1                                               '[L]INE
        Case 8: ClearIt_MouseUp 2, 1, 1, 1, 1                                                           '[BKSP] UNDO
        Case 46: ClearIt_Click 0                                                                        '[DEL] CLEAR EDITOR
        Case 79                                                                                         '[O]PEN FONT FILE
        Case 85: SaveCharacter                                                                          '[U]PDATE
        'FONT PANEL PAN BUTTONS
        Case 104: Ybox.text = Str(Val(Ybox.text) - 1):  FontPict.CurrentY = Val(Ybox.text): ShowFont    '[10KEY-UP]
        Case 100: Xbox.text = Str(Val(Xbox.text) - 1):  FontPict.CurrentX = Val(Xbox.text): ShowFont    '[10KEY-left]
        Case 98: Ybox.text = Str(Val(Ybox.text) + 1):  FontPict.CurrentY = Val(Ybox.text): ShowFont     '[10KEY-down]
        Case 102: Xbox.text = Str(Val(Xbox.text) + 1):  FontPict.CurrentX = Val(Xbox.text): ShowFont    '[10KEY-right]
        Case 65: About.Visible = True   'about info                                                     '[A]bout box
        'main editor pan buttons
        Case 35: MoveRight                                                                              '[RIGHT ARROW]
        Case 36: MoveDown                                                                               '[DOWN ARROW]
        Case 37: MoveLeft                                                                               '[LEFT ARROW]
        Case 38: MoveUp                                                                                 '[UP ARROW]
        'font panel Previous/Next Letter
        Case 105: A = Asc(FontBox.Caption): A = A - 1: If A = 0 Then A = 1: FontBox.Caption = Chr$(A): ShowFont    '[10KEY-PGUP]
        Case 99: A = Asc(FontBox.Caption): A = A + 1: If A = 256 Then A = 255: FontBox.Caption = Chr$(A): ShowFont '[10KEY-PGDN]
        
        'font panel functions
        Case 77: Merge                                                                                  '[M]ERGE
        Case 73: ImportFont                                                                             '[I]MPORT FONT
        Case 13: FontBox.Caption = Right(CharList1.List(CharList1.ListIndex), 1)
        
        'functions - toolbar
        Case 14: CopyImage
        Case 15: PasteImage
        Case 16: AllSpace '[LEFT SHIFT]
        Case 32: SpaceOne
        Case 18: BlankFill
        'Case 19:
        
        'right hand toolbar
        
        Case 26: HelpMenu               'help file
        Case 27: Controls_Click
        Case 28:
        Case 29: MsgBox "Not yet implemented!" 'SettingsForm.Show 1
        
    End Select
    
needed:
    If KeyCode = vbKeyS And Shift = 2 Then
        QuickSaveFont
    End If
    Dim X
    X = 1
End Sub

' _______________________________________________________________________________________________________________
'|#######################################[        FORM EVENTS        ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Form_Load()
    Dim CmdPath As String
    'check if pfc extension is already associated, if not, link it to Font creator
    Dim ret
    ret = CheckFileAssociation("pfc")
    If ret = "" Or ret = App.EXEName Or ret = "PS2 Font Creator.exe" Then
        MakeFileAssociation "pfc", App.Path, App.EXEName, "PS2 Font Creator Font File" 'App.Path & "Font Creator Icon.ico"
    End If
    
    'If Right(Command, 3) = "pfc" Then: OPEN_FONT Command
    CmdPath = Replace(Command, Chr$(34), "")
    If CmdPath <> "" Then
        If ExtOf(CmdPath) = ".PFC" Then
            GlobalFileName = CmdPath
            OPEN_FONT CmdPath
        Else
            GlobalFileName = CmdPath
            GlobalSpaceName = Replace(GlobalFileName, ".", " SPACE.", , 1)
            If Dir(GlobalSpaceName) = "" Then: MakeDumFile GlobalSpaceName, 255, 1
            CharList1.ListIndex = 0
            Dim OFS As Long
            Dim GRABDATA As Integer
            For GRABDATA = 0 To 255
                OFS = ((GRABDATA * 32.4) * 10) + 1
                fontz.fdata(GRABDATA).ByteData = AccessBinSegment(OFS, 324)
                fontz.fdata(GRABDATA).spacing = AccessSpaceSegment(GRABDATA + 1)
            Next GRABDATA
            CharList1.ListIndex = 0
            MsgBox "Finished Importing data"
        End If
    Else
        GlobalFileName = AppDirectory & "temp.pfc"
        If Dir(GlobalFileName) <> "" Then: Kill GlobalFileName
        MakeDumFilePFC GlobalFileName
    End If
    
    ControlRollFade Toolbar, 0, 200, 255
    ControlFade FontFrame, 0, 200, 255
    CurrentLetter = 0
    SetLengths
    dither = 30
    CharList1.AddItem "0x00   (NULL)", 0
    CharList1.ListIndex = 0
    AddFonts
    InitGrids
    winfontlist.ListIndex = 0
    AppDirectory = App.Path & "\"
    'MakeDumFile GlobalFileName, 82944
    'MakeDumFile GlobalSpaceName, 256, 1
    PopulateList
    Drawgrid
    DrawMiniGrid
    drawtype = 6
    ascchar = AscW(FontBox.Caption)
    hexchar = Hex(ascchar)
    'SaveAsPFC AppDirectory & "Test.PFC", "FontName", "FontArray", "SpaceArray", HexToString("FFFFFFFF"), HexToString("11111111")
    'SaveSegPFC AppDirectory & "Test.PFC", HexToString("EEEEEEEE"), F_Data
    
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Form_KeyPress / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:46
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 3 Then CopyImage
    If KeyAscii = 22 Then PasteImage
    
    
    
    
    
    
End Sub

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Form_Terminate / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:46
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Form_Terminate()
    If Dir(AppDirectory & "temp.bin") <> "" Then: Kill AppDirectory & "temp.bin"
    If Dir(AppDirectory & "temp SPACE.bin") <> "" Then: Kill AppDirectory & "temp SPACE.bin"
    If Dir(AppDirectory & "temp.pfc") <> "" Then: Kill AppDirectory & "temp.pfc"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Form_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:46
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = ""
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Form_Resize / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:46
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Form_Resize()
    FormFade Me, 0, 200, 255
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : InitGrids / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:46
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Sub InitGrids()
    'FontPict.DrawWidth = 225
    FontBox.FontSize = 50
    CD.FontSize = 50
    FontPict.ScaleHeight = 18
    FontPict.ScaleWidth = 20
    FontPict.FontSize = fontsiz.FontSize
    SmallFontpic.FontSize = fontsiz.FontSize
    SmallFontpic.ScaleHeight = 18
    SmallFontpic.ScaleWidth = 18
    ControlFade FontFrame, 0, 200, 255
    Toolbar.ScaleHeight = 255
    ControlFade UndoPanel, 0, 200, 255
    min = 0
    max = 18
End Sub

' _______________________________________________________________________________________________________________
'|#######################################[    MAIN EDITOR FUNCTIONS  ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : fontPict_MouseDown / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:32
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub fontPict_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 18 Then Exit Sub
    X = Int(X) + 0.5: Y = Int(Y) + 0.5:    StartX = X:    StartY = Y:    drawing = True
    Select Case drawtype
        Case 0 ' line
           ' FontPict.Line (CurX, CurY)-(CurX + 1, CurY + 1), drawcolor, BF
        Case 1, 3
            Rectangle.Move CurX + 0.5, CurY + 0.5, 0, 0
    End Select
    Drawgrid
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : fontPict_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:31
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub fontPict_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartX = Int(StartX)
    StartY = Int(StartY)
    CurX = Int(CurX)
    CurY = Int(CurY)
    X = Int(X)
    Y = Int(Y)

    Select Case drawtype
        Case 0: HelpText = "Left click then drag to draw line"
        Case 1: HelpText = "Left click then drag to draw box"
        Case 2: HelpText = "Left click then drag to draw circle"
        Case 3: HelpText = "Left click then drag to draw filled box"
        Case 4: HelpText = "Left click /hold to free draw"
        Case 5: HelpText = "Left click /drag to line draw"
    End Select

    If drawing = False Then
        Rectangle.Visible = False
        Circles.Visible = False
        StrightLine.Visible = False
    End If


If drawing = True Then
    If CurX <> X Or CurY <> Y Then
    Select Case drawtype
        Case 0 ' line
            StrightLine.x1 = StartX + 0.5
            StrightLine.Y1 = StartY + 0.5
            StrightLine.x2 = X + 0.5
            StrightLine.Y2 = Y + 0.5
            StrightLine.Visible = True
            FontPict.MousePointer = 2
        Case 1 ' box
            
            'Right Quadrants
            If CurX > StartX Then
                'Upper Right
                If StartY > CurY Then
                    Rectangle.Move StartX + 0.5, CurY + 0.5, Abs(X - StartX), Abs(StartY - CurY)
                    FontPict.MousePointer = 6
                End If
                'Down Right
                If CurY > StartY Then
                    Rectangle.Move StartX + 0.5, StartY + 0.5, Abs(X - StartX), Abs(Y - StartY)
                    FontPict.MousePointer = 8
                End If
            End If
            
            'Left Quadrants
            If StartX > CurX Then 'Upper Left
                If StartY > CurY Then
                    Rectangle.Move CurX + 0.5, CurY + 0.5, Abs(StartX - CurX), Abs(StartY - CurY)
                    FontPict.MousePointer = 8
                End If
                If CurY > StartY Then 'Lower Left
                    Rectangle.Move X + 0.5, StartY + 0.5, Abs(StartX - X), Abs(CurY - StartY)
                    FontPict.MousePointer = 6
                    
                End If
            End If
            
            Rectangle.FillStyle = 1
            Rectangle.Visible = True
        
        Case 2 ' circle
            Circles.Move StartX + 0.5, StartY + 0.5, Abs(X - StartX), Abs(Y - StartY)
            Circles.Visible = True
            FontPict.MousePointer = 2
            
        Case 3 ' filled box
            'Right Quadrants
            If CurX > StartX Then
                If StartY > CurY Then 'Upper Right
                    Rectangle.Move StartX + 0.5, CurY + 0.5, Abs(X - StartX), Abs(StartY - CurY)
                    FontPict.MousePointer = 6
                End If
                If CurY > StartY Then 'Down Right
                    Rectangle.Move StartX + 0.5, StartY + 0.5, Abs(X - StartX), Abs(Y - StartY)
                    FontPict.MousePointer = 8
                End If
                
            End If
            
            'Left Quadrants
            If StartX > CurX Then
                If StartY > CurY Then
                    Rectangle.Move CurX + 0.5, CurY + 0.5, Abs(StartX - CurX), Abs(StartY - CurY)
                    FontPict.MousePointer = 8
                End If
                If CurY > StartY Then
                    Rectangle.Move CurX + 0.5, StartY + 0.5, Abs(StartX - CurX), Abs(CurY - StartY)
                    FontPict.MousePointer = 6
                End If
            End If
            
            Rectangle.FillStyle = 0
            Rectangle.Visible = True
        Case 4 ' Draw
            Drawgrid
            FontPict.DrawWidth = 1
            Select Case Button
                Case 1
                    If StartX + 0.5 < 18 And StartX + 0.5 >= 0 And StartY + 0.5 <= 18 And StartY + 0.5 >= 0 Then: FontPict.Line (Int(X), Int(Y))-(Int(X) + 1, Int(Y) + 1), White, BF
                Case 2
                    If StartX + 0.5 < 18 And StartX + 0.5 >= 0 And StartY + 0.5 <= 18 And StartY + 0.5 >= 0 Then: FontPict.Line (Int(X), Int(Y))-(Int(X) + 1, Int(Y) + 1), vbBlack, BF
            End Select
            FontPict.MousePointer = 2
        Case 5 ' Line draw
            Drawgrid
            Select Case Button
                Case 1: FontPict.Line (CurX + 0.5, CurY + 0.5)-(X + 0.5, Y + 0.5), White
                Case 2: FontPict.Line (CurX + 0.5, CurY + 0.5)-(X + 0.5, Y + 0.5), vbBlack
            End Select
            FontPict.MousePointer = 2
    End Select
    End If
 End If

CurX = Int(X)
CurY = Int(Y)

End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : fontPict_MouseUp / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:31
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub fontPict_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If drawing = True Then
    If CurX <> X Or CurY <> Y Then
    Dim drawcolor As Long
    Select Case Button
        Case 1: drawcolor = White
        Case 2: drawcolor = Black
    End Select
    
    Select Case drawtype
        Case 0 ' line
            StrightLine.Visible = False
            If StartX = CurX And StartY <> CurY Then
                If CurY > StartY Then FontPict.Line (StartX + 0.5, StartY)-(CurX + 0.5, CurY + 1), drawcolor, BF
                If StartY > CurY Then FontPict.Line (CurX + 0.5, CurY)-(StartX + 0.5, StartY + 1), drawcolor, BF
                
                GoTo outc
            End If
            If StartX <> CurX And StartY = CurY Then
                If CurX > StartX Then FontPict.Line (StartX, StartY + 0.5)-(CurX + 1, CurY + 0.5), drawcolor, BF
                If StartX > CurX Then FontPict.Line (CurX, CurY + 0.5)-(StartX + 1, StartY + 0.5), drawcolor, BF
                
                GoTo outc
            End If
            FontPict.DrawWidth = 24
            FontPict.Line (StartX + 0.5, StartY + 0.5)-(CurX + 0.5, CurY + 0.5), drawcolor
            
outc:
        Case 1 ' box
            'FontPict.DrawStyle = Rectangle.BorderStyle - 1
            StartY = Rectangle.Top
            StartX = Rectangle.Left
            CurY = Rectangle.Top + Rectangle.Height
            CurX = Rectangle.Left + Rectangle.Width
            FontPict.Line (StartX, StartY)-(CurX, CurY), drawcolor, B
            Rectangle.Visible = False
            
        Case 2 ' circle
            Dim acr
            Dim Dwn
            Dim cir
            acr = Circles.Left + (Circles.Width / 2)
            Dwn = Circles.Top + (Circles.Height / 2)
            cir = Int(Circles.Width / 2)
            FontPict.DrawWidth = 26
            FontPict.Circle (acr, Dwn), cir, drawcolor
            Circles.Visible = False
            
        Case 3 ' filled box
            StartY = Rectangle.Top
            StartX = Rectangle.Left
            CurY = Rectangle.Top + Rectangle.Height
            CurX = Rectangle.Left + Rectangle.Width
            FontPict.Line (StartX, StartY)-(CurX, CurY), drawcolor, BF
            Rectangle.Visible = False
        
        Case 4 ' Draw
            FontPict.DrawWidth = 2
            If StartX + 0.5 < 18 And StartX + 0.5 >= 0 And StartY + 0.5 <= 18 And StartY + 0.5 >= 0 Then: FontPict.Line (Int(X), Int(Y))-(Int(X) + 1, Int(Y) + 1), drawcolor, BF
            
        Case 5 ' Line draw
            FontPict.Line (CurX + 0.5, CurY + 0.5)-(X + 0.5, Y + 0.5), drawcolor
            
    End Select
    End If
 End If
    Rectangle.Visible = False
    Circles.Visible = False
    StrightLine.Visible = False
    Pixelate FontPict
    Drawgrid
    drawing = False
    FontPict.MousePointer = 1
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : FontPict_DragOver / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:32
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub FontPict_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

        
    Select Case Source.Name
        Case "TOPSLIDER"
            Select Case Source.Index
                Case 0: TopGuide.Y1 = Int(Y)
                        TopGuide.Y2 = Int(Y)
                        Source.Top = Int(Y) - (Source.Height / 2)
                        Source.Left = 20 - (Source.Width * 2)
                        DrawMiniGrid
                Case 1: BottomGuide.Y1 = Int(Y)
                        BottomGuide.Y2 = Int(Y)
                        Source.Top = Int(Y) - (Source.Height / 2)
                        Source.Left = 20 - (Source.Width * 2)
                        DrawMiniGrid
            End Select
    End Select
End Sub

' _______________________________________________________________________________________________________________
'|#######################################[ ACTION BUTTONS -  TOOLBAR ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' _______________________________________________________________________________________________________________
' ---------------------------------------------------------------------------------------------------------------
' Procedure : BlankFill / grid.frm
' Author    : Originally written by Dnawrkshp / Re-Written by Xodus for Main Editor mode
' Date      : 2/4/2013 09:57
' Purpose   :
' ---------------------------------------------------------------------------------------------------------------
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub BlankFill()
    Dim OldIndex As Integer
    OldIndex = CharList1.ListIndex
    Dim q

    For q = 0 To 255
        fontz.fdata(q).ByteData = String(324, "FF")
        fontz.fdata(q).spacing = 16
    Next q
    CharList1.ListIndex = OldIndex
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : ClearIt_Click / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 09:57
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub ClearIt_Click(Index As Integer)
        Select Case Index
            Case 0
                Gui_ClearEditor
                Gui_ClearPreview
            Case 2
                UndoPanel.Visible = True
                FontFrame.Visible = False
                Hexbox.Visible = False
                
        End Select
        prevue.Cls
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : ClearIt_MouseDown / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 09:57
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub ClearIt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Select Case Index
            Case 0, 1: ClearIt(Index).Picture = tool_down(7).Picture
            Case 2: ClearIt(Index).Picture = tool_down(8).Picture
        End Select
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : ClearIt_MouseUp / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 09:57
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub ClearIt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Select Case Index
            Case 0, 1: ClearIt(Index).Picture = tool_still(7).Picture
            Case 2: ClearIt(Index).Picture = tool_still(8).Picture
                 Dim x1, Y1
                 FontPict.Cls
                 Undos = Undos - 1
                 If Undos = -1 Then Undos = 11
                 If Undos = -2 Then Undos = 10
                 FontPict.DrawWidth = 1
                 For x1 = 1 To 17
                    For Y1 = 1 To 17
                        If undo(Undos).Point(X - 0.5, Y - 0.5) = White Then
                            FontPict.Line (x1 - 1, Y1 - 1)-(x1, Y1), White, BF
                        End If
                    Next Y1
                Next x1
                Drawgrid
        End Select
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : CopyImage / grid.frm
' Author    : Dnawrkshp
' Date      : 2/4/2013 09:57
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub CopyImage()

Dim Output As String
Output = ""
Dim Y
Dim X
For X = 1 To 18
    For Y = 1 To 18
        FontPict.CurrentX = X - 0.75
        FontPict.CurrentY = Y - 0.75
        Select Case FontPict.Point(X - 0.5, Y - 0.5)
            Case vbBlack, Black: Output = Output + "FF"
            Case White:  Output = Output + "01"
        End Select
    Next Y
Next X
Hexbox.Visible = True
Hextext = Output
Clipboard.Clear
Clipboard.SetText (Output)

End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Controls_Click / grid.frm
' Author    : Dnawrkshp
' Date      : 2/4/2013 09:57
' Purpose   : displays the shortcut help box
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Controls_Click()
    Dim text As String
    text = "Grid:" & vbCrLf & vbTab & "Ctrl-C; copies the current grid" & vbCrLf & vbTab & _
    "Ctrl-V; pastes the copied grid onto the current grid" & vbCrLf & vbTab & _
    "Ctrl-I; Moves ALL characters to the top pixel row" & vbCrLf & vbTab & _
    "Ctrl-J; Moves ALL characters to the left pixel column" & vbCrLf & vbTab & _
    "Ctrl-K; Moves ALL characters to the bottom pixel row" & vbCrLf & vbTab & _
    "Ctrl-L; Moves ALL characters to the right pixel column" & vbCrLf & vbCrLf & _
    "Character List:" & vbCrLf & vbTab & "Press any letter or number and you will jump to that letter/number on the list." & vbCrLf & vbCrLf & _
    "Font Editor:" & vbCrLf & vbTab & "For Draw and Draw Line, Left Clicking will draw and Right Clicking will erase." & vbCrLf
    MsgBox text, , "Controls"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : CopyGrid / grid.frm
' Author    : Originally written by Dnawrkshp / Re-Written by Xodus for Main Editor mode
' Date      : 2/4/2013 09:57
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub SaveCharacter()
    fontz.fdata(CurrentLetter).ByteData = ""
    Dim X As Single, Y As Single
    For X = 1 To 18
        For Y = 1 To 18
        
            Select Case FontPict.Point(X - 0.2, Y - 0.2)
                Case White
                    fontz.fdata(CurrentLetter).ByteData = fontz.fdata(CurrentLetter).ByteData + "01"
                Case vbBlack, Black
                    fontz.fdata(CurrentLetter).ByteData = fontz.fdata(CurrentLetter).ByteData + "FF"
            End Select
        Next Y
    Next X
done:
    HelpText = "Updated"
End Sub

Private Sub KeyCheck_Click()

If KeyCheck.Value = 0 Then
    KeyShort = False
Else
    KeyShort = True
End If

End Sub

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : LOGO_Click / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 09:57
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub LOGO_Click()
    HelpText = "PS2 Font Creator By Dnawrkshp. GUI and some coding by XoDu$ (c)2013 Dnawrkshp"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : opt_MouseDown / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 09:58
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub opt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim zz
    For zz = 0 To opt.Count - 1
        opt(zz).Picture = tool_still(zz).Picture
    Next zz
    opt(Index).Picture = tool_down(Index).Picture
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : HelpMenu / grid.frm
' Author    : Dnawrkshp
' Date      : 2/4/2013 09:58
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub HelpMenu()

MsgBox "Just click on the box to switch between black and white." & vbCrLf & vbCrLf & "Be sure to save as before quiting. 'New' just saves to a temporary file that gets deleted on termination." & _
"Ctrl-C will copy an image, and Ctrl-V will paste that image. If every letter is the same at launch, make a new font by pressing Ctrl-N."

End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : NewMenu / grid.frm
' Author    : Dnawrkshp
' Date      : 2/4/2013 09:57
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub NewMenu()

    'Make dummy file. Args: Filename, Size in bytes
    'MakeDumFile AppDirectory & "temp.bin", 82944
    'MakeDumFile AppDirectory & "temp SPACE.bin", 256, 1
    'GlobalFileName = AppDirectory & "temp.bin"
    'GlobalSpaceName = AppDirectory & "temp SPACE.bin"
    
    MakeDumFilePFC AppDirectory & "temp.pfc"
    GlobalFileName = AppDirectory & "temp.pfc"

End Sub

Private Sub undo_Click(Index As Integer)
FontPict.Cls
Dim X, Y
  For X = 0 To 17
    DoEvents
    For Y = 0 To 17
        FontPict.DrawWidth = 1
        If undo(Index).Point(X + 0.5, Y + 0.5) = White Then
            FontPict.Line (X, Y)-(X + 1, Y + 1), vbWhite, BF
        End If
    Next Y
  Next X
  Drawgrid
  If AutoSize.Value = 1 Then SpaceOne
  If AutoUpdate.Value = 1 Then SaveCharacter

End Sub

' _______________________________________________________________________________________________________________
'|#######################################[      tool  BUTTONS        ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Up_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A
FontChar.FontName = FontBox.FontName
FontChar.FontSize = FontBox.FontSize
FontChar.BackColor = FontBox.BackColor
FontChar.ForeColor = FontBox.ForeColor
    Select Case Index
        'font panel pan buttons
        Case 0: Ybox.text = Str(Val(Ybox.text) - 1):  FontPict.CurrentY = Val(Ybox.text): ShowFont ' up
        Case 1: Xbox.text = Str(Val(Xbox.text) - 1):  FontPict.CurrentX = Val(Xbox.text): ShowFont ' left
        Case 2: Ybox.text = Str(Val(Ybox.text) + 1):  FontPict.CurrentY = Val(Ybox.text): ShowFont ' down
        Case 3: Xbox.text = Str(Val(Xbox.text) + 1):  FontPict.CurrentX = Val(Xbox.text): ShowFont ' right
        
        'about box
        Case 4: About.Visible = False
        
        'main editor pan buttons
        Case 5: MoveRight
        Case 6: MoveDown
        Case 7: MoveLeft
        Case 8: MoveUp
    
        'font panel Previous/Next Letter
        Case 9:  A = Asc(FontBox.Caption): A = A - 1: If A = 0 Then A = 1
            FontBox.Caption = Chr$(A)
            ShowFont
        Case 10: A = Asc(FontBox.Caption): A = A + 1: If A = 256 Then A = 255
            FontBox.Caption = Chr$(A)
            ShowFont
        
        'font panel functions
        Case 11: Merge                  'Merge selected Font character into editor(copy 0x01 only)
        Case 12: ImportFont             'Import the currently selected win font as a pfc font
        Case 13: FontBox.Caption = Right(CharList1.List(CharList1.ListIndex), 1)
        
        'functions - toolbar
        Case 14: CopyImage              'copy current char to clipboard as hexcode
                 
        Case 15: PasteImage             'paste clipboard hexdata as pixels
        Case 16: AllSpace               'auto space character list
        Case 17: SpaceOne               'auto space current char
        Case 18: BlankFill              'Fill all chars with blank (324x 0xFF)
        Case 19: SaveCharacter          'update current character
        
        'File Buttons - Toolbar
        Case 20: Load_Font              'load pfc
        Case 21: SaveFont               'save as pfc
        Case 22: OpenMen                'open raw
        Case 23: SaveAsMenu             'save as raw
        Case 24: NewMenu                'new font
        Case 25: ExpO                   'export as .c
        
        'right hand toolbar
        Case 26: HelpMenu               'help file
        Case 27: Controls_Click
        Case 28: About.Visible = True   'about info
        Case 29: MsgBox "Not yet implemented!" 'SettingsForm.Show 1
    End Select

    'Reset Button colors
    Dim u, b
    For u = 0 To Up.Count - 1
        b = Val(Up(u).Tag)
        If b <> 0 Then
            Up(u).Picture = tool_still(b).Picture
        End If
    Next u
    UndoPanel.Visible = False
    


End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Up_MouseDown / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:43
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Up_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim b
    b = Val(Up(Index).Tag)
    If b <> 0 Then
        Up(Index).Picture = tool_down(b).Picture
    End If
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Up_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:43
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Up_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
21    showhelp Index
    
End Sub



' _______________________________________________________________________________________________________________
'|#######################################[    LEFT PANEL FUNCTIONS   ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : CharList1_KeyPress / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:44
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub CharList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 3: CopyImage: Exit Sub                'Copy - Ctrl-C
        Case 22: PasteImage: Exit Sub              'Paste - Ctrl-V
        Case 9: MoveAllUp: Exit Sub                'Quick move all chars to left - Ctrl-I
        Case 10: MoveAllLeft: Exit Sub             'Quick move all chars to left - Ctrl-J
        Case 11:  MoveAllDown: Exit Sub             'Quick move all chars to left - Ctrl-K
        Case 12:  MoveAllRight: Exit Sub            'Quick move all chars to left - Ctrl-L
        Case 19:  QuickSaveFont: Exit Sub           'Saves all characters to current GlobalFileName
    End Select
    
    If KeyAscii >= 0 And KeyAscii < 255 And KeyShort = False Then: CharList1.ListIndex = KeyAscii
End Sub

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : CharList1_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:44
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub CharList1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = "Click to select a character"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : CharList1_Click / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:44
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub CharList1_Click()
    
    If isdoing = True Then Exit Sub
    
    CurrentLetter = CharList1.ListIndex
    prevue.Cls
    prevue.Scale (0, 0)-(20, 18)
    prevue.DrawWidth = 1
    
    If Import = False Then
        drawfont
        max = fontz.fdata(CurrentLetter).spacing
        If max = 0 Then max = 18
        Text1 = max
        InitMinMax
        Drawgrid
        DrawMiniGrid
        If AutoSize.Value = 1 Then SpaceOne
    End If

End Sub
' _______________________________________________________________________________________________________________
'|#######################################[   FONT PANEL FUNCTIONS    ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : FntSiz_Change / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:27
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub FntSiz_Change()
    If Val(fontsiz.Caption) = 0 Then Exit Sub
    fontsiz.Caption = Str(Val(FntSiz.Value) - 1)
    FontPict.FontSize = FntSiz.Value
    ShowFont
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : FontBox_Change / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:27
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub FontBox_Change()
    'LoadFont_Click
    FontChar.FontName = FontBox.FontName
    If Import = False Then
      If FontBox.Caption <> "" Then
        ascchar = AscW(FontBox.Caption)
        hexchar = Hex(ascchar)
      End If
    End If
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : FntSiz_Scroll / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:30
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub FntSiz_Scroll()
    fontsiz.Caption = Str(Val(FntSiz.Value) - 1)
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Fontbox_DblClick / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:27
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Fontbox_DblClick()
    FontChar.text = ""
    FontChar.Visible = True
    FontChar.ZOrder
    FontChar.SetFocus
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : FontChar_KeyPress / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:30
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub FontChar_KeyPress(KeyAscii As Integer)
    FontBox.Caption = ChrW$(KeyAscii)
    FontChar.Visible = False
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : FontPict_KeyDown / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:27
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub FontPict_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            drawing = False
            StrightLine.Visible = False
            Rectangle.Visible = False
            Circles.Visible = False
    End Select
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : ExpO / grid.frm
' Author    : Dnawrkshp
' Date      : 2/4/2013 10:27
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub ExpO()
    Dim LongX As Long
    LongX = ArrayName.hwnd
    ArrayName.Show
    While IsWindow(LongX)
        DoEvents
    Wend
    If ReturnVal = False Then: Exit Sub
    If ExtOf(GlobalFileName) = ".PFC" Then
        SaveAsCPFC FontSaveLoc, SpaceSaveLoc, FontArray, SpaceArray
    Else
        SaveAsCBIN FontSaveLoc, SpaceSaveLoc, FontArray, SpaceArray
    End If
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : winfontlist_Click / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:28
' Purpose   : event fired when win font list is clicked
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub winfontlist_Click()
    If winfontlist.text = "winfontlist" Then Exit Sub
    If winfontlist.text = "" Then Exit Sub
    FontBox.FontName = winfontlist.text
    FontPict.FontName = winfontlist.text
    ShowFont
    FontChar.Font = FontBox.Font
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Merge / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:30
' Purpose   : Saves the editor contents to the current character array
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Merge()
  Dim X, Y
  For X = 0 To 17
    DoEvents
    For Y = 0 To 17
        FontPict.DrawWidth = 1
        If SmallFontpic.Point(X + 0.5, Y + 0.5) = White Then
            FontPict.Line (X, Y)-(X + 1, Y + 1), vbWhite, BF
        End If
    Next Y
  Next X
  Drawgrid
  If AutoSize.Value = 1 Then SpaceOne
  If AutoUpdate.Value = 1 Then SaveCharacter

End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : opt_Click / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:29
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub opt_Click(Index As Integer)
  Hexbox.Visible = False
  UndoPanel.Visible = False
  FontFrame.Enabled = False
    FontFrame.Visible = False
    
    drawtype = Index
    If Index = 6 Then
        FontFrame.Enabled = True
        FontFrame.Visible = True
    End If
    
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : ImportFont / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:29
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub ImportFont()
    Dim SelIndex As Integer, Old As Integer
    Old = CharList1.ListIndex
    SmallFontpic.Visible = Visible = False
    IMP_PROG.Visible = True
    IMP_PROG.Cls
    IMP_PROG.Scale (0, 0)-(255, 1)
    fontz.Name = winfontlist.text
    For SelIndex = 0 To 255
        CurrentLetter = SelIndex
        Import = True
        FontBox.Caption = ChrW$(SelIndex) 'HexToString(Name)
        IMP_PROG.Line (0, 0)-(SelIndex, 1), &HBCA93F, BF
        If FontBox.Caption <> "" Then
            ShowFont 'Puts character in the picture box
            GetFontData
        End If
    Next SelIndex
    SmallFontpic.Visible = True
    Import = False
    CharList1.ListIndex = Old
    IMP_PROG.Visible = False
    MsgBox "Import complete"
    CharList1.ListIndex = 0
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : ShowFont / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:29
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub ShowFont()
    SmallFontpic.Cls
    SmallFontpic.FontSize = FntSiz.Value
    SmallFontpic.FontName = FontBox.FontName
    SmallFontpic.Scale (0, 0)-(18, 18)
    SmallFontpic.CurrentX = Val(Xbox.text)
    SmallFontpic.CurrentY = Val(Ybox.text)
    SmallFontpic.ForeColor = White
    SmallFontpic.Print FontBox.Caption
    Pixelate SmallFontpic
    DrawMiniGrid
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : opt_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:29
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub opt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Index
        Case 0: HelpText = "Line tool. Click then drag [L]"
        Case 1: HelpText = "Box draw. Click and drag [S]"
        Case 2: HelpText = "Circle Tool [C])"
        Case 3: HelpText = "Filled Box [B]"
        Case 4: HelpText = "Draw Tool [P]"
        Case 5: HelpText = "Line Draw [D]"
        Case 6: HelpText = "Use Windows Font mode [F]"
    End Select
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : PasteImage / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:29
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub PasteImage()

Dim Input1 As String
Input1 = Clipboard.GetText
FontPict.Cls
Hexbox.Visible = True
Hextext = Input1

FontPict.Scale (0, 0)-(20, 18)
FontPict.DrawWidth = 1
Dim f, Y, c
c = 1
For f = 1 To 18
    For Y = 1 To 18
        Select Case Mid(Input1, c, 2)
            Case "FF": FontPict.Line (f, Y)-(f - 1, Y - 1), vbBlack, BF
            Case "01": FontPict.Line (f, Y)-(f - 1, Y - 1), White, BF
        End Select
        c = c + 2
    Next Y
Next f
Drawgrid
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : drawfont / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:29
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Sub drawfont()
    Dim X, Y, z
    z = 1
    FontPict.Cls
    FontPict.Scale (0, 0)-(20, 18)
    FontPict.DrawWidth = 1
    If fontz.fdata(CurrentLetter).ByteData = "" Then Exit Sub
    
    For X = 1 To 18
        For Y = 1 To 18
            Select Case Mid(fontz.fdata(CurrentLetter).ByteData, z, 2)
                Case "FF": FontPict.Line (X, Y)-(X - 1, Y - 1), vbBlack, BF: prevue.Line (X, Y)-(X + 1, Y + 1), vbBlack, BF
                Case "01": FontPict.Line (X, Y)-(X - 1, Y - 1), vbWhite, BF: prevue.Line (X, Y)-(X + 1, Y + 1), vbWhite, BF
            End Select
            z = z + 2
        Next Y
    Next X
    
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : updateundo / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:29
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Sub updateundo()
    Dim X, Y
    If Undos = 12 Then Undos = 0
    undo(Undos).Cls
    For X = 1 To 17
        For Y = 1 To 17
            If FontPict.Point(X - 0.5, Y - 0.5) = White Then
                undo(Undos).Line (X - 1, Y - 1)-(X, Y), White, BF
            End If
        Next Y
    Next X
    Undos = Undos + 1

End Sub
' _______________________________________________________________________________________________________________
'|#######################################[    CHARACTER SPACING      ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : AllSpace / grid.frm
' Author    : Xodus(original written by Dnawrkshp - rewritten for editor by Xodus)
' Date      : 2/4/2013 09:17
' Purpose   : auto spaces all characters in current font
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub AllSpace()
    Dim Old As Integer                                              'initialize an array to store current position
    Old = CharList1.ListIndex                                       'store current list position
    Dim q                                                           'dimension array for loop
    For q = 0 To 255                                                'loop through character list
        CharList1.ListIndex = q                                     'set list index = loop counter
        SpaceOne                                                    'Call Autosize feature
    Next q                                                          'loop
    CharList1.ListIndex = Old                                       'restore list index
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : SpaceOne / grid.frm
' Author    : Xodus (original written by Dnawrkshp - rewritten for editor by Xodus)
' Date      : 2/4/2013 09:17
' Purpose   : auto spaces current character
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub SpaceOne()
    Dim Result As Integer                                           'dimension temp array for output value
    Dim X, Y                                                        'dimension temp loop arrays
    Result = 0:                                                     'Init output array
    For X = 18 To 1 Step -1                                         'Loop for each column in grid
        For Y = 1 To 18                                             'Loop for each row in grid
            If FontPict.Point(X - 0.2, Y - 0.2) = White Then        'Is there a pixel in this column?
                Result = X: GoTo done                               'if so set result = this column (x)
            End If
        Next Y                                                      'loop
    Next X
done:
    If Result = 0 Then: max = 10                                    'If result failed, default to 10
    max = Result + 2                                                'Set Max property as result
    If CharList1.ListIndex = 32 Then: max = 8                       'In case character is a space
    fontz.fdata(CurrentLetter).spacing = max                        'Set the current characters size property
    InitMinMax                                                      'Call Spacer display function
    
    If ExtOf(GlobalFileName) <> ".PFC" And Dir(GlobalSpaceName) <> "" Then
        SaveSpaceMarker
    End If
    
End Sub

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : RightTick_MouseDown / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:25
' Purpose   : sets right clicking of pip as auto-size
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub RightTick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then                                              'if right button clicked while over pip
        SpaceOne                                                    'call auto space function
    End If
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Sliderpic_DragOver / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 10:16
' Purpose   : updates the max data and screen items associated with it while the pip is in motion
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Sliderpic_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Sliderpic.ScaleWidth = 20                                       'constrain scale
    If Source.Left < -0.31 Then Source.Left = -0.31                 'check left boundry
    If Source.Left > 20 - 0.31 Then Source.Left = 20 - 0.31         'check Right boundry
    If X < min + 1 Then X = min + 1
    
    Source.Left = Int(X) - (Source.Width / 2)                       'set pip position
    maxLine.x1 = Int(X)                                             'set guide line top position (x)
    maxLine.x2 = Int(X)                                             'set guide line bottom postion(x)
    maxLine.Y1 = 0: maxLine.Y2 = 18                                 'Set guide line horizontal position
    max = Int(X)                                                    'Set max property based on pip position
    fontz.fdata(CurrentLetter).spacing = max                        'Set the current characters size property
    Text1.text = Str(max - min)                                     'Display the size on screen
    DrawMiniGrid                                                    'Update the minigrid accordingly
    With SizeFx                                                     'Update the sizer frame display
        .Left = min                                                 'set Sizer display left
        .Width = max - min                                          'set Sizer display width
        .Top = -1                                                   'set Sizer display top
        .Height = 2                                                 'Set sizer display height
    End With
    Text1.Left = (max - ((max - min) / 2)) - (Text1.Width / 2)      'update space display text position
    Text1.Top = 0.7
End Sub

' _______________________________________________________________________________________________________________
'|#######################################[    FILE I/O FUNCTIONS     ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : SaveFont / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 09:23
' Purpose   : Saves the current font in PFC format
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub SaveFont()
    Dim filename As String
    CD.Filter = "Font Files (*.pfc)|*.pfc|All files (*.*)|*.*"
    CD.DefaultExt = "pfc"
    CD.DialogTitle = "Select File"
    CD.Action = 2
    filename = CD.filename
    If filename = "" Or Dir(filename) = "" Then Exit Sub
    
    If ExtOf(GlobalFileName) = ".PFC" Then
        Save_Font filename
    Else
        Save_BINFont filename
        MsgBox "Successfully saved as " & CD.filename
    End If
    
    GlobalFileName = CD.filename
End Sub

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Load_Font / grid.frm
' Author    : Dnawrkshp /(later revised by Xodus)
' Date      : 2/4/2013 12:47
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Load_Font()
    Dim filename As String
    CD.Filter = "Font Files (*.pfc)|*.pfc|All files (*.*)|*.*"
    CD.DefaultExt = "pfc"
    CD.DialogTitle = "Select File"
    CD.Action = 1
    filename = CD.filename
    If Dir(filename) = "" Or filename = "" Then Exit Sub
    OPEN_FONT filename
    GlobalFileName = CD.filename
    If ExtOf(GlobalFileName) <> ".PFC" Then: GlobalSpaceName = Replace(GlobalFileName, ".", " SPACE.", , 1)
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : SaveAsMenu / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub SaveAsMenu()
    InitCD
    CD.Action = 2
    If CD.filename = "" Or Dir(CD.filename) = "" Then: Exit Sub
    'If CD.filename = GlobalFileName Then: Exit Sub
    If ExtOf(GlobalFileName) = ".PFC" Then
        SaveBINfromPFC CD.filename, Replace(CD.filename, ".", " SPACE.", , 1)
    Else
        'FileCopy GlobalFileName, CD.filename
        'FileCopy GlobalSpaceName, Replace(CD.filename, ".", " SPACE.", , 1)
        SaveBIN CD.filename, Replace(CD.filename, ".", " SPACE.", , 1)
    End If
    GlobalFileName = CD.filename
    GlobalSpaceName = Replace(GlobalFileName, ".", " SPACE.", , 1)
    UpdateSaveFile
    CD.filename = ""
    MsgBox "Successfully Saved As " & GlobalFileName
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : OpenMen / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub OpenMen()

    InitCD
    CD.Action = 1
    If CD.filename = "" Or Dir(CD.filename) = "" Or CD.filename = GlobalFileName Then: Exit Sub
    
    If Dir(AppDirectory & "temp.bin") <> "" Then: Kill AppDirectory & "temp.bin"
    If Dir(AppDirectory & "temp SPACE.bin") <> "" Then: Kill AppDirectory & "temp SPACE.bin"
    
    GlobalFileName = CD.filename
    GlobalSpaceName = Replace(GlobalFileName, ".", " SPACE.", , 1)
    If Dir(GlobalSpaceName) = "" Then: MakeDumFile GlobalSpaceName, 255, 1
    CharList1.ListIndex = 0
    Dim OFS As Long
    Dim GRABDATA As Integer
    For GRABDATA = 0 To 255
        OFS = ((GRABDATA * 32.4) * 10) + 1
        fontz.fdata(GRABDATA).ByteData = AccessBinSegment(OFS, 324)
        fontz.fdata(GRABDATA).spacing = AccessSpaceSegment(GRABDATA + 1)
    Next GRABDATA
    CharList1.ListIndex = 0
    MsgBox "Finished Importing data"
    CD.filename = ""
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : InitCD / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Sub InitCD()
    CD.Filter = "RAW's (*.raw)|*.raw|BIN's (*.bin)|*.bin|All files (*.*)|*.*"
    CD.DefaultExt = "raw"
    CD.DialogTitle = "Select File"
    
End Sub

' _______________________________________________________________________________________________________________
'|#######################################[    GRID MAINTENANCE       ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Sub GetFontData()
    Dim X, Y
    fontz.fdata(CurrentLetter).ByteData = ""
      For X = 0 To 17
       DoEvents
        For Y = 0 To 17
            FontPict.DrawWidth = 1
            Select Case SmallFontpic.Point(X + 0.5, Y + 0.5)
                Case White: fontz.fdata(CurrentLetter).ByteData = fontz.fdata(CurrentLetter).ByteData + "01"
                Case Black, vbBlack: fontz.fdata(CurrentLetter).ByteData = fontz.fdata(CurrentLetter).ByteData + "FF"
            End Select
        Next Y
      Next X
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Gui_ClearEditor / grid.frm
' Author    : Xodus
' Date      : 1/28/2013 03:58
' Purpose   : Clears the Main Editor
'---------------------------------------------------------------------------------------
Function Gui_ClearEditor()
    FontPict.Cls
    FontPict.ScaleWidth = 20
    FontPict.ScaleHeight = 18
    Drawgrid
End Function

'---------------------------------------------------------------------------------------
' Procedure : Gui_ClearPreview / grid.frm
' Author    : Xodus
' Date      : 1/28/2013 03:59
' Purpose   : Clears mini Font preview pic
'---------------------------------------------------------------------------------------
Function Gui_ClearPreview()
    SmallFontpic.Cls
    SmallFontpic.ScaleWidth = 80
    SmallFontpic.ScaleHeight = 18
    DrawMiniGrid
End Function

'---------------------------------------------------------------------------------------
' Procedure : DrawMiniGrid
' Author    : Xodus
' Date      : 1/28/2013 11:10
' Purpose   : REDRAWS THE EDITOR / MINIFONT OUTPUT GRIDS
'---------------------------------------------------------------------------------------
Sub DrawMiniGrid()
    'MiniFont
    SmallFontpic.Scale (0, 0)-(18, 18)
    SmallFontpic.DrawStyle = vbSolid
    Dim gx As Integer, x1, x2
    For gx = 1 To 17: SmallFontpic.Line (gx, 0)-(gx, 18), darkgrey:  Next gx
    For gx = 1 To 17: SmallFontpic.Line (0, gx)-(18, gx), darkgrey:  Next gx
    x1 = Int(TopGuide.Y1)
    x2 = Int(BottomGuide.Y1)
    SmallFontpic.DrawStyle = vbDot
    SmallFontpic.Line (0, x1)-(18, x1), QBColor(3)
    SmallFontpic.Line (0, x2)-(18, x2), QBColor(3)
    SmallFontpic.Line (max, 0)-(max, 18), QBColor(10)
    SmallFontpic.Line (min, 0)-(min, 18), QBColor(6)
    SmallFontpic.DrawStyle = vbSolid
    SmallFontpic.Line (0, 0)-(17.9, 17.9), QBColor(11), B
End Sub

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Drawgrid / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Sub Drawgrid()
    'Main Editor Grid
    '-----------------------------------------
    FontPict.FontSize = 8
    FontPict.DrawWidth = 1
    FontPict.Line (18, 0)-(22, 18), &H0, BF
    Dim gx, gy, COL As Long
    For gx = 1 To 22
        For gy = 1 To 18
        
            FontPict.Line (gx - 1, gy - 1)-(gx, gy), &H404000, B
            SmallFontpic.Line (gx - 1, gy - 1)-(gx, gy), &H404000, B
            Sliderpic.Line (gx, 0)-(gx, 0.5), &H404000
            If Showbyte.Value = 1 Then
                FontPict.CurrentX = gx - 0.75
                FontPict.CurrentY = gy - 0.75
                COL = FontPict.Point(gx - 0.1, gy - 0.1)
                If COL = White Then FontPict.ForeColor = vbBlack: FontPict.Print "01"
                If COL = vbBlack Then FontPict.ForeColor = White: FontPict.Print "FF"
            End If
            
        Next gy
    Next gx
    
    FontPict.Line (18, 0)-(18, 18), vbRed
    FontPict.Line (0, 0)-(19.9, 17.9), QBColor(11), B
    FontPict.DrawWidth = 20
    FontPict.ForeColor = Black
    
    'Slider Pic
    '-----------------------------------------
    Sliderpic.Line (0, 0)-(Sliderpic.ScaleWidth - 0.1, Sliderpic.ScaleHeight - 0.1), QBColor(11), B
    
    
End Sub



' _______________________________________________________________________________________________________________
'|#######################################[    DRAWING FUNCTIONS      ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'---------------------------------------------------------------------------------------
' Procedure : Pixelate / grid.frm
' Author    : Xodus
' Date      : 1/28/2013 11:08
' Purpose   : CONVERTS THE CURVED INPUT TO PIXELATED OUTPUT USING FILL PERCENTAGE ENGINE
'---------------------------------------------------------------------------------------

Sub Pixelate(where As Control)
'where.Cls
Dim pixels As Integer
Dim xs As Single, ys As Single
DoEvents
prevue.Cls
prevue.Scale (0, 0)-(18, 18)
Dim X As Single, Y As Single
where.DrawWidth = 1
For Y = 0 To 17
    For X = 0 To 17
        pixels = 0
        For xs = (X + 0.1) To X + 0.9 Step 0.1
            For ys = (Y + 0.1) To Y + 0.9 Step 0.1
                If where.Point(ys, xs) = White Then
                    pixels = pixels + 1
                End If
            Next ys
        Next xs
        If pixels > Val(dither) Then 'FILLED PERCENT TO CHECK FOR (MIN 1 - MAX 81)
            where.Line (Y, X)-(Y + 1, X + 1), vbWhite, BF
            If Import = False Then prevue.Line (Y, X)-(Y + 1, X + 1), vbWhite, BF
          Else
            where.Line (Y, X)-(Y + 1, X + 1), vbBlack, BF
            If Import = False Then prevue.Line (Y, X)-(Y + 1, X + 1), vbBlack, BF
        End If
    Next X
Next Y
done:
If AutoSize.Value = 1 Then SpaceOne
If AutoUpdate.Value = 1 Then SaveCharacter

If Import = False Then updateundo
End Sub


' _______________________________________________________________________________________________________________
'|#######################################[    HELP EVENT FUNCTIONS   ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : prevue_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub prevue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = "An actual size preview will be displayed here"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Sliderpic_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Sliderpic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = "Set spacing margins by moving the arrows left and right"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : SmallFontpic_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub SmallFontpic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = "This is a preview of the imported font, used for ajustment"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : TOPSLIDER_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub TOPSLIDER_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = "These lines serve as guidelines for alignment. drag to move up/dn"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : RightTick_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub RightTick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = "Max margin - press and hold left mouse and drag pip (currently:" & max & ")"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : ClearIt_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub ClearIt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Select Case Index
            Case 0: HelpText = "Clear Editbox"
            Case 1: HelpText = "Clear Current Character"
            Case 2: HelpText = "Undo (up to 12 steps)"
        End Select
End Sub

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : FontBox_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub FontBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = "Selected Character. Double Click and type a letter to change"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : winfontlist_MouseMove / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub winfontlist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpText = "Click to select a windows font"
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : showhelp / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Sub showhelp(Index As Integer)
    Select Case Index
        Case 0, 8: HelpText = "move selected letter up one row"
        Case 1, 7: HelpText = "move selected letter left one row"
        Case 2, 6: HelpText = "move selected letter down one row"
        Case 3, 5: HelpText = "move selected letter right one row"
        Case 9: HelpText = "previous letter in font set"
        Case 10: HelpText = "Next letter in font set"
        Case 11: HelpText = "Merge the above letter into the editor"
        Case 12: HelpText = "Import entire windows font into character set"
        Case 13: HelpText = "Use current character in set list as selection"
        Case 14: HelpText = "Copy character image as Hex data to clipboard"
        Case 15: HelpText = "Import Hex data from clipboard as pixels"
        Case 16: HelpText = "Auto Size all characters in character set"
        Case 17: HelpText = "Auto Size current character"
        Case 18: HelpText = "reset all characters to blank"
        Case 19: HelpText = "Copy Editor contents to current charcter"
        Case 20: HelpText = "Load a saved font file"
        Case 21: HelpText = "Save the current font set"
        Case 22: HelpText = "Import a font from raw hex data"
        Case 23: HelpText = "export the current font set as raw hex data"
        Case 24: HelpText = "New Font.. Caution you will loose any unsaved data"
        Case 25: HelpText = "export font as .c"
        Case 26: HelpText = "Font Creator Help"
        Case 28: HelpText = "Information about Font Creator"
        Case 29: HelpText = "Program Settings"
        Case Else
            HelpText = "You dont have a help text for this button"
  End Select
End Sub
' _______________________________________________________________________________________________________________
'|#######################################[    EDITOR PAN FUNCTIONS   ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'=============================== PAN UP FUNCTIONS (MOVES THE EDITOR PICTURE 1 PIXEL UP == ========================
Sub MoveAllUp()
Dim SelIndex As Integer, Old As Integer
Old = CharList1.ListIndex
For SelIndex = 0 To 255: CharList1.ListIndex = SelIndex: Up_MouseUp 9, 1, 1, 1, 1: Up_MouseUp 19, 1, 1, 1, 1: Next SelIndex
CharList1.ListIndex = Old
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : MoveUp / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:48
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub MoveUp()
    Pixelate FontPict
    FontPict.DrawWidth = 1
    Dim buffer(18) As Long
    Dim X As Single, Y As Single
    Dim COL As Long
    For X = 1 To 18: buffer(X) = FontPict.Point(X - 0.5, 0.5): Next X
    For X = 1 To 18: For Y = 2 To 18: COL = FontPict.Point(X - 0.5, Y - 0.5): FontPict.Line (X - 1, Y - 2)-(X, Y - 1), COL, BF: Next Y:  Next X
    For X = 1 To 18: COL = buffer(X): FontPict.Line (X - 1, 17)-(X, 18), COL, BF: Next X
    Drawgrid
End Sub
'=============================== PAN DOWN  FUNCTIONS (MOVES THE EDITOR PICTURE 1 PIXEL DOWN  =====================
Sub MoveAllDown()
Dim SelIndex As Integer, Old As Integer
Old = CharList1.ListIndex
For SelIndex = 0 To 255: CharList1.ListIndex = SelIndex: Up_MouseUp 6, 1, 1, 1, 1: Up_MouseUp 19, 1, 1, 1, 1: Next SelIndex
CharList1.ListIndex = Old
End Sub
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : MoveDown / grid.frm
' Author    : Xodus
' Date      : 2/4/2013 12:49
' Purpose   : Move the contents of the Editor down 1 row
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub MoveDown()
    Pixelate FontPict
    FontPict.DrawWidth = 1
    Dim buffer(18) As Long
    Dim X As Single, Y As Single
    Dim COL As Long
    For X = 1 To 18: buffer(X) = FontPict.Point(X - 0.5, 17.5):  Next X
    For X = 1 To 18: For Y = 17 To 1 Step -1: COL = FontPict.Point(X - 0.5, Y - 0.5): FontPict.Line (X - 1, Y + 1)-(X, Y), COL, BF: Next Y: Next X
    For X = 1 To 18: COL = buffer(X): FontPict.Line (X - 1, 0)-(X, 1), COL, BF: Next X
    Drawgrid
End Sub
'=============================== PAN RIGHT FUNCTIONS (MOVES THE EDITOR PICTURE 1 PIXEL RIGHT =====================
Sub MoveAllRight()
    Dim SelIndex As Integer, Old As Integer
    Old = CharList1.ListIndex
    For SelIndex = 0 To 255: CharList1.ListIndex = SelIndex:  Up_MouseUp 5, 1, 1, 1, 1: Up_MouseUp 19, 1, 1, 1, 1: Next SelIndex
    CharList1.ListIndex = Old
End Sub
Private Sub MoveRight()
    Pixelate FontPict
    FontPict.DrawWidth = 1
    Dim buffer(18) As Long
    Dim X As Single, Y As Single
    Dim COL As Long
    For Y = 1 To 18: buffer(Y) = FontPict.Point(17.5, Y - 0.5):    Next Y
    For Y = 1 To 18: For X = 17 To 1 Step -1: COL = FontPict.Point(X - 0.5, Y - 0.5):  FontPict.Line (X, Y - 1)-(X + 1, Y), COL, BF: Next X:  Next Y
    For Y = 1 To 18: COL = buffer(Y): FontPict.Line (0, Y - 1)-(1, Y), COL, BF: Next Y
    Drawgrid
End Sub
'=============================== PAN LEFT FUNCTIONS (MOVES THE EDITOR PICTURE 1 PIXEL LEFT =======================

Sub MoveAllLeft()
    Dim SelIndex As Integer, Old As Integer
    Old = CharList1.ListIndex
    For SelIndex = 0 To 255: CharList1.ListIndex = SelIndex: Up_MouseUp 7, 1, 1, 1, 1: Up_MouseUp 19, 1, 1, 1, 1: Next SelIndex
    CharList1.ListIndex = Old
End Sub
Private Sub MoveLeft()
    Pixelate FontPict
    FontPict.DrawWidth = 1
    Dim buffer(18) As Long
    Dim X As Single, Y As Single
    Dim COL As Long
    For Y = 1 To 18: buffer(Y) = FontPict.Point(0.5, Y - 0.5): Next Y
    For Y = 1 To 18: For X = 2 To 18: COL = FontPict.Point(X - 0.5, Y - 0.5): FontPict.Line (X - 2, Y - 1)-(X - 1, Y), COL, BF:  Next X: Next Y
    For Y = 1 To 18: COL = buffer(Y): FontPict.Line (17, Y - 1)-(18, Y), COL, BF: Next Y:
    Drawgrid
End Sub
' _______________________________________________________________________________________________________________
'|#######################################[ FONT ALTERATION FUNCTIONS ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯



' _______________________________________________________________________________________________________________
'|#######################################[     TRIGGER FUNCTIONS     ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Timer1_Timer()
    If Val(fontsiz.Caption) = 0 Then Exit Sub
    fontsiz.Caption = Int(fscale.slidervalue) - 1
    FontPict.FontSize = Int(fscale.slidervalue) - 1
    FntSiz.Value = fscale.slidervalue
    Timer1.Interval = 0
End Sub
Private Sub Timer2_Timer()
    dither = DitherPick.slidervalue
    If dither = 0 Then dither = 10
    ShowFont
    Timer2.Interval = 0
    Dith.Caption = "Dither Tolerance " & dither & " %"
End Sub


' _______________________________________________________________________________________________________________
'|#######################################[    FORM SKIN FUNCTIONS    ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'---------------------------------------------------------------------------------------
' Procedure : CONTROL FADE
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Paints a Gradient on a CONTROL
'---------------------------------------------------------------------------------------
Sub ControlFade(Ctl As Control, Rc As Long, Gc As Long, bc As Long)
        'This code works best when called in the paint event
          Dim t1, t2, t3
10        t1 = Rc / 255
20        t2 = Gc / 255
30        t3 = bc / 255
          Dim intLoop As Integer
50        On Error Resume Next
60        With Ctl
70            .DrawStyle = vbInsideSolid
80            .DrawMode = vbCopyPen
'90            .ScaleMode = vbPixels
100           .DrawWidth = 2
110           .ScaleHeight = 256
120           .AutoRedraw = True
130           .ClipControls = False
140       End With
            
150       For intLoop = 0 To 256
160           Ctl.Line (-1, intLoop)-(Ctl.Width, intLoop - 1), RGB(t1 * intLoop, t2 * intLoop, t3 * intLoop), BF    'Draw boxes With specified color of loop
170       Next intLoop

End Sub
Sub ControlRollFade(Ctl As Control, Rc As Long, Gc As Long, bc As Long)
        'This code works best when called in the paint event
          Dim t1, t2, t3
10        t1 = (Rc / 255)
20        t2 = (Gc / 255)
30        t3 = (bc / 255)
          Dim intLoop As Integer
50        On Error Resume Next
60        With Ctl
70            .DrawStyle = vbInsideSolid
80            .DrawMode = vbCopyPen
'90            .ScaleMode = vbPixels
100           .DrawWidth = 2
110           .ScaleHeight = 256
120           .AutoRedraw = True
130           .ClipControls = False
140       End With
            
150       For intLoop = 0 To 127
160           Ctl.Line (-1, intLoop)-(Ctl.Width, intLoop - 1), RGB(t1 * intLoop, t2 * intLoop, t3 * intLoop), BF    'Draw boxes With specified color of loop
170       Next intLoop

          For intLoop = 1 To 128
              Ctl.Line (-1, 255 - intLoop)-(Ctl.Width, 255 - (intLoop - 1)), RGB(t1 * intLoop, t2 * intLoop, t3 * intLoop), BF 'Draw boxes With specified color of loop
          Next intLoop

End Sub

' _______________________________________________________________________________________________________________
'|———————————————————————————————————————[     INITIALIZATIONS       ]———————————————————————————————————————————|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Sub AddFonts()
    Dim I As Integer
    For I = 1 To Screen.FontCount
        winfontlist.AddItem Screen.Fonts(I)
    Next I
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitMinMax / grid.frm
' Author    : Xodus
' Date      : 1/28/2013 11:07
' Purpose   : INITIALIZES AND REFRESHES THE MARGIN SIZE SLIDER
'---------------------------------------------------------------------------------------
Sub InitMinMax()
        Sliderpic.ScaleWidth = 20
        MinLine.x1 = min
        MinLine.x2 = min
        MinLine.Y1 = 0: MinLine.Y2 = 18
        maxLine.x1 = max
        maxLine.x2 = max
        maxLine.Y1 = 0: maxLine.Y2 = 18
        With SizeFx
            .Left = min
            .Width = max - min
            .Top = -1
            .Height = 2
        End With
        Text1.Left = (max - ((max - min) / 2)) - (Text1.Width / 2)
        Text1.Top = 0.7
        Text1 = max
        'LeftTick.Left = min - (LeftTick.Width / 2)
        RightTick.Left = max - (RightTick.Width / 2)
              
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PopulateList / grid.frm
' Author    : Dnawrkshp
' Date      : 1/28/2013 13:50
' Purpose   : Populates the Character List at start-up
'---------------------------------------------------------------------------------------

Sub PopulateList()
CharList1.AddItem "0x01"
CharList1.AddItem "0x02"
CharList1.AddItem "0x03"
CharList1.AddItem "0x04"
CharList1.AddItem "0x05"
CharList1.AddItem "0x06"
CharList1.AddItem "0x07"
CharList1.AddItem "0x08"
CharList1.AddItem "0x09"
CharList1.AddItem "0x0A"
CharList1.AddItem "0x0B"
CharList1.AddItem "0x0C"
CharList1.AddItem "0x0D"
CharList1.AddItem "0x0E"
CharList1.AddItem "0x0F"
CharList1.AddItem "0x10"
CharList1.AddItem "0x11"
CharList1.AddItem "0x12"
CharList1.AddItem "0x13"
CharList1.AddItem "0x14"
CharList1.AddItem "0x15"
CharList1.AddItem "0x16"
CharList1.AddItem "0x17"
CharList1.AddItem "0x18"
CharList1.AddItem "0x19"
CharList1.AddItem "0x1A"
CharList1.AddItem "0x1B"
CharList1.AddItem "0x1C"
CharList1.AddItem "0x1D"
CharList1.AddItem "0x1E"
CharList1.AddItem "0x1F"
CharList1.AddItem "0x20   (Space)"
CharList1.AddItem "0x21   !"
CharList1.AddItem "0x22   " & Chr$(34)
CharList1.AddItem "0x23   #"
CharList1.AddItem "0x24   $"
CharList1.AddItem "0x25   %"
CharList1.AddItem "0x26   &"
CharList1.AddItem "0x27   '"
CharList1.AddItem "0x28   ("
CharList1.AddItem "0x29   )"
CharList1.AddItem "0x2A   *"
CharList1.AddItem "0x2B   +"
CharList1.AddItem "0x2C   ,"
CharList1.AddItem "0x2D   -"
CharList1.AddItem "0x2E   ."
CharList1.AddItem "0x2F   /"
CharList1.AddItem "0x30   0"
CharList1.AddItem "0x31   1"
CharList1.AddItem "0x32   2"
CharList1.AddItem "0x33   3"
CharList1.AddItem "0x34   4"
CharList1.AddItem "0x35   5"
CharList1.AddItem "0x36   6"
CharList1.AddItem "0x37   7"
CharList1.AddItem "0x38   8"
CharList1.AddItem "0x39   9"
CharList1.AddItem "0x3A   :"
CharList1.AddItem "0x3B   ;"
CharList1.AddItem "0x3C   <"
CharList1.AddItem "0x3D   ="
CharList1.AddItem "0x3E   >"
CharList1.AddItem "0x3F   ?"
CharList1.AddItem "0x40   @"
CharList1.AddItem "0x41   A"
CharList1.AddItem "0x42   B"
CharList1.AddItem "0x43   C"
CharList1.AddItem "0x44   D"
CharList1.AddItem "0x45   E"
CharList1.AddItem "0x46   F"
CharList1.AddItem "0x47   G"
CharList1.AddItem "0x48   H"
CharList1.AddItem "0x49   I"
CharList1.AddItem "0x4A   J"
CharList1.AddItem "0x4B   K"
CharList1.AddItem "0x4C   L"
CharList1.AddItem "0x4D   M"
CharList1.AddItem "0x4E   N"
CharList1.AddItem "0x4F   O"
CharList1.AddItem "0x50   P"
CharList1.AddItem "0x51   Q"
CharList1.AddItem "0x52   R"
CharList1.AddItem "0x53   S"
CharList1.AddItem "0x54   T"
CharList1.AddItem "0x55   U"
CharList1.AddItem "0x56   V"
CharList1.AddItem "0x57   W"
CharList1.AddItem "0x58   X"
CharList1.AddItem "0x59   Y"
CharList1.AddItem "0x5A   Z"
CharList1.AddItem "0x5B   ["
CharList1.AddItem "0x5C   \"
CharList1.AddItem "0x5D   ]"
CharList1.AddItem "0x5E   ^"
CharList1.AddItem "0x5F   _"
CharList1.AddItem "0x60   `"
CharList1.AddItem "0x61   a"
CharList1.AddItem "0x62   b"
CharList1.AddItem "0x63   c"
CharList1.AddItem "0x64   d"
CharList1.AddItem "0x65   e"
CharList1.AddItem "0x66   f"
CharList1.AddItem "0x67   g"
CharList1.AddItem "0x68   h"
CharList1.AddItem "0x69   i"
CharList1.AddItem "0x6A   j"
CharList1.AddItem "0x6B   k"
CharList1.AddItem "0x6C   l"
CharList1.AddItem "0x6D   m"
CharList1.AddItem "0x6E   n"
CharList1.AddItem "0x6F   o"
CharList1.AddItem "0x70   p"
CharList1.AddItem "0x71   q"
CharList1.AddItem "0x72   r"
CharList1.AddItem "0x73   s"
CharList1.AddItem "0x74   t"
CharList1.AddItem "0x75   u"
CharList1.AddItem "0x76   v"
CharList1.AddItem "0x77   w"
CharList1.AddItem "0x78   x"
CharList1.AddItem "0x79   y"
CharList1.AddItem "0x7A   z"
CharList1.AddItem "0x7B   {"
CharList1.AddItem "0x7C   |"
CharList1.AddItem "0x7D   }"
CharList1.AddItem "0x7E   ~"
CharList1.AddItem "0x7F"
CharList1.AddItem "0x80"
CharList1.AddItem "0x81"
CharList1.AddItem "0x82"
CharList1.AddItem "0x83"
CharList1.AddItem "0x84"
CharList1.AddItem "0x85"
CharList1.AddItem "0x86"
CharList1.AddItem "0x87"
CharList1.AddItem "0x88"
CharList1.AddItem "0x89"
CharList1.AddItem "0x8A"
CharList1.AddItem "0x8B"
CharList1.AddItem "0x8C"
CharList1.AddItem "0x8D"
CharList1.AddItem "0x8E"
CharList1.AddItem "0x8F"
CharList1.AddItem "0x90"
CharList1.AddItem "0x91"
CharList1.AddItem "0x92"
CharList1.AddItem "0x93"
CharList1.AddItem "0x94"
CharList1.AddItem "0x95"
CharList1.AddItem "0x96"
CharList1.AddItem "0x97"
CharList1.AddItem "0x98"
CharList1.AddItem "0x99"
CharList1.AddItem "0x9A"
CharList1.AddItem "0x9B"
CharList1.AddItem "0x9C"
CharList1.AddItem "0x9D"
CharList1.AddItem "0x9E"
CharList1.AddItem "0x9F"
CharList1.AddItem "0xA0"
CharList1.AddItem "0xA1"
CharList1.AddItem "0xA2"
CharList1.AddItem "0xA3"
CharList1.AddItem "0xA4"
CharList1.AddItem "0xA5"
CharList1.AddItem "0xA6"
CharList1.AddItem "0xA7"
CharList1.AddItem "0xA8"
CharList1.AddItem "0xA9"
CharList1.AddItem "0xAA"
CharList1.AddItem "0xAB"
CharList1.AddItem "0xAC"
CharList1.AddItem "0xAD"
CharList1.AddItem "0xAE"
CharList1.AddItem "0xAF"
CharList1.AddItem "0xB0"
CharList1.AddItem "0xB1"
CharList1.AddItem "0xB2"
CharList1.AddItem "0xB3"
CharList1.AddItem "0xB4"
CharList1.AddItem "0xB5"
CharList1.AddItem "0xB6"
CharList1.AddItem "0xB7"
CharList1.AddItem "0xB8"
CharList1.AddItem "0xB9"
CharList1.AddItem "0xBA"
CharList1.AddItem "0xBB"
CharList1.AddItem "0xBC"
CharList1.AddItem "0xBD"
CharList1.AddItem "0xBE"
CharList1.AddItem "0xBF"
CharList1.AddItem "0xC0"
CharList1.AddItem "0xC1"
CharList1.AddItem "0xC2"
CharList1.AddItem "0xC3"
CharList1.AddItem "0xC4"
CharList1.AddItem "0xC5"
CharList1.AddItem "0xC6"
CharList1.AddItem "0xC7"
CharList1.AddItem "0xC8"
CharList1.AddItem "0xC9"
CharList1.AddItem "0xCA"
CharList1.AddItem "0xCB"
CharList1.AddItem "0xCC"
CharList1.AddItem "0xCD"
CharList1.AddItem "0xCE"
CharList1.AddItem "0xCF"
CharList1.AddItem "0xD0"
CharList1.AddItem "0xD1"
CharList1.AddItem "0xD2"
CharList1.AddItem "0xD3"
CharList1.AddItem "0xD4"
CharList1.AddItem "0xD5"
CharList1.AddItem "0xD6"
CharList1.AddItem "0xD7"
CharList1.AddItem "0xD8"
CharList1.AddItem "0xD9"
CharList1.AddItem "0xDA"
CharList1.AddItem "0xDB"
CharList1.AddItem "0xDC"
CharList1.AddItem "0xDD"
CharList1.AddItem "0xDE"
CharList1.AddItem "0xDF"
CharList1.AddItem "0xE0"
CharList1.AddItem "0xE1"
CharList1.AddItem "0xE2"
CharList1.AddItem "0xE3"
CharList1.AddItem "0xE4"
CharList1.AddItem "0xE5"
CharList1.AddItem "0xE6"
CharList1.AddItem "0xE7"
CharList1.AddItem "0xE8"
CharList1.AddItem "0xE9"
CharList1.AddItem "0xEA"
CharList1.AddItem "0xEB"
CharList1.AddItem "0xEC"
CharList1.AddItem "0xED"
CharList1.AddItem "0xEE"
CharList1.AddItem "0xEF"
CharList1.AddItem "0xF0"
CharList1.AddItem "0xF1"
CharList1.AddItem "0xF2"
CharList1.AddItem "0xF3"
CharList1.AddItem "0xF4"
CharList1.AddItem "0xF5"
CharList1.AddItem "0xF6"
CharList1.AddItem "0xF7"
CharList1.AddItem "0xF8"
CharList1.AddItem "0xF9"
CharList1.AddItem "0xFA"
CharList1.AddItem "0xFB"
CharList1.AddItem "0xFC"
CharList1.AddItem "0xFD"
CharList1.AddItem "0xFE"
CharList1.AddItem "0xFF"
'  !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}

CharList1.ListIndex = 0
End Sub
