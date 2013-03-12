VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ArrayName 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Enter An Array Name"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog OpenDialog1 
      Left            =   8880
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton SpaceBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox SpaceLocBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00BCA93F&
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   5655
   End
   Begin VB.CommandButton BrowseFont 
      Caption         =   "Browse"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox FontLocBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00BCA93F&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   5655
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3360
      Width           =   3495
   End
   Begin VB.CommandButton Okay 
      Caption         =   "Okay"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox SpaceBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00BCA93F&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "LDSpace"
      Top             =   1080
      Width           =   7095
   End
   Begin VB.TextBox FontBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00BCA93F&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "LDFont"
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Spacing save location"
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   7095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Font save location"
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the array name for the font file. Default: LDv3Font"
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
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the array name for the spacing file. Default: LDv3Space"
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7095
   End
End
Attribute VB_Name = "ArrayName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BrowseFont_Click()
Ed.CD.Filter = "C Array's (*.c)|*.c|All files (*.*)|*.*"
Ed.CD.DialogTitle = "Select File"
Ed.CD.Action = 2
FontLocBox.text = Ed.CD.filename
End Sub

Private Sub Cancel_Click()
ReturnVal = False
Unload ArrayName
End Sub

Private Sub Form_Load()
If FontArray <> "" Then: Fontbox.text = FontArray
If SpaceArray <> "" Then: SpaceBox.text = SpaceArray
FontLocBox.text = FontSaveLoc
SpaceLocBox.text = SpaceSaveLoc
End Sub

Private Sub Form_Resize()
FormFade Me, 0, 200, 255
End Sub

Private Sub Form_Terminate()
ReturnVal = False
End Sub

Private Sub Okay_Click()
FontArray = Fontbox.text
SpaceArray = SpaceBox.text
FontSaveLoc = FontLocBox.text
SpaceSaveLoc = SpaceLocBox.text
ReturnVal = True
Unload ArrayName
End Sub

Private Sub SpaceBrowse_Click()
Ed.CD.Filter = "C Array's (*.c)|*.c|All files (*.*)|*.*"
Ed.CD.DialogTitle = "Select File"
Ed.CD.ShowOpen
SpaceLocBox.text = Ed.CD.filename
End Sub
