VERSION 5.00
Begin VB.UserControl UserControl2 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   ScaleHeight     =   300
   ScaleMode       =   0  'User
   ScaleWidth      =   98.571
   Begin VB.Image Up 
      Height          =   300
      Index           =   3
      Left            =   1755
      MousePointer    =   99  'Custom
      Picture         =   "slider2.ctx":0000
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Up 
      Height          =   300
      Index           =   1
      Left            =   0
      MousePointer    =   99  'Custom
      Picture         =   "slider2.ctx":04F2
      Top             =   0
      Width           =   300
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   6.429
      X2              =   87.857
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Image pip 
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   720
      MousePointer    =   9  'Size W E
      Picture         =   "slider2.ctx":09E4
      Top             =   0
      Width           =   45
   End
   Begin VB.Shape Shapex 
      BorderColor     =   &H00BCA93F&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   10
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2085
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Caption As String
Property Get slidervalue() As String
   slidervalue = Caption
End Property
Property Let slidervalue(text As String)
   Caption = pip.Left
   UserControl.Tag = pip.Left
   PropertyChanged
End Property


Private Sub Up_Click(Index As Integer)
    Dim X As Integer
   ' x = Int(Val(Ed.Fscale.slidervalue))
    
        X = Int(Val(dither))
    Select Case Index
        Case 1
            X = X - 1
            dither = X
        Case 3
            X = X + 1
            dither = X
    End Select
    pip.Left = X - (pip.Width / 2)
    Caption = X
    Ed.Timer2.Interval = 1
End Sub

Private Sub UserControl_DragDrop(Source As Control, X As Single, Y As Single)
pip.Left = X - (pip.Width / 2)
Caption = Int(X - (pip.Width / 2))
Ed.Timer2.Interval = 1
End Sub

Private Sub UserControl_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
pip.Top = 0
Caption = Int(X - (pip.Width / 2))
'Ed.Timer1.Interval = 1
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ed.HelpText = "Set tolerance used in dithering algorithm"
End Sub
