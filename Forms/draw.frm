VERSION 5.00
Begin VB.Form Toolbox 
   Caption         =   "Toolbox Draw"
   ClientHeight    =   5955
   ClientLeft      =   8550
   ClientTop       =   5085
   ClientWidth     =   5640
   LinkTopic       =   "Form2"
   ScaleHeight     =   5955
   ScaleWidth      =   5640
   Begin VB.PictureBox fontPict 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawStyle       =   6  'Inside Solid
      ForeColor       =   &H80000008&
      Height          =   5355
      Left            =   90
      ScaleHeight     =   18
      ScaleMode       =   0  'User
      ScaleWidth      =   18
      TabIndex        =   0
      Top             =   45
      Width           =   5355
   End
End
Attribute VB_Name = "Toolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    fontPict.DrawWidth = 20
    DRAWGRID
End Sub


Sub DRAWGRID()
    fontPict.DrawWidth = 1
    For GX = 1 To 17
            fontPict.Line (GX, 0)-(GX, 18), &H666666
    Next GX
    
    
    For GY = 1 To 17
            fontPict.Line (0, GY)-(18, GY), &H666666
    Next GY
    fontPict.DrawWidth = 20
End Sub
