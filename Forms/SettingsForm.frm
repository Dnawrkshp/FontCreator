VERSION 5.00
Begin VB.Form SettingsForm 
   Caption         =   "Settings"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelSett 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton OkaySett 
      Caption         =   "Okay"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   2055
   End
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelSett_Click()
Unload SettingsForm
End Sub

Private Sub Form_Load()

'Dim Value As Integer

'Value = Val(Replace(LoadSett("C:\Users\Dan G\Documents\Visual Basic 6\Font Creator\sett.inf", "inv"), "InvertColors  ", ""))
'If Boolean1 <> 0 Or Boolean1 <> 1 Then: MsgBox "ERROR: Invalid value for Invert Colors. File is either corrupt or does not exist.": Exit Sub
'PS2FontCreator.SettingsForm.InvertColors.Value = Value


End Sub

Private Sub OkaySett_Click()
'StoreSett ("C:\Users\Dan G\Documents\Visual Basic 6\Font Creator\sett.inf")
'PS2FontCreator.Ed.Update_Invert
Unload SettingsForm
End Sub
