Attribute VB_Name = "Declarations"
Option Explicit
'Declarations module for PS2FontCreator/Font Creator
Global dither

'The directory that the Font Creator has been opened in. Will be where temp.bin and temp SPACE.bin will be stored
Global AppDirectory As String

'This is the file location of whatever your making
Global GlobalFileName As String

'The space file location for font
Global GlobalSpaceName As String

'Tool in use selected from Toolbox
Global ToolUse As Integer

'Settings constant strings in an array
Global SettArray(1) As String

'Settings, key shortcuts on or off
Global KeyShort As Boolean

'Spacing settings; min is left most marker, max is right most marker
Global min As Integer
Global max As Integer

'If true, will stop CharList1_Change from updating grid. Increases speed of import font
Global Import As Boolean

'Global space and font file array names
Global FontArray As String
Global SpaceArray As String
Global FontSaveLoc As String
Global SpaceSaveLoc As String

'Return value for functions/forms
Global ReturnVal As Boolean

'Color constants
Global Const White = vbWhite
Global Const Black = &H222222
Global Const Grey = &HCCCCCC

'Constant mode for PFC segment saving
Global Const F_Off = 0
Global Const S_Off = 1
Global Const FName_Off = 2
Global Const FCName_Off = 3
Global Const SCName_Off = 4
Global Const FName = 5
Global Const FCName = 6
Global Const SCName = 7
Global Const F_Data = 8
Global Const S_Data = 9


Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'Declaration For DataTree
Type Letter
    pixel(325) As Boolean
    Size As Integer
    max As Integer
    min As Integer
End Type
Global data(100) As Letter


'---------------------------------------------------------------------------------------
' Procedure : FormFade
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Paints a Gradient on a form
'---------------------------------------------------------------------------------------
Sub FormFade(Frm As Form, Rc As Long, Gc As Long, bc As Long)
        'This code works best when called in the paint event
          Dim t1, t2, t3
10        t1 = Rc / 255
20        t2 = Gc / 255
30        t3 = bc / 255
          Dim intLoop As Integer
50        On Error Resume Next
60        With Frm
70            .DrawStyle = vbInsideSolid
80            .FillStyle = 7
              .DrawMode = vbCopyPen
90            .ScaleMode = vbPixels
100           .DrawWidth = 2
110           .ScaleHeight = 256
120           .AutoRedraw = True
130           .ClipControls = False
140       End With
            
150       For intLoop = 0 To 255
160           Frm.Line (-1, intLoop)-(Frm.Width, intLoop - 1), RGB(t1 * intLoop, t2 * intLoop, t3 * intLoop), B    'Draw boxes With specified color of loop
170       Next intLoop

End Sub

