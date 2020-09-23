VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Help"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form2"
   ScaleHeight     =   6975
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  txtHelp = "This application draws sketch from your photos automatically or manually!" + vbCrLf + _
           "The original photo will be shown in left picture box, The result in Right" + vbCrLf + _
           vbCrLf + _
           "-If you move the mouse in 'Left picture box' when the 'Left button' is pressed" + vbCrLf + _
           " its trace will be drawn in 'Right picture box'" + vbCrLf + _
           " Also the original color will be used for drawing." + vbCrLf + _
           vbCrLf + _
           "-If you move the mouse in 'Left Picture Box',  the corresponding" + vbCrLf + _
           " color will be shown in upper right box of this picture box." + vbCrLf + _
           " Now, if you Right Click, this color will be shown in upper left " + vbCrLf + _
           " box of Right Picture Box." + vbCrLf + _
           vbCrLf + _
           "-If you move the mouse in 'Right picture box' when the 'Left button' is pressed " + vbCrLf + _
           " its trace will be drawn, in this case the color which is shown in" + vbCrLf + _
           " upper left box of this picture box will be used." + vbCrLf + _
          vbCrLf + _
           "-The horizontal scroll bar is used for Auto Sketching" + vbCrLf + _
           " I suggest to use Dif=10 for first try, if you use 'Auto Sketch 1'" + vbCrLf + _
           " and Dif=100 if you use 'Auto Sketch 2'" + vbCrLf + _
          vbCrLf + _
           "-'Open and Fit' loads a picture in origin picture box (Left picture box)" + vbCrLf + _
           " And fits it! 'Open' just loads without fitting!" + vbCrLf + _
           " 'Save' , 'Clear' & 'Help' don't need any comment! You better know what they do ..." + vbCrLf + _
           vbCrLf + _
           "<<<  Programmer : Mahdi Shakouri rad, Mahdi_Rad@yahoo.com , July 2004 >>>"
          

End Sub
