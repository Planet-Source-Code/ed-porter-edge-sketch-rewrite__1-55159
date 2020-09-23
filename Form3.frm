VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Sketch"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   5340
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "H E L P"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   1890
      Left            =   3600
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   6
      Top             =   1560
      Width           =   1560
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   1920
      Picture         =   "Form3.frx":237F
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1920
      Left            =   240
      Picture         =   "Form3.frx":5794
      ScaleHeight     =   1860
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   1560
      Width           =   1560
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Sketch 2"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Sketch 1"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Author"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Mahdi_Rad@yahoo.com"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Programmer : Mahdi Shakouri rad"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Ver 1.00  -  July 2004"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Sketch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHelp_Click()
  Form2.Show
End Sub

