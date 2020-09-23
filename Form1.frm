VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Sketch !"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "About"
      Height          =   495
      Left            =   14280
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdSketch2 
      Caption         =   "Auto Sketch 2"
      Height          =   495
      Left            =   10560
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.HScrollBar hsDif 
      Height          =   375
      Left            =   3960
      Max             =   125
      Min             =   1
      TabIndex        =   8
      Top             =   120
      Value           =   10
      Width           =   1695
   End
   Begin VB.CommandButton cmdSketch 
      Caption         =   "Auto Sketch 1"
      Height          =   495
      Left            =   8280
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Clear"
      Height          =   495
      Left            =   13320
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   11880
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenFit 
      Caption         =   "Open and Fit"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7680
      ScaleHeight     =   122
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   720
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   120
      MouseIcon       =   "Form1.frx":0442
      ScaleHeight     =   681
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   0
      Top             =   720
      Width           =   7395
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblDif 
      Caption         =   "Dif=10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblColor2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblColor1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is a rewrite for the code that appeared in Planet Source Code
'http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=55153&lngWId=1

'The original code was very poorly written and as a result was painfully slow,
'difficult to use, understand and generally substandard

'Original Programmer : Mahdi Shakouri rad, Mahdi_Rad@yahoo.com

'Changes:
'1. Use drag/drop on the pictureboxes so that the results can easily be "scrolled"
'onto the screen for analysis. Some times objects will not fit on the available
'real estate - have some respect for your potential user!
'2. Use "beginner" APIs - GetPixel, SetPixel for speed
'3. Place RGB calcs in loops - do not exit loop to an outside subprocedure when
'parsing a bitmap
'4. Use industry standard notation for variables. Precede a string with "s",
'a Long with "l" etc.
'5. Indent and use spacing - this is the most valuable productivity tool that you
'will ever need to learn, so start now!
'6. Use Long variables unless Singles are required. RAM "savings" are no longer
'necessary, but the speed enhancements are!
'7. Develop discipline - Stand out - do not be proud of "lazy" code.
Dim sPicW0, sPicH0, sWidth, sHeight As Single
Dim XPos As Long, YPos As Long

Private Sub cmdCls_Click()
  Picture2.Cls
  Picture2.Picture = LoadPicture()
End Sub

Private Sub cmdHelp_Click()
  Form3.Show
End Sub

Private Sub cmdOpen_Click()
  CommonDialog1.CancelError = True
  On Error GoTo ja
  
  CommonDialog1.Filter = "Image|*.bmp;*.gif;*.jpg"
  CommonDialog1.ShowOpen
  ' open picture in Picture2 and then fit itin Picture1!
  Picture1.Picture = LoadPicture(CommonDialog1.FileName)
  Picture2.Picture = LoadPicture(CommonDialog1.FileName)
  
  sPicW0 = Picture2.Width
  sPicH0 = Picture2.Height
  
  sWidth = 493
  sHeight = 681
  
  If sPicW0 < sWidth Then sWidth = sPicW0
  If sPicH0 < sHeight Then sHeight = sPicH0
  
  Picture1.Width = sWidth
  Picture1.Height = sHeight
  
  Picture2.Width = Picture1.Width
  Picture2.Height = Picture1.Height
  
  'Picture1.Picture = LoadPicture()
  'Picture1.PaintPicture Picture2.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0, sPicW0, sPicH0, vbSrcCopy
  Picture2.Picture = LoadPicture()
  
ja:
End Sub

Private Sub cmdOpenFit_Click()
  CommonDialog1.CancelError = True
  On Error GoTo ja
  
  CommonDialog1.Filter = "Image|*.bmp;*.gif;*.jpg"
  CommonDialog1.ShowOpen
  ' open picture in Picture2 and then fit itin Picture1!
  Picture1.Picture = LoadPicture(CommonDialog1.FileName)
  Picture2.Picture = LoadPicture(CommonDialog1.FileName)
  ' Fit
  sPicW0 = Picture2.Width
  sPicH0 = Picture2.Height
  
  sWidth = 493
  sHeight = 681
  
  If sPicW0 / sPicH0 > sWidth / sHeight Then
    ' resize based on width
    Picture1.Width = sWidth
    Picture1.Height = sWidth * sPicH0 / sPicW0
    Else
    ' resize based on height
    Picture1.Height = sHeight
    Picture1.Width = sHeight * sPicW0 / sPicH0
  End If
  Picture2.Width = Picture1.Width
  Picture2.Height = Picture1.Height
  
  Picture1.Picture = LoadPicture()
  Picture1.PaintPicture Picture2.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0, sPicW0, sPicH0, vbSrcCopy
  Picture2.Picture = LoadPicture()
  
ja:
End Sub

Private Sub cmdSave_Click()
  Dim FName As String
  CommonDialog1.CancelError = True
  On Error GoTo ja
  CommonDialog1.Filter = "*.jpg"
  CommonDialog1.ShowSave
  FName = CommonDialog1.FileName
  If Right$(FName, 4) <> ".jpg" Then FName = FName + ".jpg"
  SavePicture Picture2.Image, FName
ja:
End Sub


Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)

With Source
    .Move X - XPos, Y - YPos
    .Drag vbEndDrag
End With

End Sub

Private Sub Form_Load()

With Me
    .Top = 0
    .Left = 0
    .Width = Screen.Width
    .Height = Screen.Height - 450
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload Me
End

End Sub

Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)

With Source
    .Move Picture1.Left + X - XPos, Picture1.Top + Y - YPos
    .Drag vbEndDrag
End With

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
If Button = vbLeftButton Then
    XPos = X
    YPos = Y
    Picture1.Drag vbBeginDrag
End If

  If Button = vbRightButton Then
    lblColor2.BackColor = Picture1.Point(X, Y)
  End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If X < 0 Or X >= Picture1.ScaleWidth Or Y < 0 Or Y >= Picture1.ScaleHeight Then Exit Sub
  lblColor1.BackColor = Picture1.Point(X, Y)
  
  If Button = vbLeftButton Then
    Picture2.ForeColor = lblColor1.BackColor
    Picture2.PSet (X, Y)
  End If

End Sub

Private Sub Picture2_DragDrop(Source As Control, X As Single, Y As Single)

With Source
    .Move Picture2.Left + X - XPos, Picture2.Top + Y - YPos
    .Drag vbEndDrag
End With

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    XPos = X
    YPos = Y
    Picture2.Drag vbBeginDrag
End If

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = vbLeftButton Then
    Picture2.ForeColor = lblColor2.BackColor
    Picture2.PSet (X, Y)
  End If

End Sub

Private Function BW(c As Long) As Integer
Dim R, G, B As Integer
  R = c Mod 256
  G = (c \ 256) Mod 256
  B = (c \ 256 \ 256) Mod 256
  BW = (R + G + B) / 3
End Function

Private Sub cmdSketch_Click()
'Changes:
'1. Use "beginner" APIs - GetPixel, SetPixel for speed
'2. Place RGB calcs in loops - do not exit loop to an outside subprocedure when
'parsing a bitmap
'3. Use industry standard notation for variables. Precede a string with "s",
'a Long with "l" etc.
'4. Indent and use spacing - this is the most valuable productivity tool that you
'will ever need to learn, so start now!
'5. Use Long variables unless Singles are required. RAM "savings" are no longer
'necessary, but the speed enhancements are!
'6. Develop discipline - Stand out - do not be proud of "lazy" code.

Dim lX As Long, lY As Long, R As Long, G As Long, B As Long
Dim lPoint As Long, lP2 As Long, lC1 As Long, lC2 As Long
Dim lTotal, lDone As Long, lDiff As Long


Picture2.Cls
lDiff = hsDif.Value 'Always query objects before going into loop (again, watch overheads)

With Picture1
    lTotal = .Width + .Height
    lDone = 0
    ' in lY direction :
    For lX = 0 To .Width - 1
        lDone = lDone + 1
        lblPercent = Str$(Int(100 * lDone / lTotal)) + "%"
        DoEvents
        For lY = 0 To .Height - 2
            lPoint = GetPixel(.hdc, lX, lY)
            R = lPoint Mod 256
            G = (lPoint \ 256) Mod 256
            B = (lPoint \ 256 \ 256) Mod 256
            lC1 = (R + G + B) / 3
            
            lPoint = GetPixel(.hdc, lX, lY + 1)
            R = lPoint Mod 256
            G = (lPoint \ 256) Mod 256
            B = (lPoint \ 256 \ 256) Mod 256
            lC2 = (R + G + B) / 3
            
            If Abs(lC1 - lC2) > lDiff Then
                SetPixel Picture2.hdc, lX, lY, vbBlack
            End If
        Next lY
    Next lX

    'in lX direction :
    For lY = 0 To .Height - 1
        lDone = lDone + 1
        lblPercent = Str$(Int(100 * lDone / lTotal)) + "%"
        DoEvents
        For lX = 0 To .Width - 2
            lPoint = GetPixel(.hdc, lX + 1, lY)
            'Never leave loop to a sub procedure - overhead is costly
            R = lPoint Mod 256
            G = (lPoint \ 256) Mod 256
            B = (lPoint \ 256 \ 256) Mod 256
            lC1 = (R + G + B) / 3
            
            lPoint = GetPixel(.hdc, lX, lY + 1)
            R = lPoint Mod 256
            G = (lPoint \ 256) Mod 256
            B = (lPoint \ 256 \ 256) Mod 256
            lC2 = (R + G + B) / 3
            
            If Abs(lC1 - lC2) > lDiff Then SetPixel Picture2.hdc, lX, lY, vbBlack
        Next lX
    Next lY
    Picture2.Refresh 'Refresh a end of routine - test speed by placing 'refresh'
                        'in loop -  after SetPixel line
End With

lblPercent = "%"
End Sub
Private Sub cmdSketch2_Click()
Dim lX As Long, lY As Long, R As Long, G As Long, B As Long
Dim lDiff As Long
Dim lC1 As Long, lC2 As Long
Dim lPoint As Long, lPoint1 As Long
Dim lTotal, lDone As Long

Picture2.Cls

lTotal = Picture1.Width
lDone = 0
lDiff = hsDif.Value 'Always query objects before going into loop (again, watch overheads)

' in lY direction :
With Picture1
    lPoint = GetPixel(.hdc, 0, 0)
    For lX = 0 To .Width - 1
        lDone = lDone + 1
        lblPercent = Str$(Int(100 * lDone / lTotal)) + "%"
        DoEvents
        For lY = 0 To .Height - 2
            lPoint1 = GetPixel(.hdc, lX, lY)
            R = lPoint1 Mod 256
            G = (lPoint1 \ 256) Mod 256
            B = (lPoint1 \ 256 \ 256) Mod 256
            lC1 = (R + G + B) / 3
            If lC1 > lDiff Then
                lPoint = vbWhite
            Else
                lPoint = vbBlack
            End If
            
            SetPixel Picture2.hdc, lX, lY, lPoint
        Next lY
    Next lX
End With

Picture2.Refresh '
lblPercent = "%"

End Sub

Private Sub hsDif_Change()
  lblDif = "Dif=" & Str$(hsDif.Value)
End Sub

