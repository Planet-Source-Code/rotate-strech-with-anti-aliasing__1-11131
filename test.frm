VERSION 5.00
Begin VB.Form frm_rotate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotate a bitmap"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   526
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      Height          =   555
      Left            =   1260
      Picture         =   "test.frx":0000
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   18
      Top             =   480
      Width           =   975
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1995
      Left            =   5520
      Max             =   10
      Min             =   1
      TabIndex        =   11
      Top             =   5580
      Value           =   1
      Width           =   1035
   End
   Begin VB.CommandButton Command5 
      Caption         =   "strange(sine)"
      Height          =   435
      Left            =   4500
      TabIndex        =   10
      ToolTipText     =   "You can create strange effects by using something else instead of atan"
      Top             =   120
      Width           =   1035
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      Height          =   2040
      Left            =   60
      Picture         =   "test.frx":17FA
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   142
      TabIndex        =   9
      Top             =   1080
      Width           =   2190
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   60
      Picture         =   "test.frx":F4EC
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   8
      Top             =   480
      Width           =   555
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Normal"
      Height          =   435
      Left            =   5460
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "standard"
      Height          =   435
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AA"
      Height          =   435
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AA"
      Height          =   735
      Left            =   5460
      TabIndex        =   4
      Top             =   3900
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   3675
      Left            =   60
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   3
      Top             =   3900
      Width           =   5355
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3195
      LargeChange     =   45
      Left            =   5820
      Max             =   360
      TabIndex        =   2
      Top             =   600
      Width           =   795
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   2340
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   660
      Picture         =   "test.frx":1012E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label6 
      Caption         =   "Code by Twan van Laarhoven (twanvl@hotmail.com)"
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   7680
      Width           =   6795
   End
   Begin VB.Label Label5 
      Caption         =   "Select a bitmap to rotate or strech"
      Height          =   495
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   1995
   End
   Begin VB.Label Label4 
      Caption         =   "1x"
      Height          =   255
      Left            =   5640
      TabIndex        =   15
      Top             =   5340
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Angle"
      Height          =   255
      Left            =   5820
      TabIndex        =   14
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   195
      Left            =   5880
      TabIndex        =   13
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Scale"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   5160
      Width           =   975
   End
End
Attribute VB_Name = "frm_rotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selectedBox As PictureBox
Dim x As Integer

Private Sub Command1_Click()
p1dc = selectedBox.hdc
p3dc = Picture3.hdc
Picture3.Cls
sc = VScroll2.Value
For x = 0 To selectedBox.ScaleWidth * sc
For y = 0 To selectedBox.ScaleHeight * sc
    col = pPoint(x / sc, y / sc, p1dc)
    'col = GetPixel(p1dc, Int(x / 4), Int(y / 4))
    SetPixel p3dc, x, y, col
Next
Next
End Sub

Private Sub Command2_Click()
Picture2.Cls
bmp_rotate2 selectedBox, Picture2, VScroll1.Value * Trans

End Sub

Private Sub Command3_Click()
Picture2.Cls
bmp_rotate selectedBox, Picture2, VScroll1.Value * Trans

End Sub

Private Sub Command4_Click()
p1dc = selectedBox.hdc
p3dc = Picture3.hdc
Picture3.Cls
sc = VScroll2.Value
For x = 0 To selectedBox.ScaleWidth * sc
For y = 0 To selectedBox.ScaleHeight * sc
    'col = pPoint(x / 4, y / 4, p1dc)
    col = GetPixel(p1dc, Int(x / sc), Int(y / sc))
    SetPixel p3dc, x, y, col
Next
Next
End Sub

Private Sub Command5_Click()
Picture2.Cls
bmp_rotate3 selectedBox, Picture2, VScroll1.Value * Trans

End Sub

Private Sub Form_Load()
Set selectedBox = Picture1
End Sub

Private Sub Picture1_Click()
Set selectedBox = Picture1
End Sub

Private Sub Picture4_Click()
Set selectedBox = Picture4
End Sub

Private Sub Picture5_Click()
Set selectedBox = Picture5
End Sub

Private Sub Picture6_Click()
Set selectedBox = Picture6
End Sub

Private Sub VScroll1_Change()
Label2.Caption = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Label2.Caption = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
Label4.Caption = VScroll2.Value & "x"
End Sub

Private Sub VScroll2_Scroll()
Label4.Caption = VScroll2.Value & "x"
End Sub
