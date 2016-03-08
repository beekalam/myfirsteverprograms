VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memory Helper 1.1"
   ClientHeight    =   5655
   ClientLeft      =   4320
   ClientTop       =   1935
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFF80&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6555
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5760
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   2880
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   0
      Width           =   6555
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "wrong:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Right:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   3840
      Width           =   615
   End
   Begin VB.Menu option_menue 
      Caption         =   "Option"
      Begin VB.Menu Number_menue 
         Caption         =   "Number"
         Begin VB.Menu menue4 
            Caption         =   "4-Digits"
         End
         Begin VB.Menu menue5 
            Caption         =   "5-Digits"
         End
         Begin VB.Menu menue6 
            Caption         =   "6-Digits"
         End
         Begin VB.Menu menue7 
            Caption         =   "7-Digits"
         End
         Begin VB.Menu menue8 
            Caption         =   "8-Digits"
         End
         Begin VB.Menu menue9 
            Caption         =   "9-Digits"
         End
         Begin VB.Menu menue10 
            Caption         =   "10-Digits"
         End
      End
      Begin VB.Menu Time_menue 
         Caption         =   "Time"
         Begin VB.Menu time1 
            Caption         =   "1-Second"
         End
         Begin VB.Menu time2 
            Caption         =   "2-Seconds"
         End
         Begin VB.Menu time3 
            Caption         =   "3-Seconds"
         End
         Begin VB.Menu time4 
            Caption         =   "4-Seconds"
         End
         Begin VB.Menu time5 
            Caption         =   "5-Seconds"
         End
         Begin VB.Menu time6 
            Caption         =   "6-Seconds"
         End
         Begin VB.Menu time7 
            Caption         =   "7-Seconds"
         End
      End
      Begin VB.Menu Font_menue 
         Caption         =   "Font"
         Begin VB.Menu font_name_menue 
            Caption         =   "Name"
            Begin VB.Menu arial_menue 
               Caption         =   "Arial"
            End
            Begin VB.Menu Comic_menue 
               Caption         =   "Comic Sans"
            End
            Begin VB.Menu Courier 
               Caption         =   "Courier"
            End
            Begin VB.Menu Courier_New 
               Caption         =   "Courier New"
            End
            Begin VB.Menu Times 
               Caption         =   "Times New Roman"
            End
         End
         Begin VB.Menu size 
            Caption         =   "size"
            Begin VB.Menu cc 
               Caption         =   "24"
            End
            Begin VB.Menu aa 
               Caption         =   "36"
            End
            Begin VB.Menu bb 
               Caption         =   "48"
            End
         End
      End
      Begin VB.Menu Exit_menue 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help_menue 
      Caption         =   "Help"
      Begin VB.Menu about_menue 
         Caption         =   "about"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Double
Dim b As Double
Dim right As Integer
Dim wrong As Integer
Dim score As Integer
Dim today As Variant
Dim today1 As Date
Dim today2 As Date
Dim today3 As Date
Dim t1 As Double
Dim t2 As Double
Dim t3 As Double
Dim t4 As Double
Dim t5 As Double
Dim t6 As Double
Dim t7 As Double
Dim num1 As Double
Dim num2 As Double



Private Sub aa_Click()
Picture1.Font.size = 36
aa.Checked = True
bb.Checked = False
cc.Checked = False
End Sub

Private Sub about_menue_Click()
frmAbout.Show
End Sub

Private Sub arial_menue_Click()
Picture1.Font.Name = "arial"
arial_menue.Checked = True
Comic_menue.Checked = False
Courier.Checked = False
Times.Checked = False
End Sub

Private Sub bb_Click()
Picture1.Font.size = 48
bb.Checked = True
aa.Checked = False
cc.Checked = False
End Sub

Private Sub cc_Click()
Picture1.Font.size = 24
aa.Checked = False
bb.Checked = False
cc.Checked = True
End Sub

Private Sub Comic_menue_Click()
Picture1.Font.Name = "Comic Sans MS"
Comic_menue.Checked = True
Courier.Checked = False
arial_menue.Checked = False
Times.Checked = False
End Sub

Private Sub Courier_Click()
Picture1.Font.Name = "Courier"
Courier_New.Checked = False
arial_menue.Checked = False
Comic_menue.Checked = False
Courier.Checked = True
Times.Checked = False
End Sub

Private Sub Courier_New_Click()
Picture1.Font.Name = "Courier_New"
Courier_New.Checked = True
arial_menue.Checked = False
Comic_menue.Checked = False
Courier.Checked = False
Times.Checked = False
End Sub

Private Sub Exit_menue_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Enabled = False
right = 1
wrong = 1
score = 1
Call menue6_Click
Call time2_Click
Call Times_Click
Call bb_Click
End Sub

Private Sub menue10_Click()
menue4.Checked = False
menue5.Checked = False
menue6.Checked = False
menue7.Checked = False
menue8.Checked = False
menue9.Checked = False
menue10.Checked = True
If menue10.Checked = True Then
    num1 = 1000000000
    num2 = 9999999999#
End If
End Sub

Private Sub menue4_Click()
menue4.Checked = True
menue5.Checked = False
menue6.Checked = False
menue7.Checked = False
menue8.Checked = False
menue9.Checked = False
menue10.Checked = False
If menue4.Checked = True Then
    num1 = 1000
    num2 = 9999
End If
End Sub

Private Sub menue5_Click()
menue4.Checked = False
menue5.Checked = True
menue6.Checked = False
menue7.Checked = False
menue8.Checked = False
menue9.Checked = False
menue10.Checked = False
If menue5.Checked = True Then
    num1 = 10000
    num2 = 99999
End If
End Sub

Private Sub menue6_Click()
menue4.Checked = False
menue5.Checked = False
menue6.Checked = True
menue7.Checked = False
menue8.Checked = False
menue9.Checked = False
menue10.Checked = False
If menue6.Checked = True Then
    num1 = 100000
    num2 = 999999
End If
End Sub

Private Sub menue7_Click()
menue4.Checked = False
menue5.Checked = False
menue6.Checked = False
menue7.Checked = True
menue8.Checked = False
menue9.Checked = False
menue10.Checked = False
If menue7.Checked = True Then
    num1 = 1000000
    num2 = 9999999
End If
End Sub

Private Sub menue8_Click()
menue4.Checked = False
menue5.Checked = False
menue6.Checked = False
menue7.Checked = False
menue8.Checked = True
menue9.Checked = False
menue10.Checked = False
If menue8.Checked = True Then
    num1 = 10000000
    num2 = 99999999
End If
End Sub

Private Sub menue9_Click()
menue4.Checked = False
menue5.Checked = False
menue6.Checked = False
menue7.Checked = False
menue8.Checked = False
menue9.Checked = True
menue10.Checked = False
If menue9.Checked = True Then
    num1 = 100000000
    num2 = 999999999
End If
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
Picture1.Cls
Text1.Enabled = False
If KeyAscii = vbKeyReturn Then
Randomize Timer
1: a = Int(num2 * Rnd) + num1
If a > num2 Then
 GoTo 1
End If

 Picture1.Print "  "; a;
 
today1 = Label4.Caption

today3 = today1 + today2
Label9.Caption = Format(today3, "hh:mm:ss")
 
 End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 0 And KeyAscii <= 9) Or KeyAscii = vbKeyBack Then
        Beep
        ElseIf KeyAscii = vbKeyReturn Then
            b = Val(Text1.Text)
            If a = b Then
            
                Label3.Caption = "right"
                right = right + 1
                Label1.Caption = right
                Text1.Text = ""
                score = (right - wrong)
                Label3.Caption = " RIGHT" + " The  answer is   " + Format(a, "######")
                Label3.BackColor = vbGreen
                Label8.Caption = score
                Picture1.SetFocus
                
            Else
            
                Label2.Caption = "wrong"
                wrong = wrong + 1
                Label2.Caption = wrong
                Text1.Text = ""
                score = (right - wrong)
                Label3.Caption = " WRONG" + " The answer is   " + Format(a, "######")
                Label3.BackColor = vbYellow
                Label8.Caption = score
                Picture1.SetFocus
                
            End If
        
        End If
End Sub

Private Sub time1_Click()
time1.Checked = True
time2.Checked = False
time3.Checked = False
time4.Checked = False
time5.Checked = False
time6.Checked = False
time7.Checked = False
If time1.Checked = True Then
    today2 = "00:00:01"
End If
End Sub

Private Sub time2_Click()
time1.Checked = False
time2.Checked = True
time3.Checked = False
time4.Checked = False
time5.Checked = False
time6.Checked = False
time7.Checked = False
If time2.Checked = True Then
    today2 = "00:00:02"
End If
End Sub

Private Sub time3_Click()
time1.Checked = False
time2.Checked = False
time3.Checked = True
time4.Checked = False
time5.Checked = False
time6.Checked = False
time7.Checked = False
If time3.Checked = True Then
    today2 = "00:00:03"
End If
End Sub

Private Sub time4_Click()
time1.Checked = False
time2.Checked = False
time3.Checked = False
time4.Checked = True
time5.Checked = False
time6.Checked = False
time7.Checked = False
If time4.Checked = True Then
    today2 = "00:00:04"
End If
End Sub

Private Sub time5_Click()
time1.Checked = False
time2.Checked = False
time3.Checked = False
time4.Checked = False
time5.Checked = True
time6.Checked = False
time7.Checked = False
If time5.Checked = True Then
    today2 = "00:00:05"
End If
End Sub

Private Sub time6_Click()
time1.Checked = False
time2.Checked = False
time3.Checked = False
time4.Checked = False
time5.Checked = False
time6.Checked = True
time7.Checked = False
If time6.Checked = True Then
    today2 = "00:00:06"
End If
End Sub

Private Sub time7_Click()
time1.Checked = False
time2.Checked = False
time3.Checked = False
time4.Checked = False
time5.Checked = False
time6.Checked = False
time7.Checked = True
If time7.Checked = True Then
    today2 = "00:00:07"
End If
End Sub

Private Sub Timer1_Timer()
today = Now
Label4.Caption = Format(today, "hh:mm:ss")
If Label4.Caption = Label9.Caption Then
    Picture1.Cls
    Text1.Enabled = True
    Text1.SetFocus
    End If
End Sub

Private Sub Times_Click()
Picture1.Font.Name = "Times New Roman"
Courier_New.Checked = False
arial_menue.Checked = False
Comic_menue.Checked = False
Courier.Checked = False
Times.Checked = True
End Sub
