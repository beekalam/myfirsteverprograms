VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Imaginative Memory"
   ClientHeight    =   9165
   ClientLeft      =   1170
   ClientTop       =   1380
   ClientWidth     =   11130
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   11130
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture89 
      BackColor       =   &H00FF0000&
      FillColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   11640
      ScaleHeight     =   435
      ScaleWidth      =   3435
      TabIndex        =   95
      Top             =   7560
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Show  Pictures (Advanced) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   94
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Show Pictures (Intermediate)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   93
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show Pictures (Novice)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   92
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show Pictures (Beginner)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   91
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Check Answers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   90
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide Pictures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   89
      Top             =   5520
      Width           =   1455
   End
   Begin VB.PictureBox Picture88 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   10200
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   88
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture87 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   9240
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   87
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture86 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   8280
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   86
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture85 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   7320
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   85
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture84 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   6360
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   84
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture83 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   5400
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   83
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture82 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   4440
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   82
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture81 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   81
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture80 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   80
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture79 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   1560
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   79
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture78 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   78
      Top             =   9000
      Width           =   855
   End
   Begin VB.PictureBox Picture77 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   10200
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   77
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture76 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   9240
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   76
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture75 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   8280
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   75
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture74 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   7320
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   74
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture73 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   6360
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   73
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture72 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   5400
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   72
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture71 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   4440
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   71
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture70 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   70
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture69 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   69
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture68 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   1560
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   68
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture67 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   67
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture66 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   10200
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   66
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture65 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   9240
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   65
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture64 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   8280
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   64
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture63 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   7320
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   63
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture62 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   6360
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   62
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture61 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   5400
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   61
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture60 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   4440
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   60
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture59 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   59
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture58 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   58
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture57 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   1560
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   57
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture56 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   56
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture55 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   10200
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   55
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture54 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   9240
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   54
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture53 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   8280
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   53
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture52 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   7320
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   52
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture51 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   6360
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   51
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture50 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   5400
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   50
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture49 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   4440
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   49
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture48 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   48
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture47 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   47
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture46 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   1560
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   46
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox Picture45 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   45
      Top             =   5400
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   14760
      Top             =   10680
   End
   Begin VB.PictureBox Picture44 
      Height          =   1095
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   44
      Top             =   4080
      Width           =   855
   End
   Begin VB.PictureBox Picture43 
      Height          =   1095
      Left            =   5640
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   43
      Top             =   4080
      Width           =   855
   End
   Begin VB.PictureBox Picture42 
      Height          =   1095
      Left            =   14280
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   42
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture41 
      Height          =   1095
      Left            =   13200
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   41
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture40 
      Height          =   1095
      Left            =   12120
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   40
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture39 
      Height          =   1095
      Left            =   11040
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   39
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture38 
      Height          =   1095
      Left            =   9960
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   38
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture37 
      Height          =   1095
      Left            =   8880
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   37
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture36 
      Height          =   1095
      Left            =   7800
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   36
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture35 
      Height          =   1095
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   35
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture34 
      Height          =   1095
      Left            =   5640
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   34
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture33 
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   33
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture32 
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   32
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture31 
      Height          =   1095
      Left            =   2400
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   31
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture30 
      Height          =   1095
      Left            =   1320
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   30
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture29 
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   29
      Top             =   2760
      Width           =   855
   End
   Begin VB.PictureBox Picture28 
      Height          =   1095
      Left            =   14280
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   28
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture27 
      Height          =   1095
      Left            =   13200
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   27
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture26 
      Height          =   1095
      Left            =   12120
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   26
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture25 
      Height          =   1095
      Left            =   11040
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   25
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture24 
      Height          =   1095
      Left            =   9960
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   24
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture23 
      Height          =   1095
      Left            =   8880
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   23
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture22 
      Height          =   1095
      Left            =   7800
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture21 
      Height          =   1095
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   21
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture20 
      Height          =   1095
      Left            =   5640
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture19 
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture18 
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   18
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture17 
      Height          =   1095
      Left            =   2400
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture16 
      Height          =   1095
      Left            =   1320
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture15 
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture14 
      Height          =   1095
      Left            =   14280
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture13 
      Height          =   1095
      Left            =   13200
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture12 
      Height          =   1095
      Left            =   12120
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture11 
      Height          =   1095
      Left            =   11040
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture10 
      Height          =   1095
      Left            =   9960
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture9 
      Height          =   1095
      Left            =   8880
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture8 
      Height          =   1095
      Left            =   7800
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture7 
      Height          =   1095
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture6 
      Height          =   1095
      Left            =   5640
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture5 
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   2400
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   1320
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Pictures (Expert)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Menu file_menue 
      Caption         =   "File"
      Begin VB.Menu level_menue 
         Caption         =   "Level"
         Begin VB.Menu beginner_menue 
            Caption         =   "Beginner"
         End
         Begin VB.Menu novice_menue 
            Caption         =   "Novice"
         End
         Begin VB.Menu intermediate_menue 
            Caption         =   "Intermediate"
         End
         Begin VB.Menu advanced_menue 
            Caption         =   "Advanced"
         End
         Begin VB.Menu Expert_menue 
            Caption         =   "Expert"
         End
      End
      Begin VB.Menu exit_menue 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu hlep_menue 
      Caption         =   "Help"
      Begin VB.Menu Help2_menue 
         Caption         =   "Help"
      End
      Begin VB.Menu about_menue 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim z(88) As String


Dim v(44) As String

Dim m(44) As Integer

Dim w As String



Private Sub about_menue_Click()
frmAbout.Show
End Sub

Private Sub advanced_menue_Click()
Picture6.Visible = True
Picture6.Enabled = True

Picture7.Visible = True
Picture7.Enabled = True

Picture8.Visible = True
Picture8.Enabled = True

Picture9.Visible = True
Picture9.Enabled = True

Picture10.Visible = True
Picture10.Enabled = True


Picture11.Visible = True
Picture11.Enabled = True


Picture12.Visible = True
Picture12.Enabled = True


Picture13.Visible = True
Picture13.Enabled = True


Picture14.Visible = True
Picture14.Enabled = True


Picture15.Visible = True
Picture15.Enabled = True


Picture16.Visible = True
Picture16.Enabled = True


Picture17.Visible = True
Picture17.Enabled = True


Picture18.Visible = True
Picture18.Enabled = True


Picture19.Visible = True
Picture19.Enabled = True


Picture20.Visible = True
Picture20.Enabled = True


Picture21.Visible = True
Picture21.Enabled = True


Picture22.Visible = True
Picture22.Enabled = True


Picture23.Visible = True
Picture23.Enabled = True


Picture24.Visible = True
Picture24.Enabled = True


Picture25.Visible = True
Picture25.Enabled = True


Picture26.Visible = True
Picture26.Enabled = True


Picture27.Visible = True
Picture27.Enabled = True


Picture28.Visible = True
Picture28.Enabled = True


Picture29.Visible = False
Picture29.Enabled = False


Picture30.Visible = False
Picture30.Enabled = False


Picture31.Visible = False
Picture31.Enabled = False


Picture32.Visible = False
Picture32.Enabled = False


Picture33.Visible = False
Picture33.Enabled = False


Picture34.Visible = False
Picture34.Enabled = False


Picture35.Visible = False
Picture35.Enabled = False


Picture36.Visible = False
Picture36.Enabled = False


Picture37.Visible = False
Picture37.Enabled = False


Picture38.Visible = False
Picture38.Enabled = False


Picture39.Visible = False
Picture39.Enabled = False


Picture40.Visible = False
Picture40.Enabled = False


Picture41.Visible = False
Picture41.Enabled = False


Picture42.Visible = False
Picture42.Enabled = False


Picture43.Visible = False
Picture43.Enabled = False


Picture44.Visible = False
Picture44.Enabled = False


beginner_menue.Checked = False
novice_menue.Checked = False
intermediate_menue.Checked = False
advanced_menue.Checked = True
Expert_menue.Checked = False

Command1.Enabled = False
Command1.Visible = False

Command4.Visible = False
Command4.Enabled = False

Command5.Visible = False
Command5.Enabled = False

Command6.Visible = False
Command6.Enabled = False

Command7.Visible = True
Command7.Enabled = True

End Sub

Private Sub beginner_menue_Click()

Picture6.Visible = False
Picture6.Enabled = False

Picture7.Visible = False
Picture7.Enabled = False


Picture8.Visible = False
Picture8.Enabled = False


Picture9.Visible = False
Picture9.Enabled = False


Picture10.Visible = False
Picture10.Enabled = False


Picture11.Visible = False
Picture11.Enabled = False


Picture12.Visible = False
Picture12.Enabled = False


Picture13.Visible = False
Picture13.Enabled = False


Picture14.Visible = False
Picture14.Enabled = False


Picture15.Visible = False
Picture15.Enabled = False


Picture16.Visible = False
Picture16.Enabled = False


Picture17.Visible = False
Picture17.Enabled = False


Picture18.Visible = False
Picture18.Enabled = False


Picture19.Visible = False
Picture19.Enabled = False


Picture20.Visible = False
Picture20.Enabled = False


Picture21.Visible = False
Picture21.Enabled = False


Picture22.Visible = False
Picture22.Enabled = False


Picture23.Visible = False
Picture23.Enabled = False


Picture24.Visible = False
Picture24.Enabled = False


Picture25.Visible = False
Picture25.Enabled = False


Picture26.Visible = False
Picture26.Enabled = False


Picture27.Visible = False
Picture27.Enabled = False


Picture28.Visible = False
Picture28.Enabled = False


Picture29.Visible = False
Picture29.Enabled = False


Picture30.Visible = False
Picture30.Enabled = False


Picture31.Visible = False
Picture31.Enabled = False


Picture32.Visible = False
Picture32.Enabled = False


Picture33.Visible = False
Picture33.Enabled = False


Picture34.Visible = False
Picture34.Enabled = False


Picture35.Visible = False
Picture35.Enabled = False


Picture36.Visible = False
Picture36.Enabled = False


Picture37.Visible = False
Picture37.Enabled = False


Picture38.Visible = False
Picture38.Enabled = False


Picture39.Visible = False
Picture39.Enabled = False


Picture40.Visible = False
Picture40.Enabled = False


Picture41.Visible = False
Picture41.Enabled = False


Picture42.Visible = False
Picture42.Enabled = False


Picture43.Visible = False
Picture43.Enabled = False


Picture44.Visible = False
Picture44.Enabled = False


beginner_menue.Checked = True
novice_menue.Checked = False
intermediate_menue.Checked = False
advanced_menue.Checked = False
Expert_menue.Checked = False

Command1.Enabled = False
Command1.Visible = False

Command4.Visible = True
Command4.Enabled = True

Command5.Visible = False
Command5.Enabled = False

Command6.Visible = False
Command6.Enabled = False

Command7.Visible = False
Command7.Enabled = False
0
End Sub


Private Sub Command1_Click()
 ' this part
Dim x As Integer
For x = 1 To 44
    v(x) = 0
    z(x) = ""
Next x


Dim i, j, f As Integer
For j = 1 To 44
1:     m(j) = Int(Rnd * 44) + 1
            For f = 1 To (j - 1)
                If m(f) = m(j) Then
                 GoTo 1
                End If
            Next f
Next j

 z(1) = App.Path & "\" & m(1) & ".jpg"
Picture1.Picture = LoadPicture(z(1))

z(2) = App.Path & "\" & m(2) & ".jpg"
Picture2.Picture = LoadPicture(z(2))

z(3) = App.Path & "\" & m(3) & ".jpg"
Picture3.Picture = LoadPicture(z(3))

z(4) = App.Path & "\" & m(4) & ".jpg"
Picture4.Picture = LoadPicture(z(4))

z(5) = App.Path & "\" & m(5) & ".jpg"
Picture5.Picture = LoadPicture(z(5))

z(6) = App.Path & "\" & m(6) & ".jpg"
Picture6.Picture = LoadPicture(z(6))


z(7) = App.Path & "\" & m(7) & ".jpg"
Picture7.Picture = LoadPicture(z(7))


z(8) = App.Path & "\" & m(8) & ".jpg"
Picture8.Picture = LoadPicture(z(8))


z(9) = App.Path & "\" & m(9) & ".jpg"
Picture9.Picture = LoadPicture(z(9))


z(10) = App.Path & "\" & m(10) & ".jpg"
Picture10.Picture = LoadPicture(z(10))


z(11) = App.Path & "\" & m(11) & ".jpg"
Picture11.Picture = LoadPicture(z(11))


z(12) = App.Path & "\" & m(12) & ".jpg"
Picture12.Picture = LoadPicture(z(12))



z(13) = App.Path & "\" & m(13) & ".jpg"
Picture13.Picture = LoadPicture(z(13))


z(14) = App.Path & "\" & m(14) & ".jpg"
Picture14.Picture = LoadPicture(z(14))


z(15) = App.Path & "\" & m(15) & ".jpg"
Picture15.Picture = LoadPicture(z(15))


z(16) = App.Path & "\" & m(16) & ".jpg"
Picture16.Picture = LoadPicture(z(16))


z(17) = App.Path & "\" & m(17) & ".jpg"
Picture17.Picture = LoadPicture(z(17))


z(18) = App.Path & "\" & m(18) & ".jpg"
Picture18.Picture = LoadPicture(z(18))


z(19) = App.Path & "\" & m(19) & ".jpg"
Picture19.Picture = LoadPicture(z(19))


z(20) = App.Path & "\" & m(20) & ".jpg"
Picture20.Picture = LoadPicture(z(20))


z(21) = App.Path & "\" & m(21) & ".jpg"
Picture21.Picture = LoadPicture(z(21))


z(22) = App.Path & "\" & m(22) & ".jpg"
Picture22.Picture = LoadPicture(z(22))


z(23) = App.Path & "\" & m(23) & ".jpg"
Picture23.Picture = LoadPicture(z(23))


z(24) = App.Path & "\" & m(24) & ".jpg"
Picture24.Picture = LoadPicture(z(24))


z(25) = App.Path & "\" & m(25) & ".jpg"
Picture25.Picture = LoadPicture(z(25))


z(26) = App.Path & "\" & m(26) & ".jpg"
Picture26.Picture = LoadPicture(z(26))

z(27) = App.Path & "\" & m(27) & ".jpg"
Picture27.Picture = LoadPicture(z(27))

z(28) = App.Path & "\" & m(28) & ".jpg"
Picture28.Picture = LoadPicture(z(28))

z(29) = App.Path & "\" & m(29) & ".jpg"
Picture29.Picture = LoadPicture(z(29))

z(30) = App.Path & "\" & m(30) & ".jpg"
Picture30.Picture = LoadPicture(z(30))


z(31) = App.Path & "\" & m(31) & ".jpg"
Picture31.Picture = LoadPicture(z(31))


z(32) = App.Path & "\" & m(32) & ".jpg"
Picture32.Picture = LoadPicture(z(32))

z(33) = App.Path & "\" & m(33) & ".jpg"
Picture33.Picture = LoadPicture(z(33))


z(34) = App.Path & "\" & m(34) & ".jpg"
Picture34.Picture = LoadPicture(z(34))


z(35) = App.Path & "\" & m(35) & ".jpg"
Picture35.Picture = LoadPicture(z(35))


z(36) = App.Path & "\" & m(36) & ".jpg"
Picture36.Picture = LoadPicture(z(36))



z(37) = App.Path & "\" & m(37) & ".jpg"
Picture37.Picture = LoadPicture(z(37))


z(38) = App.Path & "\" & m(38) & ".jpg"
Picture38.Picture = LoadPicture(z(38))


z(39) = App.Path & "\" & m(39) & ".jpg"
Picture39.Picture = LoadPicture(z(39))


z(40) = App.Path & "\" & m(40) & ".jpg"
Picture40.Picture = LoadPicture(z(40))



z(41) = App.Path & "\" & m(41) & ".jpg"
Picture41.Picture = LoadPicture(z(41))



z(42) = App.Path & "\" & m(42) & ".jpg"
Picture42.Picture = LoadPicture(z(42))

z(43) = App.Path & "\" & m(43) & ".jpg"
Picture43.Picture = LoadPicture(z(43))

z(44) = App.Path & "\" & m(44) & ".jpg"
Picture44.Picture = LoadPicture(z(44))
      
End Sub

Private Sub Command2_Click()

Picture89.Cls

w = App.Path & "\clear.jpg"

Picture1.Picture = LoadPicture(w)
Picture2.Picture = LoadPicture(w)
Picture3.Picture = LoadPicture(w)
Picture4.Picture = LoadPicture(w)
Picture5.Picture = LoadPicture(w)
Picture6.Picture = LoadPicture(w)
Picture7.Picture = LoadPicture(w)
Picture8.Picture = LoadPicture(w)
Picture9.Picture = LoadPicture(w)
Picture10.Picture = LoadPicture(w)
Picture11.Picture = LoadPicture(w)
Picture12.Picture = LoadPicture(w)
Picture13.Picture = LoadPicture(w)
Picture14.Picture = LoadPicture(w)
Picture15.Picture = LoadPicture(w)
Picture16.Picture = LoadPicture(w)
Picture17.Picture = LoadPicture(w)
Picture18.Picture = LoadPicture(w)
Picture19.Picture = LoadPicture(w)
Picture20.Picture = LoadPicture(w)
Picture21.Picture = LoadPicture(w)
Picture22.Picture = LoadPicture(w)
Picture23.Picture = LoadPicture(w)
Picture24.Picture = LoadPicture(w)
Picture25.Picture = LoadPicture(w)
Picture26.Picture = LoadPicture(w)
Picture27.Picture = LoadPicture(w)
Picture28.Picture = LoadPicture(w)
Picture29.Picture = LoadPicture(w)
Picture30.Picture = LoadPicture(w)
Picture31.Picture = LoadPicture(w)
Picture32.Picture = LoadPicture(w)
Picture33.Picture = LoadPicture(w)
Picture34.Picture = LoadPicture(w)
Picture35.Picture = LoadPicture(w)
Picture36.Picture = LoadPicture(w)
Picture37.Picture = LoadPicture(w)
Picture38.Picture = LoadPicture(w)
Picture39.Picture = LoadPicture(w)
Picture40.Picture = LoadPicture(w)
Picture41.Picture = LoadPicture(w)
Picture42.Picture = LoadPicture(w)
Picture43.Picture = LoadPicture(w)
Picture44.Picture = LoadPicture(w)

End Sub

Private Sub Command3_Click()
w = App.Path & "\clear.jpg"
If v(1) = 0 Then
    Picture1.Picture = LoadPicture(w)
End If
If v(2) = 0 Then
    Picture2.Picture = LoadPicture(w)
End If
If v(3) = 0 Then
    Picture3.Picture = LoadPicture(w)
End If
If v(4) = 0 Then
    Picture4.Picture = LoadPicture(w)
End If
If v(5) = 0 Then
    Picture5.Picture = LoadPicture(w)
End If
If v(6) = 0 Then
    Picture6.Picture = LoadPicture(w)
End If
If v(7) = 0 Then
     Picture7.Picture = LoadPicture(w)
End If
If v(8) = 0 Then
     Picture8.Picture = LoadPicture(w)
End If
If v(9) = 0 Then
     Picture9.Picture = LoadPicture(w)
End If
If v(10) = 0 Then
     Picture10.Picture = LoadPicture(w)
End If
If v(11) = 0 Then
     Picture11.Picture = LoadPicture(w)
    End If
If v(12) = 0 Then
     Picture12.Picture = LoadPicture(w)
    End If
If v(13) = 0 Then
     Picture13.Picture = LoadPicture(w)
    End If
If v(14) = 0 Then
     Picture14.Picture = LoadPicture(w)
    End If
If v(15) = 0 Then
     Picture15.Picture = LoadPicture(w)
    End If
If v(16) = 0 Then
     Picture16.Picture = LoadPicture(w)
     End If
If v(17) = 0 Then
     Picture17.Picture = LoadPicture(w)
    End If
If v(18) = 0 Then
     Picture18.Picture = LoadPicture(w)
    End If
If v(19) = 0 Then
     Picture19.Picture = LoadPicture(w)
     End If
If v(20) = 0 Then
     Picture20.Picture = LoadPicture(w)
     End If
If v(21) = 0 Then
     Picture21.Picture = LoadPicture(w)
     End If
If v(22) = 0 Then
     Picture22.Picture = LoadPicture(w)
    End If
If v(23) = 0 Then
     Picture23.Picture = LoadPicture(w)
    End If
If v(24) = 0 Then
     Picture24.Picture = LoadPicture(w)
     End If
If v(25) = 0 Then
     Picture25.Picture = LoadPicture(w)
     End If
If v(26) = 0 Then
     Picture26.Picture = LoadPicture(w)
     End If
If v(27) = 0 Then
     Picture27.Picture = LoadPicture(w)
     End If
If v(28) = 0 Then
     Picture28.Picture = LoadPicture(w)
     End If
If v(29) = 0 Then
     Picture29.Picture = LoadPicture(w)
     End If
If v(30) = 0 Then
     Picture30.Picture = LoadPicture(w)
     End If
If v(31) = 0 Then
     Picture31.Picture = LoadPicture(w)
     End If
If v(32) = 0 Then
     Picture32.Picture = LoadPicture(w)
     End If
If v(33) = 0 Then
     Picture33.Picture = LoadPicture(w)
     End If
If v(34) = 0 Then
     Picture34.Picture = LoadPicture(w)
     End If
If v(35) = 0 Then
     Picture35.Picture = LoadPicture(w)
     End If
If v(36) = 0 Then
     Picture36.Picture = LoadPicture(w)
     End If
If v(37) = 0 Then
     Picture37.Picture = LoadPicture(w)
     End If
If v(38) = 0 Then
     Picture38.Picture = LoadPicture(w)
     End If
If v(39) = 0 Then
     Picture39.Picture = LoadPicture(w)
     End If
If v(40) = 0 Then
     Picture40.Picture = LoadPicture(w)
     End If
If v(41) = 0 Then
     Picture41.Picture = LoadPicture(w)
     End If
If v(42) = 0 Then
     Picture42.Picture = LoadPicture(w)
     End If
If v(43) = 0 Then
     Picture43.Picture = LoadPicture(w)
     End If
If v(44) = 0 Then
     Picture44.Picture = LoadPicture(w)
     End If


Dim total As Integer
Dim contor As Integer
For contor = 1 To 44
    total = total + v(contor)
Next contor

Picture89.Cls
Picture89.Print "  "; total; "RIGHT ANSWERS"

End Sub

Private Sub Command4_Click()
Dim x As Integer
For x = 1 To 44
    v(x) = 0
    z(x) = ""
Next x

Dim i, j, f As Integer
For j = 1 To 5
1:     m(j) = Int(Rnd * 44) + 1
            For f = 1 To (j - 1)
                If m(f) = m(j) Then
                 GoTo 1
                End If
            Next f
Next j

 z(1) = App.Path & "\" & m(1) & ".jpg"
Picture1.Picture = LoadPicture(z(1))

z(2) = App.Path & "\" & m(2) & ".jpg"
Picture2.Picture = LoadPicture(z(2))

z(3) = App.Path & "\" & m(3) & ".jpg"
Picture3.Picture = LoadPicture(z(3))

z(4) = App.Path & "\" & m(4) & ".jpg"
Picture4.Picture = LoadPicture(z(4))

z(5) = App.Path & "\" & m(5) & ".jpg"
Picture5.Picture = LoadPicture(z(5))

End Sub

Private Sub Command5_Click()
Dim x As Integer
For x = 1 To 44
    v(x) = 0
    z(x) = ""
Next x

Dim i, j, f As Integer
For j = 1 To 10
1:     m(j) = Int(Rnd * 44) + 1
            For f = 1 To (j - 1)
                If m(f) = m(j) Then
                 GoTo 1
                End If
            Next f
Next j
  

 z(1) = App.Path & "\" & m(1) & ".jpg"
Picture1.Picture = LoadPicture(z(1))

z(2) = App.Path & "\" & m(2) & ".jpg"
Picture2.Picture = LoadPicture(z(2))

z(3) = App.Path & "\" & m(3) & ".jpg"
Picture3.Picture = LoadPicture(z(3))

z(4) = App.Path & "\" & m(4) & ".jpg"
Picture4.Picture = LoadPicture(z(4))

z(5) = App.Path & "\" & m(5) & ".jpg"
Picture5.Picture = LoadPicture(z(5))

z(6) = App.Path & "\" & m(6) & ".jpg"
Picture6.Picture = LoadPicture(z(6))


z(7) = App.Path & "\" & m(7) & ".jpg"
Picture7.Picture = LoadPicture(z(7))


z(8) = App.Path & "\" & m(8) & ".jpg"
Picture8.Picture = LoadPicture(z(8))


z(9) = App.Path & "\" & m(9) & ".jpg"
Picture9.Picture = LoadPicture(z(9))


z(10) = App.Path & "\" & m(10) & ".jpg"
Picture10.Picture = LoadPicture(z(10))


End Sub


Private Sub Command6_Click()
Dim x As Integer
For x = 1 To 44
    v(x) = 0
    z(x) = ""
Next x

Dim i, j, f As Integer
For j = 1 To 20
1:     m(j) = Int(Rnd * 44) + 1
            For f = 1 To (j - 1)
                If m(f) = m(j) Then
                 GoTo 1
                End If
            Next f
Next j
    
 z(1) = App.Path & "\" & m(1) & ".jpg"
Picture1.Picture = LoadPicture(z(1))

z(2) = App.Path & "\" & m(2) & ".jpg"
Picture2.Picture = LoadPicture(z(2))

z(3) = App.Path & "\" & m(3) & ".jpg"
Picture3.Picture = LoadPicture(z(3))

z(4) = App.Path & "\" & m(4) & ".jpg"
Picture4.Picture = LoadPicture(z(4))

z(5) = App.Path & "\" & m(5) & ".jpg"
Picture5.Picture = LoadPicture(z(5))

z(6) = App.Path & "\" & m(6) & ".jpg"
Picture6.Picture = LoadPicture(z(6))


z(7) = App.Path & "\" & m(7) & ".jpg"
Picture7.Picture = LoadPicture(z(7))


z(8) = App.Path & "\" & m(8) & ".jpg"
Picture8.Picture = LoadPicture(z(8))


z(9) = App.Path & "\" & m(9) & ".jpg"
Picture9.Picture = LoadPicture(z(9))


z(10) = App.Path & "\" & m(10) & ".jpg"
Picture10.Picture = LoadPicture(z(10))


z(11) = App.Path & "\" & m(11) & ".jpg"
Picture11.Picture = LoadPicture(z(11))


z(12) = App.Path & "\" & m(12) & ".jpg"
Picture12.Picture = LoadPicture(z(12))



z(13) = App.Path & "\" & m(13) & ".jpg"
Picture13.Picture = LoadPicture(z(13))


z(14) = App.Path & "\" & m(14) & ".jpg"
Picture14.Picture = LoadPicture(z(14))


z(15) = App.Path & "\" & m(15) & ".jpg"
Picture15.Picture = LoadPicture(z(15))


z(16) = App.Path & "\" & m(16) & ".jpg"
Picture16.Picture = LoadPicture(z(16))


z(17) = App.Path & "\" & m(17) & ".jpg"
Picture17.Picture = LoadPicture(z(17))


z(18) = App.Path & "\" & m(18) & ".jpg"
Picture18.Picture = LoadPicture(z(18))


z(19) = App.Path & "\" & m(19) & ".jpg"
Picture19.Picture = LoadPicture(z(19))


z(20) = App.Path & "\" & m(20) & ".jpg"
Picture20.Picture = LoadPicture(z(20))

End Sub

Private Sub Command7_Click()
Dim x As Integer
For x = 1 To 44
    v(x) = 0
    z(x) = ""
Next x


Dim i, j, f As Integer
For j = 1 To 28
1:     m(j) = Int(Rnd * 44) + 1
            For f = 1 To (j - 1)
                If m(f) = m(j) Then
                 GoTo 1
                End If
            Next f
Next j

 Picture1.Picture = LoadPicture(App.Path & "\" & m(1) & ".jpg")

 z(1) = App.Path & "\" & m(1) & ".jpg"
Picture1.Picture = LoadPicture(z(1))

z(2) = App.Path & "\" & m(2) & ".jpg"
Picture2.Picture = LoadPicture(z(2))

z(3) = App.Path & "\" & m(3) & ".jpg"
Picture3.Picture = LoadPicture(z(3))

z(4) = App.Path & "\" & m(4) & ".jpg"
Picture4.Picture = LoadPicture(z(4))

z(5) = App.Path & "\" & m(5) & ".jpg"
Picture5.Picture = LoadPicture(z(5))

z(6) = App.Path & "\" & m(6) & ".jpg"
Picture6.Picture = LoadPicture(z(6))


z(7) = App.Path & "\" & m(7) & ".jpg"
Picture7.Picture = LoadPicture(z(7))


z(8) = App.Path & "\" & m(8) & ".jpg"
Picture8.Picture = LoadPicture(z(8))


z(9) = App.Path & "\" & m(9) & ".jpg"
Picture9.Picture = LoadPicture(z(9))


z(10) = App.Path & "\" & m(10) & ".jpg"
Picture10.Picture = LoadPicture(z(10))


z(11) = App.Path & "\" & m(11) & ".jpg"
Picture11.Picture = LoadPicture(z(11))


z(12) = App.Path & "\" & m(12) & ".jpg"
Picture12.Picture = LoadPicture(z(12))



z(13) = App.Path & "\" & m(13) & ".jpg"
Picture13.Picture = LoadPicture(z(13))


z(14) = App.Path & "\" & m(14) & ".jpg"
Picture14.Picture = LoadPicture(z(14))


z(15) = App.Path & "\" & m(15) & ".jpg"
Picture15.Picture = LoadPicture(z(15))


z(16) = App.Path & "\" & m(16) & ".jpg"
Picture16.Picture = LoadPicture(z(16))


z(17) = App.Path & "\" & m(17) & ".jpg"
Picture17.Picture = LoadPicture(z(17))


z(18) = App.Path & "\" & m(18) & ".jpg"
Picture18.Picture = LoadPicture(z(18))


z(19) = App.Path & "\" & m(19) & ".jpg"
Picture19.Picture = LoadPicture(z(19))


z(20) = App.Path & "\" & m(20) & ".jpg"
Picture20.Picture = LoadPicture(z(20))


z(21) = App.Path & "\" & m(21) & ".jpg"
Picture21.Picture = LoadPicture(z(21))


z(22) = App.Path & "\" & m(22) & ".jpg"
Picture22.Picture = LoadPicture(z(22))


z(23) = App.Path & "\" & m(23) & ".jpg"
Picture23.Picture = LoadPicture(z(23))


z(24) = App.Path & "\" & m(24) & ".jpg"
Picture24.Picture = LoadPicture(z(24))


z(25) = App.Path & "\" & m(25) & ".jpg"
Picture25.Picture = LoadPicture(z(25))


z(26) = App.Path & "\" & m(26) & ".jpg"
Picture26.Picture = LoadPicture(z(26))

z(27) = App.Path & "\" & m(27) & ".jpg"
Picture27.Picture = LoadPicture(z(27))

z(28) = App.Path & "\" & m(28) & ".jpg"
Picture28.Picture = LoadPicture(z(28))
    
End Sub

Private Sub exit_menue_Click()
Unload Me
End Sub

Private Sub Expert_menue_Click()
Picture6.Visible = True
Picture6.Enabled = True

Picture7.Visible = True
Picture7.Enabled = True

Picture8.Visible = True
Picture8.Enabled = True

Picture9.Visible = True
Picture9.Enabled = True

Picture10.Visible = True
Picture10.Enabled = True


Picture11.Visible = True
Picture11.Enabled = True


Picture12.Visible = True
Picture12.Enabled = True


Picture13.Visible = True
Picture13.Enabled = True


Picture14.Visible = True
Picture14.Enabled = True


Picture15.Visible = True
Picture15.Enabled = True


Picture16.Visible = True
Picture16.Enabled = True


Picture17.Visible = True
Picture17.Enabled = True


Picture18.Visible = True
Picture18.Enabled = True


Picture19.Visible = True
Picture19.Enabled = True


Picture20.Visible = True
Picture20.Enabled = True


Picture21.Visible = True
Picture21.Enabled = True


Picture22.Visible = True
Picture22.Enabled = True


Picture23.Visible = True
Picture23.Enabled = True


Picture24.Visible = True
Picture24.Enabled = True


Picture25.Visible = True
Picture25.Enabled = True


Picture26.Visible = True
Picture26.Enabled = True


Picture27.Visible = True
Picture27.Enabled = True


Picture28.Visible = True
Picture28.Enabled = True


Picture29.Visible = True
Picture29.Enabled = True


Picture30.Visible = True
Picture30.Enabled = True


Picture31.Visible = True
Picture31.Enabled = True


Picture32.Visible = True
Picture32.Enabled = True


Picture33.Visible = True
Picture33.Enabled = True


Picture34.Visible = True
Picture34.Enabled = True


Picture35.Visible = True
Picture35.Enabled = True


Picture36.Visible = True
Picture36.Enabled = True


Picture37.Visible = True
Picture37.Enabled = True


Picture38.Visible = True
Picture38.Enabled = True


Picture39.Visible = True
Picture39.Enabled = True


Picture40.Visible = True
Picture40.Enabled = True


Picture41.Visible = True
Picture41.Enabled = True


Picture42.Visible = True
Picture42.Enabled = True


Picture43.Visible = True
Picture43.Enabled = True


Picture44.Visible = True
Picture44.Enabled = True


beginner_menue.Checked = False
novice_menue.Checked = False
intermediate_menue.Checked = False
advanced_menue.Checked = False
Expert_menue.Checked = True

Command1.Enabled = True
Command1.Visible = True

Command4.Visible = False
Command4.Enabled = False

Command5.Visible = False
Command5.Enabled = False

Command6.Visible = False
Command6.Enabled = False

Command7.Visible = False
Command7.Enabled = False
End Sub

Private Sub Form_Load()
Randomize Timer

Dim d As Integer
For d = 1 To 44
    v(d) = 0
Next d


z(45) = App.Path & "\" & 1 & ".jpg"
Picture45.Picture = LoadPicture(z(45))

z(46) = App.Path & "\" & 2 & ".jpg"
Picture46.Picture = LoadPicture(z(46))


z(47) = App.Path & "\" & 3 & ".jpg"
Picture47.Picture = LoadPicture(z(47))



z(48) = App.Path & "\" & 4 & ".jpg"
Picture48.Picture = LoadPicture(z(48))


z(49) = App.Path & "\" & 5 & ".jpg"
Picture49.Picture = LoadPicture(z(49))


z(50) = App.Path & "\" & 6 & ".jpg"
Picture50.Picture = LoadPicture(z(50))


z(51) = App.Path & "\" & 7 & ".jpg"
Picture51.Picture = LoadPicture(z(51))



z(52) = App.Path & "\" & 8 & ".jpg"
Picture52.Picture = LoadPicture(z(52))



z(53) = App.Path & "\" & 9 & ".jpg"
Picture53.Picture = LoadPicture(z(53))

z(54) = App.Path & "\" & 10 & ".jpg"
Picture54.Picture = LoadPicture(z(54))

z(55) = App.Path & "\" & 11 & ".jpg"
Picture55.Picture = LoadPicture(z(55))


z(56) = App.Path & "\" & 12 & ".jpg"
Picture56.Picture = LoadPicture(z(56))


z(57) = App.Path & "\" & 13 & ".jpg"
Picture57.Picture = LoadPicture(z(57))


z(58) = App.Path & "\" & 14 & ".jpg"
Picture58.Picture = LoadPicture(z(58))


z(59) = App.Path & "\" & 15 & ".jpg"
Picture59.Picture = LoadPicture(z(59))


z(60) = App.Path & "\" & 16 & ".jpg"
Picture60.Picture = LoadPicture(z(60))



z(61) = App.Path & "\" & 17 & ".jpg"
Picture61.Picture = LoadPicture(z(61))


z(62) = App.Path & "\" & 18 & ".jpg"
Picture62.Picture = LoadPicture(z(62))


z(63) = App.Path & "\" & 19 & ".jpg"
Picture63.Picture = LoadPicture(z(63))


z(64) = App.Path & "\" & 20 & ".jpg"
Picture64.Picture = LoadPicture(z(64))


z(65) = App.Path & "\" & 21 & ".jpg"
Picture65.Picture = LoadPicture(z(65))



z(66) = App.Path & "\" & 22 & ".jpg"
Picture66.Picture = LoadPicture(z(66))




z(67) = App.Path & "\" & 23 & ".jpg"
Picture67.Picture = LoadPicture(z(67))



z(68) = App.Path & "\" & 24 & ".jpg"
Picture68.Picture = LoadPicture(z(68))

z(69) = App.Path & "\" & 25 & ".jpg"
Picture69.Picture = LoadPicture(z(69))

z(70) = App.Path & "\" & 26 & ".jpg"
Picture70.Picture = LoadPicture(z(70))

z(71) = App.Path & "\" & 27 & ".jpg"
Picture71.Picture = LoadPicture(z(71))

z(72) = App.Path & "\" & 28 & ".jpg"
Picture72.Picture = LoadPicture(z(72))


z(73) = App.Path & "\" & 29 & ".jpg"
Picture73.Picture = LoadPicture(z(73))

z(74) = App.Path & "\" & 30 & ".jpg"
Picture74.Picture = LoadPicture(z(74))

z(75) = App.Path & "\" & 31 & ".jpg"
Picture75.Picture = LoadPicture(z(75))

z(76) = App.Path & "\" & 32 & ".jpg"
Picture76.Picture = LoadPicture(z(76))

z(77) = App.Path & "\" & 33 & ".jpg"
Picture77.Picture = LoadPicture(z(77))


z(78) = App.Path & "\" & 34 & ".jpg"
Picture78.Picture = LoadPicture(z(78))


z(79) = App.Path & "\" & 35 & ".jpg"
Picture79.Picture = LoadPicture(z(79))


z(80) = App.Path & "\" & 36 & ".jpg"
Picture80.Picture = LoadPicture(z(80))



z(81) = App.Path & "\" & 37 & ".jpg"
Picture81.Picture = LoadPicture(z(81))



z(82) = App.Path & "\" & 38 & ".jpg"
Picture82.Picture = LoadPicture(z(82))



z(83) = App.Path & "\" & 39 & ".jpg"
Picture83.Picture = LoadPicture(z(83))

z(84) = App.Path & "\" & 40 & ".jpg"
Picture84.Picture = LoadPicture(z(84))

z(85) = App.Path & "\" & 41 & ".jpg"
Picture85.Picture = LoadPicture(z(85))


z(86) = App.Path & "\" & 42 & ".jpg"
Picture86.Picture = LoadPicture(z(86))


z(87) = App.Path & "\" & 43 & ".jpg"
Picture87.Picture = LoadPicture(z(87))


z(88) = App.Path & "\" & 44 & ".jpg"
Picture88.Picture = LoadPicture(z(88))


Call Expert_menue_Click

End Sub



Private Sub Help2_menue_Click()
Dialog.Show
End Sub

Private Sub intermediate_menue_Click()

Picture6.Visible = True
Picture6.Enabled = True

Picture7.Visible = True
Picture7.Enabled = True

Picture8.Visible = True
Picture8.Enabled = True

Picture9.Visible = True
Picture9.Enabled = True

Picture10.Visible = True
Picture10.Enabled = True


Picture11.Visible = True
Picture11.Enabled = True


Picture12.Visible = True
Picture12.Enabled = True


Picture13.Visible = True
Picture13.Enabled = True


Picture14.Visible = True
Picture14.Enabled = True


Picture15.Visible = True
Picture15.Enabled = True


Picture16.Visible = True
Picture16.Enabled = True


Picture17.Visible = True
Picture17.Enabled = True


Picture18.Visible = True
Picture18.Enabled = True


Picture19.Visible = True
Picture19.Enabled = True


Picture20.Visible = True
Picture20.Enabled = True


Picture21.Visible = False
Picture21.Enabled = False


Picture22.Visible = False
Picture22.Enabled = False


Picture23.Visible = False
Picture23.Enabled = False


Picture24.Visible = False
Picture24.Enabled = False


Picture25.Visible = False
Picture25.Enabled = False


Picture26.Visible = False
Picture26.Enabled = False


Picture27.Visible = False
Picture27.Enabled = False


Picture28.Visible = False
Picture28.Enabled = False


Picture29.Visible = False
Picture29.Enabled = False


Picture30.Visible = False
Picture30.Enabled = False


Picture31.Visible = False
Picture31.Enabled = False


Picture32.Visible = False
Picture32.Enabled = False


Picture33.Visible = False
Picture33.Enabled = False


Picture34.Visible = False
Picture34.Enabled = False


Picture35.Visible = False
Picture35.Enabled = False


Picture36.Visible = False
Picture36.Enabled = False


Picture37.Visible = False
Picture37.Enabled = False


Picture38.Visible = False
Picture38.Enabled = False


Picture39.Visible = False
Picture39.Enabled = False


Picture40.Visible = False
Picture40.Enabled = False


Picture41.Visible = False
Picture41.Enabled = False


Picture42.Visible = False
Picture42.Enabled = False


Picture43.Visible = False
Picture43.Enabled = False


Picture44.Visible = False
Picture44.Enabled = False


beginner_menue.Checked = False
novice_menue.Checked = False
intermediate_menue.Checked = True
advanced_menue.Checked = False
Expert_menue.Checked = False

Command1.Enabled = False
Command1.Visible = False

Command4.Visible = False
Command4.Enabled = False

Command5.Visible = False
Command5.Enabled = False

Command6.Visible = True
Command6.Enabled = True

Command7.Visible = False
Command7.Enabled = False

End Sub

Private Sub novice_menue_Click()
Picture6.Visible = True
Picture6.Enabled = True

Picture7.Visible = True
Picture7.Enabled = True

Picture8.Visible = True
Picture8.Enabled = True

Picture9.Visible = True
Picture9.Enabled = True

Picture10.Visible = True
Picture10.Enabled = True




Picture11.Visible = False
Picture11.Enabled = False


Picture12.Visible = False
Picture12.Enabled = False


Picture13.Visible = False
Picture13.Enabled = False


Picture14.Visible = False
Picture14.Enabled = False


Picture15.Visible = False
Picture15.Enabled = False


Picture16.Visible = False
Picture16.Enabled = False


Picture17.Visible = False
Picture17.Enabled = False


Picture18.Visible = False
Picture18.Enabled = False


Picture19.Visible = False
Picture19.Enabled = False


Picture20.Visible = False
Picture20.Enabled = False


Picture21.Visible = False
Picture21.Enabled = False


Picture22.Visible = False
Picture22.Enabled = False


Picture23.Visible = False
Picture23.Enabled = False


Picture24.Visible = False
Picture24.Enabled = False


Picture25.Visible = False
Picture25.Enabled = False


Picture26.Visible = False
Picture26.Enabled = False


Picture27.Visible = False
Picture27.Enabled = False


Picture28.Visible = False
Picture28.Enabled = False


Picture29.Visible = False
Picture29.Enabled = False


Picture30.Visible = False
Picture30.Enabled = False


Picture31.Visible = False
Picture31.Enabled = False


Picture32.Visible = False
Picture32.Enabled = False


Picture33.Visible = False
Picture33.Enabled = False


Picture34.Visible = False
Picture34.Enabled = False


Picture35.Visible = False
Picture35.Enabled = False


Picture36.Visible = False
Picture36.Enabled = False


Picture37.Visible = False
Picture37.Enabled = False


Picture38.Visible = False
Picture38.Enabled = False


Picture39.Visible = False
Picture39.Enabled = False


Picture40.Visible = False
Picture40.Enabled = False


Picture41.Visible = False
Picture41.Enabled = False


Picture42.Visible = False
Picture42.Enabled = False


Picture43.Visible = False
Picture43.Enabled = False


Picture44.Visible = False
Picture44.Enabled = False


beginner_menue.Checked = False
novice_menue.Checked = True
intermediate_menue.Checked = False
advanced_menue.Checked = False
Expert_menue.Checked = False

Command1.Enabled = False
Command1.Visible = False

Command4.Visible = False
Command4.Enabled = False

Command5.Visible = True
Command5.Enabled = True

Command6.Visible = False
Command6.Enabled = False

Command7.Visible = False
Command7.Enabled = False
End Sub

Private Sub Picture1_DblClick()
w = App.Path & "\clear.jpg"
Picture1.Picture = LoadPicture(w)
End Sub

Private Sub Picture1_DragDrop(Source As Control, x As Single, y As Single)
v(1) = 0
If Source = Picture45 Then
    Picture1.Picture = LoadPicture(z(45))
    If z(1) = z(45) Then
        v(1) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture1.Picture = LoadPicture(z(46))
    If z(1) = z(46) Then
        v(1) = v(1) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture1.Picture = LoadPicture(z(47))
    If z(1) = z(47) Then
       v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture1.Picture = LoadPicture(z(48))
    If z(1) = z(48) Then
        v(1) = v(1) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture1.Picture = LoadPicture(z(49))
    If z(1) = z(49) Then
      v(1) = v(1) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture1.Picture = LoadPicture(z(50))
    If z(1) = z(50) Then
      v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture1.Picture = LoadPicture(z(51))
    If z(1) = z(51) Then
        v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture1.Picture = LoadPicture(z(52))
    If z(1) = z(52) Then
        v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture1.Picture = LoadPicture(z(53))
    If z(1) = z(53) Then
      v(1) = v(1) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture1.Picture = LoadPicture(z(54))
    If z(1) = z(54) Then
     v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture1.Picture = LoadPicture(z(55))
    If z(1) = z(55) Then
        v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture1.Picture = LoadPicture(z(56))
    If z(1) = z(56) Then
        v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture1.Picture = LoadPicture(z(57))
    If z(1) = z(57) Then
       v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture1.Picture = LoadPicture(z(58))
    If z(1) = z(58) Then
       v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture1.Picture = LoadPicture(z(59))
    If z(1) = z(59) Then
       v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture1.Picture = LoadPicture(z(60))
    If z(1) = z(60) Then
        v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture1.Picture = LoadPicture(z(61))
    If z(1) = z(61) Then
       v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture1.Picture = LoadPicture(z(62))
    If z(1) = z(62) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture1.Picture = LoadPicture(z(63))
    If z(1) = z(63) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture1.Picture = LoadPicture(z(64))
    If z(1) = z(64) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture1.Picture = LoadPicture(z(65))
    If z(1) = z(65) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture1.Picture = LoadPicture(z(66))
    If z(1) = z(66) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture1.Picture = LoadPicture(z(67))
    If z(1) = z(67) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture1.Picture = LoadPicture(z(68))
    If z(1) = z(68) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture1.Picture = LoadPicture(z(69))
    If z(1) = z(69) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture1.Picture = LoadPicture(z(70))
    If z(1) = z(70) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture1.Picture = LoadPicture(z(71))
    If z(1) = z(71) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture1.Picture = LoadPicture(z(72))
    If z(1) = z(72) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture1.Picture = LoadPicture(z(73))
    If z(1) = z(73) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture1.Picture = LoadPicture(z(74))
    If z(1) = z(74) Then
        v(1) = v(1) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture1.Picture = LoadPicture(z(75))
    If z(1) = z(75) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture1.Picture = LoadPicture(z(76))
    If z(1) = z(76) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture1.Picture = LoadPicture(z(77))
    If z(1) = z(77) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture1.Picture = LoadPicture(z(78))
    If z(1) = z(78) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture1.Picture = LoadPicture(z(79))
    If z(1) = z(79) Then
        v(1) = v(1) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture1.Picture = LoadPicture(z(80))
    If z(1) = z(80) Then
        v(1) = v(1) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture1.Picture = LoadPicture(z(81))
    If z(1) = z(81) Then
        v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture1.Picture = LoadPicture(z(82))
    If z(1) = z(82) Then
        v(1) = v(1) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture1.Picture = LoadPicture(z(83))
    If z(1) = z(83) Then
        v(1) = v(1) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture1.Picture = LoadPicture(z(84))
    If z(1) = z(84) Then
        v(1) = v(1) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture1.Picture = LoadPicture(z(85))
    If z(1) = z(85) Then
        v(1) = v(1) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture1.Picture = LoadPicture(z(86))
    If z(1) = z(86) Then
        v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture1.Picture = LoadPicture(z(87))
    If z(1) = z(87) Then
        v(1) = v(1) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture1.Picture = LoadPicture(z(88))
    If z(1) = z(88) Then
        v(1) = v(1) + 1
        End If
      
End If
    
End Sub



Private Sub Picture10_DblClick()
w = App.Path & "\clear.jpg"
Picture10.Picture = LoadPicture(w)
End Sub

Private Sub Picture10_DragDrop(Source As Control, x As Single, y As Single)
v(10) = 0
If Source = Picture45 Then
    Picture10.Picture = LoadPicture(z(45))
    If z(10) = z(45) Then
        v(10) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture10.Picture = LoadPicture(z(46))
    If z(10) = z(46) Then
        v(10) = v(10) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture10.Picture = LoadPicture(z(47))
    If z(10) = z(47) Then
       v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture10.Picture = LoadPicture(z(48))
    If z(10) = z(48) Then
        v(10) = v(10) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture10.Picture = LoadPicture(z(49))
    If z(10) = z(49) Then
      v(10) = v(10) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture10.Picture = LoadPicture(z(50))
    If z(10) = z(50) Then
      v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture10.Picture = LoadPicture(z(51))
    If z(10) = z(51) Then
        v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture10.Picture = LoadPicture(z(52))
    If z(10) = z(52) Then
        v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture10.Picture = LoadPicture(z(53))
    If z(10) = z(53) Then
      v(10) = v(10) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture10.Picture = LoadPicture(z(54))
    If z(10) = z(54) Then
     v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture10.Picture = LoadPicture(z(55))
    If z(10) = z(55) Then
        v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture10.Picture = LoadPicture(z(56))
    If z(10) = z(56) Then
        v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture10.Picture = LoadPicture(z(57))
    If z(10) = z(57) Then
       v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture10.Picture = LoadPicture(z(58))
    If z(10) = z(58) Then
       v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture10.Picture = LoadPicture(z(59))
    If z(10) = z(59) Then
       v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture10.Picture = LoadPicture(z(60))
    If z(10) = z(60) Then
        v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture10.Picture = LoadPicture(z(61))
    If z(10) = z(61) Then
       v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture10.Picture = LoadPicture(z(62))
    If z(10) = z(62) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture10.Picture = LoadPicture(z(63))
    If z(10) = z(63) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture10.Picture = LoadPicture(z(64))
    If z(10) = z(64) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture10.Picture = LoadPicture(z(65))
    If z(10) = z(65) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture10.Picture = LoadPicture(z(66))
    If z(10) = z(66) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture10.Picture = LoadPicture(z(67))
    If z(10) = z(67) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture10.Picture = LoadPicture(z(68))
    If z(10) = z(68) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture10.Picture = LoadPicture(z(69))
    If z(10) = z(69) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture10.Picture = LoadPicture(z(70))
    If z(10) = z(70) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture10.Picture = LoadPicture(z(71))
    If z(10) = z(71) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture10.Picture = LoadPicture(z(72))
    If z(10) = z(72) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture10.Picture = LoadPicture(z(73))
    If z(10) = z(73) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture10.Picture = LoadPicture(z(74))
    If z(10) = z(74) Then
        v(10) = v(10) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture10.Picture = LoadPicture(z(75))
    If z(10) = z(75) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture10.Picture = LoadPicture(z(76))
    If z(10) = z(76) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture10.Picture = LoadPicture(z(77))
    If z(10) = z(77) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture10.Picture = LoadPicture(z(78))
    If z(10) = z(78) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture10.Picture = LoadPicture(z(79))
    If z(10) = z(79) Then
        v(10) = v(10) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture10.Picture = LoadPicture(z(80))
    If z(10) = z(80) Then
        v(10) = v(10) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture10.Picture = LoadPicture(z(81))
    If z(10) = z(81) Then
        v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture10.Picture = LoadPicture(z(82))
    If z(10) = z(82) Then
        v(10) = v(10) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture10.Picture = LoadPicture(z(83))
    If z(10) = z(83) Then
        v(10) = v(10) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture10.Picture = LoadPicture(z(84))
    If z(10) = z(84) Then
        v(10) = v(10) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture10.Picture = LoadPicture(z(85))
    If z(10) = z(85) Then
        v(10) = v(10) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture10.Picture = LoadPicture(z(86))
    If z(10) = z(86) Then
        v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture10.Picture = LoadPicture(z(87))
    If z(10) = z(87) Then
        v(10) = v(10) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture10.Picture = LoadPicture(z(88))
    If z(10) = z(88) Then
        v(10) = v(10) + 1
        End If
      
End If
End Sub

Private Sub Picture11_DblClick()
w = App.Path & "\clear.jpg"
Picture11.Picture = LoadPicture(w)
End Sub

Private Sub Picture11_DragDrop(Source As Control, x As Single, y As Single)
v(11) = 0
If Source = Picture45 Then
    Picture11.Picture = LoadPicture(z(45))
    If z(11) = z(45) Then
        v(11) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture11.Picture = LoadPicture(z(46))
    If z(11) = z(46) Then
        v(11) = v(11) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture11.Picture = LoadPicture(z(47))
    If z(11) = z(47) Then
       v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture11.Picture = LoadPicture(z(48))
    If z(11) = z(48) Then
        v(11) = v(11) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture11.Picture = LoadPicture(z(49))
    If z(11) = z(49) Then
      v(11) = v(11) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture11.Picture = LoadPicture(z(50))
    If z(11) = z(50) Then
      v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture11.Picture = LoadPicture(z(51))
    If z(11) = z(51) Then
        v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture11.Picture = LoadPicture(z(52))
    If z(11) = z(52) Then
        v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture11.Picture = LoadPicture(z(53))
    If z(11) = z(53) Then
      v(11) = v(11) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture11.Picture = LoadPicture(z(54))
    If z(11) = z(54) Then
     v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture11.Picture = LoadPicture(z(55))
    If z(11) = z(55) Then
        v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture11.Picture = LoadPicture(z(56))
    If z(11) = z(56) Then
        v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture11.Picture = LoadPicture(z(57))
    If z(11) = z(57) Then
       v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture11.Picture = LoadPicture(z(58))
    If z(11) = z(58) Then
       v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture11.Picture = LoadPicture(z(59))
    If z(11) = z(59) Then
       v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture11.Picture = LoadPicture(z(60))
    If z(11) = z(60) Then
        v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture11.Picture = LoadPicture(z(61))
    If z(11) = z(61) Then
       v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture11.Picture = LoadPicture(z(62))
    If z(11) = z(62) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture11.Picture = LoadPicture(z(63))
    If z(11) = z(63) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture11.Picture = LoadPicture(z(64))
    If z(11) = z(64) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture11.Picture = LoadPicture(z(65))
    If z(11) = z(65) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture11.Picture = LoadPicture(z(66))
    If z(11) = z(66) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture11.Picture = LoadPicture(z(67))
    If z(11) = z(67) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture11.Picture = LoadPicture(z(68))
    If z(11) = z(68) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture11.Picture = LoadPicture(z(69))
    If z(11) = z(69) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture11.Picture = LoadPicture(z(70))
    If z(11) = z(70) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture11.Picture = LoadPicture(z(71))
    If z(11) = z(71) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture11.Picture = LoadPicture(z(72))
    If z(11) = z(72) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture11.Picture = LoadPicture(z(73))
    If z(11) = z(73) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture11.Picture = LoadPicture(z(74))
    If z(11) = z(74) Then
        v(11) = v(11) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture11.Picture = LoadPicture(z(75))
    If z(11) = z(75) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture11.Picture = LoadPicture(z(76))
    If z(11) = z(76) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture11.Picture = LoadPicture(z(77))
    If z(11) = z(77) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture11.Picture = LoadPicture(z(78))
    If z(11) = z(78) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture11.Picture = LoadPicture(z(79))
    If z(11) = z(79) Then
        v(11) = v(11) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture11.Picture = LoadPicture(z(80))
    If z(11) = z(80) Then
        v(11) = v(11) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture11.Picture = LoadPicture(z(81))
    If z(11) = z(81) Then
        v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture11.Picture = LoadPicture(z(82))
    If z(11) = z(82) Then
        v(11) = v(11) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture11.Picture = LoadPicture(z(83))
    If z(11) = z(83) Then
        v(11) = v(11) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture11.Picture = LoadPicture(z(84))
    If z(11) = z(84) Then
        v(11) = v(11) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture11.Picture = LoadPicture(z(85))
    If z(11) = z(85) Then
        v(11) = v(11) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture11.Picture = LoadPicture(z(86))
    If z(11) = z(86) Then
        v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture11.Picture = LoadPicture(z(87))
    If z(11) = z(87) Then
        v(11) = v(11) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture11.Picture = LoadPicture(z(88))
    If z(11) = z(88) Then
        v(11) = v(11) + 1
        End If
      
End If
End Sub

Private Sub Picture12_DblClick()
w = App.Path & "\clear.jpg"
Picture12.Picture = LoadPicture(w)
End Sub

Private Sub Picture12_DragDrop(Source As Control, x As Single, y As Single)
v(12) = 0
If Source = Picture45 Then
    Picture12.Picture = LoadPicture(z(45))
    If z(12) = z(45) Then
        v(12) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture12.Picture = LoadPicture(z(46))
    If z(12) = z(46) Then
        v(12) = v(12) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture12.Picture = LoadPicture(z(47))
    If z(12) = z(47) Then
       v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture12.Picture = LoadPicture(z(48))
    If z(12) = z(48) Then
        v(12) = v(12) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture12.Picture = LoadPicture(z(49))
    If z(12) = z(49) Then
      v(12) = v(12) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture12.Picture = LoadPicture(z(50))
    If z(12) = z(50) Then
      v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture12.Picture = LoadPicture(z(51))
    If z(12) = z(51) Then
        v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture12.Picture = LoadPicture(z(52))
    If z(12) = z(52) Then
        v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture12.Picture = LoadPicture(z(53))
    If z(12) = z(53) Then
      v(12) = v(12) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture12.Picture = LoadPicture(z(54))
    If z(12) = z(54) Then
     v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture12.Picture = LoadPicture(z(55))
    If z(12) = z(55) Then
        v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture12.Picture = LoadPicture(z(56))
    If z(12) = z(56) Then
        v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture12.Picture = LoadPicture(z(57))
    If z(12) = z(57) Then
       v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture12.Picture = LoadPicture(z(58))
    If z(12) = z(58) Then
       v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture12.Picture = LoadPicture(z(59))
    If z(12) = z(59) Then
       v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture12.Picture = LoadPicture(z(60))
    If z(12) = z(60) Then
        v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture12.Picture = LoadPicture(z(61))
    If z(12) = z(61) Then
       v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture12.Picture = LoadPicture(z(62))
    If z(12) = z(62) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture12.Picture = LoadPicture(z(63))
    If z(12) = z(63) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture12.Picture = LoadPicture(z(64))
    If z(12) = z(64) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture12.Picture = LoadPicture(z(65))
    If z(12) = z(65) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture12.Picture = LoadPicture(z(66))
    If z(12) = z(66) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture12.Picture = LoadPicture(z(67))
    If z(12) = z(67) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture12.Picture = LoadPicture(z(68))
    If z(12) = z(68) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture12.Picture = LoadPicture(z(69))
    If z(12) = z(69) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture12.Picture = LoadPicture(z(70))
    If z(12) = z(70) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture12.Picture = LoadPicture(z(71))
    If z(12) = z(71) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture12.Picture = LoadPicture(z(72))
    If z(12) = z(72) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture12.Picture = LoadPicture(z(73))
    If z(12) = z(73) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture12.Picture = LoadPicture(z(74))
    If z(12) = z(74) Then
        v(12) = v(12) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture12.Picture = LoadPicture(z(75))
    If z(12) = z(75) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture12.Picture = LoadPicture(z(76))
    If z(12) = z(76) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture12.Picture = LoadPicture(z(77))
    If z(12) = z(77) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture12.Picture = LoadPicture(z(78))
    If z(12) = z(78) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture12.Picture = LoadPicture(z(79))
    If z(12) = z(79) Then
        v(12) = v(12) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture12.Picture = LoadPicture(z(80))
    If z(12) = z(80) Then
        v(12) = v(12) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture12.Picture = LoadPicture(z(81))
    If z(12) = z(81) Then
        v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture12.Picture = LoadPicture(z(82))
    If z(12) = z(82) Then
        v(12) = v(12) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture12.Picture = LoadPicture(z(83))
    If z(12) = z(83) Then
        v(12) = v(12) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture12.Picture = LoadPicture(z(84))
    If z(12) = z(84) Then
        v(12) = v(12) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture12.Picture = LoadPicture(z(85))
    If z(12) = z(85) Then
        v(12) = v(12) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture12.Picture = LoadPicture(z(86))
    If z(12) = z(86) Then
        v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture12.Picture = LoadPicture(z(87))
    If z(12) = z(87) Then
        v(12) = v(12) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture12.Picture = LoadPicture(z(88))
    If z(12) = z(88) Then
        v(12) = v(12) + 1
        End If
      
End If
        
       
End Sub

Private Sub Picture13_DblClick()
w = App.Path & "\clear.jpg"
Picture13.Picture = LoadPicture(w)
End Sub

Private Sub Picture13_DragDrop(Source As Control, x As Single, y As Single)
v(13) = 0
If Source = Picture45 Then
    Picture13.Picture = LoadPicture(z(45))
    If z(13) = z(45) Then
        v(13) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture13.Picture = LoadPicture(z(46))
    If z(13) = z(46) Then
        v(13) = v(13) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture13.Picture = LoadPicture(z(47))
    If z(13) = z(47) Then
       v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture13.Picture = LoadPicture(z(48))
    If z(13) = z(48) Then
        v(13) = v(13) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture13.Picture = LoadPicture(z(49))
    If z(13) = z(49) Then
      v(13) = v(13) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture13.Picture = LoadPicture(z(50))
    If z(13) = z(50) Then
      v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture13.Picture = LoadPicture(z(51))
    If z(13) = z(51) Then
        v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture13.Picture = LoadPicture(z(52))
    If z(13) = z(52) Then
        v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture13.Picture = LoadPicture(z(53))
    If z(13) = z(53) Then
      v(13) = v(13) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture13.Picture = LoadPicture(z(54))
    If z(13) = z(54) Then
     v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture13.Picture = LoadPicture(z(55))
    If z(13) = z(55) Then
        v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture13.Picture = LoadPicture(z(56))
    If z(13) = z(56) Then
        v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture13.Picture = LoadPicture(z(57))
    If z(13) = z(57) Then
       v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture13.Picture = LoadPicture(z(58))
    If z(13) = z(58) Then
       v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture13.Picture = LoadPicture(z(59))
    If z(13) = z(59) Then
       v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture13.Picture = LoadPicture(z(60))
    If z(13) = z(60) Then
        v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture13.Picture = LoadPicture(z(61))
    If z(13) = z(61) Then
       v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture13.Picture = LoadPicture(z(62))
    If z(13) = z(62) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture13.Picture = LoadPicture(z(63))
    If z(13) = z(63) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture13.Picture = LoadPicture(z(64))
    If z(13) = z(64) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture13.Picture = LoadPicture(z(65))
    If z(13) = z(65) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture13.Picture = LoadPicture(z(66))
    If z(13) = z(66) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture13.Picture = LoadPicture(z(67))
    If z(13) = z(67) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture13.Picture = LoadPicture(z(68))
    If z(13) = z(68) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture13.Picture = LoadPicture(z(69))
    If z(13) = z(69) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture13.Picture = LoadPicture(z(70))
    If z(13) = z(70) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture13.Picture = LoadPicture(z(71))
    If z(13) = z(71) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture13.Picture = LoadPicture(z(72))
    If z(13) = z(72) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture13.Picture = LoadPicture(z(73))
    If z(13) = z(73) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture13.Picture = LoadPicture(z(74))
    If z(13) = z(74) Then
        v(13) = v(13) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture13.Picture = LoadPicture(z(75))
    If z(13) = z(75) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture13.Picture = LoadPicture(z(76))
    If z(13) = z(76) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture13.Picture = LoadPicture(z(77))
    If z(13) = z(77) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture13.Picture = LoadPicture(z(78))
    If z(13) = z(78) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture13.Picture = LoadPicture(z(79))
    If z(13) = z(79) Then
        v(13) = v(13) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture13.Picture = LoadPicture(z(80))
    If z(13) = z(80) Then
        v(13) = v(13) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture13.Picture = LoadPicture(z(81))
    If z(13) = z(81) Then
        v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture13.Picture = LoadPicture(z(82))
    If z(13) = z(82) Then
        v(13) = v(13) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture13.Picture = LoadPicture(z(83))
    If z(13) = z(83) Then
        v(13) = v(13) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture13.Picture = LoadPicture(z(84))
    If z(13) = z(84) Then
        v(13) = v(13) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture13.Picture = LoadPicture(z(85))
    If z(13) = z(85) Then
        v(13) = v(13) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture13.Picture = LoadPicture(z(86))
    If z(13) = z(86) Then
        v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture13.Picture = LoadPicture(z(87))
    If z(13) = z(87) Then
        v(13) = v(13) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture13.Picture = LoadPicture(z(88))
    If z(13) = z(88) Then
        v(13) = v(13) + 1
        End If
      
End If

End Sub

Private Sub Picture14_DblClick()
w = App.Path & "\clear.jpg"
Picture14.Picture = LoadPicture(w)
End Sub

Private Sub Picture14_DragDrop(Source As Control, x As Single, y As Single)
v(14) = 0
If Source = Picture45 Then
    Picture14.Picture = LoadPicture(z(45))
    If z(14) = z(45) Then
        v(14) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture14.Picture = LoadPicture(z(46))
    If z(14) = z(46) Then
        v(14) = v(14) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture14.Picture = LoadPicture(z(47))
    If z(14) = z(47) Then
       v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture14.Picture = LoadPicture(z(48))
    If z(14) = z(48) Then
        v(14) = v(14) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture14.Picture = LoadPicture(z(49))
    If z(14) = z(49) Then
      v(14) = v(14) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture14.Picture = LoadPicture(z(50))
    If z(14) = z(50) Then
      v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture14.Picture = LoadPicture(z(51))
    If z(14) = z(51) Then
        v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture14.Picture = LoadPicture(z(52))
    If z(14) = z(52) Then
        v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture14.Picture = LoadPicture(z(53))
    If z(14) = z(53) Then
      v(14) = v(14) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture14.Picture = LoadPicture(z(54))
    If z(14) = z(54) Then
     v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture14.Picture = LoadPicture(z(55))
    If z(14) = z(55) Then
        v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture14.Picture = LoadPicture(z(56))
    If z(14) = z(56) Then
        v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture14.Picture = LoadPicture(z(57))
    If z(14) = z(57) Then
       v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture14.Picture = LoadPicture(z(58))
    If z(14) = z(58) Then
       v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture14.Picture = LoadPicture(z(59))
    If z(14) = z(59) Then
       v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture14.Picture = LoadPicture(z(60))
    If z(14) = z(60) Then
        v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture14.Picture = LoadPicture(z(61))
    If z(14) = z(61) Then
       v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture14.Picture = LoadPicture(z(62))
    If z(14) = z(62) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture14.Picture = LoadPicture(z(63))
    If z(14) = z(63) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture14.Picture = LoadPicture(z(64))
    If z(14) = z(64) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture14.Picture = LoadPicture(z(65))
    If z(14) = z(65) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture14.Picture = LoadPicture(z(66))
    If z(14) = z(66) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture14.Picture = LoadPicture(z(67))
    If z(14) = z(67) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture14.Picture = LoadPicture(z(68))
    If z(14) = z(68) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture14.Picture = LoadPicture(z(69))
    If z(14) = z(69) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture14.Picture = LoadPicture(z(70))
    If z(14) = z(70) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture14.Picture = LoadPicture(z(71))
    If z(14) = z(71) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture14.Picture = LoadPicture(z(72))
    If z(14) = z(72) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture14.Picture = LoadPicture(z(73))
    If z(14) = z(73) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture14.Picture = LoadPicture(z(74))
    If z(14) = z(74) Then
        v(14) = v(14) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture14.Picture = LoadPicture(z(75))
    If z(14) = z(75) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture14.Picture = LoadPicture(z(76))
    If z(14) = z(76) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture14.Picture = LoadPicture(z(77))
    If z(14) = z(77) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture14.Picture = LoadPicture(z(78))
    If z(14) = z(78) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture14.Picture = LoadPicture(z(79))
    If z(14) = z(79) Then
        v(14) = v(14) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture14.Picture = LoadPicture(z(80))
    If z(14) = z(80) Then
        v(14) = v(14) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture14.Picture = LoadPicture(z(81))
    If z(14) = z(81) Then
        v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture14.Picture = LoadPicture(z(82))
    If z(14) = z(82) Then
        v(14) = v(14) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture14.Picture = LoadPicture(z(83))
    If z(14) = z(83) Then
        v(14) = v(14) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture14.Picture = LoadPicture(z(84))
    If z(14) = z(84) Then
        v(14) = v(14) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture14.Picture = LoadPicture(z(85))
    If z(14) = z(85) Then
        v(14) = v(14) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture14.Picture = LoadPicture(z(86))
    If z(14) = z(86) Then
        v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture14.Picture = LoadPicture(z(87))
    If z(14) = z(87) Then
        v(14) = v(14) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture14.Picture = LoadPicture(z(88))
    If z(14) = z(88) Then
        v(14) = v(14) + 1
        End If
      
End If


End Sub

Private Sub Picture15_DblClick()
w = App.Path & "\clear.jpg"
Picture15.Picture = LoadPicture(w)
End Sub

Private Sub Picture15_DragDrop(Source As Control, x As Single, y As Single)
v(15) = 0
If Source = Picture45 Then
    Picture15.Picture = LoadPicture(z(45))
    If z(15) = z(45) Then
        v(15) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture15.Picture = LoadPicture(z(46))
    If z(15) = z(46) Then
        v(15) = v(15) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture15.Picture = LoadPicture(z(47))
    If z(15) = z(47) Then
       v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture15.Picture = LoadPicture(z(48))
    If z(15) = z(48) Then
        v(15) = v(15) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture15.Picture = LoadPicture(z(49))
    If z(15) = z(49) Then
      v(15) = v(15) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture15.Picture = LoadPicture(z(50))
    If z(15) = z(50) Then
      v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture15.Picture = LoadPicture(z(51))
    If z(15) = z(51) Then
        v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture15.Picture = LoadPicture(z(52))
    If z(15) = z(52) Then
        v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture15.Picture = LoadPicture(z(53))
    If z(15) = z(53) Then
      v(15) = v(15) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture15.Picture = LoadPicture(z(54))
    If z(15) = z(54) Then
     v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture15.Picture = LoadPicture(z(55))
    If z(15) = z(55) Then
        v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture15.Picture = LoadPicture(z(56))
    If z(15) = z(56) Then
        v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture15.Picture = LoadPicture(z(57))
    If z(15) = z(57) Then
       v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture15.Picture = LoadPicture(z(58))
    If z(15) = z(58) Then
       v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture15.Picture = LoadPicture(z(59))
    If z(15) = z(59) Then
       v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture15.Picture = LoadPicture(z(60))
    If z(15) = z(60) Then
        v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture15.Picture = LoadPicture(z(61))
    If z(15) = z(61) Then
       v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture15.Picture = LoadPicture(z(62))
    If z(15) = z(62) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture15.Picture = LoadPicture(z(63))
    If z(15) = z(63) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture15.Picture = LoadPicture(z(64))
    If z(15) = z(64) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture15.Picture = LoadPicture(z(65))
    If z(15) = z(65) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture15.Picture = LoadPicture(z(66))
    If z(15) = z(66) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture15.Picture = LoadPicture(z(67))
    If z(15) = z(67) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture15.Picture = LoadPicture(z(68))
    If z(15) = z(68) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture15.Picture = LoadPicture(z(69))
    If z(15) = z(69) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture15.Picture = LoadPicture(z(70))
    If z(15) = z(70) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture15.Picture = LoadPicture(z(71))
    If z(15) = z(71) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture15.Picture = LoadPicture(z(72))
    If z(15) = z(72) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture15.Picture = LoadPicture(z(73))
    If z(15) = z(73) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture15.Picture = LoadPicture(z(74))
    If z(15) = z(74) Then
        v(15) = v(15) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture15.Picture = LoadPicture(z(75))
    If z(15) = z(75) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture15.Picture = LoadPicture(z(76))
    If z(15) = z(76) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture15.Picture = LoadPicture(z(77))
    If z(15) = z(77) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture15.Picture = LoadPicture(z(78))
    If z(15) = z(78) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture15.Picture = LoadPicture(z(79))
    If z(15) = z(79) Then
        v(15) = v(15) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture15.Picture = LoadPicture(z(80))
    If z(15) = z(80) Then
        v(15) = v(15) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture15.Picture = LoadPicture(z(81))
    If z(15) = z(81) Then
        v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture15.Picture = LoadPicture(z(82))
    If z(15) = z(82) Then
        v(15) = v(15) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture15.Picture = LoadPicture(z(83))
    If z(15) = z(83) Then
        v(15) = v(15) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture15.Picture = LoadPicture(z(84))
    If z(15) = z(84) Then
        v(15) = v(15) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture15.Picture = LoadPicture(z(85))
    If z(15) = z(85) Then
        v(15) = v(15) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture15.Picture = LoadPicture(z(86))
    If z(15) = z(86) Then
        v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture15.Picture = LoadPicture(z(87))
    If z(15) = z(87) Then
        v(15) = v(15) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture15.Picture = LoadPicture(z(88))
    If z(15) = z(88) Then
        v(15) = v(15) + 1
        End If
      
End If
End Sub

Private Sub Picture16_DblClick()
w = App.Path & "\clear.jpg"
Picture16.Picture = LoadPicture(w)
End Sub

Private Sub Picture16_DragDrop(Source As Control, x As Single, y As Single)
v(16) = 0
If Source = Picture45 Then
    Picture16.Picture = LoadPicture(z(45))
    If z(16) = z(45) Then
        v(16) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture16.Picture = LoadPicture(z(46))
    If z(16) = z(46) Then
        v(16) = v(16) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture16.Picture = LoadPicture(z(47))
    If z(16) = z(47) Then
       v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture16.Picture = LoadPicture(z(48))
    If z(16) = z(48) Then
        v(16) = v(16) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture16.Picture = LoadPicture(z(49))
    If z(16) = z(49) Then
      v(16) = v(16) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture16.Picture = LoadPicture(z(50))
    If z(16) = z(50) Then
      v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture16.Picture = LoadPicture(z(51))
    If z(16) = z(51) Then
        v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture16.Picture = LoadPicture(z(52))
    If z(16) = z(52) Then
        v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture16.Picture = LoadPicture(z(53))
    If z(16) = z(53) Then
      v(16) = v(16) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture16.Picture = LoadPicture(z(54))
    If z(16) = z(54) Then
     v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture16.Picture = LoadPicture(z(55))
    If z(16) = z(55) Then
        v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture16.Picture = LoadPicture(z(56))
    If z(16) = z(56) Then
        v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture16.Picture = LoadPicture(z(57))
    If z(16) = z(57) Then
       v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture16.Picture = LoadPicture(z(58))
    If z(16) = z(58) Then
       v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture16.Picture = LoadPicture(z(59))
    If z(16) = z(59) Then
       v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture16.Picture = LoadPicture(z(60))
    If z(16) = z(60) Then
        v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture16.Picture = LoadPicture(z(61))
    If z(16) = z(61) Then
       v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture16.Picture = LoadPicture(z(62))
    If z(16) = z(62) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture16.Picture = LoadPicture(z(63))
    If z(16) = z(63) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture16.Picture = LoadPicture(z(64))
    If z(16) = z(64) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture16.Picture = LoadPicture(z(65))
    If z(16) = z(65) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture16.Picture = LoadPicture(z(66))
    If z(16) = z(66) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture16.Picture = LoadPicture(z(67))
    If z(16) = z(67) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture16.Picture = LoadPicture(z(68))
    If z(16) = z(68) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture16.Picture = LoadPicture(z(69))
    If z(16) = z(69) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture16.Picture = LoadPicture(z(70))
    If z(16) = z(70) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture16.Picture = LoadPicture(z(71))
    If z(16) = z(71) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture16.Picture = LoadPicture(z(72))
    If z(16) = z(72) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture16.Picture = LoadPicture(z(73))
    If z(16) = z(73) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture16.Picture = LoadPicture(z(74))
    If z(16) = z(74) Then
        v(16) = v(16) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture16.Picture = LoadPicture(z(75))
    If z(16) = z(75) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture16.Picture = LoadPicture(z(76))
    If z(16) = z(76) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture16.Picture = LoadPicture(z(77))
    If z(16) = z(77) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture16.Picture = LoadPicture(z(78))
    If z(16) = z(78) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture16.Picture = LoadPicture(z(79))
    If z(16) = z(79) Then
        v(16) = v(16) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture16.Picture = LoadPicture(z(80))
    If z(16) = z(80) Then
        v(16) = v(16) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture16.Picture = LoadPicture(z(81))
    If z(16) = z(81) Then
        v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture16.Picture = LoadPicture(z(82))
    If z(16) = z(82) Then
        v(16) = v(16) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture16.Picture = LoadPicture(z(83))
    If z(16) = z(83) Then
        v(16) = v(16) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture16.Picture = LoadPicture(z(84))
    If z(16) = z(84) Then
        v(16) = v(16) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture16.Picture = LoadPicture(z(85))
    If z(16) = z(85) Then
        v(16) = v(16) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture16.Picture = LoadPicture(z(86))
    If z(16) = z(86) Then
        v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture16.Picture = LoadPicture(z(87))
    If z(16) = z(87) Then
        v(16) = v(16) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture16.Picture = LoadPicture(z(88))
    If z(16) = z(88) Then
        v(16) = v(16) + 1
        End If
      
End If
End Sub

Private Sub Picture17_DblClick()
w = App.Path & "\clear.jpg"
Picture17.Picture = LoadPicture(w)
End Sub

Private Sub Picture17_DragDrop(Source As Control, x As Single, y As Single)
v(17) = 0
If Source = Picture45 Then
    Picture17.Picture = LoadPicture(z(45))
    If z(17) = z(45) Then
        v(17) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture17.Picture = LoadPicture(z(46))
    If z(17) = z(46) Then
        v(17) = v(17) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture17.Picture = LoadPicture(z(47))
    If z(17) = z(47) Then
       v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture17.Picture = LoadPicture(z(48))
    If z(17) = z(48) Then
        v(17) = v(17) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture17.Picture = LoadPicture(z(49))
    If z(17) = z(49) Then
      v(17) = v(17) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture17.Picture = LoadPicture(z(50))
    If z(17) = z(50) Then
      v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture17.Picture = LoadPicture(z(51))
    If z(17) = z(51) Then
        v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture17.Picture = LoadPicture(z(52))
    If z(17) = z(52) Then
        v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture17.Picture = LoadPicture(z(53))
    If z(17) = z(53) Then
      v(17) = v(17) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture17.Picture = LoadPicture(z(54))
    If z(17) = z(54) Then
     v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture17.Picture = LoadPicture(z(55))
    If z(17) = z(55) Then
        v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture17.Picture = LoadPicture(z(56))
    If z(17) = z(56) Then
        v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture17.Picture = LoadPicture(z(57))
    If z(17) = z(57) Then
       v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture17.Picture = LoadPicture(z(58))
    If z(17) = z(58) Then
       v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture17.Picture = LoadPicture(z(59))
    If z(17) = z(59) Then
       v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture17.Picture = LoadPicture(z(60))
    If z(17) = z(60) Then
        v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture17.Picture = LoadPicture(z(61))
    If z(17) = z(61) Then
       v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture17.Picture = LoadPicture(z(62))
    If z(17) = z(62) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture17.Picture = LoadPicture(z(63))
    If z(17) = z(63) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture17.Picture = LoadPicture(z(64))
    If z(17) = z(64) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture17.Picture = LoadPicture(z(65))
    If z(17) = z(65) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture17.Picture = LoadPicture(z(66))
    If z(17) = z(66) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture17.Picture = LoadPicture(z(67))
    If z(17) = z(67) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture17.Picture = LoadPicture(z(68))
    If z(17) = z(68) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture17.Picture = LoadPicture(z(69))
    If z(17) = z(69) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture17.Picture = LoadPicture(z(70))
    If z(17) = z(70) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture17.Picture = LoadPicture(z(71))
    If z(17) = z(71) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture17.Picture = LoadPicture(z(72))
    If z(17) = z(72) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture17.Picture = LoadPicture(z(73))
    If z(17) = z(73) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture17.Picture = LoadPicture(z(74))
    If z(17) = z(74) Then
        v(17) = v(17) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture17.Picture = LoadPicture(z(75))
    If z(17) = z(75) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture17.Picture = LoadPicture(z(76))
    If z(17) = z(76) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture17.Picture = LoadPicture(z(77))
    If z(17) = z(77) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture17.Picture = LoadPicture(z(78))
    If z(17) = z(78) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture17.Picture = LoadPicture(z(79))
    If z(17) = z(79) Then
        v(17) = v(17) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture17.Picture = LoadPicture(z(80))
    If z(17) = z(80) Then
        v(17) = v(17) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture17.Picture = LoadPicture(z(81))
    If z(17) = z(81) Then
        v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture17.Picture = LoadPicture(z(82))
    If z(17) = z(82) Then
        v(17) = v(17) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture17.Picture = LoadPicture(z(83))
    If z(17) = z(83) Then
        v(17) = v(17) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture17.Picture = LoadPicture(z(84))
    If z(17) = z(84) Then
        v(17) = v(17) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture17.Picture = LoadPicture(z(85))
    If z(17) = z(85) Then
        v(17) = v(17) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture17.Picture = LoadPicture(z(86))
    If z(17) = z(86) Then
        v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture17.Picture = LoadPicture(z(87))
    If z(17) = z(87) Then
        v(17) = v(17) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture17.Picture = LoadPicture(z(88))
    If z(17) = z(88) Then
        v(17) = v(17) + 1
        End If
      
End If

End Sub

Private Sub Picture18_DblClick()
w = App.Path & "\clear.jpg"
Picture18.Picture = LoadPicture(w)
End Sub

Private Sub Picture18_DragDrop(Source As Control, x As Single, y As Single)
v(18) = 0
If Source = Picture45 Then
    Picture18.Picture = LoadPicture(z(45))
    If z(18) = z(45) Then
        v(18) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture18.Picture = LoadPicture(z(46))
    If z(18) = z(46) Then
        v(18) = v(18) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture18.Picture = LoadPicture(z(47))
    If z(18) = z(47) Then
       v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture18.Picture = LoadPicture(z(48))
    If z(18) = z(48) Then
        v(18) = v(18) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture18.Picture = LoadPicture(z(49))
    If z(18) = z(49) Then
      v(18) = v(18) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture18.Picture = LoadPicture(z(50))
    If z(18) = z(50) Then
      v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture18.Picture = LoadPicture(z(51))
    If z(18) = z(51) Then
        v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture18.Picture = LoadPicture(z(52))
    If z(18) = z(52) Then
        v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture18.Picture = LoadPicture(z(53))
    If z(18) = z(53) Then
      v(18) = v(18) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture18.Picture = LoadPicture(z(54))
    If z(18) = z(54) Then
     v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture18.Picture = LoadPicture(z(55))
    If z(18) = z(55) Then
        v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture18.Picture = LoadPicture(z(56))
    If z(18) = z(56) Then
        v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture18.Picture = LoadPicture(z(57))
    If z(18) = z(57) Then
       v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture18.Picture = LoadPicture(z(58))
    If z(18) = z(58) Then
       v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture18.Picture = LoadPicture(z(59))
    If z(18) = z(59) Then
       v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture18.Picture = LoadPicture(z(60))
    If z(18) = z(60) Then
        v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture18.Picture = LoadPicture(z(61))
    If z(18) = z(61) Then
       v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture18.Picture = LoadPicture(z(62))
    If z(18) = z(62) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture18.Picture = LoadPicture(z(63))
    If z(18) = z(63) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture18.Picture = LoadPicture(z(64))
    If z(18) = z(64) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture18.Picture = LoadPicture(z(65))
    If z(18) = z(65) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture18.Picture = LoadPicture(z(66))
    If z(18) = z(66) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture18.Picture = LoadPicture(z(67))
    If z(18) = z(67) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture18.Picture = LoadPicture(z(68))
    If z(18) = z(68) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture18.Picture = LoadPicture(z(69))
    If z(18) = z(69) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture18.Picture = LoadPicture(z(70))
    If z(18) = z(70) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture18.Picture = LoadPicture(z(71))
    If z(18) = z(71) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture18.Picture = LoadPicture(z(72))
    If z(18) = z(72) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture18.Picture = LoadPicture(z(73))
    If z(18) = z(73) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture18.Picture = LoadPicture(z(74))
    If z(18) = z(74) Then
        v(18) = v(18) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture18.Picture = LoadPicture(z(75))
    If z(18) = z(75) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture18.Picture = LoadPicture(z(76))
    If z(18) = z(76) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture18.Picture = LoadPicture(z(77))
    If z(18) = z(77) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture18.Picture = LoadPicture(z(78))
    If z(18) = z(78) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture18.Picture = LoadPicture(z(79))
    If z(18) = z(79) Then
        v(18) = v(18) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture18.Picture = LoadPicture(z(80))
    If z(18) = z(80) Then
        v(18) = v(18) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture18.Picture = LoadPicture(z(81))
    If z(18) = z(81) Then
        v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture18.Picture = LoadPicture(z(82))
    If z(18) = z(82) Then
        v(18) = v(18) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture18.Picture = LoadPicture(z(83))
    If z(18) = z(83) Then
        v(18) = v(18) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture18.Picture = LoadPicture(z(84))
    If z(18) = z(84) Then
        v(18) = v(18) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture18.Picture = LoadPicture(z(85))
    If z(18) = z(85) Then
        v(18) = v(18) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture18.Picture = LoadPicture(z(86))
    If z(18) = z(86) Then
        v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture18.Picture = LoadPicture(z(87))
    If z(18) = z(87) Then
        v(18) = v(18) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture18.Picture = LoadPicture(z(88))
    If z(18) = z(88) Then
        v(18) = v(18) + 1
        End If
      
End If

End Sub

Private Sub Picture19_DblClick()
w = App.Path & "\clear.jpg"
Picture19.Picture = LoadPicture(w)
End Sub

Private Sub Picture19_DragDrop(Source As Control, x As Single, y As Single)
v(19) = 0
If Source = Picture45 Then
    Picture19.Picture = LoadPicture(z(45))
    If z(19) = z(45) Then
        v(19) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture19.Picture = LoadPicture(z(46))
    If z(19) = z(46) Then
        v(19) = v(19) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture19.Picture = LoadPicture(z(47))
    If z(19) = z(47) Then
       v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture19.Picture = LoadPicture(z(48))
    If z(19) = z(48) Then
        v(19) = v(19) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture19.Picture = LoadPicture(z(49))
    If z(19) = z(49) Then
      v(19) = v(19) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture19.Picture = LoadPicture(z(50))
    If z(19) = z(50) Then
      v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture19.Picture = LoadPicture(z(51))
    If z(19) = z(51) Then
        v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture19.Picture = LoadPicture(z(52))
    If z(19) = z(52) Then
        v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture19.Picture = LoadPicture(z(53))
    If z(19) = z(53) Then
      v(19) = v(19) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture19.Picture = LoadPicture(z(54))
    If z(19) = z(54) Then
     v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture19.Picture = LoadPicture(z(55))
    If z(19) = z(55) Then
        v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture19.Picture = LoadPicture(z(56))
    If z(19) = z(56) Then
        v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture19.Picture = LoadPicture(z(57))
    If z(19) = z(57) Then
       v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture19.Picture = LoadPicture(z(58))
    If z(19) = z(58) Then
       v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture19.Picture = LoadPicture(z(59))
    If z(19) = z(59) Then
       v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture19.Picture = LoadPicture(z(60))
    If z(19) = z(60) Then
        v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture19.Picture = LoadPicture(z(61))
    If z(19) = z(61) Then
       v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture19.Picture = LoadPicture(z(62))
    If z(19) = z(62) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture19.Picture = LoadPicture(z(63))
    If z(19) = z(63) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture19.Picture = LoadPicture(z(64))
    If z(19) = z(64) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture19.Picture = LoadPicture(z(65))
    If z(19) = z(65) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture19.Picture = LoadPicture(z(66))
    If z(19) = z(66) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture19.Picture = LoadPicture(z(67))
    If z(19) = z(67) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture19.Picture = LoadPicture(z(68))
    If z(19) = z(68) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture19.Picture = LoadPicture(z(69))
    If z(19) = z(69) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture19.Picture = LoadPicture(z(70))
    If z(19) = z(70) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture19.Picture = LoadPicture(z(71))
    If z(19) = z(71) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture19.Picture = LoadPicture(z(72))
    If z(19) = z(72) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture19.Picture = LoadPicture(z(73))
    If z(19) = z(73) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture19.Picture = LoadPicture(z(74))
    If z(19) = z(74) Then
        v(19) = v(19) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture19.Picture = LoadPicture(z(75))
    If z(19) = z(75) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture19.Picture = LoadPicture(z(76))
    If z(19) = z(76) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture19.Picture = LoadPicture(z(77))
    If z(19) = z(77) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture19.Picture = LoadPicture(z(78))
    If z(19) = z(78) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture19.Picture = LoadPicture(z(79))
    If z(19) = z(79) Then
        v(19) = v(19) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture19.Picture = LoadPicture(z(80))
    If z(19) = z(80) Then
        v(19) = v(19) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture19.Picture = LoadPicture(z(81))
    If z(19) = z(81) Then
        v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture19.Picture = LoadPicture(z(82))
    If z(19) = z(82) Then
        v(19) = v(19) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture19.Picture = LoadPicture(z(83))
    If z(19) = z(83) Then
        v(19) = v(19) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture19.Picture = LoadPicture(z(84))
    If z(19) = z(84) Then
        v(19) = v(19) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture19.Picture = LoadPicture(z(85))
    If z(19) = z(85) Then
        v(19) = v(19) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture19.Picture = LoadPicture(z(86))
    If z(19) = z(86) Then
        v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture19.Picture = LoadPicture(z(87))
    If z(19) = z(87) Then
        v(19) = v(19) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture19.Picture = LoadPicture(z(88))
    If z(19) = z(88) Then
        v(19) = v(19) + 1
        End If
      
End If

End Sub

Private Sub Picture2_DblClick()
w = App.Path & "\clear.jpg"
Picture2.Picture = LoadPicture(w)
End Sub

Private Sub Picture2_DragDrop(Source As Control, x As Single, y As Single)
v(2) = 0
If Source = Picture45 Then
    Picture2.Picture = LoadPicture(z(45))
    If z(2) = z(45) Then
        v(2) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture2.Picture = LoadPicture(z(46))
    If z(2) = z(46) Then
        v(2) = v(2) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture2.Picture = LoadPicture(z(47))
    If z(2) = z(47) Then
       v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture2.Picture = LoadPicture(z(48))
    If z(2) = z(48) Then
        v(2) = v(2) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture2.Picture = LoadPicture(z(49))
    If z(2) = z(49) Then
      v(2) = v(2) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture2.Picture = LoadPicture(z(50))
    If z(2) = z(50) Then
      v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture2.Picture = LoadPicture(z(51))
    If z(2) = z(51) Then
        v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture2.Picture = LoadPicture(z(52))
    If z(2) = z(52) Then
        v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture2.Picture = LoadPicture(z(53))
    If z(2) = z(53) Then
      v(2) = v(2) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture2.Picture = LoadPicture(z(54))
    If z(2) = z(54) Then
     v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture2.Picture = LoadPicture(z(55))
    If z(2) = z(55) Then
        v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture2.Picture = LoadPicture(z(56))
    If z(2) = z(56) Then
        v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture2.Picture = LoadPicture(z(57))
    If z(2) = z(57) Then
       v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture2.Picture = LoadPicture(z(58))
    If z(2) = z(58) Then
       v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture2.Picture = LoadPicture(z(59))
    If z(2) = z(59) Then
       v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture2.Picture = LoadPicture(z(60))
    If z(2) = z(60) Then
        v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture2.Picture = LoadPicture(z(61))
    If z(2) = z(61) Then
       v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture2.Picture = LoadPicture(z(62))
    If z(2) = z(62) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture2.Picture = LoadPicture(z(63))
    If z(2) = z(63) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture2.Picture = LoadPicture(z(64))
    If z(2) = z(64) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture2.Picture = LoadPicture(z(65))
    If z(2) = z(65) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture2.Picture = LoadPicture(z(66))
    If z(2) = z(66) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture2.Picture = LoadPicture(z(67))
    If z(2) = z(67) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture2.Picture = LoadPicture(z(68))
    If z(2) = z(68) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture2.Picture = LoadPicture(z(69))
    If z(2) = z(69) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture2.Picture = LoadPicture(z(70))
    If z(2) = z(70) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture2.Picture = LoadPicture(z(71))
    If z(2) = z(71) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture2.Picture = LoadPicture(z(72))
    If z(2) = z(72) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture2.Picture = LoadPicture(z(73))
    If z(2) = z(73) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture2.Picture = LoadPicture(z(74))
    If z(2) = z(74) Then
        v(2) = v(2) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture2.Picture = LoadPicture(z(75))
    If z(2) = z(75) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture2.Picture = LoadPicture(z(76))
    If z(2) = z(76) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture2.Picture = LoadPicture(z(77))
    If z(2) = z(77) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture2.Picture = LoadPicture(z(78))
    If z(2) = z(78) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture2.Picture = LoadPicture(z(79))
    If z(2) = z(79) Then
        v(2) = v(2) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture2.Picture = LoadPicture(z(80))
    If z(2) = z(80) Then
        v(2) = v(2) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture2.Picture = LoadPicture(z(81))
    If z(2) = z(81) Then
        v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture2.Picture = LoadPicture(z(82))
    If z(2) = z(82) Then
        v(2) = v(2) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture2.Picture = LoadPicture(z(83))
    If z(2) = z(83) Then
        v(2) = v(2) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture2.Picture = LoadPicture(z(84))
    If z(2) = z(84) Then
        v(2) = v(2) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture2.Picture = LoadPicture(z(85))
    If z(2) = z(85) Then
        v(2) = v(2) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture2.Picture = LoadPicture(z(86))
    If z(2) = z(86) Then
        v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture2.Picture = LoadPicture(z(87))
    If z(2) = z(87) Then
        v(2) = v(2) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture2.Picture = LoadPicture(z(88))
    If z(2) = z(88) Then
        v(2) = v(2) + 1
        End If
      
End If
End Sub

Private Sub Picture20_DblClick()
w = App.Path & "\clear.jpg"
Picture20.Picture = LoadPicture(w)
End Sub

Private Sub Picture20_DragDrop(Source As Control, x As Single, y As Single)
v(20) = 0
If Source = Picture45 Then
    Picture20.Picture = LoadPicture(z(45))
    If z(20) = z(45) Then
        v(20) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture20.Picture = LoadPicture(z(46))
    If z(20) = z(46) Then
        v(20) = v(20) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture20.Picture = LoadPicture(z(47))
    If z(20) = z(47) Then
       v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture20.Picture = LoadPicture(z(48))
    If z(20) = z(48) Then
        v(20) = v(20) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture20.Picture = LoadPicture(z(49))
    If z(20) = z(49) Then
      v(20) = v(20) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture20.Picture = LoadPicture(z(50))
    If z(20) = z(50) Then
      v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture20.Picture = LoadPicture(z(51))
    If z(20) = z(51) Then
        v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture20.Picture = LoadPicture(z(52))
    If z(20) = z(52) Then
        v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture20.Picture = LoadPicture(z(53))
    If z(20) = z(53) Then
      v(20) = v(20) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture20.Picture = LoadPicture(z(54))
    If z(20) = z(54) Then
     v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture20.Picture = LoadPicture(z(55))
    If z(20) = z(55) Then
        v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture20.Picture = LoadPicture(z(56))
    If z(20) = z(56) Then
        v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture20.Picture = LoadPicture(z(57))
    If z(20) = z(57) Then
       v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture20.Picture = LoadPicture(z(58))
    If z(20) = z(58) Then
       v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture20.Picture = LoadPicture(z(59))
    If z(20) = z(59) Then
       v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture20.Picture = LoadPicture(z(60))
    If z(20) = z(60) Then
        v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture20.Picture = LoadPicture(z(61))
    If z(20) = z(61) Then
       v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture20.Picture = LoadPicture(z(62))
    If z(20) = z(62) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture20.Picture = LoadPicture(z(63))
    If z(20) = z(63) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture20.Picture = LoadPicture(z(64))
    If z(20) = z(64) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture20.Picture = LoadPicture(z(65))
    If z(20) = z(65) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture20.Picture = LoadPicture(z(66))
    If z(20) = z(66) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture20.Picture = LoadPicture(z(67))
    If z(20) = z(67) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture20.Picture = LoadPicture(z(68))
    If z(20) = z(68) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture20.Picture = LoadPicture(z(69))
    If z(20) = z(69) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture20.Picture = LoadPicture(z(70))
    If z(20) = z(70) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture20.Picture = LoadPicture(z(71))
    If z(20) = z(71) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture20.Picture = LoadPicture(z(72))
    If z(20) = z(72) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture20.Picture = LoadPicture(z(73))
    If z(20) = z(73) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture20.Picture = LoadPicture(z(74))
    If z(20) = z(74) Then
        v(20) = v(20) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture20.Picture = LoadPicture(z(75))
    If z(20) = z(75) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture20.Picture = LoadPicture(z(76))
    If z(20) = z(76) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture20.Picture = LoadPicture(z(77))
    If z(20) = z(77) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture20.Picture = LoadPicture(z(78))
    If z(20) = z(78) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture20.Picture = LoadPicture(z(79))
    If z(20) = z(79) Then
        v(20) = v(20) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture20.Picture = LoadPicture(z(80))
    If z(20) = z(80) Then
        v(20) = v(20) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture20.Picture = LoadPicture(z(81))
    If z(20) = z(81) Then
        v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture20.Picture = LoadPicture(z(82))
    If z(20) = z(82) Then
        v(20) = v(20) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture20.Picture = LoadPicture(z(83))
    If z(20) = z(83) Then
        v(20) = v(20) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture20.Picture = LoadPicture(z(84))
    If z(20) = z(84) Then
        v(20) = v(20) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture20.Picture = LoadPicture(z(85))
    If z(20) = z(85) Then
        v(20) = v(20) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture20.Picture = LoadPicture(z(86))
    If z(20) = z(86) Then
        v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture20.Picture = LoadPicture(z(87))
    If z(20) = z(87) Then
        v(20) = v(20) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture20.Picture = LoadPicture(z(88))
    If z(20) = z(88) Then
        v(20) = v(20) + 1
        End If
      
End If
End Sub

Private Sub Picture21_DblClick()
w = App.Path & "\clear.jpg"
Picture21.Picture = LoadPicture(w)
End Sub

Private Sub Picture21_DragDrop(Source As Control, x As Single, y As Single)
v(21) = 0
If Source = Picture45 Then
    Picture21.Picture = LoadPicture(z(45))
    If z(21) = z(45) Then
        v(21) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture21.Picture = LoadPicture(z(46))
    If z(21) = z(46) Then
        v(21) = v(21) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture21.Picture = LoadPicture(z(47))
    If z(21) = z(47) Then
       v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture21.Picture = LoadPicture(z(48))
    If z(21) = z(48) Then
        v(21) = v(21) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture21.Picture = LoadPicture(z(49))
    If z(21) = z(49) Then
      v(21) = v(21) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture21.Picture = LoadPicture(z(50))
    If z(21) = z(50) Then
      v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture21.Picture = LoadPicture(z(51))
    If z(21) = z(51) Then
        v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture21.Picture = LoadPicture(z(52))
    If z(21) = z(52) Then
        v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture21.Picture = LoadPicture(z(53))
    If z(21) = z(53) Then
      v(21) = v(21) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture21.Picture = LoadPicture(z(54))
    If z(21) = z(54) Then
     v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture21.Picture = LoadPicture(z(55))
    If z(21) = z(55) Then
        v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture21.Picture = LoadPicture(z(56))
    If z(21) = z(56) Then
        v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture21.Picture = LoadPicture(z(57))
    If z(21) = z(57) Then
       v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture21.Picture = LoadPicture(z(58))
    If z(21) = z(58) Then
       v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture21.Picture = LoadPicture(z(59))
    If z(21) = z(59) Then
       v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture21.Picture = LoadPicture(z(60))
    If z(21) = z(60) Then
        v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture21.Picture = LoadPicture(z(61))
    If z(21) = z(61) Then
       v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture21.Picture = LoadPicture(z(62))
    If z(21) = z(62) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture21.Picture = LoadPicture(z(63))
    If z(21) = z(63) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture21.Picture = LoadPicture(z(64))
    If z(21) = z(64) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture21.Picture = LoadPicture(z(65))
    If z(21) = z(65) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture21.Picture = LoadPicture(z(66))
    If z(21) = z(66) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture21.Picture = LoadPicture(z(67))
    If z(21) = z(67) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture21.Picture = LoadPicture(z(68))
    If z(21) = z(68) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture21.Picture = LoadPicture(z(69))
    If z(21) = z(69) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture21.Picture = LoadPicture(z(70))
    If z(21) = z(70) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture21.Picture = LoadPicture(z(71))
    If z(21) = z(71) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture21.Picture = LoadPicture(z(72))
    If z(21) = z(72) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture21.Picture = LoadPicture(z(73))
    If z(21) = z(73) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture21.Picture = LoadPicture(z(74))
    If z(21) = z(74) Then
        v(21) = v(21) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture21.Picture = LoadPicture(z(75))
    If z(21) = z(75) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture21.Picture = LoadPicture(z(76))
    If z(21) = z(76) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture21.Picture = LoadPicture(z(77))
    If z(21) = z(77) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture21.Picture = LoadPicture(z(78))
    If z(21) = z(78) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture21.Picture = LoadPicture(z(79))
    If z(21) = z(79) Then
        v(21) = v(21) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture21.Picture = LoadPicture(z(80))
    If z(21) = z(80) Then
        v(21) = v(21) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture21.Picture = LoadPicture(z(81))
    If z(21) = z(81) Then
        v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture21.Picture = LoadPicture(z(82))
    If z(21) = z(82) Then
        v(21) = v(21) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture21.Picture = LoadPicture(z(83))
    If z(21) = z(83) Then
        v(21) = v(21) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture21.Picture = LoadPicture(z(84))
    If z(21) = z(84) Then
        v(21) = v(21) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture21.Picture = LoadPicture(z(85))
    If z(21) = z(85) Then
        v(21) = v(21) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture21.Picture = LoadPicture(z(86))
    If z(21) = z(86) Then
        v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture21.Picture = LoadPicture(z(87))
    If z(21) = z(87) Then
        v(21) = v(21) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture21.Picture = LoadPicture(z(88))
    If z(21) = z(88) Then
        v(21) = v(21) + 1
        End If
      
End If

End Sub

Private Sub Picture22_DblClick()
w = App.Path & "\clear.jpg"
Picture22.Picture = LoadPicture(w)
End Sub

Private Sub Picture22_DragDrop(Source As Control, x As Single, y As Single)
v(22) = 0
If Source = Picture45 Then
    Picture22.Picture = LoadPicture(z(45))
    If z(22) = z(45) Then
        v(22) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture22.Picture = LoadPicture(z(46))
    If z(22) = z(46) Then
        v(22) = v(22) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture22.Picture = LoadPicture(z(47))
    If z(22) = z(47) Then
       v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture22.Picture = LoadPicture(z(48))
    If z(22) = z(48) Then
        v(22) = v(22) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture22.Picture = LoadPicture(z(49))
    If z(22) = z(49) Then
      v(22) = v(22) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture22.Picture = LoadPicture(z(50))
    If z(22) = z(50) Then
      v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture22.Picture = LoadPicture(z(51))
    If z(22) = z(51) Then
        v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture22.Picture = LoadPicture(z(52))
    If z(22) = z(52) Then
        v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture22.Picture = LoadPicture(z(53))
    If z(22) = z(53) Then
      v(22) = v(22) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture22.Picture = LoadPicture(z(54))
    If z(22) = z(54) Then
     v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture22.Picture = LoadPicture(z(55))
    If z(22) = z(55) Then
        v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture22.Picture = LoadPicture(z(56))
    If z(22) = z(56) Then
        v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture22.Picture = LoadPicture(z(57))
    If z(22) = z(57) Then
       v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture22.Picture = LoadPicture(z(58))
    If z(22) = z(58) Then
       v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture22.Picture = LoadPicture(z(59))
    If z(22) = z(59) Then
       v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture22.Picture = LoadPicture(z(60))
    If z(22) = z(60) Then
        v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture22.Picture = LoadPicture(z(61))
    If z(22) = z(61) Then
       v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture22.Picture = LoadPicture(z(62))
    If z(22) = z(62) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture22.Picture = LoadPicture(z(63))
    If z(22) = z(63) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture22.Picture = LoadPicture(z(64))
    If z(22) = z(64) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture22.Picture = LoadPicture(z(65))
    If z(22) = z(65) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture22.Picture = LoadPicture(z(66))
    If z(22) = z(66) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture22.Picture = LoadPicture(z(67))
    If z(22) = z(67) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture22.Picture = LoadPicture(z(68))
    If z(22) = z(68) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture22.Picture = LoadPicture(z(69))
    If z(22) = z(69) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture22.Picture = LoadPicture(z(70))
    If z(22) = z(70) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture22.Picture = LoadPicture(z(71))
    If z(22) = z(71) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture22.Picture = LoadPicture(z(72))
    If z(22) = z(72) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture22.Picture = LoadPicture(z(73))
    If z(22) = z(73) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture22.Picture = LoadPicture(z(74))
    If z(22) = z(74) Then
        v(22) = v(22) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture22.Picture = LoadPicture(z(75))
    If z(22) = z(75) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture22.Picture = LoadPicture(z(76))
    If z(22) = z(76) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture22.Picture = LoadPicture(z(77))
    If z(22) = z(77) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture22.Picture = LoadPicture(z(78))
    If z(22) = z(78) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture22.Picture = LoadPicture(z(79))
    If z(22) = z(79) Then
        v(22) = v(22) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture22.Picture = LoadPicture(z(80))
    If z(22) = z(80) Then
        v(22) = v(22) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture22.Picture = LoadPicture(z(81))
    If z(22) = z(81) Then
        v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture22.Picture = LoadPicture(z(82))
    If z(22) = z(82) Then
        v(22) = v(22) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture22.Picture = LoadPicture(z(83))
    If z(22) = z(83) Then
        v(22) = v(22) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture22.Picture = LoadPicture(z(84))
    If z(22) = z(84) Then
        v(22) = v(22) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture22.Picture = LoadPicture(z(85))
    If z(22) = z(85) Then
        v(22) = v(22) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture22.Picture = LoadPicture(z(86))
    If z(22) = z(86) Then
        v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture22.Picture = LoadPicture(z(87))
    If z(22) = z(87) Then
        v(22) = v(22) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture22.Picture = LoadPicture(z(88))
    If z(22) = z(88) Then
        v(22) = v(22) + 1
        End If
      
End If

End Sub

Private Sub Picture23_DblClick()
w = App.Path & "\clear.jpg"
Picture23.Picture = LoadPicture(w)
End Sub

Private Sub Picture23_DragDrop(Source As Control, x As Single, y As Single)
v(23) = 0
If Source = Picture45 Then
    Picture23.Picture = LoadPicture(z(45))
    If z(23) = z(45) Then
        v(23) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture23.Picture = LoadPicture(z(46))
    If z(23) = z(46) Then
        v(23) = v(23) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture23.Picture = LoadPicture(z(47))
    If z(23) = z(47) Then
       v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture23.Picture = LoadPicture(z(48))
    If z(23) = z(48) Then
        v(23) = v(23) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture23.Picture = LoadPicture(z(49))
    If z(23) = z(49) Then
      v(23) = v(23) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture23.Picture = LoadPicture(z(50))
    If z(23) = z(50) Then
      v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture23.Picture = LoadPicture(z(51))
    If z(23) = z(51) Then
        v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture23.Picture = LoadPicture(z(52))
    If z(23) = z(52) Then
        v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture23.Picture = LoadPicture(z(53))
    If z(23) = z(53) Then
      v(23) = v(23) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture23.Picture = LoadPicture(z(54))
    If z(23) = z(54) Then
     v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture23.Picture = LoadPicture(z(55))
    If z(23) = z(55) Then
        v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture23.Picture = LoadPicture(z(56))
    If z(23) = z(56) Then
        v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture23.Picture = LoadPicture(z(57))
    If z(23) = z(57) Then
       v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture23.Picture = LoadPicture(z(58))
    If z(23) = z(58) Then
       v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture23.Picture = LoadPicture(z(59))
    If z(23) = z(59) Then
       v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture23.Picture = LoadPicture(z(60))
    If z(23) = z(60) Then
        v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture23.Picture = LoadPicture(z(61))
    If z(23) = z(61) Then
       v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture23.Picture = LoadPicture(z(62))
    If z(23) = z(62) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture23.Picture = LoadPicture(z(63))
    If z(23) = z(63) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture23.Picture = LoadPicture(z(64))
    If z(23) = z(64) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture23.Picture = LoadPicture(z(65))
    If z(23) = z(65) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture23.Picture = LoadPicture(z(66))
    If z(23) = z(66) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture23.Picture = LoadPicture(z(67))
    If z(23) = z(67) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture23.Picture = LoadPicture(z(68))
    If z(23) = z(68) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture23.Picture = LoadPicture(z(69))
    If z(23) = z(69) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture23.Picture = LoadPicture(z(70))
    If z(23) = z(70) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture23.Picture = LoadPicture(z(71))
    If z(23) = z(71) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture23.Picture = LoadPicture(z(72))
    If z(23) = z(72) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture23.Picture = LoadPicture(z(73))
    If z(23) = z(73) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture23.Picture = LoadPicture(z(74))
    If z(23) = z(74) Then
        v(23) = v(23) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture23.Picture = LoadPicture(z(75))
    If z(23) = z(75) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture23.Picture = LoadPicture(z(76))
    If z(23) = z(76) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture23.Picture = LoadPicture(z(77))
    If z(23) = z(77) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture23.Picture = LoadPicture(z(78))
    If z(23) = z(78) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture23.Picture = LoadPicture(z(79))
    If z(23) = z(79) Then
        v(23) = v(23) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture23.Picture = LoadPicture(z(80))
    If z(23) = z(80) Then
        v(23) = v(23) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture23.Picture = LoadPicture(z(81))
    If z(23) = z(81) Then
        v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture23.Picture = LoadPicture(z(82))
    If z(23) = z(82) Then
        v(23) = v(23) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture23.Picture = LoadPicture(z(83))
    If z(23) = z(83) Then
        v(23) = v(23) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture23.Picture = LoadPicture(z(84))
    If z(23) = z(84) Then
        v(23) = v(23) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture23.Picture = LoadPicture(z(85))
    If z(23) = z(85) Then
        v(23) = v(23) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture23.Picture = LoadPicture(z(86))
    If z(23) = z(86) Then
        v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture23.Picture = LoadPicture(z(87))
    If z(23) = z(87) Then
        v(23) = v(23) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture23.Picture = LoadPicture(z(88))
    If z(23) = z(88) Then
        v(23) = v(23) + 1
        End If
      
End If

End Sub

Private Sub Picture24_DblClick()
w = App.Path & "\clear.jpg"
Picture24.Picture = LoadPicture(w)
End Sub

Private Sub Picture24_DragDrop(Source As Control, x As Single, y As Single)
v(24) = 0
If Source = Picture45 Then
    Picture24.Picture = LoadPicture(z(45))
    If z(24) = z(45) Then
        v(24) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture24.Picture = LoadPicture(z(46))
    If z(24) = z(46) Then
        v(24) = v(24) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture24.Picture = LoadPicture(z(47))
    If z(24) = z(47) Then
       v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture24.Picture = LoadPicture(z(48))
    If z(24) = z(48) Then
        v(24) = v(24) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture24.Picture = LoadPicture(z(49))
    If z(24) = z(49) Then
      v(24) = v(24) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture24.Picture = LoadPicture(z(50))
    If z(24) = z(50) Then
      v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture24.Picture = LoadPicture(z(51))
    If z(24) = z(51) Then
        v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture24.Picture = LoadPicture(z(52))
    If z(24) = z(52) Then
        v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture24.Picture = LoadPicture(z(53))
    If z(24) = z(53) Then
      v(24) = v(24) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture24.Picture = LoadPicture(z(54))
    If z(24) = z(54) Then
     v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture24.Picture = LoadPicture(z(55))
    If z(24) = z(55) Then
        v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture24.Picture = LoadPicture(z(56))
    If z(24) = z(56) Then
        v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture24.Picture = LoadPicture(z(57))
    If z(24) = z(57) Then
       v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture24.Picture = LoadPicture(z(58))
    If z(24) = z(58) Then
       v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture24.Picture = LoadPicture(z(59))
    If z(24) = z(59) Then
       v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture24.Picture = LoadPicture(z(60))
    If z(24) = z(60) Then
        v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture24.Picture = LoadPicture(z(61))
    If z(24) = z(61) Then
       v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture24.Picture = LoadPicture(z(62))
    If z(24) = z(62) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture24.Picture = LoadPicture(z(63))
    If z(24) = z(63) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture24.Picture = LoadPicture(z(64))
    If z(24) = z(64) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture24.Picture = LoadPicture(z(65))
    If z(24) = z(65) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture24.Picture = LoadPicture(z(66))
    If z(24) = z(66) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture24.Picture = LoadPicture(z(67))
    If z(24) = z(67) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture24.Picture = LoadPicture(z(68))
    If z(24) = z(68) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture24.Picture = LoadPicture(z(69))
    If z(24) = z(69) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture24.Picture = LoadPicture(z(70))
    If z(24) = z(70) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture24.Picture = LoadPicture(z(71))
    If z(24) = z(71) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture24.Picture = LoadPicture(z(72))
    If z(24) = z(72) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture24.Picture = LoadPicture(z(73))
    If z(24) = z(73) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture24.Picture = LoadPicture(z(74))
    If z(24) = z(74) Then
        v(24) = v(24) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture24.Picture = LoadPicture(z(75))
    If z(24) = z(75) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture24.Picture = LoadPicture(z(76))
    If z(24) = z(76) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture24.Picture = LoadPicture(z(77))
    If z(24) = z(77) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture24.Picture = LoadPicture(z(78))
    If z(24) = z(78) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture24.Picture = LoadPicture(z(79))
    If z(24) = z(79) Then
        v(24) = v(24) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture24.Picture = LoadPicture(z(80))
    If z(24) = z(80) Then
        v(24) = v(24) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture24.Picture = LoadPicture(z(81))
    If z(24) = z(81) Then
        v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture24.Picture = LoadPicture(z(82))
    If z(24) = z(82) Then
        v(24) = v(24) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture24.Picture = LoadPicture(z(83))
    If z(24) = z(83) Then
        v(24) = v(24) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture24.Picture = LoadPicture(z(84))
    If z(24) = z(84) Then
        v(24) = v(24) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture24.Picture = LoadPicture(z(85))
    If z(24) = z(85) Then
        v(24) = v(24) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture24.Picture = LoadPicture(z(86))
    If z(24) = z(86) Then
        v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture24.Picture = LoadPicture(z(87))
    If z(24) = z(87) Then
        v(24) = v(24) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture24.Picture = LoadPicture(z(88))
    If z(24) = z(88) Then
        v(24) = v(24) + 1
        End If
      
End If

End Sub

Private Sub Picture25_DblClick()
w = App.Path & "\clear.jpg"
Picture25.Picture = LoadPicture(w)
End Sub

Private Sub Picture25_DragDrop(Source As Control, x As Single, y As Single)
v(25) = 0
If Source = Picture45 Then
    Picture25.Picture = LoadPicture(z(45))
    If z(25) = z(45) Then
        v(25) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture25.Picture = LoadPicture(z(46))
    If z(25) = z(46) Then
        v(25) = v(25) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture25.Picture = LoadPicture(z(47))
    If z(25) = z(47) Then
       v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture25.Picture = LoadPicture(z(48))
    If z(25) = z(48) Then
        v(25) = v(25) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture25.Picture = LoadPicture(z(49))
    If z(25) = z(49) Then
      v(25) = v(25) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture25.Picture = LoadPicture(z(50))
    If z(25) = z(50) Then
      v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture25.Picture = LoadPicture(z(51))
    If z(25) = z(51) Then
        v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture25.Picture = LoadPicture(z(52))
    If z(25) = z(52) Then
        v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture25.Picture = LoadPicture(z(53))
    If z(25) = z(53) Then
      v(25) = v(25) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture25.Picture = LoadPicture(z(54))
    If z(25) = z(54) Then
     v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture25.Picture = LoadPicture(z(55))
    If z(25) = z(55) Then
        v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture25.Picture = LoadPicture(z(56))
    If z(25) = z(56) Then
        v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture25.Picture = LoadPicture(z(57))
    If z(25) = z(57) Then
       v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture25.Picture = LoadPicture(z(58))
    If z(25) = z(58) Then
       v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture25.Picture = LoadPicture(z(59))
    If z(25) = z(59) Then
       v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture25.Picture = LoadPicture(z(60))
    If z(25) = z(60) Then
        v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture25.Picture = LoadPicture(z(61))
    If z(25) = z(61) Then
       v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture25.Picture = LoadPicture(z(62))
    If z(25) = z(62) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture25.Picture = LoadPicture(z(63))
    If z(25) = z(63) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture25.Picture = LoadPicture(z(64))
    If z(25) = z(64) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture25.Picture = LoadPicture(z(65))
    If z(25) = z(65) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture25.Picture = LoadPicture(z(66))
    If z(25) = z(66) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture25.Picture = LoadPicture(z(67))
    If z(25) = z(67) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture25.Picture = LoadPicture(z(68))
    If z(25) = z(68) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture25.Picture = LoadPicture(z(69))
    If z(25) = z(69) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture25.Picture = LoadPicture(z(70))
    If z(25) = z(70) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture25.Picture = LoadPicture(z(71))
    If z(25) = z(71) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture25.Picture = LoadPicture(z(72))
    If z(25) = z(72) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture25.Picture = LoadPicture(z(73))
    If z(25) = z(73) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture25.Picture = LoadPicture(z(74))
    If z(25) = z(74) Then
        v(25) = v(25) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture25.Picture = LoadPicture(z(75))
    If z(25) = z(75) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture25.Picture = LoadPicture(z(76))
    If z(25) = z(76) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture25.Picture = LoadPicture(z(77))
    If z(25) = z(77) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture25.Picture = LoadPicture(z(78))
    If z(25) = z(78) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture25.Picture = LoadPicture(z(79))
    If z(25) = z(79) Then
        v(25) = v(25) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture25.Picture = LoadPicture(z(80))
    If z(25) = z(80) Then
        v(25) = v(25) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture25.Picture = LoadPicture(z(81))
    If z(25) = z(81) Then
        v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture25.Picture = LoadPicture(z(82))
    If z(25) = z(82) Then
        v(25) = v(25) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture25.Picture = LoadPicture(z(83))
    If z(25) = z(83) Then
        v(25) = v(25) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture25.Picture = LoadPicture(z(84))
    If z(25) = z(84) Then
        v(25) = v(25) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture25.Picture = LoadPicture(z(85))
    If z(25) = z(85) Then
        v(25) = v(25) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture25.Picture = LoadPicture(z(86))
    If z(25) = z(86) Then
        v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture25.Picture = LoadPicture(z(87))
    If z(25) = z(87) Then
        v(25) = v(25) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture25.Picture = LoadPicture(z(88))
    If z(25) = z(88) Then
        v(25) = v(25) + 1
        End If
      
End If

End Sub

Private Sub Picture26_DblClick()
w = App.Path & "\clear.jpg"
Picture26.Picture = LoadPicture(w)
End Sub

Private Sub Picture26_DragDrop(Source As Control, x As Single, y As Single)
v(26) = 0
If Source = Picture45 Then
    Picture26.Picture = LoadPicture(z(45))
    If z(26) = z(45) Then
        v(26) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture26.Picture = LoadPicture(z(46))
    If z(26) = z(46) Then
        v(26) = v(26) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture26.Picture = LoadPicture(z(47))
    If z(26) = z(47) Then
       v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture26.Picture = LoadPicture(z(48))
    If z(26) = z(48) Then
        v(26) = v(26) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture26.Picture = LoadPicture(z(49))
    If z(26) = z(49) Then
      v(26) = v(26) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture26.Picture = LoadPicture(z(50))
    If z(26) = z(50) Then
      v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture26.Picture = LoadPicture(z(51))
    If z(26) = z(51) Then
        v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture26.Picture = LoadPicture(z(52))
    If z(26) = z(52) Then
        v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture26.Picture = LoadPicture(z(53))
    If z(26) = z(53) Then
      v(26) = v(26) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture26.Picture = LoadPicture(z(54))
    If z(26) = z(54) Then
     v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture26.Picture = LoadPicture(z(55))
    If z(26) = z(55) Then
        v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture26.Picture = LoadPicture(z(56))
    If z(26) = z(56) Then
        v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture26.Picture = LoadPicture(z(57))
    If z(26) = z(57) Then
       v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture26.Picture = LoadPicture(z(58))
    If z(26) = z(58) Then
       v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture26.Picture = LoadPicture(z(59))
    If z(26) = z(59) Then
       v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture26.Picture = LoadPicture(z(60))
    If z(26) = z(60) Then
        v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture26.Picture = LoadPicture(z(61))
    If z(26) = z(61) Then
       v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture26.Picture = LoadPicture(z(62))
    If z(26) = z(62) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture26.Picture = LoadPicture(z(63))
    If z(26) = z(63) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture26.Picture = LoadPicture(z(64))
    If z(26) = z(64) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture26.Picture = LoadPicture(z(65))
    If z(26) = z(65) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture26.Picture = LoadPicture(z(66))
    If z(26) = z(66) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture26.Picture = LoadPicture(z(67))
    If z(26) = z(67) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture26.Picture = LoadPicture(z(68))
    If z(26) = z(68) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture26.Picture = LoadPicture(z(69))
    If z(26) = z(69) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture26.Picture = LoadPicture(z(70))
    If z(26) = z(70) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture26.Picture = LoadPicture(z(71))
    If z(26) = z(71) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture26.Picture = LoadPicture(z(72))
    If z(26) = z(72) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture26.Picture = LoadPicture(z(73))
    If z(26) = z(73) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture26.Picture = LoadPicture(z(74))
    If z(26) = z(74) Then
        v(26) = v(26) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture26.Picture = LoadPicture(z(75))
    If z(26) = z(75) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture26.Picture = LoadPicture(z(76))
    If z(26) = z(76) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture26.Picture = LoadPicture(z(77))
    If z(26) = z(77) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture26.Picture = LoadPicture(z(78))
    If z(26) = z(78) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture26.Picture = LoadPicture(z(79))
    If z(26) = z(79) Then
        v(26) = v(26) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture26.Picture = LoadPicture(z(80))
    If z(26) = z(80) Then
        v(26) = v(26) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture26.Picture = LoadPicture(z(81))
    If z(26) = z(81) Then
        v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture26.Picture = LoadPicture(z(82))
    If z(26) = z(82) Then
        v(26) = v(26) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture26.Picture = LoadPicture(z(83))
    If z(26) = z(83) Then
        v(26) = v(26) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture26.Picture = LoadPicture(z(84))
    If z(26) = z(84) Then
        v(26) = v(26) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture26.Picture = LoadPicture(z(85))
    If z(26) = z(85) Then
        v(26) = v(26) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture26.Picture = LoadPicture(z(86))
    If z(26) = z(86) Then
        v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture26.Picture = LoadPicture(z(87))
    If z(26) = z(87) Then
        v(26) = v(26) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture26.Picture = LoadPicture(z(88))
    If z(26) = z(88) Then
        v(26) = v(26) + 1
        End If
      
End If

End Sub

Private Sub Picture27_DblClick()
w = App.Path & "\clear.jpg"
Picture27.Picture = LoadPicture(w)
End Sub

Private Sub Picture27_DragDrop(Source As Control, x As Single, y As Single)
v(27) = 0
If Source = Picture45 Then
    Picture27.Picture = LoadPicture(z(45))
    If z(27) = z(45) Then
        v(27) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture27.Picture = LoadPicture(z(46))
    If z(27) = z(46) Then
        v(27) = v(27) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture27.Picture = LoadPicture(z(47))
    If z(27) = z(47) Then
       v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture27.Picture = LoadPicture(z(48))
    If z(27) = z(48) Then
        v(27) = v(27) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture27.Picture = LoadPicture(z(49))
    If z(27) = z(49) Then
      v(27) = v(27) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture27.Picture = LoadPicture(z(50))
    If z(27) = z(50) Then
      v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture27.Picture = LoadPicture(z(51))
    If z(27) = z(51) Then
        v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture27.Picture = LoadPicture(z(52))
    If z(27) = z(52) Then
        v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture27.Picture = LoadPicture(z(53))
    If z(27) = z(53) Then
      v(27) = v(27) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture27.Picture = LoadPicture(z(54))
    If z(27) = z(54) Then
     v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture27.Picture = LoadPicture(z(55))
    If z(27) = z(55) Then
        v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture27.Picture = LoadPicture(z(56))
    If z(27) = z(56) Then
        v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture27.Picture = LoadPicture(z(57))
    If z(27) = z(57) Then
       v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture27.Picture = LoadPicture(z(58))
    If z(27) = z(58) Then
       v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture27.Picture = LoadPicture(z(59))
    If z(27) = z(59) Then
       v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture27.Picture = LoadPicture(z(60))
    If z(27) = z(60) Then
        v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture27.Picture = LoadPicture(z(61))
    If z(27) = z(61) Then
       v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture27.Picture = LoadPicture(z(62))
    If z(27) = z(62) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture27.Picture = LoadPicture(z(63))
    If z(27) = z(63) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture27.Picture = LoadPicture(z(64))
    If z(27) = z(64) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture27.Picture = LoadPicture(z(65))
    If z(27) = z(65) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture27.Picture = LoadPicture(z(66))
    If z(27) = z(66) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture27.Picture = LoadPicture(z(67))
    If z(27) = z(67) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture27.Picture = LoadPicture(z(68))
    If z(27) = z(68) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture27.Picture = LoadPicture(z(69))
    If z(27) = z(69) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture27.Picture = LoadPicture(z(70))
    If z(27) = z(70) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture27.Picture = LoadPicture(z(71))
    If z(27) = z(71) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture27.Picture = LoadPicture(z(72))
    If z(27) = z(72) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture27.Picture = LoadPicture(z(73))
    If z(27) = z(73) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture27.Picture = LoadPicture(z(74))
    If z(27) = z(74) Then
        v(27) = v(27) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture27.Picture = LoadPicture(z(75))
    If z(27) = z(75) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture27.Picture = LoadPicture(z(76))
    If z(27) = z(76) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture27.Picture = LoadPicture(z(77))
    If z(27) = z(77) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture27.Picture = LoadPicture(z(78))
    If z(27) = z(78) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture27.Picture = LoadPicture(z(79))
    If z(27) = z(79) Then
        v(27) = v(27) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture27.Picture = LoadPicture(z(80))
    If z(27) = z(80) Then
        v(27) = v(27) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture27.Picture = LoadPicture(z(81))
    If z(27) = z(81) Then
        v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture27.Picture = LoadPicture(z(82))
    If z(27) = z(82) Then
        v(27) = v(27) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture27.Picture = LoadPicture(z(83))
    If z(27) = z(83) Then
        v(27) = v(27) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture27.Picture = LoadPicture(z(84))
    If z(27) = z(84) Then
        v(27) = v(27) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture27.Picture = LoadPicture(z(85))
    If z(27) = z(85) Then
        v(27) = v(27) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture27.Picture = LoadPicture(z(86))
    If z(27) = z(86) Then
        v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture27.Picture = LoadPicture(z(87))
    If z(27) = z(87) Then
        v(27) = v(27) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture27.Picture = LoadPicture(z(88))
    If z(27) = z(88) Then
        v(27) = v(27) + 1
        End If
      
End If

End Sub

Private Sub Picture28_DblClick()
w = App.Path & "\clear.jpg"
Picture28.Picture = LoadPicture(w)
End Sub

Private Sub Picture28_DragDrop(Source As Control, x As Single, y As Single)
v(28) = 0
If Source = Picture45 Then
    Picture28.Picture = LoadPicture(z(45))
    If z(28) = z(45) Then
        v(28) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture28.Picture = LoadPicture(z(46))
    If z(28) = z(46) Then
        v(28) = v(28) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture28.Picture = LoadPicture(z(47))
    If z(28) = z(47) Then
       v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture28.Picture = LoadPicture(z(48))
    If z(28) = z(48) Then
        v(28) = v(28) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture28.Picture = LoadPicture(z(49))
    If z(28) = z(49) Then
      v(28) = v(28) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture28.Picture = LoadPicture(z(50))
    If z(28) = z(50) Then
      v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture28.Picture = LoadPicture(z(51))
    If z(28) = z(51) Then
        v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture28.Picture = LoadPicture(z(52))
    If z(28) = z(52) Then
        v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture28.Picture = LoadPicture(z(53))
    If z(28) = z(53) Then
      v(28) = v(28) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture28.Picture = LoadPicture(z(54))
    If z(28) = z(54) Then
     v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture28.Picture = LoadPicture(z(55))
    If z(28) = z(55) Then
        v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture28.Picture = LoadPicture(z(56))
    If z(28) = z(56) Then
        v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture28.Picture = LoadPicture(z(57))
    If z(28) = z(57) Then
       v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture28.Picture = LoadPicture(z(58))
    If z(28) = z(58) Then
       v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture28.Picture = LoadPicture(z(59))
    If z(28) = z(59) Then
       v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture28.Picture = LoadPicture(z(60))
    If z(28) = z(60) Then
        v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture28.Picture = LoadPicture(z(61))
    If z(28) = z(61) Then
       v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture28.Picture = LoadPicture(z(62))
    If z(28) = z(62) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture28.Picture = LoadPicture(z(63))
    If z(28) = z(63) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture28.Picture = LoadPicture(z(64))
    If z(28) = z(64) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture28.Picture = LoadPicture(z(65))
    If z(28) = z(65) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture28.Picture = LoadPicture(z(66))
    If z(28) = z(66) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture28.Picture = LoadPicture(z(67))
    If z(28) = z(67) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture28.Picture = LoadPicture(z(68))
    If z(28) = z(68) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture28.Picture = LoadPicture(z(69))
    If z(28) = z(69) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture28.Picture = LoadPicture(z(70))
    If z(28) = z(70) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture28.Picture = LoadPicture(z(71))
    If z(28) = z(71) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture28.Picture = LoadPicture(z(72))
    If z(28) = z(72) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture28.Picture = LoadPicture(z(73))
    If z(28) = z(73) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture28.Picture = LoadPicture(z(74))
    If z(28) = z(74) Then
        v(28) = v(28) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture28.Picture = LoadPicture(z(75))
    If z(28) = z(75) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture28.Picture = LoadPicture(z(76))
    If z(28) = z(76) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture28.Picture = LoadPicture(z(77))
    If z(28) = z(77) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture28.Picture = LoadPicture(z(78))
    If z(28) = z(78) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture28.Picture = LoadPicture(z(79))
    If z(28) = z(79) Then
        v(28) = v(28) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture28.Picture = LoadPicture(z(80))
    If z(28) = z(80) Then
        v(28) = v(28) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture28.Picture = LoadPicture(z(81))
    If z(28) = z(81) Then
        v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture28.Picture = LoadPicture(z(82))
    If z(28) = z(82) Then
        v(28) = v(28) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture28.Picture = LoadPicture(z(83))
    If z(28) = z(83) Then
        v(28) = v(28) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture28.Picture = LoadPicture(z(84))
    If z(28) = z(84) Then
        v(28) = v(28) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture28.Picture = LoadPicture(z(85))
    If z(28) = z(85) Then
        v(28) = v(28) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture28.Picture = LoadPicture(z(86))
    If z(28) = z(86) Then
        v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture28.Picture = LoadPicture(z(87))
    If z(28) = z(87) Then
        v(28) = v(28) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture28.Picture = LoadPicture(z(88))
    If z(28) = z(88) Then
        v(28) = v(28) + 1
        End If
      
End If

End Sub

Private Sub Picture29_DblClick()
w = App.Path & "\clear.jpg"
Picture29.Picture = LoadPicture(w)
End Sub

Private Sub Picture29_DragDrop(Source As Control, x As Single, y As Single)
v(29) = 0
If Source = Picture45 Then
    Picture29.Picture = LoadPicture(z(45))
    If z(29) = z(45) Then
        v(29) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture29.Picture = LoadPicture(z(46))
    If z(29) = z(46) Then
        v(29) = v(29) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture29.Picture = LoadPicture(z(47))
    If z(29) = z(47) Then
       v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture29.Picture = LoadPicture(z(48))
    If z(29) = z(48) Then
        v(29) = v(29) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture29.Picture = LoadPicture(z(49))
    If z(29) = z(49) Then
      v(29) = v(29) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture29.Picture = LoadPicture(z(50))
    If z(29) = z(50) Then
      v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture29.Picture = LoadPicture(z(51))
    If z(29) = z(51) Then
        v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture29.Picture = LoadPicture(z(52))
    If z(29) = z(52) Then
        v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture29.Picture = LoadPicture(z(53))
    If z(29) = z(53) Then
      v(29) = v(29) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture29.Picture = LoadPicture(z(54))
    If z(29) = z(54) Then
     v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture29.Picture = LoadPicture(z(55))
    If z(29) = z(55) Then
        v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture29.Picture = LoadPicture(z(56))
    If z(29) = z(56) Then
        v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture29.Picture = LoadPicture(z(57))
    If z(29) = z(57) Then
       v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture29.Picture = LoadPicture(z(58))
    If z(29) = z(58) Then
       v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture29.Picture = LoadPicture(z(59))
    If z(29) = z(59) Then
       v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture29.Picture = LoadPicture(z(60))
    If z(29) = z(60) Then
        v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture29.Picture = LoadPicture(z(61))
    If z(29) = z(61) Then
       v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture29.Picture = LoadPicture(z(62))
    If z(29) = z(62) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture29.Picture = LoadPicture(z(63))
    If z(29) = z(63) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture29.Picture = LoadPicture(z(64))
    If z(29) = z(64) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture29.Picture = LoadPicture(z(65))
    If z(29) = z(65) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture29.Picture = LoadPicture(z(66))
    If z(29) = z(66) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture29.Picture = LoadPicture(z(67))
    If z(29) = z(67) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture29.Picture = LoadPicture(z(68))
    If z(29) = z(68) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture29.Picture = LoadPicture(z(69))
    If z(29) = z(69) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture29.Picture = LoadPicture(z(70))
    If z(29) = z(70) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture29.Picture = LoadPicture(z(71))
    If z(29) = z(71) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture29.Picture = LoadPicture(z(72))
    If z(29) = z(72) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture29.Picture = LoadPicture(z(73))
    If z(29) = z(73) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture29.Picture = LoadPicture(z(74))
    If z(29) = z(74) Then
        v(29) = v(29) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture29.Picture = LoadPicture(z(75))
    If z(29) = z(75) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture29.Picture = LoadPicture(z(76))
    If z(29) = z(76) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture29.Picture = LoadPicture(z(77))
    If z(29) = z(77) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture29.Picture = LoadPicture(z(78))
    If z(29) = z(78) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture29.Picture = LoadPicture(z(79))
    If z(29) = z(79) Then
        v(29) = v(29) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture29.Picture = LoadPicture(z(80))
    If z(29) = z(80) Then
        v(29) = v(29) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture29.Picture = LoadPicture(z(81))
    If z(29) = z(81) Then
        v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture29.Picture = LoadPicture(z(82))
    If z(29) = z(82) Then
        v(29) = v(29) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture29.Picture = LoadPicture(z(83))
    If z(29) = z(83) Then
        v(29) = v(29) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture29.Picture = LoadPicture(z(84))
    If z(29) = z(84) Then
        v(29) = v(29) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture29.Picture = LoadPicture(z(85))
    If z(29) = z(85) Then
        v(29) = v(29) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture29.Picture = LoadPicture(z(86))
    If z(29) = z(86) Then
        v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture29.Picture = LoadPicture(z(87))
    If z(29) = z(87) Then
        v(29) = v(29) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture29.Picture = LoadPicture(z(88))
    If z(29) = z(88) Then
        v(29) = v(29) + 1
        End If
      
End If

End Sub

Private Sub Picture3_DblClick()
w = App.Path & "\clear.jpg"
Picture3.Picture = LoadPicture(w)
End Sub

Private Sub Picture3_DragDrop(Source As Control, x As Single, y As Single)
v(3) = 0
If Source = Picture45 Then
    Picture3.Picture = LoadPicture(z(45))
    If z(3) = z(45) Then
        v(3) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture3.Picture = LoadPicture(z(46))
    If z(3) = z(46) Then
        v(3) = v(3) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture3.Picture = LoadPicture(z(47))
    If z(3) = z(47) Then
       v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture3.Picture = LoadPicture(z(48))
    If z(3) = z(48) Then
        v(3) = v(3) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture3.Picture = LoadPicture(z(49))
    If z(3) = z(49) Then
      v(3) = v(3) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture3.Picture = LoadPicture(z(50))
    If z(3) = z(50) Then
      v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture3.Picture = LoadPicture(z(51))
    If z(3) = z(51) Then
        v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture3.Picture = LoadPicture(z(52))
    If z(3) = z(52) Then
        v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture3.Picture = LoadPicture(z(53))
    If z(3) = z(53) Then
      v(3) = v(3) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture3.Picture = LoadPicture(z(54))
    If z(3) = z(54) Then
     v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture3.Picture = LoadPicture(z(55))
    If z(3) = z(55) Then
        v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture3.Picture = LoadPicture(z(56))
    If z(3) = z(56) Then
        v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture3.Picture = LoadPicture(z(57))
    If z(3) = z(57) Then
       v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture3.Picture = LoadPicture(z(58))
    If z(3) = z(58) Then
       v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture3.Picture = LoadPicture(z(59))
    If z(3) = z(59) Then
       v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture3.Picture = LoadPicture(z(60))
    If z(3) = z(60) Then
        v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture3.Picture = LoadPicture(z(61))
    If z(3) = z(61) Then
       v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture3.Picture = LoadPicture(z(62))
    If z(3) = z(62) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture3.Picture = LoadPicture(z(63))
    If z(3) = z(63) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture3.Picture = LoadPicture(z(64))
    If z(3) = z(64) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture3.Picture = LoadPicture(z(65))
    If z(3) = z(65) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture3.Picture = LoadPicture(z(66))
    If z(3) = z(66) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture3.Picture = LoadPicture(z(67))
    If z(3) = z(67) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture3.Picture = LoadPicture(z(68))
    If z(3) = z(68) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture3.Picture = LoadPicture(z(69))
    If z(3) = z(69) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture3.Picture = LoadPicture(z(70))
    If z(3) = z(70) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture3.Picture = LoadPicture(z(71))
    If z(3) = z(71) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture3.Picture = LoadPicture(z(72))
    If z(3) = z(72) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture3.Picture = LoadPicture(z(73))
    If z(3) = z(73) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture3.Picture = LoadPicture(z(74))
    If z(3) = z(74) Then
        v(3) = v(3) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture3.Picture = LoadPicture(z(75))
    If z(3) = z(75) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture3.Picture = LoadPicture(z(76))
    If z(3) = z(76) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture3.Picture = LoadPicture(z(77))
    If z(3) = z(77) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture3.Picture = LoadPicture(z(78))
    If z(3) = z(78) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture3.Picture = LoadPicture(z(79))
    If z(3) = z(79) Then
        v(3) = v(3) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture3.Picture = LoadPicture(z(80))
    If z(3) = z(80) Then
        v(3) = v(3) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture3.Picture = LoadPicture(z(81))
    If z(3) = z(81) Then
        v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture3.Picture = LoadPicture(z(82))
    If z(3) = z(82) Then
        v(3) = v(3) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture3.Picture = LoadPicture(z(83))
    If z(3) = z(83) Then
        v(3) = v(3) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture3.Picture = LoadPicture(z(84))
    If z(3) = z(84) Then
        v(3) = v(3) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture3.Picture = LoadPicture(z(85))
    If z(3) = z(85) Then
        v(3) = v(3) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture3.Picture = LoadPicture(z(86))
    If z(3) = z(86) Then
        v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture3.Picture = LoadPicture(z(87))
    If z(3) = z(87) Then
        v(3) = v(3) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture3.Picture = LoadPicture(z(88))
    If z(3) = z(88) Then
        v(3) = v(3) + 1
        End If
      
End If
End Sub

Private Sub Picture30_DblClick()
w = App.Path & "\clear.jpg"
Picture30.Picture = LoadPicture(w)
End Sub

Private Sub Picture30_DragDrop(Source As Control, x As Single, y As Single)
v(30) = 0
If Source = Picture45 Then
    Picture30.Picture = LoadPicture(z(45))
    If z(30) = z(45) Then
        v(30) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture30.Picture = LoadPicture(z(46))
    If z(30) = z(46) Then
        v(30) = v(30) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture30.Picture = LoadPicture(z(47))
    If z(30) = z(47) Then
       v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture30.Picture = LoadPicture(z(48))
    If z(30) = z(48) Then
        v(30) = v(30) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture30.Picture = LoadPicture(z(49))
    If z(30) = z(49) Then
      v(30) = v(30) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture30.Picture = LoadPicture(z(50))
    If z(30) = z(50) Then
      v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture30.Picture = LoadPicture(z(51))
    If z(30) = z(51) Then
        v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture30.Picture = LoadPicture(z(52))
    If z(30) = z(52) Then
        v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture30.Picture = LoadPicture(z(53))
    If z(30) = z(53) Then
      v(30) = v(30) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture30.Picture = LoadPicture(z(54))
    If z(30) = z(54) Then
     v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture30.Picture = LoadPicture(z(55))
    If z(30) = z(55) Then
        v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture30.Picture = LoadPicture(z(56))
    If z(30) = z(56) Then
        v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture30.Picture = LoadPicture(z(57))
    If z(30) = z(57) Then
       v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture30.Picture = LoadPicture(z(58))
    If z(30) = z(58) Then
       v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture30.Picture = LoadPicture(z(59))
    If z(30) = z(59) Then
       v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture30.Picture = LoadPicture(z(60))
    If z(30) = z(60) Then
        v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture30.Picture = LoadPicture(z(61))
    If z(30) = z(61) Then
       v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture30.Picture = LoadPicture(z(62))
    If z(30) = z(62) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture30.Picture = LoadPicture(z(63))
    If z(30) = z(63) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture30.Picture = LoadPicture(z(64))
    If z(30) = z(64) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture30.Picture = LoadPicture(z(65))
    If z(30) = z(65) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture30.Picture = LoadPicture(z(66))
    If z(30) = z(66) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture30.Picture = LoadPicture(z(67))
    If z(30) = z(67) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture30.Picture = LoadPicture(z(68))
    If z(30) = z(68) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture30.Picture = LoadPicture(z(69))
    If z(30) = z(69) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture30.Picture = LoadPicture(z(70))
    If z(30) = z(70) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture30.Picture = LoadPicture(z(71))
    If z(30) = z(71) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture30.Picture = LoadPicture(z(72))
    If z(30) = z(72) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture30.Picture = LoadPicture(z(73))
    If z(30) = z(73) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture30.Picture = LoadPicture(z(74))
    If z(30) = z(74) Then
        v(30) = v(30) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture30.Picture = LoadPicture(z(75))
    If z(30) = z(75) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture30.Picture = LoadPicture(z(76))
    If z(30) = z(76) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture30.Picture = LoadPicture(z(77))
    If z(30) = z(77) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture30.Picture = LoadPicture(z(78))
    If z(30) = z(78) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture30.Picture = LoadPicture(z(79))
    If z(30) = z(79) Then
        v(30) = v(30) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture30.Picture = LoadPicture(z(80))
    If z(30) = z(80) Then
        v(30) = v(30) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture30.Picture = LoadPicture(z(81))
    If z(30) = z(81) Then
        v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture30.Picture = LoadPicture(z(82))
    If z(30) = z(82) Then
        v(30) = v(30) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture30.Picture = LoadPicture(z(83))
    If z(30) = z(83) Then
        v(30) = v(30) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture30.Picture = LoadPicture(z(84))
    If z(30) = z(84) Then
        v(30) = v(30) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture30.Picture = LoadPicture(z(85))
    If z(30) = z(85) Then
        v(30) = v(30) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture30.Picture = LoadPicture(z(86))
    If z(30) = z(86) Then
        v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture30.Picture = LoadPicture(z(87))
    If z(30) = z(87) Then
        v(30) = v(30) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture30.Picture = LoadPicture(z(88))
    If z(30) = z(88) Then
        v(30) = v(30) + 1
        End If
      
End If

End Sub

Private Sub Picture31_DblClick()
w = App.Path & "\clear.jpg"
Picture31.Picture = LoadPicture(w)
End Sub

Private Sub Picture31_DragDrop(Source As Control, x As Single, y As Single)
v(31) = 0
If Source = Picture45 Then
    Picture31.Picture = LoadPicture(z(45))
    If z(31) = z(45) Then
        v(31) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture31.Picture = LoadPicture(z(46))
    If z(31) = z(46) Then
        v(31) = v(31) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture31.Picture = LoadPicture(z(47))
    If z(31) = z(47) Then
       v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture31.Picture = LoadPicture(z(48))
    If z(31) = z(48) Then
        v(31) = v(31) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture31.Picture = LoadPicture(z(49))
    If z(31) = z(49) Then
      v(31) = v(31) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture31.Picture = LoadPicture(z(50))
    If z(31) = z(50) Then
      v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture31.Picture = LoadPicture(z(51))
    If z(31) = z(51) Then
        v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture31.Picture = LoadPicture(z(52))
    If z(31) = z(52) Then
        v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture31.Picture = LoadPicture(z(53))
    If z(31) = z(53) Then
      v(31) = v(31) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture31.Picture = LoadPicture(z(54))
    If z(31) = z(54) Then
     v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture31.Picture = LoadPicture(z(55))
    If z(31) = z(55) Then
        v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture31.Picture = LoadPicture(z(56))
    If z(31) = z(56) Then
        v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture31.Picture = LoadPicture(z(57))
    If z(31) = z(57) Then
       v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture31.Picture = LoadPicture(z(58))
    If z(31) = z(58) Then
       v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture31.Picture = LoadPicture(z(59))
    If z(31) = z(59) Then
       v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture31.Picture = LoadPicture(z(60))
    If z(31) = z(60) Then
        v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture31.Picture = LoadPicture(z(61))
    If z(31) = z(61) Then
       v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture31.Picture = LoadPicture(z(62))
    If z(31) = z(62) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture31.Picture = LoadPicture(z(63))
    If z(31) = z(63) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture31.Picture = LoadPicture(z(64))
    If z(31) = z(64) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture31.Picture = LoadPicture(z(65))
    If z(31) = z(65) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture31.Picture = LoadPicture(z(66))
    If z(31) = z(66) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture31.Picture = LoadPicture(z(67))
    If z(31) = z(67) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture31.Picture = LoadPicture(z(68))
    If z(31) = z(68) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture31.Picture = LoadPicture(z(69))
    If z(31) = z(69) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture31.Picture = LoadPicture(z(70))
    If z(31) = z(70) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture31.Picture = LoadPicture(z(71))
    If z(31) = z(71) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture31.Picture = LoadPicture(z(72))
    If z(31) = z(72) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture31.Picture = LoadPicture(z(73))
    If z(31) = z(73) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture31.Picture = LoadPicture(z(74))
    If z(31) = z(74) Then
        v(31) = v(31) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture31.Picture = LoadPicture(z(75))
    If z(31) = z(75) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture31.Picture = LoadPicture(z(76))
    If z(31) = z(76) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture31.Picture = LoadPicture(z(77))
    If z(31) = z(77) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture31.Picture = LoadPicture(z(78))
    If z(31) = z(78) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture31.Picture = LoadPicture(z(79))
    If z(31) = z(79) Then
        v(31) = v(31) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture31.Picture = LoadPicture(z(80))
    If z(31) = z(80) Then
        v(31) = v(31) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture31.Picture = LoadPicture(z(81))
    If z(31) = z(81) Then
        v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture31.Picture = LoadPicture(z(82))
    If z(31) = z(82) Then
        v(31) = v(31) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture31.Picture = LoadPicture(z(83))
    If z(31) = z(83) Then
        v(31) = v(31) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture31.Picture = LoadPicture(z(84))
    If z(31) = z(84) Then
        v(31) = v(31) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture31.Picture = LoadPicture(z(85))
    If z(31) = z(85) Then
        v(31) = v(31) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture31.Picture = LoadPicture(z(86))
    If z(31) = z(86) Then
        v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture31.Picture = LoadPicture(z(87))
    If z(31) = z(87) Then
        v(31) = v(31) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture31.Picture = LoadPicture(z(88))
    If z(31) = z(88) Then
        v(31) = v(31) + 1
        End If
      
End If

End Sub

Private Sub Picture32_DblClick()
w = App.Path & "\clear.jpg"
Picture32.Picture = LoadPicture(w)
End Sub

Private Sub Picture32_DragDrop(Source As Control, x As Single, y As Single)
v(32) = 0
If Source = Picture45 Then
    Picture32.Picture = LoadPicture(z(45))
    If z(32) = z(45) Then
        v(32) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture32.Picture = LoadPicture(z(46))
    If z(32) = z(46) Then
        v(32) = v(32) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture32.Picture = LoadPicture(z(47))
    If z(32) = z(47) Then
       v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture32.Picture = LoadPicture(z(48))
    If z(32) = z(48) Then
        v(32) = v(32) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture32.Picture = LoadPicture(z(49))
    If z(32) = z(49) Then
      v(32) = v(32) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture32.Picture = LoadPicture(z(50))
    If z(32) = z(50) Then
      v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture32.Picture = LoadPicture(z(51))
    If z(32) = z(51) Then
        v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture32.Picture = LoadPicture(z(52))
    If z(32) = z(52) Then
        v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture32.Picture = LoadPicture(z(53))
    If z(32) = z(53) Then
      v(32) = v(32) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture32.Picture = LoadPicture(z(54))
    If z(32) = z(54) Then
     v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture32.Picture = LoadPicture(z(55))
    If z(32) = z(55) Then
        v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture32.Picture = LoadPicture(z(56))
    If z(32) = z(56) Then
        v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture32.Picture = LoadPicture(z(57))
    If z(32) = z(57) Then
       v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture32.Picture = LoadPicture(z(58))
    If z(32) = z(58) Then
       v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture32.Picture = LoadPicture(z(59))
    If z(32) = z(59) Then
       v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture32.Picture = LoadPicture(z(60))
    If z(32) = z(60) Then
        v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture32.Picture = LoadPicture(z(61))
    If z(32) = z(61) Then
       v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture32.Picture = LoadPicture(z(62))
    If z(32) = z(62) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture32.Picture = LoadPicture(z(63))
    If z(32) = z(63) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture32.Picture = LoadPicture(z(64))
    If z(32) = z(64) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture32.Picture = LoadPicture(z(65))
    If z(32) = z(65) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture32.Picture = LoadPicture(z(66))
    If z(32) = z(66) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture32.Picture = LoadPicture(z(67))
    If z(32) = z(67) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture32.Picture = LoadPicture(z(68))
    If z(32) = z(68) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture32.Picture = LoadPicture(z(69))
    If z(32) = z(69) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture32.Picture = LoadPicture(z(70))
    If z(32) = z(70) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture32.Picture = LoadPicture(z(71))
    If z(32) = z(71) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture32.Picture = LoadPicture(z(72))
    If z(32) = z(72) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture32.Picture = LoadPicture(z(73))
    If z(32) = z(73) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture32.Picture = LoadPicture(z(74))
    If z(32) = z(74) Then
        v(32) = v(32) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture32.Picture = LoadPicture(z(75))
    If z(32) = z(75) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture32.Picture = LoadPicture(z(76))
    If z(32) = z(76) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture32.Picture = LoadPicture(z(77))
    If z(32) = z(77) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture32.Picture = LoadPicture(z(78))
    If z(32) = z(78) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture32.Picture = LoadPicture(z(79))
    If z(32) = z(79) Then
        v(32) = v(32) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture32.Picture = LoadPicture(z(80))
    If z(32) = z(80) Then
        v(32) = v(32) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture32.Picture = LoadPicture(z(81))
    If z(32) = z(81) Then
        v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture32.Picture = LoadPicture(z(82))
    If z(32) = z(82) Then
        v(32) = v(32) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture32.Picture = LoadPicture(z(83))
    If z(32) = z(83) Then
        v(32) = v(32) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture32.Picture = LoadPicture(z(84))
    If z(32) = z(84) Then
        v(32) = v(32) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture32.Picture = LoadPicture(z(85))
    If z(32) = z(85) Then
        v(32) = v(32) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture32.Picture = LoadPicture(z(86))
    If z(32) = z(86) Then
        v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture32.Picture = LoadPicture(z(87))
    If z(32) = z(87) Then
        v(32) = v(32) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture32.Picture = LoadPicture(z(88))
    If z(32) = z(88) Then
        v(32) = v(32) + 1
        End If
      
End If

End Sub

Private Sub Picture33_DblClick()
w = App.Path & "\clear.jpg"
Picture33.Picture = LoadPicture(w)
End Sub

Private Sub Picture33_DragDrop(Source As Control, x As Single, y As Single)
v(33) = 0
If Source = Picture45 Then
    Picture33.Picture = LoadPicture(z(45))
    If z(33) = z(45) Then
        v(33) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture33.Picture = LoadPicture(z(46))
    If z(33) = z(46) Then
        v(33) = v(33) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture33.Picture = LoadPicture(z(47))
    If z(33) = z(47) Then
       v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture33.Picture = LoadPicture(z(48))
    If z(33) = z(48) Then
        v(33) = v(33) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture33.Picture = LoadPicture(z(49))
    If z(33) = z(49) Then
      v(33) = v(33) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture33.Picture = LoadPicture(z(50))
    If z(33) = z(50) Then
      v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture33.Picture = LoadPicture(z(51))
    If z(33) = z(51) Then
        v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture33.Picture = LoadPicture(z(52))
    If z(33) = z(52) Then
        v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture33.Picture = LoadPicture(z(53))
    If z(33) = z(53) Then
      v(33) = v(33) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture33.Picture = LoadPicture(z(54))
    If z(33) = z(54) Then
     v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture33.Picture = LoadPicture(z(55))
    If z(33) = z(55) Then
        v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture33.Picture = LoadPicture(z(56))
    If z(33) = z(56) Then
        v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture33.Picture = LoadPicture(z(57))
    If z(33) = z(57) Then
       v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture33.Picture = LoadPicture(z(58))
    If z(33) = z(58) Then
       v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture33.Picture = LoadPicture(z(59))
    If z(33) = z(59) Then
       v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture33.Picture = LoadPicture(z(60))
    If z(33) = z(60) Then
        v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture33.Picture = LoadPicture(z(61))
    If z(33) = z(61) Then
       v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture33.Picture = LoadPicture(z(62))
    If z(33) = z(62) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture33.Picture = LoadPicture(z(63))
    If z(33) = z(63) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture33.Picture = LoadPicture(z(64))
    If z(33) = z(64) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture33.Picture = LoadPicture(z(65))
    If z(33) = z(65) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture33.Picture = LoadPicture(z(66))
    If z(33) = z(66) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture33.Picture = LoadPicture(z(67))
    If z(33) = z(67) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture33.Picture = LoadPicture(z(68))
    If z(33) = z(68) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture33.Picture = LoadPicture(z(69))
    If z(33) = z(69) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture33.Picture = LoadPicture(z(70))
    If z(33) = z(70) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture33.Picture = LoadPicture(z(71))
    If z(33) = z(71) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture33.Picture = LoadPicture(z(72))
    If z(33) = z(72) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture33.Picture = LoadPicture(z(73))
    If z(33) = z(73) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture33.Picture = LoadPicture(z(74))
    If z(33) = z(74) Then
        v(33) = v(33) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture33.Picture = LoadPicture(z(75))
    If z(33) = z(75) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture33.Picture = LoadPicture(z(76))
    If z(33) = z(76) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture33.Picture = LoadPicture(z(77))
    If z(33) = z(77) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture33.Picture = LoadPicture(z(78))
    If z(33) = z(78) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture33.Picture = LoadPicture(z(79))
    If z(33) = z(79) Then
        v(33) = v(33) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture33.Picture = LoadPicture(z(80))
    If z(33) = z(80) Then
        v(33) = v(33) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture33.Picture = LoadPicture(z(81))
    If z(33) = z(81) Then
        v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture33.Picture = LoadPicture(z(82))
    If z(33) = z(82) Then
        v(33) = v(33) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture33.Picture = LoadPicture(z(83))
    If z(33) = z(83) Then
        v(33) = v(33) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture33.Picture = LoadPicture(z(84))
    If z(33) = z(84) Then
        v(33) = v(33) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture33.Picture = LoadPicture(z(85))
    If z(33) = z(85) Then
        v(33) = v(33) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture33.Picture = LoadPicture(z(86))
    If z(33) = z(86) Then
        v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture33.Picture = LoadPicture(z(87))
    If z(33) = z(87) Then
        v(33) = v(33) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture33.Picture = LoadPicture(z(88))
    If z(33) = z(88) Then
        v(33) = v(33) + 1
        End If
      
End If

End Sub

Private Sub Picture34_DblClick()
w = App.Path & "\clear.jpg"
Picture34.Picture = LoadPicture(w)
End Sub

Private Sub Picture34_DragDrop(Source As Control, x As Single, y As Single)
v(34) = 0
If Source = Picture45 Then
    Picture34.Picture = LoadPicture(z(45))
    If z(34) = z(45) Then
        v(34) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture34.Picture = LoadPicture(z(46))
    If z(34) = z(46) Then
        v(34) = v(34) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture34.Picture = LoadPicture(z(47))
    If z(34) = z(47) Then
       v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture34.Picture = LoadPicture(z(48))
    If z(34) = z(48) Then
        v(34) = v(34) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture34.Picture = LoadPicture(z(49))
    If z(34) = z(49) Then
      v(34) = v(34) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture34.Picture = LoadPicture(z(50))
    If z(34) = z(50) Then
      v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture34.Picture = LoadPicture(z(51))
    If z(34) = z(51) Then
        v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture34.Picture = LoadPicture(z(52))
    If z(34) = z(52) Then
        v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture34.Picture = LoadPicture(z(53))
    If z(34) = z(53) Then
      v(34) = v(34) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture34.Picture = LoadPicture(z(54))
    If z(34) = z(54) Then
     v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture34.Picture = LoadPicture(z(55))
    If z(34) = z(55) Then
        v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture34.Picture = LoadPicture(z(56))
    If z(34) = z(56) Then
        v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture34.Picture = LoadPicture(z(57))
    If z(34) = z(57) Then
       v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture34.Picture = LoadPicture(z(58))
    If z(34) = z(58) Then
       v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture34.Picture = LoadPicture(z(59))
    If z(34) = z(59) Then
       v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture34.Picture = LoadPicture(z(60))
    If z(34) = z(60) Then
        v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture34.Picture = LoadPicture(z(61))
    If z(34) = z(61) Then
       v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture34.Picture = LoadPicture(z(62))
    If z(34) = z(62) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture34.Picture = LoadPicture(z(63))
    If z(34) = z(63) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture34.Picture = LoadPicture(z(64))
    If z(34) = z(64) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture34.Picture = LoadPicture(z(65))
    If z(34) = z(65) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture34.Picture = LoadPicture(z(66))
    If z(34) = z(66) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture34.Picture = LoadPicture(z(67))
    If z(34) = z(67) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture34.Picture = LoadPicture(z(68))
    If z(34) = z(68) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture34.Picture = LoadPicture(z(69))
    If z(34) = z(69) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture34.Picture = LoadPicture(z(70))
    If z(34) = z(70) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture34.Picture = LoadPicture(z(71))
    If z(34) = z(71) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture34.Picture = LoadPicture(z(72))
    If z(34) = z(72) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture34.Picture = LoadPicture(z(73))
    If z(34) = z(73) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture34.Picture = LoadPicture(z(74))
    If z(34) = z(74) Then
        v(34) = v(34) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture34.Picture = LoadPicture(z(75))
    If z(34) = z(75) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture34.Picture = LoadPicture(z(76))
    If z(34) = z(76) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture34.Picture = LoadPicture(z(77))
    If z(34) = z(77) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture34.Picture = LoadPicture(z(78))
    If z(34) = z(78) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture34.Picture = LoadPicture(z(79))
    If z(34) = z(79) Then
        v(34) = v(34) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture34.Picture = LoadPicture(z(80))
    If z(34) = z(80) Then
        v(34) = v(34) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture34.Picture = LoadPicture(z(81))
    If z(34) = z(81) Then
        v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture34.Picture = LoadPicture(z(82))
    If z(34) = z(82) Then
        v(34) = v(34) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture34.Picture = LoadPicture(z(83))
    If z(34) = z(83) Then
        v(34) = v(34) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture34.Picture = LoadPicture(z(84))
    If z(34) = z(84) Then
        v(34) = v(34) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture34.Picture = LoadPicture(z(85))
    If z(34) = z(85) Then
        v(34) = v(34) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture34.Picture = LoadPicture(z(86))
    If z(34) = z(86) Then
        v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture34.Picture = LoadPicture(z(87))
    If z(34) = z(87) Then
        v(34) = v(34) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture34.Picture = LoadPicture(z(88))
    If z(34) = z(88) Then
        v(34) = v(34) + 1
        End If
      
End If

End Sub

Private Sub Picture35_DblClick()
w = App.Path & "\clear.jpg"
Picture35.Picture = LoadPicture(w)
End Sub

Private Sub Picture35_DragDrop(Source As Control, x As Single, y As Single)
v(35) = 0
If Source = Picture45 Then
    Picture35.Picture = LoadPicture(z(45))
    If z(35) = z(45) Then
        v(35) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture35.Picture = LoadPicture(z(46))
    If z(35) = z(46) Then
        v(35) = v(35) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture35.Picture = LoadPicture(z(47))
    If z(35) = z(47) Then
       v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture35.Picture = LoadPicture(z(48))
    If z(35) = z(48) Then
        v(35) = v(35) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture35.Picture = LoadPicture(z(49))
    If z(35) = z(49) Then
      v(35) = v(35) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture35.Picture = LoadPicture(z(50))
    If z(35) = z(50) Then
      v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture35.Picture = LoadPicture(z(51))
    If z(35) = z(51) Then
        v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture35.Picture = LoadPicture(z(52))
    If z(35) = z(52) Then
        v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture35.Picture = LoadPicture(z(53))
    If z(35) = z(53) Then
      v(35) = v(35) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture35.Picture = LoadPicture(z(54))
    If z(35) = z(54) Then
     v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture35.Picture = LoadPicture(z(55))
    If z(35) = z(55) Then
        v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture35.Picture = LoadPicture(z(56))
    If z(35) = z(56) Then
        v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture35.Picture = LoadPicture(z(57))
    If z(35) = z(57) Then
       v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture35.Picture = LoadPicture(z(58))
    If z(35) = z(58) Then
       v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture35.Picture = LoadPicture(z(59))
    If z(35) = z(59) Then
       v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture35.Picture = LoadPicture(z(60))
    If z(35) = z(60) Then
        v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture35.Picture = LoadPicture(z(61))
    If z(35) = z(61) Then
       v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture35.Picture = LoadPicture(z(62))
    If z(35) = z(62) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture35.Picture = LoadPicture(z(63))
    If z(35) = z(63) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture35.Picture = LoadPicture(z(64))
    If z(35) = z(64) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture35.Picture = LoadPicture(z(65))
    If z(35) = z(65) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture35.Picture = LoadPicture(z(66))
    If z(35) = z(66) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture35.Picture = LoadPicture(z(67))
    If z(35) = z(67) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture35.Picture = LoadPicture(z(68))
    If z(35) = z(68) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture35.Picture = LoadPicture(z(69))
    If z(35) = z(69) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture35.Picture = LoadPicture(z(70))
    If z(35) = z(70) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture35.Picture = LoadPicture(z(71))
    If z(35) = z(71) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture35.Picture = LoadPicture(z(72))
    If z(35) = z(72) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture35.Picture = LoadPicture(z(73))
    If z(35) = z(73) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture35.Picture = LoadPicture(z(74))
    If z(35) = z(74) Then
        v(35) = v(35) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture35.Picture = LoadPicture(z(75))
    If z(35) = z(75) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture35.Picture = LoadPicture(z(76))
    If z(35) = z(76) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture35.Picture = LoadPicture(z(77))
    If z(35) = z(77) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture35.Picture = LoadPicture(z(78))
    If z(35) = z(78) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture35.Picture = LoadPicture(z(79))
    If z(35) = z(79) Then
        v(35) = v(35) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture35.Picture = LoadPicture(z(80))
    If z(35) = z(80) Then
        v(35) = v(35) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture35.Picture = LoadPicture(z(81))
    If z(35) = z(81) Then
        v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture35.Picture = LoadPicture(z(82))
    If z(35) = z(82) Then
        v(35) = v(35) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture35.Picture = LoadPicture(z(83))
    If z(35) = z(83) Then
        v(35) = v(35) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture35.Picture = LoadPicture(z(84))
    If z(35) = z(84) Then
        v(35) = v(35) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture35.Picture = LoadPicture(z(85))
    If z(35) = z(85) Then
        v(35) = v(35) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture35.Picture = LoadPicture(z(86))
    If z(35) = z(86) Then
        v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture35.Picture = LoadPicture(z(87))
    If z(35) = z(87) Then
        v(35) = v(35) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture35.Picture = LoadPicture(z(88))
    If z(35) = z(88) Then
        v(35) = v(35) + 1
        End If
      
End If

End Sub

Private Sub Picture36_DblClick()
w = App.Path & "\clear.jpg"
Picture36.Picture = LoadPicture(w)
End Sub

Private Sub Picture36_DragDrop(Source As Control, x As Single, y As Single)
v(36) = 0
If Source = Picture45 Then
    Picture36.Picture = LoadPicture(z(45))
    If z(36) = z(45) Then
        v(36) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture36.Picture = LoadPicture(z(46))
    If z(36) = z(46) Then
        v(36) = v(36) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture36.Picture = LoadPicture(z(47))
    If z(36) = z(47) Then
       v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture36.Picture = LoadPicture(z(48))
    If z(36) = z(48) Then
        v(36) = v(36) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture36.Picture = LoadPicture(z(49))
    If z(36) = z(49) Then
      v(36) = v(36) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture36.Picture = LoadPicture(z(50))
    If z(36) = z(50) Then
      v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture36.Picture = LoadPicture(z(51))
    If z(36) = z(51) Then
        v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture36.Picture = LoadPicture(z(52))
    If z(36) = z(52) Then
        v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture36.Picture = LoadPicture(z(53))
    If z(36) = z(53) Then
      v(36) = v(36) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture36.Picture = LoadPicture(z(54))
    If z(36) = z(54) Then
     v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture36.Picture = LoadPicture(z(55))
    If z(36) = z(55) Then
        v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture36.Picture = LoadPicture(z(56))
    If z(36) = z(56) Then
        v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture36.Picture = LoadPicture(z(57))
    If z(36) = z(57) Then
       v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture36.Picture = LoadPicture(z(58))
    If z(36) = z(58) Then
       v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture36.Picture = LoadPicture(z(59))
    If z(36) = z(59) Then
       v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture36.Picture = LoadPicture(z(60))
    If z(36) = z(60) Then
        v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture36.Picture = LoadPicture(z(61))
    If z(36) = z(61) Then
       v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture36.Picture = LoadPicture(z(62))
    If z(36) = z(62) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture36.Picture = LoadPicture(z(63))
    If z(36) = z(63) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture36.Picture = LoadPicture(z(64))
    If z(36) = z(64) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture36.Picture = LoadPicture(z(65))
    If z(36) = z(65) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture36.Picture = LoadPicture(z(66))
    If z(36) = z(66) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture36.Picture = LoadPicture(z(67))
    If z(36) = z(67) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture36.Picture = LoadPicture(z(68))
    If z(36) = z(68) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture36.Picture = LoadPicture(z(69))
    If z(36) = z(69) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture36.Picture = LoadPicture(z(70))
    If z(36) = z(70) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture36.Picture = LoadPicture(z(71))
    If z(36) = z(71) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture36.Picture = LoadPicture(z(72))
    If z(36) = z(72) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture36.Picture = LoadPicture(z(73))
    If z(36) = z(73) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture36.Picture = LoadPicture(z(74))
    If z(36) = z(74) Then
        v(36) = v(36) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture36.Picture = LoadPicture(z(75))
    If z(36) = z(75) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture36.Picture = LoadPicture(z(76))
    If z(36) = z(76) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture36.Picture = LoadPicture(z(77))
    If z(36) = z(77) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture36.Picture = LoadPicture(z(78))
    If z(36) = z(78) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture36.Picture = LoadPicture(z(79))
    If z(36) = z(79) Then
        v(36) = v(36) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture36.Picture = LoadPicture(z(80))
    If z(36) = z(80) Then
        v(36) = v(36) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture36.Picture = LoadPicture(z(81))
    If z(36) = z(81) Then
        v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture36.Picture = LoadPicture(z(82))
    If z(36) = z(82) Then
        v(36) = v(36) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture36.Picture = LoadPicture(z(83))
    If z(36) = z(83) Then
        v(36) = v(36) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture36.Picture = LoadPicture(z(84))
    If z(36) = z(84) Then
        v(36) = v(36) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture36.Picture = LoadPicture(z(85))
    If z(36) = z(85) Then
        v(36) = v(36) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture36.Picture = LoadPicture(z(86))
    If z(36) = z(86) Then
        v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture36.Picture = LoadPicture(z(87))
    If z(36) = z(87) Then
        v(36) = v(36) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture36.Picture = LoadPicture(z(88))
    If z(36) = z(88) Then
        v(36) = v(36) + 1
        End If
      
End If

End Sub

Private Sub Picture37_DblClick()
w = App.Path & "\clear.jpg"
Picture37.Picture = LoadPicture(w)
End Sub

Private Sub Picture37_DragDrop(Source As Control, x As Single, y As Single)
v(37) = 0
If Source = Picture45 Then
    Picture37.Picture = LoadPicture(z(45))
    If z(37) = z(45) Then
        v(37) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture37.Picture = LoadPicture(z(46))
    If z(37) = z(46) Then
        v(37) = v(37) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture37.Picture = LoadPicture(z(47))
    If z(37) = z(47) Then
       v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture37.Picture = LoadPicture(z(48))
    If z(37) = z(48) Then
        v(37) = v(37) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture37.Picture = LoadPicture(z(49))
    If z(37) = z(49) Then
      v(37) = v(37) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture37.Picture = LoadPicture(z(50))
    If z(37) = z(50) Then
      v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture37.Picture = LoadPicture(z(51))
    If z(37) = z(51) Then
        v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture37.Picture = LoadPicture(z(52))
    If z(37) = z(52) Then
        v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture37.Picture = LoadPicture(z(53))
    If z(37) = z(53) Then
      v(37) = v(37) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture37.Picture = LoadPicture(z(54))
    If z(37) = z(54) Then
     v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture37.Picture = LoadPicture(z(55))
    If z(37) = z(55) Then
        v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture37.Picture = LoadPicture(z(56))
    If z(37) = z(56) Then
        v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture37.Picture = LoadPicture(z(57))
    If z(37) = z(57) Then
       v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture37.Picture = LoadPicture(z(58))
    If z(37) = z(58) Then
       v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture37.Picture = LoadPicture(z(59))
    If z(37) = z(59) Then
       v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture37.Picture = LoadPicture(z(60))
    If z(37) = z(60) Then
        v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture37.Picture = LoadPicture(z(61))
    If z(37) = z(61) Then
       v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture37.Picture = LoadPicture(z(62))
    If z(37) = z(62) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture37.Picture = LoadPicture(z(63))
    If z(37) = z(63) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture37.Picture = LoadPicture(z(64))
    If z(37) = z(64) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture37.Picture = LoadPicture(z(65))
    If z(37) = z(65) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture37.Picture = LoadPicture(z(66))
    If z(37) = z(66) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture37.Picture = LoadPicture(z(67))
    If z(37) = z(67) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture37.Picture = LoadPicture(z(68))
    If z(37) = z(68) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture37.Picture = LoadPicture(z(69))
    If z(37) = z(69) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture37.Picture = LoadPicture(z(70))
    If z(37) = z(70) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture37.Picture = LoadPicture(z(71))
    If z(37) = z(71) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture37.Picture = LoadPicture(z(72))
    If z(37) = z(72) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture37.Picture = LoadPicture(z(73))
    If z(37) = z(73) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture37.Picture = LoadPicture(z(74))
    If z(37) = z(74) Then
        v(37) = v(37) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture37.Picture = LoadPicture(z(75))
    If z(37) = z(75) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture37.Picture = LoadPicture(z(76))
    If z(37) = z(76) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture37.Picture = LoadPicture(z(77))
    If z(37) = z(77) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture37.Picture = LoadPicture(z(78))
    If z(37) = z(78) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture37.Picture = LoadPicture(z(79))
    If z(37) = z(79) Then
        v(37) = v(37) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture37.Picture = LoadPicture(z(80))
    If z(37) = z(80) Then
        v(37) = v(37) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture37.Picture = LoadPicture(z(81))
    If z(37) = z(81) Then
        v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture37.Picture = LoadPicture(z(82))
    If z(37) = z(82) Then
        v(37) = v(37) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture37.Picture = LoadPicture(z(83))
    If z(37) = z(83) Then
        v(37) = v(37) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture37.Picture = LoadPicture(z(84))
    If z(37) = z(84) Then
        v(37) = v(37) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture37.Picture = LoadPicture(z(85))
    If z(37) = z(85) Then
        v(37) = v(37) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture37.Picture = LoadPicture(z(86))
    If z(37) = z(86) Then
        v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture37.Picture = LoadPicture(z(87))
    If z(37) = z(87) Then
        v(37) = v(37) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture37.Picture = LoadPicture(z(88))
    If z(37) = z(88) Then
        v(37) = v(37) + 1
        End If
      
End If

End Sub

Private Sub Picture38_DblClick()
w = App.Path & "\clear.jpg"
Picture38.Picture = LoadPicture(w)
End Sub

Private Sub Picture38_DragDrop(Source As Control, x As Single, y As Single)
v(38) = 0
If Source = Picture45 Then
    Picture38.Picture = LoadPicture(z(45))
    If z(38) = z(45) Then
        v(38) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture38.Picture = LoadPicture(z(46))
    If z(38) = z(46) Then
        v(38) = v(38) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture38.Picture = LoadPicture(z(47))
    If z(38) = z(47) Then
       v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture38.Picture = LoadPicture(z(48))
    If z(38) = z(48) Then
        v(38) = v(38) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture38.Picture = LoadPicture(z(49))
    If z(38) = z(49) Then
      v(38) = v(38) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture38.Picture = LoadPicture(z(50))
    If z(38) = z(50) Then
      v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture38.Picture = LoadPicture(z(51))
    If z(38) = z(51) Then
        v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture38.Picture = LoadPicture(z(52))
    If z(38) = z(52) Then
        v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture38.Picture = LoadPicture(z(53))
    If z(38) = z(53) Then
      v(38) = v(38) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture38.Picture = LoadPicture(z(54))
    If z(38) = z(54) Then
     v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture38.Picture = LoadPicture(z(55))
    If z(38) = z(55) Then
        v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture38.Picture = LoadPicture(z(56))
    If z(38) = z(56) Then
        v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture38.Picture = LoadPicture(z(57))
    If z(38) = z(57) Then
       v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture38.Picture = LoadPicture(z(58))
    If z(38) = z(58) Then
       v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture38.Picture = LoadPicture(z(59))
    If z(38) = z(59) Then
       v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture38.Picture = LoadPicture(z(60))
    If z(38) = z(60) Then
        v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture38.Picture = LoadPicture(z(61))
    If z(38) = z(61) Then
       v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture38.Picture = LoadPicture(z(62))
    If z(38) = z(62) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture38.Picture = LoadPicture(z(63))
    If z(38) = z(63) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture38.Picture = LoadPicture(z(64))
    If z(38) = z(64) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture38.Picture = LoadPicture(z(65))
    If z(38) = z(65) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture38.Picture = LoadPicture(z(66))
    If z(38) = z(66) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture38.Picture = LoadPicture(z(67))
    If z(38) = z(67) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture38.Picture = LoadPicture(z(68))
    If z(38) = z(68) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture38.Picture = LoadPicture(z(69))
    If z(38) = z(69) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture38.Picture = LoadPicture(z(70))
    If z(38) = z(70) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture38.Picture = LoadPicture(z(71))
    If z(38) = z(71) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture38.Picture = LoadPicture(z(72))
    If z(38) = z(72) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture38.Picture = LoadPicture(z(73))
    If z(38) = z(73) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture38.Picture = LoadPicture(z(74))
    If z(38) = z(74) Then
        v(38) = v(38) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture38.Picture = LoadPicture(z(75))
    If z(38) = z(75) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture38.Picture = LoadPicture(z(76))
    If z(38) = z(76) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture38.Picture = LoadPicture(z(77))
    If z(38) = z(77) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture38.Picture = LoadPicture(z(78))
    If z(38) = z(78) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture38.Picture = LoadPicture(z(79))
    If z(38) = z(79) Then
        v(38) = v(38) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture38.Picture = LoadPicture(z(80))
    If z(38) = z(80) Then
        v(38) = v(38) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture38.Picture = LoadPicture(z(81))
    If z(38) = z(81) Then
        v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture38.Picture = LoadPicture(z(82))
    If z(38) = z(82) Then
        v(38) = v(38) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture38.Picture = LoadPicture(z(83))
    If z(38) = z(83) Then
        v(38) = v(38) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture38.Picture = LoadPicture(z(84))
    If z(38) = z(84) Then
        v(38) = v(38) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture38.Picture = LoadPicture(z(85))
    If z(38) = z(85) Then
        v(38) = v(38) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture38.Picture = LoadPicture(z(86))
    If z(38) = z(86) Then
        v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture38.Picture = LoadPicture(z(87))
    If z(38) = z(87) Then
        v(38) = v(38) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture38.Picture = LoadPicture(z(88))
    If z(38) = z(88) Then
        v(38) = v(38) + 1
        End If
      
End If

End Sub

Private Sub Picture39_DblClick()
w = App.Path & "\clear.jpg"
Picture39.Picture = LoadPicture(w)
End Sub

Private Sub Picture39_DragDrop(Source As Control, x As Single, y As Single)
v(39) = 0
If Source = Picture45 Then
    Picture39.Picture = LoadPicture(z(45))
    If z(39) = z(45) Then
        v(39) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture39.Picture = LoadPicture(z(46))
    If z(39) = z(46) Then
        v(39) = v(39) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture39.Picture = LoadPicture(z(47))
    If z(39) = z(47) Then
       v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture39.Picture = LoadPicture(z(48))
    If z(39) = z(48) Then
        v(39) = v(39) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture39.Picture = LoadPicture(z(49))
    If z(39) = z(49) Then
      v(39) = v(39) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture39.Picture = LoadPicture(z(50))
    If z(39) = z(50) Then
      v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture39.Picture = LoadPicture(z(51))
    If z(39) = z(51) Then
        v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture39.Picture = LoadPicture(z(52))
    If z(39) = z(52) Then
        v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture39.Picture = LoadPicture(z(53))
    If z(39) = z(53) Then
      v(39) = v(39) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture39.Picture = LoadPicture(z(54))
    If z(39) = z(54) Then
     v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture39.Picture = LoadPicture(z(55))
    If z(39) = z(55) Then
        v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture39.Picture = LoadPicture(z(56))
    If z(39) = z(56) Then
        v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture39.Picture = LoadPicture(z(57))
    If z(39) = z(57) Then
       v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture39.Picture = LoadPicture(z(58))
    If z(39) = z(58) Then
       v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture39.Picture = LoadPicture(z(59))
    If z(39) = z(59) Then
       v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture39.Picture = LoadPicture(z(60))
    If z(39) = z(60) Then
        v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture39.Picture = LoadPicture(z(61))
    If z(39) = z(61) Then
       v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture39.Picture = LoadPicture(z(62))
    If z(39) = z(62) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture39.Picture = LoadPicture(z(63))
    If z(39) = z(63) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture39.Picture = LoadPicture(z(64))
    If z(39) = z(64) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture39.Picture = LoadPicture(z(65))
    If z(39) = z(65) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture39.Picture = LoadPicture(z(66))
    If z(39) = z(66) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture39.Picture = LoadPicture(z(67))
    If z(39) = z(67) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture39.Picture = LoadPicture(z(68))
    If z(39) = z(68) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture39.Picture = LoadPicture(z(69))
    If z(39) = z(69) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture39.Picture = LoadPicture(z(70))
    If z(39) = z(70) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture39.Picture = LoadPicture(z(71))
    If z(39) = z(71) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture39.Picture = LoadPicture(z(72))
    If z(39) = z(72) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture39.Picture = LoadPicture(z(73))
    If z(39) = z(73) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture39.Picture = LoadPicture(z(74))
    If z(39) = z(74) Then
        v(39) = v(39) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture39.Picture = LoadPicture(z(75))
    If z(39) = z(75) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture39.Picture = LoadPicture(z(76))
    If z(39) = z(76) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture39.Picture = LoadPicture(z(77))
    If z(39) = z(77) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture39.Picture = LoadPicture(z(78))
    If z(39) = z(78) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture39.Picture = LoadPicture(z(79))
    If z(39) = z(79) Then
        v(39) = v(39) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture39.Picture = LoadPicture(z(80))
    If z(39) = z(80) Then
        v(39) = v(39) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture39.Picture = LoadPicture(z(81))
    If z(39) = z(81) Then
        v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture39.Picture = LoadPicture(z(82))
    If z(39) = z(82) Then
        v(39) = v(39) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture39.Picture = LoadPicture(z(83))
    If z(39) = z(83) Then
        v(39) = v(39) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture39.Picture = LoadPicture(z(84))
    If z(39) = z(84) Then
        v(39) = v(39) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture39.Picture = LoadPicture(z(85))
    If z(39) = z(85) Then
        v(39) = v(39) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture39.Picture = LoadPicture(z(86))
    If z(39) = z(86) Then
        v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture39.Picture = LoadPicture(z(87))
    If z(39) = z(87) Then
        v(39) = v(39) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture39.Picture = LoadPicture(z(88))
    If z(39) = z(88) Then
        v(39) = v(39) + 1
        End If
      
End If

End Sub

Private Sub Picture4_DblClick()
w = App.Path & "\clear.jpg"
Picture4.Picture = LoadPicture(w)
End Sub

Private Sub Picture4_DragDrop(Source As Control, x As Single, y As Single)
v(4) = 0
If Source = Picture45 Then
    Picture4.Picture = LoadPicture(z(45))
    If z(4) = z(45) Then
        v(4) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture4.Picture = LoadPicture(z(46))
    If z(4) = z(46) Then
        v(4) = v(4) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture4.Picture = LoadPicture(z(47))
    If z(4) = z(47) Then
       v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture4.Picture = LoadPicture(z(48))
    If z(4) = z(48) Then
        v(4) = v(4) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture4.Picture = LoadPicture(z(49))
    If z(4) = z(49) Then
      v(4) = v(4) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture4.Picture = LoadPicture(z(50))
    If z(4) = z(50) Then
      v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture4.Picture = LoadPicture(z(51))
    If z(4) = z(51) Then
        v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture4.Picture = LoadPicture(z(52))
    If z(4) = z(52) Then
        v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture4.Picture = LoadPicture(z(53))
    If z(4) = z(53) Then
      v(4) = v(4) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture4.Picture = LoadPicture(z(54))
    If z(4) = z(54) Then
     v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture4.Picture = LoadPicture(z(55))
    If z(4) = z(55) Then
        v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture4.Picture = LoadPicture(z(56))
    If z(4) = z(56) Then
        v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture4.Picture = LoadPicture(z(57))
    If z(4) = z(57) Then
       v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture4.Picture = LoadPicture(z(58))
    If z(4) = z(58) Then
       v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture4.Picture = LoadPicture(z(59))
    If z(4) = z(59) Then
       v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture4.Picture = LoadPicture(z(60))
    If z(4) = z(60) Then
        v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture4.Picture = LoadPicture(z(61))
    If z(4) = z(61) Then
       v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture4.Picture = LoadPicture(z(62))
    If z(4) = z(62) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture4.Picture = LoadPicture(z(63))
    If z(4) = z(63) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture4.Picture = LoadPicture(z(64))
    If z(4) = z(64) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture4.Picture = LoadPicture(z(65))
    If z(4) = z(65) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture4.Picture = LoadPicture(z(66))
    If z(4) = z(66) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture4.Picture = LoadPicture(z(67))
    If z(4) = z(67) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture4.Picture = LoadPicture(z(68))
    If z(4) = z(68) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture4.Picture = LoadPicture(z(69))
    If z(4) = z(69) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture4.Picture = LoadPicture(z(70))
    If z(4) = z(70) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture4.Picture = LoadPicture(z(71))
    If z(4) = z(71) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture4.Picture = LoadPicture(z(72))
    If z(4) = z(72) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture4.Picture = LoadPicture(z(73))
    If z(4) = z(73) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture4.Picture = LoadPicture(z(74))
    If z(4) = z(74) Then
        v(4) = v(4) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture4.Picture = LoadPicture(z(75))
    If z(4) = z(75) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture4.Picture = LoadPicture(z(76))
    If z(4) = z(76) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture4.Picture = LoadPicture(z(77))
    If z(4) = z(77) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture4.Picture = LoadPicture(z(78))
    If z(4) = z(78) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture4.Picture = LoadPicture(z(79))
    If z(4) = z(79) Then
        v(4) = v(4) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture4.Picture = LoadPicture(z(80))
    If z(4) = z(80) Then
        v(4) = v(4) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture4.Picture = LoadPicture(z(81))
    If z(4) = z(81) Then
        v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture4.Picture = LoadPicture(z(82))
    If z(4) = z(82) Then
        v(4) = v(4) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture4.Picture = LoadPicture(z(83))
    If z(4) = z(83) Then
        v(4) = v(4) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture4.Picture = LoadPicture(z(84))
    If z(4) = z(84) Then
        v(4) = v(4) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture4.Picture = LoadPicture(z(85))
    If z(4) = z(85) Then
        v(4) = v(4) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture4.Picture = LoadPicture(z(86))
    If z(4) = z(86) Then
        v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture4.Picture = LoadPicture(z(87))
    If z(4) = z(87) Then
        v(4) = v(4) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture4.Picture = LoadPicture(z(88))
    If z(4) = z(88) Then
        v(4) = v(4) + 1
        End If
      
End If
End Sub

Private Sub Picture40_DblClick()
w = App.Path & "\clear.jpg"
Picture40.Picture = LoadPicture(w)
End Sub

Private Sub Picture40_DragDrop(Source As Control, x As Single, y As Single)
v(40) = 0
If Source = Picture45 Then
    Picture40.Picture = LoadPicture(z(45))
    If z(40) = z(45) Then
        v(40) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture40.Picture = LoadPicture(z(46))
    If z(40) = z(46) Then
        v(40) = v(40) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture40.Picture = LoadPicture(z(47))
    If z(40) = z(47) Then
       v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture40.Picture = LoadPicture(z(48))
    If z(40) = z(48) Then
        v(40) = v(40) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture40.Picture = LoadPicture(z(49))
    If z(40) = z(49) Then
      v(40) = v(40) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture40.Picture = LoadPicture(z(50))
    If z(40) = z(50) Then
      v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture40.Picture = LoadPicture(z(51))
    If z(40) = z(51) Then
        v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture40.Picture = LoadPicture(z(52))
    If z(40) = z(52) Then
        v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture40.Picture = LoadPicture(z(53))
    If z(40) = z(53) Then
      v(40) = v(40) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture40.Picture = LoadPicture(z(54))
    If z(40) = z(54) Then
     v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture40.Picture = LoadPicture(z(55))
    If z(40) = z(55) Then
        v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture40.Picture = LoadPicture(z(56))
    If z(40) = z(56) Then
        v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture40.Picture = LoadPicture(z(57))
    If z(40) = z(57) Then
       v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture40.Picture = LoadPicture(z(58))
    If z(40) = z(58) Then
       v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture40.Picture = LoadPicture(z(59))
    If z(40) = z(59) Then
       v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture40.Picture = LoadPicture(z(60))
    If z(40) = z(60) Then
        v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture40.Picture = LoadPicture(z(61))
    If z(40) = z(61) Then
       v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture40.Picture = LoadPicture(z(62))
    If z(40) = z(62) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture40.Picture = LoadPicture(z(63))
    If z(40) = z(63) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture40.Picture = LoadPicture(z(64))
    If z(40) = z(64) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture40.Picture = LoadPicture(z(65))
    If z(40) = z(65) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture40.Picture = LoadPicture(z(66))
    If z(40) = z(66) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture40.Picture = LoadPicture(z(67))
    If z(40) = z(67) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture40.Picture = LoadPicture(z(68))
    If z(40) = z(68) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture40.Picture = LoadPicture(z(69))
    If z(40) = z(69) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture40.Picture = LoadPicture(z(70))
    If z(40) = z(70) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture40.Picture = LoadPicture(z(71))
    If z(40) = z(71) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture40.Picture = LoadPicture(z(72))
    If z(40) = z(72) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture40.Picture = LoadPicture(z(73))
    If z(40) = z(73) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture40.Picture = LoadPicture(z(74))
    If z(40) = z(74) Then
        v(40) = v(40) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture40.Picture = LoadPicture(z(75))
    If z(40) = z(75) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture40.Picture = LoadPicture(z(76))
    If z(40) = z(76) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture40.Picture = LoadPicture(z(77))
    If z(40) = z(77) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture40.Picture = LoadPicture(z(78))
    If z(40) = z(78) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture40.Picture = LoadPicture(z(79))
    If z(40) = z(79) Then
        v(40) = v(40) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture40.Picture = LoadPicture(z(80))
    If z(40) = z(80) Then
        v(40) = v(40) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture40.Picture = LoadPicture(z(81))
    If z(40) = z(81) Then
        v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture40.Picture = LoadPicture(z(82))
    If z(40) = z(82) Then
        v(40) = v(40) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture40.Picture = LoadPicture(z(83))
    If z(40) = z(83) Then
        v(40) = v(40) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture40.Picture = LoadPicture(z(84))
    If z(40) = z(84) Then
        v(40) = v(40) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture40.Picture = LoadPicture(z(85))
    If z(40) = z(85) Then
        v(40) = v(40) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture40.Picture = LoadPicture(z(86))
    If z(40) = z(86) Then
        v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture40.Picture = LoadPicture(z(87))
    If z(40) = z(87) Then
        v(40) = v(40) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture40.Picture = LoadPicture(z(88))
    If z(40) = z(88) Then
        v(40) = v(40) + 1
        End If
      
End If

End Sub

Private Sub Picture41_DblClick()
w = App.Path & "\clear.jpg"
Picture41.Picture = LoadPicture(w)
End Sub

Private Sub Picture41_DragDrop(Source As Control, x As Single, y As Single)
v(41) = 0
If Source = Picture45 Then
    Picture41.Picture = LoadPicture(z(45))
    If z(41) = z(45) Then
        v(41) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture41.Picture = LoadPicture(z(46))
    If z(41) = z(46) Then
        v(41) = v(41) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture41.Picture = LoadPicture(z(47))
    If z(41) = z(47) Then
       v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture41.Picture = LoadPicture(z(48))
    If z(41) = z(48) Then
        v(41) = v(41) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture41.Picture = LoadPicture(z(49))
    If z(41) = z(49) Then
      v(41) = v(41) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture41.Picture = LoadPicture(z(50))
    If z(41) = z(50) Then
      v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture41.Picture = LoadPicture(z(51))
    If z(41) = z(51) Then
        v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture41.Picture = LoadPicture(z(52))
    If z(41) = z(52) Then
        v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture41.Picture = LoadPicture(z(53))
    If z(41) = z(53) Then
      v(41) = v(41) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture41.Picture = LoadPicture(z(54))
    If z(41) = z(54) Then
     v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture41.Picture = LoadPicture(z(55))
    If z(41) = z(55) Then
        v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture41.Picture = LoadPicture(z(56))
    If z(41) = z(56) Then
        v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture41.Picture = LoadPicture(z(57))
    If z(41) = z(57) Then
       v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture41.Picture = LoadPicture(z(58))
    If z(41) = z(58) Then
       v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture41.Picture = LoadPicture(z(59))
    If z(41) = z(59) Then
       v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture41.Picture = LoadPicture(z(60))
    If z(41) = z(60) Then
        v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture41.Picture = LoadPicture(z(61))
    If z(41) = z(61) Then
       v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture41.Picture = LoadPicture(z(62))
    If z(41) = z(62) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture41.Picture = LoadPicture(z(63))
    If z(41) = z(63) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture41.Picture = LoadPicture(z(64))
    If z(41) = z(64) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture41.Picture = LoadPicture(z(65))
    If z(41) = z(65) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture41.Picture = LoadPicture(z(66))
    If z(41) = z(66) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture41.Picture = LoadPicture(z(67))
    If z(41) = z(67) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture41.Picture = LoadPicture(z(68))
    If z(41) = z(68) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture41.Picture = LoadPicture(z(69))
    If z(41) = z(69) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture41.Picture = LoadPicture(z(70))
    If z(41) = z(70) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture41.Picture = LoadPicture(z(71))
    If z(41) = z(71) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture41.Picture = LoadPicture(z(72))
    If z(41) = z(72) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture41.Picture = LoadPicture(z(73))
    If z(41) = z(73) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture41.Picture = LoadPicture(z(74))
    If z(41) = z(74) Then
        v(41) = v(41) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture41.Picture = LoadPicture(z(75))
    If z(41) = z(75) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture41.Picture = LoadPicture(z(76))
    If z(41) = z(76) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture41.Picture = LoadPicture(z(77))
    If z(41) = z(77) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture41.Picture = LoadPicture(z(78))
    If z(41) = z(78) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture41.Picture = LoadPicture(z(79))
    If z(41) = z(79) Then
        v(41) = v(41) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture41.Picture = LoadPicture(z(80))
    If z(41) = z(80) Then
        v(41) = v(41) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture41.Picture = LoadPicture(z(81))
    If z(41) = z(81) Then
        v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture41.Picture = LoadPicture(z(82))
    If z(41) = z(82) Then
        v(41) = v(41) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture41.Picture = LoadPicture(z(83))
    If z(41) = z(83) Then
        v(41) = v(41) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture41.Picture = LoadPicture(z(84))
    If z(41) = z(84) Then
        v(41) = v(41) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture41.Picture = LoadPicture(z(85))
    If z(41) = z(85) Then
        v(41) = v(41) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture41.Picture = LoadPicture(z(86))
    If z(41) = z(86) Then
        v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture41.Picture = LoadPicture(z(87))
    If z(41) = z(87) Then
        v(41) = v(41) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture41.Picture = LoadPicture(z(88))
    If z(41) = z(88) Then
        v(41) = v(41) + 1
        End If
      
End If

End Sub

Private Sub Picture42_DblClick()
w = App.Path & "\clear.jpg"
Picture42.Picture = LoadPicture(w)
End Sub

Private Sub Picture42_DragDrop(Source As Control, x As Single, y As Single)
v(42) = 0
If Source = Picture45 Then
    Picture42.Picture = LoadPicture(z(45))
    If z(42) = z(45) Then
        v(42) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture42.Picture = LoadPicture(z(46))
    If z(42) = z(46) Then
        v(42) = v(42) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture42.Picture = LoadPicture(z(47))
    If z(42) = z(47) Then
       v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture42.Picture = LoadPicture(z(48))
    If z(42) = z(48) Then
        v(42) = v(42) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture42.Picture = LoadPicture(z(49))
    If z(42) = z(49) Then
      v(42) = v(42) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture42.Picture = LoadPicture(z(50))
    If z(42) = z(50) Then
      v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture42.Picture = LoadPicture(z(51))
    If z(42) = z(51) Then
        v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture42.Picture = LoadPicture(z(52))
    If z(42) = z(52) Then
        v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture42.Picture = LoadPicture(z(53))
    If z(42) = z(53) Then
      v(42) = v(42) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture42.Picture = LoadPicture(z(54))
    If z(42) = z(54) Then
     v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture42.Picture = LoadPicture(z(55))
    If z(42) = z(55) Then
        v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture42.Picture = LoadPicture(z(56))
    If z(42) = z(56) Then
        v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture42.Picture = LoadPicture(z(57))
    If z(42) = z(57) Then
       v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture42.Picture = LoadPicture(z(58))
    If z(42) = z(58) Then
       v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture42.Picture = LoadPicture(z(59))
    If z(42) = z(59) Then
       v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture42.Picture = LoadPicture(z(60))
    If z(42) = z(60) Then
        v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture42.Picture = LoadPicture(z(61))
    If z(42) = z(61) Then
       v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture42.Picture = LoadPicture(z(62))
    If z(42) = z(62) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture42.Picture = LoadPicture(z(63))
    If z(42) = z(63) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture42.Picture = LoadPicture(z(64))
    If z(42) = z(64) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture42.Picture = LoadPicture(z(65))
    If z(42) = z(65) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture42.Picture = LoadPicture(z(66))
    If z(42) = z(66) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture42.Picture = LoadPicture(z(67))
    If z(42) = z(67) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture42.Picture = LoadPicture(z(68))
    If z(42) = z(68) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture42.Picture = LoadPicture(z(69))
    If z(42) = z(69) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture42.Picture = LoadPicture(z(70))
    If z(42) = z(70) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture42.Picture = LoadPicture(z(71))
    If z(42) = z(71) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture42.Picture = LoadPicture(z(72))
    If z(42) = z(72) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture42.Picture = LoadPicture(z(73))
    If z(42) = z(73) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture42.Picture = LoadPicture(z(74))
    If z(42) = z(74) Then
        v(42) = v(42) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture42.Picture = LoadPicture(z(75))
    If z(42) = z(75) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture42.Picture = LoadPicture(z(76))
    If z(42) = z(76) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture42.Picture = LoadPicture(z(77))
    If z(42) = z(77) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture42.Picture = LoadPicture(z(78))
    If z(42) = z(78) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture42.Picture = LoadPicture(z(79))
    If z(42) = z(79) Then
        v(42) = v(42) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture42.Picture = LoadPicture(z(80))
    If z(42) = z(80) Then
        v(42) = v(42) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture42.Picture = LoadPicture(z(81))
    If z(42) = z(81) Then
        v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture42.Picture = LoadPicture(z(82))
    If z(42) = z(82) Then
        v(42) = v(42) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture42.Picture = LoadPicture(z(83))
    If z(42) = z(83) Then
        v(42) = v(42) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture42.Picture = LoadPicture(z(84))
    If z(42) = z(84) Then
        v(42) = v(42) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture42.Picture = LoadPicture(z(85))
    If z(42) = z(85) Then
        v(42) = v(42) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture42.Picture = LoadPicture(z(86))
    If z(42) = z(86) Then
        v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture42.Picture = LoadPicture(z(87))
    If z(42) = z(87) Then
        v(42) = v(42) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture42.Picture = LoadPicture(z(88))
    If z(42) = z(88) Then
        v(42) = v(42) + 1
        End If
      
End If

End Sub

Private Sub Picture43_DblClick()
w = App.Path & "\clear.jpg"
Picture43.Picture = LoadPicture(w)
End Sub

Private Sub Picture43_DragDrop(Source As Control, x As Single, y As Single)
v(43) = 0
If Source = Picture45 Then
     Picture43.Picture = LoadPicture(z(45))
    If z(43) = z(45) Then
        v(43) = 1
        End If
    ElseIf Source = Picture46 Then
     Picture43.Picture = LoadPicture(z(46))
    If z(43) = z(46) Then
        v(43) = v(43) + 1
        End If
    ElseIf Source = Picture47 Then
     Picture43.Picture = LoadPicture(z(47))
    If z(43) = z(47) Then
       v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture48 Then
     Picture43.Picture = LoadPicture(z(48))
    If z(43) = z(48) Then
        v(43) = v(43) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
     Picture43.Picture = LoadPicture(z(49))
    If z(43) = z(49) Then
      v(43) = v(43) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
     Picture43.Picture = LoadPicture(z(50))
    If z(43) = z(50) Then
      v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture51 Then
     Picture43.Picture = LoadPicture(z(51))
    If z(43) = z(51) Then
        v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture52 Then
     Picture43.Picture = LoadPicture(z(52))
    If z(43) = z(52) Then
        v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture53 Then
     Picture43.Picture = LoadPicture(z(53))
    If z(43) = z(53) Then
      v(43) = v(43) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
     Picture43.Picture = LoadPicture(z(54))
    If z(43) = z(54) Then
     v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture55 Then
     Picture43.Picture = LoadPicture(z(55))
    If z(43) = z(55) Then
        v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture56 Then
     Picture43.Picture = LoadPicture(z(56))
    If z(43) = z(56) Then
        v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture57 Then
     Picture43.Picture = LoadPicture(z(57))
    If z(43) = z(57) Then
       v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture58 Then
     Picture43.Picture = LoadPicture(z(58))
    If z(43) = z(58) Then
       v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture59 Then
     Picture43.Picture = LoadPicture(z(59))
    If z(43) = z(59) Then
       v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture60 Then
     Picture43.Picture = LoadPicture(z(60))
    If z(43) = z(60) Then
        v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture61 Then
     Picture43.Picture = LoadPicture(z(61))
    If z(43) = z(61) Then
       v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture62 Then
     Picture43.Picture = LoadPicture(z(62))
    If z(43) = z(62) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture63 Then
     Picture43.Picture = LoadPicture(z(63))
    If z(43) = z(63) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture64 Then
     Picture43.Picture = LoadPicture(z(64))
    If z(43) = z(64) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture65 Then
     Picture43.Picture = LoadPicture(z(65))
    If z(43) = z(65) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture66 Then
     Picture43.Picture = LoadPicture(z(66))
    If z(43) = z(66) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
     Picture43.Picture = LoadPicture(z(67))
    If z(43) = z(67) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture68 Then
     Picture43.Picture = LoadPicture(z(68))
    If z(43) = z(68) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture69 Then
     Picture43.Picture = LoadPicture(z(69))
    If z(43) = z(69) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture70 Then
     Picture43.Picture = LoadPicture(z(70))
    If z(43) = z(70) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
     Picture43.Picture = LoadPicture(z(71))
    If z(43) = z(71) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
     Picture43.Picture = LoadPicture(z(72))
    If z(43) = z(72) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
     Picture43.Picture = LoadPicture(z(73))
    If z(43) = z(73) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture74 Then
     Picture43.Picture = LoadPicture(z(74))
    If z(43) = z(74) Then
        v(43) = v(43) + 1
        End If
        
          ElseIf Source = Picture75 Then
     Picture43.Picture = LoadPicture(z(75))
    If z(43) = z(75) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
     Picture43.Picture = LoadPicture(z(76))
    If z(43) = z(76) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
     Picture43.Picture = LoadPicture(z(77))
    If z(43) = z(77) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
     Picture43.Picture = LoadPicture(z(78))
    If z(43) = z(78) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
     Picture43.Picture = LoadPicture(z(79))
    If z(43) = z(79) Then
        v(43) = v(43) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
     Picture43.Picture = LoadPicture(z(80))
    If z(43) = z(80) Then
        v(43) = v(43) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
     Picture43.Picture = LoadPicture(z(81))
    If z(43) = z(81) Then
        v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture82 Then
     Picture43.Picture = LoadPicture(z(82))
    If z(43) = z(82) Then
        v(43) = v(43) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
     Picture43.Picture = LoadPicture(z(83))
    If z(43) = z(83) Then
        v(43) = v(43) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
     Picture43.Picture = LoadPicture(z(84))
    If z(43) = z(84) Then
        v(43) = v(43) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
     Picture43.Picture = LoadPicture(z(85))
    If z(43) = z(85) Then
        v(43) = v(43) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
     Picture43.Picture = LoadPicture(z(86))
    If z(43) = z(86) Then
        v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture87 Then
     Picture43.Picture = LoadPicture(z(87))
    If z(43) = z(87) Then
        v(43) = v(43) + 1
        End If
        
        ElseIf Source = Picture88 Then
     Picture43.Picture = LoadPicture(z(88))
    If z(43) = z(88) Then
        v(43) = v(43) + 1
        End If
      
End If

End Sub

Private Sub Picture44_DblClick()
w = App.Path & "\clear.jpg"
Picture44.Picture = LoadPicture(w)
End Sub

Private Sub Picture44_DragDrop(Source As Control, x As Single, y As Single)
v(44) = 0
If Source = Picture45 Then
      Picture44.Picture = LoadPicture(z(45))
    If z(44) = z(45) Then
        v(44) = 1
        End If
    ElseIf Source = Picture46 Then
      Picture44.Picture = LoadPicture(z(46))
    If z(44) = z(46) Then
        v(44) = v(44) + 1
        End If
    ElseIf Source = Picture47 Then
      Picture44.Picture = LoadPicture(z(47))
    If z(44) = z(47) Then
       v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture48 Then
      Picture44.Picture = LoadPicture(z(48))
    If z(44) = z(48) Then
        v(44) = v(44) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
      Picture44.Picture = LoadPicture(z(49))
    If z(44) = z(49) Then
      v(44) = v(44) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
      Picture44.Picture = LoadPicture(z(50))
    If z(44) = z(50) Then
      v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture51 Then
      Picture44.Picture = LoadPicture(z(51))
    If z(44) = z(51) Then
        v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture52 Then
      Picture44.Picture = LoadPicture(z(52))
    If z(44) = z(52) Then
        v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture53 Then
      Picture44.Picture = LoadPicture(z(53))
    If z(44) = z(53) Then
      v(44) = v(44) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
      Picture44.Picture = LoadPicture(z(54))
    If z(44) = z(54) Then
     v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture55 Then
      Picture44.Picture = LoadPicture(z(55))
    If z(44) = z(55) Then
        v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture56 Then
      Picture44.Picture = LoadPicture(z(56))
    If z(44) = z(56) Then
        v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture57 Then
      Picture44.Picture = LoadPicture(z(57))
    If z(44) = z(57) Then
       v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture58 Then
      Picture44.Picture = LoadPicture(z(58))
    If z(44) = z(58) Then
       v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture59 Then
      Picture44.Picture = LoadPicture(z(59))
    If z(44) = z(59) Then
       v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture60 Then
      Picture44.Picture = LoadPicture(z(60))
    If z(44) = z(60) Then
        v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture61 Then
      Picture44.Picture = LoadPicture(z(61))
    If z(44) = z(61) Then
       v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture62 Then
      Picture44.Picture = LoadPicture(z(62))
    If z(44) = z(62) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture63 Then
      Picture44.Picture = LoadPicture(z(63))
    If z(44) = z(63) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture64 Then
      Picture44.Picture = LoadPicture(z(64))
    If z(44) = z(64) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture65 Then
      Picture44.Picture = LoadPicture(z(65))
    If z(44) = z(65) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture66 Then
      Picture44.Picture = LoadPicture(z(66))
    If z(44) = z(66) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
      Picture44.Picture = LoadPicture(z(67))
    If z(44) = z(67) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture68 Then
      Picture44.Picture = LoadPicture(z(68))
    If z(44) = z(68) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture69 Then
      Picture44.Picture = LoadPicture(z(69))
    If z(44) = z(69) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture70 Then
      Picture44.Picture = LoadPicture(z(70))
    If z(44) = z(70) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
      Picture44.Picture = LoadPicture(z(71))
    If z(44) = z(71) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
      Picture44.Picture = LoadPicture(z(72))
    If z(44) = z(72) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
      Picture44.Picture = LoadPicture(z(73))
    If z(44) = z(73) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture74 Then
      Picture44.Picture = LoadPicture(z(74))
    If z(44) = z(74) Then
        v(44) = v(44) + 1
        End If
        
          ElseIf Source = Picture75 Then
      Picture44.Picture = LoadPicture(z(75))
    If z(44) = z(75) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
      Picture44.Picture = LoadPicture(z(76))
    If z(44) = z(76) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
      Picture44.Picture = LoadPicture(z(77))
    If z(44) = z(77) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
      Picture44.Picture = LoadPicture(z(78))
    If z(44) = z(78) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
      Picture44.Picture = LoadPicture(z(79))
    If z(44) = z(79) Then
        v(44) = v(44) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
      Picture44.Picture = LoadPicture(z(80))
    If z(44) = z(80) Then
        v(44) = v(44) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
      Picture44.Picture = LoadPicture(z(81))
    If z(44) = z(81) Then
        v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture82 Then
      Picture44.Picture = LoadPicture(z(82))
    If z(44) = z(82) Then
        v(44) = v(44) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
      Picture44.Picture = LoadPicture(z(83))
    If z(44) = z(83) Then
        v(44) = v(44) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
      Picture44.Picture = LoadPicture(z(84))
    If z(44) = z(84) Then
        v(44) = v(44) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
      Picture44.Picture = LoadPicture(z(85))
    If z(44) = z(85) Then
        v(44) = v(44) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
      Picture44.Picture = LoadPicture(z(86))
    If z(44) = z(86) Then
        v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture87 Then
      Picture44.Picture = LoadPicture(z(87))
    If z(44) = z(87) Then
        v(44) = v(44) + 1
        End If
        
        ElseIf Source = Picture88 Then
      Picture44.Picture = LoadPicture(z(88))
    If z(44) = z(88) Then
        v(44) = v(44) + 1
        End If
      
End If

End Sub

Private Sub Picture5_DblClick()
w = App.Path & "\clear.jpg"
Picture5.Picture = LoadPicture(w)
End Sub

Private Sub Picture5_DragDrop(Source As Control, x As Single, y As Single)
v(5) = 0
If Source = Picture45 Then
    Picture5.Picture = LoadPicture(z(45))
    If z(5) = z(45) Then
        v(5) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture5.Picture = LoadPicture(z(46))
    If z(5) = z(46) Then
        v(5) = v(5) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture5.Picture = LoadPicture(z(47))
    If z(5) = z(47) Then
       v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture5.Picture = LoadPicture(z(48))
    If z(5) = z(48) Then
        v(5) = v(5) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture5.Picture = LoadPicture(z(49))
    If z(5) = z(49) Then
      v(5) = v(5) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture5.Picture = LoadPicture(z(50))
    If z(5) = z(50) Then
      v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture5.Picture = LoadPicture(z(51))
    If z(5) = z(51) Then
        v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture5.Picture = LoadPicture(z(52))
    If z(5) = z(52) Then
        v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture5.Picture = LoadPicture(z(53))
    If z(5) = z(53) Then
      v(5) = v(5) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture5.Picture = LoadPicture(z(54))
    If z(5) = z(54) Then
     v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture5.Picture = LoadPicture(z(55))
    If z(5) = z(55) Then
        v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture5.Picture = LoadPicture(z(56))
    If z(5) = z(56) Then
        v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture5.Picture = LoadPicture(z(57))
    If z(5) = z(57) Then
       v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture5.Picture = LoadPicture(z(58))
    If z(5) = z(58) Then
       v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture5.Picture = LoadPicture(z(59))
    If z(5) = z(59) Then
       v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture5.Picture = LoadPicture(z(60))
    If z(5) = z(60) Then
        v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture5.Picture = LoadPicture(z(61))
    If z(5) = z(61) Then
       v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture5.Picture = LoadPicture(z(62))
    If z(5) = z(62) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture5.Picture = LoadPicture(z(63))
    If z(5) = z(63) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture5.Picture = LoadPicture(z(64))
    If z(5) = z(64) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture5.Picture = LoadPicture(z(65))
    If z(5) = z(65) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture5.Picture = LoadPicture(z(66))
    If z(5) = z(66) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture5.Picture = LoadPicture(z(67))
    If z(5) = z(67) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture5.Picture = LoadPicture(z(68))
    If z(5) = z(68) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture5.Picture = LoadPicture(z(69))
    If z(5) = z(69) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture5.Picture = LoadPicture(z(70))
    If z(5) = z(70) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture5.Picture = LoadPicture(z(71))
    If z(5) = z(71) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture5.Picture = LoadPicture(z(72))
    If z(5) = z(72) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture5.Picture = LoadPicture(z(73))
    If z(5) = z(73) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture5.Picture = LoadPicture(z(74))
    If z(5) = z(74) Then
        v(5) = v(5) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture5.Picture = LoadPicture(z(75))
    If z(5) = z(75) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture5.Picture = LoadPicture(z(76))
    If z(5) = z(76) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture5.Picture = LoadPicture(z(77))
    If z(5) = z(77) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture5.Picture = LoadPicture(z(78))
    If z(5) = z(78) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture5.Picture = LoadPicture(z(79))
    If z(5) = z(79) Then
        v(5) = v(5) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture5.Picture = LoadPicture(z(80))
    If z(5) = z(80) Then
        v(5) = v(5) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture5.Picture = LoadPicture(z(81))
    If z(5) = z(81) Then
        v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture5.Picture = LoadPicture(z(82))
    If z(5) = z(82) Then
        v(5) = v(5) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture5.Picture = LoadPicture(z(83))
    If z(5) = z(83) Then
        v(5) = v(5) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture5.Picture = LoadPicture(z(84))
    If z(5) = z(84) Then
        v(5) = v(5) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture5.Picture = LoadPicture(z(85))
    If z(5) = z(85) Then
        v(5) = v(5) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture5.Picture = LoadPicture(z(86))
    If z(5) = z(86) Then
        v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture5.Picture = LoadPicture(z(87))
    If z(5) = z(87) Then
        v(5) = v(5) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture5.Picture = LoadPicture(z(88))
    If z(5) = z(88) Then
        v(5) = v(5) + 1
        End If
      
End If
          
       
End Sub

Private Sub Picture6_DblClick()
w = App.Path & "\clear.jpg"
Picture6.Picture = LoadPicture(w)
End Sub

Private Sub Picture6_DragDrop(Source As Control, x As Single, y As Single)
v(6) = 0
If Source = Picture45 Then
    Picture6.Picture = LoadPicture(z(45))
    If z(6) = z(45) Then
        v(6) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture6.Picture = LoadPicture(z(46))
    If z(6) = z(46) Then
        v(6) = v(6) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture6.Picture = LoadPicture(z(47))
    If z(6) = z(47) Then
       v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture6.Picture = LoadPicture(z(48))
    If z(6) = z(48) Then
        v(6) = v(6) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture6.Picture = LoadPicture(z(49))
    If z(6) = z(49) Then
      v(6) = v(6) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture6.Picture = LoadPicture(z(50))
    If z(6) = z(50) Then
      v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture6.Picture = LoadPicture(z(51))
    If z(6) = z(51) Then
        v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture6.Picture = LoadPicture(z(52))
    If z(6) = z(52) Then
        v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture6.Picture = LoadPicture(z(53))
    If z(6) = z(53) Then
      v(6) = v(6) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture6.Picture = LoadPicture(z(54))
    If z(6) = z(54) Then
     v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture6.Picture = LoadPicture(z(55))
    If z(6) = z(55) Then
        v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture6.Picture = LoadPicture(z(56))
    If z(6) = z(56) Then
        v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture6.Picture = LoadPicture(z(57))
    If z(6) = z(57) Then
       v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture6.Picture = LoadPicture(z(58))
    If z(6) = z(58) Then
       v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture6.Picture = LoadPicture(z(59))
    If z(6) = z(59) Then
       v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture6.Picture = LoadPicture(z(60))
    If z(6) = z(60) Then
        v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture6.Picture = LoadPicture(z(61))
    If z(6) = z(61) Then
       v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture6.Picture = LoadPicture(z(62))
    If z(6) = z(62) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture6.Picture = LoadPicture(z(63))
    If z(6) = z(63) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture6.Picture = LoadPicture(z(64))
    If z(6) = z(64) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture6.Picture = LoadPicture(z(65))
    If z(6) = z(65) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture6.Picture = LoadPicture(z(66))
    If z(6) = z(66) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture6.Picture = LoadPicture(z(67))
    If z(6) = z(67) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture6.Picture = LoadPicture(z(68))
    If z(6) = z(68) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture6.Picture = LoadPicture(z(69))
    If z(6) = z(69) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture6.Picture = LoadPicture(z(70))
    If z(6) = z(70) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture6.Picture = LoadPicture(z(71))
    If z(6) = z(71) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture6.Picture = LoadPicture(z(72))
    If z(6) = z(72) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture6.Picture = LoadPicture(z(73))
    If z(6) = z(73) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture6.Picture = LoadPicture(z(74))
    If z(6) = z(74) Then
        v(6) = v(6) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture6.Picture = LoadPicture(z(75))
    If z(6) = z(75) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture6.Picture = LoadPicture(z(76))
    If z(6) = z(76) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture6.Picture = LoadPicture(z(77))
    If z(6) = z(77) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture6.Picture = LoadPicture(z(78))
    If z(6) = z(78) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture6.Picture = LoadPicture(z(79))
    If z(6) = z(79) Then
        v(6) = v(6) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture6.Picture = LoadPicture(z(80))
    If z(6) = z(80) Then
        v(6) = v(6) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture6.Picture = LoadPicture(z(81))
    If z(6) = z(81) Then
        v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture6.Picture = LoadPicture(z(82))
    If z(6) = z(82) Then
        v(6) = v(6) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture6.Picture = LoadPicture(z(83))
    If z(6) = z(83) Then
        v(6) = v(6) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture6.Picture = LoadPicture(z(84))
    If z(6) = z(84) Then
        v(6) = v(6) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture6.Picture = LoadPicture(z(85))
    If z(6) = z(85) Then
        v(6) = v(6) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture6.Picture = LoadPicture(z(86))
    If z(6) = z(86) Then
        v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture6.Picture = LoadPicture(z(87))
    If z(6) = z(87) Then
        v(6) = v(6) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture6.Picture = LoadPicture(z(88))
    If z(6) = z(88) Then
        v(6) = v(6) + 1
        End If
      
End If

      

End Sub

Private Sub Picture7_DblClick()
w = App.Path & "\clear.jpg"
Picture7.Picture = LoadPicture(w)
End Sub

Private Sub Picture7_DragDrop(Source As Control, x As Single, y As Single)
v(7) = 0
If Source = Picture45 Then
    Picture7.Picture = LoadPicture(z(45))
    If z(7) = z(45) Then
        v(7) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture7.Picture = LoadPicture(z(46))
    If z(7) = z(46) Then
        v(7) = v(7) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture7.Picture = LoadPicture(z(47))
    If z(7) = z(47) Then
       v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture7.Picture = LoadPicture(z(48))
    If z(7) = z(48) Then
        v(7) = v(7) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture7.Picture = LoadPicture(z(49))
    If z(7) = z(49) Then
      v(7) = v(7) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture7.Picture = LoadPicture(z(50))
    If z(7) = z(50) Then
      v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture7.Picture = LoadPicture(z(51))
    If z(7) = z(51) Then
        v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture7.Picture = LoadPicture(z(52))
    If z(7) = z(52) Then
        v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture7.Picture = LoadPicture(z(53))
    If z(7) = z(53) Then
      v(7) = v(7) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture7.Picture = LoadPicture(z(54))
    If z(7) = z(54) Then
     v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture7.Picture = LoadPicture(z(55))
    If z(7) = z(55) Then
        v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture7.Picture = LoadPicture(z(56))
    If z(7) = z(56) Then
        v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture7.Picture = LoadPicture(z(57))
    If z(7) = z(57) Then
       v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture7.Picture = LoadPicture(z(58))
    If z(7) = z(58) Then
       v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture7.Picture = LoadPicture(z(59))
    If z(7) = z(59) Then
       v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture7.Picture = LoadPicture(z(60))
    If z(7) = z(60) Then
        v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture7.Picture = LoadPicture(z(61))
    If z(7) = z(61) Then
       v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture7.Picture = LoadPicture(z(62))
    If z(7) = z(62) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture7.Picture = LoadPicture(z(63))
    If z(7) = z(63) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture7.Picture = LoadPicture(z(64))
    If z(7) = z(64) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture7.Picture = LoadPicture(z(65))
    If z(7) = z(65) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture7.Picture = LoadPicture(z(66))
    If z(7) = z(66) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture7.Picture = LoadPicture(z(67))
    If z(7) = z(67) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture7.Picture = LoadPicture(z(68))
    If z(7) = z(68) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture7.Picture = LoadPicture(z(69))
    If z(7) = z(69) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture7.Picture = LoadPicture(z(70))
    If z(7) = z(70) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture7.Picture = LoadPicture(z(71))
    If z(7) = z(71) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture7.Picture = LoadPicture(z(72))
    If z(7) = z(72) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture7.Picture = LoadPicture(z(73))
    If z(7) = z(73) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture7.Picture = LoadPicture(z(74))
    If z(7) = z(74) Then
        v(7) = v(7) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture7.Picture = LoadPicture(z(75))
    If z(7) = z(75) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture7.Picture = LoadPicture(z(76))
    If z(7) = z(76) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture7.Picture = LoadPicture(z(77))
    If z(7) = z(77) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture7.Picture = LoadPicture(z(78))
    If z(7) = z(78) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture7.Picture = LoadPicture(z(79))
    If z(7) = z(79) Then
        v(7) = v(7) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture7.Picture = LoadPicture(z(80))
    If z(7) = z(80) Then
        v(7) = v(7) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture7.Picture = LoadPicture(z(81))
    If z(7) = z(81) Then
        v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture7.Picture = LoadPicture(z(82))
    If z(7) = z(82) Then
        v(7) = v(7) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture7.Picture = LoadPicture(z(83))
    If z(7) = z(83) Then
        v(7) = v(7) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture7.Picture = LoadPicture(z(84))
    If z(7) = z(84) Then
        v(7) = v(7) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture7.Picture = LoadPicture(z(85))
    If z(7) = z(85) Then
        v(7) = v(7) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture7.Picture = LoadPicture(z(86))
    If z(7) = z(86) Then
        v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture7.Picture = LoadPicture(z(87))
    If z(7) = z(87) Then
        v(7) = v(7) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture7.Picture = LoadPicture(z(88))
    If z(7) = z(88) Then
        v(7) = v(7) + 1
        End If
      
End If
End Sub

Private Sub Picture8_DblClick()
w = App.Path & "\clear.jpg"
Picture8.Picture = LoadPicture(w)
End Sub

Private Sub Picture8_DragDrop(Source As Control, x As Single, y As Single)
v(8) = 0
If Source = Picture45 Then
    Picture8.Picture = LoadPicture(z(45))
    If z(8) = z(45) Then
        v(8) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture8.Picture = LoadPicture(z(46))
    If z(8) = z(46) Then
        v(8) = v(8) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture8.Picture = LoadPicture(z(47))
    If z(8) = z(47) Then
       v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture8.Picture = LoadPicture(z(48))
    If z(8) = z(48) Then
        v(8) = v(8) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture8.Picture = LoadPicture(z(49))
    If z(8) = z(49) Then
      v(8) = v(8) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture8.Picture = LoadPicture(z(50))
    If z(8) = z(50) Then
      v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture8.Picture = LoadPicture(z(51))
    If z(8) = z(51) Then
        v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture8.Picture = LoadPicture(z(52))
    If z(8) = z(52) Then
        v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture8.Picture = LoadPicture(z(53))
    If z(8) = z(53) Then
      v(8) = v(8) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture8.Picture = LoadPicture(z(54))
    If z(8) = z(54) Then
     v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture8.Picture = LoadPicture(z(55))
    If z(8) = z(55) Then
        v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture8.Picture = LoadPicture(z(56))
    If z(8) = z(56) Then
        v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture8.Picture = LoadPicture(z(57))
    If z(8) = z(57) Then
       v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture8.Picture = LoadPicture(z(58))
    If z(8) = z(58) Then
       v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture8.Picture = LoadPicture(z(59))
    If z(8) = z(59) Then
       v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture8.Picture = LoadPicture(z(60))
    If z(8) = z(60) Then
        v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture8.Picture = LoadPicture(z(61))
    If z(8) = z(61) Then
       v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture8.Picture = LoadPicture(z(62))
    If z(8) = z(62) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture8.Picture = LoadPicture(z(63))
    If z(8) = z(63) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture8.Picture = LoadPicture(z(64))
    If z(8) = z(64) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture8.Picture = LoadPicture(z(65))
    If z(8) = z(65) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture8.Picture = LoadPicture(z(66))
    If z(8) = z(66) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture8.Picture = LoadPicture(z(67))
    If z(8) = z(67) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture8.Picture = LoadPicture(z(68))
    If z(8) = z(68) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture8.Picture = LoadPicture(z(69))
    If z(8) = z(69) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture8.Picture = LoadPicture(z(70))
    If z(8) = z(70) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture8.Picture = LoadPicture(z(71))
    If z(8) = z(71) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture8.Picture = LoadPicture(z(72))
    If z(8) = z(72) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture8.Picture = LoadPicture(z(73))
    If z(8) = z(73) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture8.Picture = LoadPicture(z(74))
    If z(8) = z(74) Then
        v(8) = v(8) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture8.Picture = LoadPicture(z(75))
    If z(8) = z(75) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture8.Picture = LoadPicture(z(76))
    If z(8) = z(76) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture8.Picture = LoadPicture(z(77))
    If z(8) = z(77) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture8.Picture = LoadPicture(z(78))
    If z(8) = z(78) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture8.Picture = LoadPicture(z(79))
    If z(8) = z(79) Then
        v(8) = v(8) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture8.Picture = LoadPicture(z(80))
    If z(8) = z(80) Then
        v(8) = v(8) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture8.Picture = LoadPicture(z(81))
    If z(8) = z(81) Then
        v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture8.Picture = LoadPicture(z(82))
    If z(8) = z(82) Then
        v(8) = v(8) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture8.Picture = LoadPicture(z(83))
    If z(8) = z(83) Then
        v(8) = v(8) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture8.Picture = LoadPicture(z(84))
    If z(8) = z(84) Then
        v(8) = v(8) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture8.Picture = LoadPicture(z(85))
    If z(8) = z(85) Then
        v(8) = v(8) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture8.Picture = LoadPicture(z(86))
    If z(8) = z(86) Then
        v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture8.Picture = LoadPicture(z(87))
    If z(8) = z(87) Then
        v(8) = v(8) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture8.Picture = LoadPicture(z(88))
    If z(8) = z(88) Then
        v(8) = v(8) + 1
        End If
      
End If
End Sub

Private Sub Picture9_DblClick()
w = App.Path & "\clear.jpg"
Picture9.Picture = LoadPicture(w)
End Sub

Private Sub Picture9_DragDrop(Source As Control, x As Single, y As Single)
v(9) = 0
If Source = Picture45 Then
    Picture9.Picture = LoadPicture(z(45))
    If z(9) = z(45) Then
        v(9) = 1
        End If
    ElseIf Source = Picture46 Then
    Picture9.Picture = LoadPicture(z(46))
    If z(9) = z(46) Then
        v(9) = v(9) + 1
        End If
    ElseIf Source = Picture47 Then
    Picture9.Picture = LoadPicture(z(47))
    If z(9) = z(47) Then
       v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture48 Then
    Picture9.Picture = LoadPicture(z(48))
    If z(9) = z(48) Then
        v(9) = v(9) + 1
        End If
        
        
        ElseIf Source = Picture49 Then
    Picture9.Picture = LoadPicture(z(49))
    If z(9) = z(49) Then
      v(9) = v(9) + 1
        End If
        
        
        ElseIf Source = Picture50 Then
    Picture9.Picture = LoadPicture(z(50))
    If z(9) = z(50) Then
      v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture51 Then
    Picture9.Picture = LoadPicture(z(51))
    If z(9) = z(51) Then
        v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture52 Then
    Picture9.Picture = LoadPicture(z(52))
    If z(9) = z(52) Then
        v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture53 Then
    Picture9.Picture = LoadPicture(z(53))
    If z(9) = z(53) Then
      v(9) = v(9) + 1
        End If
        
        
        ElseIf Source = Picture54 Then
    Picture9.Picture = LoadPicture(z(54))
    If z(9) = z(54) Then
     v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture55 Then
    Picture9.Picture = LoadPicture(z(55))
    If z(9) = z(55) Then
        v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture56 Then
    Picture9.Picture = LoadPicture(z(56))
    If z(9) = z(56) Then
        v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture57 Then
    Picture9.Picture = LoadPicture(z(57))
    If z(9) = z(57) Then
       v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture58 Then
    Picture9.Picture = LoadPicture(z(58))
    If z(9) = z(58) Then
       v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture59 Then
    Picture9.Picture = LoadPicture(z(59))
    If z(9) = z(59) Then
       v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture60 Then
    Picture9.Picture = LoadPicture(z(60))
    If z(9) = z(60) Then
        v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture61 Then
    Picture9.Picture = LoadPicture(z(61))
    If z(9) = z(61) Then
       v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture62 Then
    Picture9.Picture = LoadPicture(z(62))
    If z(9) = z(62) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture63 Then
    Picture9.Picture = LoadPicture(z(63))
    If z(9) = z(63) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture64 Then
    Picture9.Picture = LoadPicture(z(64))
    If z(9) = z(64) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture65 Then
    Picture9.Picture = LoadPicture(z(65))
    If z(9) = z(65) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture66 Then
    Picture9.Picture = LoadPicture(z(66))
    If z(9) = z(66) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture67 Then
    Picture9.Picture = LoadPicture(z(67))
    If z(9) = z(67) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture68 Then
    Picture9.Picture = LoadPicture(z(68))
    If z(9) = z(68) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture69 Then
    Picture9.Picture = LoadPicture(z(69))
    If z(9) = z(69) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture70 Then
    Picture9.Picture = LoadPicture(z(70))
    If z(9) = z(70) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture71 Then
    Picture9.Picture = LoadPicture(z(71))
    If z(9) = z(71) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture72 Then
    Picture9.Picture = LoadPicture(z(72))
    If z(9) = z(72) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture73 Then
    Picture9.Picture = LoadPicture(z(73))
    If z(9) = z(73) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture74 Then
    Picture9.Picture = LoadPicture(z(74))
    If z(9) = z(74) Then
        v(9) = v(9) + 1
        End If
        
          ElseIf Source = Picture75 Then
    Picture9.Picture = LoadPicture(z(75))
    If z(9) = z(75) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture76 Then
    Picture9.Picture = LoadPicture(z(76))
    If z(9) = z(76) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture77 Then
    Picture9.Picture = LoadPicture(z(77))
    If z(9) = z(77) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture78 Then
    Picture9.Picture = LoadPicture(z(78))
    If z(9) = z(78) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture79 Then
    Picture9.Picture = LoadPicture(z(79))
    If z(9) = z(79) Then
        v(9) = v(9) + 1
        End If
        
        
          ElseIf Source = Picture80 Then
    Picture9.Picture = LoadPicture(z(80))
    If z(9) = z(80) Then
        v(9) = v(9) + 1
        End If
        
        
        ElseIf Source = Picture81 Then
    Picture9.Picture = LoadPicture(z(81))
    If z(9) = z(81) Then
        v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture82 Then
    Picture9.Picture = LoadPicture(z(82))
    If z(9) = z(82) Then
        v(9) = v(9) + 1
        End If
        
        
        ElseIf Source = Picture83 Then
    Picture9.Picture = LoadPicture(z(83))
    If z(9) = z(83) Then
        v(9) = v(9) + 1
        End If
        
        
        ElseIf Source = Picture84 Then
    Picture9.Picture = LoadPicture(z(84))
    If z(9) = z(84) Then
        v(9) = v(9) + 1
        End If
        
        
        ElseIf Source = Picture85 Then
    Picture9.Picture = LoadPicture(z(85))
    If z(9) = z(85) Then
        v(9) = v(9) + 1
        End If
        
        
        ElseIf Source = Picture86 Then
    Picture9.Picture = LoadPicture(z(86))
    If z(9) = z(86) Then
        v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture87 Then
    Picture9.Picture = LoadPicture(z(87))
    If z(9) = z(87) Then
        v(9) = v(9) + 1
        End If
        
        ElseIf Source = Picture88 Then
    Picture9.Picture = LoadPicture(z(88))
    If z(9) = z(88) Then
        v(9) = v(9) + 1
        End If
      
End If
End Sub
