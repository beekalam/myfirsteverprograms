VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   3120
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6495
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Dialog.frx":030A
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4050
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OKButton_Click()
Unload Me
End Sub
