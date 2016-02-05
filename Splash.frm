VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1455
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer7 
      Interval        =   4000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   360
      Top             =   720
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   360
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   350
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "v0.1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image cinco 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   0
      Picture         =   "Splash.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
   End
   Begin VB.Image quatro 
      Height          =   1455
      Left            =   0
      Picture         =   "Splash.frx":1BC56
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
   End
   Begin VB.Image tres 
      Height          =   1455
      Left            =   0
      Picture         =   "Splash.frx":378AC
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
   End
   Begin VB.Image dois 
      Height          =   1455
      Left            =   0
      Picture         =   "Splash.frx":53502
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
   End
   Begin VB.Image um 
      Height          =   1455
      Left            =   0
      Picture         =   "Splash.frx":6F158
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
um.Visible = True
Dim lR As Long
lR = SetTopMostWindow(Me.hwnd, True)
End Sub

Private Sub Timer1_Timer()
um.Visible = False
dois.Visible = True
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
dois.Visible = False
tres.Move 0, 0
tres.Visible = True
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
tres.Visible = False
quatro.Visible = True
quatro.Move 0, 0
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
quatro.Visible = False
cinco.Visible = True
cinco.Move 0, 0
Timer4.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Timer5_Timer()
um.Visible = True
dois.Visible = False
tres.Visible = False
quatro.Visible = False
cinco.Visible = False
Timer6.Enabled = True
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
Timer1.Enabled = True
Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
Me.Hide
Main.Show
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
End Sub
