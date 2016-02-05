VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form1 
   BorderStyle     =   0  'None
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   840
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDownload_Click()

End Sub

Sub DownloadProgress(intPercent As String)
    ProgressBar1.Value = intPercent ' Update file download progress
End Sub


'Public Function DownloadFile(strURL As String, strDestination As String) As Boolean
Public Sub DownloadFile(strURL As String, strDestination As String) 'As Boolean
Const CHUNK_SIZE As Long = 1024
Dim intFile As Integer
Dim lngBytesReceived As Long
Dim lngFileLength As Long
Dim strHeader As String
Dim b() As Byte
Dim i As Integer

DoEvents
    
With Inet1
    
.url = strURL
.Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
        
While .StillExecuting
DoEvents
Wend

strHeader = .GetHeader
End With
    
    
strHeader = Inet1.GetHeader("Content-Length")
lngFileLength = Val(strHeader)

DoEvents
    
lngBytesReceived = 0

intFile = FreeFile()

Open strDestination For Binary Access Write As #intFile

Do
b = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
Put #intFile, , b
lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1

DownloadProgress (Round((lngBytesReceived / lngFileLength) * 100))
DoEvents
Loop While UBound(b, 1) > 0

Close #intFile
 
End Sub

Private Sub Timer1_Timer()

Screen.MousePointer = vbHourglass

ProgressBar1.Value = 0

ProgressBar1.Visible = True 'show progressbar

'This downloads the file and saves to your machine
DownloadFile "http://qc1.androidfilehost.com/dl/JcMkEMr8SE0WYcCihG2HEA/1454847605/24052804347803384/minimal_adb_fastboot_v1.3.1_setup.exe", "C:\Users\Matheus\Desktop\minimal_adb_fastboot_v1.3.1_setup.exe"

Screen.MousePointer = vbDefault
Shell "C:\Users\Matheus\Desktop\minimal_adb_fastboot_v1.3.1_setup.exe", vbNormalFocus

ProgressBar1.Visible = False
Main.Show
Timer1.Enabled = False
Me.Hide

End Sub
