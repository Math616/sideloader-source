VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instalador de roms"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14670
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Install 
      BorderStyle     =   0  'None
      Caption         =   "Install"
      Height          =   4575
      Left            =   4920
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3840
         Width           =   4575
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "link."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "os drivers ADB através deste "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   4575
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Se não puder vê-lo, provavelmente você terá que instalar"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   4575
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "a instalação seguirá automaticamente."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   4575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "em alguns segundos, e se seu aparelho aparecer na lista"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   4575
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Você deverá ver uma janela de comando do Windows"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "A instalação foi iniciada"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Image Image3 
         Height          =   1200
         Left            =   1800
         Picture         =   "Principal.frx":1CFA
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Help 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   9840
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OK, compreendo"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   1200
         Left            =   1800
         Picture         =   "Principal.frx":683C
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "e não tive tempo de criar um tópico de ajuda"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   4575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "O programa ainda está em desenvolvimento,"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   4575
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Instalar"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Tamanho: "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Nenhum arquivo ZIP selecionado."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   1800
      Picture         =   "Principal.frx":B37E
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Como coloco meu smartphone em modo ""sideload""?"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Ele será flasheado automaticamente ao concluir a cópia"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Arraste para esta janela, o arquivo ZIP à ser instalado"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   4575
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Private Sub OpenUrl(ByVal url As String)
    r = ShellExecute(0, "open", url, 0, 0, 1)
End Sub
Private Sub Timer1_Timer()
Slider1.Value = Slider1.Value + 3
If Slider1.Value = 255 Then
    Timer1.Enabled = False
End If
End Sub
Private Sub Form_Load()
Me.OLEDropMode = 1 ' Ativa o modo de arrastar e soltar
Label3.ForeColor = vbBlack ' Cor do link
Me.Width = 4905
Me.Height = 4875
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim j As Long
If Data.GetFormat(vbCFFiles) = True Then
    If Not (GetAttr(Data.Files.Item(1)) And vbDirectory) Then
        For j = 1 To Data.Files.Count
            Label4.Caption = Data.Files.Item(j)
            Dim lngFileSize As Long
            lngFileSize = FileLen(Data.Files.Item(j))
            Me.Caption = lngFileSize
            Label8.Visible = True
            Label8.Caption = "Tamanho: " & lngFileSize & "B"
            Label1.Caption = "Arquivo pronto. Ligue seu aparelho no Recovery"
            Label2.Caption = "conecte o cabo USB e selecione o Sideload Mode"
            Label9.Visible = True
            Next
        End If
    End If
End Sub
Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) = True Then
        If (GetAttr(Data.Files.Item(1)) And vbDirectory) Then ' don't accept folders
            Effect = vbDropEffectNone
        Else
            Effect = vbDropEffectCopy
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.ForeColor = vbBlack
End Sub
Private Sub Image1_Click()
form1.Show
End Sub
Private Sub Label17_Click()
Dim Opcao As String
Opcao = MsgBox("Deseja voltar ao menu inicial?", vbYesNo + vbQuestion, "Cancelar Instalação")
    If Opcao = vbYes Then
Install.Visible = False
    Else
End If
End Sub
Private Sub Label3_Click()
Help.Enabled = True
Help.Visible = True
Help.Move 0, 0
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.ForeColor = vbBlue
End Sub
Private Sub Label5_Click()
Help.Visible = False
End Sub
Private Sub Label9_Click()
If Dir("C:\Program Files\Minimal ADB and Fastboot\adb.exe") <> "" Then
    Install.Visible = True
    Install.Move 0, 0
    MsgBox "1 - Desligue o celular" & vbCrLf & "2 - Entre no modo recuperação (CWM/TWRP)" & vbCrLf & "3 - Selecione Install ZIP/ADB Sideload" & vbCrLf & "4 - Caso haja alguma confirmação, clique ou arraste o botão" & vbCrLf & "5 - Conecte o cabo USB", vbExclamation, "Instruções"
        Set fs = CreateObject("Scripting.FileSystemObject")
        MkDir "C:\.tmp-rom"
        fs.CopyFile Label4.Caption, "C:\.tmp-rom\"
        On Error Resume Next
Else
    Dim Resposta As Integer
    Resposta = MsgBox("Parece que você não tem o Minimal ADB instalado," & vbCrLf & "deseja instalá-lo rapidamente?", vbYesNo + vbInformation, "ADB")
    If Resposta = vbYes Then
        Me.Hide
        form1.Show
    End If
    If Resposta = vbNo Then
    End If
End If
End Sub
Private Sub Timer3_Timer()
Slider1.Value = Slider1.Value - 3
If Slider1.Value = 0 Then
    End
    Timer3.Enabled = False
End If
End Sub
