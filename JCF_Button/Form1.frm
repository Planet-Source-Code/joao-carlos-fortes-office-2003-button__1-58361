VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   390
      ScaleWidth      =   5490
      TabIndex        =   7
      Top             =   345
      Width           =   5490
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   0
         Picture         =   "Form1.frx":925B
         ScaleHeight     =   390
         ScaleWidth      =   5085
         TabIndex        =   8
         Top             =   0
         Width           =   5085
         Begin Project1.JCF_Button btn1 
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   9
            Top             =   15
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "&Landscape"
            IsCheckButton   =   -1  'True
            Picture         =   "Form1.frx":9853
         End
         Begin Project1.JCF_Button btn1 
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            Top             =   15
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   661
            Caption         =   "&Portrait"
            IsCheckButton   =   -1  'True
            Picture         =   "Form1.frx":9AD5
         End
         Begin Project1.JCF_Button btn1 
            Height          =   375
            Index           =   0
            Left            =   270
            TabIndex        =   11
            ToolTipText     =   "Enviar por correio electrónico"
            Top             =   15
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   661
            Caption         =   "&Send"
            Picture         =   "Form1.frx":9D57
            PictureDisabled =   "Form1.frx":9FED
         End
         Begin Project1.JCF_Button btn1 
            Height          =   375
            Index           =   3
            Left            =   3855
            TabIndex        =   12
            Top             =   15
            Width           =   1035
            _ExtentX        =   1746
            _ExtentY        =   661
            Caption         =   "&Justified"
            IsCheckButton   =   -1  'True
            Picture         =   "Form1.frx":A197
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   4875
            Picture         =   "Form1.frx":A334
            Top             =   0
            Width           =   210
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   1170
            Picture         =   "Form1.frx":A4DD
            Top             =   0
            Width           =   45
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   3705
            Picture         =   "Form1.frx":A594
            Top             =   0
            Width           =   45
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      Picture         =   "Form1.frx":A64B
      ScaleHeight     =   345
      ScaleWidth      =   5490
      TabIndex        =   6
      Top             =   0
      Width           =   5490
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disabled"
      Height          =   390
      Index           =   4
      Left            =   4320
      TabIndex        =   4
      Top             =   1410
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OverDown"
      Height          =   390
      Index           =   3
      Left            =   3270
      TabIndex        =   3
      Top             =   1410
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Down"
      Height          =   390
      Index           =   2
      Left            =   2235
      TabIndex        =   2
      Top             =   1410
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Over"
      Height          =   390
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   1410
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Normal"
      Height          =   390
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   1410
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Apply states on &send button:"
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   1110
      Width           =   4275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    btn1(0).State = Index
End Sub

Private Sub btn1_Click(Index As Integer)

    ' prevent reentrant calls (we have a DoEvents in here)
    Static bBusy%
    If bBusy Then Exit Sub
    bBusy = True
    
    If Index = 1 Then
        If btn1(2).Value = 0 Then
            btn1(2).Value = 1
        Else
            btn1(2).Value = 0
        End If
    End If

    If Index = 2 Then
        If btn1(1).Value = 0 Then
            btn1(1).Value = 1
        Else
            btn1(1).Value = 0
        End If
    End If

    bBusy = False
End Sub

Private Sub Form_Load()
    btn1(1).Value = 1
End Sub
