VERSION 5.00
Begin VB.UserControl JCF_Button 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   LockControls    =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   2685
   ToolboxBitmap   =   "JCF_Button.ctx":0000
   Begin VB.PictureBox PicOverDown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "JCF_Button.ctx":0312
      ScaleHeight     =   330
      ScaleWidth      =   2250
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   300
      Picture         =   "JCF_Button.ctx":0508
      ScaleHeight     =   375
      ScaleWidth      =   30
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.PictureBox picLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      Picture         =   "JCF_Button.ctx":0612
      ScaleHeight     =   375
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.PictureBox picBottom 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   0
      Picture         =   "JCF_Button.ctx":071C
      ScaleHeight     =   60
      ScaleWidth      =   2250
      TabIndex        =   4
      Top             =   320
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picUp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      Picture         =   "JCF_Button.ctx":0E6E
      ScaleHeight     =   30
      ScaleWidth      =   2250
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Timer tmrToolTip 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   1440
   End
   Begin VB.PictureBox PicDown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "JCF_Button.ctx":1238
      ScaleHeight     =   330
      ScaleWidth      =   2250
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox PicOver 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "JCF_Button.ctx":3952
      ScaleHeight     =   330
      ScaleWidth      =   2250
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox PicNormal 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      Picture         =   "JCF_Button.ctx":606C
      ScaleHeight     =   375
      ScaleWidth      =   2250
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   50
      ToolTipText     =   "ToolTipText"
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   90
      Width           =   360
   End
End
Attribute VB_Name = "JCF_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==========================================================================
'Copyright © 2004 João Fortes, todos os direitos reservados
'
'   JCF_ToolButton
'   Data: 19-Ago-2004
'
'   Notas:
'
'
'==========================================================================
Option Explicit

'API constants
Private Const SW_SHOWNOACTIVATE = 4
Private Const HWND_TOP = 0
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40

'API functions
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As TYPERECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Type
Private Type TYPERECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

'constants
Const TEXT_ACTIVE = &H0&
Const TEXT_INACTIVE = &H80000011

'state constants
Const STA_NORMAL = 0
Const STA_OVER = 1
Const STA_DOWN = 2
Const STA_OVERDOWN = 3
Const STA_DISABLED = 4

'value constants
Const VAL_UNCHECKED = 0
Const VAL_CHECKED = 1
Const VAL_GRAY = 2

'events
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'members
Dim m_Picture As StdPicture
Dim m_PictureOver As StdPicture
Dim m_PictureDown As StdPicture
Dim m_PictureOverDown As StdPicture
Dim m_PictureDisabled As StdPicture

Dim m_State As Integer
Dim m_Value As Integer
Dim m_Enabled As Boolean
Dim m_Caption As String
Dim m_IsCheckButton As Boolean

Dim m_ToolTipBackColor As OLE_COLOR
Dim m_ToolTipForeColor As OLE_COLOR
Dim m_ToolTipText As String

'local
Dim tmpState As Integer
Dim tmpDrawState As Integer
'==========================================================================
' Init, Read & Write UserControl
'==========================================================================
Private Sub UserControl_InitProperties()
    m_ToolTipBackColor = &H80000018
    m_ToolTipForeColor = &H80000012
    m_ToolTipText = Extender.ToolTipText
End Sub

Private Sub UserControl_Initialize()
    tmpDrawState = -1
    Set UserControl.Picture = PicNormal.Picture
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 375
    picRight.Left = UserControl.Width - 30
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_State = .ReadProperty("State", 0)
        m_Enabled = .ReadProperty("Enabled", True)
        m_Caption = .ReadProperty("Caption", Empty)
        m_IsCheckButton = .ReadProperty("IsCheckButton", False)
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Set m_PictureOver = .ReadProperty("PictureOver", Nothing)
        Set m_PictureDown = .ReadProperty("PictureDown", Nothing)
        Set m_PictureOverDown = .ReadProperty("PictureOverDown", Nothing)
        Set m_PictureDisabled = .ReadProperty("PictureDisabled", Nothing)
    End With
    
    Call DrawUserControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("State", m_State, 0)
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("Caption", m_Caption, Empty)
        Call .WriteProperty("IsCheckButton", m_IsCheckButton, False)
        Call .WriteProperty("Picture", m_Picture, Nothing)
        Call .WriteProperty("PictureOver", m_PictureOver, Nothing)
        Call .WriteProperty("PictureDown", m_PictureDown, Nothing)
        Call .WriteProperty("PictureOverDown", m_PictureOverDown, Nothing)
        Call .WriteProperty("PictureDisabled", m_PictureDisabled, Nothing)
    End With
End Sub
'==========================================================================
' Down
'==========================================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UnallowToolTip
    
    'Is disabled
    If m_State = STA_DISABLED Or Not m_Enabled Then Exit Sub
    'only LeftButton
    If Button <> vbLeftButton Then Exit Sub
    
    tmpState = m_State
    m_State = STA_DOWN
    DrawUserControl
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub
'==========================================================================
' Up
'==========================================================================
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Is disabled
    If m_State = STA_DISABLED Or Not m_Enabled Then Exit Sub
    
    'type of button
    If m_IsCheckButton Then
        If m_Value = VAL_UNCHECKED Then
            m_Value = VAL_CHECKED
        ElseIf m_Value = VAL_CHECKED Then
            m_Value = VAL_UNCHECKED
        End If
    Else
        m_State = STA_NORMAL
    End If
    
    'refresh control
    DrawUserControl
    
    'Fire event
    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent Click
End Sub
'==========================================================================
' Move
'==========================================================================
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Is disabled
    If m_State = STA_DISABLED Or Not m_Enabled Then Exit Sub
    
    ' Either Capture/ReleaseCapture hwnd first, so that if MouseOver is on lblDisp
    ' or picDisp, tooltiptext will still be displayed
    If GetCapture() <> UserControl.hwnd Then
         SetCapture UserControl.hwnd
    End If
    If X < 0 Or X > UserControl.ScaleWidth Or Y < 0 Or Y > UserControl.ScaleHeight Then
         ' Ensure no tooltip
         RemoveToolTip
         ReleaseCapture
    End If
        
    If Len(Extender.ToolTipText) > 0 Then
        ShowToolTip Extender.ToolTipText
    Else
        UnallowToolTip
    End If
    
    If CheckMouseOver Then
        If m_IsCheckButton Then
            If m_Value = VAL_UNCHECKED Then
                m_State = STA_OVER
            ElseIf m_Value = VAL_CHECKED Then
                m_State = STA_OVERDOWN
            End If
            DrawUserControl True
        Else
            If m_State = STA_NORMAL Then
                m_State = STA_OVER
            End If
            DrawUserControl True
        End If
    Else
        If m_State = STA_OVER Then
            m_State = STA_NORMAL
        ElseIf m_State = STA_OVERDOWN Then
            m_State = STA_DOWN
        End If
        
        DrawUserControl False
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

'==========================================================================
' ToolTipText
'==========================================================================
Private Sub tmrToolTip_Timer()
    On Error Resume Next
    If GetCapture() = UserControl.hwnd Then
        doToolTip tmrToolTip.Tag
    Else
        UnallowToolTip
    End If
End Sub

Private Sub UnallowToolTip()
    On Error Resume Next
    tmrToolTip.Enabled = False
    RemoveToolTip
End Sub

Private Sub RemoveToolTip()
    On Error Resume Next
    Unload frmTooltip
End Sub

Private Sub ShowToolTip(ByVal inText As String)
    On Error Resume Next
    If inText = "" Then
        RemoveToolTip
        tmrToolTip.Enabled = False
    Else
        tmrToolTip.Enabled = True              ' Let timer do it
        tmrToolTip.Tag = inText
    End If
End Sub

Private Sub doToolTip(ByVal inText)
    On Error GoTo errHandler
    Dim X As Long, Y As Long
    Dim adjW As Long, adjH As Long
    Dim textW As Long, textH As Long
    Dim typRect As TYPERECT
    Dim i As Integer
    
    GetWindowRect UserControl.hwnd, typRect
    X = (typRect.Left + (typRect.Right - typRect.Left) / 3) * Screen.TwipsPerPixelX
    Y = (typRect.Bottom + 8) * Screen.TwipsPerPixelY
    
    adjW = 10 * Screen.TwipsPerPixelX
    adjH = 8 * Screen.TwipsPerPixelY
    
        ' Ensure tooltiptext is not too long
    i = frmTooltip.TextWidth(inText)
    Do While i > (Screen.Width * 80 / 100)
         inText = Left(inText, Len(inText) - 1)
         i = frmTooltip.TextWidth(inText)
    Loop
    
    textW = frmTooltip.TextWidth(inText) + adjW
    textH = frmTooltip.TextHeight(inText) + adjH
    
    If X < 0 Then
         X = 0
    ElseIf (X + textW) > Screen.Width Then
         X = Screen.Width - textW
    End If
    If (Y + textH) > Screen.Height Then
         Y = (typRect.Top - 2) * Screen.TwipsPerPixelY - textH
    End If
    
    With frmTooltip
        .BackColor = &H80000018
        .lblToolTipText.Width = textW
        .lblToolTipText.Height = textH
        .lblToolTipText.BackColor = &H80000018
        .lblToolTipText.ForeColor = m_ToolTipForeColor
        .lblToolTipText.Caption = inText
        .lblToolTipText.Refresh
        .Move X, Y, textW, textH
              
        SetWindowPos .hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or _
           SWP_NOSIZE Or SWP_SHOWWINDOW
    End With
    Exit Sub
errHandler:
    RemoveToolTip
End Sub
'==========================================================================
' Properties
'==========================================================================
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    DrawUserControl
    PropertyChanged "Value"
End Property

Public Property Get IsCheckButton() As Boolean
    IsCheckButton = m_IsCheckButton
End Property

Public Property Let IsCheckButton(ByVal New_Value As Boolean)
    m_IsCheckButton = New_Value
    DrawUserControl
    PropertyChanged "IsCheckButton"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Value As String)
    m_Caption = New_Value
    DrawUserControl
    PropertyChanged "Caption"
End Property

Public Property Get State() As Integer
    State = m_State
End Property

Public Property Let State(ByVal New_Value As Integer)
    m_State = New_Value
    DrawUserControl
    PropertyChanged "State"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Value As Boolean)
    m_Enabled = New_Value
    PropertyChanged "Enabled"
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    DrawUserControl
    PropertyChanged "Picture"
End Property

Public Property Get PictureOver() As StdPicture
    Set PictureOver = m_PictureOver
End Property

Public Property Set PictureOver(ByVal New_Picture As StdPicture)
    Set m_PictureOver = New_Picture
    PropertyChanged "PictureOver"
End Property

Public Property Get PictureDown() As StdPicture
    Set PictureDown = m_PictureDown
End Property

Public Property Set PictureDown(ByVal New_Picture As StdPicture)
    Set m_PictureDown = New_Picture
    PropertyChanged "PictureDown"
End Property

Public Property Get PictureOverDown() As StdPicture
    Set PictureOverDown = m_PictureOverDown
End Property

Public Property Set PictureOverDown(ByVal New_Picture As StdPicture)
    Set m_PictureOverDown = New_Picture
    PropertyChanged "PictureOverDown"
End Property

Public Property Get PictureDisabled() As StdPicture
    Set PictureDisabled = m_PictureDisabled
End Property

Public Property Set PictureDisabled(ByVal New_Picture As StdPicture)
    Set m_PictureDisabled = New_Picture
    PropertyChanged "PictureDisabled"
End Property
'==========================================================================
' Functions
'==========================================================================
Private Sub DrawUserControl(Optional IsOver As Boolean = False)

    If m_State <> STA_DISABLED Then
        If m_IsCheckButton Then
            If IsOver Then
                If m_Value = VAL_CHECKED Then
                    m_State = STA_OVERDOWN
                Else
                    m_State = STA_OVER
                End If
            Else
                If m_Value = VAL_CHECKED Then
                    m_State = STA_DOWN
                Else
                    m_State = STA_NORMAL
                End If
            End If
        End If
    End If
    
    'No need to redraw it
    If tmpDrawState > 0 And m_State = tmpDrawState Then Exit Sub
    
    Select Case m_State
        Case STA_NORMAL
            UserControl.Picture = PicNormal
            picUp.Visible = False
            picBottom.Visible = False
            picLeft.Visible = False
            picRight.Visible = False
            Label1.ForeColor = TEXT_ACTIVE
            Image1.Picture = m_Picture
            Label1.Caption = m_Caption
        Case STA_OVER
            UserControl.Picture = PicOver
            picUp.Visible = True
            picBottom.Visible = True
            picLeft.Visible = True
            picRight.Visible = True
            Label1.ForeColor = TEXT_ACTIVE
            Image1.Picture = m_Picture
            Label1.Caption = m_Caption
        Case STA_DOWN
            UserControl.Picture = PicDown
            picUp.Visible = True
            picBottom.Visible = True
            picLeft.Visible = True
            picRight.Visible = True
            Label1.ForeColor = TEXT_ACTIVE
            Image1.Picture = m_Picture
            Label1.Caption = m_Caption
        Case STA_OVERDOWN
            UserControl.Picture = PicOverDown
            picUp.Visible = True
            picBottom.Visible = True
            picLeft.Visible = True
            picRight.Visible = True
            Label1.ForeColor = TEXT_ACTIVE
            Image1.Picture = m_Picture
            Label1.Caption = m_Caption
        Case STA_DISABLED
            UserControl.Picture = PicNormal
            Label1.ForeColor = TEXT_INACTIVE
            Image1.Picture = m_PictureDisabled
            Label1.Caption = m_Caption
    End Select

    tmpDrawState = m_State
End Sub

Private Function CheckMouseOver() As Boolean
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.X, pt.Y) = UserControl.hwnd)
End Function

