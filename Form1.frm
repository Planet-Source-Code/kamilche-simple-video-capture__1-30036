VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   Begin VB.CheckBox chkPaused 
      BackColor       =   &H00FF00FF&
      Caption         =   "Paused"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   4185
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   2670
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1830
      Top             =   1350
   End
   Begin VB.Label lblPaused 
      BackStyle       =   0  'Transparent
      Caption         =   "Paused"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4440
      TabIndex        =   4
      Top             =   2670
      Width           =   690
   End
   Begin VB.Label lblSource 
      BackStyle       =   0  'Transparent
      Caption         =   "Set video source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   255
      Width           =   1500
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "Set video size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   1365
   End
   Begin VB.Label lblQuit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4920
      TabIndex        =   0
      Top             =   -60
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants
Private Const WM_CAP_DRIVER_CONNECT As Long = 1034
Private Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
Private Const WM_CAP_GRAB_FRAME As Long = 1084
Private Const WM_CAP_EDIT_COPY As Long = 1054
Private Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Private Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Private Const WM_CLOSE = &H10

'Declarations
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long

'Variables
Private mCapHwnd As Long
Private mLivePreview As Boolean

'Methods
Private Sub Form_Load()
    mCapHwnd = capCreateCaptureWindow("My Own Capture Window", 0, 0, 0, 320, 240, Me.hwnd, 0)
    LivePreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LivePreview = False
    SendMessage mCapHwnd, WM_CLOSE, 0, 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_Resize()
    Static Running As Boolean
    If Running = True Then
        Exit Sub
    End If
    Running = True
    If WindowState <> vbMinimized Then
        GrabFrame
        Width = ScaleX(Picture.Width, vbHimetric, vbTwips)
        Height = ScaleY(Picture.Height, vbHimetric, vbTwips)
        lblQuit.Move ScaleWidth - lblQuit.Width
        lblPaused.Move ScaleWidth - lblPaused.Width, ScaleHeight - lblPaused.Height
        chkPaused.Move lblPaused.Left - chkPaused.Width - 3, lblPaused.Top
    End If
    Running = False
End Sub

Private Sub lblQuit_Click()
    Unload Me
End Sub

Private Sub lblSize_Click()
    SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0
    Form_Resize
End Sub

Private Sub lblSource_Click()
    SendMessage mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0
    Form_Resize
End Sub

Private Sub Timer1_Timer()
    GrabFrame
End Sub

Private Sub GrabFrame()
    On Error Resume Next
    SendMessage mCapHwnd, WM_CAP_GRAB_FRAME, 0, 0
    SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
    Set Picture = Clipboard.GetData
End Sub

Private Sub lblPaused_Click()
    chkPaused.Value = IIf(chkPaused.Value = 0, 1, 0)
End Sub

Private Sub chkPaused_Click()
    If chkPaused.Value = 1 Then
        LivePreview = False
    Else
        LivePreview = True
    End If
End Sub

Friend Property Let LivePreview(ByVal b As Boolean)
    'Toggle video on or off
    If b = False Then
        'Turn off video
        If mLivePreview = True Then
            SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
        End If
    Else
        'Turn on video
        If mLivePreview = False Then
            SendMessage mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0
        End If
    End If
    mLivePreview = b
    Timer1.Enabled = b
End Property


