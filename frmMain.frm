VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mPopupSys 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mPopHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "HideIE" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    RegisterHotKey Me.hwnd, MIN_HOTKEY, MOD_WIN, VK_Z
    RegisterHotKey Me.hwnd, RST_HOTKEY, MOD_WIN + MOD_SHIFT, VK_Z
    lngOldWindowProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf SubProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnregisterHotKey Me.hwnd, MIN_HOTKEY
    UnregisterHotKey Me.hwnd, RST_HOTKEY
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, lngOldWindowProc)
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
    Dim Result As Long
    Dim msg As Long

'the value of X will vary depending upon the scalemode setting
If Me.ScaleMode = vbPixels Then
    msg = X
Else
    msg = X / Screen.TwipsPerPixelX
End If

Select Case msg
    Case WM_RBUTTONUP        '517 display popup menu
        Result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu Me.mPopupSys
End Select

End Sub

Private Sub mPopExit_Click()
    Unload Me
End Sub

Private Sub mPopHide_Click()
    MinimizeIE
End Sub

Private Sub mPopShow_Click()
    RestoreIE
End Sub
