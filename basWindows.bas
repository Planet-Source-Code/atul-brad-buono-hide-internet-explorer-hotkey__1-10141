Attribute VB_Name = "basWindows"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Public Type IE_STATE_SAVE
    hWnd As Long
    wp As WINDOWPLACEMENT
End Type

'used for debugging
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassNameA Lib "user32" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemoryH Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)

Global Const GWL_WNDPROC = (-4)     'used as paramater of SetWindowLong, sets the window procedure
Global Const SW_HIDE = 0            'constant passed to showCmd member of WINDOWPLACEMENT structure to hide the window
Global Const WM_HOTKEY = &H312      'message sent when hotkey registered with RegisterHotKey is pressed
Public Const MOD_SHIFT = &H4        'check the WM_HOTKEY message to see if the shift key was held down
Global Const MOD_WIN = &H8          'check the WM_HOTKEY message to see if the windows key was held down
Global Const VK_Z = &H5A            'virtual key code for Z

Global Const MIN_HOTKEY = &H5F      'user defined, this is my unique (to this process) id for the minimize hotkey
Global Const RST_HOTKEY = &H6F      'user defined, this is my unique (to this process) id for the restore hotkey

Public lngOldWindowProc As Long     'this stores the address of the original window procedure
Private arrayIESS() As IE_STATE_SAVE 'array used to store the windows and their positions

'used for debugging only, so i could find out the ie class names
'Public Function GetCaption(hWnd As Long) As String
'    Dim strBuffer As String
'    Dim lngReturn As Long
'
'    strBuffer = Space(255)
'
'    lngReturn = GetWindowText(hWnd, strBuffer, Len(strBuffer))
'
'    GetCaption = Left(strBuffer, lngReturn)
'
'End Function

Public Sub Main()
    Load frmMain
End Sub

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    Dim l As Long
    Dim sClsNm As String
    Dim ss As IE_STATE_SAVE
    Dim wp As WINDOWPLACEMENT

    wp.Length = Len(wp)

    sClsNm = GetClassName(hWnd)

    'Debug.Print hWnd, sClsNm, GetCaption(hWnd)
    'used for debugging to find the class names for the ie windows

    If sClsNm = "IEFrame" Or sClsNm = "CabinetWClass" Then 'If it is an IE window then...
        GetWindowPlacement hWnd, wp         'get the windowplacement
        ss.hWnd = hWnd                      'store the hwnd
        ss.wp = wp                          'store the windowplacement
        l = UBound(arrayIESS)               'we are going to place it at the ubound of the array
        arrayIESS(l) = ss                   'save our IE_STATE_SAVE to the array
        ReDim Preserve arrayIESS(l + 1)     'make room for the next one
        wp.showCmd = SW_HIDE                'modify the windowplacement to hide the window
        SetWindowPlacement hWnd, wp         'reapply it to the IE window
    End If

    EnumWindowsProc = True                  'this must return true to keep enumerating

End Function

Public Function SubProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If wMsg = WM_HOTKEY Then    'the hotkey message was fired
        If LoWord(lParam) = MOD_WIN And HiWord(lParam) = VK_Z And wParam = MIN_HOTKEY Then
            MinimizeIE
        End If
        If LoWord(lParam) = (MOD_WIN + MOD_SHIFT) And HiWord(lParam) = VK_Z And wParam = RST_HOTKEY Then
            RestoreIE
        End If
    End If
    SubProc = CallWindowProc(lngOldWindowProc, hWnd, wMsg, wParam, lParam)
End Function

Public Sub MinimizeIE()

    ReDim arrayIESS(0) 'clear out all the current IE_STATE_SAVEs

    EnumWindows AddressOf EnumWindowsProc, 0 'lParam is unused

End Sub

Public Sub RestoreIE()
'this just restores all the windows to the positions they had when MinimizeIE was called
On Error GoTo exit_RestoreIE

    Dim ieSS As IE_STATE_SAVE
    Dim l As Long

    For l = UBound(arrayIESS) To LBound(arrayIESS) Step -1
        ieSS = arrayIESS(l)
        With ieSS
            If .hWnd > 0 Then
                SetWindowPlacement .hWnd, .wp
            End If
        End With
    Next

exit_RestoreIE:
    Exit Sub

End Sub

Private Function GetClassName(ByVal hWnd As Long) As String
    Dim lngReturn As Long
    Dim strReturn As String
    
    strReturn = Space(255)

    lngReturn = GetClassNameA(hWnd, strReturn, Len(strReturn))

    GetClassName = Left(strReturn, lngReturn)

End Function

Public Function LoWord(ByVal dw As Long) As Integer
On Error GoTo Err_LoWord

    CopyMemoryH LoWord, ByVal VarPtr(dw), 2

Exit_LoWord:
    Exit Function

Err_LoWord:
    GoTo Exit_LoWord

End Function

Public Function HiWord(ByVal dw As Long) As Integer
On Error GoTo Err_HiWord

    CopyMemoryH HiWord, ByVal VarPtr(dw) + 2, 2

Exit_HiWord:
    Exit Function

Err_HiWord:
    GoTo Exit_HiWord

End Function
