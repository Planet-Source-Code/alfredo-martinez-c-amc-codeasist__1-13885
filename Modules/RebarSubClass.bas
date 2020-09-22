Attribute VB_Name = "RebarSubClass"
Option Private Module
Option Explicit
Public NextProcs As Long
 
Public Nodef As Boolean

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long

 
 
Public Const WM_COMMAND = &H111
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = -4

Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Type POINTAPI
        x As Long
        y As Long
End Type
Private Const WM_USER = &H400

Public Const RBHT_CAPTION = &H2
Public Const RBHT_CLIENT = &H3
Public Const RBHT_GRABBER = &H4

Public Const RB_HITTEST = (WM_USER + 8)
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_NOTIFY = &H4E


Public Type RBHITTESTINFO
    ptApi As POINTAPI
    flags As Long
    iBand As Long
End Type


 Public Type NMREBAR
       NMHDR As Long
        uBand As Long
       wID As Long
       cyChild As Long
        cyBand As Long
        
End Type

 
Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

 
 
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
  
 Select Case hwnd
 
        
        Case MDIGen.hwnd
             MDIGen.ProcMsg hwnd, uMsg, wParam, lParam, 0& ', 0&
      
    End Select
    If Nodef = True Then
    WindowProc = CallWindowProc(NextProcs, hwnd, uMsg, wParam, ByVal lParam)
    Else
    Nodef = False
    Nodef = True
    End If
End Function

   

 
' Public Sub SubClass(hwnd As Long)
'NextProcs = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
' End Sub
'Public Sub UnSubClass()
'Dim hWndCur As Long
'    hWndCur = Form1.hwnd
'    If NextProcs Then
'        SetWindowLong hWndCur, GWL_WNDPROC, NextProcs
 '       NextProcs = 0
  '  End If
'End Function

