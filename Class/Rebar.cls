VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public RebarWindow As Long
Private RebarChildWin As Object
Dim RebarBand As tagRebarBandInfo

Dim a As Variant
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Dim i As Integer

Public Enum BandPosition
AddNewRow = 1
AddToEnd = 2
End Enum

Private Const HWND_TOPMOST = -1
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREDRAW = &H8
Private Const SWP_SHOWWINDOW = &H40

Private RebarPic As Object

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wcmd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendTBMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
ByVal lParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const WM_USER = &H400
Dim b As Variant
'Toolbar Const
Private Const TBSTYLE_TRANSPARENT = &H8000 'Haven't gotton this one to work with regular toolbars yet.
Private Const TBSTYLE_FLAT = &H800
Private Const TB_SETSTYLE = (WM_USER + 56)
Private Const TB_GETSTYLE = (WM_USER + 57)
Private Const TBSTYLE_LIST = &H1000

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
 
'System Color Constants
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNTEXT = 18

Private Const REBARCLASSNAME = "ReBarWindow32"
Private Const RBN_FIRST = 0 - 831
Private Const RBN_LAST = 0 - 859
Private Const RBIM_IMAGELIST = &H1
Private Const ICC_COOL_CLASSES = &H400
'Rebar Styles
Private Const RBS_AUTOSIZE = &H2000
Private Const RBS_VERTICALGRIPPER = &H4000 '  // this always has the vertical gripper (default for horizontal mode)
Private Const RBS_TOOLTIPS = &H100
Private Const RBS_VARHEIGHT = &H200
Private Const RBS_BANDBORDERS = &H400
Private Const RBBS_VARIABLEHEIGHT = &H40
Private Const RBS_FIXEDORDER = &H800
Private Const RBBS_GRIPPERALWAYS = &H80      ' always show the gripper
Private Const RBBS_BREAK = &H1               ' break to new line
Private Const RBBS_FIXEDSIZE = &H2           ' band can't be sized
Private Const RBBS_CHILDEDGE = &H4           ' edge around top & bottom of child window
Private Const RBBS_HIDDEN = &H8              ' don't show
Private Const RBBS_NOVERT = &H10             ' don't show when vertical
Private Const RBBS_FIXEDBMP = &H20           ' bitmap doesn't move during band resize
Private Const RBBIM_STYLE = &H1
Private Const RBBIM_COLORS = &H2
Private Const RBBIM_TEXT = &H4
Private Const RBBIM_IMAGE = &H8
Private Const RBBIM_CHILD = &H10
Private Const RBBIM_CHILDSIZE = &H20
Private Const RBBIM_SIZE = &H40
Private Const RBBIM_BACKGROUND = &H80
Private Const RBBIM_ID = &H100
Private Const RB_BEGINDRAG = (WM_USER + 24)
Private Const RB_ENDDRAG = (WM_USER + 25)
Private Const RB_DRAGMOVE = (WM_USER + 26)
Private Const RB_HITTEST = (WM_USER + 8)
Private Const RB_INSERTBANDA = (WM_USER + 1)
Private Const RB_DELETEBAND = (WM_USER + 2)
Private Const RB_GETBARINFO = (WM_USER + 3)
Private Const RB_SETBARINFO = (WM_USER + 4)
Private Const RB_GETBANDINFO = (WM_USER + 5)
Private Const RB_SETBANDINFOA = (WM_USER + 6)
Private Const RB_SETPARENT = (WM_USER + 7)
Private Const RB_INSERTBANDW = (WM_USER + 10)
Private Const RB_SETBANDINFOW = (WM_USER + 11)
Private Const RB_GETBANDCOUNT = (WM_USER + 12)
Private Const RB_GETROWCOUNT = (WM_USER + 13)
Private Const RB_GETROWHEIGHT = (WM_USER + 14)
Private Const RB_SETBKCOLOR = (WM_USER + 19)
Private Const RB_GETBKCOLOR = (WM_USER + 20)
Private Const RB_SETTEXTCOLOR = (WM_USER + 21)
Private Const RB_GETTEXTCOLOR = (WM_USER + 22)
Private Const RBHT_NOWHERE = &H1
Private Const RBHT_CAPTION = &H2
Private Const RBHT_CLIENT = &H3
Private Const RBHT_GRABBER = &H4
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const RB_INSERTBAND = RB_INSERTBANDA
Private Const RB_SETBANDINFO = RB_SETBANDINFOA
Private Const RBN_HEIGHTCHANGE = (RBN_FIRST - 0)

'CreateWindowEx Constants
Private Const WS_BORDER = &H800000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
Private Const CCS_NORESIZE = &H4
Private Const CCS_NOPARENTALIGN = &H8
Private Const CCS_NODIVIDER = &H40
Private Const CCS_VERT = &H80

Private Type tagRebarInfo
    cbSize As Integer
    fMask As Integer
    himl As Long
End Type

 Private Type tagRebarBandInfo
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As String
    cch As Long
    iImage As Long
    hWndChild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long
    wID As Long
End Type
 
Private RebarhWnd As Long
Private RebarParent As Object
 
 
Sub TBMakeFlat(Tb As Object, Optional TBList As Boolean)
 
   Dim Style As Long
   Dim lRet As Long
   Dim ToolbarHandle As Long


   ToolbarHandle = FindWindowEx(Tb.hwnd, 0&, "ToolbarWindow32", vbNullString)

   Style = SendTBMessage(ToolbarHandle, TB_GETSTYLE, 0&, 0&)
   
   If TBList = True Then
   Style = Style Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT Or CCS_NODIVIDER Or TBSTYLE_LIST
   Else
   Style = Style Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT Or CCS_NODIVIDER
   End If
   
   lRet = SendTBMessage(ToolbarHandle, TB_SETSTYLE, 0, Style)

   Tb.Refresh
   
End Sub
Public Sub AddBands(BandText As String, BandName As Integer, _
ChildWin As Variant, NewRow As BandPosition, Optional mWidth As Variant)
      
On Error Resume Next
'Set structure

RebarBand.cbSize = LenB(RebarBand)

'Add Mask for all possibilities
RebarBand.fMask = RBBIM_COLORS Or RBBIM_IMAGE Or RBBIM_BACKGROUND Or RBBS_FIXEDBMP Or RBBIM_CHILD Or RBBIM_CHILDSIZE Or _
RBBIM_ID Or RBBIM_STYLE Or RBBIM_TEXT Or RBBS_BREAK

'Set RebarBands Colors
RebarBand.clrFore = GetSysColor(COLOR_BTNTEXT)
RebarBand.clrBack = GetSysColor(COLOR_BTNFACE)

If NewRow = AddNewRow Then
RebarBand.fStyle = RBBS_FIXEDBMP Or RBBS_CHILDEDGE Or RBS_VARHEIGHT Or RBBS_GRIPPERALWAYS
Else
RebarBand.fStyle = RBBS_FIXEDBMP Or RBBS_CHILDEDGE Or RBS_VARHEIGHT Or RBBS_BREAK Or RBBS_GRIPPERALWAYS
End If


'Add Band Text if any
RebarBand.lpText = BandText
 
'Set BackGround Picture
RebarBand.hbmBack = RebarPic.Picture
 
If ChildWin <> "" Then
RebarBand.hWndChild = ChildWin
Set RebarChildWin = ChildWin
End If

RebarBand.cxMinChild = ChildWin.Width / Screen.TwipsPerPixelX
If mWidth Then RebarBand.cxMinChild = mWidth * Screen.TwipsPerPixelX

'Band height
Dim ChildRect As RECT
Call GetWindowRect(ChildWin, ChildRect)
RebarBand.cyMinChild = (ChildRect.Bottom - ChildRect.Top)

'SetMin Height
RebarBand.cx = 10
 
RebarBand.wID = BandName
 
Call SendMessage(RebarhWnd, RB_INSERTBAND, -1, RebarBand) '-1 = add to end
Call UpdateWindow(RebarhWnd)
 
End Sub

Public Function GetBandId(id As Integer)
If IsNumeric(id) = False Then Exit Function

Dim RebarBand As tagRebarBandInfo
RebarBand.cbSize = LenB(RebarBand)

'Add Mask for all possibilities
RebarBand.fMask = RBBIM_COLORS Or RBBIM_IMAGE Or RBBIM_BACKGROUND Or RBBS_FIXEDBMP Or RBBIM_CHILD Or RBBIM_CHILDSIZE Or _
RBBIM_ID Or RBBIM_STYLE Or RBBIM_TEXT Or RBBS_BREAK Or RBS_FIXEDORDER Or RBBS_FIXEDSIZE
'Set RebarBands Colors
Dim xReturn As Long
xReturn = SendMessage(RebarhWnd, RB_GETBANDINFO, id, RebarBand)
GetBandId = RebarBand.wID

End Function


Public Sub SetBandColors()
 
'This procedure is used to set the band colors in case the system color changes

Dim RebarBand As tagRebarBandInfo
RebarBand.cbSize = LenB(RebarBand)

'Add Mask for all possibilities
RebarBand.fMask = RBBIM_COLORS Or RBBIM_IMAGE Or RBBIM_BACKGROUND Or RBBS_FIXEDBMP Or RBBIM_CHILD Or RBBIM_CHILDSIZE Or _
RBBIM_ID Or RBBIM_STYLE Or RBBIM_TEXT Or RBBS_BREAK Or RBS_FIXEDORDER Or RBBS_FIXEDSIZE
'Set RebarBands Colors
 
Dim xReturn As Long
Dim xCount As Integer
xCount = SendMessage(RebarhWnd, RB_GETBANDCOUNT, 0, 0)

For xCount = 0 To xCount - 1
xReturn = SendMessage(RebarhWnd, RB_GETBANDINFO, xCount, RebarBand)

RebarBand.clrFore = GetSysColor(COLOR_BTNTEXT)
RebarBand.clrBack = GetSysColor(COLOR_BTNFACE)

xReturn = SendMessage(RebarhWnd, RB_SETBANDINFO, xCount, RebarBand)
Next

End Sub

Public Sub DestroyRebar()
 On Error Resume Next
'We Must return all children that we did not create
'back to the form.

Dim hwnd As Long
Call ShowWindow(RebarhWnd, SW_HIDE)
'Get Window Handle of 1st child
hwnd = GetWindow(RebarhWnd, GW_CHILD)

Do

'Return Child
 Call SetParent(hwnd, Parent.hwnd)

'Get Next Window
hwnd = GetWindow(hwnd, GW_CHILD)

'Loop until we run out
Loop While hwnd <> 0

Dim BandCount As Integer, i As Integer

'Get Number of bands
BandCount = SendMessage(RebarhWnd, RB_GETBANDCOUNT, 0, 0)

For i = 0 To BandCount - 1
Call SendMessage(RebarhWnd, RB_DELETEBAND, 0, 0)
Next
 
'Now it's safe to kill the rebar
If RebarhWnd <> 0 Then
   Call DestroyWindow(RebarhWnd)
End If
End Sub

Public Function GetRebarWindow()
GetRebarWindow = RebarhWnd
End Function

Public Sub RemoveBands(Optional BandPosition As Integer = 0)
On Error Resume Next
Call SendMessage(RebarhWnd, RB_DELETEBAND, BandPosition, 0)
End Sub


 
 
Public Sub Resize(frm As Object)
 On Error Resume Next
 Call MoveWindow(RebarhWnd, 0, 0, frm.Width / Screen.TwipsPerPixelX - 8, 0, True)
 Call UpdateWindow(RebarhWnd)
End Sub

 

 
 
Public Sub Move(Left As Long, Top As Long, Width As Long, Height As Long)
On Error Resume Next
    If hwnd <> 0 Then
        Call MoveWindow(hwnd, Left, Top, Width, Height, True)
    End If
   
End Sub

Public Property Get hwnd() As Long
    hwnd = RebarhWnd
    RebarWindow = RebarhWnd
End Property

Public Function Create()
On Error Resume Next
    
    If Parent Is Nothing Then
       Create = False
       Exit Function
    End If
    
     
    'RBS_AUTOSIZE Or
    
    RebarhWnd = CreateWindowEX(0, "ReBarWindow32", "", _
    WS_VISIBLE Or WS_BORDER Or WS_CHILD Or _
    RBS_VARHEIGHT Or RBS_BANDBORDERS Or _
    CCS_NODIVIDER Or RBS_VERTICALGRIPPER Or _
    RBBS_VARIABLEHEIGHT Or CCS_NOPARENTALIGN Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS, _
    0, 0, Parent.Width, 60, _
    Parent.hwnd, 0&, App.hInstance, 0&)
  
     
    Call Move(CLng(0), CLng(0), CLng(Parent.Width), CLng(60))
    Call ShowWindow(RebarhWnd, SW_SHOWNORMAL)
    'Set Parent to receive messages
    Call SetParent(RebarhWnd, Parent.hwnd)
    Create = (RebarhWnd <> 0)
      
  
End Function

Public Property Set Parent(frm As Object)
    Set RebarParent = frm
End Property

Public Property Get Parent() As Object
    Set Parent = RebarParent
End Property

 
Public Property Set ImageForRebar(Img As Object)
On Error Resume Next
Set RebarPic = Img
End Property

 
 
Private Sub Class_Initialize()
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_COOL_CLASSES
    End With
    Call InitCommonControlsEx(iccex)
    RebarhWnd = 0
End Sub

Public Sub SetMainParent(Obj As Object)
  Call SendMessage(RebarhWnd, RB_SETPARENT, Obj.hwnd, 0)
End Sub
 
Private Sub Class_Terminate()
'Place this in your form Unload event
'YourDimName.DestroyRebar
End Sub

Public Function GetHeight()

a = 0
For i = 0 To SendMessage(RebarhWnd, RB_GETROWCOUNT, 0, 0) - 1
  a = a + SendMessage(RebarhWnd, RB_GETROWHEIGHT, i, 0)
Next
GetHeight = (a * Screen.TwipsPerPixelY) + (a) + 30
End Function
