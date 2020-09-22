VERSION 5.00
Begin VB.UserControl SplitPanel 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1170
   ControlContainer=   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   1170
   ToolboxBitmap   =   "SplitPanel.ctx":0000
   Begin VB.PictureBox Splitter 
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   855
      ScaleHeight     =   3540
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   30
      Width           =   165
   End
End
Attribute VB_Name = "SplitPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const SPLITWIDTH As Single = 80

Private mHorizontalSplit As Boolean
Private mControl1 As Object
Private mControl2 As Object
Private mSplitPercent As Single

Public Event Resize()



Public Property Get HorizontalSplit() As Boolean
    HorizontalSplit = mHorizontalSplit
End Property
Public Property Let HorizontalSplit(val As Boolean)
    mHorizontalSplit = val
    If mHorizontalSplit Then
        Splitter.MousePointer = 7
    Else
        Splitter.MousePointer = 9
    End If
    PropertyChanged "HorizontalSplit"
    UserControl_Resize
End Property

Public Property Get Control1() As Object
    Set Control1 = mControl1
End Property
Public Property Set Control1(ctl As Object)
    Set mControl1 = ctl
    PropertyChanged "Control1"
    UserControl_Resize
End Property

Public Property Get Control2() As Object
    Set Control2 = mControl2
End Property
Public Property Set Control2(ctl As Object)
    Set mControl2 = ctl
    PropertyChanged "Control2"
    UserControl_Resize
End Property

Public Property Get SplitPercent() As Byte
    SplitPercent = mSplitPercent * 100
End Property
Public Property Let SplitPercent(val As Byte)
    mSplitPercent = val / 100
    PropertyChanged "SplitPercent"
    UserControl_Resize
End Property

Private Sub UserControl_InitProperties()
    HorizontalSplit = False
    SplitPercent = 50
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    HorizontalSplit = PropBag.ReadProperty("HorizontalSplit", False)
    SplitPercent = PropBag.ReadProperty("SplitPercent", 50)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HorizontalSplit", HorizontalSplit, False
    PropBag.WriteProperty "SplitPercent", SplitPercent, 50
End Sub

Private Sub splitter_mousedown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Splitter.BackColor = &H80000008
    Splitter.ZOrder
End Sub

Private Sub splitter_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If mHorizontalSplit Then
            y = Splitter.Top - (SPLITWIDTH - y)
            mSplitPercent = y / UserControl.Height
            Splitter.Move 0, y
        Else
            x = Splitter.Left - (SPLITWIDTH - x)
            mSplitPercent = x / UserControl.Width
            Splitter.Move x
        End If
        If mSplitPercent < 0.1 Then mSplitPercent = 0.1
        If mSplitPercent > 0.9 Then mSplitPercent = 0.9
    End If
End Sub

Private Sub splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Splitter.BackColor = &H8000000F
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    If UserControl.Ambient.UserMode Then
        UserControl.BorderStyle = 0
    End If
    
    Dim pane1 As Single
    Dim pane2 As Single
    Dim totwidth As Single
    Dim totheight As Single
    totwidth = UserControl.Width
    totheight = UserControl.Height
    If mHorizontalSplit Then
        pane1 = (totheight - SPLITWIDTH) * mSplitPercent
        pane2 = (totheight - SPLITWIDTH) * (1 - mSplitPercent)
        mControl1.Move 0, 0, totwidth, pane1
        mControl2.Move 0, pane1 + SPLITWIDTH, totwidth, pane2
        Splitter.Move 0, pane1, totwidth, SPLITWIDTH
    Else
        pane1 = (totwidth - SPLITWIDTH) * mSplitPercent
        pane2 = (totwidth - SPLITWIDTH) * (1 - mSplitPercent)
        mControl1.Move 0, 0, pane1, totheight
        mControl2.Move pane1 + SPLITWIDTH, 0, pane2, totheight
        Splitter.Move pane1, 0, SPLITWIDTH, totheight
    End If
    RaiseEvent Resize
End Sub



Public Sub Refresh()
   UserControl.Refresh
End Sub
