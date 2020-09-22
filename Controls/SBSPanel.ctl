VERSION 5.00
Begin VB.UserControl SBSPanel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   ControlContainer=   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   2445
   ToolboxBitmap   =   "SBSPanel.ctx":0000
End
Attribute VB_Name = "SBSPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum BorderType
    None = 0
    Flat = 1
    Frame = 2
    Inset = 3
    Raised = 4
End Enum

Private Const iOffSet = 4
Private Const iOffSet2 = 15

Private mintBorder As BorderType

Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Resize()

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As BorderType
    BorderStyle = mintBorder
    DrawBorder
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderType)
    mintBorder = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_Resize()
    DrawBorder
    RaiseEvent Resize
End Sub

Public Property Get ScaleHeight() As Single
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

Public Property Get ScaleLeft() As Single
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

Public Property Get ScaleTop() As Single
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

Public Property Get ScaleWidth() As Single
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    mintBorder = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 6300)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 8025)
    
    DrawBorder
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", mintBorder, 0)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 6300)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 8025)
End Sub


Private Sub DrawBorder()
    Dim x1 As Single
    Dim X2 As Single
    Dim y1 As Single
    Dim Y2 As Single
    Dim c1 As Long
    Dim c2 As Long
    Dim c3 As Long
    Dim c4 As Long
        
    UserControl.ScaleLeft = 0
    UserControl.ScaleWidth = UserControl.Width
    UserControl.ScaleTop = 0
    UserControl.ScaleHeight = UserControl.Height
        
    x1 = iOffSet
    y1 = iOffSet
    X2 = UserControl.Width - (iOffSet * 2)
    Y2 = UserControl.Height - (iOffSet * 2)
    UserControl.Cls
        
    Select Case mintBorder
    Case Flat
        c1 = &H808080
        c2 = &H808080
        c3 = &H808080
        c4 = &H808080
    Case Frame
        c1 = &H808080
        c2 = &HFFFFFF
        c3 = &HFFFFFF
        c4 = &H808080
        
    Case Inset
        c1 = &H808080
        c2 = &HFFFFFF
        c3 = &HFFFFFF
        c4 = &H808080

    Case Raised
        c1 = &HFFFFFF
        c2 = &H808080
        c3 = &H808080
        c4 = &HFFFFFF
    End Select
            
    If mintBorder <> None Then
        UserControl.Line (x1, y1)-(X2, y1), c1
        UserControl.Line (X2, y1)-(X2, Y2), c2
        UserControl.Line (x1, Y2)-(X2, Y2), c3
        UserControl.Line (x1, y1)-(x1, Y2), c4
        
        If mintBorder = Frame Then
            c1 = &HFFFFFF
            c2 = &H808080
            c3 = &H808080
            c4 = &HFFFFFF
        
            x1 = x1 + iOffSet2
            X2 = X2 - iOffSet2
            y1 = y1 + iOffSet2
            Y2 = Y2 - iOffSet2
        
            UserControl.Line (x1, y1)-(X2, y1), c1
            UserControl.Line (X2, y1)-(X2, Y2), c2
            UserControl.Line (x1, Y2)-(X2, Y2), c3
            UserControl.Line (x1, y1)-(x1, Y2), c4
        End If
        
        UserControl.ScaleLeft = x1 + iOffSet
        UserControl.ScaleWidth = X2 - iOffSet
        UserControl.ScaleTop = y1 + iOffSet
        UserControl.ScaleHeight = Y2 - iOffSet
    End If
    
    UserControl.Refresh
End Sub

