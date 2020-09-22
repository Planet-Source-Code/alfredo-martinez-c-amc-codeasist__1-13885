VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBaseTemp 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3405
   ClientLeft      =   3585
   ClientTop       =   3210
   ClientWidth     =   2835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBaseTemp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   2835
   Begin MSComctlLib.TreeView tvServer 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1720
      _Version        =   393217
      Indentation     =   9
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgPrincipal"
      Appearance      =   1
   End
   Begin AMCCodeAssist.FormDragger FormDragger1 
      Align           =   1  'Align Top
      Height          =   165
      Left            =   0
      Top             =   0
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   291
   End
   Begin MSComctlLib.ImageList imgPrincipal 
      Left            =   1860
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16744703
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":0328
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":077C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":0A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":31F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":3644
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":3A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":3C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":4194
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":44B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":460C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":4B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":50B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseTemp.frx":5508
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBaseTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public variables used elsewhere to set values for this form's position
'and size.
Dim lFloatingWidth As Long
Dim lFloatingHeight As Long
Dim lFloatingLeft As Long
Dim lFloatingTop As Long
Dim bMoving As Boolean

'Private variables used to track moving/sizing etc.
Public bDocked As Boolean
Public lDockedWidth As Long
Public lDockedHeight As Long

Private Sub Form_Load()
    'Initialize the positions/sizes of this form
    lDockedWidth = MDIGen.PicMenu.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = MDIGen.PicMenu.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    lFloatingLeft = Me.Left
    lFloatingTop = Me.Top
    lFloatingWidth = Me.Width
    lFloatingHeight = Me.Height
    'Start with the form docked in Picture1 on the MDI Form
    'put Form1 in the 'Dock' and position it so its resizing border is
    'hidden outside the confines of Picture1
    bDocked = True
    SetParent Me.hwnd, MDIGen!PicMenu.hwnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    
    MDIGen!PicMenu.Visible = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'reset this form's owner to prevent a crash
    Call SetWindowWord(Me.hwnd, SWW_HPARENT, 0&)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        'Update the stored Values
        StoreFormDimensions
        'position and size the listbox
        tvServer.Move 3 * Screen.TwipsPerPixelX, FormDragger1.Height + (3 * Screen.TwipsPerPixelY), Me.ScaleWidth - (7 * Screen.TwipsPerPixelX), Me.ScaleHeight - (FormDragger1.Height + (6 * Screen.TwipsPerPixelY))
    End If
    tvServer.Refresh
    
End Sub

Private Sub FormDragger1_DblClick()

    'Snap the form in or out of the dock (Picture1)
    bMoving = True 'stop the new dimensions being stored
    If bDocked Then
        'Undock
        Me.Visible = False
        bDocked = False
        SetParent Me.hwnd, 0
        Me.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
        MDIGen!PicMenu.Visible = False
        Me.Visible = True
        'make this form 'float' above the MDI form
        Call SetWindowWord(Me.hwnd, SWW_HPARENT, MDIGen.hwnd)
    Else
        'Dock
        bDocked = True
        SetParent Me.hwnd, MDIGen!PicMenu.hwnd
        Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        MDIGen!PicMenu.Visible = True
        MDIGen!Picture2.Top = frmPlant.Top - 10
    End If
    bMoving = False

End Sub

Private Sub FormDragger1_FormDropped(FormLeft As Long, FormTop As Long, formWidth As Long, formHeight As Long)
    
    Dim rct As RECT

    'If over Picture1 on MDIForm1 which we are using as a Dock, set parent
    'of this form to Picture1, and position it at -4,-4 pixels, otherwise
    'set this Form's parent to the desktop and postion it at Left,Top
    'We dont need to size the form, as the DragForm control will have done
    'this for us.
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    'Get the screen based coordinates of Picture1
    GetWindowRect MDIGen!PicMenu.hwnd, rct
    'Inflate the rect because we want the form to be bigger than Picture1
    'to hide it's border
    With rct
        .Left = .Left - 4
        .Top = .Top - 4
        .Right = .Right + 4
        .Bottom = .Bottom + 4
    End With
    'See if the top/left corner of this form is in Picture1's screen rectangle
    'As we have set RepositionForm to false, we are responsible for positioning the form
    If PtInRect(rct, FormLeft, FormTop) Then
        bDocked = True
        SetParent Me.hwnd, MDIGen!PicMenu.hwnd
        Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        MDIGen!PicMenu.Visible = True
        MDIGen!Picture2.Top = frmPlant.Top - 10
    Else
        Me.Visible = False
        bDocked = False
        SetParent Me.hwnd, 0
        Me.Move FormLeft * Screen.TwipsPerPixelX, FormTop * Screen.TwipsPerPixelY, lFloatingWidth, lFloatingHeight
        MDIGen!PicMenu.Visible = False
        Me.Visible = True
        'make this form 'float' above the MDI form
        Call SetWindowWord(Me.hwnd, SWW_HPARENT, MDIGen.hwnd)
        MDIGen!Picture2.Top = frmPlant.Top - 10
    End If
    
    'reset the moving flag and store the form dimensions
    bMoving = False
    StoreFormDimensions

End Sub

Private Sub FormDragger1_FormMoved(FormLeft As Long, FormTop As Long, formWidth As Long, formHeight As Long)
    
    Dim rct As RECT
    
    'Set the moving flag so we dont store the wrong dimensions
    bMoving = True
    
    'If over Picture1 on MDIForm1 which we are using as a Dock, change the width to that of
    'Picture1, else change it to the 'floating width and height
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    'Get the screen based coordinates of Picture1
    GetWindowRect MDIGen!PicMenu.hwnd, rct
    'Inflate the rect because we want the form to be bigger than Picture1
    'to hide it's border
    With rct
        .Left = .Left - 4
        .Top = .Top - 4
        .Right = .Right + 4
        .Bottom = .Bottom + 4
    End With
    'See if the top/left corner of this form is in Picture1's screen rectangle
    
    If PtInRect(rct, FormLeft, FormTop) Then
        formWidth = lDockedWidth / Screen.TwipsPerPixelX
        formHeight = lDockedHeight / Screen.TwipsPerPixelY
    Else
        formWidth = lFloatingWidth / Screen.TwipsPerPixelX
        formHeight = lFloatingHeight / Screen.TwipsPerPixelY
    End If

End Sub

Private Sub StoreFormDimensions()

   'Store the height/width values
    If Not bMoving Then
        If bDocked Then
            lDockedWidth = Me.Width
            lDockedHeight = Me.Height
        Else
            lFloatingLeft = Me.Left
            lFloatingTop = Me.Top
            lFloatingWidth = Me.Width
            lFloatingHeight = Me.Height
        End If
    End If
End Sub




Private Sub tvServer_Collapse(ByVal Node As MSComctlLib.Node)
   If Node.Key = "DB" Then
      Node.SelectedImage = 3
   End If

End Sub

Private Sub tvServer_Expand(ByVal Node As MSComctlLib.Node)
   If Node.Key = "DB" Then
      Node.SelectedImage = 4
   End If

End Sub

Private Sub tvServer_NodeClick(ByVal Node As MSComctlLib.Node)
   
   frmConsola.txtBase = ""
   Screen.MousePointer = vbHourglass
   Select Case Mid(Node.Key, 1, 2)
      Case "TB"
         Call frmConsola.MuestraTablas(Mid(Node.Key, 3, Len(Node.Key)))
      Case "SP"
         Call frmConsola.MuestraProcedimientos(Mid(Node.Key, 3, Len(Node.Key)))
      Case "VI"
         Call frmConsola.MuestraVistas(Mid(Node.Key, 3, Len(Node.Key)))
      Case "US"
         Call frmConsola.MuestraUsuarios(Mid(Node.Key, 3, Len(Node.Key)))
   End Select
   
   If Node.Parent Is Nothing Then
   Else
      If Node.Parent.Key = "DB" Then Call frmConsola.MuestraBase(Node.Text): frmConsola.Accion = "DB"
   End If
   Screen.MousePointer = vbDefault
End Sub





