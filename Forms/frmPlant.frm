VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlant 
   ClientHeight    =   1620
   ClientLeft      =   1755
   ClientTop       =   6420
   ClientWidth     =   5445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   5445
   Begin MSComctlLib.ListView lvPlants 
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1085
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvLibs 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1085
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgPics 
      Left            =   4080
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlant.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlant.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlant.frx":08A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPlantilla 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   660
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   795
   End
   Begin MSComctlLib.TabStrip tsPlant 
      Height          =   1155
      Left            =   -60
      TabIndex        =   1
      Top             =   0
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   2037
      HotTracking     =   -1  'True
      ImageList       =   "imgPics"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Librerias"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Plantillas"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Propiedades"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin AMCCodeAssist.FormDragger FormDragger1 
      Align           =   1  'Align Top
      Height          =   165
      Left            =   0
      Top             =   0
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   291
   End
End
Attribute VB_Name = "frmPlant"
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
   Call ConstruyeListas
   'Initialize the positions/sizes of this form
   lDockedWidth = MDIGen.picCodigo.ScaleWidth + (8 * Screen.TwipsPerPixelX)
   lDockedHeight = MDIGen.picCodigo.ScaleHeight + (8 * Screen.TwipsPerPixelY)
   lFloatingLeft = Me.Left
   lFloatingTop = Me.Top
   lFloatingWidth = Me.Width
   lFloatingHeight = Me.Height
   'Start with the form docked in Picture1 on the MDI Form
   'put Form1 in the 'Dock' and position it so its resizing border is
   'hidden outside the confines of Picture1
   bDocked = True
   SetParent Me.hwnd, MDIGen!picCodigo.hwnd
   Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
   
   MDIGen!picCodigo.Visible = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'reset this form's owner to prevent a crash
    Call SetWindowWord(Me.hwnd, SWW_HPARENT, 0&)
End Sub

Private Sub Form_Resize()
On Error Resume Next
   
   tsPlant.Height = Me.Height - 50
   tsPlant.Width = Me.Width - 50
   
   lvLibs.Left = tsPlant.Left + 50
   lvLibs.Height = Me.Height - 500
   lvLibs.Width = Me.Width - 150
   
   lvPlants.Left = tsPlant.Left + 50
   lvPlants.Height = Me.Height - 500
   lvPlants.Width = Me.Width - 150
   
   txtPlantilla.Left = tsPlant.Left + 50
   txtPlantilla.Height = Me.Height - 500
   txtPlantilla.Width = Me.Width - 150
   
   
   If Me.WindowState <> vbMinimized Then
       'Update the stored Values
       StoreFormDimensions
   End If
   txtPlantilla.Refresh
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
        MDIGen!picCodigo.Visible = False
        Me.Visible = True
        'make this form 'float' above the MDI form
        Call SetWindowWord(Me.hwnd, SWW_HPARENT, MDIGen.hwnd)
    Else
        'Dock
        bDocked = True
        SetParent Me.hwnd, MDIGen!picCodigo.hwnd
        Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        MDIGen!picCodigo.Visible = True
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
    GetWindowRect MDIGen!picCodigo.hwnd, rct
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
        SetParent Me.hwnd, MDIGen!picCodigo.hwnd
        Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        MDIGen!picCodigo.Visible = True
    Else
        Me.Visible = False
        bDocked = False
        SetParent Me.hwnd, 0
        Me.Move FormLeft * Screen.TwipsPerPixelX, FormTop * Screen.TwipsPerPixelY, lFloatingWidth, lFloatingHeight
        MDIGen!picCodigo.Visible = False
        Me.Visible = True
        'make this form 'float' above the MDI form
        Call SetWindowWord(Me.hwnd, SWW_HPARENT, MDIGen.hwnd)
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
    GetWindowRect MDIGen!picCodigo.hwnd, rct
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


'Procedimiento para construir las listas
Private Sub ConstruyeListas()
   With lvLibs.ColumnHeaders
      .Add , , , 300
      .Add , , "Título", 2500
      .Add , , "Autor", 2500
      .Add , , "Fecha", 800
      .Add , , "Lenguaje", 900
      .Add , , "Plantillas", 800
      .Add , , "Ruta", 3200
      .Add , , "Libreria", 0
   End With
   
   With lvPlants.ColumnHeaders
      .Add , , , 300
      .Add , , "Nombre", 5000
      .Add , , "Archivo", 4000
   End With
   
   
End Sub



Private Sub lvLibs_ItemClick(ByVal Item As MSComctlLib.ListItem)
   With lvLibs.SelectedItem
      Call ExtraPlantillas(.SubItems(7), val(.SubItems(5)))
   End With
End Sub

Private Sub lvPlants_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Call AbrePlantilla(lvLibs.SelectedItem.SubItems(6), lvPlants.SelectedItem.SubItems(2), lvPlants.SelectedItem.SubItems(1))
End Sub

Private Sub tsPlant_Click()
   Select Case tsPlant.SelectedItem.Index
      Case 1
         lvLibs.Visible = True
         lvPlants.Visible = False
         txtPlantilla.Visible = False
      Case 2
         lvLibs.Visible = False
         lvPlants.Visible = True
         txtPlantilla.Visible = False
      Case 3
         lvLibs.Visible = False
         lvPlants.Visible = False
         txtPlantilla.Visible = True
   End Select
End Sub


