VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIGen 
   BackColor       =   &H00FFFFFF&
   Caption         =   "AMC CodeAssist"
   ClientHeight    =   4590
   ClientLeft      =   1125
   ClientTop       =   1740
   ClientWidth     =   7035
   Icon            =   "MDIGen.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   30
      ScaleWidth      =   7035
      TabIndex        =   5
      Top             =   2670
      Width           =   7035
   End
   Begin VB.PictureBox picCodigo 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   7035
      TabIndex        =   4
      Top             =   2700
      Visible         =   0   'False
      Width           =   7035
   End
   Begin MSComctlLib.ImageList imgGen 
      Left            =   4080
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":2ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":5272
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":53CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":552A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":5686
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":59A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":5AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":5F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGen.frx":63A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4020
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbGen 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   4275
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicMenu 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   0
      ScaleHeight     =   2310
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox picSplit 
      Align           =   3  'Align Left
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2310
      Left            =   2895
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2310
      ScaleWidth      =   30
      TabIndex        =   0
      Top             =   360
      Width           =   30
   End
   Begin MSComctlLib.Toolbar tbEst 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgGen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Conectar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Desconectar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Guardar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Abrir libreria"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar libreria"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Abrir plantilla"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Propiedades"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar código"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Abrir asistente"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArcConectar 
         Caption         =   "&Conectar con SQL Server"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArcDesconectar 
         Caption         =   "&Desconectar de SQL Server"
      End
      Begin VB.Menu arcS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcGuardar 
         Caption         =   "&Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuArcGuardarComo 
         Caption         =   "G&uardar como..."
      End
      Begin VB.Menu arcS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuArcConfigPrint 
         Caption         =   "Configurar impresora"
      End
      Begin VB.Menu arcS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuEdiCortar 
         Caption         =   "Cor&tar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdiCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdiPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu ediS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdiBorrar 
         Caption         =   "Borr&ar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdiSelAll 
         Caption         =   "&Seleccionar todo"
         Shortcut        =   ^E
      End
      Begin VB.Menu ediS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdiBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEdiReemplazar 
         Caption         =   "Reempla&zar"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuLibreria 
      Caption         =   "&Libreria"
      Begin VB.Menu mnuLibAbrirLibreria 
         Caption         =   "&Abrir libreria"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLibCerrarLibreria 
         Caption         =   "&Cerrar Libreria"
      End
      Begin VB.Menu libS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibNuevaPlant 
         Caption         =   "&Nueva plantilla"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLibAgregaPlant 
         Caption         =   "&Agregar plantilla"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuLibRemovPlant 
         Caption         =   "&Remover plantilla"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuBarraHerr 
         Caption         =   "&Barra de herramientas"
         Begin VB.Menu mnuBarraEstandar 
            Caption         =   "Estándar"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuStatusBar 
            Caption         =   "Barra de estado"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu verS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerPropiedades 
         Caption         =   "&Propiedades"
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuHerGenCodigo 
         Caption         =   "&Generar código"
      End
      Begin VB.Menu herS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHerAssistCodigo 
         Caption         =   "&Asistente para generar código"
      End
      Begin VB.Menu herS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHerOpciones 
         Caption         =   "&Opciones"
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuVenOrizontal 
         Caption         =   "Alinear orizontal"
      End
      Begin VB.Menu mnuVenVertical 
         Caption         =   "Alinear vertical"
      End
      Begin VB.Menu mnuVenCascada 
         Caption         =   "Cascada"
      End
      Begin VB.Menu mnuVenOrganizar 
         Caption         =   "Organizar iconos"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu mnuAyuAcercaDe 
         Caption         =   "&Acerca de ..."
      End
   End
End
Attribute VB_Name = "MDIGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub MDIForm_Load()
   
   Call DeshabilitaMenus
   Load frmPlant
   frmPlant.Show
   
   Load frmBaseTemp
   frmBaseTemp.Show
   
End Sub


Private Sub MDIForm_Resize()
On Error Resume Next
   If Me.Height < 3000 Then Me.Height = 3000
   If Me.Width < 3000 Then Me.Width = 3000
   
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Unload frmBaseTemp
   Unload frmPlant
End Sub

Private Sub DeshabilitaMenus()
   mnuArcDesconectar.Enabled = False
   tbEst.Buttons(2).Enabled = False
   tbEst.Buttons(10).Enabled = False
End Sub

Private Sub mnuArcConectar_Click()
   Load frmServers
   frmServers.Show 1
End Sub

Private Sub mnuArcDesconectar_Click()
   On Error Resume Next
   objServer.DisConnect
   frmBaseTemp.tvServer.Nodes.Clear
   mnuArcConectar.Enabled = True
   mnuArcDesconectar.Enabled = False
   tbEst.Buttons(2).Enabled = False
   tbEst.Buttons(1).Enabled = True
End Sub

Private Sub mnuArcGuardar_Click()
   MsgBox "Guardar", vbInformation, App.Title
End Sub

Private Sub mnuArcSalir_Click()
   Call mnuArcDesconectar_Click
   Unload Me
   End
End Sub



Private Sub mnuAyuAcercaDe_Click()
   frmAbout.Show 1
End Sub

Private Sub mnuBarraEstandar_Click()
   If tbEst.Visible = True Then
      tbEst.Visible = False
      mnuBarraEstandar.Checked = False
   Else
      tbEst.Visible = True
      mnuBarraEstandar.Checked = True
   End If
End Sub

Private Sub mnuHerAssistCodigo_Click()
   MsgBox "El asistente esta aun en desarrollo", vbInformation, App.Title
End Sub

Private Sub mnuHerGenCodigo_Click()
Dim strBase As String, strTabla As String
Dim strPantilla As String
   
   With frmBaseTemp
      If .tvServer.Nodes.Count <= 0 Then MsgBox "No se ha establecido conexion con el servidor", vbExclamation, App.Title: Exit Sub
      
      If frmPlant.lvLibs.ListItems.Count <= 0 Then MsgBox "No se ha abierto ninguna libreria de pantillas", vbExclamation, App.Title: Exit Sub
      If frmPlant.lvLibs.SelectedItem Is Nothing Then MsgBox "Tiene que seleccionar una libreria", vbInformation, App.Title: Exit Sub
      If frmPlant.lvPlants.ListItems.Count <= 0 Then MsgBox "La libreria seleccionada no contiene pantillas", vbExclamation, App.Title: Exit Sub
      If frmPlant.lvPlants.SelectedItem Is Nothing Then MsgBox "Tiene que seleccionar una plantilla", vbInformation, App.Title: Exit Sub
      strPantilla = frmPlant.lvLibs.SelectedItem.SubItems(6) & frmPlant.lvPlants.SelectedItem.SubItems(2)
   End With
   
   With frmConsola
      If .Accion <> "TB" Then MsgBox "Seleccione por favor una base de datos y una tabla", vbExclamation, App.Title: Exit Sub
      If .lvConsola.ListItems.Count <= 0 Then MsgBox "No existen tablas en la lista", vbInformation, App.Title: Exit Sub
      If .lvConsola.SelectedItem Is Nothing Then MsgBox "Tiene que seleccionar una tabla", vbInformation, App.Title: Exit Sub
      strBase = .txtBase
      strTabla = .lvConsola.SelectedItem.SubItems(1)
   End With
   
   Screen.MousePointer = vbHourglass
   
   Dim objGenCode As New clsInterprete
   Dim objLineas As New clsLineas, objCampos As New clsCampos
   Dim objNewCode As New frmCodigo
   
   With objServer.Databases(strBase).Tables(strTabla).Columns
      For i = 1 To .Count
         objCampos.Add .Item(i).Name, .Item(i).Datatype, _
                       .Item(i).Length, .Item(i).NumericPrecision, _
                       .Item(i).NumericScale, .Item(i).InPrimaryKey
      Next i
      
   End With
   
   Load objNewCode
   
   With objNewCode
      If Trim$(frmPlant.lvLibs.SelectedItem.SubItems(4)) = "VB" Then
         .chlCodigo.DelimiterColor = &H8000&
         .chlCodigo.ForeColor = &H0&
         .chlCodigo.FunctionColor = &HFF00FF
         .chlCodigo.KeywordColor = &H9F9644
         .chlCodigo.OperatorColor = &H80FF&
         .chlCodigo.Language = hlVisualBasic
      End If
      
      If Trim$(frmPlant.lvLibs.SelectedItem.SubItems(4)) = "SQL" Then
         .chlCodigo.DelimiterColor = &HA58B70
         .chlCodigo.ForeColor = &H0&
         .chlCodigo.FunctionColor = &HFF00FF
         .chlCodigo.KeywordColor = &HC00000
         .chlCodigo.OperatorColor = &HFF&
         .chlCodigo.Language = [SQL Server]
      End If
   End With
   
   
   With objGenCode
      .Tabla = strTabla
      .Template = strPantilla
      .Campos = objCampos
      objNewCode.chlCodigo.Text = .GeneraCodigo
      Set objCampos = Nothing
      Screen.MousePointer = vbDefault
   End With
   
   objNewCode.Show
   
   
   
End Sub

Private Sub mnuHerOpciones_Click()
   MsgBox "Las opciones del sistema estan en desarrollo", vbInformation, App.Title
End Sub

Private Sub mnuLibAbrirLibreria_Click()
   With cdlg
      .FileName = ""
      .DialogTitle = "Abrir libreria"
      .Filter = "Libreria de plantillas (*.AMC)|*.AMC"
      .ShowOpen
      If Trim$(.FileName) = "" Then Exit Sub
      If AbrirLibreria(.FileName, Mid(.FileName, 1, Len(.FileName) - Len(.FileTitle))) = True Then
      End If
   End With
   
   
End Sub

Private Sub mnuLibCerrarLibreria_Click()
   If frmPlant.lvLibs.ListItems.Count <= 0 Then
      MsgBox "No se han cargado librerias", vbInformation, App.Title
      Exit Sub
   End If
   If frmPlant.lvLibs.SelectedItem Is Nothing Then
      MsgBox "Seleccione la libreria cerrar", vbInformation, App.Title
      Exit Sub
   End If
   
   frmPlant.lvLibs.ListItems.Remove frmPlant.lvLibs.SelectedItem.Index
   
   For i = 1 To frmPlant.lvPlants.ListItems.Count
      frmPlant.lvPlants.ListItems.Remove 1
   Next i
   
   strTipo = ""
   strFile = ""

End Sub



Private Sub mnuVerPropiedades_Click()
   On Error Resume Next
   Call frmConsola.lvConsola_DblClick
End Sub

Private Sub picCodigo_Resize()
   If frmPlant.bDocked Then
      frmPlant.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, _
                  picCodigo.ScaleWidth + (8 * Screen.TwipsPerPixelX), _
                  picCodigo.ScaleHeight + (8 * Screen.TwipsPerPixelY)
      Picture2.Top = frmPlant.Top - 10
   End If
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If PicMenu.Visible Then
        ReleaseCapture
        SendMessage PicMenu.hwnd, WM_NCLBUTTONDOWN, HTRIGHT, ByVal &O0
   End If
End Sub

Private Sub PicMenu_Resize()
On Error Resume Next
   
   If frmBaseTemp.bDocked Then
      frmBaseTemp.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, _
                  PicMenu.ScaleWidth + (8 * Screen.TwipsPerPixelX), _
                  PicMenu.ScaleHeight + (8 * Screen.TwipsPerPixelY)
   End If
   
End Sub


Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If picCodigo.Visible Then
      ReleaseCapture
      SendMessage picCodigo.hwnd, WM_NCLBUTTONDOWN, HTTOP, ByVal &O0
   Else
      Picture2.Top = frmPlant.Top - 10
   End If

End Sub

Private Sub tbEst_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1: Call mnuArcConectar_Click
      Case 2: Call mnuArcDesconectar_Click
      Case 6: Call mnuLibAbrirLibreria_Click
      Case 7: Call mnuLibCerrarLibreria_Click
      Case 10: Call mnuVerPropiedades_Click
      Case 12: Call mnuHerGenCodigo_Click
      Case 13: Call mnuHerAssistCodigo_Click
      Case 15: Call mnuArcSalir_Click
   End Select
End Sub


