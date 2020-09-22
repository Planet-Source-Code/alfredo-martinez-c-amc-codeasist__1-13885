VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsola 
   Caption         =   "Consola"
   ClientHeight    =   3690
   ClientLeft      =   6660
   ClientTop       =   3975
   ClientWidth     =   4635
   ControlBox      =   0   'False
   Icon            =   "frmConsola.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   4635
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBase 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Accion 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList imgConsola 
      Left            =   3780
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsola.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsola.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsola.frx":0B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsola.frx":0E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsola.frx":0FBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvConsola 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgConsola"
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmConsola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum eTipoLista
   Base
   Tablas
   Vistas
   Procedimientos
   Usuarios
End Enum

Private Sub Form_Resize()
   lvConsola.Width = Me.Width - 100
   lvConsola.Height = Me.Height - 400
End Sub



Public Sub MuestraTablas(ByVal strBase As String)
   Call ContruyeListas(Tablas)
   Accion = "TB"
   txtBase = strBase
   With objServer.Databases(strBase).Tables
      For i = 1 To .Count
         lvConsola.ListItems.Add , , , , SmallIcon:=1
         lvConsola.ListItems(i).SubItems(1) = .Item(i).Name
         lvConsola.ListItems(i).SubItems(2) = .Item(i).Owner
         If .Item(i).SystemObject = True Then
            lvConsola.ListItems(i).SubItems(3) = "Sistema"
         Else
            lvConsola.ListItems(i).SubItems(3) = "Usuario"
         End If
         lvConsola.ListItems(i).SubItems(4) = .Item(i).CreateDate
      Next i
   End With
End Sub


Public Sub MuestraVistas(ByVal strBase As String)
   Call ContruyeListas(Tablas)
   Accion = "VI"
   txtBase = strBase
   With objServer.Databases(strBase).Views
      For i = 1 To .Count
         lvConsola.ListItems.Add , , , , SmallIcon:=2
         lvConsola.ListItems(i).SubItems(1) = .Item(i).Name
         lvConsola.ListItems(i).SubItems(2) = .Item(i).Owner
         If .Item(i).SystemObject = True Then
            lvConsola.ListItems(i).SubItems(3) = "Sistema"
         Else
            lvConsola.ListItems(i).SubItems(3) = "Usuario"
         End If
         lvConsola.ListItems(i).SubItems(4) = .Item(i).CreateDate
      Next i
   End With
End Sub


Public Sub MuestraProcedimientos(ByVal strBase As String)
   Call ContruyeListas(Tablas)
   Accion = "SP"
   txtBase = strBase
   With objServer.Databases(strBase).StoredProcedures
      For i = 1 To .Count
         lvConsola.ListItems.Add , , , , SmallIcon:=3
         lvConsola.ListItems(i).SubItems(1) = .Item(i).Name
         lvConsola.ListItems(i).SubItems(2) = .Item(i).Owner
         If .Item(i).SystemObject = True Then
            lvConsola.ListItems(i).SubItems(3) = "Sistema"
         Else
            lvConsola.ListItems(i).SubItems(3) = "Usuario"
         End If
         lvConsola.ListItems(i).SubItems(4) = .Item(i).CreateDate
      Next i
   End With
End Sub


Public Sub MuestraBase(ByVal strBase As String)
Dim lngTamañoDB As Long, datFechaCreacion As Variant
Dim lngUsadoDB As Long, lngLibreDB As Long, strID As String
Dim lngUsadoIDX As Long, intStatus As Integer, strVersion As String
   
   Call ContruyeListas(Base)
   txtBase = strBase
   With objServer.Databases(strBase)
      lngTamañoDB = .Size
      lngLibreDB = .SpaceAvailableInMB
      datFechaCreacion = .CreateDate
      lngUsadoDB = .DataSpaceUsage
      strID = .ID
      lngUsadoIDX = .IndexSpaceUsage
      intStatus = .Status
      strVersion = .Version
   End With
   
   With Me.lvConsola.ListItems
      For i = 1 To 9
         .Add , , , , SmallIcon:=5
      Next i
      .Item(1).SubItems(1) = "Identificador ": .Item(1).SubItems(2) = strID
      .Item(2).SubItems(1) = "Fecha de creacion ": .Item(2).SubItems(2) = datFechaCreacion
      .Item(3).SubItems(1) = "Nombre ": .Item(3).SubItems(2) = "strBase"
      .Item(4).SubItems(1) = "Tamaño ": .Item(4).SubItems(2) = lngTamañoDB & " MB"
      .Item(5).SubItems(1) = "Espacio usado ": .Item(5).SubItems(2) = lngUsadoDB / 1000 & " MB"
      .Item(6).SubItems(1) = "Espacio libre ": .Item(6).SubItems(2) = lngLibreDB & " MB"
      .Item(7).SubItems(1) = "Espacio usado por Idices": .Item(7).SubItems(2) = lngUsadoIDX / 1000 & " MB"
      .Item(8).SubItems(1) = "Version ": .Item(8).SubItems(2) = strVersion
      .Item(9).SubItems(1) = "Estatus ": .Item(9).SubItems(2) = intStatus
   End With
   
End Sub



Public Sub MuestraUsuarios(ByVal strBase As String)
   Call ContruyeListas(Usuarios)
   Accion = "TB"
   txtBase = strBase
   With objServer.Databases(strBase).Users
      For i = 1 To .Count
         lvConsola.ListItems.Add , , , , SmallIcon:=4
         lvConsola.ListItems(i).SubItems(1) = .Item(i).Name
         lvConsola.ListItems(i).SubItems(2) = .Item(i).Login
         If .Item(i).HasDBAccess = True Then
            lvConsola.ListItems(i).SubItems(3) = "Permitido"
         Else
            lvConsola.ListItems(i).SubItems(3) = "No Permitido"
         End If
      Next i
   End With
End Sub





Private Sub ContruyeListas(ByVal enmTipo As eTipoLista)
   With lvConsola
      .ListItems.Clear
      .ColumnHeaders.Clear
      
      If enmTipo = Base Then
         .ColumnHeaders.Add , , "", 300
         .ColumnHeaders.Add , , "Propiedades", 3000
         .ColumnHeaders.Add , , "Valores", 3000
      End If
      
      If enmTipo = Tablas Or enmTipo = Vistas Or enmTipo = Procedimientos Then
         .ColumnHeaders.Add , , "", 300
         .ColumnHeaders.Add , "Nombre", "Nombre", 3000
         .ColumnHeaders.Add , "Propietario", "Propietario", 1000
         .ColumnHeaders.Add , "Tipo", "Tipo", 1000
         .ColumnHeaders.Add , "Fecha de Creacion", "Fecha de Creacion", 3000
      End If
      
      If enmTipo = Usuarios Then
         .ColumnHeaders.Add , , "", 300
         .ColumnHeaders.Add , "Nombre", "Nombre"
         .ColumnHeaders.Add , "Usuario", "Usuario"
         .ColumnHeaders.Add , "Acceso", "Acceso"
      End If
      
   End With
   
   
End Sub




Public Sub lvConsola_DblClick()
Dim strNombreObjeto As String, strBase As String
   If Accion = "DB" Then Exit Sub
   If Accion = "SP" Or Accion = "VI" Then
      If lvConsola.ListItems.Count <= 0 Then Exit Sub
      If lvConsola.SelectedItem Is Nothing Then Exit Sub
      strNombreObjeto = lvConsola.SelectedItem.SubItems(1)
      strBase = txtBase.Text
      Screen.MousePointer = vbHourglass
      Load frmSPView
      If Accion = "SP" Then
         frmSPView.chlCodigo.Text = objServer.Databases(strBase).StoredProcedures(strNombreObjeto).Text
         frmSPView.Caption = "Propiedades del Procedimiento Almacenado - " & lvConsola.SelectedItem.SubItems(1)
         frmSPView.txtBase = txtBase
         frmSPView.txtNombre = lvConsola.SelectedItem.SubItems(1)
         frmSPView.Accion = "C"
         frmSPView.txtTipo = Accion
      Else
         frmSPView.chlCodigo.Text = objServer.Databases(strBase).Views(strNombreObjeto).Text
         frmSPView.Caption = "Propiedades de la Vista - " & lvConsola.SelectedItem.SubItems(1)
         frmSPView.txtBase = txtBase
         frmSPView.txtNombre = lvConsola.SelectedItem.SubItems(1)
         frmSPView.Accion = "C"
         frmSPView.txtTipo = Accion
      End If
      Screen.MousePointer = vbDefault
      frmSPView.Show
      
   End If
   
   If Accion = "TB" Then
      If lvConsola.ListItems.Count <= 0 Then Exit Sub
      If lvConsola.SelectedItem Is Nothing Then Exit Sub
      Load frmVTable
      frmVTable.Caption = "Propiedades de la Tabla - " & lvConsola.SelectedItem.SubItems(1)
      Call frmVTable.CargaTabla(txtBase, lvConsola.SelectedItem.SubItems(1))
      Screen.MousePointer = vbDefault
      frmVTable.Show 1
   End If
   
End Sub

