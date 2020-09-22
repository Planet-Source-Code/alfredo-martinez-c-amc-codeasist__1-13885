VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVTable 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3210
   ClientLeft      =   2385
   ClientTop       =   3765
   ClientWidth     =   6105
   Icon            =   "frmVTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgTablas 
      Left            =   1080
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVTable.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvTabla 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgTablas"
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmVTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
   objMPG.CentrarForma Me
End Sub





'Procedimiento para cargar los datos de una tabla
Public Sub CargaTabla(ByVal strBase As String, ByVal strTable As String)
   If strBase = "DB" Then Exit Sub
   Call ConstruyeLista
   With objServer.Databases(strBase).Tables(strTable).Columns
      For i = 1 To .Count
         lvTabla.ListItems.Add
         If .Item(i).InPrimaryKey = True Then
            lvTabla.ListItems(i).SmallIcon = 1
         End If
         lvTabla.ListItems(i).SubItems(1) = .Item(i).id
         lvTabla.ListItems(i).SubItems(2) = .Item(i).Name
         lvTabla.ListItems(i).SubItems(3) = .Item(i).Datatype
         lvTabla.ListItems(i).SubItems(4) = .Item(i).Length
         If .Item(i).AllowNulls = True Then
            lvTabla.ListItems(i).SubItems(5) = "S"
         Else
            lvTabla.ListItems(i).SubItems(5) = "N"
         End If
         lvTabla.ListItems(i).SubItems(6) = .Item(i).Default
      Next i
   End With
End Sub


'Procedimiento para construir la lista
Private Sub ConstruyeLista()
   With lvTabla.ColumnHeaders
      .Add , , "PK", 400
      .Add , , "ID", 400
      .Add , , "Nombre", 2000
      .Add , , "Tipo", 1200
      .Add , , "Tamaño", 800
      .Add , , "Nulos", 800
      .Add , , "Default", 3000
   End With
End Sub
