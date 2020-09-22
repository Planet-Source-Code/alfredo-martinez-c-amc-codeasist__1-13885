VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSPView 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5715
   ClientLeft      =   3030
   ClientTop       =   2010
   ClientWidth     =   6585
   Icon            =   "frmSPView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin AMCCodeAssist.CodeHighlight chlCodigo 
      Height          =   2955
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   5212
      BackColor       =   16777215
      KeywordColor    =   12582912
      OperatorColor   =   255
      DelimiterColor  =   10849136
      ForeColor       =   0
      FunctionColor   =   12583104
      HighlightCode   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTipo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   3060
      TabIndex        =   3
      Top             =   4260
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Accion 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   4260
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   4260
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtBase 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   4260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgConsola 
      Left            =   5940
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPView.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPView.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPView.frx":0B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPView.frx":0E56
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSPView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Form_Load()
   chlCodigo.LineIndex = 1
   chlCodigo.HighlightCode = hlOnNewLine
   chlCodigo.Language = [SQL Server]
End Sub

Private Sub Form_Resize()
   chlCodigo.Width = Me.Width - 100
   chlCodigo.Height = Me.Height - 400
End Sub


Private Function Modifica() As Boolean
Dim strBase As String, strNombreObjeto As String
Dim objStoreProc As New SQLDMO.StoredProcedure
Dim objVista As New SQLDMO.View
On Error GoTo ErrorModifica
   strBase = txtBase
   strNombreObjeto = txtNombre
   Screen.MousePointer = vbHourglass
   With objServer.Databases(strBase)
      If txtTipo = "SP" Then
         Set objStoreProc = .StoredProcedures(strNombreObjeto)
         objStoreProc.Alter chlCodigo.Text
      Else
         Set objVista = .Views(strNombreObjeto)
         objVista.Alter chlCodigo.Text
      End If
   End With
   Screen.MousePointer = vbDefault
   Modifica = True
Exit Function
ErrorModifica:
   MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
   Screen.MousePointer = vbDefault
End Function



'Procedimiento para generar un nuevo store procedure
Private Function Nuevo() As Boolean
Dim objSP As New SQLDMO.StoredProcedure
Dim objView As New SQLDMO.View
   'objSP.Name
End Function


