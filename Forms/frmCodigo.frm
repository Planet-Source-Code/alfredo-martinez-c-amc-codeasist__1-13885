VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCodigo 
   Caption         =   "Código Generado SQL"
   ClientHeight    =   1605
   ClientLeft      =   2535
   ClientTop       =   1860
   ClientWidth     =   3255
   Icon            =   "frmCodigo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   3255
   WindowState     =   2  'Maximized
   Begin AMCCodeAssist.CodeHighlight chlCodigo 
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1614
      Language        =   4
      KeywordColor    =   12582912
      OperatorColor   =   10849136
      DelimiterColor  =   32768
      ForeColor       =   0
      FunctionColor   =   8421631
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
   Begin MSComctlLib.ImageList imgEdi 
      Left            =   0
      Top             =   1560
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
            Picture         =   "frmCodigo.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodigo.frx":05AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodigo.frx":0712
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodigo.frx":087A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodigo.frx":09E2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Resize()
   On Error Resume Next
   chlCodigo.Width = Me.Width - 100
   chlCodigo.Height = Me.Height - 400
End Sub

