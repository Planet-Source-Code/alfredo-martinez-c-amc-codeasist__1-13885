VERSION 5.00
Begin VB.Form frmServers 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Establecer conexion"
   ClientHeight    =   3150
   ClientLeft      =   2985
   ClientTop       =   3615
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin AMCCodeAssist.SBSPanel SBSPanel2 
      Height          =   2595
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   4577
      BackColor       =   14737632
      BorderStyle     =   4
      ScaleHeight     =   2583
      ScaleLeft       =   8
      ScaleTop        =   8
      ScaleWidth      =   4263
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   1
         Top             =   300
         Width           =   2295
      End
      Begin VB.OptionButton optOpciones 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usar Autentificacion de S&QL Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   1020
         Value           =   -1  'True
         Width           =   3435
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1740
         TabIndex        =   4
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1740
         TabIndex        =   6
         Top             =   1980
         Width           =   2295
      End
      Begin VB.Image picSvr 
         Height          =   480
         Left            =   180
         Picture         =   "frmConexion.frx":0000
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblFields 
         BackStyle       =   0  'Transparent
         Caption         =   "&Servidor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblFields 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   3
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label lblFields 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   5
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   240
         X2              =   4020
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   240
         X2              =   4020
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin AMCCodeAssist.SBSCoolButton cmdAceptar 
      Height          =   345
      Left            =   2040
      TabIndex        =   7
      Top             =   2730
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BackColor       =   12632256
      Caption         =   "&Aceptar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   12582912
      ForeColor       =   0
   End
   Begin VB.TextBox Accion 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   4020
      Visible         =   0   'False
      Width           =   375
   End
   Begin AMCCodeAssist.SBSCoolButton cmdCancelar 
      Height          =   345
      Left            =   3270
      TabIndex        =   8
      Top             =   2730
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BackColor       =   12632256
      Caption         =   "&Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   12582912
      ForeColor       =   0
   End
End
Attribute VB_Name = "frmServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdAceptar_Click(ByVal ClickReason As b2kClickReason)
On Error GoTo ErrorConexion
   
   If ValidaDatos = True Then
      Screen.MousePointer = vbHourglass
      objServer.Start True, Trim$(txtFields(0)), Trim$(txtFields(1)), Trim$(txtFields(2))
      Call ConstruyeSQLArbol(frmBaseTemp.tvServer)
      Screen.MousePointer = vbDefault
      Unload Me
      Load frmConsola
      frmConsola.Show
      MDIGen.mnuArcConectar.Enabled = False
      MDIGen.mnuArcDesconectar.Enabled = True
      MDIGen.tbEst.Buttons(2).Enabled = True
      MDIGen.tbEst.Buttons(1).Enabled = False
      MDIGen.tbEst.Buttons(10).Enabled = True
      Exit Sub
   End If
Exit Sub
ErrorConexion:
   MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click(ByVal ClickReason As b2kClickReason)
   Unload Me
End Sub


Private Sub Form_Load()
   objMPG.CentrarForma Me
   objMPG.Explosion Me.hwnd, 400, Negro
End Sub



Private Function ValidaDatos() As Boolean
   ValidaDatos = objMPG.CamposRequeridos(txtFields(0), txtFields(1))
End Function


Private Sub txtFields_GotFocus(Index As Integer)
   txtFields(Index).SelStart = 0
   txtFields(Index).SelLength = Len(txtFields(Index))
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      objMPG.Mayusculas KeyAscii
   End If
End Sub
