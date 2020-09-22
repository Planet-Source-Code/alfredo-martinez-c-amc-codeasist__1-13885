VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colLineas As Collection
Private I As Long, I2 As Long


Public Sub Add(ByVal strTexto As String)
Dim objLinea As New clsLinea
   objLinea.Texto = strTexto
   With m_colLineas
      .Add objLinea
   End With
   Set objLinea = Nothing
End Sub


Public Sub Clear()
   With m_colLineas
      For I = 1 To .Count
         m_colLineas.Remove (1)
      Next I
   End With
End Sub

Public Function Item(ByVal Index As Long) As clsLinea
   Set Item = m_colLineas(Index)
End Function


Public Function Count() As Long
   Count = m_colLineas.Count
End Function


Private Sub Class_Initialize()
   Set m_colLineas = New Collection
End Sub


Private Sub Class_Terminate()
   Set m_colLineas = Nothing
End Sub
