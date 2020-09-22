Attribute VB_Name = "VarProcAMCCodeAssist"
Option Explicit

Public objMPG As New clsMPG
Public objServer As New SQLDMO.SQLServer
Public i As Long, I2 As Long, I3 As Long
Public strFile As String
Public strTipo As String



'Procedimiento para mover contoles
Public Sub DragObject(objSourceMove As Object)
On Local Error Resume Next
   Call ReleaseCapture
   Call SendMessage(objSourceMove.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub


Public Function ConstruyeSQLArbol(ByVal objTreeView As TreeView) As Boolean
Dim xNode As Node
   'Limpia el arbol
   objTreeView.Nodes.Clear
   'Carga el nodo principal
   Set xNode = objTreeView.Nodes.Add(, , "Servers", "Microsoft SQL Servers", Image:=1, SelectedImage:=1)
   xNode.Expanded = True
   Set xNode = objTreeView.Nodes.Add("Servers", tvwChild, objServer.Name, objServer.Name, Image:=2, SelectedImage:=2)
   xNode.Expanded = True
   Set xNode = objTreeView.Nodes.Add(objServer.Name, tvwChild, "DB", "Bases de datos", Image:=4, SelectedImage:=4)
   xNode.Expanded = True
   
   With objServer
      For i = 1 To .Databases.Count
         Set xNode = objTreeView.Nodes.Add("DB", tvwChild, .Databases(i).Name, .Databases(i).Name, Image:=13, SelectedImage:=13)
         'xNode.Expanded = True
         Set xNode = objTreeView.Nodes.Add(.Databases(i).Name, tvwChild, "TB" & .Databases(i).Name, "Tablas", Image:=8, SelectedImage:=8)
         Set xNode = objTreeView.Nodes.Add(.Databases(i).Name, tvwChild, "VI" & .Databases(i).Name, "Vistas", Image:=9, SelectedImage:=9)
         Set xNode = objTreeView.Nodes.Add(.Databases(i).Name, tvwChild, "SP" & .Databases(i).Name, "Procedimientos", Image:=10, SelectedImage:=10)
      Next i
   End With

End Function



'Funcion para abrir una libreria de plantillas
Public Function AbrirLibreria(ByVal strArchivo As String, ByVal strPath As String) As Boolean
Dim xNode As Node, intNumPlt As Integer
Dim strResult As String, strResult2 As String
   
   strResult = objMPG.ReadINI(strArchivo, "Configuracion", "Titulo")
   
   With frmPlant.lvLibs.ListItems
      For i = 1 To .Count
         If Trim$(strResult) = Trim$(.Item(i).SubItems(1)) Then
            MsgBox "La libreria ya se encuentra abierta", vbExclamation, App.Title
            AbrirLibreria = False
            Exit Function
         End If
      Next i
   End With
   
   Dim intItemLib As Integer
   With frmPlant.lvLibs.ListItems
      .Add
      intItemLib = .Count
      .Item(intItemLib).SubItems(1) = objMPG.ReadINI(strArchivo, "Configuracion", "Titulo")
      .Item(intItemLib).SubItems(2) = objMPG.ReadINI(strArchivo, "Configuracion", "Autor")
      .Item(intItemLib).SubItems(3) = objMPG.ReadINI(strArchivo, "Configuracion", "Fecha")
      .Item(intItemLib).SubItems(4) = objMPG.ReadINI(strArchivo, "Configuracion", "Tipo")
      .Item(intItemLib).SubItems(5) = objMPG.ReadINI(strArchivo, "Configuracion", "Plantillas")
      
      .Item(intItemLib).SubItems(6) = strPath
      
      .Item(intItemLib).SubItems(7) = strArchivo
   End With
   
   AbrirLibreria = True
End Function


Public Function ExtraPlantillas(ByVal strArchivo As String, ByVal intNumPlant As Integer) As Boolean
Dim strResult As String, strResult2 As String
   With frmPlant.lvPlants.ListItems
      .Clear
      
      For i = 1 To intNumPlant
         strResult = objMPG.ReadINI(strArchivo, "PLT" & Format(i, "00"), "Titulo")
         strResult2 = objMPG.ReadINI(strArchivo, "PLT" & Format(i, "00"), "Nombre")
         .Add
         .Item(i).SubItems(1) = strResult
         .Item(i).SubItems(2) = strResult2
      Next i
      
   End With


   ExtraPlantillas = True

End Function




'Procedimiento para abrir una plantilla
Public Sub AbrePlantilla(ByVal strRuta As String, ByVal strArchivo As String, ByVal strTitulo As String)
Dim strResult As String, varResult, strLinea As String
Dim strPlant As String
On Error GoTo ErrorAbrePlantilla
   
   varResult = Dir(strRuta & strArchivo)
   If varResult <> "" Then
      frmPlant.txtPlantilla.Text = ""
      Open strFile & strArchivo For Input As #1
      While Not EOF(1)
          Line Input #1, strLinea
          frmPlant.txtPlantilla.Text = frmPlant.txtPlantilla.Text & strLinea & vbCrLf
      Wend
      Close #1
      frmPlant.FormDragger1.Caption = "Plantilla - " & Trim$(strTitulo)
   Else
      frmPlant.txtPlantilla.Text = ""
   End If
Exit Sub
ErrorAbrePlantilla:
   MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
End Sub

