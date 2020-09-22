VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variables miembro
Private m_strNombre As String
Private m_strTipo As String
Private m_intLongitud As Integer
Private m_intPresicion As Integer
Private m_intEscala As Integer
Private m_boolPK As Boolean


'Propiedades y Metodos
Public Property Get Nombre() As String
   Nombre = m_strNombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
   m_strNombre = vNewValue
End Property

Public Property Get Tipo() As String
   Tipo = m_strTipo
End Property

Public Property Let Tipo(ByVal vNewValue As String)
   m_strTipo = vNewValue
End Property

Public Property Get Longitud() As Integer
   Longitud = m_intLongitud
End Property

Public Property Let Longitud(ByVal vNewValue As Integer)
   m_intLongitud = vNewValue
End Property

Public Property Get Presicion() As Integer
   Presicion = m_intPresicion
End Property

Public Property Let Presicion(ByVal vNewValue As Integer)
   m_intPresicion = vNewValue
End Property

Public Property Get Escala() As Integer
   Escala = m_intEscala
End Property

Public Property Let Escala(ByVal vNewValue As Integer)
   m_intEscala = vNewValue
End Property

Public Property Get PK() As Boolean
   PK = m_boolPK
End Property

Public Property Let PK(ByVal vNewValue As Boolean)
   m_boolPK = vNewValue
End Property
