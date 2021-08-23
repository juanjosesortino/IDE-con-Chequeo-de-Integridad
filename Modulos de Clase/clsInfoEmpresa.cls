VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfoEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCodigoEmpresa As String
Private mvarNombreEmpresa As String
Private mvarEjercicioVigente As String
Private mvarSucursalActiva As String
Public Property Let SucursalActiva(ByVal vData As String)
    mvarSucursalActiva = vData
End Property
Public Property Get SucursalActiva() As String
    SucursalActiva = mvarSucursalActiva
End Property
Public Property Let EjercicioVigente(ByVal vData As String)
    mvarEjercicioVigente = vData
End Property
Public Property Get EjercicioVigente() As String
    EjercicioVigente = mvarEjercicioVigente
End Property
Public Property Let NombreEmpresa(ByVal vData As String)
    mvarNombreEmpresa = vData
End Property
Public Property Get NombreEmpresa() As String
    NombreEmpresa = mvarNombreEmpresa
End Property
Public Property Let CodigoEmpresa(ByVal vData As String)
    mvarCodigoEmpresa = vData
End Property
Public Property Get CodigoEmpresa() As String
Attribute CodigoEmpresa.VB_UserMemId = 0
    CodigoEmpresa = mvarCodigoEmpresa
End Property



