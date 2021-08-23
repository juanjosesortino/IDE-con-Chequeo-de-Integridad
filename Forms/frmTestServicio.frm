VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTestServicio 
   Caption         =   "Form1"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   2175
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTestServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Connected As Boolean

Private Sub Form_Load()
Dim StrConexionSevicio As String
Dim strServerRoot As String

   Connected = True
   ' Si estamos en Algoritmo no vamos a exit sub
   If GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultDomainName") = "ALGORITMO" Then Exit Sub
   
   StrConexionSevicio = GetSetting("Algoritmo", "Inicio", "ConexionSevicio", "Nada")
   If StrConexionSevicio <> NullString And InStr(StrConexionSevicio, ";") > 0 Then
      '   en la PC donde se requiere que evitar que levanten el servicio definir una rama en
      '   CURRENT_USER\Software\VB and VBA Program Settings\Algoritmo\Inicio
      '   Parametro ConexionSevicio
      '   valor: NombreServer;1590
      '
      Dim aSocketsData() As String
      
      aSocketsData = Split(StrConexionSevicio, ";")
      
      strServerRoot = aSocketsData(0)
      Winsock1.RemoteHost = aSocketsData(0)        '-> Nombre del Server
      Winsock1.RemotePort = aSocketsData(1)        '-> numero del puerto
   Else
      'obtengo la ubicación en el server de la version del producto
      ' primero me fijo si soy server - tuve que agregar el ModRegistry para que no dependa del SPM
      strServerRoot = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo\Environment\Application Server", "Remote Host", REG_SZ, "", False)
      
      If strServerRoot = NullString Then
         strServerRoot = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo\Environment\Application Server", "Server Update Root", REG_SZ, "", False)
      End If
      
      If Left(strServerRoot, 2) = "\\" Then
         strServerRoot = Right(strServerRoot, Len(strServerRoot) - 2)
      End If
      If InStr(strServerRoot, "\") > 0 Then
         strServerRoot = Left(strServerRoot, InStr(strServerRoot, "\") - 1)
      End If
      
      'elimino los : del drive
      strServerRoot = Replace(strServerRoot, ":", "")
      
      ' Aca busco de irme si hay algun problema o esta mapeado el recurso del server
      If strServerRoot = NullString Then Exit Sub
      If Len(strServerRoot) = 1 Then Exit Sub
      
      Winsock1.RemoteHost = strServerRoot
      Winsock1.RemotePort = 1590
   End If

   Connected = False
   
   'si la conexion falla, se produce el evento tcpClient_Error
   Winsock1.Connect
   
   DoEvents
     
End Sub

Private Sub Winsock1_Connect()
    Connected = True
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '
    '   si le da error en la conexión termina el programa
    '
    Select Case Number
      Case 11001
         MsgBox "El Servidor (" & Winsock1.RemoteHost & ") no se encuentra disponible." & vbCrLf
      Case 10061
         MsgBox "El ""Servicio Inicial de Aplicaciones Algoritmo"" no esta en ejecución en el Servidor: " & Winsock1.RemoteHost
      Case 10060
         MsgBox "El ""Servicio Inicial de Aplicaciones Algoritmo"" no responde en el Servidor: " & Winsock1.RemoteHost & " Puerto: " & Winsock1.RemotePort & vbCrLf & vbCrLf _
         & "ConexionServicio: " & GetSetting("Algoritmo", "Inicio", "ConexionSevicio", "Nada") & vbCrLf _
         & "Remote Host: " & GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo\Environment\Application Server", "Remote Host", REG_SZ, "", False) & vbCrLf _
         & "Server Update Root: " & GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo\Environment\Application Server", "Server Update Root", REG_SZ, "", False)
      Case Else
         MsgBox "Error: " & Number & vbCrLf & Description & " en Winsock1_Error (form1)"
    End Select
    
    End
   
   CancelDisplay = True
   tcpClient.Close

End Sub
