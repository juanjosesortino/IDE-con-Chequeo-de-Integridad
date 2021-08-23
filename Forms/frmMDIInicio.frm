VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{B97E3E11-CC61-11D3-95C0-00C0F0161F05}#163.0#0"; "ALGControls.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.MDIForm frmMDIInicio 
   BackColor       =   &H8000000C&
   Caption         =   "Inicio General de Herramientas"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "frmMDIInicio.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   1365
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   11820
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   11880
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   9000
         Left            =   -90
         Picture         =   "frmMDIInicio.frx":058A
         ScaleHeight     =   9000
         ScaleWidth      =   12000
         TabIndex        =   4
         Top             =   -30
         Width           =   12000
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   3930
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   3
         Top             =   -1680
         Width           =   4095
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3420
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   2820
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin ALGControls.MDITaskBar MDITaskBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   7620
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMDIInicio.frx":1A910
   End
   Begin VB.Timer Timer1 
      Left            =   855
      Top             =   0
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8010
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Estado"
            TextSave        =   "Estado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/08/2009"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "13:01"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5371
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   60
      Top             =   2190
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgEnabled 
      Left            =   660
      Top             =   1700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1A92C
            Key             =   "ANTERIOR"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1AD40
            Key             =   "SIGUIENTE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1B154
            Key             =   "NUEVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1B658
            Key             =   "CANCELAR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1BA6C
            Key             =   "SALVAR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1BB80
            Key             =   "ELIMINAR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1C24C
            Key             =   "VISTA"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1C9D8
            Key             =   "NAVEGAR"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1CEDC
            Key             =   "ORDEN"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1D32C
            Key             =   "FILTROS"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1D658
            Key             =   "IMPRIMIR"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1DD8C
            Key             =   "DETENER"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1E468
            Key             =   "REPETIR"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1E5C2
            Key             =   "ACTUALIZAR"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1EBD6
            Key             =   "BUSQUEDA"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1F02A
            Key             =   "OFFICE"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1F34A
            Key             =   "APLICAR"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1F4A6
            Key             =   "FILTROS_PFT"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1F9A8
            Key             =   "FILTROS_VUELTA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDisabled 
      Left            =   720
      Top             =   2820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":1FEAA
            Key             =   "ANTERIOR"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":20466
            Key             =   "SIGUIENTE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":2096A
            Key             =   "NUEVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":20E6E
            Key             =   "CANCELAR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":21402
            Key             =   "SALVAR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":21956
            Key             =   "ELIMINAR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":21FD2
            Key             =   "VISTA"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":2268E
            Key             =   "NAVEGAR"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":22C72
            Key             =   "ORDEN"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":230C2
            Key             =   "FILTROS"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":2362E
            Key             =   "IMPRIMIR"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":23CEA
            Key             =   "DETENER"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":243C6
            Key             =   "REPETIR"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":24520
            Key             =   "ACTUALIZAR"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":24B9C
            Key             =   "BUSQUEDA"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":24FF0
            Key             =   "OFFICE"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":25310
            Key             =   "APLICAR"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":2546C
            Key             =   "FILTROS_PFT"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIInicio.frx":25AA6
            Key             =   "FILTROS_VUELTA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16x16 
      Left            =   45
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1545
      Top             =   1530
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuToolsItems 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMDIInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmForm    As Form
Attribute frmForm.VB_VarHelpID = -1
Private Const exRunForm       As Integer = 1
Private Const exModulo        As Integer = 2
Private Const exRunShell      As Integer = 3

Private bIsConnected          As Boolean
Private strClientName         As String
Private m_DoneLoading         As Boolean
Private ErrorLog              As ErrType

Private mvarControlData       As DataShare.udtControlData
Private mvarObject            As Object

Private Sub MDIForm_Load()
   
   On Error GoTo GestErr

   tcpClient.RemoteHost = GetSPMProperty(APPSERVERRemoteHost)
   tcpClient.RemotePort = GetSPMProperty(APPSERVERPuerto)
  
  'si la conexion falla, se produce el evento tcpClient_Error
   tcpClient.Connect
   
   'Carga las imagenes
   'iconos para Cambio de Pwd
   img16x16.ListImages.Add , "USUARIO", LoadPicture(Icons & "RNASERV1.ico")
   img16x16.ListImages.Add , "TODOS", LoadPicture(Icons & "RNASERV2.ico")
   
   'iconos para frmConfirmarTickets
   img16x16.ListImages.Add , "Confirmando", LoadPicture(Icons & "flecha16.ico")
   img16x16.ListImages.Add , "Confirmado", LoadPicture(Icons & "ok16.ico")
   img16x16.ListImages.Add , "Error", LoadPicture(Icons & "error16.ico")
   
   'iconos para frmMenuAdmin
   img16x16.ListImages.Add , "LibroCerrado", LoadPicture(Icons & "07_book.ico")
   img16x16.ListImages.Add , "LibroAbierto", LoadPicture(Icons & "08_book.ico")
   img16x16.ListImages.Add , "Root", LoadPicture(Icons & "Mycomp.ico")
   img16x16.ListImages.Add , "Hoja", LoadPicture(Icons & "01853.ico")
   
   'iconos para Monitoreo de Ingresos
   img16x16.ListImages.Add , "Monitoreo", LoadPicture(Icons & "fullscrn.ico")
   img16x16.ListImages.Add , "Retroceder", LoadPicture(Icons & "00592.ico")
   img16x16.ListImages.Add , "Avanzar", LoadPicture(Icons & "00593.ico")
   img16x16.ListImages.Add , "Actualizar", LoadPicture(Icons & "refresh.ico")
   img16x16.ListImages.Add , "Imprimir", LoadPicture(Icons & "printer.ico")
   img16x16.ListImages.Add , "DetalleIngreso", LoadPicture(Icons & "i16-32.ico")
   img16x16.ListImages.Add , "DetalleCamion", LoadPicture(Icons & "i32.ico")
   
   'iconos para la ToolBar Subir/Bajar/Agregar/Quitar/Separador
   img16x16.ListImages.Add , "Nuevo", LoadPicture(Icons & "New.ico")
   img16x16.ListImages.Add , "Delete", LoadPicture(Icons & "Delete.ico")
   img16x16.ListImages.Add , "Subir", LoadPicture(Icons & "00594.ico")
   img16x16.ListImages.Add , "Bajar", LoadPicture(Icons & "00591.ico")
   img16x16.ListImages.Add , "Editar", LoadPicture(Icons & "Pen.ico")
   img16x16.ListImages.Add , "Check", LoadPicture(Icons & "TildeOK.ico")
   img16x16.ListImages.Add , "Separador", LoadPicture(Icons & "Separador.ico")
   
   'iconos para la ToolBar Administracion Listas de Precios
   img16x16.ListImages.Add , "Individual", LoadPicture(Icons & "01865.ico")
   img16x16.ListImages.Add , "Porcentual", LoadPicture(Icons & "01866.ico")
   img16x16.ListImages.Add , "Dirigido", LoadPicture(Icons & "01225.ico")
   img16x16.ListImages.Add , "Generar", LoadPicture(Icons & "00038.ico")
   img16x16.ListImages.Add , "Filtro", LoadPicture(Icons & "Filter.ico")
   
   'iconos para la ToolBar Emisión Comprobantes CV
   img16x16.ListImages.Add , "Origen", LoadPicture(Icons & "Origen.ico")
   img16x16.ListImages.Add , "Vencimientos", LoadPicture(Icons & "00444.ico")
   img16x16.ListImages.Add , "Comentarios", LoadPicture(Icons & "edit.ico")
   img16x16.ListImages.Add , "Percepciones", LoadPicture(Icons & "Percepciones.ico")
   img16x16.ListImages.Add , "Retenciones", LoadPicture(Icons & "Retenciones.ico")
   img16x16.ListImages.Add , "Transporte", LoadPicture(Icons & "train.ico")
   img16x16.ListImages.Add , "InformarCAI", LoadPicture(Icons & "InformarCAI.ico")
   img16x16.ListImages.Add , "Canjes", LoadPicture(Icons & "Canjes.ico")
   img16x16.ListImages.Add , "DatosAdicionales", LoadPicture(Icons & "DatosAdicionales.ico")
   img16x16.ListImages.Add , "MediosPago", LoadPicture(Icons & "MediosPago.ico")
   img16x16.ListImages.Add , "Monedas", LoadPicture(Icons & "pesos.ico")
   img16x16.ListImages.Add , "AjusteTipoCambio", LoadPicture(Icons & "Calculadora16.ico")
   
   'iconos para el menu Edit
   img16x16.ListImages.Add , "Cut", LoadPicture(Icons & "Cut.ico")
   img16x16.ListImages.Add , "Paste", LoadPicture(Icons & "Paste.ico")
   img16x16.ListImages.Add , "Copy", LoadPicture(Icons & "Copy.ico")
   img16x16.ListImages.Add , "Find", LoadPicture(Icons & "Find.ico")
   img16x16.ListImages.Add , "Props", LoadPicture(Icons & "Props.ico")
   img16x16.ListImages.Add , "Undo", LoadPicture(Icons & "Undo.ico")
   
   'iconos para la Emisión de Comprobantes de Tesorería
   img16x16.ListImages.Add , "Cheques", LoadPicture(Icons & "simple.ico")
   'iconos para Administración de Roles
   img16x16.ListImages.Add , "Llave", LoadPicture(Icons & "secur08.ico")
   
   'Iconos Produccion
   'Partes de Produccion
   img16x16.ListImages.Add , "Estado1", LoadPicture(Icons & "TreeEnabled.ico")
   img16x16.ListImages.Add , "Estado2", LoadPicture(Icons & "TreeDisabled.ico")
      
   'CashFlow
   img16x16.ListImages.Add , "Office", LoadPicture(Icons & "ms_office.ico")
   img16x16.ListImages.Add , "Guardar", LoadPicture(Icons & "mdSalva.ico")

   
   CargarFondo
   
   InitMainMenu
   
   Timer1.Interval = 1000
   Timer1.Enabled = True
   
   Do Until m_DoneLoading
      DoEvents
   Loop

   Exit Sub

GestErr:
   LoadError ErrorLog, "MDIForm_Load"
   ShowErrMsg ErrorLog

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ix As Integer

   ' descargo todos los forms menos el MDI
   On Error GoTo GestErr

   Set MRUForms = Nothing
   
   'elimino la coleccion de los modulos
   If Not objSeguridad Is Nothing Then Set objSeguridad.FormsMRU = Nothing
   If Not objContabilidad Is Nothing Then Set objContabilidad.FormsMRU = Nothing
   If Not objFiscal Is Nothing Then Set objFiscal.FormsMRU = Nothing
   If Not objGesCom Is Nothing Then Set objGesCom.FormsMRU = Nothing
   If Not objGeneral Is Nothing Then Set objGeneral.FormsMRU = Nothing
   If Not objCereales Is Nothing Then Set objCereales.FormsMRU = Nothing
   If Not objProduccion Is Nothing Then Set objProduccion.FormsMRU = Nothing

   On Error Resume Next
   
   For ix = MDIExtend1.ExForms.Count To 1 Step -1
      If InStr("frmMDIInicio ", MDIExtend1.ExForms.Item(ix).Name) = 0 Then
         Unload MDIExtend1.ExForms.Item(ix)
      End If
   Next ix
  
   'descargo los forms de Inicio
   For ix = Forms.Count - 1 To 1 Step -1
      If InStr("frmMDIInicio ", Forms(ix).Name) = 0 Then
         Unload Forms(ix)
      End If
   Next ix
   
   
   Exit Sub

GestErr:
   LoadError ErrorLog, "MDIForm_QueryUnload"
   ShowErrMsg ErrorLog
   
End Sub

Private Sub mnuFileExit_Click()
   Unload Me
End Sub

Private Sub mnuToolsItems_Click(Index As Integer)
   CallToolsItem Index
End Sub

Private Sub Timer1_Timer()
   
   m_DoneLoading = True
   Timer1.Enabled = False

End Sub

Private Sub tcpClient_Close()

   If tcpClient.State = sckClosing Then
      
      If bIsConnected = True Then
        
        'necesario para evitar un loop infinito
         bIsConnected = False
         
        'informo al usuario
         MsgBox "La conexión a ' " & strClientName & "' ha sido terminada."

         End
         
      Else
        'informo al usuario
         MsgBox "El servidor no esta corriendo." & vbCrLf & vbCrLf & _
                "Causas Posibles:" & vbCrLf & vbCrLf & _
                " - Es posible que el servdior no haya sido correctamente activado." & vbCrLf & _
                " - El Servicio ha sido momentáneamente suspendido"
        End
         
      End If
                
      'cierro para permitir reconectarme
       tcpClient.Close
      
   End If
   
End Sub
Private Sub tcpClient_Connect()
   If bIsConnected = False Then

      'esta el la primera vez que me conecto al server
      'le trasnmito al server mi nombre asi me identifica
      If tcpClient.State = sckConnected Then
      
         'Aqui ahora devuelve user@servidor cuando es user del T.Service
         tcpClient.SendData CSysEnvironment.Machine & ";" & App.EXEName
      End If
      
   End If
   
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
Dim strData As String  'guardo los datos recibidos
   
   If bIsConnected = True Then
   
     'la conexion esta establecida. Los datos que llegan forman parte del dialogo
      tcpClient.GetData strData
      
      If strData = "Islogued" Then
         bUserLogued = 1
         Exit Sub
      End If
      If strData = "Notlogued" Then
         bUserLogued = 2
         Exit Sub
      End If

      MsgBox strData, vbOKOnly, "Consola"
      
   Else
   
      ' es el primer dato recibido del cliente (su nombre)
      bIsConnected = True
      
     'le transmito mi nombre al cliente
      tcpClient.GetData strData
      
      strClientName = strData
      
'      Me.Caption = "TCP Client : Dialogando con " & strData
      
   End If

End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, _
                            Description As String, _
                            ByVal Scode As Long, _
                            ByVal source As String, _
                            ByVal HelpFile As String, _
                            ByVal HelpContext As Long, _
                            CancelDisplay As Boolean)

   Select Case Number
      Case 10061
      
         MsgBox "Error: " & Number & vbCrLf & Description & vbCrLf & vbCrLf & _
                "El server no esta corriendo o no ha sido establecida la conexión." & vbCrLf & _
                "Causas Posibles:" & vbCrLf & _
                "  Es posible que el server no haya sido correctamente activado." & vbCrLf & _
                "  El Servicio ha sido momentaneamente suspendido"
                
         End
         
      Case Else
         MsgBox "Error: " & Number & vbCrLf & Description & " en tcpClient_Error (frmMDIInicio)"
         End
   End Select
   
   CancelDisplay = True
   tcpClient.Close

End Sub

Public Sub InitMainMenu(Optional f As Form)

   If f Is Nothing Then
      Set f = Me
   End If

   '  menu Herramientas
   aMenuTools(0) = "&Chequeo de Integridad del Sistema ...;frmChequearIntegridad"
   'aMenuTools(1) = "&Establecer Ejercicio Contable Activo...;frmEjercicioContableActivo" ' QUITAR
   
   LoadMenu "Tools", f, aMenuTools

End Sub

Private Sub LoadMenu(strMenuName As String, frm As Form, aMenuTools() As String)
Dim ix   As Integer
Dim iPos As Integer

   '  carga el menu indicado por el parametro strMenuname en el form frm
   
   On Error Resume Next
   
   Select Case strMenuName
   
      Case "Tools"
         '  menu Herramientas
         For ix = 0 To UBound(aMenuTools)
            If Len(aMenuTools(ix)) > 0 Then
               If ix > 0 Then
                  Load frm.mnuToolsItems(ix)
               End If
               iPos = InStr(aMenuTools(ix), ";")
               If iPos > 0 Then
                  frm.mnuToolsItems(ix).Caption = Left(aMenuTools(ix), iPos - 1)
               Else
                  frm.mnuToolsItems(ix).Caption = aMenuTools(ix)
               End If
            End If
         Next ix
         
   End Select

End Sub

Public Sub CallToolsItem(ByVal iItem As Integer)
Dim strForm          As String
Dim Form             As Form
Dim ix               As Integer

   ' llama una opcíon del menú Herramientas

   If iItem > UBound(aMenuTools) Then Exit Sub

   ix = InStr(aMenuTools(iItem), ";")
   If ix > 0 Then
      strForm = Mid(aMenuTools(iItem), ix + 1)
   Else
      strForm = ""
   End If
   
   rstMenu.Filter = adFilterNone

   Select Case True
      Case UCase(strForm) = UCase("frmChequearIntegridad")

         Set Form = New frmChequearIntegridad
         With mvarControlData
            .Empresa = objInfoEmpresa.CodigoEmpresa
            .Sucursal = objInfoEmpresa.SucursalActiva
            .Ejercicio = objInfoEmpresa.EjercicioVigente
            .Usuario = CUsuario.Usuario
            .Maquina = CSysEnvironment.Machine
            .MenuKey = "ChequearIntegridad"
         End With
         Form.ControlData = mvarControlData
         Form.Show
               
         Case Else
            Set Form = Forms.Add(Trim(strForm))
            Form.Show vbModal
   End Select
   
End Sub

Public Function CallAdmin(ByVal strMenuKeyAdmin As String, ByVal ControlData As Variant, Optional bMRUMode As Boolean = False) As Long

   '-- llamo al Form para la administracion del control activo
   '-- ControlInfo es del tipo ControlType
   
   On Error GoTo GestErr
   
   If strMenuKeyAdmin = NullString Then
      MsgBox "Es probable que no se haya definido una Clave de Menu Administrar", vbOKOnly, App.ProductName
      Exit Function
   Else
      ControlData.MenuKey = strMenuKeyAdmin
   End If
   
   rstMenu.Filter = "MNU_CLAVE = '" & strMenuKeyAdmin & "'"
   
   With ErrorLog
      .Form = Me.Name
      .Empresa = ControlData.Empresa
   End With
   
   Select Case True
      Case UCase(rstMenu("MNU_MODULO")) = UCase(App.ProductName)
     
         Select Case rstMenu("MNU_TIPO_EXEC")
               
            Case exRunForm
            Case exModulo
                  
            Case exRunShell
               Shell rstMenu("MNU_NOMBRE_EXEC")
               
         End Select
         
         Exit Function
         
       Case rstMenu("MNU_MODULO") <> UCase(App.ProductName)
       
         Select Case UCase(rstMenu("MNU_MODULO"))
            Case "ADMINISTRADORGENERAL"
            
               Select Case rstMenu("MNU_TIPO_EXEC")
                  Case exRunForm
                  
                     If objGeneral Is Nothing Then
                        Set objGeneral = CreateObject("AdministradorGeneral.Application")
                        Set objGeneral = SetApplication(objGeneral)
                     End If
                     
                     CallAdmin = ShowForm(ControlData, rstMenu("MNU_CLAVE"), objGeneral, bMRUMode)
                     
                  Case exModulo
                  Case exRunShell
                     Shell rstMenu("MNU_NOMBRE_EXEC")
               End Select
            
            Case "SEGURIDAD"
            
               Select Case rstMenu("MNU_TIPO_EXEC")
                  Case exRunForm
                  
                     If objSeguridad Is Nothing Then
                        Set objSeguridad = CreateObject("Seguridad.Application")
                        Set objSeguridad = SetApplication(objSeguridad)
                     End If
                  
                     CallAdmin = ShowForm(ControlData, rstMenu("MNU_CLAVE"), objSeguridad, bMRUMode)
                     
                  Case exModulo
                  Case exRunShell
                     Shell rstMenu("MNU_NOMBRE_EXEC")
               End Select
               
            Case "CONTABILIDAD"
            
               Select Case rstMenu("MNU_TIPO_EXEC")
                  Case exRunForm
                  
                     If objContabilidad Is Nothing Then
                        Set objContabilidad = CreateObject("Contabilidad.Application")
                        Set objContabilidad = SetApplication(objContabilidad)
                     End If
                  
                     CallAdmin = ShowForm(ControlData, rstMenu("MNU_CLAVE"), objContabilidad, bMRUMode)
                     
                  Case exModulo
                  Case exRunShell
                     Shell rstMenu("MNU_NOMBRE_EXEC")
               End Select
               
            Case "CEREALES"
            
               Select Case rstMenu("MNU_TIPO_EXEC")
                  Case exRunForm
                    
                     If objCereales Is Nothing Then
                        Set objCereales = CreateObject("Cereales.Application")
                        Set objCereales = SetApplication(objCereales)
                     End If
                  
                     CallAdmin = ShowForm(ControlData, rstMenu("MNU_CLAVE"), objCereales, bMRUMode)
                     
                  Case exModulo
                  Case exRunShell
                     Shell rstMenu("MNU_NOMBRE_EXEC")
               End Select
               
            Case "FISCAL"
            
               Select Case rstMenu("MNU_TIPO_EXEC")
                  Case exRunForm
                  
                     If objFiscal Is Nothing Then
                        Set objFiscal = CreateObject("Fiscal.Application")
                        Set objFiscal = SetApplication(objFiscal)
                     End If
                  
                     CallAdmin = ShowForm(ControlData, rstMenu("MNU_CLAVE"), objFiscal, bMRUMode)
                     
                  Case exModulo
                  Case exRunShell
                     Shell rstMenu("MNU_NOMBRE_EXEC")
               End Select
               
            Case "GESTIONCOMERCIAL"
            
               Select Case rstMenu("MNU_TIPO_EXEC")
                  Case exRunForm
                  
                     If objGesCom Is Nothing Then
                        Set objGesCom = CreateObject("GestionComercial.Application")
                        Set objGesCom = SetApplication(objGesCom)
                     End If
                  
                     CallAdmin = ShowForm(ControlData, rstMenu("MNU_CLAVE"), objGesCom, bMRUMode)
                     
                  Case exModulo
                  Case exRunShell
                     Shell rstMenu("MNU_NOMBRE_EXEC")
               End Select
            
            Case "PRODUCCION"
            
               Select Case rstMenu("MNU_TIPO_EXEC")
                  Case exRunForm
                  
                     If objProduccion Is Nothing Then
                        Set objProduccion = CreateObject("Produccion.Application")
                        Set objProduccion = SetApplication(objProduccion)
                     End If
                  
                     CallAdmin = ShowForm(ControlData, rstMenu("MNU_CLAVE"), objProduccion, bMRUMode)
                     
                  Case exModulo
                  Case exRunShell
                     Shell rstMenu("MNU_NOMBRE_EXEC")
               End Select

         End Select
         
   End Select
   
   Exit Function
   
GestErr:
   LoadError ErrorLog, "CallAdmin"
   ShowErrMsg ErrorLog
End Function

Public Function ShowForm(ByVal ControlData As Variant, _
                        ByVal strMenuKey As String, _
                        ByVal objApp As Object, _
                        Optional bMRUMode As Boolean = False) As Long
   
Dim retval As Long
   
   '/
   '  Muestra un form.
   '  Si el form que se desea visualizar existe en el MRUForms y no esta invisible entonces
   '  hago el Show de dicho form
   '  Si el form no esta en el MRUForms o bien esta pero se requiere una nueva instancia,
   '  entonces cargo la nueva instancia
   '
   '  Devuelve siempre el hWnd del form mostrado
   '/
   
   On Error Resume Next

   Set frmForm = Nothing
   
   'La busqueda en MRUForms la hago con la clave del menu
   Set frmForm = MRUForms(Trim(strMenuKey))
   If Not (frmForm Is Nothing) Then

      On Error GoTo GestErr

       '/ es un form presente en MRUForms
       
      If frmForm.Visible = False And frmMDIInicio.MDITaskBar1.FormInTaskBar(frmForm.hWnd) = False Then
         'form MRU no visible y no presente en la barra
         If Not frmForm Is Nothing Then frmForm.Visible = True
         If Not frmForm Is Nothing Then frmForm.MenuKey = strMenuKey
         If Not frmForm Is Nothing Then frmForm.ControlData = ControlData
         
         If Not frmForm Is Nothing Then frmForm.InitForm
         If Not frmForm Is Nothing Then frmForm.PostInitForm
         If Not frmForm Is Nothing Then ShowForm = frmForm.hWnd
         
         'agrego el form en la TaskBar
         If Not frmForm Is Nothing Then frmForm.MDIExtend1.WindowState = vbNormal
         If Not frmForm Is Nothing Then frmMDIInicio.MDITaskBar1.AddFormToTaskBar frmForm
         If Not frmForm Is Nothing Then
            If Not frmMDIInicio.MDITaskBar1.FormInTaskBar(frmForm.hWnd) Then
               CenterMDIActiveXChild frmForm
            End If
         End If
         
         Exit Function
      End If
   End If

   Set frmForm = objApp.GetMDIChild(strMenuKey, bMRUMode)
   
   '/ deshabilito el form
   retval = EnableWindow(frmForm.hWnd, 0)
   
   If Not frmForm Is Nothing Then frmForm.MenuKey = strMenuKey
   If Not frmForm Is Nothing Then frmForm.ControlData = ControlData
   
   If Not frmForm Is Nothing Then frmForm.InitForm
   
   '/ habilito el form
   retval = EnableWindow(frmForm.hWnd, 1)
   If Not frmForm Is Nothing Then frmForm.PostInitForm
   If Not frmForm Is Nothing Then frmForm.ZOrder

   If Not frmForm Is Nothing Then ShowForm = frmForm.hWnd
   
   Screen.MousePointer = vbDefault
   
   Exit Function
   
GestErr:
   LoadError ErrorLog, "ShowForm"
   ShowErrMsg ErrorLog
End Function

Private Sub frmForm_Unload(Cancel As Integer)
   Set frmForm = Nothing
End Sub

Public Function GetForm(ByVal strFormName As String) As Form
   
   On Error GoTo GestErr
   
   Set GetForm = Forms.Add(strFormName)
   
   Exit Function

GestErr:
   LoadError ErrorLog, "GetForm"
   ShowErrMsg ErrorLog
End Function
Public Function GetInstance(ByVal strInstanceName As String) As Object

   Select Case UCase(strInstanceName)
      Case "ADMINISTRADORGENERAL"
         If objGeneral Is Nothing Then
            Set objGeneral = CreateObject("AdministradorGeneral.Application")
            Set objGeneral = SetApplication(objGeneral)
         End If
         Set GetInstance = objGeneral
         
      Case "SEGURIDAD"
         If objSeguridad Is Nothing Then
            Set objSeguridad = CreateObject("Seguridad.Application")
            Set objSeguridad = SetApplication(objSeguridad)
         End If
         Set GetInstance = objSeguridad
         
      Case "CONTABILIDAD"
         If objContabilidad Is Nothing Then
            Set objContabilidad = CreateObject("Contabilidad.Application")
            Set objContabilidad = SetApplication(objContabilidad)
         End If
         Set GetInstance = objContabilidad
         
      Case "CEREALES"
         If objCereales Is Nothing Then
            Set objCereales = CreateObject("Cereales.Application")
            Set objCereales = SetApplication(objCereales)
         End If
         Set GetInstance = objCereales
         
      Case "FISCAL"
         If objFiscal Is Nothing Then
            Set objFiscal = CreateObject("Fiscal.Application")
            Set objFiscal = SetApplication(objFiscal)
         End If
         Set GetInstance = objFiscal
         
      Case "GESTIONCOMERCIAL"
         If objGesCom Is Nothing Then
            Set objGesCom = CreateObject("GestionComercial.Application")
            Set objGesCom = SetApplication(objGesCom)
         End If
         Set GetInstance = objGesCom
      
      Case "PRODUCCION"
         If objProduccion Is Nothing Then
            Set objProduccion = CreateObject("Produccion.Application")
            Set objProduccion = SetApplication(objProduccion)
         End If
         Set GetInstance = objProduccion
   End Select

End Function
Public Property Get Objeto() As Object
   Set Objeto = mvarObject
End Property
Public Property Set Objeto(ByVal obj As Object)
   Set mvarObject = obj
End Property
Private Sub CargarFondo()
   picStretched.Move 0, 0, _
      Screen.Width, Screen.Height

   picStretched.PaintPicture _
      picOriginal.Picture, _
      0, 0, _
      picStretched.ScaleWidth, _
      picStretched.ScaleHeight, _
      0, 0, _
      picOriginal.ScaleWidth, _
      picOriginal.ScaleHeight

   Set Picture = picStretched.Image
End Sub
Public Sub Largar() '(Cancel As Integer, UnloadMode As Integer)
Dim ix As Integer

   ' descargo todos los forms menos el MDI
   On Error GoTo GestErr

   Set MRUForms = Nothing
   
   'elimino la coleccion de los modulos
   If Not objSeguridad Is Nothing Then Set objSeguridad.FormsMRU = Nothing
   If Not objContabilidad Is Nothing Then Set objContabilidad.FormsMRU = Nothing
   If Not objFiscal Is Nothing Then Set objFiscal.FormsMRU = Nothing
   If Not objGesCom Is Nothing Then Set objGesCom.FormsMRU = Nothing
   If Not objGeneral Is Nothing Then Set objGeneral.FormsMRU = Nothing
   If Not objCereales Is Nothing Then Set objCereales.FormsMRU = Nothing
   If Not objProduccion Is Nothing Then Set objProduccion.FormsMRU = Nothing

   On Error Resume Next
   
   For ix = MDIExtend1.ExForms.Count To 1 Step -1
      If InStr("frmMDIInicio ", MDIExtend1.ExForms.Item(ix).Name) = 0 Then
         If InStr("frmChequearIntegridad ", MDIExtend1.ExForms.Item(ix).Name) = 0 Then
            Unload MDIExtend1.ExForms.Item(ix)
         End If
      End If
   Next ix
  
   'descargo los forms de Inicio
   For ix = Forms.Count - 1 To 1 Step -1
      If InStr("frmMDIInicio ", Forms(ix).Name) = 0 Then
         If InStr("frmChequearIntegridad ", Forms(ix).Name) = 0 Then
            Unload Forms(ix)
         End If
      End If
   Next ix
   
   
   Exit Sub

GestErr:
   LoadError ErrorLog, "Release"
   ShowErrMsg ErrorLog
   
End Sub
