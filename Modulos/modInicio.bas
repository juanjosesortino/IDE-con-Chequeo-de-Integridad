Attribute VB_Name = "modInicio"
Option Explicit
Private mvarControlData       As DataShare.udtControlData

'***********************************************************************
' Constantes Propias
'***********************************************************************
Public Const Si                  As String = "Sí"
Public Const No                  As String = "No"
Public Const NullString          As String = ""
Public Const UNKNOWN_ERRORSOURCE As String = "[Fuente de Error Desconocida]"
Public Const KNOWN_ERRORSOURCE   As String = "[Fuente de Error Conocida]"
'***********************************************************************
'Private Const MODULE_NAME        As String = "[ModInicio]"
Private ErrorLog                 As ErrType

Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const SPI_GETWORKAREA = 48

'Private objTextBox As New AlgStdFunc.clsTextBoxEdit

Public CUsuario                  As BOSeguridad.clsUsuario
Public CSysEnvironment           As AlgStdFunc.clsSysEnvironment
Public rstContextMenu            As ADODB.Recordset
Public rstMenu                   As ADODB.Recordset
Public rstVistasPersonalizadas   As ADODB.Recordset
Public rstVistasExportacion      As ADODB.Recordset
Public rstEmpresas               As ADODB.Recordset
Public RegistrySubKeys()         As Variant
Public SystemOptions()           As Variant
Public strSucursalElegida        As String

'/ deben ser visibles por frmMenu y frmMDIInicio
Public objSeguridad              As Object
Public objContabilidad           As Object
Public objFiscal                 As Object
Public objGesCom                 As Object
Public objGeneral                As Object
Public objCereales               As Object
Public objProduccion             As Object

'/ deben ser visibles por frmLlogin
Public bUserLogued               As Long

Public Enum EnumRegistrySubKeys
   Environment = 0
   DataBaseSettings = 1
   KeyMRUForms = 2
   MRUEmpresas = 3
   GridQueries = 4
   NavigationQueries = 5
   PrintQueries = 6
   QueryDBQueries = 7
   DataComboQueries = 8
   [_MAX_Value] = 8
End Enum

Public Enum EnumSystemOptions
   iCacheSize = 0             'valor del parámetro cachesize (registro del sistema)
   iZoom = 1                  'valor del Zoom por defecto en Vista Previa
   iFetchMode = 2             'indica el modo en el que vendran capturados los registros del server
   lngFetchLimit = 3          'si alFetchMode = 2, es el limite de registros recuperados en manera sincronica
   iFetchModeSearch = 4       'indica el modo en el que vendran capturados los registros del server (para la busqueda)
   lngFetchLimitSearch = 5    'si alFetchMode = 2, es el limite de registros recuperados en manera sincronica (para la busqueda)
   UseLocalCopy = 6           'Sí=Usa copias locales; No=Usa copias locales (Vista-Lista, Navegación e Impresión)
   UseLocalCopySearch = 7     'Sí=Usa copias locales; No=Usa copias locales (para la búsqueda)
   AskOldLocalCopy = 8        'Sí=Pregunta si usa copias locales desactualizadas;(Vista-Lista, Navegación e Impresión)
   UseMRUEnterprise = 9       'Si=recuerda las ultimas empresas;No=No recuerda
   MaxMRUForms = 10           'Dimension de la colecion MRUForms
   [_MAX_Value] = 10
End Enum

Public Enum ContextMenuEnum
   mnxNombre = 0
   mnxOrden = 1
   mnxForms = 2
   mnxCaption = 3
   mnxTarea = 4
   mnxClave = 5
End Enum


'Mensajes enviados por la Filter
Public Const FILTER_CALL_ADMIN = &H1
Public Const FILTER_QUERY_USER   As Long = &H2
Public Const FILTER_QUERY_CONTROLDATA   As Long = &H3

'  constantes para identificar los mensajes devueltos por Filter
Public Const MSG_CANCEL  As String = "CANCELFILTRO"
Public Const MSG_CONFIRM As String = "CONFIRMAFILTRO"
Public Const MSG_APPLY   As String = "APLICARFILTRO"

'  constantes para identificar los paneles del Status Bar
Public Const STB_PANEL1              As Integer = 1
Public Const STB_PANEL2              As Integer = 2
Public Const STB_PANEL3              As Integer = 3
Public Const STB_PANEL4              As Integer = 4

Public Enum alFetchMode
   alAsync = 1
   alSync = 2
   alTable = 3
End Enum

Public Const IX_CAMBIO_PWD          As Integer = 0
Public Const IX_ESTABLECER_EJERCI   As Integer = 1
Public Const IX_CAMBIO_LOGIN        As Integer = 2
Public Const IX_EMPRESAS            As Integer = 3
Public Const IX_SEPARA1             As Integer = 4
Public Const IX_LOG_ERRORES         As Integer = 5
Public Const IX_EDITOR_SQL          As Integer = 6
Public Const IX_SEPARA2             As Integer = 7
Public Const IX_BUSCAR_MENU         As Integer = 8
Public Const IX_BUSCAR_SIGUIENTE    As Integer = 9
Public Const IX_SEPARA3             As Integer = 10
Public Const IX_PROPIEDADES         As Integer = 11
Public Const IX_ARCHIVOS            As Integer = 12
Public Const IX_SEPARA4             As Integer = 13
Public Const IX_VER_FAVORITOS       As Integer = 14
Public Const IX_ORGANIZAR_FAVORITOS As Integer = 15
Public Const IX_SEPARA5             As Integer = 16
Public Const IX_OPCIONES            As Integer = 17 ' ---> ver aMenuTools(17)

Public aMenuTools(17)            As String                                                 'matriz elementos del menu Herramientas

Public mvarMDIForm               As MDIForm
Public MRUForms                  As New Collection

Public bRestart                  As Boolean    'sirve para saber si me estoy logeando como otro usuario

Public colInfoEmpresas As clsInfoEmpresas
Public objInfoEmpresa As clsInfoEmpresa

Private bStartUp  As Boolean

Sub Main()
Dim objSPM  As DataShare.SPM
Dim aEmpresas()    As Variant
Dim ix             As Integer
Dim strEjercicioActivo As String
   On Error GoTo GestErr
   
   Dim strServerRoot As String
   Dim ServerDLLFolder As String
   '
   '  intenta establecer una conexón con el servicio
   '
   Load frmTestServicio
   Do While Not frmTestServicio.Connected
      DoEvents
   Loop
   Unload frmTestServicio
  
'
'        Esto cambia el mensaje "Cambiar - Reintentar" por éste un poco más amigable
'
   App.OleRequestPendingMsgText = "El Servidor está ocupado procesando su requerimiento." & vbCrLf & vbCrLf & "Por favor, aguarde unos instantes."
   App.OleRequestPendingMsgTitle = "Aplicaciones Algoritmo"
  
   bStartUp = True
   
   Set colInfoEmpresas = New clsInfoEmpresas
   
   Set CUsuario = New BOSeguridad.clsUsuario
   
   Set CSysEnvironment = New AlgStdFunc.clsSysEnvironment
   
   SetRegistryEntries
   
   ReadSystemOptions

   ' defino el application path para la clase clsEnvironment
   SetAppPath App.Path
   
   With ErrorLog
      .Empresa = GetSPMProperty(DBSEmpresaPrimaria)
      .Maquina = CSysEnvironment.Machine
      .Aplicacion = App.EXEName
   End With
   
   ' Leo el registro de la empresa primaria para ver si usa o no cache local
   '  ReadDefault
   ReDim aKeys(1, 1)
   aKeys(0, 0) = "Opciones\Utiliza cache local;" & No
   aKeys(1, 0) = "Opciones\Permite Múltiples Instancias de la Aplicacion;" & No
   Set objSPM = GetMyObject("DataShare.SPM")
   objSPM.GetKeyValues objSPM.GetSPMProperty(DBSEmpresaPrimaria), aKeys
   
   Set rstMenu = objSPM.GetSPMProperty(MNURecordset)
   Set rstContextMenu = objSPM.GetSPMProperty(CMURecordset)
  
   Set objSPM = Nothing
   
   'muestro el form MDI
   frmMDIInicio.Show
   
   Set mvarMDIForm = frmMDIInicio
   
   '  login usuario
   With mvarControlData
       .Empresa = NullString
       .Sucursal = NullString
       .Maquina = CSysEnvironment.Machine
   End With
   
   CUsuario.ControlData = mvarControlData
   CUsuario.Load "ADMIN"

'------------
   CrearRstEmpresas
   aEmpresas = EnumRegistryValues(HKEY_LOCAL_MACHINE, "SoftWare\Algoritmo\MRU Empresas\" & CUsuario.Usuario)
   On Error Resume Next
   ix = LBound(aEmpresas, 2)
   
   If Err.Number = 0 Then
      Err.Clear
      On Error GoTo GestErr
      For ix = LBound(aEmpresas, 2) To UBound(aEmpresas, 2)
         If aEmpresas(0, ix) <> NullString Then
            rstEmpresas.Filter = "EMP_CODIGO_EMPRESA = '" & aEmpresas(0, ix) & "'"
            
            strEjercicioActivo = GetEjercicioActivo(aEmpresas(0, ix)) 'Inc.: 41774
            If strEjercicioActivo <> "ERROR" Then
               If rstEmpresas.RecordCount <> 0 Then ' ESTE IF ES PARA NO
                                           '  AGREGAR EN EL TAB EMPRESAS NO HABILITADAS
                                           '  PARA EL USUARIO
                  
                  Set objInfoEmpresa = New clsInfoEmpresa
                  
                  With objInfoEmpresa
                  
                     .CodigoEmpresa = aEmpresas(0, ix)
                     .NombreEmpresa = NombreEmpresa(.CodigoEmpresa)
                     .EjercicioVigente = strEjercicioActivo
                        
                     .SucursalActiva = GetSucursalActiva(.CodigoEmpresa)

                     colInfoEmpresas.Add objInfoEmpresa, .CodigoEmpresa
                  End With
               End If
            End If
         End If
         rstEmpresas.Filter = adFilterNone
         
      Next ix
   End If
   Set objInfoEmpresa = colInfoEmpresas.Item(1)
   mvarControlData.Ejercicio = objInfoEmpresa.EjercicioVigente
   mvarControlData.Empresa = objInfoEmpresa.CodigoEmpresa
   mvarControlData.Sucursal = objInfoEmpresa.SucursalActiva
   mvarControlData.Usuario = "ADMIN"
'------------
   CUsuario.ControlData = mvarControlData
   
   SetRegistryEntries CUsuario.Usuario
   
   frmMDIInicio.stbMain.Panels(1).Text = "Cargando Opciones del Sistema ..."
   frmMDIInicio.stbMain.Panels(1).Text = CUsuario.Usuario
   
   bStartUp = False
   
   Exit Sub
   
GestErr:
   Set objSPM = Nothing

   LoadError ErrorLog, "Main" & Erl
   
   If bStartUp Then
   
      Select Case True

         Case CUsuario Is Nothing
               MsgBox "Se produjo un error en la linea " & Erl & vbCrLf & _
                      "Error en la creación del objeto Usuario durante la fase de Inicio del sistema. " & _
                      "Controle que todos sus componentes este correctamente registrados", vbExclamation, App.ProductName
         Case Else
   
               MsgBox "Se produjo un error en la linea " & Erl & vbCrLf & _
                      "Error durante la inicialización del Sistema. Las posibles causas de este inconveniente podrían ser: " & vbCrLf & vbCrLf & _
                      "    - uno o mas Componentes no estan correctamente registrados en su PC Local" & vbCrLf & _
                      "    - el equipo Servidor no esta activo" & vbCrLf & _
                      "    - se produjo un error al intentar establecer un conexíon con el Paquete MTS" & vbCrLf & vbCrLf & _
                      "Intente Reiniciar el Paquete MTS en el servidor eliminando previamente todos los procesos mtx.exe visibles en el Administrador de Tareas del Servidor." & _
                      "Si el problema persiste, retroceda a la versión precedente. Si aún no ha podido iniciar el sistema contacte a su proveedor", vbExclamation, App.ProductName

      End Select

      End
   Else
      ShowErrMsg ErrorLog
   End If
   
End Sub

Public Function CallAdmin(ByVal ControlInfo As Variant, ByVal ControlData As Variant) As Long
   
   On Error GoTo GestErr
   
   CallAdmin = mvarMDIForm.CallAdmin(ControlInfo.MenuKeyAdmin, ControlData)
   
   Exit Function
   
GestErr:
   LoadError ErrorLog, "CallAdmin"
   ShowErrMsg ErrorLog
End Function

Public Sub LoadError(ByRef ErrLog As ErrType, ByVal strSource As String)
Dim PropBag As PropertyBag

   ' carga la información del error en la variable ErrorLog
   SetError ErrLog, App.ProductName, strSource
   
   Set PropBag = New PropertyBag
   
   With PropBag
      .WriteProperty "ERR_EMPRESA", ErrLog.Empresa
      .WriteProperty "ERR_APLICACION", ErrLog.Aplicacion
      .WriteProperty "ERR_COMENTARIO", ErrLog.Comentario
      .WriteProperty "ERR_DESCRIPCION", ErrLog.Descripcion
      .WriteProperty "ERR_ERRORNATIVO", ErrLog.ErrorNativo
      .WriteProperty "ERR_FORM", ErrLog.Form
      .WriteProperty "ERR_MAQUINA", ErrLog.Maquina
      .WriteProperty "ERR_MODULO", ErrLog.Modulo
      .WriteProperty "ERR_NUMERROR", ErrLog.NumError
      .WriteProperty "ERR_SOURCE", ErrLog.source
      .WriteProperty "ERR_USUARIO", ErrLog.Usuario
      .WriteProperty "WRITE_ERROR", ErrLog.WriteError
   End With
   
   TrapError PropBag.Contents
   Set PropBag = Nothing
   
   
End Sub
Public Sub SetError(ByRef ErrLog As ErrType, ByVal strModuleName As String, ByVal strSource As String)

   With ErrLog
   
      If Not CUsuario Is Nothing Then
         .Usuario = CUsuario.Usuario
      End If
      
      .Modulo = strSource
      .NumError = Err.Number
      .source = Err.source
      .Aplicacion = UCase(App.ProductName)
      .WriteError = Si
      
      If InStr(.source, KNOWN_ERRORSOURCE) = 0 Then
         If InStr(.source, UNKNOWN_ERRORSOURCE) = 0 Then
            .source = UNKNOWN_ERRORSOURCE & vbCrLf & .source
         End If
      Else
         .source = Replace(.source, KNOWN_ERRORSOURCE, NullString)
         .WriteError = False
      End If
      
      If InStr(.source, strModuleName) > 0 Then
         .source = .source & vbCrLf & "[" & strSource & "]"
      Else
         .source = .source & vbCrLf & strModuleName & "[" & strSource & "]"
      End If
      
      .Descripcion = Err.Description

   End With

   
End Sub

Public Sub ShowErrMsg(ByRef ErrorLog As ErrType)
Dim iErrNumber         As Long                          ' numero de error (sin vbObjectError)
Dim bAlgError          As Boolean                       ' identifica un error de Algoritmo
'Dim ix                 As Integer
Dim strSource          As String
Dim n                  As Integer
Dim frmMsg             As frmMsgBox
Dim strMensaje         As String
Dim strDetalle         As String

   '  muestra en manera amigable un mensaje de error
   

   strSource = Trim(ErrorLog.source)
   
   bAlgError = True
   n = InStr(strSource, UNKNOWN_ERRORSOURCE)
   If n > 0 Then
      ' es un error generado por alguna aplicacion de Algoritmo
      bAlgError = False
   End If
   
   strSource = Replace(strSource, UNKNOWN_ERRORSOURCE, NullString)
   
   If bAlgError Then
      ' errores de Algortimo
      iErrNumber = ErrorLog.NumError - vbObjectError
      Select Case iErrNumber
         Case Is < 10000
            ' warnings de Algortimo
            
            Set frmMsg = New frmMsgBox
            
            strMensaje = ErrorLog.Descripcion
            strDetalle = strSource
            
            frmMsg.ShowMsg App.ProductName, strMensaje, strDetalle, Warning
            
         Case 10000 To 20000
            'Errores Severos de Algoritmo
            
            Set frmMsg = New frmMsgBox
            
            strMensaje = ErrorLog.Descripcion
            frmMsg.ShowMsg App.ProductName, strMensaje, strDetalle, Error
            
         Case Else
         
            Set frmMsg = New frmMsgBox
            
            strMensaje = ErrorLog.Descripcion
            frmMsg.ShowMsg App.ProductName, strMensaje, strDetalle, Error
         
      End Select
   Else
      ' errores no generados por Algoritmo
             
         Set frmMsg = New frmMsgBox
          
         ErrorLog.Descripcion = Replace(ErrorLog.Descripcion, vbCr, NullString)
         ErrorLog.Descripcion = Replace(ErrorLog.Descripcion, vbLf, NullString)
          
         strMensaje = ErrorLog.Descripcion
         strDetalle = "Número     : " & ErrorLog.NumError & vbCrLf & strSource
         
         frmMsg.ShowMsg App.ProductName, strMensaje, strDetalle, Error
            
             
   End If
   
   ' una vez visualizado el mensaje de error, este viene limpiado
   With ErrorLog
      .Modulo = NullString
      .NumError = 0
      .source = NullString
      .Descripcion = NullString
   End With
   
   Screen.MousePointer = vbDefault
   
End Sub

Public Sub CenterMDIActiveXChild(ByVal frmChild As Form)

   '--  centra el form MDIActiveX Child

   frmChild.Move (mvarMDIForm.ScaleWidth - frmChild.Width) / 2, (mvarMDIForm.ScaleHeight - frmChild.Height) / 2

End Sub

Public Sub CenterForm(ByRef frm As Form)
Dim r As RECT
Dim lRes As Long
Dim lw As Long
Dim lh As Long

   lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, r, 0)

   If lRes Then
      With r
         .Left = Screen.TwipsPerPixelX * .Left
         .Top = Screen.TwipsPerPixelY * .Top
         .Right = Screen.TwipsPerPixelX * .Right
         .Bottom = Screen.TwipsPerPixelY * .Bottom
         lw = .Right - .Left
         lh = .Bottom - .Top
         
         frm.Move .Left + (lw - frm.Width) \ 2, .Top + (lh - frm.Height) \ 2
      End With
   End If

End Sub

Public Sub SetRegistryEntries(Optional ByVal strUser As String)

         '  setea la ubicación de las claves del registro de windows
      
10       On Error GoTo GestErr
      
20       ReDim RegistrySubKeys(EnumRegistrySubKeys.[_MAX_Value])
   
30       If Len(strUser) = 0 Then
40          RegistrySubKeys(EnumRegistrySubKeys.DataBaseSettings) = "Software\Algoritmo\DataBaseSettings"
50          RegistrySubKeys(EnumRegistrySubKeys.Environment) = "Software\Algoritmo\Environment"
'60          RegistrySubKeys(EnumRegistrySubKeys.NavigationQueries) = "Software\Algoritmo\MRU Queries\NavigationStoredQueries"
'70          RegistrySubKeys(EnumRegistrySubKeys.GridQueries) = "Software\Algoritmo\MRU Queries\GridStoredQueries"
'80          RegistrySubKeys(EnumRegistrySubKeys.PrintQueries) = "Software\Algoritmo\MRU Queries\PrintStoredQueries"
'90          RegistrySubKeys(EnumRegistrySubKeys.QueryDBQueries) = "Software\Algoritmo\MRU Queries\QueryDBStoredQueries"
'100         RegistrySubKeys(EnumRegistrySubKeys.DataComboQueries) = "Software\Algoritmo\MRU Queries\DataComboStoredQueries"
110      Else
120         RegistrySubKeys(EnumRegistrySubKeys.MRUEmpresas) = "Software\Algoritmo\MRU Empresas\" & strUser
'130         RegistrySubKeys(EnumRegistrySubKeys.KeyMRUForms) = "Software\Algoritmo\MRU Forms\" & strUser
140      End If
      
150      Exit Sub
   
GestErr:
160      LoadError ErrorLog, "SetRegistryEntries" & Erl
170      ShowErrMsg ErrorLog
End Sub

Public Sub ReadSystemOptions()
      Dim vValue As Variant

         ' lectura de los parametros internos
   
10       On Error GoTo GestErr

20       ReDim SystemOptions(EnumSystemOptions.[_MAX_Value])
   
         '  CacheSize
30       vValue = GetKeyValuePI("ADO\CacheSize")
40       SystemOptions(EnumSystemOptions.iCacheSize) = IIf(IsNull(vValue), 1, vValue)
   
         '  Zoom
50       vValue = GetKeyValuePI("Opciones\Zoom Vista Previa\Valor Generico", 80)
60       SystemOptions(EnumSystemOptions.iZoom) = IIf(IsNull(vValue), 70, vValue)
   
         '  Fetch Mode
70       vValue = GetKeyValuePI("Performance\FetchMode")
80       SystemOptions(EnumSystemOptions.iFetchMode) = IIf(IsNull(vValue), 1, vValue)
90       If (SystemOptions(EnumSystemOptions.iFetchMode) <> alAsync) And SystemOptions(EnumSystemOptions.iFetchMode) <> alSync And (SystemOptions(EnumSystemOptions.iFetchMode) <> alTable) Then
100         MsgBox "El valor del parámetro 'Performance\FetchMode' admite los siguientes valores:" & vbCrLf & _
                   "    1 - Fetch Asincrónico" & vbCrLf & _
                   "    2 - Fetch Sincrónico" & vbCrLf & _
                   "    3 - Variable según 'Performance\Limite Fetch Sincrónico'" & vbCrLf & _
                   "En caso de omisión, asume la opción 2", vbInformation, App.ProductName
110      End If
   
120      If SystemOptions(EnumSystemOptions.iFetchMode) = alTable Then
130         vValue = GetKeyValuePI("Performance\Limite Fetch Sincronico")
140         SystemOptions(EnumSystemOptions.lngFetchLimit) = IIf(IsNull(vValue), 1000, vValue)
150      End If
   
   
         '  Fetch Mode Busqueda
160      vValue = GetKeyValuePI("Performance\FetchMode en Busqueda")
170      SystemOptions(EnumSystemOptions.iFetchModeSearch) = IIf(IsNull(vValue), 1, vValue)
180      If (SystemOptions(EnumSystemOptions.iFetchModeSearch) <> alAsync) And (SystemOptions(EnumSystemOptions.iFetchModeSearch) <> alSync) And (SystemOptions(EnumSystemOptions.iFetchModeSearch) <> alTable) Then
190         MsgBox "El valor del parámetro 'Performance\FetchMode' admite los siguientes valores:" & vbCrLf & _
                   "    1 - Fetch Asincrónico" & vbCrLf & _
                   "    2 - Fetch Sincrónico" & vbCrLf & _
                   "    3 - Variable según 'Performance\Limite Fetch Sincrónico'" & vbCrLf & _
                   "En caso de omisión, asume la opción 2", vbInformation, App.ProductName
200      End If
   
210      If SystemOptions(EnumSystemOptions.iFetchModeSearch) = alTable Then
220         vValue = GetKeyValuePI("Performance\Limite Fetch Sincronico en Busqueda")
230         SystemOptions(EnumSystemOptions.lngFetchLimitSearch) = IIf(IsNull(vValue), 300, vValue)
240      End If
   
   
         '  Usa copias locales
250      vValue = GetKeyValuePI("Performance\Usa Copias Locales")
260      SystemOptions(EnumSystemOptions.UseLocalCopy) = IIf(IsNull(vValue), Si, vValue)
270      If (SystemOptions(EnumSystemOptions.UseLocalCopy) <> Si) And (SystemOptions(EnumSystemOptions.UseLocalCopy) <> No) Then
280         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales' puede ser Sí o No:" & vbCrLf & _
                   "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
290      End If

         '  Usa copias locales en Búsquedas
300      vValue = GetKeyValuePI("Performance\Usa Copias Locales en Busqueda")
310      SystemOptions(EnumSystemOptions.UseLocalCopySearch) = IIf(IsNull(vValue), Si, vValue)
320      If (SystemOptions(EnumSystemOptions.UseLocalCopySearch) <> Si) And (SystemOptions(EnumSystemOptions.UseLocalCopySearch) <> No) Then
330         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales en Búsqueda' puede ser Sí o No:" & vbCrLf & _
                   "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
340      End If

         '  Pregunta si usa copias locales desactualizadas
350      vValue = GetKeyValuePI("Performance\Usa Copias Locales Desactualizadas")
360      SystemOptions(EnumSystemOptions.AskOldLocalCopy) = IIf(IsNull(vValue), Si, vValue)
370      If (SystemOptions(EnumSystemOptions.AskOldLocalCopy) <> Si) And (SystemOptions(EnumSystemOptions.AskOldLocalCopy) <> No) Then
380         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales Desactualizadas' puede ser Sí o No:" & vbCrLf & _
                   "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
390      End If
   
         '  Pregunta si usa copias locales desactualizadas
400      vValue = GetKeyValuePI("Opciones\Empresas\Usa MRU de Empresas")
410      SystemOptions(EnumSystemOptions.UseMRUEnterprise) = IIf(IsNull(vValue), Si, vValue)
420      If (SystemOptions(EnumSystemOptions.UseMRUEnterprise) <> Si) And (SystemOptions(EnumSystemOptions.UseMRUEnterprise) <> No) Then
430         MsgBox "El valor del parámetro 'Opciones\Empresas\Usa MRU de Empresas' puede ser Sí o No:" & vbCrLf & _
                   "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
440      End If
   
         '  Dimension de MRUForms
450      vValue = GetKeyValuePI("Performance\Dimension MRUForms")
460      SystemOptions(EnumSystemOptions.MaxMRUForms) = IIf(IsNull(vValue), 0, vValue)
   
470      Exit Sub

GestErr:
480      LoadError ErrorLog, "ReadSystemOptions " & Erl
490      ShowErrMsg ErrorLog
   
End Sub

Public Function SetApplication(ByVal objApp As Object) As Object

   If objApp Is Nothing Then Exit Function

   Set objApp.CurrentUser = CUsuario
   Set objApp.FormMDI = frmMDIInicio
   Set objApp.Menus = rstMenu
   Set objApp.ContextMenus = rstContextMenu
   Set objApp.CustomViews = rstVistasPersonalizadas
   Set objApp.ExportViews = rstVistasExportacion
   Set objApp.SysEnvironment = CSysEnvironment
   Set objApp.FormsMRU = MRUForms
   objApp.SystemOptionsProperty = SystemOptions
   objApp.RegistrySubKeysProperty = RegistrySubKeys

   Set SetApplication = objApp
   
End Function

Public Function IsMRUForm(ByVal lngHndW As Long) As Boolean
Dim frm As Form

   '  determina si un forms esta cargado en la colección MRUForms
   
   For Each frm In MRUForms

      If frm.hWnd = lngHndW Then IsMRUForm = True: Exit For

   Next frm
   
End Function

Public Function GetEjercicioActivo(ByVal strEmpresa As String) As String
Dim rst As ADODB.Recordset
   
   On Error GoTo GestErr
   
   Set rst = Fetch(strEmpresa, "SELECT EJERCICIOS.EJE_CODIGO FROM EJERCICIOS WHERE EJERCICIOS.EJE_ESTADO = 'V'")
   If Not rst.EOF Then
      GetEjercicioActivo = rst("EJE_CODIGO").Value
   End If
   
   If Not rst Is Nothing Then
      If rst.State <> adStateClosed Then rst.Close
   End If
   Set rst = Nothing
   
   Exit Function

GestErr:

   MsgBox "No es posible establecer una conexion con la Empresa " & strEmpresa & "." & vbCrLf & vbCrLf & _
          "Verifique en el Servidor si el Servicio OracleService" & strEmpresa & " ha sido iniciado." & vbCrLf & _
          "Asegurese que el Servicio este iniciado o modifique localmente su Registro para evitar el uso de dicha empresa"
          

   Err.Raise vbObjectError + 100, "GetEjercicioActivo [modMain]" & KNOWN_ERRORSOURCE, "Ejercicio Activo de la Empresa " & strEmpresa & " no disponible."
                              

End Function

Public Sub CrearRstEmpresas()
Dim sql As String

   sql = " SELECT EMPRESAS.* "
   sql = sql & " FROM EMPRESAS, "
   sql = sql & "     USUARIOS"
   sql = sql & " WHERE"
   sql = sql & "     USUARIOS.USU_USUARIO = '" & CUsuario.Usuario & "'"
   sql = sql & "     AND"
   sql = sql & "     ("
   sql = sql & "       (   USUARIOS.USU_PERMISO_EMPRESA = 'P' AND"
   sql = sql & "           EMPRESAS.EMP_CODIGO_EMPRESA  IN (  SELECT USUARIOS_EMPRESAS.UEM_EMPRESA"
   sql = sql & "                                          FROM USUARIOS_EMPRESAS"
   sql = sql & "                                          WHERE USUARIOS_EMPRESAS.UEM_USUARIO = '" & CUsuario.Usuario & "'"
   sql = sql & "                                              AND USUARIOS_EMPRESAS.UEM_EMPRESA = EMPRESAS.EMP_CODIGO_EMPRESA"
   sql = sql & "                                  )"
   sql = sql & "       ) OR"
   sql = sql & "       (   USUARIOS.USU_PERMISO_EMPRESA = 'D' AND"
   sql = sql & "           EMPRESAS.EMP_CODIGO_EMPRESA NOT IN (  SELECT USUARIOS_EMPRESAS.UEM_EMPRESA"
   sql = sql & "                                          FROM USUARIOS_EMPRESAS"
   sql = sql & "                                          WHERE USUARIOS_EMPRESAS.UEM_USUARIO = '" & CUsuario.Usuario & "'"
   sql = sql & "                                              AND USUARIOS_EMPRESAS.UEM_EMPRESA = EMPRESAS.EMP_CODIGO_EMPRESA"
   sql = sql & "                                  )"
   sql = sql & "       ) OR"
   sql = sql & "       (   NVL(USUARIOS.USU_PERMISO_EMPRESA,'N') = 'N' "
   sql = sql & "          "
   sql = sql & "           "
   sql = sql & "       )                                      "
   sql = sql & "     )"
   sql = sql & " ORDER BY EMP_DESCRIPCION    "
   Set rstEmpresas = Fetch(GetSPMProperty(DBSEmpresaPrimaria), sql, adOpenStatic, adLockReadOnly, adUseClient)
End Sub

Public Function GetMyObject(ByVal strComponentClass As String, Optional ByVal strServerName As String = NullString) As Object
10       On Error GoTo GestErr

         ' Sin este Objeto Local (que termina en Nothing) se queda vivo el objeto en el servidor
         Dim objetoLocal   As Object
         Dim ix As Integer
   
20       ix = 0
30       Set objetoLocal = CreateObject(strComponentClass, strServerName)
40       Set GetMyObject = objetoLocal
   
50       Set objetoLocal = Nothing
   
60       Exit Function

GestErr:
70       ix = ix + 1
80       If ix < 3 Then
90          Resume
100      End If
   
110      Set objetoLocal = Nothing
120      Set GetMyObject = Nothing

130      LoadError ErrorLog, "Objeto: " & strComponentClass & vbCrLf & "Servidor: " & strServerName
140      ShowErrMsg ErrorLog

End Function

