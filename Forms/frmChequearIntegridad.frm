VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChequearIntegridad 
   Caption         =   "Chequear Integridad"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "frmChequearIntegridad"
   ScaleHeight     =   1665
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   1260
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   885
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   4080
      Begin VB.OptionButton optCompleto 
         Caption         =   "Chequeo &completo"
         Height          =   525
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1635
      End
      Begin VB.OptionButton optRapido 
         Caption         =   "Chequeo &Rápido"
         Height          =   525
         Left            =   210
         TabIndex        =   6
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame freButtons 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   4245
      TabIndex        =   0
      Top             =   60
      Width           =   1215
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   60
         TabIndex        =   2
         ToolTipText     =   "Cancela"
         Top             =   450
         Width           =   1125
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Acepta"
         Top             =   30
         Width           =   1125
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1350
      Top             =   0
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmChequearIntegridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmChequearIntegridad
' DateTime  : 14/07/2009 16:11
' Author    : Juan José Sortino
' Purpose   : Chequear la Integridad del sistema instanciando todos los objetos
'---------------------------------------------------------------------------------------
Option Explicit

Private ErrorLog            As ErrType

Private mvarControlData     As DataShare.udtControlData         'información de control

Dim intValor                As Integer
Dim strdll                  As String

Dim tliApp                  As Object
Dim objobjeto               As Object

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim bContinuarProceso       As Boolean

Private Sub cmdOK_Click()

10       On Error GoTo GestErr
   
20       bContinuarProceso = True
   
30       If optRapido Then
40          Chequeo_Rapido
50       Else
60          Chequeo_Completo
70       End If
   
80       stb1.Panels(STB_PANEL1).Text = NullString
90       stb1.Panels(STB_PANEL2).Text = NullString
   
100      ProgressBar1.Value = 0
   
         'Unload Me
   
110      Exit Sub

GestErr:
120      LoadError ErrorLog, "cmdOK_Click" & Erl
130      ShowErrMsg ErrorLog
  
End Sub

Private Sub cmdCancel_Click()

   bContinuarProceso = False
   
   stb1.Panels(STB_PANEL1).Text = NullString
   stb1.Panels(STB_PANEL2).Text = NullString
   
   'Unload Me
    
End Sub
Private Sub Chequeo_Rapido()

10       On Error GoTo GestErr
   
20       intValor = 0
   
         ' Cliente
30       CrearObjeto "BOFiscal.clsLibroIVA"
40       CrearObjeto "BOGeneral.clsCodigoPostal"
50       CrearObjeto "BOGesCom.clsClienteProveedor"
60       CrearObjeto "BOCereales.clsAnalisis"
70       CrearObjeto "BOContabilidad.clsAsiento"
80       CrearObjeto "BOSeguridad.clsUsuario"
90       CrearObjeto "BOProduccion.clsActividad"
   
100      CrearObjeto "Cereales.Application"
110      CrearObjeto "Fiscal.Application"
120      CrearObjeto "ReportsCereales.Application"
130      CrearObjeto "GestionComercial.Application"
140      CrearObjeto "Contabilidad.Application"
150      CrearObjeto "AdministradorGeneral.Application"
160      CrearObjeto "Seguridad.Application"
170      CrearObjeto "AlgStdFunc.clsStdFunctions"
180      CrearObjeto "Produccion.Application"
      '   CrearObjeto "PowerMaskControl.ocx"

190      CrearObjeto "ALGControls.clsControls"

         ' Server
200      CrearObjeto "DSGeneral.clsEntidadDS"
210      CrearObjeto "DSCereales.clsAlmacenajeDS"
220      CrearObjeto "DataAccess.clsDataAccess"
230      CrearObjeto "DSContabilidad.clsAsientoDS"
240      CrearObjeto "DSFiscal.clsLibroIvaDS"
250      CrearObjeto "DSGesCom.clsHistoricoTesoreriaDS"
260      CrearObjeto "DSProduccion.clsActividadCampoDS"
270      CrearObjeto "SPCereales.Liquidacion1116BC"

280      If bContinuarProceso = True Then
290         MsgBox "Chequeo de integridad Terminado"
300      Else
310         MsgBox "Chequeo de integridad Cancelado"
320      End If
   
330      Exit Sub

GestErr:
340      LoadError ErrorLog, "Chequeo_Rapido" & " --> " & strdll & " " & Erl
350      Set objobjeto = Nothing
360      ShowErrMsg ErrorLog
  
End Sub
Private Sub CrearObjeto(strdll As String)

10       On Error GoTo GestErr
   
20       If bContinuarProceso = False Then Exit Sub
   
30       MostrarAvance Left(strdll, InStr(strdll, ".") - 1) & ".dll"
   
40       Set objobjeto = CreateObject(strdll)
50       Set objobjeto = Nothing
   
60       Exit Sub

GestErr:
70       LoadError ErrorLog, "CrearObjeto" & " --> " & strdll & " " & Erl
80       Set objobjeto = Nothing
90       ShowErrMsg ErrorLog
End Sub
Private Sub Chequeo_Completo()

   On Error GoTo GestErr
   
   Set tliApp = CreateObject("TLI.TLIApplication")
   
   ChequearForms "Fiscal" '34
   ChequearForms "AdministradorGeneral" '26
   ChequearForms "Contabilidad" '28
   ChequearForms "Seguridad" '3
   ChequearForms "Produccion" '28
   ChequearForms "GesCom" '107
   ChequearForms "Cereales" '191
'ReportsCereales.dll
'COMFiscalPrinter.dll
'AlgStdFunc.dll

'   ChequearBO "BOFiscal"
'   ChequearBO "BOGeneral"
'   ChequearBO "BOContabilidad"
'   ChequearBO "BOSeguridad"
'   ChequearBO "BOProduccion"
'   ChequearBO "BOGesCom"
'   ChequearBO "BOCereales"
   Set tliApp = Nothing

'   ChequearDS

   If bContinuarProceso = True Then
      MsgBox "Chequeo de integridad Terminado"
   Else
      MsgBox "Chequeo de integridad Cancelado"
   End If
   
   Exit Sub

GestErr:
   LoadError ErrorLog, "Chequeo_Completo" & Erl
   ShowErrMsg ErrorLog
   
End Sub

Private Sub ChequearForms(strdll As String)
'Error: cannot create mdi child window heap memory
'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\SubSystems
'%SystemRoot%\system32\csrss.exe ObjectDirectory=\Windows SharedSection=1024,3072,512 Windows=On SubSystemType=Windows ServerDll=basesrv,1 ServerDll=winsrv:UserServerDllInitialization,3 ServerDll=winsrv:ConServerDllInitialization,2 ProfileControl=Off MaxRequestThreads=16

   Dim rst       As ADODB.Recordset
   Dim sql       As String
   Dim hWndAdmin As Long
   Dim ix        As Long
   Dim objForms  As Object
   Dim ObjForm   As Form
   
   On Error GoTo GestErr
   
   sql = "SELECT MAX(MNU_CLAVE) MNU_CLAVE, MNU_NOMBRE_EXEC "
   sql = sql & " FROM MENU "
   sql = sql & " WHERE MNU_MODULO = '%1' "
   sql = sql & " AND MNU_NOMBRE_EXEC IS NOT NULL "
   sql = sql & " GROUP BY MNU_NOMBRE_EXEC "
   
   Select Case strdll
      Case "Fiscal"
         sql = Replace(sql, "%1", "FISCAL")
      Case "AdministradorGeneral"
         sql = Replace(sql, "%1", "ADMINISTRADORGENERAL")
      Case "Contabilidad"
         sql = Replace(sql, "%1", "CONTABILIDAD")
      Case "Seguridad"
         sql = Replace(sql, "%1", "SEGURIDAD")
      Case "Produccion"
         sql = Replace(sql, "%1", "PRODUCCION")
      Case "GesCom"
         sql = Replace(sql, "%1", "GESTIONCOMERCIAL")
      Case "Cereales"
         sql = Replace(sql, "%1", "CEREALES")
   End Select
   
   Set rst = Fetch(GetSPMProperty(DBSEmpresaPrimaria), sql)
   If rst.RecordCount > 0 Then
      Set objGesCom = mvarMDIForm.GetInstance(strdll)
      
      ProgressBar1.Value = 0
      ProgressBar1.Min = 0
      ProgressBar1.Max = rst.RecordCount

      Do While Not rst.EOF And bContinuarProceso = True
         'Sleep 5000
         stb1.Panels(STB_PANEL1).Text = "Formularios de " & strdll & ".dll"
         stb1.Panels(STB_PANEL2).Text = rst("MNU_NOMBRE_EXEC").Value
         
         'Me.Hide
         hWndAdmin = mvarMDIForm.CallAdmin(rst("MNU_CLAVE").Value, mvarControlData, True)
         DoEvents
         'Me.Show
         
         Set objForms = objGesCom.CollectionForms
         DoEvents

         For ix = 0 To objForms.Count - 1
            If objForms(ix).hWnd = hWndAdmin Then
               Set ObjForm = objForms(ix)
               Unload ObjForm
               Set ObjForm = Nothing
               Exit For
            End If
         Next
         Set objForms = Nothing
            
         ProgressBar1.Value = ProgressBar1.Value + 1
         rst.MoveNext
      Loop
   End If
   
   Set objGesCom = Nothing
   Set objForms = Nothing
   Set ObjForm = Nothing
   If Not rst Is Nothing Then
      If rst.State <> adStateClosed Then rst.Close
   End If
   Set rst = Nothing
      
   mvarMDIForm.Largar
   '? frmMDIInicio.Count
   '21
   
   Exit Sub

GestErr:
   LoadError ErrorLog, "ChequearForms " & Erl & vbCrLf & strdll & "." & objForms(ix).Caption
   
   Set objGesCom = Nothing
   Set objForms = Nothing
   Set ObjForm = Nothing
   If Not rst Is Nothing Then
      If rst.State <> adStateClosed Then rst.Close
   End If
   Set rst = Nothing
   
   ShowErrMsg ErrorLog
End Sub
Private Sub ChequearDS()
   
         Dim objCatlog           As Object
         Dim objApplications     As Object
         Dim objApplication      As Object
         Dim objComponents       As Object
         Dim objComponentItem    As Object
   
10       On Error GoTo GestErr
   
20       If bContinuarProceso = False Then Exit Sub
   
30       stb1.Panels(STB_PANEL1).Text = "Esperando COM+ ..."
40       stb1.Panels(STB_PANEL2).Text = ""
   
50       Set objCatlog = CreateObject("COMAdmin.COMAdminCatalog")
60       objCatlog.Connect ("")
70       Set objApplications = objCatlog.GetCollection("Applications")
80       objApplications.Populate
90       For Each objApplication In objApplications
100          Set objComponents = objApplications.GetCollection("Components", objApplication.Key)
110          objComponents.Populate

120          ProgressBar1.Value = 0
130          ProgressBar1.Min = 0
140          ProgressBar1.Max = objComponents.Count
 
150          If objApplication.Name = "Algoritmo" Then
160             For Each objComponentItem In objComponents
170                 Set objobjeto = CreateObject(objComponentItem.Name)
180                 Set objobjeto = Nothing
  
190                 DoEvents
  
200                 stb1.Panels(STB_PANEL1).Text = objApplication.Name & " COM+"
210                 stb1.Panels(STB_PANEL2).Text = objComponentItem.Name
  
220                 ProgressBar1.Value = ProgressBar1.Value + 1
230             Next
240         End If
250      Next

260      Set objCatlog = Nothing
270      Set objApplications = Nothing
280      Set objApplication = Nothing
290      Set objComponents = Nothing
300      Set objComponentItem = Nothing
   
310      Exit Sub

GestErr:
320      LoadError ErrorLog, "ChequearDS " & Erl & vbCrLf & objComponentItem.Name

330      Set objCatlog = Nothing
340      Set objApplications = Nothing
350      Set objApplication = Nothing
360      Set objComponents = Nothing
370      Set objComponentItem = Nothing
380      Set objobjeto = Nothing
   
390      ShowErrMsg ErrorLog
End Sub

Private Sub ChequearBO(strdll As String)
   
10       On Error GoTo GestErr
   
         Dim tlibi     As Object
         Dim ti        As Object
         Dim objobjeto As Object
   
20       If bContinuarProceso = False Then Exit Sub
   
30       Set tlibi = tliApp.TypeLibInfoFromFile("C:\Archivos de programa\Algoritmo\" & strdll & ".dll")
   
40       ProgressBar1.Value = 0
50       ProgressBar1.Min = 0
60       ProgressBar1.Max = tlibi.TypeInfos.Count
   
70       For Each ti In tlibi.TypeInfos
80          If ti.AttributeMask = 2 Then
90             If Len(ti.Name) > 0 Then
100               Set objobjeto = CreateObject(strdll & "." & ti.Name)
110               Set objobjeto = Nothing

120               stb1.Panels(STB_PANEL1).Text = strdll
130               stb1.Panels(STB_PANEL2).Text = ti.Name

140               DoEvents
150            End If
160         End If
170         ProgressBar1.Value = ProgressBar1.Value + 1
180      Next
   
190      Set tlibi = Nothing
200      Set ti = Nothing
210      Set objobjeto = Nothing

220      Exit Sub

GestErr:
230      LoadError ErrorLog, "ChequearBO " & Erl & vbCrLf & strdll & "." & ti.Name & "Nro: " & ProgressBar1.Value

240      Set tlibi = Nothing
   Set ti = Nothing
   Set objobjeto = Nothing
   
   ShowErrMsg ErrorLog
End Sub
Private Sub MostrarAvance(strlocaldll As String)

   DoEvents
   
   strdll = strlocaldll
   
   stb1.Panels(STB_PANEL1).Text = strlocaldll
   stb1.Panels(STB_PANEL2).Text = ""
      
   ProgressBar1.Value = intValor
   
   intValor = intValor + 1
End Sub
Private Sub Form_Load()

   On Error GoTo GestErr

   optRapido = True
   intValor = 0
   
   With ProgressBar1
      .Min = 0
      .Max = 25
   End With
   DoEvents

   Exit Sub

GestErr:
   LoadError ErrorLog, "Form_Load" & Erl
   ShowErrMsg ErrorLog
End Sub
 
Public Property Let ControlData(ByVal vData As Variant)
    mvarControlData = vData
    
   With ErrorLog
      .Form = Me.Name
      .Empresa = frmMenu.TreeView1(frmMenu.sst1.TabIndex).Tag
   End With
    
End Property

Public Property Get ControlData() As Variant
    ControlData = mvarControlData
End Property
