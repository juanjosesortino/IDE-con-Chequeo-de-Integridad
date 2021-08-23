VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMenu 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Menú Principal"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   4890
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   3330
      Picture         =   "frmMenu.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar siguiente"
      Top             =   75
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2955
      Picture         =   "frmMenu.frx":28D4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar opción"
      Top             =   75
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3990
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   1
      Top             =   90
      Width           =   72
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4170
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2A06
            Key             =   "LibroCerrado"
            Object.Tag             =   "LibroCerrado16x16"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2B60
            Key             =   "LibroAbierto"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2FB4
            Key             =   "Hoja"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sst1 
      Height          =   6210
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   10954
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   706
      BackColor       =   12632256
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMenu.frx":30B0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TreeView1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5175
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   570
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   9128
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4320
      Top             =   1980
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Image imgSplitter 
      BorderStyle     =   1  'Fixed Single
      Height          =   4785
      Left            =   4140
      MousePointer    =   9  'Size W E
      Top             =   120
      Width           =   150
   End
   Begin VB.Menu mnuContextMenu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuContextExpandirItem 
         Caption         =   "&Expandir"
      End
      Begin VB.Menu mnuContextContraerItem 
         Caption         =   "&Contraer"
      End
      Begin VB.Menu mnuContextContraerTodos 
         Caption         =   "Contraer &Todos"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditarAyuda 
         Caption         =   "Editar Ayuda..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextEjecutarItem 
         Caption         =   "&Ejecutar"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextBuscar 
         Caption         =   "&Buscar Opción"
         Shortcut        =   {F3}
      End
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
   Begin VB.Menu mnuWindow 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowItems 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuAyudaItem 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents frmForm    As Form
Attribute frmForm.VB_VarHelpID = -1

Public Function ShowForm(ByVal ControlData As Variant, _
                        ByVal strMenuKey As String, _
                        ByVal objApp As Object) As Long
   
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
   
10       On Error Resume Next

20       Set frmForm = Nothing
   
         'La busqueda en MRUForms la hago con la clave del menu
30       Set frmForm = MRUForms(Trim(strMenuKey))
40       If Not (frmForm Is Nothing) Then

             '/ es un form presente en MRUForms
             '
50          If frmForm.Visible = False And frmMDIInicio.MDITaskBar1.FormInTaskBar(frmForm.hWnd) = False Then
               'form MRU no visible y no presente en la barra
60             If Not frmForm Is Nothing Then frmForm.Visible = True
70             If Not frmForm Is Nothing Then frmForm.MenuKey = strMenuKey
80             If Not frmForm Is Nothing Then frmForm.ControlData = ControlData
         
90             If Not frmForm Is Nothing Then frmForm.InitForm
100            If Not frmForm Is Nothing Then frmForm.PostInitForm
110            If Not frmForm Is Nothing Then ShowForm = frmForm.hWnd
         
               'agrego el form en la TaskBar
120            If Not frmForm Is Nothing Then frmForm.MDIExtend1.WindowState = vbNormal
130            If Not frmForm Is Nothing Then frmMDIInicio.MDITaskBar1.AddFormToTaskBar frmForm
140            If Not frmForm Is Nothing Then
150               If Not frmMDIInicio.MDITaskBar1.FormInTaskBar(frmForm.hWnd) Then
160                  CenterMDIActiveXChild frmForm
170               End If
180            End If
         
200            Exit Function
210         End If
220      End If
   
230      On Error GoTo GestErr
   
240      Set frmForm = objApp.GetMDIChild(strMenuKey)
        
250      If Not frmForm Is Nothing Then frmForm.MenuKey = strMenuKey
260      If Not frmForm Is Nothing Then frmForm.ControlData = ControlData
   
270      If Not frmForm Is Nothing Then frmForm.InitForm
   
         '/ habilito el form
280      retval = EnableWindow(frmForm.hWnd, 1)

290      If Not frmForm Is Nothing Then frmForm.PostInitForm
   
310      If Not frmForm Is Nothing Then frmForm.ZOrder

320      If Not frmForm Is Nothing Then ShowForm = frmForm.hWnd
   
330      Screen.MousePointer = vbDefault
   
340      Exit Function
   
GestErr:
   
350      Screen.MousePointer = vbDefault
360      Err.Raise Err.Number, Err.source & vbCrLf & TypeName(Me) & ".ShowForm" & Erl, Err.Description
End Function

