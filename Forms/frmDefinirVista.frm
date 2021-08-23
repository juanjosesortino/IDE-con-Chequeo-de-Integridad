VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDefinirVista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personalizar Vista Lista"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "Anular Selección"
      Height          =   420
      Index           =   2
      Left            =   4410
      TabIndex        =   4
      Top             =   1530
      Width           =   1410
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Seleccionar Todo"
      Height          =   420
      Index           =   1
      Left            =   4410
      TabIndex        =   3
      Top             =   1035
      Width           =   1410
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Guardar Vista"
      Height          =   420
      Index           =   3
      Left            =   4410
      TabIndex        =   2
      Top             =   2385
      Width           =   1410
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   420
      Index           =   0
      Left            =   4410
      TabIndex        =   1
      Top             =   225
      Width           =   1410
   End
   Begin MSComctlLib.ListView lvwCampos 
      Height          =   3885
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   6853
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre de la Columna"
         Object.Width           =   6050
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DataField"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmDefinirVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ErrorLog  As ErrType                          'información del error generado

Private Values    As Collection    ' coleccion devuelta al llamador con todos los valores de los controles
Private itmX      As ListItem
Private mvarMenuKey As String

Private Const BUTTON_ACCEPT As Integer = 0
Private Const BUTTON_SELECT_ALL As Integer = 1
Private Const BUTTON_SELECT_CANCEL As Integer = 2
Private Const BUTTON_SAVE As Integer = 3

Private Sub cmdButton_Click(Index As Integer)

   On Error GoTo GestErr

   On Error GoTo GestErr

   Select Case Index
      Case BUTTON_ACCEPT
         
         If Not AlmenosUno() Then Exit Sub
         
         Set Values = New Collection
         For Each itmX In lvwCampos.ListItems
            Values.Add itmX.Checked, itmX.SubItems(1)
         Next itmX
      
         Me.Visible = False
      Case BUTTON_SELECT_ALL
         For Each itmX In lvwCampos.ListItems
            itmX.Checked = True
         Next itmX
      Case BUTTON_SELECT_CANCEL
         For Each itmX In lvwCampos.ListItems
            itmX.Checked = False
         Next itmX
      Case BUTTON_SAVE
         If Not AlmenosUno() Then Exit Sub
         GuardarVista
   End Select

   Exit Sub

GestErr:
   LoadError ErrorLog, "cmdButton_Click"
   ShowErrMsg ErrorLog

End Sub
Public Sub LoadListView(ByVal dtgGrilla As DataGrid)
      Dim ix As Integer
'Esta sub no se usa en ningun lugar, fue remplasada por una sub en clsABM1.LoadListView
'si se esta por romper compatibilidad se podria sacar ya que es Public
10       On Error GoTo GestErr

20       rstVistasPersonalizadas.Filter = "MenuKey = '" & mvarMenuKey & "'"
                                    
         'Carga de los campos en el ListView
30       If lvwCampos.ListItems.Count = 0 Then

40          For ix = 0 To dtgGrilla.Columns.Count - 1
50             If dtgGrilla.Columns(ix).Width > 0 Then
60                Set itmX = lvwCampos.ListItems.Add
70                itmX.Text = dtgGrilla.Columns(ix).Caption
80                itmX.SubItems(1) = dtgGrilla.Columns(ix).DataField & Format(ix, "000")

90                rstVistasPersonalizadas.Find "Column = '" & dtgGrilla.Columns(ix).DataField & Format(ix, "000") & "'", , adSearchForward, 1
100               If rstVistasPersonalizadas.EOF Then
110                  itmX.Checked = True
120               Else
130                  itmX.Checked = (rstVistasPersonalizadas("Visible") = Si)
140               End If

150               Values.Add itmX.Checked, itmX.SubItems(1)
160            End If
170         Next ix
180      End If

190      Exit Sub

GestErr:
200      LoadError ErrorLog, "LoadListView" & Erl
210      ShowErrMsg ErrorLog

End Sub
Function value(ByVal Key As String) As Variant

   On Error Resume Next
   value = Values.Item(Key)
    
End Function
Public Property Let MenuKey(ByVal vData As String)
   mvarMenuKey = vData
End Property

Private Sub GuardarVista()
   
   '
   ' borro la vieja definicion
   '
   On Error GoTo GestErr

   If rstVistasPersonalizadas.RecordCount > 0 Then rstVistasPersonalizadas.MoveFirst
   Do While Not rstVistasPersonalizadas.EOF
      rstVistasPersonalizadas.Delete
      rstVistasPersonalizadas.MoveNext
   Loop
   
   For Each itmX In lvwCampos.ListItems
   
      With rstVistasPersonalizadas
         
         .AddNew
         
         .Fields("User") = vbNullString
         .Fields("MenuKey") = mvarMenuKey
         .Fields("Column") = itmX.SubItems(1)
         .Fields("Visible") = IIf(itmX.Checked = True, Si, No)
         
         .Update
         
      End With
      
   Next itmX
   
   rstVistasPersonalizadas.Filter = adFilterNone
   
   If Dir(LocalPath, vbDirectory) = NullString Then MkDir LocalPath
   
   If Len(Dir(LocalPath & "VistasPersonalizadas")) > 0 Then Kill LocalPath & "VistasPersonalizadas"
   
   rstVistasPersonalizadas.Save LocalPath & "VistasPersonalizadas", adPersistADTG

   Exit Sub

GestErr:
   LoadError ErrorLog, "GuardarVista"
   ShowErrMsg ErrorLog
   
End Sub

Private Sub Form_Load()
   On Error GoTo GestErr

   Set Values = New Collection
   
   With ErrorLog
      .Form = Me.Name
      .Empresa = GetSPMProperty(DBSEmpresaPrimaria)
   End With

   Exit Sub

GestErr:
   LoadError ErrorLog, "Form_Load"
   ShowErrMsg ErrorLog
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Values = Nothing
   Set frmDefinirVista = Nothing
End Sub

Private Function AlmenosUno() As Boolean

   If Me.Visible = False Then Exit Function

   AlmenosUno = False
   
   For Each itmX In lvwCampos.ListItems
      If itmX.Checked Then AlmenosUno = True: Exit For
   Next itmX
   
   If Not AlmenosUno Then
      MsgBox "Debe seleccionar almenos una columna", vbInformation, App.ProductName
   End If

End Function
