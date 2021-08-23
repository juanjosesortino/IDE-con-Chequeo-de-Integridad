VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freButtons 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton cmdOK 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   60
         TabIndex        =   2
         ToolTipText     =   "Acepta"
         Top             =   30
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "Cancela"
         Top             =   450
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   4080
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   795
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   255
         Visible         =   0   'False
         Width           =   645
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu mnuContextMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuContextItem 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmDialog
' DateTime  : 06/05/2004 10:51
' Author    : tony
' Purpose   : Esta versión del Dialogo es solo para Inicio
'---------------------------------------------------------------------------------------
Option Explicit

Private Const FRAME_LEFT = 100                          ' left del frame
Private Const FRAME_DISTANCE = 100                      ' distancia entre frames
'Private Const OPTION_LEFT = 200                         ' valor del Left para todos los controles check, option dentro de un frame
'Private Const OPTION_DISTANCE = 50                      ' distancia entre options/checks

'  posicion de los datos en coleccion collSearch
'Private Const TABLE_FIELD = 0
'Private Const CONTROL_LIST = 1
'Private Const FIELD_LIST = 2
'Private Const LABEL_CONTROL = 3
'Private Const LABEL_FIELDNAME = 4

Private ErrorLog                As ErrType

Private nextBtnTop               As Integer            ' es el Top para proximo control
Private maxWidth                 As Integer            ' es el ancho del frame
'Private TabIndexCounter          As Integer           ' es el contador del TabIndex

Private Values                   As New Collection    ' coleccion devuelta al llamador con todos los valores de los controles
Private fi                       As New Collection    ' coleccion con FieldInformatio

Private mvarControlData          As DataShare.udtControlData         'información de control

Private aTableProperties         As Variant

'Private aKeys()                  As Variant
'Private aArray()                 As Variant
'Private vValue                   As Variant

' definicion de los eventos del form
Public Event ValidateDialog(ByRef Response As String)
Public Event LostFocus(ByVal ControlKey As String)
Public Event ButtonClick(ByVal StrButton As String)

Public Enum eDlgAlignText
   AlineaDerecha = 0
   AlineaInferior = 1
End Enum
'Private bSetDefault              As Boolean

Private Sub AddFrame(ByVal Key As String, ByVal Caption As String)
                    
Dim frameIndex    As Integer
Dim thisFrame     As Frame
Dim prevFrame     As Frame
    
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   On Error GoTo GestErr
   
   frameIndex = Frame1.UBound + 1
   
   ' cargo un nuevo frame
   Load Frame1(frameIndex)
   Set thisFrame = Frame1(frameIndex)
   
   ' setea las propiedades del frame
   thisFrame.Caption = Caption
   thisFrame.Tag = Key
   thisFrame.Visible = True
   
   thisFrame.Width = 0
   thisFrame.Height = 0
   
   maxWidth = 0  ' ***
   
   nextBtnTop = 200
   
   'obtengo el frame anterior
   Set prevFrame = Frame1(frameIndex - 1)
   
   thisFrame.Move FRAME_LEFT, FRAME_DISTANCE, thisFrame.Width, thisFrame.Height
         
   Exit Sub

GestErr:
   LoadError ErrorLog, "AddFrame"
   ShowErrMsg ErrorLog
    
End Sub

Public Sub AddText(ByVal Key As String, ByVal Caption As String)
                   
Dim frameIndex          As Integer
Dim thisText            As TextBox
Dim thisLabel           As Label
Dim thisFrame           As Frame
Dim iDimension           As Integer

   ' agrego un textbox al grupo actual
   
   On Error GoTo GestErr

   aTableProperties = GetFieldInformation(Key)

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   Load Label1(Label1.UBound + 1)
   
   Set thisLabel = Label1(Label1.UBound)
   Set thisLabel.Container = thisFrame
   thisLabel.Caption = Caption & ":"
   thisLabel.Width = TextWidth(thisLabel.Caption)
   thisLabel.Move 100, nextBtnTop + 40
   thisLabel.Visible = True
   thisLabel.Tag = TypeName(thisLabel) & Key
      
   Load Text2(Text2.UBound + 1)
   Set thisText = Text2(Text2.UBound)
   
   thisText.Height = 315
   
   Set thisText.Container = thisFrame
   
   Select Case FieldProperty(aTableProperties, Key, dsTipoDato)
      Case "135"
         thisText.Alignment = vbLeftJustify
      Case "131"
         thisText.Alignment = vbRightJustify
      Case Else
         thisText.Alignment = vbLeftJustify
   End Select
      
   thisText.Tag = TypeName(thisText) & Key
   thisText.Visible = True
   
    thisText.Move 1500, nextBtnTop
   
   ' calculo la posicion del proximo control
   nextBtnTop = nextBtnTop + thisText.Height + 200
   
   fi.Add aTableProperties, TypeName(thisText) & Key
   
   iDimension = FieldProperty(aTableProperties, Key, dsDimension)
   If iDimension = 0 Then iDimension = 10
   thisText.Width = iDimension * TextWidth("Z") + 200
   thisText.MaxLength = iDimension
   '
   '  unifico los anchos por rango
   '
   Select Case thisText.Width
      Case Is <= 500
         thisText.Width = 500
      Case Is <= 800
         thisText.Width = 800
      Case Is <= 1300
         thisText.Width = 1300
      Case Is <= 1500
         thisText.Width = 1500
   End Select
      
   Exit Sub

GestErr:
   LoadError ErrorLog, "AddText"
   ShowErrMsg ErrorLog
   
End Sub

Function Value(ByVal Key As String) As Variant
   On Error Resume Next
   
   Value = Values.Item(Key)
   
End Function

Private Sub cmdOK_Click()
Dim Response As String
   
   'salvo las propiedades
   SaveProperties
      
   Values.Add "OK", "ButtonPressed"
      
   'dejo que el objeto llamador valide los datos del form
   RaiseEvent ValidateDialog(Response)
   
   If Len(Response) > 0 Then
   
      MsgBox Response, vbExclamation, Me.Caption
      
      Set Values = Nothing
      
      Exit Sub
   End If
   
   'informo al objeto llamado que los datos han sido aceptados
   RaiseEvent ButtonClick(cmdOK.Caption)
   
   Unload Me
   
End Sub

Private Sub cmdCancel_Click()

   Set Values = New Collection

   Values.Add "Cancel", "ButtonPressed"

   RaiseEvent ButtonClick(cmdCancel.Caption)
   Unload Me
    
End Sub

Public Sub ShowDialog(Optional ByVal ShowModal As FormShowConstants = vbModal)
Dim Ctrl       As Control
Dim iMaxLabelWidth     As Integer
Dim iMaxTextWidth     As Integer

   On Error Resume Next
   
   For Each Ctrl In Me.Controls
      
      If TypeOf Ctrl Is Label Then
         If Ctrl.Width > iMaxLabelWidth Then iMaxLabelWidth = Ctrl.Width
      End If
      
   Next Ctrl
   
   For Each Ctrl In Me.Controls
      
      If TypeOf Ctrl Is TextBox Then
         Ctrl.Left = iMaxLabelWidth + 300
         If Ctrl.Left + Ctrl.Width > iMaxTextWidth Then
            iMaxTextWidth = Ctrl.Left + Ctrl.Width
         End If
      End If
      
   Next Ctrl
   
   Frame1(Frame1.UBound).Width = iMaxTextWidth + 300
   Frame1(Frame1.UBound).Height = nextBtnTop
   
   Me.freButtons.Left = Frame1(Frame1.UBound).Left + Frame1(Frame1.UBound).Width + 550
   
   Me.Width = Me.freButtons.Left + Me.freButtons.Width + 300
   Me.Height = Frame1(Frame1.UBound).Height + 800
   
   'el dialogo es siempre modal (y no es MDIChild)
   Me.Show vbModal
   
End Sub

Private Sub Form_Initialize()
   
   On Error GoTo GestErr
   
   nextBtnTop = 200
   
   AddFrame "fre1", ""
   
   Exit Sub

GestErr:
   LoadError ErrorLog, "Form_Initialize"
   ShowErrMsg ErrorLog
   
End Sub

Private Sub Form_Load()

   ' reseteo todos los valores
   On Error GoTo GestErr

   Set Values = Nothing
   
   Me.Icon = LoadPicture(Icons & "Forms.ico")
   DoEvents
   
   DoEvents

   Exit Sub

GestErr:
   LoadError ErrorLog, "Form_Load"
   ShowErrMsg ErrorLog
End Sub

Private Sub Form_Terminate()

   Set Values = Nothing
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmDialog = Nothing
   
End Sub

Public Function GetControl(ByVal Key As String) As Control
Dim cntl As Control

   On Error GoTo GestErr

   For Each cntl In Me.Controls
      If UCase(cntl.Tag) = UCase(Key) Then
         Set GetControl = cntl
         Exit For
      End If
   Next cntl

   Exit Function

GestErr:
   LoadError ErrorLog, "GetControl"
   ShowErrMsg ErrorLog

End Function

Private Sub SaveProperties()
Dim Ctrl As Control

   ' salvo las propiedades
    
   On Error Resume Next
   
   Set Values = New Collection
   
   For Each Ctrl In Controls
       If Ctrl.Visible = True Then
       
         If Ctrl.Tag <> "" Then
         
            Values.Add Ctrl.Text, Ctrl.Tag
            
         End If
      End If
   Next
      

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

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
Dim thisText As TextBox
Dim Key As String
Dim StrCase As String
Dim iTipoDato As Integer

   Set thisText = Text2(Index)
      
   Key = Replace(thisText.Tag, "TextBox", "")
      
   aTableProperties = GetFieldInformation(Key)
   
   StrCase = FieldProperty(aTableProperties, Key, dsCaseMode)
   iTipoDato = FieldProperty(aTableProperties, Key, dsTipoDato)
   
   Select Case StrCase
      Case "U"
         KeyAscii = Asc(UCase(Chr(KeyAscii)))
      Case "L"
         KeyAscii = Asc(LCase(Chr(KeyAscii)))
   End Select
   
   Select Case iTipoDato
      Case 131
         Select Case KeyAscii
            Case vbKeyBack, Asc("+"), Asc("-"), Asc(DecimalChar)
            Case Else
               If Not IsNumeric(Chr(KeyAscii)) Then
                  KeyAscii = 0
               End If
         End Select
   End Select
      

End Sub

Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
Dim thisText As TextBox
Dim Key As String
Dim iTipoDato As Integer

   Set thisText = Text2(Index)
   
   If thisText.Text = "" Then Exit Sub
      
   Key = Replace(thisText.Tag, "TextBox", "")
      
   aTableProperties = GetFieldInformation(Key)
   
   iTipoDato = FieldProperty(aTableProperties, Key, dsTipoDato)
   
   If iTipoDato = 135 Then
   
      thisText.Text = IIf(thisText.Text = "", "", Format(thisText.Text, , , vbUseSystem))
   
      If Not IsDate(thisText.Text) Then
         MsgBox "Ingrese una fecha válida", vbInformation, App.ProductName
         Cancel = True
      End If
      
   End If

End Sub
