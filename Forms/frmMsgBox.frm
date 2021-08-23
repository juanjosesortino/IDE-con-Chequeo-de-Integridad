VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B97E3E11-CC61-11D3-95C0-00C0F0161F05}#163.0#0"; "ALGControls.ocx"
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caption"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDetalle 
      BackColor       =   &H80000004&
      Height          =   1635
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "frmMsgBox.frx":0000
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   585
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   4620
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4365
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":0006
            Key             =   "INFO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":0322
            Key             =   "QUESTION"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":063E
            Key             =   "WARNING"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBox.frx":095A
            Key             =   "ERROR"
         EndProperty
      EndProperty
   End
   Begin ALGControls.ALGLine ALGLine1 
      Height          =   45
      Left            =   45
      TabIndex        =   3
      Top             =   1695
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      X1              =   50
      X2              =   5115
      Y1              =   10
      Y2              =   10
   End
   Begin VB.PictureBox picSeverity 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   45
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   210
      Width           =   495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   2670
      TabIndex        =   1
      Top             =   1290
      Width           =   1215
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "D&etalles >>"
      Height          =   345
      Left            =   3975
      TabIndex        =   0
      Top             =   1290
      Width           =   1215
   End
   Begin VB.Label lblDescripcion 
      Height          =   240
      Left            =   585
      TabIndex        =   5
      Top             =   45
      Width           =   4650
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HEIGHT_NORMAL As Integer = 2150
Private Const HEIGHT_DETAILS As Integer = 4100

Public Enum eMsgSeverity
   Information = 1
   Question = 2
   Warning = 3
   Error = 4
End Enum

Private Sub cmdAceptar_Click()
   Unload Me
End Sub

Private Sub cmdDetalle_Click()

   Me.Height = IIf(cmdDetalle.Caption = "D&etalles >>", HEIGHT_DETAILS, HEIGHT_NORMAL)
   cmdDetalle.Caption = IIf(cmdDetalle.Caption = "D&etalles >>", "<< D&etalles", "D&etalles >>")
   
End Sub

Private Sub Form_Activate()
   Me.cmdAceptar.SetFocus
End Sub

Private Sub Form_Load()
   Me.Height = HEIGHT_NORMAL
   
   Me.Icon = LoadPicture(Icons & "Forms.ico")
   DoEvents
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmMsgBox = Nothing
End Sub

Public Sub ShowMsg(ByVal strTitle As String, _
                   ByVal strMsg As String, _
                   Optional ByVal strDetail As String, _
                   Optional ByVal iSeverity As eMsgSeverity = Information)

   Me.picSeverity.Picture = ImageList1.ListImages(iSeverity).Picture
   Me.Caption = strTitle
   If iSeverity = Error Then
      Me.lblDescripcion.Caption = "Se produjo el siguiente error:"
      Me.txtMsg.Font.Bold = True
      cmdDetalle_Click
   End If
   Me.txtMsg.Text = strMsg
   Me.txtDetalle.Text = strDetail
   
    
   
   Me.Show vbModal
   
End Sub

