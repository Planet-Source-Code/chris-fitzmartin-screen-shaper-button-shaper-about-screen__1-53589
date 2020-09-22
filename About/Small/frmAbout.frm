VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   4260
      Picture         =   "frmAbout.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close This Screen"
      Top             =   2460
      Width           =   615
   End
   Begin VB.CommandButton cmdSysInfo 
      Height          =   585
      Left            =   3720
      Picture         =   "frmAbout.frx":0D26
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "View System Information"
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4260
      Picture         =   "frmAbout.frx":1B48
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      ToolTipText     =   "Application !"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Left            =   600
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "email . . ."
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Warning: © "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Copyright notice"
      Top             =   2460
      Width           =   3930
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<Version>"
      Height          =   225
      Left            =   600
      TabIndex        =   1
      Top             =   1500
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<COMMENT>"
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   3900
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSysInfo_Click()

On Local Error GoTo ERROR_HANDLER
 
  Dim sCommand As String
  sCommand = "compmgmt.msc "
  Call ShellExecute(0&, "open", sCommand, vbNullString, "", 1&)

Exit Sub
ERROR_HANDLER:
    HandleError Me.Name & " " & "Sub" & " " & "mnuComputerManager_Click" & " Err# " & Err.Number & " " & Err.Description

End Sub

Private Sub Form_Load()

On Local Error GoTo ERROR_HANDLER

  Dim objShaper As New Region
  objShaper.MakeFormUsingResource Me, "ABOUT"
  objShaper.MakeButtonUsingSelf cmdSysInfo
  objShaper.MakeButtonUsingSelf cmdCancel
  
  'Me.Icon = Forms(0).Icon
  'picIcon.picture = Forms(0).Icon
  
  lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  lblTitle = App.Title
  lblDisclaimer = App.LegalCopyright
  lblDescription = App.Comments
      
Exit Sub
ERROR_HANDLER:
    HandleError Me.Name & " " & "Sub" & " " & "Form_Load" & " Err# " & Err.Number & " " & Err.Description

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Me.Caption = "" Then FormMove Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' ESC to quit
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub pctShape_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseDown Button, Shift, x, y
End Sub

Sub HandleError(sMessage As String)
  Debug.Print sMessage
End Sub


Private Sub cmdCancel_Click()

On Local Error GoTo ERROR_HANDLER

  Unload Me
Exit Sub
ERROR_HANDLER:
    HandleError Me.Name & " " & "Sub" & " " & "cmdCancel_Click" & " Err# " & Err.Number & " " & Err.Description

End Sub

Private Sub lblContact_Click()

On Local Error GoTo ERROR_HANDLER

Const SW_SHOWNORMAL = 1

Dim sMessage As String
sMessage = "?subject=" & "From " & App.ProductName & " -- Feedback&body="
sMessage = Replace(sMessage, " ", "%20")

lblContact.ForeColor = &HFF00&
Screen.MousePointer = vbHourglass
DoEvents

Dim dReturn
Dim sCommand As String
sCommand = "mailto:chris_fitzmartin@yahoo.com " & sMessage
dReturn = ShellExecute(Me.hWnd, "open", sCommand, "", "C:\", SW_SHOWNORMAL)

Screen.MousePointer = vbNormal

Exit Sub
ERROR_HANDLER:
    'HandleError "frmAbout" & " " & "Sub" & " " & "lblContact_Click" & " Err# " & Err.Number & " " & Err.Description

End Sub

Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Form_MouseDown Button, Shift, x, y
End Sub
Private Sub lblDisclaimer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Form_MouseDown Button, Shift, x, y
End Sub
Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Form_MouseDown Button, Shift, x, y
End Sub
Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Form_MouseDown Button, Shift, x, y
End Sub
Private Sub pctFrame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Form_MouseDown Button, Shift, x, y
End Sub
Private Sub picIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Form_MouseDown Button, Shift, x, y
End Sub


