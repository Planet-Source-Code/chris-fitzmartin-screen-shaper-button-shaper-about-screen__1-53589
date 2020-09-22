VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5580
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox pctFrame 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   675
      Left            =   5100
      ScaleHeight     =   675
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   3540
      Width           =   1275
      Begin VB.CommandButton cmdSysInfo 
         Height          =   585
         Left            =   0
         Picture         =   "frmAbout.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "View System Information"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdOK 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   600
         Picture         =   "frmAbout.frx":112C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Close This Screen"
         Top             =   60
         Width           =   615
      End
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
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3780
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
      Left            =   720
      TabIndex        =   5
      Top             =   3540
      Width           =   3390
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<Version>"
      Height          =   225
      Left            =   720
      TabIndex        =   4
      Top             =   1620
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "About <app>"
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
      Height          =   300
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<COMMENT>"
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   5940
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
  objShaper.MakeButtonUsingSelf cmdOK
  
  Me.Icon = Forms(0).Icon
  picIcon.picture = Me.Icon
  
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


Private Sub cmdOK_Click()

On Local Error GoTo ERROR_HANDLER

  Unload Me
Exit Sub
ERROR_HANDLER:
    HandleError Me.Name & " " & "Sub" & " " & "cmdOK_Click" & " Err# " & Err.Number & " " & Err.Description

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


