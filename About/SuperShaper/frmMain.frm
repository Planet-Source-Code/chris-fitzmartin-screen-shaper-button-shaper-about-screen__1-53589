VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   900
      Width           =   435
   End
   Begin VB.PictureBox pctShape 
      Height          =   495
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdLoadPicture 
      Caption         =   "PIC"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   900
      Width           =   435
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'----------------------------------------------
' You really only need a few lines of code . . .
'----------------------------------------------
On Error Resume Next

If Not FileExists(App.Path & "\" & "SUPER_SHAPER.RES") Then
  MsgBox "file not found: SUPER_SHAPER.RES" & vbCrLf & vbCrLf & "Please check your VB project"
  FileCopy "C:\Program Files\Microsoft Visual Studio\VB98\Template\Projects\SUPER_SHAPER.RES", App.Path & "\" & "SUPER_SHAPER.RES"
  Exit Sub
End If

Dim R As New Region
R.MakeFormUsingResource Me, "HOMER"
'R.MakeFormUsingPictureBox Me, pctShape
'R.MakeButtonUsingSelf cmdLoadPicture   'put a pic in the button
'----------------------------------------------

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

Private Sub cmdLoadPicture_Click()

On Local Error GoTo ERROR_HANDLER

With CommonDialog1
  .CancelError = True
  .DefaultExt = "*.bmp;*.gif;*.jpg;*.jpeg"
  .DialogTitle = ""
  .Filter = "Pictures (*.bmp;*.gif;*.jpg;*.jpeg)|*.bmp;*.gif;*.jpg;*.jpeg"
  .InitDir = ""
  .ShowOpen
  If CommonDialog1.FileName = "" Then Exit Sub
  pctShape.picture = LoadPicture(CommonDialog1.FileName)
End With

'----------------------------------------------
' You really only need 2 lines of code
Dim R As New Region
R.MakeFormUsingPictureBox Me, pctShape
'----------------------------------------------

Exit Sub
ERROR_HANDLER:
  HandleError Me.Name & vbTab & "cmdLoadPicture_Click" & vbTab & Err.Number & vbTab & Err.Description

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Sub HandleError(sMessage)
  Debug.Print sMessage
End Sub

Public Function FileExists(sFileName As String) As Boolean

On Local Error GoTo ERROR_HANDLER


On Local Error Resume Next
Dim sResult As String

If Trim$(sFileName) <> "" Then
  sResult = Dir(sFileName)
  If sResult <> "" Then FileExists = True
  
  If FileExists Then
    If GetAttr(sFileName) And vbDirectory Then FileExists = False
  End If
End If


Exit Function
ERROR_HANDLER:
    HandleError "basGlobal" & " " & "Function" & " " & "FileExists" & " Err# " & Err.Number & " " & Err.Description

End Function

