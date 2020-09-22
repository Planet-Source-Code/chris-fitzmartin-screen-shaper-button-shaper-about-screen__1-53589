VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Region"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Const RGN_OR = 2

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private hndRegion As Long
Private DIB As cDIBSection
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
'vbPicTypeNone, vbPicTypeBitmap, vbPicTypeMetafile, vbPicTypeIcon, vbPicTypeEMetafile

Private Sub Class_Initialize()
  Set DIB = New cDIBSection
End Sub

Private Sub Class_Terminate()
  If hndRegion <> 0 Then Call DeleteObject(hndRegion)
End Sub

Private Sub ApplyRgn(ByVal hWnd As Long)

'-----------------------------------------------------------------------------
'   When the setWinowRgn function gets a handle to a region it applies and
'   deletes the region.  Therefore you cannot call the function twice with
'   the handle to the same region.  Therefore we make a copy of the region
'   for Windows to apply and delete.
'-----------------------------------------------------------------------------
  Dim hndRegionCopy As Long   'handle
  
  hndRegionCopy = CreateRectRgn(0, 0, 0, 0) ' Create a blank region

  Call CombineRgn(hndRegionCopy, hndRegion, hndRegionCopy, RGN_OR) ' Copy the region
  Call SetWindowRgn(hWnd, hndRegionCopy, True)
  
End Sub

Private Sub ScanPicture(ByVal picture As StdPicture, Optional transColor As Long = vbNull)
  Dim Rgn2 As Long
  
  hndRegion = CreateRectRgn(0&, 0&, 0&, 0&)
  
  Dim x As Long, y As Long
  Dim SPos As Long, EPos As Long
  Dim bDib() As Byte
  Dim tSA As SAFEARRAY2D
    
  'get the picture size of the form
  DIB.CreateFromPicture picture
  
  ' have the local matrix point to bitmap pixels
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DIB.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = DIB.BytesPerScanLine
    .pvData = DIB.DIBSectionBitsPtr
  End With
  
  Call CopyMemory(ByVal VarPtrArray(bDib), VarPtr(tSA), 4)
  
  ' if there is no transColor specified, use the first pixel as the transparent color
  If transColor = vbNull Then transColor = RGB(bDib(0, 0), bDib(1, 0), bDib(2, 0))
  
  For y = 0 To DIB.Height - 1 'line scan
      x = -3
      Do
          Rgn2 = 0
          x = x + 3
          
          While RGB(bDib(x, y), bDib(x + 1, y), bDib(x + 2, y)) = transColor And (x < DIB.Width * 3 - 3)
              x = x + 3 'skip the transparent point
          Wend
          
          SPos = x / 3
          
          While RGB(bDib(x, y), bDib(x + 1, y), bDib(x + 2, y)) <> transColor And (x < DIB.Width * 3 - 3)
              x = x + 3 'skip the nontransparent point
          Wend
          
          EPos = x / 3
          
          'combine the region
          If SPos <= EPos Then
              Rgn2 = CreateRectRgn(SPos, DIB.Height - y, EPos, DIB.Height - 1 - y)
              Call CombineRgn(hndRegion, hndRegion, Rgn2, RGN_OR)
              Call DeleteObject(Rgn2)
          End If
      Loop Until x >= DIB.Width * 3 - 3
  Next y
  
  Call CopyMemory(ByVal VarPtrArray(bDib), 0&, 4)
  
End Sub

Private Sub OffsetHeader(ByRef tmpForm As Object)
  With tmpForm
      If .BorderStyle <> 0 Then
          Dim xoff As Long, yoff As Long
          
          .ScaleMode = vbPixels
          
          xoff = (.ScaleX(.Width, vbTwips, vbPixels) - .ScaleWidth) / 2
          yoff = .ScaleY(.Height, vbTwips, vbPixels) - .ScaleHeight - xoff
  
          Call OffsetRgn(hndRegion, xoff, yoff)
      End If
  End With
  
End Sub

Public Sub MakeButtonUsingPictureBox(cmd As CommandButton, PIC As PictureBox)
'-------------------------------------------------------------------------
' The CommandButton's shape comes from a PictureBox.
'-------------------------------------------------------------------------
' Can load image into a "standard picture" from a file, resource or imagelist
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
If cmd.Style <> 1 Then Beep
If cmd.Style <> 1 Then Exit Sub
If cmd.picture <> 0 And PIC.picture = 0 Then Set PIC.picture = cmd.picture
Dim P As stdole.StdPicture
Set P = PIC.picture
If P.Type <> PictureTypeConstants.vbPicTypeBitmap Then Exit Sub    ' check for BMP
Set P = Nothing
'-------------------------------------------------------------------------

With PIC
  .BackColor = vbRed
  .AutoRedraw = False
  .AutoSize = True
  .Visible = False
  .BorderStyle = vbFixedSingle
End With

With cmd
  .Appearance = 0
  .BackColor = vbBlack
  .Caption = ""
  .DEFAULT = False
  .Width = PIC.Width
  .Height = PIC.Height
End With

Call ScanPicture(PIC.picture) ' get the bytes from the BMP (as StdPicture)
Call OffsetHeader(PIC)        ' adjust top/left starting point
Call ApplyRgn(cmd.hWnd)       ' set the shape to the object
 
Set PIC.picture = Nothing

End Sub

Public Sub MakePictureBox(PIC As PictureBox)

'-------------------------------------------------------------------------
' The PictureBox's shape comes from it's own .Picture property
'-------------------------------------------------------------------------
' can load into a standard picture from a file or resource or imagelist
'-------------------------------------------------------------------------
If PIC.picture = 0 Then Exit Sub  ' no pic was sent in ???
Dim P As stdole.StdPicture
Set P = PIC.picture
If P.Type <> vbPicTypeBitmap Then Exit Sub  ' check for BMP
Set P = Nothing
'-------------------------------------------------------------------------

With PIC
  .BackColor = vbRed
  .AutoRedraw = False
  .AutoSize = True
  .Visible = True
  .BorderStyle = vbBSNone
End With

Call ScanPicture(PIC.picture) ' get the bytes from the BMP (as StdPicture)
Call OffsetHeader(PIC)        ' Offset the Shape to allow for the border.
Call ApplyRgn(PIC.hWnd)       ' set the shape to the object

End Sub

Public Sub MakeFormUsingPictureBox(F As Form, PIC As PictureBox)

'-------------------------------------------------------------------------
' The form's shape comes from a PictureBox.
'-------------------------------------------------------------------------
' can load into a standard picture from a file, resource or imagelist
'-------------------------------------------------------------------------
If PIC.picture = 0 Then Exit Sub  ' no pic was sent in ???
Dim P As stdole.StdPicture
Set P = PIC.picture
If P.Type <> vbPicTypeBitmap Then Exit Sub  ' check for BMP
Set P = Nothing
'-------------------------------------------------------------------------
' To Do: set form's style to hide the titlebar, and setup form dragging.
'-------------------------------------------------------------------------
With PIC
  .BackColor = vbRed
  .AutoRedraw = False
  .AutoSize = True
  .Visible = False
  .BorderStyle = vbBSNone
End With

With F
  .Visible = False
  Set .picture = PIC.picture ' Set the pic
  If .ControlBox = True Then Beep
  If .BorderStyle <> vbBSNone Then Beep
  If .KeyPreview = False Then Beep
  If .Caption <> "" Then Beep
  If .picture = 0 Then Beep
  If .picture = 0 Then Exit Sub
End With

Call ScanPicture(PIC.picture) ' get the bytes of the BMP (as StdPicture)
Call OffsetHeader(PIC)        ' Offset the Shape to allow for the form header.
Call ApplyRgn(F.hWnd)         ' set the shape to the object

With F
  .Width = DIB.Width * Screen.TwipsPerPixelX
  .Height = DIB.Height * Screen.TwipsPerPixelY
  .Visible = True
End With

Set PIC.picture = Nothing

End Sub

Public Sub MakeFormUsingSelf(F As Form)

'-------------------------------------------------------------------------
' The form's shape comes from the Form's own .Picture property
'-------------------------------------------------------------------------
' can load into a standard picture from a file or resource or imagelist
'-------------------------------------------------------------------------
 Dim P As stdole.StdPicture
 Set P = F.picture
 If P.Type <> vbPicTypeBitmap Then Exit Sub   ' check for BMP
 
'-------------------------------------------------------------------------
' To Do: set form's style to hide the titlebar, and setup form dragging.
'-------------------------------------------------------------------------
With F
  .Visible = False
  If .ControlBox = True Then Beep
  If .BorderStyle <> vbBSNone Then Beep
  If .KeyPreview = False Then Beep
  If .Caption <> "" Then Beep
  If .picture = 0 Then Beep
  If .picture = 0 Then Exit Sub
End With

Call ScanPicture(F.picture)   ' get the bytes of the BMP (as StdPicture)
Call ApplyRgn(F.hWnd)         ' set the shape to the object

With F
  .Width = DIB.Width * Screen.TwipsPerPixelX
  .Height = DIB.Height * Screen.TwipsPerPixelY
  .Visible = True
End With

End Sub

Public Sub MakeFormUsingResource(F As Form, vResource As Variant)

'-------------------------------------------------------------------------
' The form's shape comes from a resource of type bitmap (~gif, ~ico, ~jpg)
'-------------------------------------------------------------------------
Set F.picture = LoadResPicture(vResource, vbResBitmap)
MakeFormUsingSelf F

End Sub

Public Sub MakeButtonUsingSelf(cmd As CommandButton)
'-------------------------------------------------------------------------
' The CommandButton's shape comes from itself.
'-------------------------------------------------------------------------
' can load image into a "standard picture" from a file or resource or imagelist
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
If cmd.picture = 0 Then Exit Sub
If cmd.Style <> 1 Then Beep
If cmd.Style <> 1 Then Exit Sub
Dim P As stdole.StdPicture
Set P = cmd.picture
If P.Type <> vbPicTypeBitmap Then Exit Sub    ' check for BMP
'-------------------------------------------------------------------------

Call ScanPicture(P)               ' get the bytes from the BMP (as StdPicture)
Call OffsetRgn(hndRegion, 2&, 1&) ' offset for a cmd button
Call ApplyRgn(cmd.hWnd)           ' set the shape to the object
 
With cmd
  .Appearance = 0
  .BackColor = vbBlack
  .Caption = ""
  .DEFAULT = False
  .Width = DIB.Width * Screen.TwipsPerPixelX + 20
  .Height = DIB.Height * Screen.TwipsPerPixelY + 20
  .UseMaskColor = True
  .MaskColor = vbRed
End With

Set P = Nothing

End Sub

Public Sub MakeButtonUsingSelf_O(opt As OptionButton)
'-------------------------------------------------------------------------
' The CommandButton's shape comes from itself.
'-------------------------------------------------------------------------
' can load image into a "standard picture" from a file or resource or imagelist
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
If opt.picture = 0 Then Exit Sub
If opt.Style <> 1 Then Beep
If opt.Style <> 1 Then Exit Sub
Dim P As stdole.StdPicture
Set P = opt.picture
If P.Type <> vbPicTypeBitmap Then Exit Sub    ' check for BMP
'-------------------------------------------------------------------------

Call ScanPicture(P)               ' get the bytes from the BMP (as StdPicture)
Call OffsetRgn(hndRegion, 2&, 0&) ' offset for a opt button
Call ApplyRgn(opt.hWnd)           ' set the shape to the object
 
With opt
  .Appearance = 0
  .BackColor = vbBlack
  .Caption = ""
  .ForeColor = &H80000010
  .Width = DIB.Width * Screen.TwipsPerPixelX + 20
  .Height = DIB.Height * Screen.TwipsPerPixelY + 20
End With

Set P = Nothing

End Sub

