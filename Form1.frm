VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCap 
   Caption         =   "Screen Capture The Inventive Way"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1200
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Picture"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdME 
      Caption         =   "Capture This Form "
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtPause 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "2"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   6255
      Left            =   2160
      ScaleHeight     =   6195
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   240
      Width           =   7215
   End
   Begin VB.CommandButton cmdActWin 
      Caption         =   "Capture Active Window"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Wait (seconds):"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frmCap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Capture Active Window The Inventive Way
'  This module is just another way to capture the screen but
'  it does not use the API calls needed!
'=========================================================================================
'  Created By: Behrooz Sangani
'  Published Date: 19/04/2002
'  Email:   bs20014@yahoo.com
'  WebSite: http://www.geocities.com/bs20014
'  Legal Copyright: Behrooz Sangani © 19/04/2002
'  Use and modify for free but keep the copyright!
'=========================================================================================
'  Description:
'  There are well known routines in VB to capture the screen
'  or a part of it. If you do not want to deal with advanced
'  Create Bitmap API calls and ... this is for you!
'  Set the capture timer and be sure you activate the window
'  you want to capture.
'  If you want to capture desktop click anywhere on the desktop
'  after timer starts.
'=========================================================================================
'  All comments and votes are welcome. Enjoy!

'Spare me one API call
'Unfortunately VB SendKeys doesn't send Print Screen else I would use that
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_SNAPSHOT = &H2C  'Snapshot button

Private Sub cmdActWin_Click() 'Capture Active Window
    
    'Hide me before capture
    Me.WindowState = vbMinimized
    Me.Hide
    
    'Wait
    Pause txtPause
    'Clear the clipboard
    Clipboard.Clear
    'Call Keyboard event
    Call keybd_event(VK_SNAPSHOT, 1, 0, 0)
    
    DoEvents  'required
    
    'Get the picture from clipboard
    Picture1.Picture = Clipboard.GetData()
    
    'Show me, task done!
    Me.Show
    Me.WindowState = vbNormal
    cmdSave.Enabled = True

End Sub

Private Sub cmdClear_Click() 'Clear picture box

    Set Picture1.Picture = Nothing
    cmdSave.Enabled = False
    
End Sub

Private Sub cmdME_Click()   'Capture me

    Clipboard.Clear
    Call keybd_event(VK_SNAPSHOT, 1, 0, 0)
    DoEvents
    Picture1.Picture = Clipboard.GetData()
    cmdSave.Enabled = True

End Sub

Private Sub cmdSave_Click()  'Save picture

    On Error GoTo Error
    'Common dialog save routine
    With CD1
        .DialogTitle = "Save Capture To..."
        .Filter = "Bitmap (*.bmp)|*.bmp"
        .CancelError = True
        .Flags = &H2  'Overwrite prompt
        .ShowSave
        If .FileName = "" Then GoTo Error
        'Save capture as bitmap
        SavePicture Picture1.Picture, .FileName
    End With
    
Error:   'Canceled

End Sub

Sub Pause(interval)  'Pause an interval

    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
    
End Sub


