VERSION 5.00
Begin VB.Form StickyForm 
   BorderStyle     =   0  'None
   ClientHeight    =   3780
   ClientLeft      =   1080
   ClientTop       =   3015
   ClientWidth     =   4350
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   360
   End
   Begin VB.Menu mnuControl 
      Caption         =   "Example Menu"
      Begin VB.Menu mnuControlAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuControlSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControlClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "StickyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -[ StickyForm, by crown ]-
' Run this together with Explorer...

Private Sub Form_Load()
 StickyForm.Hide
 Visible = False
End Sub

Private Sub mnuControlAbout_Click()
MsgBox ("a small StickyForm Example...")
End Sub

Private Sub mnuControlClose_Click()
Unload StickyForm
End Sub

Private Sub Timer1_Timer()
On Error GoTo Err_Handler
 Dim sWindowText As String
 Dim r As Long
 Dim hwnd As Long
 Dim wp As WINDOWPLACEMENT
 
' change "Explorer" to whatever...
 hwnd = FindWindowWild("*Explorer*", False)
 
 wp.Length = Len(wp)
 r = GetWindowPlacement(hwnd, wp)
 sWindowText = Space(255)
 r = GetWindowText(hwnd, sWindowText, 255)
 sWindowText = Left(sWindowText, r)
    
 If wp.showCmd = 1 Then
    If Visible = False Then
       StickyForm.Show
       AppActivate sWindowText
       Visible = True
    End If
    
' window positions, demensions
 StickyForm.Move wp.rcNormalPosition.Right * 15 - 1300, (wp.rcNormalPosition.Top * 15) - 350, 1300, 320
    
    Else
       StickyForm.Hide
       Visible = False
 End If

Err_Handler:
End Sub

