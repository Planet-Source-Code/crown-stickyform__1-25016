Attribute VB_Name = "Sticky"
' Sticky.bas, by crown

' Thanks to: Pascal van de Wijdeven (HandleTreeCreator)
'            Arkadiy Olovyannikov   (FindWindowWild)


Dim sPattern As String, hFind As Long

Public Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Public Type POINTAPI
  x       As Long
  y       As Long
End Type

Public Type WINDOWPLACEMENT
  Length            As Long
  flags             As Long
  showCmd           As Long
  ptMinPosition     As POINTAPI
  ptMaxPosition     As POINTAPI
  rcNormalPosition  As RECT
End Type

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function IsWindowVisible& Lib "user32" (ByVal hwnd As Long)
Private Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
    

Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim k As Long, sName As String

    If IsWindowVisible(hwnd) And GetParent(hwnd) = 0 Then
        sName = Space$(128)
        k = GetWindowText(hwnd, sName, 128)
        
        If k > 0 Then
            sName = Left$(sName, k)
            If lParam = 0 Then sName = UCase(sName)
            
            If sName Like sPattern Then
                hFind = hwnd

EnumWinProc = 0
    Exit Function
End If
End If
End If

EnumWinProc = 1
End Function
'By: Arkadiy Olovyannikov
Public Function FindWindowWild(sWild As String, Optional bMatchCase As Boolean = True) As Long
    sPattern = sWild
    If Not bMatchCase Then sPattern = UCase(sPattern)

EnumWindows AddressOf EnumWinProc, bMatchCase
    FindWindowWild = hFind
End Function

