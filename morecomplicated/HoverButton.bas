Attribute VB_Name = "HoverButton"
'this example creates a timer instead of using an endless loop
'the result is an ressource-saving method
'(compare the cpu-usage in windows-taskmanager)


Option Explicit

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef _
lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint Lib "user32.dll" ( _
     ByVal xPoint As Long, _
     ByVal yPoint As Long) As Long

Private Declare Function SetTimer Lib "user32.dll" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long _
    ) As Long

Private Declare Function KillTimer Lib "user32.dll" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long _
    ) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim WhichButton As Object

Public Sub MakeHover(ByVal Button As Object)
    Const CheckInterval = 1 'ms     (- check every ms)
    
    If WhichButton Is Nothing Then
        Set WhichButton = Button
        SetEffect Button
        SetTimer Button.hWnd, 1, CheckInterval, AddressOf ReleaseHoverOrNot
    End If
End Sub


Private Sub ReleaseHoverOrNot()
    Dim PntApi As POINTAPI
    
    Do Until GetCursorPos(PntApi) <> 0: DoEvents: Loop
    'wait until Mouse isnt over the button...
    If Not WindowFromPoint(PntApi.X, PntApi.Y) = WhichButton.hWnd Then
        ResetEffect WhichButton
        KillTimer WhichButton.hWnd, 1
        Set WhichButton = Nothing
    End If
    
End Sub


'Feel free to change the following lines
'to create your individuel hoovereffect...

Private Sub SetEffect(Button As Object)
    'because command-buttons do not support forecolor...
    On Error Resume Next
    
    With Button
    
        .BackColor = RGB(200, 200, 255)
        .ForeColor = vbRed
        .FontBold = True
        .MouseIcon = LoadResPicture(101, vbResCursor)
        .MousePointer = 99
    
    End With
End Sub

Private Sub ResetEffect(Button As Object)
    'because command-buttons do not support forecolor...
    On Error Resume Next
    
    With Button
    
        .BackColor = vbMenuBar
        .ForeColor = vbMenuText
        .FontBold = False
        .MousePointer = 0
    
    End With
End Sub



