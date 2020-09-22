Attribute VB_Name = "Module1"
Option Explicit
Type RECT
    Left As Long
    Top As Long
    Height As Long
    Width As Long
End Type

Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

Public ButtonsOn As Boolean
Public hRectRgn(7) As Long
Public rectangle(7) As RECT
Public sMenuButtonText(7) As String
Public sGraphics(7) As String
Public sRollOvr(7) As String



Public Sub Main()
    
    Dim intCount As Integer
    'Initialize main menu option "button" rectangles
    rectangle(0).Top = 1315
    rectangle(1).Top = 2350
    rectangle(2).Top = 2940
    rectangle(3).Top = 3520
    rectangle(4).Top = 4170
    rectangle(5).Top = 4820
    rectangle(6).Top = 6190
    
    For intCount = 0 To 6
        rectangle(intCount).Left = 160
        rectangle(intCount).Width = 1330
        rectangle(intCount).Height = 300
        hRectRgn(intCount) = CreateRectRgn(rectangle(intCount).Left, _
                                            rectangle(intCount).Top, _
                                            (rectangle(intCount).Left + rectangle(intCount).Width), _
                                            (rectangle(intCount).Top + rectangle(intCount).Height))

                                        
    Next intCount
    
   
    
    'Initialize main menu option "button" tooltip_label1 text
    sMenuButtonText(0) = "This is a Rollover Menu Program"
    sMenuButtonText(1) = "This takes you to Form1"
    sMenuButtonText(2) = "This takes you to Form2"
    sMenuButtonText(3) = "This takes you to Form3"
    sMenuButtonText(4) = "This takes you to Form4"
    sMenuButtonText(5) = "This takes you to Form5"
    sMenuButtonText(6) = "This takes you to the Exit"
    
    'Initilize graphics on button's
    sGraphics(0) = App.Path & "\buttons\abutton.bmp"
    sGraphics(1) = App.Path & "\buttons\f1.bmp"
    sGraphics(2) = App.Path & "\buttons\f2.bmp"
    sGraphics(3) = App.Path & "\buttons\f3.bmp"
    sGraphics(4) = App.Path & "\buttons\f4.bmp"
    sGraphics(5) = App.Path & "\buttons\f5.bmp"
    sGraphics(6) = App.Path & "\buttons\exit.bmp"
    
    'Initilize graphics for rollovers
    sRollOvr(0) = App.Path & "\graphics\globe.bmp"
    sRollOvr(1) = App.Path & "\graphics\pic20.bmp"
    sRollOvr(2) = App.Path & "\graphics\globe.bmp"
    sRollOvr(3) = App.Path & "\graphics\pic20.bmp"
    sRollOvr(4) = App.Path & "\graphics\globe.bmp"
    sRollOvr(5) = App.Path & "\graphics\pic20.bmp"
    sRollOvr(6) = App.Path & "\graphics\bye.bmp"
    
    frmMenu.Show
End Sub
Public Function CursorOnButton(ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim Counter As Integer
    
    'This function determines whether the cursor is over a "Menu Button"
    CursorOnButton = -1
    
    If (X >= rectangle(0).Left And X <= rectangle(0).Left + rectangle(0).Width) And ButtonsOn Then
        For Counter = 0 To 6
            If Y >= rectangle(Counter).Top And Y <= rectangle(Counter).Top _
            + rectangle(Counter).Height Then
                CursorOnButton = Counter
            Exit Function
            End If
        Next Counter
    End If
        
End Function


