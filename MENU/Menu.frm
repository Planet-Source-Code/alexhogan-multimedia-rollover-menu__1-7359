VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   7110
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgRollOvr 
      Height          =   1455
      Left            =   5640
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgButton 
      Height          =   300
      Left            =   120
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4245
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3690
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LastButton As Integer

Private Sub Form_Load()
    LastButton = -1
    ButtonsOn = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intButton As Integer
    
    'This is defining the value of intButton for menu button positions
    intButton = CursorOnButton(X, Y)
    
        'This controls the rollover for the buttons
        If intButton > -1 Then
            If LastButton <> intButton Then
                lblAbout = sMenuButtonText(intButton)
                lblAbout.Visible = True
                imgButton.Move rectangle(intButton).Left, rectangle(intButton).Top
                imgButton.Picture = LoadPicture(sGraphics(intButton))
                imgButton.Visible = True
                imgButton.Enabled = True
                imgRollOvr.Picture = LoadPicture(sRollOvr(intButton))
                imgRollOvr.Visible = True
                LastButton = intButton
            End If
        Else
            lblAbout.Visible = False
            imgButton.Visible = False
            imgButton.Enabled = False
            imgRollOvr.Visible = False
            LastButton = -1
        End If
    
End Sub


Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
        
        'This controls what happens when a menu button is clisked
        If PtInRegion(hRectRgn(0), X + imgButton.Left, Y + imgButton.Top) Then
            MsgBox "This is the About"
            
        ElseIf PtInRegion(hRectRgn(1), X + imgButton.Left, Y + imgButton.Top) Then
            MsgBox "This is Form 1"
            
        ElseIf PtInRegion(hRectRgn(2), X + imgButton.Left, Y + imgButton.Top) Then
            MsgBox "This is Form 2"
            
        ElseIf PtInRegion(hRectRgn(3), X + imgButton.Left, Y + imgButton.Top) Then
            MsgBox "This is Form 3"
            
        ElseIf PtInRegion(hRectRgn(4), X + imgButton.Left, Y + imgButton.Top) Then
            MsgBox "This is Form 4"
            
        ElseIf PtInRegion(hRectRgn(5), X + imgButton.Left, Y + imgButton.Top) Then
            MsgBox "This is Form 5"
            
        ElseIf PtInRegion(hRectRgn(6), X + imgButton.Left, Y + imgButton.Top) Then
            MsgBox "Thank you for viewing the ""Rollover Menu"" program." & vbCrLf & vbCrLf _
            & "         For questions and comments send e-mail to;" & vbCrLf _
            & "                     hogana@usapathway.com"
            Unload Me
        End If
       
End Sub
