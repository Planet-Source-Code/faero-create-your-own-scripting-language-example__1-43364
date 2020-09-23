VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4875
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   1995
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmScript.frx":0442
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_Resize
    rtfText.SelIndent = 16
    rtfText.text = "' Simple Hello World! program:" & vbCrLf & _
    "MessageBox(" & Chr(34) & "Hello World" & Chr(34) & ", " & Chr(34) & "Hi there" & Chr(34) & ", 0)" & _
    vbCrLf & "' Puts Hello World in a message box!" & vbCrLf & _
    "Faero()" & vbCrLf & "'Uses the all special Faero command!"
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 16, 0, Me.ScaleWidth - 16, Me.ScaleHeight
End Sub

