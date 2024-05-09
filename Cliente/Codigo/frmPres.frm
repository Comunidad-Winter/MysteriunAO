VERSION 5.00
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPres.frx":0000
   ScaleHeight     =   8640
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   3400
      Left            =   1125
      Top             =   1230
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then finpres = True
End Sub

Private Sub Timer1_Timer()
Static ticks As Long

ticks = ticks + 1

If ticks = 1 Then
    Me.Picture = LoadPicture(App.Path & "\Graficos\Dragoon.jpg")
ElseIf ticks < 13 Then
    Me.Picture = LoadPicture(App.Path & "\Graficos\intro.jpg")
Else
 finpres = True
End If

End Sub
