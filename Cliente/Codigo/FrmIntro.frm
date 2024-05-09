VERSION 5.00
Begin VB.Form FrmIntro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   4170
   ClientTop       =   2565
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   Picture         =   "FrmIntro.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ONLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   3000
      TabIndex        =   1
      Top             =   6840
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   480
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":1A7FB
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":1AB05
      MousePointer    =   99  'Custom
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":1AE0F
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":1B119
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   840
      MouseIcon       =   "FrmIntro.frx":1B423
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   3495
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@Dragoonao.com.ar
'www.Dragoonao.com.ar

Private Sub Form_Load()


Dim corriendo As Integer
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim pepe As String

Dim exeName As String
snap = CreateToolhelpSnapshot(TH32CS_SNAPALL, 0)
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)
i = 0
While theloop <> 0
    exeName = proc.szexeFile
    Text1.Text = proc.szexeFile
    If Text1.Text = "DragoonAONoDinamico.exe" Or Text1.Text = "DragoonAO.exe" Then
        corriendo = corriendo + 1
        Text1.Text = ""
    End If
    i = i + 1
    theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap

End Sub

Private Sub Image1_Click()
ShellExecute Me.hwnd, "open", App.Path & "/Autoupdate.exe", "", "", 1
Unload Me
End Sub

Private Sub Image2_Click()
If FindWindow(vbNullString, UCase$("Dragoon AO")) Then
    MsgBox "No está permitido el uso de doble cliente", vbExclamation
    End
Else
Call Main
End If
End Sub

Private Sub Image3_Click()
ShellExecute Me.hwnd, "open", App.Path & "/aosetup.exe", "", "", 1
End Sub

Private Sub Image4_Click()
ShellExecute Me.hwnd, "open", "http://mysteriun-ao.ucoz.com", "", "", 1

End Sub

Private Sub Image5_Click()
ShellExecute Me.hwnd, "open", "http://mysteriun-ao.ucoz.com", "", "", 1

End Sub

Private Sub Image6_Click()
Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton Then

      DX = X

      dy = Y

      bmoving = True

   End If

   

End Sub

 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving And ((X <> DX) Or (Y <> dy)) Then

      Move Left + (X - DX), Top + (Y - dy)

   End If

   

End Sub

 

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then

      bmoving = False

   End If

   

End Sub
