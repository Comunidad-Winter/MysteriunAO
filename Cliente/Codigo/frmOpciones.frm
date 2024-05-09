VERSION 5.00
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   5970
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdKeys 
      Caption         =   "Controles"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox PictureSanado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   4200
      MouseIcon       =   "frmOpciones.frx":15245
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   3960
      Width           =   335
   End
   Begin VB.PictureBox PictureRecuMana 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   4440
      MouseIcon       =   "frmOpciones.frx":1554F
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   1920
      Width           =   335
   End
   Begin VB.PictureBox PictureVestirse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   4560
      MouseIcon       =   "frmOpciones.frx":15859
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   2640
      Width           =   335
   End
   Begin VB.PictureBox PictureMenosCansado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   3480
      MouseIcon       =   "frmOpciones.frx":15B63
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   3360
      Width           =   335
   End
   Begin VB.PictureBox PictureNoHayNada 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2760
      MouseIcon       =   "frmOpciones.frx":15E6D
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   4560
      Width           =   335
   End
   Begin VB.PictureBox PictureOcultarse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   4440
      MouseIcon       =   "frmOpciones.frx":16177
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   1080
      Width           =   335
   End
   Begin VB.PictureBox PictureFxs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   600
      MouseIcon       =   "frmOpciones.frx":16481
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   2040
      Width           =   335
   End
   Begin VB.PictureBox PictureMusica 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   600
      MouseIcon       =   "frmOpciones.frx":1678B
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1200
      Width           =   335
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   120
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   360
      MouseIcon       =   "frmOpciones.frx":16A95
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()


Me.Picture = LoadPicture(DirGraficos & "OpcionesDelJuego.gif")

If Musica = 0 Then
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If FX = 0 Then
    PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelOcultarse = 1 Then
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelMenosCansado = 1 Then
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelVestirse = 1 Then
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelNoHayNada = 1 Then
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelRecuMana = 1 Then
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelSanado = 1 Then
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

End Sub
Private Sub Image1_Click()

Me.Visible = False

End Sub
Private Sub Picture1_Click()

If NoRes = 0 Then
    NoRes = 1
    Picture1.Picture = LoadPicture(DirGraficos & "tick1.gif")
    Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 1)
Else
    NoRes = 0
    Picture1.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 0)
End If

MsgBox "Este cambio hará efecto recién la próxima vez que ejecutes el juego."

End Sub

Private Sub Image2_Click()
ShellExecute Me.hWnd, "open", "http://dragoon-ao.site50.net/reglamento.html", "", "", 1
End Sub

Private Sub PictureFxs_Click()

Select Case FX
    Case 0
        FX = 1
        PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Case 1
        FX = 0
        PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
End Select

End Sub
Private Sub PictureMenosCansado_Click()

If CartelMenosCansado = 0 Then
    CartelMenosCansado = 1
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelMenosCansado = 0
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "MenosCansado", Str(CartelMenosCansado))

End Sub

Private Sub PictureMusica_Click()

If Not IsPlayingCheck Then
    Musica = 0
    Play_Midi
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    Musica = 1
    Stop_Midi
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

End Sub

Private Sub PictureNoHayNada_Click()
If CartelNoHayNada = 0 Then
    CartelNoHayNada = 1
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelNoHayNada = 0
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "NoHayNada", Str(CartelNoHayNada))

End Sub

Private Sub PictureOcultarse_Click()

If CartelOcultarse = 0 Then
    CartelOcultarse = 1
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelOcultarse = 0
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Ocultarse", Str(CartelOcultarse))
End Sub

Private Sub PictureRecuMana_Click()
If CartelRecuMana = 0 Then
    CartelRecuMana = 1
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelRecuMana = 0
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "RecuMana", Str(CartelRecuMana))

End Sub

Private Sub PictureSanado_Click()
If CartelSanado = 0 Then
    CartelSanado = 1
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelSanado = 0
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Sanado", Str(CartelSanado))

End Sub

Private Sub PictureVestirse_Click()
If CartelVestirse = 0 Then
    CartelVestirse = 1
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelVestirse = 0
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Vestirse", Str(CartelVestirse))

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

Private Sub cmdKeys_Click()
Unload Me
    Call frmCustomKeys.Show(vbModeless, frmMain)
End Sub
