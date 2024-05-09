VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   Caption         =   "Sistema de Canjeo"
   ClientHeight    =   4485
   ClientLeft      =   540
   ClientTop       =   765
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   7245
   Begin VB.CommandButton Command1 
      Caption         =   "Canjear"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   3960
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   480
      Width           =   540
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lblPermisos 
      Height          =   975
      Left            =   3600
      TabIndex        =   10
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label lblStat 
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lblPrecio 
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblNombre 
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clases Permitidas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   2400
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stats:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio:"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   360
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista Completa de Items de Canjeo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   -15
      TabIndex        =   1
      Top             =   120
      Width           =   3645
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
If List1.Text = "Tunica Faccionaria Gris (Altos)" Then Call SendData("/CANJEO T1")
If List1.Text = "Tunica Faccionaria Roja (Bajos)" Then Call SendData("/CANJEO T2")
If List1.Text = "Tunica Faccionaria Roja (Altos)" Then Call SendData("/CANJEO T3")
If List1.Text = "Tunica Faccionaria Azul (Bajos)" Then Call SendData("/CANJEO T4")
If List1.Text = "Tunica Faccionaria Azul (Altos)" Then Call SendData("/CANJEO T5")
If List1.Text = "Ropa del Mal" Then Call SendData("/CANJEO T6")
If List1.Text = "Ropa de la Alianza" Then Call SendData("/CANJEO T7")
If List1.Text = "Tunica Transparencias" Then Call SendData("/CANJEO T8")
If List1.Text = "Chupines Grises" Then Call SendData("/CANJEO T9")
If List1.Text = "Chupines Azules" Then Call SendData("/CANJEO T10")
If List1.Text = "Chupines Rojos" Then Call SendData("/CANJEO T11")
If List1.Text = "Tunica Amarrilla" Then Call SendData("/CANJEO T12")
If List1.Text = "Tunica Celeste" Then Call SendData("/CANJEO T13")
If List1.Text = "Tunica Blanca" Then Call SendData("/CANJEO T14")
If List1.Text = "Tunica Viva" Then Call SendData("/CANJEO T15")
If List1.Text = "Tunica Verde" Then Call SendData("/CANJEO T16")
If List1.Text = "Armadura de Paloma (Bajos)" Then Call SendData("/CANJEO T17")
If List1.Text = "Armadura de Paloma (Altos)" Then Call SendData("/CANJEO T18")
If List1.Text = "Tunica de Rey (Bajos)" Then Call SendData("/CANJEO T19")
If List1.Text = "Tunica de Rey (Altos)" Then Call SendData("/CANJEO T20")
If List1.Text = "Daga +5" Then Call SendData("/CANJEO T21")
If List1.Text = "Espada Clerica" Then Call SendData("/CANJEO T22")
If List1.Text = "Espada Imperial" Then Call SendData("/CANJEO T23")
If List1.Text = "Escudo Estrella" Then Call SendData("/CANJEO T24")
If List1.Text = "Escudo Plaxus" Then Call SendData("/CANJEO T25")
If List1.Text = "Sombrero de Mago Verde" Then Call SendData("/CANJEO T26")
If List1.Text = "Sombrero de Mago Rojo" Then Call SendData("/CANJEO T27")
If List1.Text = "Tiara de la Vida" Then Call SendData("/CANJEO T28")
If List1.Text = "Coronita" Then Call SendData("/CANJEO T29")
If List1.Text = "Corona Verde" Then Call SendData("/CANJEO T30")
If List1.Text = "Coronita Dorada" Then Call SendData("/CANJEO T31")
If List1.Text = "Corona de Rey" Then Call SendData("/CANJEO T32")
If List1.Text = "Espada Fantasmal" Then Call SendData("/CANJEO T33")
If List1.Text = "Gorro de Navidad" Then Call SendData("/CANJEO T34")
If List1.Text = "Chupines Dorados" Then Call SendData("/CANJEO T35")
If List1.Text = "Chupines Negros" Then Call SendData("/CANJEO T36")
If List1.Text = "Chupines Transparentes" Then Call SendData("/CANJEO T37")
End Sub

Private Sub Form_Load()
List1.AddItem "Tunica Faccionaria Gris (Altos)"
List1.AddItem "Tunica Faccionaria Roja (Bajos)"
List1.AddItem "Tunica Faccionaria Roja (Altos)"
List1.AddItem "Tunica Faccionaria Azul (Bajos)"
List1.AddItem "Tunica Faccionaria Azul (Altos)"
List1.AddItem "Ropa del Mal"
List1.AddItem "Ropa de la Alianza"
List1.AddItem "Tunica Transparencias"
List1.AddItem "Chupines Grises"
List1.AddItem "Chupines Azules"
List1.AddItem "Chupines Rojos"
List1.AddItem "Tunica Amarrilla"
List1.AddItem "Tunica Celeste"
List1.AddItem "Tunica Blanca"
List1.AddItem "Tunica Viva"
List1.AddItem "Tunica Verde"
List1.AddItem "Armadura de Paloma (Bajos)"
List1.AddItem "Armadura de Paloma (Altos)"
List1.AddItem "Tunica de Rey (Bajos)"
List1.AddItem "Tunica de Rey (Altos)"
List1.AddItem "Daga +5"
List1.AddItem "Espada Clerica"
List1.AddItem "Espada Fantasmal"
List1.AddItem "Espada Imperial"
List1.AddItem "Escudo Estrella"
List1.AddItem "Escudo Plaxus"
List1.AddItem "Sombrero de Mago Verde"
List1.AddItem "Sombrero de Mago Rojo"
List1.AddItem "Tiara de la Vida"
List1.AddItem "Coronita"
List1.AddItem "Corona Verde"
List1.AddItem "Coronita Dorada"
List1.AddItem "Corona de Rey"
List1.AddItem "Gorro de Navidad"
List1.AddItem "Chupines Dorados"
List1.AddItem "Chupines Negros"
List1.AddItem "Chupines Transparentes"
End Sub


Private Sub list1_Click()
If List1.Text = "Tunica Faccionaria Gris (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16119.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "15 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Faccionaria Roja (Bajos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16126.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "15 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Faccionaria Roja (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16126.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "15 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Faccionaria Azul (Bajos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16129.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "15 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Faccionaria Azul (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16129.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "15 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Ropa del Mal" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16132.bmp")
    lblPrecio.Caption = "20 Puntos de Canje"
    lblPrecio.Caption = ""
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Ropa de la Alianza" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16134.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Transparencias" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16136.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Chupines Grises" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16184.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Chupines Azules" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16182.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Chupines Rojos" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16180.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Amarrilla" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16214.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Celeste" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16216.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Blanca" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16231.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Viva" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16233.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica Verde" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16235.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Armadura de Paloma (Bajos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16123.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "40 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Paladin, Guerrero, Cazador, Arkero"
End If

If List1.Text = "Armadura de Paloma (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16123.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "40 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Paladin, Guerrero, Cazador, Arkero"
End If

If List1.Text = "Tunica de Rey (Bajos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16089.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "4 Puntos de Canje"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Tunica de Rey (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16089.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "4 Puntos de Canje"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Daga +5" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16150.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 8 / Max: 12"
    lblPermisos.Caption = "Solo Asesino (Cualquier Raza)"
End If

If List1.Text = "Espada Clerica" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16200.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 18 / Max: 21"
    lblPermisos.Caption = "Solo Clerigos"
End If

If List1.Text = "Espada Fantasmal" Then
    Picture1.Picture = LoadPicture(DirGraficos & "9630.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 21 / Max: 24"
    lblPermisos.Caption = "Solo Guerreros"
End If

If List1.Text = "Espada Imperial" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16083.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 20 / Max: 22"
    lblPermisos.Caption = "Solo Paladines"
End If

If List1.Text = "Escudo Estrella" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16152.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 30 / Max: 40"
    lblPermisos.Caption = "Guerrero, Paladin, Clerigo"
End If

If List1.Text = "Escudo Plaxus" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16154.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 15 / Max: 20"
    lblPermisos.Caption = "Asesino, Bardo y Druida"
End If

If List1.Text = "Sombrero de Mago Verde" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16253.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 10 / Max: 13"
    lblPermisos.Caption = "Mago y Nigromante"
End If

If List1.Text = "Sombrero de Mago Rojo" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16237.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 10 / Max: 13"
    lblPermisos.Caption = "Mago y Nigromante"
End If

If List1.Text = "Tiara de la Vida" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16226.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "5 Puntos de Canje"
    lblStat.Caption = "Min: 30 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Coronita" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16190.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "10 Puntos de Canje"
    lblStat.Caption = "Min: 35 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Corona Verde" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16245.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 60"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Coronita Dorada" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16255.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 60"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Corona de Rey" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16192.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 60"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Gorro de Navidad" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16023.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "2 Puntos de Canje"
    lblStat.Caption = "Min: 30 / Max: 40"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Chupines Dorados" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16264.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Chupines Negros" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16266.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

If List1.Text = "Chupines Transparentes" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16268.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas las Clases"
End If

End Sub
