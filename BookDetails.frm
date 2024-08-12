VERSION 5.00
Begin VB.Form BookDetails 
   Caption         =   "BookDetails"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRecommendBook 
      Caption         =   "Recomendar Libro"
      Height          =   375
      Left            =   9600
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton btnGoBack 
      Caption         =   "Volver"
      Height          =   375
      Left            =   10200
      TabIndex        =   10
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton btnAddReadBooks 
      Caption         =   "Añadir a Libros Leidos"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton btnAddToReadLater 
      Caption         =   "Leer mas Tarde"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton btnAddToFavoriteGenre 
      Caption         =   "Añadir Genero a Favorito"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton btnLibroNomegusta 
      Caption         =   "No me gusta"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtDescription 
      Height          =   1575
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "BookDetails.frx":0000
      Top             =   2400
      Width           =   6975
   End
   Begin VB.TextBox txtGenre 
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Genre"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtYear 
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Year"
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtAuthor 
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Author"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Title"
      Top             =   600
      Width           =   2895
   End
   Begin VB.PictureBox imgCover 
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3195
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Descripción:"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Genero:"
      Height          =   255
      Left            =   7680
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Autor:"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Año:"
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo:"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "BookDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BookID As Integer
Public GenreID As Integer

Private Sub btnAddReadBooks_Click()
    Dim sql As String
    sql = "INSERT INTO ReadBook (BookID) VALUES (" & BookID & ")"
    conn.Execute sql
    MsgBox "El libro '" & txtTitle.Text & "' se ha añadido a la lista de libros leídos.", vbInformation
End Sub

Private Sub btnAddToFavoriteGenre_Click()
    Dim sql As String
    sql = "INSERT INTO FavoriteGenre (GenreID) VALUES (" & GenreID & ")"
    conn.Execute sql
    MsgBox "El género '" & txtGenre.Text & "' se ha añadido a la lista de géneros favoritos.", vbInformation
End Sub

Private Sub btnAddToReadLater_Click()
    Dim sql As String
    sql = "INSERT INTO ToReadBook (BookID) VALUES (" & BookID & ")"
    conn.Execute sql
    MsgBox "El libro '" & txtTitle.Text & "' se ha añadido a la lista de libros para leer más tarde.", vbInformation
End Sub

Private Sub btnLibroNomegusta_Click()
    Dim sql As String
    sql = "INSERT INTO DislikedBook (BookID) VALUES (" & BookID & ")"
    conn.Execute sql
    MsgBox "El libro '" & txtTitle.Text & "' se ha añadido a la lista de libros que no te gustan.", vbInformation
End Sub

Private Sub btnGoBack_Click()
    Unload Me
End Sub

Private Sub btnRecommendBook_Click()
    Dim sql As String
    sql = "INSERT INTO RecommendedBook (BookID) VALUES (" & BookID & ")"
    conn.Execute sql
    MsgBox "El libro '" & txtTitle.Text & "' se ha añadido a la lista de libros recomendados.", vbInformation
End Sub

