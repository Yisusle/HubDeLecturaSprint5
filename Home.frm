VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Home 
   Caption         =   "Home"
   ClientHeight    =   6240
   ClientLeft      =   3360
   ClientTop       =   2850
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnLibroNomegustaVista 
      Caption         =   "Ver no me gusta"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton btnLibrosParaLeer 
      Caption         =   "Libros Para Leer"
      Height          =   375
      Index           =   4
      Left            =   7680
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnGeneroFavorito 
      Caption         =   "Generos Favoritos"
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnLibrosRecomendados 
      Caption         =   "Libros Recomendados"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton btnLibrosLeidos 
      Caption         =   "Libros Leidos"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvBooks 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8493
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10200
      TabIndex        =   0
      Top             =   5760
      Width           =   735
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExit_Click()
    If MsgBox("¿Desea Salir?", vbQuestion + vbYesNo, "¿Desea Salir de la Aplicación?") = vbYes Then
        End
    End If
End Sub

Private Sub btnGeneroFavorito_Click(Index As Integer)
    FavoriteGenres.Show
End Sub


Private Sub btnLibroNomegustaVista_Click(Index As Integer)
    DislikedBooks.Show
End Sub

Private Sub btnLibrosLeidos_Click(Index As Integer)
    ReadBooks.Show
End Sub

Private Sub btnLibrosParaLeer_Click(Index As Integer)
    ToReadBooks.Show
End Sub

Private Sub btnLibrosRecomendados_Click(Index As Integer)
    RecommendedBooks.Show
End Sub

Private Sub Form_Load()
    ' Abrir la conexión a la base de datos
    OpenConnection
    
    ' Configurar las columnas del ListView
    With lvBooks
        .View = lvwReport
        .ColumnHeaders.Add , , "Título", 2200
    End With
    
    ' Cargar los datos de los libros
    LoadBooks
End Sub

Private Sub LoadBooks()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT Book.BookID,Genre.GenreID, Genre.Name AS Genre, Book.Title, Book.Author, Book.Year, Book.Description, Book.CoverImage FROM Book INNER JOIN Genre ON Book.GenreID = Genre.GenreID"
    
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Limpiar el ListView antes de cargar nuevos datos
    lvBooks.ListItems.Clear
    
    ' Cargar los datos en el ListView
    Do While Not rs.EOF
        Dim item As ListItem
        Set item = lvBooks.ListItems.Add(, , rs("Title"))
        ' Almacenar la información en la propiedad Tag
        item.Tag = rs("BookID") & "|" & rs("GenreID") & "|" & rs("Title") & "|" & rs("Author") & "|" & rs("Year") & "|" & rs("Genre") & "|" & rs("Description") & "|" & rs("CoverImage")
        rs.MoveNext
    Loop
    
    rs.Close
End Sub

Private Sub lvBooks_Click()
    On Error GoTo ErrorHandler

    If Not lvBooks.SelectedItem Is Nothing Then
        Dim selectedBook As ListItem
        Set selectedBook = lvBooks.SelectedItem

        Dim bookInfo() As String
        bookInfo = Split(selectedBook.Tag, "|")

        ' Instanciar
        Dim BookDetails As New BookDetails

        ' Verifica la cantidad de datos antes de asignar
        If UBound(bookInfo) >= 6 Then
            ' Abrir el formulario de detalles del libro y pasarle la información
            With BookDetails
                 .BookID = CInt(bookInfo(0)) ' BookID
                 .GenreID = CInt(bookInfo(1)) ' GenreID
                .txtTitle.Text = bookInfo(2)
                .txtAuthor.Text = bookInfo(3)
                .txtYear.Text = bookInfo(4)
                .txtGenre.Text = bookInfo(5)
                .txtDescription.Text = bookInfo(6)
                
                ' Descargar y cargar la imagen de portada
                On Error Resume Next
                Dim imagePath As String
                imagePath = DownloadImage(bookInfo(7))
                .imgCover.Picture = LoadPicture(imagePath)
                If Err.Number <> 0 Then
                    MsgBox "Error al cargar la imagen: " & Err.Description, vbExclamation
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
                ' Mostrar el formulario de detalles
                .Show

                ' Eliminar Imagen temp
                On Error Resume Next
                Kill imagePath
                If Err.Number <> 0 Then
                    MsgBox "Error al eliminar el archivo temporal: " & Err.Description, vbExclamation
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End With
        Else
            MsgBox "Información del libro incompleta.", vbExclamation
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error en lvBooks_Click: " & Err.Description, vbCritical
    Err.Clear
End Sub

Function DownloadImage(imageUrl As String) As String
    Dim http As Object
    Dim stream As Object
    Dim tempFile As String

    ' objeto
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", imageUrl, False
    http.Send

    ' guardar imagen
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write http.responseBody
    tempFile = "C:\Temp\tempImage.jpg" ' Ruta del archivo temporal
    stream.SaveToFile tempFile, 2 '
    stream.Close

    ' Devolver la ruta del archivo temporal
    DownloadImage = tempFile
End Function
