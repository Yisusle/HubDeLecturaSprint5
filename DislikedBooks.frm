VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DislikedBooks 
   Caption         =   "DislikedBooks"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnGoBack 
      Caption         =   "Volver"
      Height          =   375
      Left            =   9840
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvDislikedBooks 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "DislikedBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGoBack_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 ' Configurar las columnas del ListView
    With lvDislikedBooks
        .View = lvwReport
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Título", 2200
        .ColumnHeaders.Add , , "Autor", 2000
        .ColumnHeaders.Add , , "Año", 800
        .ColumnHeaders.Add , , "Género", 1500
        .ColumnHeaders.Add , , "Descripción", 6000
    End With
    
    ' Cargar los datos de los libros no gustados
    LoadDislikedBooks
End Sub

Private Sub LoadDislikedBooks()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT Book.Title, Book.Author, Book.Year, Genre.Name AS Genre, Book.Description FROM DislikedBook INNER JOIN Book ON DislikedBook.BookID = Book.BookID INNER JOIN Genre ON Book.GenreID = Genre.GenreID"
    
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Limpiar el ListView antes de cargar nuevos datos
    lvDislikedBooks.ListItems.Clear
    
    ' Cargar los datos en el ListView
    Do While Not rs.EOF
        Dim item As ListItem
        Set item = lvDislikedBooks.ListItems.Add(, , rs("Title"))
        item.SubItems(1) = rs("Author")
        item.SubItems(2) = rs("Year")
        item.SubItems(3) = rs("Genre")
        item.SubItems(4) = rs("Description")
        rs.MoveNext
    Loop
    
    rs.Close
End Sub

Private Sub lvDislikedBooks_BeforeLabelEdit(Cancel As Integer)
    ' Cancelar la edición de las etiquetas del ListView
    Cancel = True
End Sub
