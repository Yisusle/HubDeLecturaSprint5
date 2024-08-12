VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FavoriteGenres 
   Caption         =   "FavoriteGenres"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnGoBack 
      Caption         =   "Volver"
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvFavoriteGenres 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9128
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
Attribute VB_Name = "FavoriteGenres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGoBack_Click()
    Unload Me
End Sub

Private Sub Form_Load()
' Configurar las columnas del ListView
    With lvFavoriteGenres
        .View = lvwReport
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Género", 3000
    End With
    
    ' Cargar los datos de los géneros favoritos
    LoadFavoriteGenres
End Sub

Private Sub LoadFavoriteGenres()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT Genre.Name AS Genre FROM FavoriteGenre INNER JOIN Genre ON FavoriteGenre.GenreID = Genre.GenreID"
    
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Limpiar el ListView antes de cargar nuevos datos
    lvFavoriteGenres.ListItems.Clear
    
    ' Cargar los datos en el ListView
    Do While Not rs.EOF
        Dim item As ListItem
        Set item = lvFavoriteGenres.ListItems.Add(, , rs("Genre"))
        rs.MoveNext
    Loop
    
    rs.Close
End Sub

Private Sub lvFavoriteGenres_BeforeLabelEdit(Cancel As Integer)
    ' Cancelar la edición de las etiquetas del ListView
    Cancel = True
End Sub
