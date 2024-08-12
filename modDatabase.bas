Attribute VB_Name = "modDatabase"
Public conn As New ADODB.Connection

Public Sub OpenConnection()
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=HubLibros;User ID=User;Password=Password;"
    conn.Open
End Sub
