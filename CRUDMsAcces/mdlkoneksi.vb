Imports System.Data.OleDb
Imports System.Data

Module mdlkoneksi
    Public conn As OleDbConnection
    Public CMD As OleDbCommand
    Public DS As New DataSet
    Public DA As OleDbDataAdapter
    Public RD As OleDbDataReader
    Public lokasidata As String

    Public Sub konek()
        lokasidata = "provider=microsoft.jet.oledb.4.0;data source=biodata.mdb"
        conn = New OleDbConnection(lokasidata)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
    End Sub
End Module
