Imports System.Data.OleDb
Module Access
    Public conn As New OleDbConnection
    Public cmd As OleDbCommand
    Public da As OleDbDataAdapter
    Public ds As DataSet
    Public dr As OleDbDataReader
    Public Qry As String
    Public MobilQry As String
    Public SupirQry As String
    Sub koneksi()
        conn = New OleDbConnection("provider=microsoft.jet.oledb.4.0;Data Source=db.mdb;")
        conn.Open()
    End Sub
End Module
