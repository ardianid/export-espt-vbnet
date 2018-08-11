Imports System.Text
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Data
Imports System.Data.OleDb
Imports System.Security.Cryptography

Public Class ClassMy

    Public Shared Function open_conn() As OleDbConnection

        Dim mloc As String = ""
        Dim mdbase As String = ""
        Dim muser As String = ""
        Dim mpwd As String = ""

        Dim cn = New OleDbConnection

        Try

            Dim myconnectionstring As String = String.Format("Provider=SQLOLEDB;Server={0};Database={1};Uid={2};Pwd={3};", "swtl\SQL2008R2", "db_penghubung", "sa", "")

            cn.ConnectionString = myconnectionstring
            cn.Open()

        Catch ex As OleDb.OleDbException
            Throw New Exception(ex.ToString)
        End Try

        Return cn

    End Function

    Public Shared Function open_conn_mobiz() As OleDbConnection

        Dim mloc As String = ""
        Dim mdbase As String = ""
        Dim muser As String = ""
        Dim mpwd As String = ""

        Dim cn = New OleDbConnection

        Try

            Dim myconnectionstring As String = String.Format("Provider=SQLOLEDB;Server={0};Database={1};Uid={2};Pwd={3};", "swtl\SQL2008R2", "", "sa", "")

            cn.ConnectionString = myconnectionstring
            cn.Open()

        Catch ex As OleDb.OleDbException
            Throw New Exception(ex.ToString)
        End Try

        Return cn

    End Function

    Public Shared Function GetDataSet(ByVal SQL As String, ByVal cn As OleDbConnection) As DataSet

        Dim adapter As New OleDbDataAdapter(SQL, cn)
        Dim myData As New DataSet
        adapter.Fill(myData)

        adapter.Dispose()

        Return myData
    End Function


End Class
