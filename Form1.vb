Imports System.Data.SqlClient
Imports System.Text
Public Class Form1
    Dim df1 As New DataFunctions.DataFunctions
    Dim gf1 As New GlobalFunction1.GlobalFunction1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim TxtControlFile As String = "D:\saralwork\configfile\configfile.txt"
        gf1.SetGlobalVariables(TxtControlFile)
        Dim arr1 As List(Of String) = Index_script("1_srv_1.3_mdf_3", "UserDetails1234")
        Dim txt As Encoding = Encoding.UTF8
        Dim str As String() = arr1.ToArray
        gf1.StringWriteAllLines("C:\Users\USER\Desktop\Index_Create_Script.txt", str, txt)
    End Sub
    Public Function Index_Script(ByVal server_database_name As String, ByVal table As String) As List(Of String)
        Dim arr As New List(Of String)
        Dim primary_key_element As String = df1.GetPrimaryKey(server_database_name, table)
        Dim index_element As String = ""
        Dim index_str As String = ""
        Dim index_str1 As String = ""
        Dim dt As DataTable
        Dim dt_str As String = "EXEC sp_helpindex" & " " & table
        dt = df1.SqlExecuteDataTable(server_database_name, dt_str)
        If dt.Rows.Count <> 0 Then
            index_element = Convert.ToString(dt.Rows(0).Item("index_keys"))
        End If
        For i = 0 To dt.Rows.Count - 1
            index_element = Convert.ToString(dt.Rows(i).Item("index_keys"))
            If index_element = primary_key_element Then
                Continue For
            Else
                If Convert.ToString(dt.Rows(i).Item("index_description")).Contains("unique") Then
                    index_str = "UNIQUE"
                    index_str1 = "IGNORE_DUP_KEY = OFF,"
                Else
                    index_str = ""
                    index_str1 = ""
                End If
                arr.Add("CREATE " & index_str & " NONCLUSTERED INDEX" & "[" & Convert.ToString(dt.Rows(i).Item("index_name")) & "] " & "ON" & "[" & "dbo" & "]." & "[" & table & "]")
                arr.Add("(")
                arr.Add("[" & index_element & "]" & " ASC")
                arr.Add(")" & "WITH" & "(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, " & index_str1 & " DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]")
            End If
        Next
        Return arr
    End Function
End Class