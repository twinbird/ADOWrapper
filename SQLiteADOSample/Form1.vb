Option Infer On

Imports System
Imports System.Data
Imports System.Windows.Forms

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        create()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        sample()
    End Sub

    Private Sub create()
        Using accessor As New ADOWrapper.DBAccessor
            Dim q = accessor.CreateQuery
            With q.Query
                .AppendLine("CREATE TABLE IF NOT EXISTS test (")
                .AppendLine("    id integer primary key AUTOINCREMENT")
                .AppendLine("   ,text varchar(100)")
                .AppendLine(")")
            End With
            q.ExecNonQuery()
        End Using

        Using accessor As New ADOWrapper.DBAccessor

            Try
                accessor.BeginTransaction()

                Dim q1 = accessor.CreateQuery
                With q1.Query
                    .AppendLine("INSERT INTO test(text) VALUES('test1')")
                End With
                q1.ExecNonQuery()

                Dim q2 = accessor.CreateQuery
                With q2.Query
                    .AppendLine("INSERT INTO test(text) VALUES('test2')")
                End With
                q2.ExecNonQuery()

                Dim q3 = accessor.CreateQuery
                With q3.Query
                    .AppendLine("INSERT INTO test(text) VALUES('test3')")
                End With
                q3.ExecNonQuery()
                q3.ToString()

                accessor.Commit()
            Catch ex As Exception
                accessor.RollBack()
            Finally
            End Try
        End Using

    End Sub

    Private Sub sample()
        Dim acc = New ADOWrapper.DBAccessor

        Dim query1 = acc.CreateQuery
        With query1.Query
            .AppendLine("SELECT")
            .AppendLine("   *")
            .AppendLine("FROM")
            .AppendLine("   test")
            .AppendLine("WHERE")
            .AppendLine("   text = @text")
        End With
        With query1.Parameters
            .Add("@text", "test1")
        End With
        Dim dt = query1.ExecQuery()
        MessageBox.Show(dt.Rows.Count.ToString)
    End Sub

End Class
