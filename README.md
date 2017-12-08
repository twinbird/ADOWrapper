# ADOWrapper

ADO.NETをVB.NETで使う際のうすーいラッパー.

### 単純なSelectクエリ

```vb
Using accessor As New ADOWrapper.DBAccessor(ConnectionString)
	Dim q = accessor.CreateQuery
		With q.Query
			.AppendLine("SELECT * FROM test")
		End With
	Dim dt = q.ExecQuery '=> dtは結果セットのDataTableです
End Using
```

### 単一値を取得するSelectクエリ

```vb
Using accessor As New ADOWrapper.DBAccessor(ConnectionString)
    Dim q = accessor.CreateQuery
    With q.Query
        .AppendLine("SELECT")
        .AppendLine("   COUNT(*)")
        .AppendLine("FROM")
        .AppendLine("   test")
    End With

    Dim obj = q.ExecScalar '=> obj = COUNT(*)
End Using
```

### 単純なInsert/Update/Deleteクエリ

```vb
Using accessor As New ADOWrapper.DBAccessor(ConnectionString)
    Dim q = accessor.CreateQuery
    With q.Query
        .AppendLine("INSERT INTO test(text) VALUES('text1');")
        .AppendLine("INSERT INTO test(text) VALUES('text1');")
        .AppendLine("INSERT INTO test(text) VALUES('text2');")
        .AppendLine("INSERT INTO test(text) VALUES('text2');")
    End With
    Dim ret = q.ExecNonQuery()
End Using
```

### トランザクション

```vb
Using accessor As New ADOWrapper.DBAccessor(ConnectionString)
    accessor.BeginTransaction()

    Try
        Dim q = accessor.CreateQuery
        With q.Query
            .AppendLine("INSERT INTO test(text) VALUES('text1');")
        End With
        q.ExecNonQuery()

        Dim q2 = accessor.CreateQuery
        With q2.Query
            .AppendLine("INSERT INTO test(text) VALUES('text1');")
        End With
        q2.ExecNonQuery()

        Throw New Exception("もしここで例外が発生すると")

        Dim q3 = accessor.CreateQuery
        With q3.Query
            .AppendLine("INSERT INTO test(text) VALUES('text1');")
        End With
        q3.ExecNonQuery()

		'問題が無ければここでコミットします
        accessor.Commit()

    Catch ex As Exception
		'ここでロールバックします
        accessor.RollBack()
    End Try
End Using
```

### パラメータクエリ

```vb
Using accessor As New ADOWrapper.DBAccessor(ConnectionString)
    Dim q = accessor.CreateQuery
    With q.Query
        .AppendLine("SELECT")
        .AppendLine("   *")
        .AppendLine("FROM")
        .AppendLine("   test")
        .AppendLine("WHERE")
        .AppendLine("   text = @text")
    End With

    With q.Parameters
        .Add("text", "text1")
    End With

    Dim dt = q.ExecQuery 'dtは結果セットのDataTableです
End Using
```

### Connection String

接続文字列はDBAccessor生成時に指定する事ができますが, デフォルトではApp.Config内の```mainDB```を利用します.

##### Connection String指定の場合

Dim accessor As New ADOWrapper.DBAccessor(ConnectionString)

##### デフォルトのConnection Stringを利用する場合

Dim accessor As New ADOWrapper.DBAccessor

### See Also

詳細はテストプロジェクト内のコードを参照

#### LISENCE

MIT
