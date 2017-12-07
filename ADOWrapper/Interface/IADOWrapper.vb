''' <summary>
''' ADO.NETのうすーいラッパーインターフェース
''' </summary>
Public Interface IADOWrapper
    Inherits IDisposable

    ''' <summary>
    ''' 実行するSQL
    ''' </summary>
    ''' <returns></returns>
    Property Query As QueryBuilder

    ''' <summary>
    ''' SQL実行時に利用するパラメータ
    ''' </summary>
    ''' <returns></returns>
    Property Parameters As IADOWrapperParameters

    ''' <summary>
    ''' プロパティに設定された情報でクエリを発行します
    ''' </summary>
    ''' <returns>クエリ結果のデータテーブル</returns>
    Function ExecQuery() As DataTable

    ''' <summary>
    ''' プロパティに設定された情報でクエリを発行します
    ''' </summary>
    ''' <returns>影響を与えた行数</returns>
    Function ExecNonQuery() As Integer

    ''' <summary>
    ''' 内部保持のSQLを返します
    ''' </summary>
    ''' <returns></returns>
    Function ToString() As String

End Interface
