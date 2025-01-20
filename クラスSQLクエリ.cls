VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "クラスSQLクエリ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************************************************
'*  機　能：クラスSQLクエリ
'*  概　要：SQLクエリの実行を管理する
'*  補　足：
'**************************************************************************************************************************
'更新履歴
'25/01/17: SQL実行関数をSQL Server用の他にAccess用を作成。
'          Excelシートの表に対して直接SQL実行出来るほか、Access固有関数(TRANSFORM)を使用出来るようになった。
'25/01/16: 関数「SQLクエリ実行結果出力」に書式設定を追加。オフセットに対応。
'24/10/08: 関数「SQLクエリ実行結果出力」を追加。

Private Const モジュール名 As String = "クラスSQLクエリ"

Public Function 実行結果_SQLServer(ByVal SQL文 As String) As Variant
    '--------------------------------------------------------------------------------------------------------------------------
    'エラー処理（前半）
    On Error GoTo エラー時
    Const プロシージャ名 As String = "実行結果_SQLServer"
    '--------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------------------------
    'Application.ScreenUpdating = False
    'Application.Cursor = xlWait
    'Application.DisplayAlerts = False
    'ThisWorkbook.Unprotect
    '--------------------------------------------------------------------------------------------------------------------------
    Dim データベース接続文 As String
 
    'OLE DB
    Const プロバイダー = "Provider=プロバイダー名;"
    Const サーバー = "Data Source=サーバー名;"                'サーバー,ポート
    Const 認証 = "Trusted_connection=yes;"                    'Windows認証の場合（「"Integrated Security=SSPI;"」でもよい）
    Const データベース = "Initial Catalog=データベース名;"                '接続するデータベース名
    データベース接続文 = プロバイダー & サーバー & 認証 & データベース       'Windows認証の場合
    
    '--------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------------------------
    'Const ユーザー = "User ID=sqluser;"                               'SQL Server認証の場合のみ指定
    'Const パスワード = "Password=psword;"                               'SQL Server認証の場合のみ指定
    'データベース接続文 = プロバイダー & サーバー & データベース & ユーザー & パスワード    'SQL Server認証の場合
    
    'ODBC
    'Const DSN = "DSN=ODBCユーザーデータソース名;"        'ODBCユーザーデータソース名
    'データベース接続文 = DSN                                 'Windows認証の場合
    'Const USER = "UID=sqluser;"                  'ユーザー名（SQL Server認証を使用する場合のみ指定）
    'Const PSWD = "PWD=psword;"                   'パスワード（SQL Server認証を使用する場合のみ指定）
    'データベース接続文 = DSN & USER & PSWD                   'SQL Server認証の場合
    '--------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------------------------
    
    Dim 接続データベース As New ADODB.Connection
    Dim 読込テーブル As New ADODB.Recordset
    
    接続データベース.ConnectionString = データベース接続文
    接続データベース.Open

    'SQL実行、配列へ出力、DB接続解除（コピペ可）
    読込テーブル.Source = SQL文
    読込テーブル.ActiveConnection = 接続データベース
    読込テーブル.Open
    
    Dim i, j As Long
    Dim SQL実行結果, 出力用配列 As Variant
    Dim 列数, レコード数 As Long
    列数 = 読込テーブル.Fields.Count
    
    ReDim SQL実行結果(1 To 列数, 1 To 2)
    
    レコード数 = 1
    Do Until 読込テーブル.EOF
        For i = 0 To 列数 - 1
            If レコード数 = 1 Then
                SQL実行結果(i + 1, レコード数) = 読込テーブル.Fields(i).Name
            End If
            SQL実行結果(i + 1, レコード数 + 1) = 読込テーブル(i).Value
        Next i
        読込テーブル.MoveNext
        ReDim Preserve SQL実行結果(1 To 列数, 1 To 2 + レコード数)
        レコード数 = レコード数 + 1
    Loop
    
    ReDim 出力用配列(1 To レコード数, 1 To 列数)
    For i = 1 To レコード数
        For j = 1 To 列数
            出力用配列(i, j) = SQL実行結果(j, i)
        Next j
    Next i

    '終了処理
    読込テーブル.Close
    Set 読込テーブル = Nothing
    接続データベース.Close
    Set 接続データベース = Nothing
    
    実行結果_SQLServer = 出力用配列
    
    '--------------------------------------------------------------------------------------------------------------------------
    'Application.ScreenUpdating = True
    'Application.Cursor = xlDefault
    'Application.DisplayAlerts = True
    'ThisWorkbook.Protect
    '--------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------------------------
    'エラー処理（後半）
    Exit Function
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
    '--------------------------------------------------------------------------------------------------------------------------
End Function

Public Function 実行結果_Access(ByVal SQL文 As String) As Variant
    '注意: ブックとシートを読み込む。シートのテーブルはA1セルから始まっている必要がある。
    '注意: 対象のシートに設定されたNumberFormatLocalが影響することに注意が必要
    '--------------------------------------------------------------------------------------------------------------------------
    'エラー処理（前半）
    On Error GoTo エラー時
    Const プロシージャ名 As String = "実行結果_Access"
    '--------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------------------------
    'Application.ScreenUpdating = False
    'Application.Cursor = xlWait
    'Application.DisplayAlerts = False
    'ThisWorkbook.Unprotect
    '--------------------------------------------------------------------------------------------------------------------------
    Dim ブック As Workbook
    Set ブック = ThisWorkbook
    Dim データベース接続文 As String
    データベース接続文 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ブック.FullName & ";Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=True';"    'IMEX=1;
    
    Dim 接続データベース As New ADODB.Connection
    Dim 読込テーブル As New ADODB.Recordset

    
    接続データベース.ConnectionString = データベース接続文
    接続データベース.Open
    
    'SQL実行、配列へ出力、DB接続解除（コピペ可）
    読込テーブル.Source = SQL文
    読込テーブル.ActiveConnection = 接続データベース
    読込テーブル.Open

    Dim i, j As Long
    Dim SQL実行結果, 出力用配列 As Variant
    Dim 列数, レコード数 As Long
    列数 = 読込テーブル.Fields.Count
    
    'ヘッダーのみ格納
    ReDim SQL実行結果(1 To 列数, 1 To 1)
    For i = 0 To 列数 - 1
        SQL実行結果(i + 1, 1) = 読込テーブル.Fields(i).Name
    Next i
    
    'レコードの格納
    レコード数 = 0
    Do Until 読込テーブル.EOF
        レコード数 = レコード数 + 1
        ReDim Preserve SQL実行結果(1 To 列数, 1 To レコード数 + 1)
        For i = 0 To 列数 - 1
            SQL実行結果(i + 1, レコード数 + 1) = 読込テーブル(i).Value
        Next i
        読込テーブル.MoveNext
    Loop
    
    '出力用配列に変換
    ReDim 出力用配列(1 To UBound(SQL実行結果, 2), 1 To UBound(SQL実行結果, 1))
    For i = 1 To UBound(SQL実行結果, 2)
        For j = 1 To UBound(SQL実行結果, 1)
            出力用配列(i, j) = SQL実行結果(j, i)
        Next j
    Next i

    '終了処理
    読込テーブル.Close
    Set 読込テーブル = Nothing
    接続データベース.Close
    Set 接続データベース = Nothing
    
    実行結果_Access = 出力用配列

    '--------------------------------------------------------------------------------------------------------------------------
    'Application.ScreenUpdating = True
    'Application.Cursor = xlDefault
    'Application.DisplayAlerts = True
    'ThisWorkbook.Protect
    '--------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------------------------
    'エラー処理（後半）
    Exit Function
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
    '--------------------------------------------------------------------------------------------------------------------------
End Function

Public Function SQLクエリ実行結果出力(ByVal 出力配列 As Variant, ByVal シート As Object, Optional ByVal タイトル As String, Optional ByRef Offset行 As Long, Optional ByRef Offset列 As Long, Optional ByVal 書式設定する As Boolean = True)
    '--------------------------------------------------------------------------------------------------------------------------
    'エラー処理（前半）
    On Error GoTo エラー時
    Const プロシージャ名 As String = "SQLクエリ実行結果出力"
    '--------------------------------------------------------------------------------------------------------------------------
    
    If Offset行 < -1 Then Offset行 = -1
    If Offset列 < 0 Then Offset列 = 0
    
    With シート
        '以前までの情報の全削除
        .UsedRange.ClearContents
        If 書式設定する Then
            'セルの枠線の全削除
            .Cells.Borders.LineStyle = xlLineStyleNone
            .Cells.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
            .Cells.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
            '書式設定
            If Offset行 >= 0 Then .Cells(1, UBound(出力配列, 2)).Offset(Offset行, Offset列).HorizontalAlignment = xlRight
            .Range("A2").Offset(Offset行, Offset列).Resize(UBound(出力配列, 1), UBound(出力配列, 2)).Borders.LineStyle = xlContinuous
            .Range("A2").Offset(Offset行, Offset列).Resize(UBound(出力配列, 1), UBound(出力配列, 2)).Borders(xlInsideHorizontal).LineStyle = xlDot
            .Range("A2").Offset(Offset行, Offset列).Resize(1, UBound(出力配列, 2)).Borders(xlEdgeBottom).LineStyle = xlDouble
        End If
        '更新した情報の出力
        .Range("A2").Offset(Offset行, Offset列).Resize(UBound(出力配列, 1), UBound(出力配列, 2)) = 出力配列
        If 書式設定する Then
            .Range("A2").Offset(Offset行, Offset列).AutoFilter
            .UsedRange.EntireColumn.AutoFit
        End If
        If Offset行 >= 0 Then .Range("A1").Offset(Offset行, Offset列).Value = タイトル
        If Offset行 >= 0 Then .Cells(1, UBound(出力配列, 2)).Offset(Offset行, Offset列).Value = "更新日時: " & Now
        If 書式設定する Then
            'ウィンドウ枠の固定(対象のシートはVisible = xlSheetVisibleである必要がある)
            .Select
            If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False
            .Range("B3").Offset(Offset行, Offset列).Select
            ActiveWindow.FreezePanes = True
        End If
    End With

    '--------------------------------------------------------------------------------------------------------------------------
    'エラー処理（後半）
    Exit Function
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
    '--------------------------------------------------------------------------------------------------------------------------
End Function
