VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "日付範囲指定"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3165
   OleObjectBlob   =   "日付範囲指定ユーザーフォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'--------------------------------------------------------------------------------------------------------------------------
'エラー処理用
Private Const モジュール名 As String = "UserForm1"
'--------------------------------------------------------------------------------------------------------------------------
'参考サイト:
'【VBA×WindowsAPI】UserFormに日付選択コントロール(カレンダー)を作成する
'https://liclog.net/date-picker-function-vba-api/#google_vignette
'Userform1にcombobox2つとcommandbutton1つで動く
'--------------------------------------------------------------------------------------------------------------------------
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDc As LongPtr, ByVal nIndex As Long) As Long

Private Const DTS_SHORTDATEFORMAT = &H0     'YYYY/MM/DD
Private Const DTS_LONGDATEFORMAT = &H4      'YYYY年MM月DD日

Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_GROUP = &H20000

Private Const DTM_FIRST = &H1000
Private Const DTM_GETSYSTEMTIME = (DTM_FIRST + 1)   'ｺﾝﾄﾛｰﾙの日付/時刻を取得
Private Const DTM_SETSYSTEMTIME = (DTM_FIRST + 2)   'ｺﾝﾄﾛｰﾙの日付/時刻をｾｯﾄ
Private Const DTM_GETRANGE = (DTM_FIRST + 3)        'ｺﾝﾄﾛｰﾙの日付範囲を取得
Private Const DTM_SETRANGE = (DTM_FIRST + 4)        'ｺﾝﾄﾛｰﾙの日付範囲を設定

Private 開始日_選択 As Date
Private 終了日_選択 As Date

'ｼｽﾃﾑﾀｲﾑ構造体
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private hWndDate_1, hWndDate_2 As LongPtr

'------------------------------------------------------------------
'   UserForm起動時ｲﾍﾞﾝﾄ
'------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo エラー時
    Const プロシージャ名 As String = "UserForm_Initialize"

    Dim hWndForm As LongPtr
    Dim hWndClient As LongPtr
    Dim システム時刻 As SYSTEMTIME
    Dim 今月初日, 今月末日 As Date
        
    'UserFormのｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙ取得
    hWndForm = FindWindow("ThunderDFrame", Me.Caption)
    hWndClient = FindWindowEx(hWndForm, 0, vbNullString, vbNullString)

    '日付選択ｺﾝﾄﾛｰﾙ作成
    hWndDate_1 = CreateWindowEx(0, "SysDateTimePick32", vbNullString, _
                              WS_CHILD Or WS_VISIBLE Or DTS_SHORTDATEFORMAT Or WS_GROUP, _
                              PtToPx(Me.ComboBox1.Left), PtToPx(Me.ComboBox1.Top), _
                              PtToPx(Me.ComboBox1.Width), PtToPx(Me.ComboBox1.Height), _
                              hWndClient, 0, 0, 0)
    hWndDate_2 = CreateWindowEx(0, "SysDateTimePick32", vbNullString, _
                              WS_CHILD Or WS_VISIBLE Or DTS_SHORTDATEFORMAT Or WS_GROUP, _
                              PtToPx(Me.ComboBox2.Left), PtToPx(Me.ComboBox2.Top), _
                              PtToPx(Me.ComboBox2.Width), PtToPx(Me.ComboBox2.Height), _
                              hWndClient, 0, 0, 0)
    
    '初期値を今月にする
    '今月初日と今月末日を計算
    今月初日 = DateSerial(Year(Date), Month(Date), 1)
    今月末日 = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    'hWndDate_1の日付を今月の初日に設定
    システム時刻.wYear = Year(今月初日)
    システム時刻.wMonth = Month(今月初日)
    システム時刻.wDay = 1
    SendMessage hWndDate_1, DTM_SETSYSTEMTIME, 0, システム時刻
    
    'hWndDate_2の日付を今月の末日に設定
    システム時刻.wDay = Day(今月末日)
    SendMessage hWndDate_2, DTM_SETSYSTEMTIME, 0, システム時刻

    Exit Sub
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo エラー時
    Const プロシージャ名 As String = "UserForm_QueryClose"
    
    If (CloseMode = vbFormControlMenu) Then
        End
    End If
    
    Exit Sub
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
End Sub

'------------------------------------------------------------------
'   UserForm終了時ｲﾍﾞﾝﾄ
'------------------------------------------------------------------
Private Sub UserForm_Terminate()
    On Error GoTo エラー時
    Const プロシージャ名 As String = "UserForm_Terminate"
 
    '日付選択ｺﾝﾄﾛｰﾙを破棄
    Call DestroyWindow(hWndDate_1)
    Call DestroyWindow(hWndDate_2)
 
     Exit Sub
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
End Sub

'------------------------------------------------------------------
'   ｺﾏﾝﾄﾞﾎﾞﾀﾝ押下ｲﾍﾞﾝﾄ
'------------------------------------------------------------------
Private Sub CommandButton1_Click()
    On Error GoTo エラー時
    Const プロシージャ名 As String = "CommandButton1_Click"

    Dim システム時刻 As SYSTEMTIME
    
    SendMessage hWndDate_1, DTM_GETSYSTEMTIME, 0, システム時刻
    開始日_選択 = DateSerial(システム時刻.wYear, システム時刻.wMonth, システム時刻.wDay)
    
    SendMessage hWndDate_2, DTM_GETSYSTEMTIME, 0, システム時刻
    終了日_選択 = DateSerial(システム時刻.wYear, システム時刻.wMonth, システム時刻.wDay)
    
    Me.Hide
    
    Exit Sub
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
End Sub

'------------------------------------------------------------------
'   ﾎﾟｲﾝﾄ(pt)→ﾋﾟｸｾﾙ(px)変換
'------------------------------------------------------------------
Function PtToPx(ByVal dPt As Double) As Double
    On Error GoTo エラー時
    Const プロシージャ名 As String = "PtToPx"

    Dim hDc As LongPtr
    Dim lDpiX As Long
    
    hDc = GetDC(0)
    lDpiX = GetDeviceCaps(hDc, 88) '88 = LOGPIXELSX
    Call ReleaseDC(0, hDc)
    
    PtToPx = dPt * lDpiX / 72
    
    Exit Function
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
End Function

Public Property Get 開始日() As Date
    On Error GoTo エラー時
    Const プロシージャ名 As String = "開始日"
    
    開始日 = 開始日_選択

    Exit Property
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
End Property
Public Property Get 終了日() As Date
    On Error GoTo エラー時
    Const プロシージャ名 As String = "終了日"
    
    終了日 = 終了日_選択

    Exit Property
エラー時:
    Call エラー処理.Setエラー(モジュール名, プロシージャ名)
    Call Err.Raise(Err.Number, , Err.Description)
End Property
