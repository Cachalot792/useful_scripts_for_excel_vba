VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "���t�͈͎w��"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3165
   OleObjectBlob   =   "���t�͈͎w�胆�[�U�[�t�H�[��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'--------------------------------------------------------------------------------------------------------------------------
'�G���[�����p
Private Const ���W���[���� As String = "UserForm1"
'--------------------------------------------------------------------------------------------------------------------------
'�Q�l�T�C�g:
'�yVBA�~WindowsAPI�zUserForm�ɓ��t�I���R���g���[��(�J�����_�[)���쐬����
'https://liclog.net/date-picker-function-vba-api/#google_vignette
'Userform1��combobox2��commandbutton1�œ���
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
Private Const DTS_LONGDATEFORMAT = &H4      'YYYY�NMM��DD��

Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_GROUP = &H20000

Private Const DTM_FIRST = &H1000
Private Const DTM_GETSYSTEMTIME = (DTM_FIRST + 1)   '���۰ق̓��t/�������擾
Private Const DTM_SETSYSTEMTIME = (DTM_FIRST + 2)   '���۰ق̓��t/�������
Private Const DTM_GETRANGE = (DTM_FIRST + 3)        '���۰ق̓��t�͈͂��擾
Private Const DTM_SETRANGE = (DTM_FIRST + 4)        '���۰ق̓��t�͈͂�ݒ�

Private �J�n��_�I�� As Date
Private �I����_�I�� As Date

'������э\����
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
'   UserForm�N���������
'------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo �G���[��
    Const �v���V�[�W���� As String = "UserForm_Initialize"

    Dim hWndForm As LongPtr
    Dim hWndClient As LongPtr
    Dim �V�X�e������ As SYSTEMTIME
    Dim ��������, �������� As Date
        
    'UserForm�̳���޳����َ擾
    hWndForm = FindWindow("ThunderDFrame", Me.Caption)
    hWndClient = FindWindowEx(hWndForm, 0, vbNullString, vbNullString)

    '���t�I����۰ٍ쐬
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
    
    '�����l�������ɂ���
    '���������ƍ����������v�Z
    �������� = DateSerial(Year(Date), Month(Date), 1)
    �������� = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    'hWndDate_1�̓��t�������̏����ɐݒ�
    �V�X�e������.wYear = Year(��������)
    �V�X�e������.wMonth = Month(��������)
    �V�X�e������.wDay = 1
    SendMessage hWndDate_1, DTM_SETSYSTEMTIME, 0, �V�X�e������
    
    'hWndDate_2�̓��t�������̖����ɐݒ�
    �V�X�e������.wDay = Day(��������)
    SendMessage hWndDate_2, DTM_SETSYSTEMTIME, 0, �V�X�e������

    Exit Sub
�G���[��:
    Call �G���[����.Set�G���[(���W���[����, �v���V�[�W����)
    Call Err.Raise(Err.Number, , Err.Description)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo �G���[��
    Const �v���V�[�W���� As String = "UserForm_QueryClose"
    
    If (CloseMode = vbFormControlMenu) Then
        End
    End If
    
    Exit Sub
�G���[��:
    Call �G���[����.Set�G���[(���W���[����, �v���V�[�W����)
    Call Err.Raise(Err.Number, , Err.Description)
End Sub

'------------------------------------------------------------------
'   UserForm�I���������
'------------------------------------------------------------------
Private Sub UserForm_Terminate()
    On Error GoTo �G���[��
    Const �v���V�[�W���� As String = "UserForm_Terminate"
 
    '���t�I����۰ق�j��
    Call DestroyWindow(hWndDate_1)
    Call DestroyWindow(hWndDate_2)
 
     Exit Sub
�G���[��:
    Call �G���[����.Set�G���[(���W���[����, �v���V�[�W����)
    Call Err.Raise(Err.Number, , Err.Description)
End Sub

'------------------------------------------------------------------
'   ��������݉��������
'------------------------------------------------------------------
Private Sub CommandButton1_Click()
    On Error GoTo �G���[��
    Const �v���V�[�W���� As String = "CommandButton1_Click"

    Dim �V�X�e������ As SYSTEMTIME
    
    SendMessage hWndDate_1, DTM_GETSYSTEMTIME, 0, �V�X�e������
    �J�n��_�I�� = DateSerial(�V�X�e������.wYear, �V�X�e������.wMonth, �V�X�e������.wDay)
    
    SendMessage hWndDate_2, DTM_GETSYSTEMTIME, 0, �V�X�e������
    �I����_�I�� = DateSerial(�V�X�e������.wYear, �V�X�e������.wMonth, �V�X�e������.wDay)
    
    Me.Hide
    
    Exit Sub
�G���[��:
    Call �G���[����.Set�G���[(���W���[����, �v���V�[�W����)
    Call Err.Raise(Err.Number, , Err.Description)
End Sub

'------------------------------------------------------------------
'   �߲��(pt)���߸��(px)�ϊ�
'------------------------------------------------------------------
Function PtToPx(ByVal dPt As Double) As Double
    On Error GoTo �G���[��
    Const �v���V�[�W���� As String = "PtToPx"

    Dim hDc As LongPtr
    Dim lDpiX As Long
    
    hDc = GetDC(0)
    lDpiX = GetDeviceCaps(hDc, 88) '88 = LOGPIXELSX
    Call ReleaseDC(0, hDc)
    
    PtToPx = dPt * lDpiX / 72
    
    Exit Function
�G���[��:
    Call �G���[����.Set�G���[(���W���[����, �v���V�[�W����)
    Call Err.Raise(Err.Number, , Err.Description)
End Function

Public Property Get �J�n��() As Date
    On Error GoTo �G���[��
    Const �v���V�[�W���� As String = "�J�n��"
    
    �J�n�� = �J�n��_�I��

    Exit Property
�G���[��:
    Call �G���[����.Set�G���[(���W���[����, �v���V�[�W����)
    Call Err.Raise(Err.Number, , Err.Description)
End Property
Public Property Get �I����() As Date
    On Error GoTo �G���[��
    Const �v���V�[�W���� As String = "�I����"
    
    �I���� = �I����_�I��

    Exit Property
�G���[��:
    Call �G���[����.Set�G���[(���W���[����, �v���V�[�W����)
    Call Err.Raise(Err.Number, , Err.Description)
End Property
