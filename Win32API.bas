Attribute VB_Name = "Win32API"
Option Explicit

'API
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long '��ȡ������ʽ
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '���ô�����ʽ
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long


Public Type SHFILEOPSTRUCT
    hWnd   As Long
    wFunc   As Long
    pFrom   As String
    pTo   As String
    fFlags   As Integer
    fAborted   As Boolean
    hNameMaps   As Long
    sProgress   As String
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LVHITTESTINFO
   pt As POINTAPI
   Flags As Long
   iItem As Long
   iSubItem As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetPrivateProfileInt Lib "kernel32" _
    Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal nDefault As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Public ErrorMsg As String

Public Const WM_USER = &H400

Public Const LB_SETHORIZONTALEXTENT = WM_USER + 21
Public Const LB_GETSELITEMS = &H191
Public Const SPI_GETWORKAREA = 48

'�����
Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10
Public Const VK_MENU = &H12


'��Ϣ����
Public Const GWL_STYLE = (-16) '������ʽ
Public Const WS_CAPTION = &HC00000 '�����ⴰ��
Public Const WS_MAXIMIZEBOX = &H10000 '����󻯰�ť
Public Const WS_MINIMIZEBOX = &H20000 '����С��
Public Const WS_THICKFRAME = &H40000 '����ɵ�
Public Const MY_NOT_SIZABLE = 382337024  '���岻�ɵ�
Public Const MY_SIZABLE = 382664704 '���岻�ɵ�
 
Public Const FO_DELETE = &H3

Public Const FOF_ALLOWUNDO = &H40           '   �������վ
Public Const FOF_CONFIRMMOUSE = &H2           '   ɾ�������������վ
Public Const FOF_NOCONFIRMATION = &H10           '   û����ʾ

Public Const EM_SETTARGETDEVICE = (WM_USER + 72)

'�ı������
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB

'ListView
Public Const LVM_FIRST = &H1000
Public Const LVM_HITTEST = LVM_FIRST + 18
Public Const LVM_GETCOUNTPERPAGE = LVM_FIRST + 40
Public Const LVM_SETCOLUMNWIDTH = &H1000 + 30
Public Const LVSCW_AUTOSIZE_USEHEADER = -2


'*************************************************************************
'**�� �� ����ErrorMsg_Initialize
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:43:13
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.9
'*************************************************************************
Public Sub ErrorMsg_Initialize()
    Call Class_Initialize
End Sub

'*************************************************************************
'**�� �� ����Class_Initialize
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:43:06
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.9
'*************************************************************************
Private Sub Class_Initialize()
    ErrorMsg = vbNullString
End Sub

'*************************************************************************
'**�� �� ����WriteString
'**��    �룺iniFileName(String) -
'**        ��Section(String)     -
'**        ��Key(String)         -
'**        ��Value(String)       -
'**��    ����(Boolean) -
'**����������д��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:42:48
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.13
'*************************************************************************
Public Function WriteString(iniFileName As String, Section As String, Key As String, Value As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim F As Integer
    
    WriteString = False
    ErrorMsg = vbNullString
    If Not FileExists(iniFileName) Then
        'ErrorMsg = "INI file has not been specifyed!"
        'Exit Function
        F = FreeFile
        
        Open iniFileName For Output As #F
        Close #F
    End If
    If WritePrivateProfileString(Section, Key, Value, iniFileName) = 0 Then
        ErrorMsg = "Failed to write to the ini file!"
        Exit Function
    End If
 
    WriteString = True
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Win32API", "WriteString", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����ReadString
'**��    �룺iniFileName(String) -
'**        ��Section(String)     -
'**        ��Key(String)         -
'**        ��Size(Long)          -
'**��    ����(String) -
'**���������������ַ���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:41:43
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.9
'*************************************************************************
Public Function ReadString(iniFileName As String, Section As String, Key As String, Optional Size As Long = 255) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim ReturnStr As String
    Dim ReturnLng As Long
    ErrorMsg = vbNullString
    ReadString = vbNullString
    If Not FileExists(iniFileName) Then
        ErrorMsg = "INI file has not been specifyed!"
        Exit Function
    End If
    ReturnStr = Space(Size)
    ReturnLng = GetPrivateProfileString(Section, Key, vbNullString, ReturnStr, Size, iniFileName)
    'ReadString = Left(ReturnStr, ReturnLng)
    ReadString = StripTerFlag(ReturnStr)
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Win32API", "ReadString", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����ReadInt
'**��    �룺iniFileName(String) -
'**        ��Section(String)     -
'**        ��Key(String)         -
'**��    ����(Long) -
'**����������������ֵ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:41:50
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.9
'*************************************************************************
Public Function ReadInt(iniFileName As String, Section As String, Key As String) As Long
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim ReturnLng As Long
    ReadInt = 0
    ErrorMsg = vbNullString
    If iniFileName = "" Then
        ErrorMsg = "INI file has not been specifyed!"
        Exit Function
    End If
    ReturnLng = GetPrivateProfileInt(Section, Key, 0, iniFileName)
    If ReturnLng = 0 Then
        ReturnLng = GetPrivateProfileInt(Section, Key, 1, iniFileName)
        If ReturnLng = 1 Then
            ErrorMsg = "Can not read the ini file!"
            Exit Function
        End If
    End If
    ReadInt = ReturnLng
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Win32API", "ReadInt", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����ChangeSectionName
'**��    �룺iniFileName(String) -
'**        ��oSection(String)     -
'**        ��nSection(String)         -
'**��    ����(Long) -
'**�������������ļ���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-06-16 14:01:43
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function ChangeSectionName(iniFileName As String, oSection As String, nSection As String) As Long
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim F As Long
    
    F = FreeFile
    
    Open iniFileName For Input As #F
      
    Close #F
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Win32API", "ReadInt", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����GetMyDocumentDirectory
'**��    �룺Null
'**��    ����(Long) -
'**����������������ֵ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:41:50
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.9
'*************************************************************************
Public Function GetMyDocumentDirectory() As String
Dim ObjTem As Object

Set ObjTem = CreateObject("Shell.Application")

GetMyDocumentDirectory = ObjTem.NameSpace(5).Self.Path

End Function

'*************************************************************************
'**�� �� ����SwitchSizable
'**��    �룺(Form)ObjForm
'**��    ����(Long) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-30 21:25:58
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************

Public Sub SwitchSizable(ObjForm As Form)
  Dim WinStyle     As Long
  WinStyle = GetWindowLong(ObjForm.hWnd, GWL_STYLE) 'ȡ�ô�����ʽ
  SetWindowLong ObjForm.hWnd, GWL_STYLE, WinStyle Xor (WS_MAXIMIZEBOX Or WS_THICKFRAME)  '������ʽ ��ԭ����ʽ ��� �����С��������3��ʽ���� Ҳ����ԭ���еľͳ�û�� ��֮��Ȼ
  
  ObjForm.Hide
  ObjForm.Show
End Sub

'*************************************************************************
'**�� �� ����SetFormStyle
'**��    �룺(Form)ObjForm,(Long)Style
'**��    ����(Long) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-30 22:17:15
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************

Public Sub SetFormStyle(ObjForm As Form, ByVal Style As Long)
  SetWindowLong ObjForm.hWnd, GWL_STYLE, Style   '������ʽ ��ԭ����ʽ ��� �����С��������3��ʽ���� Ҳ����ԭ���еľͳ�û�� ��֮��Ȼ
  ObjForm.Hide
  ObjForm.Show
End Sub


'*************************************************************************
'**�� �� ����DeleteFolder
'**��    �룺(String)sObject
'**��    ����-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-01 23:06:09
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub DeleteFolder(sObject As String)
    Dim SHFileOp     As SHFILEOPSTRUCT

    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sObject
        .fFlags = FOF_CONFIRMMOUSE Or FOF_NOCONFIRMATION
    End With
    SHFileOperation SHFileOp
End Sub


'*************************************************************************
'**�� �� ����GetListViewVisibleLines
'**��    �룺(ListView)ListView
'**��    ����(Long)-
'**�������������listview�ɼ�����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-11 23:15:45
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetListViewVisibleLines(ListView As ListView) As Long
    GetListViewVisibleLines = SendMessage(ListView.hWnd, LVM_GETCOUNTPERPAGE, 0, 0)
End Function


'*************************************************************************
'**�� �� ����GetListViewItemIndexUnderMousePointer
'**��    �룺(ListView)ListView
'**��    ����(Long)-
'**�������������listview���ָ���µ���Ŀindex
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-11 23:47:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetListViewItemIndexUnderMousePointer(ListView As ListView, X As Long, Y As Long) As Long
Dim lvhti As LVHITTESTINFO

   lvhti.pt.X = X / Screen.TwipsPerPixelX
   lvhti.pt.Y = Y / Screen.TwipsPerPixelY
   GetListViewItemIndexUnderMousePointer = SendMessage(ListView.hWnd, LVM_HITTEST, 0, lvhti) + 1
End Function

'*************************************************************************
'**�� �� ����GetTaskbarHeight
'**��    �룺(ListView)ListView
'**��    ����(Long)-
'**������������ÿ�ʼ�˵����߶�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-14 17:47:34
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

'*************************************************************************
'**�� �� ����StripTerFlag
'**��    �룺(String)Str
'**��    ����(String)-
'**������������ý�����ʶ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-05-26 22:25:52
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function StripTerFlag(Str As String) As String
Dim n As Integer
n = InStr(Str, Chr$(0))
If n > 0 Then
    StripTerFlag = Left(Str, n - 1)
Else
    StripTerFlag = Str
End If
End Function

'*************************************************************************
'**�� �� ����AutoSwitchLine
'**��    �룺(Long)RichText , (Boolean)bSwitch
'**��    ����-
'**�������������� RichTextBox �Զ�����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�TechnoFantasy
'**��    �ڣ�2005-02-05 23:25:19
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub AutoSwitchLine(ByRef RichText As RichTextBox, ByVal bSwitch As Boolean)
    If bSwitch Then
        '���� RichTextBox �Զ�����
        Call SendMessage(RichText.hWnd, EM_SETTARGETDEVICE, _
        GetDC(RichText.hWnd), RichText.Width / 15)
        If RichText.RightMargin = 0 Then
            RichText.RightMargin = 1
        Else
            RichText.RightMargin = 0
        End If
    Else
        '���� RichTextBox ���Զ�����
        Call SendMessage(RichText.hWnd, EM_SETTARGETDEVICE, 0, 1)
    End If
End Sub
