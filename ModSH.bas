Attribute VB_Name = "ModSH"
Option Explicit

Public Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long

Public Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const MAX_LEN = 200 '�ַ�����󳤶�

Public Const DESKTOP = &H0& '����

Public Const PROGRAMS = &H2& '����

Public Const MYDOCUMENTS = &H5& '�ҵ��ĵ�

Public Const MYFAVORITES = &H6& '�ղؼ�

Public Const STARTUP = &H7& '����

Public Const RECENT = &H8& '����򿪵��ļ�

Public Const SENDTO = &H9& '����

Public Const STARTMENU = &HB& '��ʼ�˵�

Public Const NETHOOD = &H13& '�����ھ�

Public Const FONTS = &H14& '����

Public Const SHELLNEW = &H15& 'ShellNew

Public Const APPDATA = &H1A& 'Application Data

Public Const PRINTHOOD = &H1B& 'PrintHood

Public Const PAGETMP = &H20& '��ҳ��ʱ�ļ�

Public Const COOKIES = &H21& 'CookiesĿ¼

Public Const HISTORY = &H22& '��ʷ

Public Function GetAppDataPath()
Dim sTmp As String * MAX_LEN '��Ž���Ĺ̶����ȵ��ַ���

Dim pidl As Long 'ĳ����Ŀ¼������Ŀ¼�б��е�λ��


SHGetSpecialFolderLocation 0, APPDATA, pidl

SHGetPathFromIDList pidl, sTmp

GetAppDataPath = Left(sTmp, InStr(sTmp, Chr(0)) - 1)

End Function
