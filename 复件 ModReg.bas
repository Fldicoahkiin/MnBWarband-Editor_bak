Attribute VB_Name = "ModReg"
Dim Fso As Object

'��������
    Const HKEY_CLASSES_ROOT = -2147483648#
    Const HKEY_CURRENT_USER = -2147483647#
    Const HKEY_LOCAL_MACHINE = -2147483646#
    Const HKEY_USERS = -2147483645#

    '��ֵ����
    Const REG_SZ = 1& '�ַ���ֵ
    Const REG_BINARY = 3& '������ֵ
    Const REG_DWORD = 4& 'DWORD ֵ

    '�����й�API����
    Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
        ( _
          ByVal hKey As Long, _
          ByVal lpSubKey As String, _
          ByRef phkResult As Long _
          ) As Long '����һ���µ�����

    Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
        ( _
          ByVal hKey As Long, _
          ByVal lpSubKey As String, _
          ByRef phkResult As Long _
        ) As Long '��һ������

    Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
      ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String _
      ) As Long 'ɾ��һ������

    Public Declare Function RegCloseKey Lib "advapi32.dll" _
      ( _
        ByVal hKey As Long _
      ) As Long '�ر�һ������

    Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
      ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByVal lpData As Any, _
        ByVal cbData As Long _
      ) As Long '������ı�һ����ֵ,lpDataӦ��ȱʡ��ByRef�͸�ΪByVal��

    Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
      ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByVal lpData As Any, _
        ByRef lpcbData As Long _
      ) As Long '��ѯһ����ֵ,lpDataӦ��ȱʡ��ByRef�͸�ΪByVal��

    Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
      ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String _
        ) As Long 'ɾ��һ����ֵ

Public Sub InitModReg()

Set Fso = CreateObject("WScript.Shell")

End Sub

Public Sub UnloadModReg()

Set Fso = Nothing

End Sub

Public Sub RegProtection()
Dim i As Long, V As Variant
For i = 0 To UBound(RegProtectionObjects)
   With RegProtectionObjects(i)
    V = RegRead(.KeyRoot, .KeyName)
    If V <> .KeyValue Then
    RegWrite .KeyRoot, .KeyName, .KeyType, .KeyValue
    End If
   End With
Next i
End Sub


Public Sub RegWrite(KeyRoot As String, KeyName As String, KeyType As String, CMD As Variant)
On Error Resume Next
Dim Type_Name As String, strKey As Variant, Key_Name As String

   Type_Name = KeyType
   Key_Name = KeyName

   strKey = CMD
   Fso.RegWrite KeyRoot & "\" & Key_Name, strKey, Type_Name

End Sub

Public Function RegRead(KeyRoot As String, KeyName As String) As Variant
On Error GoTo Errline
Dim TemS As String


RegRead = Fso.RegRead(KeyRoot & "\" & KeyName)
Exit Function

Errline:
Call logErr("ModReg", "RegRead", Err.Number, Err.Description)
End Function


Public Sub RegDelete(KeyRoot As String, KeyName As String)
On Error Resume Next

   Fso.RegDelete KeyRoot & "\" & Key_Name

End Sub



    '������
    Sub Main()
      Dim nKeyHandle As Long, nValueType As Long, nLength As Long
      Dim sValue As String
      sValue = "I am a winner!"
      Call RegCreateKey(HKEY_CURRENT_USER, "New ReGIStry Key", nKeyHandle)
      Call RegSetValueEx(nKeyHandle, "My Value", 0, REG_SZ, sValue, 255)
      sValue = Space(255)
      nLength = 255
      Call RegQueryValueEx(nKeyHandle, "My Value", 0, nValueType, sValue, nLength)
      MsgBox sValue
      Call RegDeleteValue(nKeyHandle, "My Value")
      Call RegDeleteKey(HKEY_CURRENT_USER, "New Registry Key")
      Call RegCloseKey(nKeyHandle)
    End Sub

