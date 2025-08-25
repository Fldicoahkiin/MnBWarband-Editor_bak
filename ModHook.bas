Attribute VB_Name = "ModHook"
'ģ�����
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'���ô�����Ϣ
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'���ô��ڴ���
Private Const WM_RBUTTONDOWN = &H204
Private Const GWL_WNDPROC = (-4)
Private lpOldWndFunc As Long
'--------------------------------------------------------------------------------------
'�� �� ��: WindowProcedure
'��    ��: ������Ϣ������
'--------------------------------------------------------------------------------------
Private Function WindowProcedure(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Select Case wMsg
               Case WM_RBUTTONDOWN
               Case Else
                    WindowProcedure = CallWindowProc(lpOldWndFunc, hWnd, wMsg, wParam, _
                    lParam)                                                             'ԭ������Ϣ����
        End Select
End Function
'--------------------------------------------------------------------------------------
'�� �� ��: Hook
'��    ��: ���໯����
'--------------------------------------------------------------------------------------
Public Sub Hook(ByVal hWnd As Long)
       lpOldWndFunc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProcedure)       '�´�����Ϣ����
End Sub
'--------------------------------------------------------------------------------------
'�� �� ��: UnHook
'��    ��: ȡ�����໯
'--------------------------------------------------------------------------------------
Public Sub UnHook(ByVal hWnd As Long)
       Call SetWindowLong(hWnd, GWL_WNDPROC, lpOldWndFunc)                              '�ָ�ԭ������Ϣ����
End Sub


