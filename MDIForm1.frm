VERSION 5.00
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "�����뿳ɱ:ս�� �籾�༭��"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10575
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Menu mFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mSelectMod 
         Caption         =   "ѡ��籾(&M)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mSave 
         Caption         =   "����籾(&S)"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "���Ϊ�籾(&A)"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mDash 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mQuit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mWindows 
      Caption         =   "����(&W)"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mEditor 
      Caption         =   "�༭��(&E)"
      Enabled         =   0   'False
      Begin VB.Menu mEditors 
         Caption         =   "�༭��"
         Index           =   0
      End
   End
   Begin VB.Menu mTool 
      Caption         =   "����(&T)"
      Enabled         =   0   'False
      Begin VB.Menu mTools 
         Caption         =   "����"
         Index           =   0
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mHelpTheme 
         Caption         =   "��������(&H)"
      End
      Begin VB.Menu mDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "����MnBWarband Editor(&A)"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub MDIForm_Load()
Dim i As Integer

Me.Show

Welcome.Show

'Load Menus

For i = 1 To 7
   If i > 0 Then
      Load mEditors(i)
      
      With mEditors(i)
        Select Case i
            Case 1
               .Caption = "����(&T)"
            Case 2
               .Caption = "��Ʒ(&I)"
            Case 3
               .Caption = "����(&P)"
            Case 4
               .Caption = "����ģ��(&M)"
            Case 5
               .Caption = "��Ӫ(&F)"
            Case 6
               .Caption = "����(&S)"
            Case 7
               .Caption = "���ͼͼ��(&C)"
        End Select
        
        .Visible = True
        .Tag = "edit_" & i
      End With

    End If
Next i

mEditors(0).Visible = False

For i = 1 To 1
   If i > 0 Then
      Load mTools(i)
      
      With mTools(i)
        Select Case i
            Case 1
               .Caption = "�������ǵ�ͼ(&C)"
        End Select
        
        .Visible = True
        .Tag = "tool_" & i
      End With

    End If
Next i

mTools(0).Visible = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("ȷ���˳� �������뿳ɱ:ս�š��籾�༭�� ��?", vbYesNo + vbExclamation, "ȷ���˳�") = vbNo Then
    Cancel = True
End If
End Sub

Private Sub mEditors_Click(Index As Integer)
Select Case mEditors(Index).Tag
      Case "edit_1"
          frmTroops.Show
      Case "edit_2"
          frmItems.Show
      Case "edit_3"
          frmParties.Show
      Case "edit_4"
          frmParty_Templates.Show
      Case "edit_5"
          frmFactions.Show
      Case "edit_6"
          frmScenes.Show
      Case "edit_7"
          frmMap_Icons.Show
End Select
End Sub

Private Sub mQuit_Click()
Unload Me
End Sub

Public Sub mSave_Click()

If MsgBox("ȷ���������趨����籾[" & MnBInfo.ModName & "]?", vbInformation + vbYesNo, "����籾") = vbYes Then
    MsgBox "��ȷ�ϱ��ݾ籾[" & MnBInfo.ModName & "]��,����ȷ������ʼ���档", vbInformation, "����籾"
    
    frmTip.ShowTip "������,���Ժ�..."
    DoEvents
    
    SaveAll
    frmTip.HideTip
        
    MsgBox "�籾[" & MnBInfo.ModName & "]����ɹ�!", vbInformation, "����籾"
End If
End Sub

Private Sub mSaveAs_Click()
frmSaveAs.Show
End Sub

Private Sub mSelectMod_Click()

Unload frmFactions
Unload frmTroops
Unload frmParties
Unload frmParty_Templates
Unload frmItems
Unload frmScenes

Unload frmMap

Unload frmModule
Unload frmSaveAs

Welcome.Show
End Sub



Private Sub mTools_Click(Index As Integer)
Select Case mTools(Index).Tag
      Case "tool_1"
          frmMap.Show

End Select
End Sub
