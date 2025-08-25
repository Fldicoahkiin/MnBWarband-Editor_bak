VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Info"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   Enabled         =   0   'False
   FillColor       =   &H80000012&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Title"
      Top             =   0
      Width           =   2055
   End
   Begin MSComctlLib.ListView ListInfo 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3413
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16761024
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const OriginWidth = 2600

Private Sub Form_Deactivate()
Me.ZOrder
End Sub

Private Sub Form_Load()

Me.Visible = False

InitListView

With txtTitle
  .Top = 0
  .Left = 0
  .Width = Me.ScaleWidth
End With

With ListInfo
  .Top = txtTitle.Top + txtTitle.Height
  .Left = 0
  .Width = Me.ScaleWidth
  .Height = Me.ScaleHeight - txtTitle.Top - txtTitle.Height
End With

End Sub

'*************************************************************************
'**�� �� ����ShowfrmInfo
'**��    �룺ItemID(-1��Ϊ��ǰ��Ʒ)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-10 15:24:50
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub ShowfrmInfo(X As Single, Y As Single)

Me.Left = X
Me.Top = Y
Me.Visible = True
ResetFormPos

End Sub

Private Sub InitListView()
Dim n As Integer
n = 2
ListInfo.View = lvwReport
ListInfo.Sorted = False
ListInfo.ListItems.Clear
ListInfo.ColumnHeaders.Clear
ListInfo.SortOrder = lvwAscending
ListInfo.FullRowSelect = True
ListInfo.AllowColumnReorder = False
ListInfo.LabelEdit = lvwManual
ListInfo.Checkboxes = False
ListInfo.GridLines = False
ListInfo.MultiSelect = False
ListInfo.HideSelection = True
ListInfo.HideColumnHeaders = True

ListInfo.ColumnHeaders.Add , , , ListInfo.Width / n / 1.2
ListInfo.ColumnHeaders.Add , , , ListInfo.Width / n

End Sub
'*************************************************************************
'**�� �� ����AutoSize
'**��    �룺
'**��    ������
'**�����������Զ�����ListView���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-13 23:33:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub AutoSize()
Dim i As Integer, Width As Long, Height As Long, HerL As Long
    
    For i = 0 To ListInfo.ColumnHeaders.Count
        SendMessage ListInfo.hWnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE_USEHEADER
    Next i
    
    If ListInfo.ListItems.Count > 0 Then
       If ListInfo.ListItems(1).Width < OriginWidth Then
           ListInfo.Width = OriginWidth
       Else
           ListInfo.Width = ListInfo.ListItems(1).Width
       End If
    Else
       ListInfo.Width = OriginWidth
    End If
    
    txtTitle.Width = ListInfo.Width
    Me.Width = ListInfo.Width
    
    For i = 1 To ListInfo.ListItems.Count
       If i = 1 Then HerL = ListInfo.ListItems(i).Height
       Height = Height + ListInfo.ListItems(i).Height
    Next i
    ListInfo.Height = Height + HerL * 2
    Me.Height = ListInfo.Height
    
End Sub
'*************************************************************************
'**�� �� ����ResetFormPos
'**��    �룺
'**��    ������
'**�����������Զ���������λ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-14 17:43:33
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub ResetFormPos()
Dim TBH As Integer

TBH = GetTaskbarHeight()
If Me.Top + Me.Height > Screen.Height - TBH Then
   Me.Top = Screen.Height - Me.Height - TBH
End If

End Sub
'*************************************************************************
'**�� �� ����LoadItemInfo
'**��    �룺ItemID(-1��Ϊ��ǰ��Ʒ)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-10 15:24:50
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub LoadItemInfo(Optional ItemID As Long = -1)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim Item As Type_Item, oItem As ListItem, TemItp As Integer
Dim Dam As Long, tpDam As Integer

If ItemID = -1 Then ItemID = CurrentItmID
Item = itm(ItemID)

'ResetFormPos

txtTitle.Text = Item.csvName

With ListInfo
   .ListItems.Clear

   Set oItem = .ListItems.Add(, "itemtype", "  " & PublicMsgs(102) & ":  ")           '��Ʒ����
       TemItp = GetItmType(Item.itmType)
       If TemItp > 0 Then
          If Not IsTwoHanded(Item.itmType) And TemItp = itp_type_two_handed_wpn Then
            oItem.SubItems(1) = PublicMsgs(104)     '��/˫��
          Else
            oItem.SubItems(1) = Item_Type(TemItp).Y
          End If
       End If
   Set oItem = .ListItems.Add(, "price", "  " & PublicMsgs(103) & ":  ")           '�۸�
       oItem.SubItems(1) = Item.price
       oItem.ForeColor = vbYellow
       oItem.ListSubItems(1).ForeColor = vbYellow

   If Not (IsHorse(Item.itmType)) Then
      Set oItem = .ListItems.Add(, "weight", "  " & PublicMsgs(105) & ":  ")           '����
      oItem.SubItems(1) = Format(Item.weight, "0.00")
   End If
  '-------------------------------��ս����-----------------------------------
   If IsMeleeWeapon(Item.itmType) Then
      If Item.swing_damage > 0 Then
        Dam = GetDamage(Item.swing_damage, tpDam)
        If Dam > 0 Then
           Set oItem = .ListItems.Add(, "swing", "  " & PublicMsgs(106) & ":  ")           '�ӿ��˺�
           oItem.SubItems(1) = Dam & PublicMsgs(108 + tpDam)
           oItem.ForeColor = vbRed
           oItem.ListSubItems(1).ForeColor = vbRed
        End If
      End If
   
      If Item.thrust_damage > 0 Then
        Dam = GetDamage(Item.thrust_damage, tpDam)
        If Dam > 0 Then
           Set oItem = .ListItems.Add(, "thrust", "  " & PublicMsgs(107) & ":  ")          '�����˺�
           oItem.SubItems(1) = Dam & PublicMsgs(108 + tpDam)
           oItem.ForeColor = vbRed
           oItem.ListSubItems(1).ForeColor = vbRed
        End If
      End If
      
      Set oItem = .ListItems.Add(, "speed", "  " & PublicMsgs(111) & ":  ")          '�ٶ�
      oItem.SubItems(1) = Item.speed_rating
      
      Set oItem = .ListItems.Add(, "length", "  " & PublicMsgs(112) & ":  ")          '��Χ
      oItem.SubItems(1) = Item.weapon_length
      
      If Item.difficulty > 0 Then
         Set oItem = .ListItems.Add(, "difficulty", "  " & PublicMsgs(113) & ":  ")          '�Ѷ�
         oItem.SubItems(1) = Item.difficulty
         oItem.ForeColor = vbGreen
         oItem.ListSubItems(1).ForeColor = vbGreen
      End If
      
      Dim nFlags As Integer
      Dim LabelStr As String
      nFlags = 0
      LabelStr = "  " & PublicMsgs(114) & ":  "
      If IsBonusAgainstShield(Item.itmType) Then
         nFlags = nFlags + 1
         Set oItem = .ListItems.Add(, "flags" & nFlags, LabelStr)          '����
         LabelStr = ""
         oItem.SubItems(1) = Itp(itp_bonus_against_shield - 12).Y
         oItem.ForeColor = vbBlue
         oItem.ListSubItems(1).ForeColor = vbBlue
      End If
      
      If IsUnbalanced(Item.itmType) Then
         nFlags = nFlags + 1
         Set oItem = .ListItems.Add(, "flags" & nFlags, LabelStr)          '����
         LabelStr = ""
         oItem.SubItems(1) = Itp(itp_unbalanced - 12).Y
         oItem.ForeColor = vbBlue
         oItem.ListSubItems(1).ForeColor = vbBlue
      End If
      
       If IsCrushThrough(Item.itmType) Then
         nFlags = nFlags + 1
         Set oItem = .ListItems.Add(, "flags" & nFlags, LabelStr)          '����
         LabelStr = ""
         oItem.SubItems(1) = Itp(itp_crush_through - 12).Y
         oItem.ForeColor = vbBlue
         oItem.ListSubItems(1).ForeColor = vbBlue
      End If
      
       If IsCantUseOnHorseBack(Item.itmType) Then
         nFlags = nFlags + 1
         Set oItem = .ListItems.Add(, "flags" & nFlags, LabelStr)          '����
         LabelStr = ""
         oItem.SubItems(1) = Itp(itp_cant_use_on_horseback - 12).Y
         oItem.ForeColor = vbBlue
         oItem.ListSubItems(1).ForeColor = vbBlue
      End If
  '-------------------------------Զ������-----------------------------------
   ElseIf IsRangedWeapon(Item.itmType) Then
      If Item.thrust_damage > 0 Then
        Dam = GetDamage(Item.thrust_damage, tpDam)
        If Dam > 0 Then
           Set oItem = .ListItems.Add(, "thrust", "  " & PublicMsgs(107) & ":  ")          '�����˺�
           oItem.SubItems(1) = Dam & PublicMsgs(108 + tpDam)
           oItem.ForeColor = vbRed
           oItem.ListSubItems(1).ForeColor = vbRed
        End If
      End If
      
      Set oItem = .ListItems.Add(, "accuracy", "  " & PublicMsgs(115) & ":  ")          '����
           oItem.SubItems(1) = Item.leg_armor
           oItem.ForeColor = &HFF00FF
           oItem.ListSubItems(1).ForeColor = &HFF00FF
      
      Set oItem = .ListItems.Add(, "speed", "  " & PublicMsgs(111) & ":  ")          '�ٶ�
           oItem.SubItems(1) = Item.speed_rating
           
      Set oItem = .ListItems.Add(, "mspeed", "  " & PublicMsgs(116) & ":  ")          '����
           oItem.SubItems(1) = Item.missile_speed
                     
      If Item.difficulty > 0 Then
         Set oItem = .ListItems.Add(, "difficulty", "  " & PublicMsgs(113) & ":  ")          '�Ѷ�
         oItem.SubItems(1) = Item.difficulty
         oItem.ForeColor = vbGreen
         oItem.ListSubItems(1).ForeColor = vbGreen
      End If
      
      'nFlags = 0
      LabelStr = "  " & PublicMsgs(114) & ":  "
      If IsCantReloadOnHorseBack(Item.itmType) Then
         'nFlags = nFlags + 1
         Set oItem = .ListItems.Add(, "flags" & nFlags, LabelStr)          '����
         'LabelStr = ""
         oItem.SubItems(1) = Itp(itp_cant_use_on_horseback - 12).Y
         oItem.ForeColor = vbBlue
         oItem.ListSubItems(1).ForeColor = vbBlue
      End If
  '-------------------------------��ҩ-----------------------------------
   ElseIf IsAmmo(Item.itmType) Then
        Set oItem = .ListItems.Add(, "num", "  " & PublicMsgs(129) & ":  ")          '����
           oItem.SubItems(1) = Item.max_ammo
           oItem.ForeColor = &HC0C000
           oItem.ListSubItems(1).ForeColor = &HC0C000
        
      If Item.thrust_damage > 0 Then
        Dam = GetDamage(Item.thrust_damage, tpDam)
        If Dam > 0 Then
           Set oItem = .ListItems.Add(, "thrust", "  " & PublicMsgs(107) & ":  ")          '�����˺�
           oItem.SubItems(1) = Dam & PublicMsgs(108 + tpDam)
           oItem.ForeColor = vbRed
           oItem.ListSubItems(1).ForeColor = vbRed
        End If
      End If
      
      'nFlags = 0
      LabelStr = "  " & PublicMsgs(114) & ":  "
      If IsCanPenetrateShield(Item.itmType) Then
         'nFlags = nFlags + 1
         Set oItem = .ListItems.Add(, "flags" & nFlags, LabelStr)          '����
         'LabelStr = ""
         oItem.SubItems(1) = PublicMsgs(117)
         oItem.ForeColor = vbBlue
         oItem.ListSubItems(1).ForeColor = vbBlue
      End If
  '-------------------------------��ƥ-----------------------------------
   ElseIf IsHorse(Item.itmType) Then
      Set oItem = .ListItems.Add(, "armor", "  " & PublicMsgs(118) & ":  ")          '����
        oItem.SubItems(1) = Item.body_armor
        oItem.ForeColor = &HC0C000
        oItem.ListSubItems(1).ForeColor = &HC0C000
   
      Set oItem = .ListItems.Add(, "speed", "  " & PublicMsgs(111) & ":  ")          '�ٶ�
        oItem.SubItems(1) = Item.missile_speed
   
      Set oItem = .ListItems.Add(, "mv", "  " & PublicMsgs(120) & ":  ")          '����
        oItem.SubItems(1) = Item.speed_rating
   
      Set oItem = .ListItems.Add(, "charge", "  " & PublicMsgs(121) & ":  ")          '���
        oItem.SubItems(1) = Item.thrust_damage
        
      Set oItem = .ListItems.Add(, "hp", "  " & PublicMsgs(122) & ":  ")          'HP
        oItem.SubItems(1) = Item.hit_points
    If Item.difficulty > 0 Then
      Set oItem = .ListItems.Add(, "difficulty", "  " & PublicMsgs(113) & ":  ")          '�Ѷ�
        oItem.SubItems(1) = Item.difficulty
        oItem.ForeColor = vbGreen
        oItem.ListSubItems(1).ForeColor = vbGreen
    End If
  '-------------------------------����-----------------------------------
   ElseIf IsArmor(Item.itmType) Then
      Set oItem = .ListItems.Add(, "ha", "  " & PublicMsgs(119) & ":  ")          'ͷ��
        oItem.SubItems(1) = Item.head_armor
   
      Set oItem = .ListItems.Add(, "charge", "  " & PublicMsgs(123) & ":  ")          '���
        oItem.SubItems(1) = Item.body_armor
        
      Set oItem = .ListItems.Add(, "hp", "  " & PublicMsgs(124) & ":  ")          '�ȷ�
        oItem.SubItems(1) = Item.leg_armor
    
    If Item.difficulty > 0 Then
      Set oItem = .ListItems.Add(, "difficulty", "  " & PublicMsgs(113) & ":  ")          '�Ѷ�
        oItem.SubItems(1) = Item.difficulty
        oItem.ForeColor = vbGreen
        oItem.ListSubItems(1).ForeColor = vbGreen
    End If
  '-------------------------------����-----------------------------------
   ElseIf IsShield(Item.itmType) Then
      Set oItem = .ListItems.Add(, "str", "  " & PublicMsgs(127) & ":  ")          'ǿ��
        oItem.SubItems(1) = Item.hit_points
   
      Set oItem = .ListItems.Add(, "armor", "  " & PublicMsgs(125) & ":  ")          '����
        oItem.SubItems(1) = Item.body_armor
        
      Set oItem = .ListItems.Add(, "size", "  " & PublicMsgs(126) & ":  ")          '�ߴ�
        oItem.SubItems(1) = Item.weapon_length
        
      Set oItem = .ListItems.Add(, "speed", "  " & PublicMsgs(111) & ":  ")          '�ٶ�
        oItem.SubItems(1) = Item.speed_rating
        
    If Item.difficulty > 0 Then
      Set oItem = .ListItems.Add(, "difficulty", "  " & PublicMsgs(113) & ":  ")          '�Ѷ�
        oItem.SubItems(1) = Item.difficulty
        oItem.ForeColor = vbGreen
        oItem.ListSubItems(1).ForeColor = vbGreen
    End If
  '-------------------------------���Ｐ����-----------------------------------
   Else
    If Item.max_ammo > 0 Then
      Set oItem = .ListItems.Add(, "num", "  " & PublicMsgs(129) & ":  ")          '����
         oItem.SubItems(1) = Item.max_ammo
    End If
    
    If Item.head_armor > 0 Then
      Set oItem = .ListItems.Add(, "quality", "  " & PublicMsgs(128) & ":  ")          '����
         oItem.SubItems(1) = Item.head_armor
    End If
         
   End If

End With

AutoSize
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("frmInfo", "LoadItemInfo", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����LoadTroopInfo
'**��    �룺TrpID(-1��Ϊ��ǰ����)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-16 23:12:33
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub LoadTroopInfo(Optional TrpID As Long = -1)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim Trp As Type_Troops, oItem As ListItem
Dim nSkill As Integer, LabelStr As String, i As Integer, Skl As Byte, tC As Long

If TrpID = -1 Then TrpID = CurrentTrpID
Trp = Trps(TrpID)

'ResetFormPos

txtTitle.Text = Trp.csvName

With ListInfo
   .ListItems.Clear

   Set oItem = .ListItems.Add(, "level", "  " & PublicMsgs(131) & ":  ")           '�ȼ�
       oItem.SubItems(1) = Trp.tAttrib.level
       oItem.ForeColor = &HC0&
       oItem.ListSubItems(1).ForeColor = &HC0&
       
   nSkill = 0
   LabelStr = "  " & PublicMsgs(132) & ":  "
   
   For i = 0 To UBound(PublicSkills)
      Skl = GetTroopSkill(CInt(Trp.ID), i + 1)
      If Skl > 0 Then
        nSkill = nSkill + 1
        Set oItem = .ListItems.Add(, "skill" & nSkill, LabelStr)          '����
            oItem.SubItems(1) = PublicSkills(i) & " " & Skl
            If i <= 8 Then
               tC = &HC000&
            ElseIf i > 8 And i <= 11 Then
               tC = vbRed
            ElseIf i > 11 And i <= 22 Then
               tC = &H80FF&
            Else
               tC = vbBlue
            End If
        oItem.ListSubItems(1).ForeColor = tC
        LabelStr = ""
      End If
   Next i

End With

AutoSize
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("frmInfo", "LoadTroopInfo", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����CopyInfotoClipBoard
'**��    �룺
'**��    ������
'**����������������Ϣ�����а�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-14 15:33:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub CopyInfotoClipBoard()
Dim temStr As String, i As Integer

temStr = Replace(txtTitle.Text, " ", "")

For i = 1 To ListInfo.ListItems.Count
  If ListInfo.ListItems(i).Text <> "" Then
    temStr = temStr & vbCrLf & Replace(ListInfo.ListItems(i).Text, " ", "") & Replace(ListInfo.ListItems(i).SubItems(1), " ", "")
  Else
    temStr = temStr & "," & Replace(ListInfo.ListItems(i).SubItems(1), " ", "")
  End If
Next i

Clipboard.Clear
Clipboard.SetText temStr & vbCrLf

End Sub
