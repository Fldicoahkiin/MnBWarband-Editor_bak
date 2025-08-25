Attribute VB_Name = "MdlLanMgr"

'*************************************************************************
'**ģ �� ����MdlLanMgr
'**˵    ����DPS4E.com ��Ȩ����2008 - 2009(C)1
'**�� �� �ˣ�kevin
'**��    �ڣ�2008-05-17 01:11:16
'**�� �� �ˣ�SSgt_Edward,Ser_Charles
'**��    �ڣ�
'**��    ����
'**��    ����V0.951.7
'*************************************************************************
    Option Explicit
    
    Private Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)

    Private LanguageFileName As String
    
    Public PublicTags(36) As String   'ends_add
    Public PublicEditors(13) As String
    Public PublicEditors_Simplified(13) As String
    Public PublicTools(5) As String
    Public PublicHelp(1) As String
    Public PublicTips(5) As String
    Public PublicSkills() As String
    Public PublicMsgs(167) As String
    Public PublicWizards(31) As String
    Public PublicAbout(37) As String
    
    
'*************************************************************************
'**�� �� ����SelectLanguage
'**��    �룺LanName(String) -
'**��    ������
'**����������ѡ������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 01:10:59
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2011-02-05 22:35:01
'**��    ����V1.1321
'*************************************************************************
Public Sub SelectLanguage(LanName As String)
    On Error Resume Next
    
LanguageFileName = ""

  With MnBInfo
    ' check empty
    If Len(LanName) = 0 Then
        GoTo Default
    End If
        
    .Language_Edit = LanName
    
    If .Language_Edit <> "cns" Then

      If FileExists(App.Path & "\" & .Language_Edit & ".lan.ini") Then
         LanguageFileName = App.Path & "\" & .Language_Edit & ".lan.ini"
         'TranslateForms
      Else
         GoTo Default
      End If
    Else
      InitPublicWords
    End If
   End With
   
   TranslatePublicWords
   
   Exit Sub
   
Default:
MnBInfo.Language_Edit = "cns"
InitPublicWords
End Sub

'*************************************************************************
'**�� �� ����GetLanguageFileName
'**��    �룺��
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:58:03
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.13
'*************************************************************************
Public Function GetLanguageFileName(LanName As String) As String

    If Len(LanguageFileName) = 0 Then
        SelectLanguage (LanName)
        GetLanguageFileName = LanguageFileName
    Else
        GetLanguageFileName = LanguageFileName
    End If

End Function

Public Function SetLanguage(LanName As String) As String

        SelectLanguage (LanName)
        SetLanguage = LanguageFileName

End Function

'*************************************************************************
'**�� �� ����TranslateStr
'**��    �룺sSection(String) -
'**        ��sKey(String)     -
'**        ��sDefVal(String)  -
'**��    ����(String) -
'**������������������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 01:10:40
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-02-10 18:13:43
'**��    ����V1.1321
'*************************************************************************
Public Function TranslateStr(sSection As String, sKey As String, sDefVal As String) As String
    Dim sValue As String * 256
    Dim n As Long

    On Error Resume Next

    TranslateStr = sDefVal
    If Not FileExists(LanguageFileName) Then
        Exit Function
    End If
    n = GetPrivateProfileString(sSection, sKey, sDefVal, sValue, 255, LanguageFileName)
    n = InStr(1, sValue, Chr(0), vbTextCompare)
    
    If n > 0 Then
        TranslateStr = Left$(sValue, n - 1)
    End If
End Function

'*************************************************************************
'**�� �� ����TranslateForm
'**��    �룺Frm(Form) -
'**��    ������
'**�����������Զ�ת��Form�ϵ�һЩ�����ؼ���Ҳ��������չ��Ŀǰ֧��CommandButton,Label,OptionButton,CheckButton��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 01:09:23
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2011-02-05 22:49:12
'**��    ����1.1321
'*************************************************************************
Public Sub TranslateForm(frm As Form)
    Dim i As Long, j As Integer

    On Error Resume Next
If LanguageFileName <> "" Then
If FileExists(LanguageFileName) Then
    frm.Caption = TranslateStr(frm.Name, "Caption", frm.Caption)
    
    For i = 0 To frm.Controls.Count - 1
        If (TypeOf frm.Controls(i) Is CommandButton) Or (TypeOf frm.Controls(i) Is Label) Or (TypeOf frm.Controls(i) Is Menu) _
           Or (TypeOf frm.Controls(i) Is OptionButton) Or (TypeOf frm.Controls(i) Is CheckBox) Or (TypeOf frm.Controls(i) Is Frame) Then
            'frm.Controls(i).Caption = TranslateStr(frm.Name, frm.Controls(i).Name, frm.Controls(i).Caption)
            If checkIndex(frm, i) >= 0 Then
                frm.Controls(i).Caption = TranslateStr(frm.Name & "_" & frm.Controls(i).Name, checkIndex(frm, i), frm.Controls(i).Caption)
            Else
                frm.Controls(i).Caption = TranslateStr(frm.Name, frm.Controls(i).Name, frm.Controls(i).Caption)
            End If
        ElseIf (TypeOf frm.Controls(i) Is TabStrip) Then
          For j = 1 To frm.Controls(i).Tabs.Count
            If checkIndex(frm, i) >= 0 Then
                frm.Controls(i).Tabs(j).Caption = TranslateStr(frm.Name & "_" & frm.Controls(i).Name & "_" & checkIndex(frm, i), CStr(j), frm.Controls(i).Tabs(j).Caption)
            Else
                frm.Controls(i).Tabs(j).Caption = TranslateStr(frm.Name & "_" & frm.Controls(i).Name, CStr(j), frm.Controls(i).Tabs(j).Caption)
            End If
          Next j
        End If
    Next i
    
    If LCase(frm.Name) = "frmitems" Then
        For i = 0 To frm.LstFlags.ListCount - 1
            frm.LstFlags.List(i) = TranslateStr(frm.Name & "_LstFlags", CStr(i), frm.LstFlags.List(i))
        Next i
        
        For i = 0 To frm.LstAction.ListCount - 1
            frm.LstAction.List(i) = TranslateStr(frm.Name & "_LstAction", CStr(i), frm.LstAction.List(i))
        Next i
        
        For i = 0 To frm.CBixmesh.ListCount - 1
            frm.CBixmesh.List(i) = TranslateStr(frm.Name & "_CBixmesh", CStr(i), frm.CBixmesh.List(i))
        Next i
    ElseIf LCase(frm.Name) = "frmpsys" Then
    
        For i = 0 To frm.LstFlags.ListCount - 1
            frm.LstFlags.List(i) = TranslateStr(frm.Name & "_LstFlags", CStr(i), frm.LstFlags.List(i))
        Next i
    ElseIf LCase(frm.Name) = "frmmap" Then
    
        For i = 0 To frm.LstFilter(0).ListCount - 1
            frm.LstFilter(0).List(i) = TranslateStr(frm.Name & "_LstFilter", CStr(i), frm.LstFilter(0).List(i))
        Next i
    End If
End If
End If
End Sub



'*************************************************************************
'**�� �� ����InitTranslateForm
'**��    �룺Frm(Form)           -
'**        ��iniFileName(String) -
'**��    ������
'**������������Form�ϵ�һЩ�����ؼ���Caption��¼��ini�ļ�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 01:09:11
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2011-02-05 22:49:12
'**��    ����1.1321
'*************************************************************************
Public Sub InitTranslateForm(frm As Form, iniFileName As String)
    Dim i As Long, j As Integer
    On Error Resume Next

    'Frm.Caption = TranslateStr(Frm.Name, "Caption", Frm.Caption)
    Call WriteString(iniFileName, frm.Name, "Caption", frm.Caption)
    
    For i = 0 To frm.Controls.Count - 1
        If (TypeOf frm.Controls(i) Is CommandButton) Or (TypeOf frm.Controls(i) Is Label) Or (TypeOf frm.Controls(i) Is Menu) _
           Or (TypeOf frm.Controls(i) Is OptionButton) Or (TypeOf frm.Controls(i) Is CheckBox) Or (TypeOf frm.Controls(i) Is Frame) Then
            'Frm.Controls(I).Caption = TranslateStr(Frm.Name, Frm.Controls(I).Name, Frm.Controls(I).Caption)
            If checkIndex(frm, i) >= 0 Then
                Call WriteString(iniFileName, frm.Name & "_" & frm.Controls(i).Name, checkIndex(frm, i), frm.Controls(i).Caption)
            Else
                Call WriteString(iniFileName, frm.Name, frm.Controls(i).Name, frm.Controls(i).Caption)
            End If
        ElseIf (TypeOf frm.Controls(i) Is TabStrip) Then
            For j = 1 To frm.Controls(i).Tabs.Count
               If checkIndex(frm, i) >= 0 Then
                Call WriteString(iniFileName, frm.Name & "_" & frm.Controls(i).Name & "_" & checkIndex(frm, i), CStr(j), frm.Controls(i).Tabs(j).Caption)
               Else
                Call WriteString(iniFileName, frm.Name & "_" & frm.Controls(i).Name, CStr(j), frm.Controls(i).Tabs(j).Caption)
               End If
            Next j
        End If
    Next i
    
    If LCase(frm.Name) = "frmitems" Then
        For i = 0 To frm.LstFlags.ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_LstFlags", CStr(i), frm.LstFlags.List(i))
        Next i
        
        For i = 0 To frm.LstAction.ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_LstAction", CStr(i), frm.LstAction.List(i))
        Next i
        
        For i = 0 To frm.CBixmesh.ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_CBixmesh", CStr(i), frm.CBixmesh.List(i))
        Next i
    ElseIf LCase(frm.Name) = "frmpsys" Then
    
        For i = 0 To frm.LstFlags.ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_LstFlags", CStr(i), frm.LstFlags.List(i))
        Next i
    ElseIf LCase(frm.Name) = "frmmap" Then
        For i = 0 To frm.LstFilter(0).ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_LstFilter", CStr(i), frm.LstFilter(0).List(i))
        Next i
        
    End If
End Sub

'*************************************************************************
'**�� �� ����TranslatePublicWords
'**��    �룺iniFileName(String)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2011-02-08 15:06:35
'**��    ����1.1321
'*************************************************************************
Public Sub TranslatePublicWords()
Dim i As Integer, j As Integer

If Not FileExists(LanguageFileName) Then
   Exit Sub
End If
 
For i = 0 To UBound(PublicTags)
    PublicTags(i) = TranslateStr("PublicTags", CStr(i), PublicTags(i))
Next i

For i = 0 To UBound(PublicEditors)
    PublicEditors(i) = TranslateStr("PublicEditors", CStr(i), PublicEditors(i))
Next i

For i = 0 To UBound(PublicEditors_Simplified)
    PublicEditors_Simplified(i) = TranslateStr("PublicEditors_Simplified", CStr(i), PublicEditors_Simplified(i))
Next i

For i = 0 To UBound(PublicTools)
    PublicTools(i) = TranslateStr("PublicTools", CStr(i), PublicTools(i))
Next i

For i = 0 To UBound(PublicHelp)
    PublicHelp(i) = TranslateStr("PublicHelp", CStr(i), PublicHelp(i))
Next i

For i = 0 To UBound(PublicHelp)
    PublicHelp(i) = TranslateStr("PublicHelp", CStr(i), PublicHelp(i))
Next i

For i = 0 To UBound(PublicTips)
    PublicTips(i) = TranslateStr("PublicTips", CStr(i), PublicTips(i))
Next i

For i = 0 To UBound(PublicMsgs)
    PublicMsgs(i) = TranslateStr("PublicMsgs", CStr(i), PublicMsgs(i))
Next i

For i = 0 To UBound(PublicSkills)
    PublicSkills(i) = TranslateStr("PublicSkills", CStr(i), PublicSkills(i))
Next i

For i = 0 To UBound(Operation)
   With Operation(i)
     .Op_CSVname = TranslateStr("Operation" & i, "Name", .Op_CSVname)
    
     For j = 1 To .ParaNum
         .Para(j).Value = TranslateStr("Operation" & i, CStr(j), .Para(j).Value)
     Next j
   End With
Next i

   For i = 0 To UBound(tiOn)
     tiOn(i).Y = TranslateStr("tiOn", CStr(i), tiOn(i).Y)
   Next i

For i = 0 To UBound(PublicWizards)
    PublicWizards(i) = TranslateStr("PublicWizards", CStr(i), PublicWizards(i))
Next i

For i = 0 To UBound(PublicAbout)
    PublicAbout(i) = TranslateStr("PublicAbout", CStr(i), PublicAbout(i))
Next i

End Sub

'*************************************************************************
'**�� �� ����WritePublicWords
'**��    �룺
'**��    ������
'**������������PublicWords��¼��ini�ļ�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-02-08 15:06:35
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����1.1321
'*************************************************************************
Public Sub WritePublicWords()
Dim i As Integer, iniFileName As String, j As Integer

iniFileName = App.Path & "\new.lan.ini"
 
For i = 0 To UBound(PublicTags)
    Call WriteString(iniFileName, "PublicTags", CStr(i), PublicTags(i))
Next i

For i = 0 To UBound(PublicEditors)
    Call WriteString(iniFileName, "PublicEditors", CStr(i), PublicEditors(i))
Next i

For i = 0 To UBound(PublicEditors_Simplified)
    Call WriteString(iniFileName, "PublicEditors_Simplified", CStr(i), PublicEditors_Simplified(i))
Next i

For i = 0 To UBound(PublicTools)
    Call WriteString(iniFileName, "PublicTools", CStr(i), PublicTools(i))
Next i

For i = 0 To UBound(PublicHelp)
    Call WriteString(iniFileName, "PublicHelp", CStr(i), PublicHelp(i))
Next i

For i = 0 To UBound(PublicTips)
    Call WriteString(iniFileName, "PublicTips", CStr(i), PublicTips(i))
Next i

For i = 0 To UBound(PublicMsgs)
    Call WriteString(iniFileName, "PublicMsgs", CStr(i), PublicMsgs(i))
Next i

For i = 0 To UBound(PublicSkills)
    Call WriteString(iniFileName, "PublicSkills", CStr(i), PublicSkills(i))
Next i

For i = 0 To UBound(Operation)
   With Operation(i)
     Call WriteString(iniFileName, "Operation" & i, "Name", .Op_CSVname)
     
     For j = 1 To .ParaNum
         Call WriteString(iniFileName, "Operation" & i, CStr(j), .Para(j).Value)
     Next j
   End With
Next i

For i = 0 To UBound(tiOn)
     Call WriteString(iniFileName, "tiOn", CStr(i), tiOn(i).Y)
Next i

For i = 0 To UBound(PublicWizards)
    Call WriteString(iniFileName, "PublicWizards", CStr(i), PublicWizards(i))
Next i

For i = 0 To UBound(PublicAbout)
    Call WriteString(iniFileName, "PublicAbout", CStr(i), PublicAbout(i))
Next i

End Sub

'*************************************************************************
'**�� �� ����AddSplash
'**��    �룺path(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 01:09:07
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.7
'*************************************************************************
Public Function AddSplash(Path As String)
    On Error Resume Next
    '------------------------------------------------
    'AddSplash(App.path)
    
    AddSplash = App.Path & "\"

End Function


'*************************************************************************
'**�� �� ����FileExists
'**��    �룺file(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 01:08:50
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.7
'*************************************************************************
Public Function FileExists(file As String) As Boolean
    On Error Resume Next
    If Trim(file) = "" Then
        FileExists = False
        Exit Function
    End If
    If Dir(file, vbNormal + vbReadOnly + vbHidden + vbSystem + vbArchive) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

'*************************************************************************
'**�� �� ����DirExists
'**��    �룺file(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-27 23:17:43
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function DirExists(Path As String) As Boolean
Dim s As String
    On Error Resume Next
    DirExists = True
    If Trim(Path) = "" Then
        DirExists = False
        Exit Function
    End If
    
    s = Dir(Path, vbNormal + vbHidden + vbSystem + vbDirectory)
    s = Trim(s)
    If s = "" Then
       DirExists = False
       Exit Function
    End If
    
    If Right(s, 1) = "." Then
       DirExists = False
       Exit Function
    End If
    
End Function


'*************************************************************************
'**�� �� ����checkIndex
'**��    �룺frm(Form) -
'**        ��i(Long)   -
'**��    ����(Integer) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 01:39:57
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.9
'*************************************************************************
Private Function checkIndex(frm As Form, i As Long) As Integer
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    checkIndex = CInt(frm.Controls(i).Index)
    

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    checkIndex = -1
End Function



'*************************************************************************
'**�� �� ����showLanMsg
'**��    �룺msgIndex(Integer) -
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 08:06:23
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.9
'*************************************************************************
Public Function showLanMsg(msgIndex As Integer) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    'showLanMsg = Form1.msgConf(msgIndex)
    'showLanMsg = gMsgConf(msgIndex)
    
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    showLanMsg = "#ERROR# NO_MESSAGE! [" & msgIndex & "]"
    Call logErr("MdlLanMgr", "showLanMsg", Err.Number, Err.Description)
End Function

Public Function InitTranslateForms()

'InitTranslateForm frmAbout, App.Path & "\new.lan.ini"
InitTranslateForm frmBackUpManager, App.Path & "\new.lan.ini"
InitTranslateForm frmFactions, App.Path & "\new.lan.ini"
InitTranslateForm frmItems, App.Path & "\new.lan.ini"
InitTranslateForm frmLine, App.Path & "\new.lan.ini"
InitTranslateForm frmMain, App.Path & "\new.lan.ini"
InitTranslateForm frmMap, App.Path & "\new.lan.ini"
InitTranslateForm frmMap_Icons, App.Path & "\new.lan.ini"
InitTranslateForm frmMesh, App.Path & "\new.lan.ini"
InitTranslateForm frmParties, App.Path & "\new.lan.ini"
InitTranslateForm frmParty_Templates, App.Path & "\new.lan.ini"
InitTranslateForm frmPSys, App.Path & "\new.lan.ini"
InitTranslateForm frmSaveAs, App.Path & "\new.lan.ini"
InitTranslateForm frmScenes, App.Path & "\new.lan.ini"
InitTranslateForm frmSelectPath, App.Path & "\new.lan.ini"
InitTranslateForm frmSoundRess, App.Path & "\new.lan.ini"
InitTranslateForm frmSounds, App.Path & "\new.lan.ini"
InitTranslateForm frmTabMat, App.Path & "\new.lan.ini"
InitTranslateForm frmTip, App.Path & "\new.lan.ini"
InitTranslateForm frmTroops, App.Path & "\new.lan.ini"
InitTranslateForm frmTrigger, App.Path & "\new.lan.ini"
InitTranslateForm Welcome, App.Path & "\new.lan.ini"

End Function


'*************************************************************************
'**�� �� ����InitPublicWords
'**��    �룺��
'**��    ������
'**������������ʼ��PublicTags
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-02-07 00:22:45
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub InitPublicWords()
Dim n As Integer, Discrabe() As Variant

'ends-add
PublicTags(0) = "����"
PublicTags(Tag_Register) = "�Ĵ���"
PublicTags(Tag_Variable) = "ȫ�ֱ���"
PublicTags(Tag_String) = "�ַ���"
PublicTags(Tag_Item) = "��Ʒ"
PublicTags(Tag_Troop) = "����"
PublicTags(Tag_Faction) = "��Ӫ"
PublicTags(Tag_Quest) = "����"
PublicTags(Tag_Party_Tpl) = "����ģ��"
PublicTags(Tag_Party) = "����"
PublicTags(Tag_Scene) = "����"
PublicTags(Tag_Mission_tpl) = "����ģ��"
PublicTags(Tag_Menu) = "�˵�"
PublicTags(Tag_Script) = "�ű�"
PublicTags(Tag_Particle_Sys) = "����ϵͳ"
PublicTags(Tag_Scene_Prop) = "��������"
PublicTags(Tag_Sound) = "����"
PublicTags(Tag_Local_Variable) = "�ֲ�����"
PublicTags(Tag_Map_Icon) = "���ͼͼ��"
PublicTags(Tag_Skill) = "����"
PublicTags(Tag_Mesh) = "����ģ��"
PublicTags(Tag_Presentation) = "չʾ"
PublicTags(Tag_Quick_String) = "�����ַ���"
PublicTags(Tag_Track) = "��Ŀ"
PublicTags(Tag_Tableau) = "�ɱ����"
PublicTags(Tag_Animation) = "����"
PublicTags(Tags_End) = "����"
PublicTags(Tags_End + 1) = "λ��"
PublicTags(Tags_End + 2) = "��Ʒ����"
PublicTags(Tags_End + 3) = "���ֱ�ǩ"
PublicTags(Tags_End + 4) = "����"
PublicTags(Tags_End + 5) = "����AI��Ϊ"
PublicTags(Tags_End + 6) = "����ѡ��"
PublicTags(Tags_End + 7) = "���ӱ�ǩ"
PublicTags(Tags_End + 8) = "����Ȩ��"
PublicTags(Tags_End + 9) = "����ֵ�趨"
PublicTags(Tags_End + 10) = "s"

PublicEditors(0) = "�༭��"
PublicEditors(1) = "����"
PublicEditors(2) = "��Ʒ"
PublicEditors(3) = "����"
PublicEditors(4) = "����ģ��"
PublicEditors(5) = "��Ӫ"
PublicEditors(6) = "����"
PublicEditors(7) = "���ͼͼ��"
PublicEditors(8) = "����ϵͳ"
PublicEditors(9) = "����"
PublicEditors(10) = "������Դ"
PublicEditors(11) = "�ɱ��ز�"
PublicEditors(12) = "����ģ��"
PublicEditors(13) = "������"

PublicEditors_Simplified(0) = "�༭��"
PublicEditors_Simplified(1) = "����"
PublicEditors_Simplified(2) = "��Ʒ"
PublicEditors_Simplified(3) = "����"
PublicEditors_Simplified(4) = "ģ��"
PublicEditors_Simplified(5) = "��Ӫ"
PublicEditors_Simplified(6) = "����"
PublicEditors_Simplified(7) = "ͼ��"
PublicEditors_Simplified(8) = "����ϵͳ"
PublicEditors_Simplified(9) = "����"
PublicEditors_Simplified(10) = "��Դ"
PublicEditors_Simplified(11) = "�ɱ��ز�"
PublicEditors_Simplified(12) = "����ģ��"
PublicEditors_Simplified(13) = "������"
        
PublicTools(0) = "����"
PublicTools(1) = "�������ǵ�ͼ"
PublicTools(2) = "���ݹ�����"
PublicTools(3) = "��Ʒ������"
PublicTools(4) = "�������д����"
PublicTools(5) = "�ַ���������"
        
PublicHelp(0) = "����"
PublicHelp(1) = "����"

PublicTips(0) = "(��)"
PublicTips(1) = "����[str0]��,���Ժ�..."
PublicTips(2) = "����ģ����,���Ժ�..."
PublicTips(3) = "������Դ��,���Ժ�..."
PublicTips(4) = "������,���Ժ�..."
PublicTips(5) = "������,���Ժ�..."

PublicMsgs(0) = "��ʾ"
PublicMsgs(1) = "ȷ�����õ�ǰ����?"
PublicMsgs(2) = "�Ƿ�ɾ��[[str0]]?"
PublicMsgs(3) = "ɾ��[str0]"
PublicMsgs(4) = "[[str0]]�Ǳ�Ҫ�ɷ�,�޷�ɾ��!"
PublicMsgs(5) = "�Ƿ����[[str0]]��������[str1]?"
PublicMsgs(6) = "��ϵ"
PublicMsgs(7) = "[[str0]]�Ǳ�Ҫ�ɷ�,�޷��ƶ�!"
PublicMsgs(8) = "�ƶ�[str0]"
PublicMsgs(9) = "������[str0]"
PublicMsgs(10) = "ȷ�����õ�ǰ����?"
PublicMsgs(11) = "û�ҵ�����ƥ����!"
PublicMsgs(12) = "��ѯ"
PublicMsgs(13) = "����"
PublicMsgs(14) = "��"
PublicMsgs(15) = "ȷ��������������?"
PublicMsgs(16) = "���������ѳɹ�������new.lan.ini��!���ȷ�������������"
PublicMsgs(17) = "ȷ��Ҫ����һ����ԭ����?"
PublicMsgs(18) = "[str0]���ֻ����[str1]!"
PublicMsgs(19) = "����!"
PublicMsgs(20) = "������ԭ��ʧ��!��ȷ�����ݾ籾��������!"
PublicMsgs(21) = "�����ɹ�!�����ļ�����[[str0]]"
PublicMsgs(22) = "ȷ���˳� �������뿳ɱ:ս�š��籾�༭�� ��?"
PublicMsgs(23) = "ȷ���˳�"
PublicMsgs(24) = "ȷ���������趨����籾[[str0]]?"
PublicMsgs(25) = "����籾"
PublicMsgs(26) = "����"
PublicMsgs(27) = "��ȷ�ϱ��ݾ籾[[str0]]��,����ȷ������ʼ���档"
PublicMsgs(28) = "��һ�����ĵڶ��ֹ���ģʽ"
PublicMsgs(29) = "�籾[[str0]]����ɹ�!"
PublicMsgs(30) = "�������:[[str0]]"
PublicMsgs(31) = "����"
PublicMsgs(32) = "ѡ��籾"
PublicMsgs(33) = "ȷ��Ҫ����ѡ��籾?"
PublicMsgs(34) = "ɾ����ԭ��"
PublicMsgs(35) = "ȷ��Ҫɾ����ѡ��ԭ����?"
PublicMsgs(36) = "����һ����ԭ��"
PublicMsgs(37) = "ȷ��Ҫ����һ����ʱ���ע�Ļ�ԭ����?"
PublicMsgs(38) = "����"
PublicMsgs(39) = "��������������!"
PublicMsgs(40) = "��ԭ"
PublicMsgs(41) = "ȷ��Ҫ��ԭ[[str0]]��?"
PublicMsgs(42) = "����"
PublicMsgs(43) = "��ԭ�ɹ�!"
PublicMsgs(44) = "��ԭʧ��!"
PublicMsgs(45) = "��������ӵ����Ʒ����(64)!���ʧ��!"
PublicMsgs(46) = "����"
PublicMsgs(47) = "��²"
PublicMsgs(48) = "����"
PublicMsgs(49) = "���"
PublicMsgs(50) = "�籾:[str0]"
PublicMsgs(51) = "���Ϊ"
PublicMsgs(52) = "����������"
PublicMsgs(53) = "������ľ籾�����Ϸ�!"
PublicMsgs(54) = "�޷�ʶ��"
PublicMsgs(55) = "����"
PublicMsgs(56) = "���ܻӿ�"
PublicMsgs(57) = "���ܴ���"
PublicMsgs(58) = "���ȸ��Ʋ���!"
PublicMsgs(59) = "��ѡ��Ҫճ���Ĳ������ڵĴ�����!"
PublicMsgs(60) = "����Ҫ��һ��ģ��!"
PublicMsgs(61) = "����Ʒ���ܱ༭"
PublicMsgs(62) = "��ѡ��Ҫ�����������ڵĲ���!"
PublicMsgs(63) = "������������½�����!"
PublicMsgs(64) = "��Ʒ������Ӫ������,����Ϊ����Ӫ"
PublicMsgs(65) = "������"
PublicMsgs(66) = "����"
PublicMsgs(67) = "ִ��"
PublicMsgs(68) = "����"
PublicMsgs(69) = "[str0][str1][str2][str3]��[str4]������,������Ϊ[str4](0)!������ð�ť�����ø�����"
PublicMsgs(70) = "����ֻ��ѡ��!"
PublicMsgs(71) = "�����ڸ�����!"
PublicMsgs(72) = "��ѡ��Ҫ��ӵĲ������ڵĴ�����!"
PublicMsgs(73) = "�������Ϸ�!"
PublicMsgs(74) = "������ƶ�!"
PublicMsgs(75) = "��ѡ��Ҫ�ƶ��Ĳ���!"
PublicMsgs(76) = "��Ʒ�������Ѹ���"
PublicMsgs(77) = "�����Ѹ���"
PublicMsgs(78) = "����ܸ���!"
PublicMsgs(79) = "��ѡ��Ҫ���ƵĴ����������!"
PublicMsgs(80) = "ȷ��Ҫɾ���ô�����ô?"
PublicMsgs(81) = "ȷ��Ҫɾ���ò���ô?"
PublicMsgs(82) = "����Ҫ��һ������!"
PublicMsgs(83) = "��ѡ��������ɾ��!"
PublicMsgs(84) = "��ѡ����ֻ�ܴ����һ��ɾ��!"
PublicMsgs(85) = "�����ɾ��!"
PublicMsgs(86) = "δѡ���κ���Ŀ!"
PublicMsgs(87) = "���ȸ�����Ʒ������!"
PublicMsgs(88) = "ֵ"
PublicMsgs(89) = "����������"
PublicMsgs(90) = "[[str0]]�Ѵ���,����ʧ��!"
PublicMsgs(91) = "[[str0]]�Ѵ���,�����ID�������á�"
PublicMsgs(92) = "�������"
PublicMsgs(93) = "������������[str0]��"
PublicMsgs(94) = "���н������[str0]��"
PublicMsgs(95) = "��ѡ��Ҫ���ƵĲ���!"
PublicMsgs(96) = "ȷ��Ҫ��ע������뵱ǰ��Ϣ��?"
PublicMsgs(97) = "ȷ��ע��"
PublicMsgs(98) = "�ѳɹ���ע������뵱ǰ��Ϣ!"
PublicMsgs(99) = "�Ƿ񰴵�ǰ�ı�����?"
PublicMsgs(100) = "����"
PublicMsgs(101) = "�������!"
PublicMsgs(102) = "�� ��"
PublicMsgs(103) = "�� ��"
PublicMsgs(104) = "��/˫������"
PublicMsgs(105) = "�� ��"
PublicMsgs(106) = "�� ��"
PublicMsgs(107) = "�� ��"
PublicMsgs(108) = "��"
PublicMsgs(109) = "��"
PublicMsgs(110) = "��"
PublicMsgs(111) = "�� ��"
PublicMsgs(112) = "�� Χ"
PublicMsgs(113) = "�� ��"
PublicMsgs(114) = "�� ��"
PublicMsgs(115) = "�� ��"
PublicMsgs(116) = "�� ��"
PublicMsgs(117) = "���Դ���"
PublicMsgs(118) = "�� ��"
PublicMsgs(119) = "ͷ ��"
PublicMsgs(120) = "�� ��"
PublicMsgs(121) = "�� ��"
PublicMsgs(122) = "�� ��"
PublicMsgs(123) = "�� ��"
PublicMsgs(124) = "�� ��"
PublicMsgs(125) = "�� ��"
PublicMsgs(126) = "�� ��"
PublicMsgs(127) = "ǿ ��"
PublicMsgs(128) = "�� ��"
PublicMsgs(129) = "�� ��"
PublicMsgs(130) = "��Ϣ�Ѹ��Ƶ����а�"
PublicMsgs(131) = "�� ��"
PublicMsgs(132) = "�� ��"
PublicMsgs(133) = "����"
PublicMsgs(134) = "�ѽ��޸���������"
PublicMsgs(135) = "�� ��"        '����
PublicMsgs(136) = "�� ��"        '����
PublicMsgs(137) = "�� ��"        '����
PublicMsgs(138) = "ǿ ��"        '����
PublicMsgs(139) = "�� ��"        '����
PublicMsgs(140) = "�� ��"        '����
PublicMsgs(141) = "��Ʒ��Ϣ�Ѹ��Ƶ����а�"        '����
PublicMsgs(142) = "ȷ������ע�����?"
PublicMsgs(143) = "�����ѳɹ�������new.op.ini��!"
PublicMsgs(144) = "����������ʹ����������Ч!"
PublicMsgs(145) = "ȷ��Ҫ����������[str0]?"
PublicMsgs(146) = "�����ѳɹ�����Ϊ[str0]!"
PublicMsgs(147) = "ǰ׺"
PublicMsgs(148) = "����������..."
PublicMsgs(149) = "����Ĳ�����"
PublicMsgs(150) = "���ַ���"
PublicMsgs(151) = "����Դ������"
PublicMsgs(152) = "����[str0]:[str1]" & vbCrLf & "��[str2]�������ĵ�[str3]��������"
PublicMsgs(153) = "����"
PublicMsgs(154) = "�ǻ�����"
PublicMsgs(155) = "������:[[str0]]"
PublicMsgs(156) = "<����[str0]>"
PublicMsgs(157) = "��"
PublicMsgs(158) = "<�߼���>"
PublicMsgs(159) = "<�������...>"
PublicMsgs(160) = "����"   '����
PublicMsgs(161) = "���ǻ�����"
PublicMsgs(162) = "�޷��ҵ�����[[str0]]!"
PublicMsgs(163) = "����"
PublicMsgs(164) = "λ��"   '����
PublicMsgs(165) = "��Ʒ����"   '����
PublicMsgs(166) = "����"
PublicMsgs(167) = "�������½��������ǽ���ǰ���������أ�ѡ��[��]�������±���,ѡ��[��]����ǰ����������"

Discrabe = Array("����", "ͳ��", "��²����", "", "", _
               "", "", "˵����", "����ѧ", "����", _
               "����", "����", "��Ʒ����", "���", "��", _
               "ս��", "����", "����", "", "", _
               "", "", "�Ӷ�", "����", "����", _
               "�ܶ�", "�ܷ�", "��������", "", "", _
               "", "", "", "ǿ��", "ǿ��", _
               "ǿ��", "����", "", "", "", _
               "", "")
               
ReDim PublicSkills(UBound(Discrabe))

For n = 0 To UBound(Discrabe)
      PublicSkills(n) = Discrabe(n)
Next n

PublicWizards(0) = "��һ��>(&N)"
PublicWizards(1) = "<��һ��(&F)"
PublicWizards(2) = "ȡ��(&C)"
PublicWizards(3) = "���(&F)"
PublicWizards(4) = "����ɹ�"
PublicWizards(5) = "����ʧ��"
PublicWizards(6) = "���󱨸�"
PublicWizards(7) = "������Ʒtxt��Ϣ(item_kinds1.txt)"
PublicWizards(8) = "ѡ���赼��װ�����ļ���"
PublicWizards(9) = "������Ʒ������Ϣ(item_kinds.csv)"
PublicWizards(10) = "ȷ��������Ʒ"
PublicWizards(11) = "������Ʒ����(��[str0]��):"
PublicWizards(12) = "����ģ�ͱ���(��[str0]��):"
PublicWizards(13) = "������ͼ����(��[str0]��):"
PublicWizards(14) = "[δ�ҵ�]"
PublicWizards(15) = "��ͼ�ļ�"
PublicWizards(16) = "ģ���ļ�"
PublicWizards(17) = "���κ�txt����!�����򵼽����������Ӧ��Ʒ,��Ҫ���Լ���[��Ʒ�༭��]���ֶ���ӡ��Ƿ������һ��?"
PublicWizards(18) = "���κ�������Ϣ����!�Ƿ������һ��?"
PublicWizards(19) = "�뽫�赼�����Ʒ��txt�и��ƽ�������ı���,���û��,��ֱ�ӵ��[��һ��]��"

PublicWizards(20) = "��ѡ����Ҫ�����װ�����ļ���,"
PublicWizards(21) = "ȷ�����ļ�������Textures(��ͼ)��Resource(ģ��)�ļ���,"
PublicWizards(22) = "��ȷ�����зֱ���DDS��BRF�ļ���"
PublicWizards(23) = "��������������ļ�����������ŵ�,"
PublicWizards(24) = "Ҳ�������˴����ȷ��������������С�"
                                    
PublicWizards(25) = "�뽫�赼�����Ʒ��������Ϣ(�����·��ı���������)���ƽ�������ı���,"
PublicWizards(26) = "���û��,��ֱ�ӵ�[��һ��]��"

PublicWizards(27) = "�������Ϣ������,�鿴�������ñ���,�����󣬵��[��һ��]�����޸ģ�ȷ����ȷ,���[��һ��]�����������"

PublicWizards(28) = "��Ʒ�ѳɹ�����,"
PublicWizards(29) = "��������[��Ʒ�༭��]����ʱ�޸����ǡ���ǵ��ڹرձ༭��ǰ[����籾]��ʹ������Ч��"
PublicWizards(30) = "��������û��txt��,���޷�������Ӧ��Ʒ����������[��Ʒ�༭��]�ﴴ������Ʒ,���Ѹ���Ʒ��ģ����Ϊ�յ����ģ�͡�"

PublicWizards(31) = "��������ж�!��ȷ��ģ���ļ�����ȷ��,txt�ı�����ȷ��,�Լ�������Ϣ����ȷ��!���[��һ��]�����޸ġ�"


PublicAbout(0) = "������Ա:"
PublicAbout(1) = "������:SSgt_Edward,Ser_Charles"
PublicAbout(2) = "����:�Ҳ��Ǹ�����(hp_honey)"

PublicAbout(3) = "�����༭�������߸�����:"
PublicAbout(24) = "����:SSgt_Edward"
PublicAbout(25) = "��Ʒ:Ser_Charles"
PublicAbout(26) = "����:SSgt_Edward"
PublicAbout(27) = "����ģ��:SSgt_Edward"
PublicAbout(28) = "��Ӫ:SSgt_Edward"
PublicAbout(29) = "����:SSgt_Edward"
PublicAbout(30) = "���ͼͼ��:SSgt_Edward"
PublicAbout(31) = "�������ǵ�ͼ:SSgt_Edward"
PublicAbout(32) = "����ϵͳ:Ser_Charles"
PublicAbout(33) = "�ɱ��ز�:Ser_Charles"
PublicAbout(34) = "����:SSgt_Edward"
PublicAbout(35) = "������Դ:SSgt_Edward"
PublicAbout(36) = "���ݹ�����:SSgt_Edward"
PublicAbout(37) = "��Ʒ������:SSgt_Edward"

PublicAbout(4) = "MnBWarband Editor����:"
PublicAbout(5) = "[MnBWarband Editor]�༭����һ�����ڲ鿴���޸�[�����뿳ɱ:ս��]�����Ϸ����ѹ��������"
PublicAbout(6) = "�������￳����վ��[����ɱTXT�༭��](kevin,kiss2003,mafei82394) ������ֲ�����ƶ���(������7��ģ��),"
PublicAbout(7) = "��������[��Ӫ]��[����]��[����]��[���ͼͼ��]��[����ϵͳ]��[����]��[������Դ]��[�ɱ��ز�]�༭���Լ�"
PublicAbout(8) = "[�������ǵ�ͼ]��[���ݹ�����]��[��Ʒ������]���ߵ�ͬʱ,"
PublicAbout(9) = "������ԭ�е�[����]��[��Ʒ]��[����ģ��]�༭��,ʹ�û������ɵ��޸��Լ���ģ�顣"

PublicAbout(10) = "ע��:"
PublicAbout(11) = "MnBWarband Editor��ս��1.132�����±�д���������ԡ�"

PublicAbout(12) = "����:"
PublicAbout(13) = "��MnBWarband Editor�汾�ȶ�֮��,�Ὺ��Դ����(VB)��"

PublicAbout(14) = "�ر��л:"
PublicAbout(15) = "��л�￳����վ��kevin,kiss2003,mafei82394������[����ɱTXT�༭��]��Դ����"
PublicAbout(16) = "��л�￳����վ[�����뿳ɱMOD�����̳̣����İ�]�ĺ�����Ϊ���Ǵ�����MOD�̳�"
PublicAbout(17) = "��л�ٶ������뿳ɱ�ɺ��￳����վ�������Ƕ�MnBWarband Editor�Ĺ�ע��֧��!"

PublicAbout(17) = "��ϵ����:"
PublicAbout(18) = "E-mail:ce1992620@yahoo.com.cn"
PublicAbout(19) = "       ce19920620@gmail.com"
PublicAbout(20) = "QQ:1061675615"
PublicAbout(21) = "   825312405"
PublicAbout(22) = "ħ������Ⱥ:74195097"

PublicAbout(23) = "��ӭ��ʱ����������ʹ�����⡢Bug��������,������������MnBWarband Editor!"
End Sub
