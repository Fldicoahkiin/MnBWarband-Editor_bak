VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParties 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "���ӱ༭��"
   ClientHeight    =   9375
   ClientLeft      =   4200
   ClientTop       =   1050
   ClientWidth     =   14355
   ForeColor       =   &H00C00000&
   Icon            =   "frmParties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_3"
   Begin VB.Frame FraProps 
      Caption         =   "FraProps3"
      Height          =   7935
      Index           =   3
      Left            =   6240
      TabIndex        =   6
      Top             =   480
      Width           =   7815
      Begin VB.Frame Frame1 
         Caption         =   "��Ա"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   7695
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   7575
         Begin VB.CommandButton CQuery2 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   7080
            TabIndex        =   98
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton CQuery2 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   6720
            TabIndex        =   97
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtQuery2 
            Height          =   330
            Left            =   2160
            TabIndex        =   96
            Top             =   240
            Width           =   4575
         End
         Begin MSComctlLib.ListView LstTroops 
            Height          =   3135
            Left            =   120
            TabIndex        =   55
            Top             =   3960
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5530
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.CheckBox chkSlave 
            Caption         =   "��²"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3120
            TabIndex        =   54
            Top             =   7250
            Width           =   855
         End
         Begin VB.TextBox txtMin 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   705
            TabIndex        =   52
            Top             =   7170
            Width           =   2295
         End
         Begin MSComctlLib.ListView boxTroops 
            Height          =   3015
            Left            =   120
            TabIndex        =   94
            Top             =   600
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5318
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label bCMD 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&D)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   3
            Left            =   5880
            MouseIcon       =   "frmParties.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   104
            Top             =   7200
            Width           =   705
         End
         Begin VB.Label bCMD 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&U)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   2
            Left            =   5040
            MouseIcon       =   "frmParties.frx":0614
            MousePointer    =   99  'Custom
            TabIndex        =   103
            Top             =   7200
            Width           =   705
         End
         Begin VB.Label bCMD 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ɾ��(&E)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   0
            Left            =   6720
            MouseIcon       =   "frmParties.frx":091E
            MousePointer    =   99  'Custom
            TabIndex        =   102
            Top             =   7200
            Width           =   705
         End
         Begin VB.Label bCMD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������ѡ����(&A)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   1
            Left            =   5715
            MouseIcon       =   "frmParties.frx":0C28
            MousePointer    =   99  'Custom
            TabIndex        =   101
            Top             =   3720
            Width           =   1680
         End
         Begin VB.Label bCMD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ȫ��ѡ(&C)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   4
            Left            =   4680
            MouseIcon       =   "frmParties.frx":0F32
            MousePointer    =   99  'Custom
            TabIndex        =   100
            Top             =   3720
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   1680
            TabIndex        =   99
            Top             =   315
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ӵ�б���:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   25
            Left            =   120
            TabIndex        =   95
            Top             =   3720
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɼ������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   23
            Left            =   120
            TabIndex        =   93
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   53
            Top             =   7275
            Width           =   495
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps1"
      Height          =   4695
      Index           =   1
      Left            =   7560
      TabIndex        =   5
      Top             =   2400
      Width           =   5415
      Begin VB.Frame Frame1 
         Caption         =   "AI"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2175
         Index           =   3
         Left            =   240
         TabIndex        =   70
         Top             =   2760
         Width           =   7215
         Begin VB.ComboBox LstTarget 
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   1200
            Width           =   5895
         End
         Begin VB.ComboBox LstAI 
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   480
            Width           =   5895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ŀ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   9
            Left            =   360
            TabIndex        =   73
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AIָ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   8
            Left            =   240
            TabIndex        =   72
            Top             =   600
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2415
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   7215
         Begin VB.CheckBox chkBanditness 
            Caption         =   "����ǿ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   45
            Top             =   1800
            Width           =   1815
         End
         Begin VB.ComboBox cbAggressiveness 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1125
            Width           =   2175
         End
         Begin VB.ComboBox cbCourage 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   525
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*����8��ʾ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   13
            Left            =   3720
            TabIndex        =   69
            Top             =   570
            Width           =   1380
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   4
            Left            =   645
            TabIndex        =   42
            Top             =   1170
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   2
            Left            =   840
            TabIndex        =   41
            Top             =   570
            Width           =   495
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps2"
      Height          =   2295
      Index           =   2
      Left            =   6360
      TabIndex        =   7
      Top             =   600
      Width           =   6495
      Begin VB.Frame Frame1 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   3255
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   7215
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   16
            Left            =   3360
            TabIndex        =   68
            Top             =   2640
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   15
            Left            =   3360
            TabIndex        =   67
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   14
            Left            =   3360
            TabIndex        =   66
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   13
            Left            =   3360
            TabIndex        =   65
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   8
            Left            =   960
            TabIndex        =   64
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   12
            Left            =   3360
            TabIndex        =   35
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   7
            Left            =   960
            TabIndex        =   34
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   6
            Left            =   960
            TabIndex        =   33
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   32
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   31
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   30
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   29
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   28
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   9
            Left            =   960
            TabIndex        =   27
            Top             =   2640
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Я��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1575
         Index           =   8
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   7215
         Begin VB.TextBox txtItemsCount 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   48
            Top             =   405
            Width           =   5175
         End
         Begin VB.TextBox txtMoney 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1470
            TabIndex        =   47
            Top             =   885
            Width           =   5175
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʒ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   7
            Left            =   645
            TabIndex        =   50
            Top             =   480
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Я�����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   6
            Left            =   495
            TabIndex        =   49
            Top             =   960
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ʶ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1695
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   7215
         Begin VB.OptionButton OptLabelSize 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   62
            Top             =   1080
            Width           =   735
         End
         Begin VB.OptionButton OptLabelSize 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   61
            Top             =   1080
            Width           =   615
         End
         Begin VB.OptionButton OptLabelSize 
            Caption         =   "С"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   60
            Top             =   1080
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.ComboBox cbIcons 
            Height          =   300
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   405
            Width           =   5295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ǩ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   12
            Left            =   720
            TabIndex        =   63
            Top             =   1095
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ͼ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   5
            Left            =   720
            TabIndex        =   46
            Top             =   480
            Width           =   495
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
      Height          =   3975
      Index           =   0
      Left            =   6720
      TabIndex        =   4
      Top             =   3360
      Width           =   4695
      Begin VB.Frame Frame1 
         Caption         =   "ģ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1095
         Index           =   10
         Left            =   240
         TabIndex        =   90
         Top             =   3000
         Width           =   7215
         Begin VB.ComboBox cbTemplate 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   400
            Width           =   5175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ģ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   24
            Left            =   720
            TabIndex        =   91
            Top             =   480
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1335
         Index           =   2
         Left            =   240
         TabIndex        =   56
         Top             =   4320
         Width           =   7215
         Begin VB.TextBox txtMenu 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   57
            Top             =   360
            Width           =   5775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*�˴������ѷϳ����Ƽ���Ϊ0"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   11
            Left            =   360
            TabIndex        =   59
            Top             =   960
            Width           =   2550
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�˵�:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   15
            Left            =   480
            TabIndex        =   58
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ʶ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2535
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   7215
         Begin VB.ComboBox cbFaction 
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1680
            Width           =   5175
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   13
            Top             =   1200
            Width           =   5175
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   12
            Top             =   720
            Width           =   5175
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   10
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ӫ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   14
            Left            =   1080
            TabIndex        =   40
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������(NOW):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   3
            Left            =   360
            TabIndex        =   14
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������(EN):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   11
            Top             =   795
            Width           =   1110
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ID:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   9
            Top             =   315
            Width           =   705
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps4"
      Height          =   2535
      Index           =   4
      Left            =   10680
      TabIndex        =   75
      Top             =   960
      Width           =   2295
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�򿪿������ǵ�ͼ(&M)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   7200
         Width           =   5175
      End
      Begin VB.Frame Frame1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   4935
         Index           =   5
         Left            =   240
         TabIndex        =   81
         Top             =   2160
         Width           =   7335
         Begin VB.PictureBox Pic1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   3255
            Left            =   1680
            MousePointer    =   2  'Cross
            ScaleHeight     =   215
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   279
            TabIndex        =   84
            Top             =   1080
            Width           =   4215
         End
         Begin VB.TextBox txtLocation 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   83
            Top             =   360
            Width           =   5295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   22
            Left            =   6120
            TabIndex        =   88
            Top             =   2520
            Width           =   210
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   21
            Left            =   1320
            TabIndex        =   87
            Top             =   2520
            Width           =   210
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   19
            Left            =   3600
            TabIndex        =   86
            Top             =   4440
            Width           =   210
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   18
            Left            =   3600
            TabIndex        =   85
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ƕ�(�Ƕ���):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   17
            Left            =   360
            TabIndex        =   82
            Top             =   405
            Width           =   1290
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "λ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1575
         Index           =   7
         Left            =   240
         TabIndex        =   76
         Top             =   360
         Width           =   7335
         Begin VB.TextBox txtLocation 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   80
            Top             =   960
            Width           =   5655
         End
         Begin VB.TextBox txtLocation 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   78
            Top             =   360
            Width           =   5655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Y:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   16
            Left            =   480
            TabIndex        =   79
            Top             =   1005
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������X:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   20
            Left            =   480
            TabIndex        =   77
            Top             =   405
            Width           =   795
         End
      End
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "�����ǰ����(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8760
      Width           =   2415
   End
   Begin VB.CommandButton CQuery 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   5520
      TabIndex        =   22
      Top             =   80
      Width           =   375
   End
   Begin VB.CommandButton CQuery 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   5160
      TabIndex        =   21
      Top             =   80
      Width           =   375
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   600
      TabIndex        =   20
      Top             =   80
      Width           =   4575
   End
   Begin VB.CommandButton CReset 
      BackColor       =   &H00C0C0C0&
      Caption         =   "����(&R)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   17
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton CApply 
      BackColor       =   &H000000FF&
      Caption         =   "����(&A)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8760
      Width           =   2175
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   720
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParties.frx":123C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParties.frx":1396
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParties.frx":14F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParties.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParties.frx":17A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   8415
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   14843
      MultiRow        =   -1  'True
      ImageList       =   "IL1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��������(&P)"
            Key             =   "PropBag"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ϊ(&A)"
            Key             =   "Actions"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&E)"
            Key             =   "Special"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ա(&M)"
            Key             =   "Member"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ʼ�趨(&I)"
            Key             =   "mInitSets"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstParties 
      Height          =   8535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   15055
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmParties.frx":18FE
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&D)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   3600
      MouseIcon       =   "frmParties.frx":5F6B
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   9135
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&U)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   2
      Left            =   2760
      MouseIcon       =   "frmParties.frx":6275
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Tag             =   "edit_3"
      Top             =   9135
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯ:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   160
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   135
      TabIndex        =   15
      Top             =   9135
      Width           =   690
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   180
      Index           =   1
      Left            =   4440
      MouseIcon       =   "frmParties.frx":657F
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   9135
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ɾ��(&E)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   5280
      MouseIcon       =   "frmParties.frx":6889
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   9135
      Width           =   705
   End
End
Attribute VB_Name = "frmParties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CustomActive As Boolean
Dim CntP As tPoint

Private Sub CApply_Click()
Dim q As Boolean
If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then

    If UCase(Parties(CurrentPartyID).strID) <> UCase(CurrentParty.strID) Then              '����
        q = ChangeStrID(Parties(CurrentPartyID).strID, CurrentParty.strID)
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentParty.strID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    Parties(CurrentPartyID) = CurrentParty
    LstParties.ListItems(CurrentPartyID + 1).SubItems(1) = Parties(CurrentPartyID).csvName
    LstParties.ListItems(CurrentPartyID + 1).SubItems(2) = Parties(CurrentPartyID).strID
    
    CurrentParty = Parties(CurrentPartyID)
    LoadPartyInfo
End If

End Sub



Private Sub cbAggressiveness_Click()
If CustomActive Then
   SetPersonalities
End If
End Sub

Private Sub cbAggressiveness_Scroll()
Call cbAggressiveness_Click
End Sub

Private Sub cbCourage_Click()
If CustomActive Then
   SetPersonalities
End If
End Sub

Private Sub cbCourage_Scroll()
Call cbCourage_Click
End Sub

Private Sub cbFaction_Click()
If CustomActive Then
    CurrentParty.Faction = cbFaction.ListIndex
    CurrentParty.Faction_strID = Factions(CurrentParty.Faction).strID
End If
End Sub

Private Sub cbFaction_Scroll()
Call cbFaction_Click
End Sub

Private Sub cbTemplate_Click()
If CustomActive Then
    CurrentParty.Template = cbTemplate.ListIndex
    CurrentParty.Template_strID = PTs(CurrentParty.Template).ptID
End If
End Sub

Private Sub cbTemplate_Scroll()
Call cbTemplate_Click
End Sub


Private Sub cbIcons_Click()
Dim fI64_NOW As Integer64b

If CustomActive Then
fI64_NOW = StrToI64(CurrentParty.Flags)
fI64_NOW.by(0) = cbIcons.ListIndex

CurrentParty.Flags = I64toStrNZ(fI64_NOW)
CurrentParty.MapIcon_strID = MapIcons(cbIcons.ListIndex).strID
End If
End Sub

Private Sub cbIcons_Scroll()
Call cbIcons_Click
End Sub



Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If Parties(CurrentPartyID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), Parties(CurrentPartyID).strID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then

              If CurrentPartyID < N_Party - 1 Then
                DelIndex Parties(CurrentPartyID).strID
                For i = CurrentPartyID To N_Party - 2 Step 1
                    ChangeID Parties(i + 1).strID, Parties(i + 1).ID - 1
                    j = Parties(i).ID
                    Parties(i) = Parties(i + 1)
                    Parties(i).ID = j
                    LstParties.ListItems(i + 1).SubItems(1) = LstParties.ListItems(i + 2).SubItems(1)
                    LstParties.ListItems(i + 1).SubItems(2) = LstParties.ListItems(i + 2).SubItems(2)
                Next i
                
                ReDim Preserve Parties(N_Party - 2)
                LstParties.ListItems.Remove N_Party
                N_Party = N_Party - 1
                
              Else
                DelIndex Parties(CurrentPartyID).strID
                ReDim Preserve Parties(N_Party - 2)
                LstParties.ListItems.Remove N_Party
                
                N_Party = N_Party - 1
                CurrentPartyID = N_Party - 1
                
              End If
               
               LstParties_ItemClick LstParties.ListItems(CurrentPartyID + 1)
               LstParties.ListItems(CurrentPartyID + 1).Selected = True
               LstParties.ListItems(CurrentPartyID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), Parties(CurrentPartyID).strID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), Parties(CurrentPartyID).strID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_Party, Parties(CurrentPartyID).strID & "_New") Then
           
           ReDim Preserve Parties(N_Party)
           N_Party = N_Party + 1
           Parties(N_Party - 1) = Parties(CurrentPartyID)
           With Parties(N_Party - 1)
                 .ID = N_Party - 1
                 .id2 = .ID
                 .strID = .strID & "_New"
                 .strName = .strName & "_New"
                 .csvName = .csvName & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstParties.ListItems.Add(, "Parties_" & Parties(N_Party - 1).ID, Parties(N_Party - 1).ID)
      
                 With oItem
                    .SubItems(1) = Parties(N_Party - 1).csvName
                    .SubItems(2) = Parties(N_Party - 1).strID
                 End With
           LstParties_ItemClick LstParties.ListItems(N_Party)
           LstParties.ListItems(N_Party).Selected = True
           LstParties.ListItems(N_Party).EnsureVisible
           
           Else
           
           MsgBox ActiveString(PublicMsgs(90), Parties(CurrentPartyID).strID & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
         
      Case 2
         If CurrentPartyID > 0 Then
           If Parties(CurrentPartyID - 1).Edit And Parties(CurrentPartyID).Edit Then
           
                SwapID Parties(CurrentPartyID - 1).strID, Parties(CurrentPartyID).strID
                SwapParties CurrentPartyID - 1, CurrentPartyID
                SwapListItem LstParties.ListItems(CurrentPartyID), LstParties.ListItems(CurrentPartyID + 1), 2, True
                
               LstParties_ItemClick LstParties.ListItems(CurrentPartyID)
               LstParties.ListItems(CurrentPartyID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), Parties(CurrentPartyID - 1).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
         End If
      Case 3
        If CurrentPartyID + 1 <= N_Party - 1 Then
           If Parties(CurrentPartyID).Edit And Parties(CurrentPartyID + 1).Edit Then
                
                SwapID Parties(CurrentPartyID).strID, Parties(CurrentPartyID + 1).strID
                SwapParties CurrentPartyID, CurrentPartyID + 1
                SwapListItem LstParties.ListItems(CurrentPartyID + 1), LstParties.ListItems(CurrentPartyID + 2), 2, True
                
                LstParties_ItemClick LstParties.ListItems(CurrentPartyID + 2)
                LstParties.ListItems(CurrentPartyID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), Parties(CurrentPartyID).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
        End If
End Select

N_Party2 = N_Party
InitAITargetList
End Sub

Private Sub chkBanditness_Click()
If CustomActive Then
   SetPersonalities
End If
End Sub

Private Sub chkFlags_Click(Index As Integer)
Dim tI(2) As Integer64b

If CustomActive Then

With CurrentParty
    tI(1) = StrToI64(.Flags)
    tI(2) = HexStrToI64(chkFlags(Index).Tag)
    If chkFlags(Index).Value = 0 Then
      tI(0) = DeleteFlags64b(tI(1), tI(2))
    Else
      tI(0) = AddFlags64b(tI(1), tI(2))
    End If
    .Flags = I64toStrNZ(tI(0))
End With

End If
End Sub



Private Sub chkSlave_Click()
If CustomActive Then
    LstTroops.SelectedItem.SubItems(4) = GetYesNoStr(chkSlave.Value)
    CurrentParty.Stacks(LstTroops.SelectedItem.Index).Flags = chkSlave.Value
End If
End Sub

Private Sub Command1_Click()
frmMap.Show
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstParties, LstParties.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstParties_ItemClick(LstParties.SelectedItem)
End If
End Sub

Private Sub CQuery2_Click(Index As Integer)
On Error GoTo EL
Dim q As Boolean
q = QueryItem(boxTroops, boxTroops.SelectedItem.Index, txtQuery2.Text, CBool(Index))

Exit Sub

EL:
  Call logErr("frmParties", "CQuery_Click2(" & Index & ")", Err.Number, Err.Description)
End Sub


Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
LstParties_ItemClick LstParties.ListItems(CurrentPartyID + 1)
LstParties.ListItems(CurrentPartyID + 1).Selected = True
LstParties.ListItems(CurrentPartyID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
CustomActive = False

InitFrames

InitIconsCombo
InitPartiesListView
InitTroopsListView
InitFlagsList
InitAIList
InitAITargetList
InitTroopBox

LoadPartiesList
LoadFactionCombo
'LoadSelectTroopCombo
LoadTroopBox
LoadPersonalityCombo
LoadTemplateCombo

CurrentPartyID = 0
CurrentParty = Parties(CurrentPartyID)
LoadPartyInfo

TranslateForm Me
InitCMDs

Label2.Caption = Label2.Caption & N_Party
Label1(3).Caption = Replace(Label1(3).Caption, "(NOW):", "(" & UCase$(MnBInfo.Language) & "):", , , vbTextCompare)

If LCase$(MnBInfo.Language) = "en" Then
    txtPropBag(2).Enabled = False
End If

CustomActive = True
Me.Show

CntP = SetPoint(Pic1.ScaleWidth / 2, Pic1.ScaleHeight / 2)
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

'*************************************************************************
'**�� �� ����DeleteCurrentStack
'**��    �룺(Long)StackIndex
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 22:17:57
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub DeleteCurrentStack(ByVal StackIndex As Long)
Dim i As Long
With CurrentParty
     If .StacksCount > 0 Then
          .StacksCount = .StacksCount - 1
          If .StacksCount = 0 Then
             'ReDim .Stacks(0)
             ClearStack .Stacks(1)
          Else
              For i = StackIndex To .StacksCount
                   .Stacks(i) = .Stacks(i + 1)
              Next i
              
              'ReDim Preserve .Stacks(1 To .StacksCount)
              ClearStack .Stacks(.StacksCount + 1)
          End If
          
     End If
End With
End Sub

'*************************************************************************
'**�� �� ����CreateCurrentStack
'**��    �룺(Long)Trp_ID
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 22:35:46
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub CreateCurrentStack(ByVal Trp_ID As Long)
Dim i As Long
If Trp_ID > -1 Then
With CurrentParty
     If .StacksCount > 0 Then
          .StacksCount = .StacksCount + 1
          ReDim Preserve .Stacks(1 To .StacksCount)
     Else
          .StacksCount = 1
          ReDim .Stacks(1 To .StacksCount)
     End If
     
     .Stacks(.StacksCount).ID = Trp_ID
     .Stacks(.StacksCount).strID = Trps(Trp_ID).strID
End With
End If
End Sub
'*************************************************************************
'**�� �� ����LoadPartiesList
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-06 10:48:30
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadPartiesList()
Dim n As Long, oItem As ListItem

LstParties.ListItems.Clear
For n = 0 To N_Party - 1

      Set oItem = LstParties.ListItems.Add(, "P_" & Parties(n).ID, Parties(n).ID)
      
      With oItem
         .SubItems(1) = Parties(n).csvName
         .SubItems(2) = Parties(n).strID
      End With
Next n

End Sub

'*************************************************************************
'**�� �� ����InitAITargetList
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 23:13:27
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitAITargetList()
Dim i As Long

LstTarget.Clear
For i = 0 To N_Party - 1
       LstTarget.AddItem i & " " & Parties(i).csvName & "|" & Parties(i).strID
Next i

End Sub
'*************************************************************************
'**�� �� ����LoadFactionCombo
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-24 23:55:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadFactionCombo()
Dim n As Long, oItem As ListItem

cbFaction.Clear
For n = 0 To UBound(Factions)
      cbFaction.AddItem Factions(n).csvName
Next n

End Sub

Private Sub InitPartiesListView()
Dim n As Integer
n = 3
LstParties.View = lvwReport
LstParties.Sorted = False
LstParties.ListItems.Clear
LstParties.ColumnHeaders.Clear
LstParties.SortOrder = lvwAscending
LstParties.FullRowSelect = True
LstParties.AllowColumnReorder = False
LstParties.LabelEdit = lvwManual
LstParties.Checkboxes = False
LstParties.GridLines = True
LstParties.MultiSelect = False
LstParties.HideSelection = False

LstParties.ColumnHeaders.Add , , PublicMsgs(13), LstParties.Width / n / 3.6
LstParties.ColumnHeaders.Add , , PublicEditors(3) & PublicMsgs(14), LstParties.Width / n * 1.5
LstParties.ColumnHeaders.Add , , PublicEditors(3) & "ID", LstParties.Width / n * 1.5


End Sub

Private Sub InitTroopsListView()
Dim n As Integer
n = 6
With LstTroops
   .View = lvwReport
   .Sorted = False
   .ListItems.Clear
   .ColumnHeaders.Clear
   .SortOrder = lvwAscending
   .FullRowSelect = True
   .AllowColumnReorder = False
   .LabelEdit = lvwManual
   .Checkboxes = True
   .GridLines = True
   .MultiSelect = False
   .HideSelection = False

   .ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(13), .Width / n / 1.8
   .ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(14), .Width / n * 2
   .ColumnHeaders.Add , , PublicEditors(1) & "ID", .Width / n * 2
   .ColumnHeaders.Add , , PublicMsgs(46), .Width / n / 1.5
   .ColumnHeaders.Add , , PublicMsgs(47), .Width / n / 1.5

End With

End Sub

Private Sub InitFrames()
Dim i As Integer

For i = 0 To FraProps.UBound
    With FraProps(i)
         .BorderStyle = 0
         .Top = Tab1.ClientTop
         .Left = Tab1.ClientLeft
         .Width = Tab1.ClientWidth
         .Height = Tab1.ClientHeight
         .ZOrder
           If i <> 0 Then
            .Visible = False
           End If
    End With
Next i

If LCase$(MnBInfo.Language) = "en" Then
     txtPropBag(2).Enabled = False
End If

End Sub



Private Sub LstAI_Click()
If CustomActive Then
      CurrentParty.AI_Behavior = LstAI.ListIndex
End If
End Sub

Private Sub LstAI_Scroll()
Call LstAI_Click
End Sub

Private Sub LstTarget_Click()
If CustomActive Then
      CurrentParty.AI_Target = LstTarget.ListIndex
      CurrentParty.AI_Target_strID = Parties(CurrentParty.AI_Target).strID
End If
End Sub

Private Sub LstTarget_Scroll()
Call LstTarget_Click
End Sub

Private Sub LstParties_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentPartyID = Val(Item.Text)

CurrentParty = Parties(CurrentPartyID)
LoadPartyInfo
End Sub

Private Sub OptLabelSize_Click(Index As Integer)
Dim tI(3) As Integer64b

If CustomActive Then
With CurrentParty
    tI(1) = StrToI64(.Flags)     'ԭ����Flags
    tI(2) = HexStrToI64(pf_label_mask)   '���в�������Flags
    
    tI(3) = HexStrToI64(OptLabelSize(Index).Tag)     'Ҫ��ӵĲ�������Flags
    
    tI(0) = DeleteFlags64b(tI(1), tI(2))       '������в�������Flags
    tI(0) = AddFlags64b(tI(0), tI(3))             '��Ӳ�������Flags
    .Flags = I64toStrNZ(tI(0))
End With
End If
End Sub



Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MosP As tPoint, TemDeg As Single

If Button = vbLeftButton Then
CustomActive = False
MosP = SetPoint(X, Y)
TemDeg = GetDegree(MosP, CntP)
TemDeg = FuncDegreeStandardize2(TemDeg + Pi / 2)

With CurrentParty
       .Degree = Format(TemDeg, "0.000000")
       txtLocation(2).Text = RadToDeg(CSng(Val(.Degree)))
       InitCompass
       DrawArrow CSng(Val(.Degree))
End With

CustomActive = True
End If

End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MosP As tPoint, TemDeg As Single

If Button = vbLeftButton Then
CustomActive = False
MosP = SetPoint(X, Y)
TemDeg = GetDegree(MosP, CntP)
TemDeg = FuncDegreeStandardize2(TemDeg + Pi / 2)

With CurrentParty
       .Degree = Format(TemDeg, "0.000000")
       txtLocation(2).Text = RadToDeg(CSng(Val(.Degree)))
       InitCompass
       DrawArrow CSng(Val(.Degree))
End With

CustomActive = True
End If

End Sub

Private Sub Tab1_Click()
Dim i As Integer

For i = 0 To FraProps.UBound
    With FraProps(i)
         .Visible = i + 1 = Tab1.SelectedItem.Index
    End With
Next i

LoadPartyInfo
End Sub

'*************************************************************************
'**�� �� ����LoadPartyInfo
'**��    �룺-
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-06 10:58:11
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadPartyInfo()
On Error GoTo EL
Dim i As Integer, oItem As ListItem, strTem As String, lngTem As Long
CustomActive = False
With CurrentParty
  Select Case Tab1.SelectedItem.Index
       Case 1
          'Recognize
          txtPropBag(0).Text = .strID
          txtPropBag(1).Text = .strName
          txtPropBag(2).Text = .csvName
          
          '.Faction = IIf(CheckExist(EditInfo_FactionsCount, .Faction), .Faction, 0)
          .Faction = GetID(.Faction_strID, , Factions(0).strID)
          cbFaction.ListIndex = .Faction

          '.Template = IIf(CheckExist(EditInfo_PartiesCount, .Template), .Template, 0)
          .Template = GetID(.Template_strID, , PTs(0).ptID)
          cbTemplate.ListIndex = .Template
          
          txtMenu.Text = .Menu
       Case 2
          LoadPersonalities
          LstAI.ListIndex = .AI_Behavior
          
          .AI_Target = GetID(.AI_Target_strID, , Parties(0).strID)
          LstTarget.ListIndex = .AI_Target
       Case 3
          LoadFlagsList
       Case 4
          LoadTroopsListView
       Case 5
          InitCompass
          txtLocation(0).Text = .InitPos(1).X
          txtLocation(1).Text = .InitPos(1).Y
          txtLocation(2).Text = RadToDeg(CSng(Val(.Degree)))
          DrawArrow CSng(Val(.Degree))
  End Select
End With
CustomActive = True

Exit Sub

EL:
   Call logErr("frmParties", "LoadPartyInfo", Err.Number, Err.Description)
End Sub


Private Sub txtItemsCount_LostFocus()
Dim fI64_NOW As Integer64b
If CustomActive Then

fI64_NOW = StrToI64(CurrentParty.Flags)
fI64_NOW.by(pf_carry_goods_bits \ 8) = Int(Val(txtItemsCount.Text))

CurrentParty.Flags = I64toStrNZ(fI64_NOW)

End If
End Sub


Private Sub txtLocation_Change(Index As Integer)
Dim m As Integer
If CustomActive Then
  With CurrentParty
    Select Case Index
         Case 0
            For m = 1 To 3
               .InitPos(m).X = Format(txtLocation(Index).Text, "0.000000")
            Next m
         Case 1
            For m = 1 To 3
               .InitPos(m).Y = Format(txtLocation(Index).Text, "0.000000")
            Next m
         Case 2
            .Degree = Format(DegToRad((CSng(Val(txtLocation(Index).Text)))), "0.000000")
            InitCompass
            DrawArrow CSng(Val(.Degree))
    End Select
  End With
End If
End Sub

Private Sub txtMenu_Change()
If CustomActive Then
    CurrentParty.Menu = Val(txtMenu.Text)
End If
End Sub

Private Sub txtMin_Change()
If CustomActive Then
    LstTroops.SelectedItem.SubItems(3) = TxtMin.Text
    CurrentParty.Stacks(LstTroops.SelectedItem.Index).Min = Val(TxtMin.Text)
End If
End Sub

Private Sub txtMoney_LostFocus()
Dim fI64_NOW As Integer64b

If CustomActive Then

fI64_NOW = StrToI64(CurrentParty.Flags)
fI64_NOW.by(pf_carry_gold_bits \ 8) = Int(Val(txtMoney.Text)) / pf_carry_gold_multiplier

CurrentParty.Flags = I64toStrNZ(fI64_NOW)

End If
End Sub

Private Sub txtPropBag_LostFocus(Index As Integer)
If CustomActive Then

With CurrentParty

If Index < 2 Then
   txtPropBag(Index).Text = Replace(txtPropBag(Index).Text, " ", "_")
End If

Select Case Index
      Case 0
            'Check Value
            If Left$(txtPropBag(Index).Text, 2) <> "p_" Then
                txtPropBag(Index).Text = "p_" & txtPropBag(Index).Text
            End If
         .strID = txtPropBag(Index).Text
      Case 1
         .strName = txtPropBag(Index).Text
           If LCase$(MnBInfo.Language) = "en" Then
             .csvName = .strName
           End If
      Case 2
           If LCase$(MnBInfo.Language) <> "en" Then
             .csvName = txtPropBag(Index).Text
           End If
End Select

End With

End If
End Sub



'*************************************************************************
'**�� �� ����LoadTemplateCombo
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-14 22:59:56
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadTemplateCombo()
Dim i As Long
cbTemplate.Clear

         For i = 0 To N_PT - 1
                  cbTemplate.AddItem "(" & i & ")" & PTs(i).csvName & "|" & PTs(i).ptID
                  cbTemplate.ItemData(i) = i
         Next i
         

End Sub

'*************************************************************************
'**�� �� ����InitFlagsList
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-05 22:52:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitFlagsList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("������", "��", "��̬", "С��ǩ", "�б�ǩ", "���ǩ", "���ǿɼ�", "��ΪĬ��", "�ڳ������Զ�ȥ��", "��������", _
               "�ޱ�ǩ", "��������", "����", "��ʾ��Ӫ", "���ɼ�", "������ƽ��", "ƽ��")
TemArray = Array(pf_disabled, pf_is_ship, pf_is_static, pf_label_small, pf_label_medium, pf_label_large, pf_always_visible, pf_default_behavior, pf_auto_remove_in_town, pf_quest_party, _
               pf_no_label, pf_limit_members, pf_hide_defenders, pf_show_faction, pf_is_hidden, pf_dont_attack_civilians, pf_civilian)
         For i = 0 To UBound(strTemArray)
            If i < 3 Or i > 5 Then
               chkFlags(i).Caption = strTemArray(i)
               chkFlags(i).Tag = TemArray(i)
               
               If i = 14 Then
                  chkFlags(i).Enabled = False
               End If
            Else
               OptLabelSize(i - 3).Tag = TemArray(i)
            End If
         Next i

End Sub

'*************************************************************************
'**�� �� ����LoadFlagsList
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-06 21:49:44
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadFlagsList()
Dim i As Integer, fI64 As Integer64b, fI64_NOW As Integer64b, resI64 As Integer64b, Money As Long, Items As Long, NeedCheckSmallLabel As Boolean
Dim lngTem As Long

With CurrentParty

fI64_NOW = StrToI64(.Flags)
For i = 0 To chkFlags.UBound
            If i < 3 Or i > 5 Then
               fI64 = HexStrToI64(chkFlags(i).Tag)
               resI64 = And64b(fI64_NOW, fI64)
               
               If IsEqual64b(resI64, fI64) Then
                  chkFlags(i).Value = 1
               Else
                  chkFlags(i).Value = 0
               End If
            Else
               
            End If
Next i

NeedCheckSmallLabel = True

For i = 4 To 5
       fI64 = HexStrToI64(OptLabelSize(i - 3).Tag)
       resI64 = And64b(fI64_NOW, fI64)
               
       If IsEqual64b(resI64, fI64) Then
           OptLabelSize(i - 3).Value = True
           NeedCheckSmallLabel = False
       Else
           OptLabelSize(i - 3).Value = False
       End If
Next i
OptLabelSize(0).Value = NeedCheckSmallLabel


Items = fI64_NOW.by(pf_carry_goods_bits \ 8)
Money = fI64_NOW.by(pf_carry_gold_bits \ 8)

txtMoney.Text = Money * pf_carry_gold_multiplier
txtItemsCount.Text = Items

lngTem = GetID(.MapIcon_strID, , MapIcons(0).strID)
cbIcons.ListIndex = lngTem

If lngTem <> fI64_NOW.by(0) Then
   fI64_NOW.by(0) = lngTem
   .Flags = I64toStrNZ(fI64_NOW)
End If

End With
End Sub

'*************************************************************************
'**�� �� ����LoadPersonalities
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-06 17:22:53
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadPersonalities()
On Error GoTo EL
Dim Cour As Long, Aggr As Long, IsBandit As Long

With CurrentParty
      Cour = .Personality(1) Mod 16
      Aggr = (.Personality(1) \ 16) Mod 16
      IsBandit = .Personality(1) \ (16 * 16)
      
      cbCourage.ListIndex = Cour
      cbAggressiveness.ListIndex = Aggr
      chkBanditness.Value = IsBandit
End With

Exit Sub

EL:
   Call logErr("frmParties", "LoadPersonalities", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SetPersonalities
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-07 14:31:31
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub SetPersonalities()
Dim Cour As Long, Aggr As Long, IsBandit As Long

With CurrentParty
      Cour = cbCourage.ListIndex
      Aggr = cbAggressiveness.ListIndex
      IsBandit = chkBanditness.Value
      
      .Personality(1) = Cour + Aggr * 16 + IsBandit * 16 * 16
      .Personality(2) = .Personality(1)
End With
End Sub

'*************************************************************************
'**�� �� ����QueryItem
'**��    �룺(ListItem)oLV,(Long)Start,(String)QueryString,(Boolean)bReverse
'**��    ����
'**��������������ListView��ѯ����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-27 21:19:17
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************

Private Function QueryItem(oLV As ListView, ByVal Start As Long, ByVal QueryString As String, Optional bReverse As Boolean = False) As Boolean
Dim oItem As ListItem
  With oLV
    Set oItem = FindItem(oLV, Start, "0|1|2", QueryString, True, vbTextCompare, bReverse)
       If Not oItem Is Nothing Then
       .ListItems(oItem.Index).Selected = True
       .ListItems(oItem.Index).EnsureVisible
       QueryItem = True
       Else
        MsgBox PublicMsgs(11), vbInformation, PublicMsgs(12)
        QueryItem = False
       End If
  End With
    Set oItem = Nothing
End Function

Private Sub InitIconsCombo()
Dim i As Integer
CustomActive = False

cbIcons.Clear
For i = 0 To N_MapIcon - 1
      cbIcons.AddItem MapIcons(i).strID
Next i

cbIcons.ListIndex = 0
CustomActive = True
End Sub

'*************************************************************************
'**�� �� ����LoadPersonalityCombo
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-06 11:36:01
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadPersonalityCombo()
Dim i As Long
CustomActive = False
cbAggressiveness.Clear
cbCourage.Clear
         For i = 0 To 15
                cbAggressiveness.AddItem i
                cbAggressiveness.ItemData(i) = i * 16
                cbCourage.AddItem i
                cbAggressiveness.ItemData(i) = i
         Next i
CustomActive = True
End Sub

Private Sub SetTroopExist(ByVal Switch As Boolean)
CustomActive = False

If Not Switch Then
  TxtMin.Text = ""
  chkSlave.Value = 0
End If

TxtMin.Enabled = Switch
chkSlave.Enabled = Switch

CustomActive = True
End Sub

'*************************************************************************
'**�� �� ����InitAIList
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-06 16:44:25
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitAIList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("����", "���е�����", "��ĳ��Ѳ��", "��ĳ���Ӵ�Ѳ��", "�������ӹ���", _
                    "�������Ӷ��", "���е�ĳ��", "�������ӽ���", "�ڳ���", "���е���", "���Ͳ���", "����������")
               
TemArray = Array(ai_bhvr_hold, ai_bhvr_travel_to_party, ai_bhvr_patrol_location, ai_bhvr_patrol_party, ai_bhvr_attack_party, _
                 ai_bhvr_avoid_party, ai_bhvr_travel_to_point, ai_bhvr_negotiate_party, ai_bhvr_in_town, ai_bhvr_travel_to_ship, ai_bhvr_escort_party, ai_bhvr_driven_by_party)

LstAI.Clear
For i = 0 To UBound(strTemArray)
      LstAI.AddItem strTemArray(i)
      LstAI.ItemData(i) = TemArray(i)
Next i

End Sub

'*************************************************************************
'**�� �� ����InitCompass
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-09 13:58:45
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitCompass()
Dim tP(1) As tPoint, i As Integer, l As Single, dL As Single

Pic1.Cls
With Pic1
    l = .ScaleHeight / 2 * 4 / 4: dL = .ScaleWidth / 2 * 1 / 10
    
    Pic1.ForeColor = vbRed
    Pic1.FillStyle = 1
    Pic1.Circle (CntP.X, CntP.Y), l
    Pic1.ForeColor = vbBlue
    For i = 0 To 7
        tP(0) = PoltoRec(l, i * Pi / 4, CntP)
        tP(1) = PoltoRec(l - dL, i * Pi / 4, CntP)
        
        Pic1.Line (tP(0).X, tP(0).Y)-(tP(1).X, tP(1).Y)
    Next i
End With
End Sub

'*************************************************************************
'**�� �� ����DrawArrow
'**��    �룺(Single)Degree
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-09 14:14:14 (��_������ѽ)
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub DrawArrow(ByVal Degree As Single)
Dim tP(3) As tPoint, l As Single, dL As Single, DrawDegree As Single

With Pic1
    DrawDegree = Degree - Pi / 2
    l = .ScaleHeight / 2 * 3 / 4: dL = .ScaleWidth / 2 * 1 / 20
    
    Pic1.ForeColor = RGB(10, 240, 10)
    Pic1.FillStyle = 0
    Pic1.FillColor = RGB(10, 240, 10)
    Pic1.Circle (CntP.X, CntP.Y), dL
        tP(0) = PoltoRec(l, DrawDegree, CntP)
        tP(1) = PoltoRec(l - dL, DrawDegree, CntP)
        tP(2) = PoltoRec(dL / 2, DrawDegree + Pi / 2, tP(1))
        tP(3) = PoltoRec(dL / 2, DrawDegree - Pi / 2, tP(1))
        
        Pic1.DrawWidth = 2
        Pic1.Line (tP(0).X, tP(0).Y)-(CntP.X, CntP.Y)
        Pic1.Line (tP(0).X, tP(0).Y)-(tP(2).X, tP(2).Y)
        Pic1.Line (tP(0).X, tP(0).Y)-(tP(3).X, tP(3).Y)
        Pic1.DrawWidth = 1
End With

End Sub

Private Sub InitTroopBox()
Dim n As Integer
n = 3
With boxTroops
   .View = lvwReport
   .Sorted = False
   .ListItems.Clear
   .ColumnHeaders.Clear
   .SortOrder = lvwAscending
   .FullRowSelect = True
   .AllowColumnReorder = False
   .LabelEdit = lvwManual
   .Checkboxes = True
   .GridLines = True
   .MultiSelect = False
   .HideSelection = False

   .ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(13), .Width / n / 2
   .ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(14), .Width / n * 1.2
   .ColumnHeaders.Add , , PublicEditors(1) & "ID", .Width / n * 1.2

End With
End Sub


Private Sub LoadTroopBox()
Dim n As Long, oItem As ListItem

For n = 0 To N_Troop - 1
  With Trps(n)
     Set oItem = boxTroops.ListItems.Add(, , .ID)
     oItem.SubItems(1) = .csvName
     oItem.SubItems(2) = .strID
  End With
Next n

End Sub

'*************************************************************************
'**�� �� ����StructurePartyStacks
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-03-20 11:15:02
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub StructurePartyStacks()
Dim i As Integer, TemList() As Type_Stacks

With CurrentParty
ReDim TemList(0)
For i = 1 To .StacksCount
        If .Stacks(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Stacks(i)
        End If
Next i

.StacksCount = UBound(TemList)

If .StacksCount > 0 Then
ReDim .Stacks(1 To .StacksCount)
   For i = 1 To .StacksCount
           .Stacks(i) = TemList(i)
   Next i
End If
End With
End Sub


'*************************************************************************
'**�� �� ����LoadTroopsListView
'**��    �룺-
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-03-20 11:26:07
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadTroopsListView()
Dim i As Integer, oItem As ListItem, q As Boolean
With CurrentParty
LstTroops.ListItems.Clear

If .StacksCount > 0 Then
         For i = 1 To .StacksCount
            .Stacks(i).ID = GetID(.Stacks(i).strID, True, "", -1)
            
            If .Stacks(i).ID = -1 Then q = True
            If .Stacks(i).ID > -1 Then
               Set oItem = LstTroops.ListItems.Add(, , Trps(.Stacks(i).ID).ID)
               
               With oItem
                     .SubItems(1) = Trps(CurrentParty.Stacks(i).ID).csvName
                     .SubItems(2) = Trps(CurrentParty.Stacks(i).ID).strID
                     .SubItems(3) = CurrentParty.Stacks(i).Min
                     .SubItems(4) = GetYesNoStr(CurrentParty.Stacks(i).Flags)
               End With
            End If
         Next i
       Set LstTroops.SelectedItem = LstTroops.ListItems(1)
       LstTroops_ItemClick LstTroops.ListItems(1)
               
       If q Then StructurePartyStacks
   SetTroopExist True
Else
   SetTroopExist False
End If

End With
End Sub

Private Sub bCMD_Click(Index As Integer)
On Error GoTo EL
Dim Index1 As Long, Index2 As Long
Dim i As Integer, q As Boolean

CustomActive = False
If LstTroops.ListItems.Count > 0 Then
Index1 = LstTroops.SelectedItem.Index
Select Case Index
       Case 2
            If Index1 >= 2 Then
                SwapStacks CurrentParty.Stacks(Index1), CurrentParty.Stacks(Index1 - 1)
                SwapListItem LstTroops.ListItems(Index1), LstTroops.ListItems(Index1 - 1), LstTroops.ListItems(Index1).ListSubItems.Count
                LstTroops.ListItems(Index1 - 1).Selected = True
                LstTroops.ListItems(Index1 - 1).EnsureVisible
            End If
       Case 3
            If Index1 <= LstTroops.ListItems.Count - 1 Then
                SwapStacks CurrentParty.Stacks(Index1), CurrentParty.Stacks(Index1 + 1)
                SwapListItem LstTroops.ListItems(Index1), LstTroops.ListItems(Index1 + 1), LstTroops.ListItems(Index1).ListSubItems.Count
                LstTroops.ListItems(Index1 + 1).Selected = True
                LstTroops.ListItems(Index1 + 1).EnsureVisible
            End If
       Case 0
            For i = 1 To LstTroops.ListItems.Count
                If LstTroops.ListItems(i).Checked Then
                   CurrentParty.Stacks(i).ID = -1
                   CurrentParty.Stacks(i).strID = ""
                End If
            Next i
            StructurePartyStacks
            LoadTroopsListView

End Select
End If

If Index = 1 Then
            For i = 1 To boxTroops.ListItems.Count
                If boxTroops.ListItems(i).Checked Then
                     CreateCurrentStack i - 1
                End If
                boxTroops.ListItems(i).Checked = False
            Next i
            LoadTroopsListView
ElseIf Index = 4 Then
            For i = 1 To boxTroops.ListItems.Count
                If boxTroops.ListItems(i).Checked Then
                    boxTroops.ListItems(i).Checked = False
                End If
            Next i
End If

CustomActive = True

Exit Sub

EL:
  Call logErr("frmParties", "bCMD_Click(" & Index & ")", Err.Number, Err.Description)
End Sub

Private Sub LstTroops_ItemClick(ByVal Item As MSComctlLib.ListItem)
With CurrentParty

   If .Stacks(Item.Index).ID >= 0 Then
       TxtMin.Text = .Stacks(Item.Index).Min
       chkSlave.Value = .Stacks(Item.Index).Flags
       
       TxtMin.Enabled = True
       chkSlave.Enabled = True
   End If

End With
End Sub

Public Sub ReLoadInfo()
LoadPartyInfo
End Sub

