VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItems 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "��Ʒ�༭��"
   ClientHeight    =   9375
   ClientLeft      =   4875
   ClientTop       =   1155
   ClientWidth     =   14355
   Icon            =   "frmItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_2"
   Begin VB.Timer Timer_KillTip 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   0
   End
   Begin VB.Timer Timer_MousePos 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   360
      Top             =   720
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   960
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":2358
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "�����ǰ��Ʒ(&O)"
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
      TabIndex        =   187
      Top             =   8760
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   720
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
      TabIndex        =   7
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton CReset 
      BackColor       =   &H000000FF&
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
      TabIndex        =   6
      Top             =   8760
      Width           =   2175
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
      TabIndex        =   5
      Top             =   120
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
      Index           =   1
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   720
      TabIndex        =   3
      Top             =   100
      Width           =   4455
   End
   Begin MSComctlLib.ListView LstItems 
      Height          =   8415
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   14843
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmItems.frx":24B2
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
      Height          =   7935
      Index           =   0
      Left            =   6240
      TabIndex        =   9
      Top             =   480
      Width           =   7815
      Begin VB.Frame Frame4 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2535
         Left            =   240
         TabIndex        =   86
         Top             =   5280
         Width           =   7335
         Begin VB.Frame Frame5 
            Caption         =   "��Ӫ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   2175
            Left            =   3600
            TabIndex        =   92
            Top             =   240
            Width           =   3615
            Begin VB.ListBox LstFaction 
               Height          =   1740
               ItemData        =   "frmItems.frx":6B1F
               Left            =   120
               List            =   "frmItems.frx":6B26
               Style           =   1  'Checkbox
               TabIndex        =   93
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.TextBox TxtAb 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   91
            Top             =   1740
            Width           =   1335
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   89
            Top             =   1260
            Width           =   1335
         End
         Begin VB.CheckBox ChkNext 
            Caption         =   "��һ��������Ϊ�ڶ�ʹ��ģʽ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   87
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��ԣ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   495
            TabIndex        =   90
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�۸�:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   720
            TabIndex        =   88
            Top             =   1320
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "��Ʒ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2655
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   7335
         Begin VB.OptionButton Option1 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   360
            MaskColor       =   &H00000000&
            TabIndex        =   43
            Top             =   2160
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   20
            Left            =   5760
            TabIndex        =   37
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��ƥ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   375
            Index           =   1
            Left            =   5760
            TabIndex        =   36
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   2
            Left            =   360
            TabIndex        =   35
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "˫������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   3
            Left            =   2160
            TabIndex        =   34
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   4
            Left            =   3840
            TabIndex        =   33
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��ʸ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Index           =   5
            Left            =   360
            TabIndex        =   32
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��ʸ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Index           =   6
            Left            =   2160
            TabIndex        =   31
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   375
            Index           =   7
            Left            =   3840
            TabIndex        =   30
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   8
            Left            =   360
            TabIndex        =   29
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   9
            Left            =   2160
            TabIndex        =   28
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ͷ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   375
            Index           =   10
            Left            =   3840
            TabIndex        =   27
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   11
            Left            =   5760
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "ͷ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   12
            Left            =   360
            TabIndex        =   25
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�ز�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   13
            Left            =   2160
            TabIndex        =   24
            Top             =   1800
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�Ȳ�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   14
            Left            =   3840
            TabIndex        =   23
            Top             =   1800
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�ֲ�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   375
            Index           =   15
            Left            =   5760
            TabIndex        =   22
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��ǹ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   16
            Left            =   360
            TabIndex        =   21
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "��ǹ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   17
            Left            =   2160
            TabIndex        =   20
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�ӵ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Index           =   18
            Left            =   3840
            TabIndex        =   19
            Top             =   1440
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Index           =   19
            Left            =   5760
            TabIndex        =   18
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame FrmType 
         Caption         =   "ʶ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2295
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   7335
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   3
            Left            =   2400
            TabIndex        =   194
            Top             =   1800
            Width           =   4335
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   0
            Left            =   2400
            TabIndex        =   13
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   1
            Left            =   2400
            TabIndex        =   12
            Top             =   840
            Width           =   4335
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   2
            Left            =   2400
            TabIndex        =   11
            Top             =   1320
            Width           =   4335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��Ʒ������(NOW):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   480
            TabIndex        =   193
            Top             =   1800
            Width           =   1845
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��ƷID:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1560
            TabIndex        =   16
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��Ʒ��(EN):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   1080
            TabIndex        =   15
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��Ʒ��(NOW):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   960
            TabIndex        =   14
            Top             =   1320
            Width           =   1395
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps2"
      Height          =   7935
      Index           =   2
      Left            =   6240
      TabIndex        =   40
      Top             =   480
      Width           =   7815
      Begin VB.Frame FrmMod 
         Caption         =   "ģ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   7575
         Left            =   0
         TabIndex        =   165
         Top             =   120
         Width           =   3735
         Begin VB.ListBox LstModel 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   120
            TabIndex        =   181
            Top             =   360
            Width           =   3495
         End
         Begin VB.Frame FrmAtt 
            Caption         =   "װ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   1935
            Left            =   120
            TabIndex        =   174
            Top             =   5520
            Width           =   3495
            Begin VB.OptionButton OptAtt 
               Caption         =   "ǿ�ư󶨵���ǰ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   360
               TabIndex        =   177
               Top             =   1200
               Width           =   3015
            End
            Begin VB.OptionButton OptAtt 
               Caption         =   "�󶨵�����"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   360
               TabIndex        =   176
               Top             =   1500
               Width           =   3015
            End
            Begin VB.OptionButton OptAtt 
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
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   175
               Top             =   300
               Width           =   3015
            End
            Begin VB.OptionButton OptAtt 
               Caption         =   "ǿ�ư󶨵�����"
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
               Left            =   360
               TabIndex        =   179
               Top             =   570
               Width           =   3015
            End
            Begin VB.OptionButton OptAtt 
               Caption         =   "ǿ�ư󶨵�����"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   178
               Top             =   900
               Width           =   3015
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "ǰ׺"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   2775
            Left            =   120
            TabIndex        =   173
            Top             =   2640
            Width           =   3495
            Begin VB.ListBox LstMIMod 
               Height          =   2370
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   182
               Top             =   240
               Width           =   3255
            End
         End
         Begin VB.TextBox TxtModName 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   170
            Top             =   1800
            Width           =   2415
         End
         Begin VB.ComboBox CBixmesh 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   169
            Top             =   2250
            Width           =   2415
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ģ������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   240
            TabIndex        =   172
            Top             =   1890
            Width           =   885
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "λ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   600
            TabIndex        =   171
            Top             =   2325
            Width           =   495
         End
         Begin VB.Label LbMdel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ɾ��"
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
            Left            =   3120
            MouseIcon       =   "frmItems.frx":6B36
            MousePointer    =   99  'Custom
            TabIndex        =   168
            Top             =   1560
            Width           =   390
         End
         Begin VB.Label LbMadd 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
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
            Left            =   2520
            MouseIcon       =   "frmItems.frx":6E40
            MousePointer    =   99  'Custom
            TabIndex        =   167
            Top             =   1560
            Width           =   390
         End
      End
      Begin VB.Frame FrmAct 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   7575
         Left            =   3840
         TabIndex        =   166
         Top             =   120
         Width           =   3855
         Begin VB.ListBox LstAction 
            Height          =   7200
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   180
            Top             =   240
            Width           =   3615
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps3"
      Height          =   7935
      Index           =   3
      Left            =   6240
      TabIndex        =   41
      Top             =   480
      Width           =   7815
      Begin VB.Frame FraCalc 
         Height          =   615
         Left            =   120
         TabIndex        =   200
         Top             =   7200
         Width           =   7575
         Begin VB.OptionButton OptHex 
            Caption         =   "ʮ����"
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   204
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptHex 
            Caption         =   "16����"
            Height          =   255
            Index           =   1
            Left            =   5280
            TabIndex        =   203
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptHex 
            Caption         =   "������"
            Height          =   255
            Index           =   2
            Left            =   6360
            TabIndex        =   202
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtCalc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   201
            Text            =   "txtCalc"
            Top             =   180
            Width           =   3855
         End
      End
      Begin MnBWarband_Editor.TriggersEditor TriggersEditor1 
         Height          =   7095
         Left            =   120
         TabIndex        =   199
         Top             =   120
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   12515
      End
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   8415
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   14843
      MultiRow        =   -1  'True
      ImageList       =   "IL1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������Ϣ(&I)"
            Key             =   "Info"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&P)"
            Key             =   "PropBag"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ģ��&����(&M)"
            Key             =   "ModelnAction"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������(&T)"
            Key             =   "Trigger"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps1"
      Height          =   7935
      Index           =   1
      Left            =   6240
      TabIndex        =   38
      Top             =   480
      Width           =   7815
      Begin VB.Frame Frame2 
         Caption         =   "Null"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   15
         Index           =   8
         Left            =   2880
         TabIndex        =   164
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Frame FrmFlags 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   4335
         Left            =   3480
         TabIndex        =   100
         Top             =   3480
         Width           =   4215
         Begin VB.ListBox LstFlags 
            Height          =   4050
            ItemData        =   "frmItems.frx":714A
            Left            =   120
            List            =   "frmItems.frx":7151
            Style           =   1  'Checkbox
            TabIndex        =   101
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ǰ׺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   7695
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   3255
         Begin VB.Frame Frame3 
            Caption         =   "Ԥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   3495
            Left            =   120
            TabIndex        =   67
            Top             =   4080
            Width           =   3015
            Begin VB.OptionButton OptIMod 
               Caption         =   "Ͷ��(��)"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   255
               Index           =   16
               Left            =   240
               TabIndex        =   85
               Top             =   3120
               Width           =   1335
            End
            Begin VB.OptionButton OptIMod 
               Caption         =   "Ͷ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   255
               Index           =   15
               Left            =   1680
               TabIndex        =   84
               Top             =   3120
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
               Caption         =   "��ʸ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   255
               Index           =   14
               Left            =   1680
               TabIndex        =   83
               Top             =   2760
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   13
               Left            =   240
               TabIndex        =   82
               Top             =   2760
               Width           =   1335
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   12
               Left            =   1680
               TabIndex        =   81
               Top             =   2400
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H0000C000&
               Height          =   255
               Index           =   11
               Left            =   240
               TabIndex        =   80
               Top             =   2400
               Width           =   1335
            End
            Begin VB.OptionButton OptIMod 
               Caption         =   "��(��Ʒ)"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Index           =   10
               Left            =   1680
               TabIndex        =   79
               Top             =   2040
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H00404040&
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   78
               Top             =   2040
               Width           =   1335
            End
            Begin VB.OptionButton OptIMod 
               Caption         =   "��ͷ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   8
               Left            =   1680
               TabIndex        =   77
               Top             =   1680
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   76
               Top             =   1680
               Width           =   1335
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H00404000&
               Height          =   255
               Index           =   6
               Left            =   1680
               TabIndex        =   75
               Top             =   1320
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H00808080&
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   74
               Top             =   1320
               Width           =   1335
            End
            Begin VB.OptionButton OptIMod 
               Caption         =   "�Զ���"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   17
               Left            =   1680
               TabIndex        =   73
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H00008080&
               Height          =   255
               Index           =   4
               Left            =   1680
               TabIndex        =   72
               Top             =   960
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
               Caption         =   "�·�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008080&
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   71
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton OptIMod 
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
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   70
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton OptIMod 
               Caption         =   "��ƥ"
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
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   69
               Top             =   600
               Width           =   1335
            End
            Begin VB.OptionButton OptIMod 
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
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   68
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.ListBox LstImodBits 
            Height          =   3630
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   42
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "��ҩ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   3135
         Index           =   4
         Left            =   3480
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox TxtADam 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   190
            Top             =   2295
            Width           =   1140
         End
         Begin VB.OptionButton OptADamR 
            Caption         =   "��"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   189
            Top             =   2370
            Width           =   1215
         End
         Begin VB.OptionButton OptADamR 
            Caption         =   "��"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   188
            Top             =   2580
            Width           =   1215
         End
         Begin VB.TextBox TxtAmMax 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   99
            Top             =   1620
            Width           =   1695
         End
         Begin VB.TextBox TxtAmWei 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   98
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox TxtAmLen 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   97
            Top             =   420
            Width           =   1695
         End
         Begin VB.OptionButton OptADamR 
            Caption         =   "��"
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   191
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�˺���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   960
            TabIndex        =   192
            Top             =   2340
            Width           =   675
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��ҩ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   96
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   960
            TabIndex        =   95
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "���ȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   960
            TabIndex        =   94
            Top             =   480
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2175
         Index           =   7
         Left            =   3480
         TabIndex        =   153
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox TxtHA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   158
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtBA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3000
            TabIndex        =   157
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtFA 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   156
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox TxtArDiff 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3000
            TabIndex        =   155
            Top             =   440
            Width           =   735
         End
         Begin VB.TextBox TxtArWei 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   154
            Top             =   440
            Width           =   735
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   163
            Top             =   1020
            Width           =   675
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ͷ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   162
            Top             =   1020
            Width           =   675
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ȷ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   161
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�Ѷȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   160
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   159
            Top             =   480
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "��ս��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   3135
         Index           =   1
         Left            =   3480
         TabIndex        =   45
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox TxtThrust 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   58
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox TxtSwing 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   57
            Top             =   1695
            Width           =   1215
         End
         Begin VB.TextBox TxtSpd 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   56
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox TxtDiff 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   54
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox TxtWeight 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   47
            Top             =   400
            Width           =   855
         End
         Begin VB.TextBox TxtLen 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   46
            Top             =   400
            Width           =   855
         End
         Begin VB.Frame FrmOptDT 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   735
            Left            =   2800
            TabIndex        =   59
            Top             =   1550
            Width           =   1335
            Begin VB.OptionButton OptDT 
               Caption         =   "����"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   61
               Top             =   220
               Width           =   1215
            End
            Begin VB.OptionButton OptDT 
               Caption         =   "����"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   60
               Top             =   430
               Width           =   1215
            End
            Begin VB.OptionButton OptDT 
               Caption         =   "����"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   62
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.Frame FrmOptDT2 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   735
            Left            =   2800
            TabIndex        =   63
            Top             =   2270
            Width           =   1335
            Begin VB.OptionButton OptDT2 
               Caption         =   "����"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   66
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton OptDT2 
               Caption         =   "����"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   64
               Top             =   430
               Width           =   1215
            End
            Begin VB.OptionButton OptDT2 
               Caption         =   "����"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   65
               Top             =   220
               Width           =   1215
            End
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   840
            TabIndex        =   196
            Top             =   2460
            Width           =   570
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ӿ�:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   840
            TabIndex        =   195
            Top             =   1755
            Width           =   570
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ٶ�:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   2400
            TabIndex        =   55
            Top             =   1035
            Width           =   570
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�Ѷ�:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   480
            TabIndex        =   53
            Top             =   1035
            Width           =   570
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   480
            TabIndex        =   49
            Top             =   480
            Width           =   570
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2400
            TabIndex        =   48
            Top             =   480
            Width           =   570
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2175
         Index           =   5
         Left            =   3480
         TabIndex        =   131
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox txtSDiff 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   197
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox TxtSWei 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   136
            Top             =   440
            Width           =   735
         End
         Begin VB.TextBox TxtSScale 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   135
            Top             =   440
            Width           =   735
         End
         Begin VB.TextBox TxtSSpdR 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   134
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox TxtArmor 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   133
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtSHP 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   132
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�Ѷȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   198
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   141
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ߴ磺"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   140
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ٶȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   139
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ǿ�ȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   138
            Top             =   1020
            Width           =   675
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   137
            Top             =   1020
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Զ����������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2655
         Index           =   2
         Left            =   3480
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox TxtRLen 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   145
            Top             =   440
            Width           =   735
         End
         Begin VB.TextBox TxtAcc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   143
            Top             =   1440
            Width           =   735
         End
         Begin VB.OptionButton OptDamR 
            Caption         =   "��"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   115
            Top             =   2250
            Width           =   735
         End
         Begin VB.OptionButton OptDamR 
            Caption         =   "��"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   114
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox TxtDamR 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   113
            Top             =   1965
            Width           =   540
         End
         Begin VB.TextBox TxtSSpd 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            TabIndex        =   112
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox TxtMax 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   111
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtSpdR 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            TabIndex        =   110
            Top             =   940
            Width           =   735
         End
         Begin VB.TextBox TxtRDiff 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            TabIndex        =   109
            Top             =   1965
            Width           =   735
         End
         Begin VB.TextBox TxtRWei 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            TabIndex        =   108
            Top             =   440
            Width           =   735
         End
         Begin VB.OptionButton OptDamR 
            Caption         =   "��"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   116
            Top             =   1830
            Width           =   735
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "���ȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2160
            TabIndex        =   144
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "���ȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2160
            TabIndex        =   142
            Top             =   1500
            Width           =   675
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2160
            TabIndex        =   107
            Top             =   1005
            Width           =   675
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�˺���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2160
            TabIndex        =   106
            Top             =   2010
            Width           =   675
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "���٣�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   105
            Top             =   1500
            Width           =   675
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ٶȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   104
            Top             =   1000
            Width           =   675
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�Ѷȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   103
            Top             =   2010
            Width           =   675
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   102
            Top             =   480
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "��ƥ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2775
         Index           =   3
         Left            =   3480
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox TxtSc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   130
            Top             =   2130
            Width           =   735
         End
         Begin VB.TextBox TxtHDiff 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   128
            Top             =   1530
            Width           =   735
         End
         Begin VB.TextBox TxtCharge 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   127
            Top             =   1530
            Width           =   735
         End
         Begin VB.TextBox TxtMV 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   126
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtHArmor 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   125
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtHSpd 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   124
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox TxtHP 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   123
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ߴ磺"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   129
            Top             =   2175
            Width           =   675
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�Ѷȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   122
            Top             =   1590
            Width           =   675
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   121
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��棺"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   120
            Top             =   1590
            Width           =   675
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ٿأ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   119
            Top             =   1020
            Width           =   675
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   118
            Top             =   1020
            Width           =   675
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ٶȣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2280
            TabIndex        =   117
            Top             =   480
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "����鱾����������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2295
         Index           =   6
         Left            =   3480
         TabIndex        =   146
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox TxtQty 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   149
            Top             =   1620
            Width           =   1335
         End
         Begin VB.TextBox TxtGWei 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   148
            Top             =   420
            Width           =   1335
         End
         Begin VB.TextBox TxtGMax 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   147
            Top             =   1020
            Width           =   1335
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ʳ��Ʒ�ʣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   152
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1080
            TabIndex        =   151
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1080
            TabIndex        =   150
            Top             =   1080
            Width           =   675
         End
      End
   End
   Begin MSComctlLib.TabStrip TabType 
      Height          =   8415
      Left            =   0
      TabIndex        =   205
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   14843
      MultiRow        =   -1  'True
      Placement       =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   13
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "all"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "t2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "˫��"
            Key             =   "t3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "t4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "t8t9"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ͷ��"
            Key             =   "t10"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "t12t13t14t15"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "t7"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��"
            Key             =   "t1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ʸ"
            Key             =   "t5t6t18"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "t16t17"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�鼮"
            Key             =   "t20"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "t19t11"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label LbTest 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1200
      TabIndex        =   44
      Top             =   9120
      Visible         =   0   'False
      Width           =   570
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
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   3
      Left            =   3480
      MouseIcon       =   "frmItems.frx":715F
      MousePointer    =   99  'Custom
      TabIndex        =   186
      Top             =   9120
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
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   2
      Left            =   2640
      MouseIcon       =   "frmItems.frx":7469
      MousePointer    =   99  'Custom
      TabIndex        =   185
      Top             =   9120
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
      Left            =   5160
      MouseIcon       =   "frmItems.frx":7773
      MousePointer    =   99  'Custom
      TabIndex        =   184
      Top             =   9120
      Width           =   705
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
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   1
      Left            =   4320
      MouseIcon       =   "frmItems.frx":7A7D
      MousePointer    =   99  'Custom
      TabIndex        =   183
      Top             =   9120
      Width           =   705
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   8
      Top             =   9120
      Width           =   690
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
      Left            =   200
      TabIndex        =   2
      Top             =   165
      Width           =   495
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean
Dim CustomActive As Boolean
Dim CalcText_De As String
Dim LstItemMPos As POINTAPI
Dim TrgCnt As New clsTriggersEditor

Private Sub InitItemsListView()
Dim n As Integer
n = 3

With LstItems
   .View = lvwReport
   .Sorted = False
   .ListItems.Clear
   .ColumnHeaders.Clear
   .SortOrder = lvwAscending
   .FullRowSelect = True
   .AllowColumnReorder = False
   .LabelEdit = lvwManual
   .Checkboxes = False
   .GridLines = True
   .MultiSelect = False
   .HideSelection = False

   .ColumnHeaders.Add , , PublicMsgs(13), LstItems.Width / n / 3.6
   .ColumnHeaders.Add , , PublicEditors(2) & PublicMsgs(14), LstItems.Width / n * 1.5
   .ColumnHeaders.Add , , PublicEditors(2) & "ID", LstItems.Width / n * 1.5
End With

End Sub

'*************************************************************************
'**�� �� ����LoadItemsList
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-11-26 23:16:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadItemsList()
Dim n As Long, oItem As ListItem, tI As Integer64b, H As Integer

LstItems.ListItems.Clear  '����б�

For n = 0 To UBound(itm)

  H = n - 1      '��һ��Ʒ
  If H < 0 Then H = 0

      Set oItem = LstItems.ListItems.Add(, "itm_" & CStr(n), itm(n).ID)
      
      With oItem
        
           If Not (HaveDoubleUsages(itm(H).itmType)) Or Not (IsWeapon(itm(H).itmType)) Or n = 0 Then
              .SubItems(1) = itm(n).csvName
           Else
              .SubItems(1) = PublicMsgs(28)
           End If
        
       .SubItems(2) = itm(n).dbName

      End With

Next n

End Sub

'*************************************************************************
'**�� �� ����loadItemsListbyType
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-03-07 15:31:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub loadItemsListbyType(TypeKey As String)
Dim n As Long, oItem As ListItem, tI As Integer64b, H As Integer, i As Integer, Continue As Boolean
Dim temArr() As String

LstItems.ListItems.Clear  '����б�

temArr() = Split(TypeKey, "t")

For n = 0 To UBound(itm)

  H = n - 1      '��һ��Ʒ
  If H < 0 Then H = 0
  Continue = False
  For i = 0 To UBound(temArr)
     If IsNumeric(temArr(i)) Then
        If GetItmType(itm(n).itmType) = CInt(temArr(i)) Then
           Continue = True
           Exit For
        End If
     End If
  Next i
  
  If Continue Then
      Set oItem = LstItems.ListItems.Add(, "itm_" & CStr(n), itm(n).ID)
      
      With oItem
        
           If Not (HaveDoubleUsages(itm(H).itmType)) Or Not (IsWeapon(itm(H).itmType)) Or n = 0 Then
              .SubItems(1) = itm(n).csvName
           Else
              .SubItems(1) = PublicMsgs(28)
           End If
        
       .SubItems(2) = itm(n).dbName

      End With
  End If
Next n

End Sub

'*************************************************************************
'**�� �� ����LoadModelList
'**��    �룺(Int)Index
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-10 13:08:40
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadModelList(itm As Type_Item)
Dim i As Integer

LstModel.Clear
With itm
      If .nmdl > 0 Then
         For i = 1 To .nmdl
             LstModel.AddItem .mdlname(i)
      
         Next i
      End If
End With

End Sub

'*************************************************************************
'**�� �� ����InitMLabel
'**��    �룺��
'**��    ������
'**��������������ģ�Ϳ��ư�ťλ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-02-14 22:51:39
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitMLabel()
LbMadd.Left = LbMdel.Left - LbMadd.Width - 100
End Sub

'*************************************************************************
'**�� �� ����InitDamChk
'**��    �룺��
'**��    ������
'**���������������˺�chkλ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-02-16 14:00:19
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitDamChk()
Label8(0).Left = TxtSwing.Left - Label8(0).Width
Label8(1).Left = TxtThrust.Left - Label8(1).Width
End Sub

'*************************************************************************
'**�� �� ����InitCBixmesh
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-10 22:26:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitCBixmesh()
CBixmesh.Clear
CBixmesh.AddItem "��ģ��"
CBixmesh.AddItem "��Ʒ��ģ��"          'ixmesh_Inventory = "1000000000000000"
CBixmesh.AddItem "����ģ��"            'ixmesh_Flying_Ammo = "2000000000000000"
CBixmesh.AddItem "Я��ģ��"            'ixmesh_Carry = "3000000000000000"
End Sub

'*************************************************************************
'**�� �� ����InitLstAction
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-11 09:57:32
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitLstAction()
Dim i As Integer
LstAction.Clear

For i = 0 To UBound(Itcf)
    LstAction.AddItem Itcf(i).Y
Next i

End Sub

'*************************************************************************
'**�� �� ����InitLstFlags
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-04 0:16:23
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitLstFlags()
Dim i As Integer

LstFlags.Clear
For i = 0 To UBound(Itp)
     LstFlags.AddItem Itp(i).Y
Next i

End Sub

Private Sub actPaste_Click()

End Sub

Private Sub CApply_Click()
Dim q As Boolean
  If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
  
    If UCase(itm(CurrentItmID).dbName) <> UCase(CurrentItm.dbName) Then             '����
        q = ChangeStrID(itm(CurrentItmID).dbName, CurrentItm.dbName)
        
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentItm.dbName), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    
    itm(CurrentItmID) = CurrentItm
    
    LstItems.SelectedItem.SubItems(1) = itm(CurrentItmID).csvName
    
    LstItems.SelectedItem.SubItems(2) = itm(CurrentItmID).dbName

    If TabType.SelectedItem.Index = 1 Then
     '���л�ģʽ���ж�
       If ChkNext.Value = 1 Then
              LstItems.ListItems(CurrentItmID + 2).SubItems(1) = PublicMsgs(28)
       ElseIf ChkNext.Value = 0 Then
          If CurrentItmID + 1 <= N_Item - 1 Then
             If itm(CurrentItmID + 1).csvName <> "" Then
                  LstItems.ListItems(CurrentItmID + 2).SubItems(1) = itm(CurrentItmID + 1).csvName
             Else
                  LstItems.ListItems(CurrentItmID + 2).SubItems(1) = itm(CurrentItmID + 1).disname
             End If
          End If
       End If
    End If
    'LoadItmInfo CInt(CurrentItmID)
  End If

End Sub

Private Sub CBixmesh_Click()
Dim Index As Integer
Index = LstModel.ListIndex + 1

If Index >= 1 Then

       If CBixmesh.ListIndex = 0 Then
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), ixmesh_Inventory_bit, 0)
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), ixmesh_Flying_Ammo_bit, 0)
       ElseIf CBixmesh.ListIndex = 1 Then
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), ixmesh_Inventory_bit, 1)
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), ixmesh_Flying_Ammo_bit, 0)
       ElseIf CBixmesh.ListIndex = 2 Then
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), ixmesh_Inventory_bit, 0)
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), ixmesh_Flying_Ammo_bit, 1)
       ElseIf CBixmesh.ListIndex = 3 Then
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), ixmesh_Inventory_bit, 1)
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), ixmesh_Flying_Ammo_bit, 1)
       End If
End If

End Sub

Private Sub CBixmesh_Scroll()
CBixmesh_Click
End Sub


Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If itm(CurrentItmID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), itm(CurrentItmID).dbName), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
              DelIndex itm(CurrentItmID).dbName
              
              If CurrentItmID < N_Item - 1 Then
                For i = CurrentItmID To N_Item - 2 Step 1
                    ChangeID itm(i + 1).dbName, itm(i + 1).ID - 1
                    j = itm(i).ID
                    itm(i) = itm(i + 1)
                    itm(i).ID = j
                    LstItems.ListItems(i + 1).SubItems(1) = LstItems.ListItems(i + 2).SubItems(1)
                    LstItems.ListItems(i + 1).SubItems(2) = LstItems.ListItems(i + 2).SubItems(2)
                Next i
                
                ReDim Preserve itm(N_Item - 2)
                LstItems.ListItems.Remove N_Item
                N_Item = N_Item - 1
                
              Else
                ReDim Preserve itm(N_Item - 2)
                LstItems.ListItems.Remove N_Item
                
                N_Item = N_Item - 1
                CurrentItmID = N_Item - 1
                
              End If
               
               LstItems_ItemClick LstItems.ListItems(CurrentItmID + 1)
               LstItems.ListItems(CurrentItmID + 1).Selected = True
               LstItems.ListItems(CurrentItmID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), itm(CurrentItmID).dbName), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), itm(CurrentItmID).dbName, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_Item - 1, itm(CurrentItmID).dbName & "_New") Then
           ReDim Preserve itm(N_Item)
           N_Item = N_Item + 1
           itm(N_Item - 1) = itm(N_Item - 2)
           itm(N_Item - 1).ID = itm(N_Item - 1).ID + 1
           itm(N_Item - 2) = itm(CurrentItmID)
           With itm(N_Item - 2)
                 .ID = N_Item - 2
                 .dbName = .dbName & "_New"
                 .disname = .disname & "_New"
                 .csvName = .csvName & "_New"
                 .csvName_pl = .csvName_pl & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstItems.ListItems.Add(, "itm_" & itm(N_Item - 1).ID, itm(N_Item - 1).ID)
           
                 With oItem
                    .SubItems(1) = itm(N_Item - 1).csvName
                    .SubItems(2) = itm(N_Item - 1).dbName
                 End With
                 
                 With LstItems.ListItems(LstItems.ListItems.Count - 1)
                    .SubItems(1) = itm(N_Item - 2).csvName
                    .SubItems(2) = itm(N_Item - 2).dbName
                 End With
                 
           LstItems_ItemClick LstItems.ListItems(N_Item - 1)
           LstItems.ListItems(N_Item - 1).Selected = True
           LstItems.ListItems(N_Item - 1).EnsureVisible
           
           Else
               MsgBox ActiveString(PublicMsgs(90), itm(CurrentItmID).dbName & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
         
      Case 2
         If CurrentItmID > 0 Then
           If itm(CurrentItmID - 1).Edit And itm(CurrentItmID).Edit Then
                SwapID itm(CurrentItmID - 1).dbName, itm(CurrentItmID).dbName
                SwapItems CurrentItmID - 1, CurrentItmID
                SwapListItem LstItems.ListItems(CurrentItmID), LstItems.ListItems(CurrentItmID + 1), 2, True
                
               LstItems_ItemClick LstItems.ListItems(CurrentItmID)
               LstItems.ListItems(CurrentItmID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), itm(CurrentItmID - 1).dbName), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If

         End If
      Case 3
        If CurrentItmID + 1 <= N_Item - 1 Then
           If itm(CurrentItmID).Edit And itm(CurrentItmID + 1).Edit Then
                SwapID itm(CurrentItmID).dbName, itm(CurrentItmID + 1).dbName
                SwapItems CurrentItmID, CurrentItmID + 1
                SwapListItem LstItems.ListItems(CurrentItmID + 1), LstItems.ListItems(CurrentItmID + 2), 2, True
                
                LstItems_ItemClick LstItems.ListItems(CurrentItmID + 2)
                LstItems.ListItems(CurrentItmID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), itm(CurrentItmID).dbName), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
        End If
End Select
End Sub

Private Sub ChkMer_Click()

     If ChkMer.Value = 1 Then
          CurrentItm.itmType = RepFlags(CurrentItm.itmType, itp_merchandise, 1)
     Else
          CurrentItm.itmType = RepFlags(CurrentItm.itmType, itp_merchandise, 0)
     End If


End Sub





Private Sub ChkNext_Click()

'If CurrentItmID >= 0 Then

   If ChkNext.Value = 1 Then
             CurrentItm.itmType = RepFlags(CurrentItm.itmType, itp_next_item_as_melee, 1)

   ElseIf ChkNext.Value = 0 Then
             CurrentItm.itmType = RepFlags(CurrentItm.itmType, itp_next_item_as_melee, 0)
   End If
   

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub CmdRefreshPY_Click()
TxtPY.Text = ExportItemPYCode(CurrentItm)
End Sub



Private Sub COutputLine_Click()
Dim t As String
t = ExportItemPYCode(CurrentItm, False)
frmLine.ShowTxtLine Me.Tag, -1
OutAsDebugTex t, PublicEditors_Simplified(2) & ":" & CurrentItm.dbName
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstItems, LstItems.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstItems_ItemClick(LstItems.SelectedItem)
End If
End Sub

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbExclamation, PublicMsgs(0)) = vbYes Then
LstItems_ItemClick LstItems.ListItems(CurrentItmID + 1)
LstItems.ListItems(CurrentItmID + 1).Selected = True
LstItems.ListItems(CurrentItmID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
'Inits
CustomActive = False

InitItemsListView
InitLstImodBits
initItmTypeOption
LoadItemsList
InitFrames
InitLstFlags
InitLstFaction
InitCBixmesh
InitLstMImod
InitLstAction

'InitLstItc

LstItems.ListItems(1).Selected = True
CurrentItmID = 0
CurrentItm = itm(0)
Call LoadItmInfo(CurrentItm)

TranslateForm Me
InitCMDs
InitMLabel
'InitDamChk

'InitTiggerEdit
TrgCnt.Initialize TriggersEditor1

Label2.Caption = Label2.Caption & N_Item

Label1(2).Caption = Replace(Label1(2).Caption, "(NOW):", "(" & UCase$(MnBInfo.Language) & "):", , , vbTextCompare)
Label1(3).Caption = Replace(Label1(3).Caption, "(NOW):", "(" & UCase$(MnBInfo.Language) & "):", , , vbTextCompare)

Loading = False
TVGotFocus = False
CustomActive = True
txtCalc.Text = ""
OptHex(0).Value = True

End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'frmInfo.Visible = False
'Unload frmInfo
End Sub

Private Sub LbMadd_Click()
If CustomActive Then
Dim i As Integer

With CurrentItm
     .nmdl = .nmdl + 1
     If .nmdl > 0 Then
          ReDim Preserve .mdlname(1 To .nmdl)
     Else
          ReDim .mdlname(1 To .nmdl)
     End If
     .mdlname(.nmdl) = "New_Model"
     ReDim Preserve .mdl_b(1 To .nmdl)
     .mdl_b(.nmdl) = "0"
End With

LoadModelList CurrentItm
LstModel.ListIndex = CurrentItm.nmdl - 1
TxtModName.Text = CurrentItm.mdlname(LstModel.ListIndex + 1)
LoadCBixmesh LstModel.ListIndex

For i = 0 To LstMIMod.ListCount - 1
    LstMIMod.Selected(i) = False
Next i

End If
End Sub

Private Sub LbMdel_Click()
If CustomActive Then
Dim Index As Integer, i As Integer
Index = LstModel.ListIndex + 1

With CurrentItm
    If Index >= 1 Then
      If .nmdl > 1 Then
              .nmdl = .nmdl - 1
         If Index = LstModel.ListCount Then              '���һ��
              
              If .nmdl = 0 Then                          'û��ģ��
                   ReDim .mdlname(0)
                   ReDim .mdl_b(0)
              Else
                   ReDim Preserve .mdlname(1 To .nmdl)
                   ReDim Preserve .mdl_b(1 To .nmdl)
              End If
         Else
              For i = Index To LstModel.ListCount - 1
                   .mdlname(i) = .mdlname(i + 1)
                   .mdl_b(i) = .mdl_b(i + 1)
              Next i
                   ReDim Preserve .mdlname(1 To .nmdl)
                   ReDim Preserve .mdl_b(1 To .nmdl)
         End If
      Else
          MsgBox PublicMsgs(60), vbOKOnly + vbCritical, PublicMsgs(0)
      End If
    End If
End With

LoadModelList CurrentItm
If CurrentItm.nmdl > 0 Then
    TxtModName.Text = CurrentItm.mdlname(1)
    LoadCBixmesh 0
    LoadLstMIMod 0
Else
    TxtModName.Text = ""
    CBixmesh.ListIndex = 0
    
    For i = 0 To LstMIMod.ListCount - 1
         LstMIMod.Selected(i) = False
    Next i
End If

End If
End Sub

Private Sub LbUp_Click()

End Sub



Private Sub LstAction_ItemCheck(Item As Integer)
Dim tI As Integer64b, i As Integer, Index As Integer

If CustomActive Then

tI = StrToI64(CurrentItm.Action)
If Item >= 0 And Item <= 32 Then
'��mask
'     If LstAction.Selected(Item) = True Then
'        tI = AddFlagsI64(tI, HexStrToI64(Itcf(Item).X))
'     ElseIf LstAction.Selected(Item) = False Then
'        tI = DeleteFlagsI64(tI, HexStrToI64(Itcf(Item).X))
'     End If
     
ElseIf Item >= 33 And Item <= 41 Then
'shoot
     If LstAction.Selected(Item) = True Then

        Index = Item
        For i = 33 To 41
             LstAction.Selected(i) = False
        Next i
             LstAction.Selected(Index) = True
        
   '     tI = DeleteFlagsI64(tI, HexStrToI64(itcf_Shoot_mask))
   '     tI = AddFlagsI64(tI, HexStrToI64(Itcf(Item).X))
     ElseIf LstAction.Selected(Item) = False Then
   '     tI = DeleteFlagsI64(tI, HexStrToI64(itcf_Shoot_mask))
     End If
ElseIf Item >= 42 And Item <= 64 Then
'carry
     If LstAction.Selected(Item) = True Then
        
        Index = Item
        For i = 42 To 64
             LstAction.Selected(i) = False
        Next i
             LstAction.Selected(Index) = True
        
   '     tI = DeleteFlagsI64(tI, HexStrToI64(itcf_Carry_mask))
   '     tI = AddFlagsI64(tI, HexStrToI64(Itcf(Item).X))
     ElseIf LstAction.Selected(Item) = False Then
   '     tI = DeleteFlagsI64(tI, HexStrToI64(itcf_Carry_mask))
     End If
ElseIf Item >= 65 And Item <= 66 Then
'reload
     If LstAction.Selected(Item) = True Then
        
        Index = Item
        For i = 65 To 66
             LstAction.Selected(i) = False
        Next i
             LstAction.Selected(Index) = True
        
     '   tI = DeleteFlagsI64(tI, HexStrToI64(itcf_Reload_mask))
     '   tI = AddFlagsI64(tI, HexStrToI64(Itcf(Item).X))
     ElseIf LstAction.Selected(Item) = False Then
     '   tI = DeleteFlagsI64(tI, HexStrToI64(itcf_Reload_mask))
     End If
End If

'CurrentItm.Action = I64toStrNZ(tI)
End If

End Sub

Private Sub LstAction_LostFocus()
Dim tI As Integer64b, i As Integer, Index As Integer
If CustomActive Then

CurrentItm.Action = ""

For i = 0 To UBound(Itcf)
     If LstAction.Selected(i) = True Then
        tI = AddFlagsI64(tI, HexStrToI64(Itcf(i).X))
     End If
Next i

CurrentItm.Action = I64toStrNZ(tI)

End If
End Sub

Private Sub LstFaction_LostFocus()
Dim i As Long, n As Long
If CustomActive Then
n = 0
For i = 0 To LstFaction.ListCount - 1
      If LstFaction.Selected(i) = True Then
          n = n + 1
      End If
Next i

If n = 0 Then
     ReDim CurrentItm.Faction(0)
     CurrentItm.FactionCount = 0
Else
     ReDim CurrentItm.Faction(1 To n)
     CurrentItm.FactionCount = n
     n = 1
     For i = 0 To LstFaction.ListCount - 1
         If LstFaction.Selected(i) = True Then
             CurrentItm.Faction(n).ID = i
             CurrentItm.Faction(n).strID = Factions(i).strID
             n = n + 1
         End If
     Next i
End If

End If
End Sub

Private Sub LstFlags_ItemCheck(Item As Integer)

If CustomActive Then
With LstFlags
       If Item = 17 And IsWeapon(CurrentItm.itmType) Then
           LstFlags.Selected(17) = HaveDoubleUsages(CurrentItm.itmType)
       ElseIf Item <= 26 Then
          If LstFlags.Selected(Item) Then
               CurrentItm.itmType = RepFlags(CurrentItm.itmType, Item + 12, 1)
          Else
               CurrentItm.itmType = RepFlags(CurrentItm.itmType, Item + 12, 0)
          End If
       Else
          If LstFlags.Selected(Item) Then
               CurrentItm.itmType = RepFlags(CurrentItm.itmType, Item + 17, 1)
          Else
               CurrentItm.itmType = RepFlags(CurrentItm.itmType, Item + 17, 0)
          End If
       End If
End With
End If

End Sub

Private Sub LstFlags_LostFocus()
'Dim i As Integer
'If CustomActive Then

'For i = 0 To LstFlags.ListCount - 1
'   If i <= 26 Then
'      If LstFlags.Selected(i) Then
'           CurrentItm.itmType = RepFlags(CurrentItm.itmType, i + 12, 1)
'      Else
'           CurrentItm.itmType = RepFlags(CurrentItm.itmType, i + 12, 0)
'      End If
'   Else
'      If LstFlags.Selected(i) Then
'           CurrentItm.itmType = RepFlags(CurrentItm.itmType, i + 17, 1)
'      Else
'           CurrentItm.itmType = RepFlags(CurrentItm.itmType, i + 17, 0)
'      End If
'   End If
'Next i

'End If
End Sub

Private Sub LstImodBits_ItemCheck(Item As Integer)

If CustomActive Then
   CustomActive = False
   OptIMod(OptIMod.UBound).Value = True
   If LstImodBits.Selected(Item) = True Then
        CurrentItm.Prefix = RepFlags(CurrentItm.Prefix, Item, 1)
   ElseIf LstImodBits.Selected(Item) = False Then
        CurrentItm.Prefix = RepFlags(CurrentItm.Prefix, Item, 0)
   End If
   LoadImodCombines CurrentItm.Prefix
   CustomActive = True
End If

End Sub

Private Sub LstImodBits_LostFocus()
Dim i As Integer
If CustomActive Then
'For i = 0 To LstImodBits.ListCount - 1
'   If LstImodBits.Selected(i) = True Then
'        CurrentItm.Prefix = RepFlags(CurrentItm.Prefix, i, 1)
'   ElseIf LstImodBits.Selected(i) = False Then
'        CurrentItm.Prefix = RepFlags(CurrentItm.Prefix, i, 0)
'   End If
'Next i

'LoadImodCombines CurrentItm.Prefix

End If
End Sub


Private Sub LstItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim n As Integer


CurrentItmID = Val(Item.Text)
CurrentItm = itm(CurrentItmID)

If CurrentItmID = N_Item - 1 Then
   MsgBox PublicMsgs(61), vbOKCancel + vbInformation, PublicMsgs(0)
     For n = 0 To FraProps.UBound
         FraProps(n).Enabled = False
     Next n
Else
     For n = 0 To FraProps.UBound
         FraProps(n).Enabled = True
     Next n
End If

If Tab1.TabIndex = 1 Then
    InitPropFrames
End If

LoadItmInfo CurrentItm

End Sub

Private Sub LstItems_KeyDown(KeyCode As Integer, Shift As Integer)

'If Shift = 2 Then
''     Timer_MousePos.Enabled = True
'End If

If Shift = 2 And KeyCode = vbKeyC Then
     frmInfo.CopyInfotoClipBoard
     frmTip.ShowTip PublicMsgs(130)
     Timer_KillTip.Enabled = True
End If

End Sub

Private Sub LstItems_KeyUp(KeyCode As Integer, Shift As Integer)

'If KeyCode = 17 Then
'    Unload frmInfo
'    Timer_MousePos.Enabled = False
'End If

End Sub

Private Sub LstItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not frmMain.mBanfrmInfo.Checked Then
'If Shift = 2 Then
   'LstItems.SetFocus
   LstItemMPos.X = X
   LstItemMPos.Y = Y
   Timer_MousePos.Enabled = True
   
   If Timer_KillTip.Enabled = True Then
      frmTip.HideTip
      Timer_KillTip.Enabled = False
   End If
'End If
End If

End Sub

Private Sub LstItems_Validate(Cancel As Boolean)
If Loading Then Cancel = True
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


Private Sub InitFrames()
Dim i As Integer

For i = 0 To 3
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
     Text1(2).Enabled = False
End If

End Sub


Private Sub LstMIMod_ItemCheck(Item As Integer)
Dim i As Integer, Index As Integer

If CustomActive Then
   Index = LstModel.ListIndex + 1

   If Index >= 1 Then
       If LstMIMod.Selected(Item) = True Then
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), Item, 1)
       ElseIf LstMIMod.Selected(Item) = False Then
           CurrentItm.mdl_b(Index) = RepFlags(CurrentItm.mdl_b(Index), Item, 0)
       End If
    End If
End If

End Sub

Private Sub LstModel_Click()
Dim Index As Integer
Index = LstModel.ListIndex

If Index >= 0 Then
   TxtModName.Text = CurrentItm.mdlname(Index + 1)

   LoadCBixmesh Index

   LoadLstMIMod Index

End If

End Sub

Private Sub LoadCBixmesh(Index As Integer)
Dim tI As Integer64b

tI = StrToI64(CurrentItm.mdl_b(Index + 1))
If ChkBit64b(tI, ixmesh_Inventory_bit) And ChkBit64b(tI, ixmesh_Flying_Ammo_bit) Then
      CBixmesh.ListIndex = 3
ElseIf ChkBit64b(tI, ixmesh_Inventory_bit) Then
      CBixmesh.ListIndex = 1
ElseIf ChkBit64b(tI, ixmesh_Flying_Ammo_bit) Then
      CBixmesh.ListIndex = 2
Else
      CBixmesh.ListIndex = 0
End If
End Sub

Private Sub LoadLstMIMod(Index As Integer)
Dim tI As Integer64b, i As Integer, H As Byte

tI = StrToI64(CurrentItm.mdl_b(Index + 1))

For i = 0 To LstMIMod.ListCount - 1
    LstMIMod.Selected(i) = False
Next i

For H = 0 To N_IMod - 1
    If ChkBit64b(tI, H) Then
          LstMIMod.Selected(H) = True
    End If
Next H

End Sub

Private Sub OptADamR_Click(Index As Integer)
CommoOptDTClick Index, TxtADam, CurrentItm.thrust_damage
End Sub

Private Sub OptAtt_Click(Index As Integer)
Dim tI As Integer64b
If CustomActive Then
 tI = StrToI64(CurrentItm.itmType)
 tI = DeleteFlagsI64(tI, HexStrToI64(itp_attachment_mask))
    Select Case Index
         Case 1
              tI = AddFlagsI64(tI, HexStrToI64(itp_force_attach_left_hand))
         Case 2
              tI = AddFlagsI64(tI, HexStrToI64(itp_force_attach_right_hand))
         Case 3
              tI = AddFlagsI64(tI, HexStrToI64(itp_force_attach_left_forearm))
         Case 4
              tI = AddFlagsI64(tI, HexStrToI64(itp_attach_armature))
    End Select
  CurrentItm.itmType = I64toStrNZ(tI)
End If
End Sub

Private Sub OptDamR_Click(Index As Integer)
CommoOptDTClick Index, TxtDamR, CurrentItm.thrust_damage
End Sub

Private Sub OptDT_Click(Index As Integer)
CommoOptDTClick Index, TxtSwing, CurrentItm.swing_damage
End Sub

Private Sub OptDT2_Click(Index As Integer)
CommoOptDTClick Index, TxtThrust, CurrentItm.thrust_damage
End Sub

Private Sub OptHex_Click(Index As Integer)
If CustomActive Then
       CustomActive = False
       Select Case Index
           Case 0
                txtCalc.Text = CalcText_De
           Case 1
                txtCalc.Text = RemoveUseless0(I64toHexStr(StrToI64(CalcText_De)))
           Case 2
                txtCalc.Text = RemoveUseless0(I64ToBinStr(StrToI64(CalcText_De)))
       End Select
       CustomActive = True
End If
End Sub


Private Sub OptIMod_Click(Index As Integer)
Dim tI As Integer64b, H As Byte

If CustomActive Then
    CustomActive = False
    If Index <= UBound(IModC) Then
      For i = 0 To LstImodBits.ListCount - 1
         LstImodBits.Selected(i) = False
      Next i

      tI = StrToI64(IModC(Index).X)
      
      CurrentItm.Prefix = "0"
      For H = 0 To N_IMod - 1
         If ChkBit64b(tI, H) Then
              LstImodBits.Selected(H) = True
              CurrentItm.Prefix = RepFlags(CurrentItm.Prefix, CInt(H), 1)
         End If
      Next H
    End If
    CustomActive = True
End If

End Sub

Private Sub OptIMod_LostFocus(Index As Integer)
'Call LstImodBits_LostFocus

'Dim i As Integer
'If CustomActive Then
'For i = 0 To LstImodBits.ListCount - 1
'   If LstImodBits.Selected(i) = True Then
'        CurrentItm.Prefix = RepFlags(CurrentItm.Prefix, i, 1)
'   ElseIf LstImodBits.Selected(i) = False Then
'        CurrentItm.Prefix = RepFlags(CurrentItm.Prefix, i, 0)
'   End If
'Next i

'LoadImodCombines CurrentItm.Prefix

'End If
End Sub

Private Sub Option1_Click(Index As Integer)
Dim tI1 As Integer64b, tI2 As Integer64b, tI3 As Integer64b, tI As Integer64b
If CustomActive Then

tI1 = StrToI64(CurrentItm.itmType)    'ԭ��ƷFlags
tI2 = StrToI64(CStr(itp_type_mask))   '������Ʒ����Flags
tI3 = StrToI64(CStr(Index))           'Ҫ��ӵ���Ʒ����Flags

tI = DeleteFlagsI64(tI1, tI2)         '���������Ʒ����Flags
tI = AddFlagsI64(tI, tI3)             '�����Ʒ����Flags
CurrentItm.itmType = I64toStrNZ(tI)

End If
End Sub

Private Sub ParamAdd_Click()

End Sub

Private Sub Tab1_BeforeClick(Cancel As Integer)
If Loading Then Cancel = 1
End Sub

Private Sub Tab1_Click()
Dim i As Integer, n As Integer

If CustomActive Then
For i = 0 To FraProps.UBound
    With FraProps(i)
         .Visible = i + 1 = Tab1.SelectedItem.Index
         If i = 1 Then
             InitPropFrames
         End If
    End With
Next i

LoadItmInfo CurrentItm
End If

End Sub

Private Sub InitLstImodBits()
Dim i As Integer, tStr As String

For i = 0 To N_IMod - 1
   If Trim(IMod(i).csvName) <> "" Then
      LstImodBits.AddItem IMod(i).csvName
   Else
      tStr = Replace(IMod(i).ID, "imod", "")
      tStr = Replace(tStr, "_", " ")
      tStr = Trim(tStr)
      LstImodBits.AddItem tStr
   End If
Next i

End Sub

Private Sub InitLstMImod()
Dim i As Integer, tStr As String

For i = 0 To N_IMod - 1
   If Trim(IMod(i).csvName) <> "" Then
      LstMIMod.AddItem IMod(i).csvName
   Else
      tStr = Replace(IMod(i).ID, "imod", "")
      tStr = Replace(tStr, "_", " ")
      tStr = Trim(tStr)
      LstMIMod.AddItem tStr
   End If
Next i

End Sub

'*************************************************************************
'**�� �� ����initItmTypeOption
'**��    �룺��
'**��    ������
'**������������ʼ����Ʒ����Option�ؼ�����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-11-28 13:42:15
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub initItmTypeOption()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim i As Integer
    
    For i = 1 To UBound(Item_Type)
          Option1(i).Caption = Item_Type(i).Y
    Next i

    Option1(0).Caption = "��"

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("frmItems", "initItmTypeList", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����FixItemFaction
'**��    �룺(Long)itemId
'**��    ������
'**����������������Ʒ��Ӫ������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-30 16:31:21
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub FixItemFaction(ItemID As Long)
Dim i As Integer, q As Boolean

With itm(ItemID)
     If .FactionCount > 0 Then
         For i = 1 To .FactionCount
             .Faction(i).ID = GetID(.Faction(i).strID, True, "", -1)
             If .Faction(i).ID = -1 Then q = True
         Next i
         
         If q Then StructureItemFactions ItemID
     End If
End With

End Sub
'*************************************************************************
'**�� �� ����LoadItmInfo
'**��    �룺��
'**��    ������
'**����������������Ʒ��Ϣ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-11-30 11:14:25
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadItmInfo(Item As Type_Item)
Dim i As Integer, tStr As String, tI As Integer64b, tpDam1 As Integer, Dam1 As Long, tpDam2 As Integer, Dam2 As Long, H As Byte, q As Boolean

    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Loading = True
CustomActive = False
For i = 0 To Text1.UBound
     Text1(i).Text = ""
Next i

Select Case Tab1.SelectedItem.Index

Case 1
'��ʾ��Ϣ
Text1(0).Text = Item.dbName
Text1(1).Text = Item.disname
Text1(2).Text = Item.csvName
Text1(3).Text = Item.csvName_pl
'��Ʒ�����ж�

Option1(GetItmType(Item.itmType)).Value = True

'�Ƿ���л��ж�
If Item.ID <= 0 Then
    If Not (IsWeapon(Item.itmType)) Or Not (IsWeapon(itm(Item.ID + 1).itmType)) Or HaveDoubleUsages(itm(Item.ID + 1).itmType) Then
        ChkNext.Enabled = False
    Else
        ChkNext.Enabled = True
    End If
Else
    If Item.ID >= N_Item - 1 Then
           If Not (IsWeapon(Item.itmType)) Or HaveDoubleUsages(itm(Item.ID - 1).itmType) Then
               ChkNext.Enabled = False
           Else
               ChkNext.Enabled = True
           End If
    Else
           If Not (IsWeapon(Item.itmType)) Or Not (IsWeapon(itm(Item.ID + 1).itmType)) Or HaveDoubleUsages(itm(Item.ID + 1).itmType) Or HaveDoubleUsages(itm(Item.ID - 1).itmType) Then
               ChkNext.Enabled = False
           Else
               ChkNext.Enabled = True
           End If
    End If
End If

'�л���ս�ж�

tI = StrToI64(Item.itmType)
If HaveDoubleUsages(Item.itmType) And IsWeapon(Item.itmType) = True Then
   ChkNext.Value = 1
Else
   ChkNext.Value = 0
End If

'��Ӫ�ж�
'FixItemFaction Item.ID
If Item.FactionCount > 0 Then
     For i = 1 To UBound(Item.Faction)
        Item.Faction(i).ID = GetID(Item.Faction(i).strID, True, "", -1)
        
        If Item.Faction(i).ID > -1 Then
          LstFaction.Selected(Item.Faction(i).ID) = True
          q = True
        End If
     Next i
     
     If q Then StructureItemFactions -1
Else
  For i = 0 To LstFaction.ListCount - 1
       LstFaction.Selected(i) = False
  Next i
End If

'�۸�
TxtPrice.Text = Item.price

'��ԣ��
TxtAb.Text = Item.abundance

Case 2

'ǰ׺�ж�
tI = StrToI64(Item.Prefix)

For i = 0 To LstImodBits.ListCount - 1
    LstImodBits.Selected(i) = False
Next i

For H = 0 To N_IMod - 1
    If ChkBit64b(tI, H) Then
          LstImodBits.Selected(H) = True
    End If
Next H

LoadImodCombines Item.Prefix

'��������
LoadLstFlags Item.itmType

If IsMeleeWeapon(Item.itmType) Then
    '����������Ϣ
    TxtWeight.Text = Format(Item.weight, "0.00")
    TxtLen.Text = Item.weapon_length
    TxtDiff.Text = Item.difficulty
    TxtSpd.Text = Item.speed_rating
    '�����˺�
    Dam1 = GetDamage(Item.swing_damage, tpDam1)
    Dam2 = GetDamage(Item.thrust_damage, tpDam2)
    OptDT(tpDam1).Value = True
    TxtSwing.Text = Dam1
    OptDT2(tpDam2).Value = True
    TxtThrust.Text = Dam2
ElseIf IsRangedWeapon(Item.itmType) Or IsFireArm(Item.itmType) Then
    TxtRWei.Text = Format(Item.weight, "0.00")
    TxtMax.Text = Item.max_ammo
    TxtSpdR.Text = Item.speed_rating
    TxtSSpd.Text = Item.missile_speed
    TxtRDiff.Text = Item.difficulty
    TxtAcc.Text = Item.leg_armor
    
    If IsFireArm(Item.itmType) Then
         TxtRLen.Text = ""
         TxtRLen.Enabled = False
         Label32.ForeColor = &H80000011
    Else
         TxtRLen.Enabled = True
         TxtRLen.Text = Item.weapon_length
         Label32.ForeColor = &H80000008
    End If
    
    Dam1 = GetDamage(Item.thrust_damage, tpDam1)
    OptDamR(tpDam1).Value = True
    TxtDamR.Text = Dam1
ElseIf IsAmmo(Item.itmType) Or IsFireArm(Item.itmType) Then
    TxtAmLen.Text = Item.weapon_length
    TxtAmWei.Text = Format(Item.weight, "0.00")
    TxtAmMax.Text = Item.max_ammo
    Dam1 = GetDamage(Item.thrust_damage, tpDam1)
    OptADamR(tpDam1).Value = True
    TxtADam.Text = Dam1
ElseIf IsHorse(Item.itmType) Then
    TxtHP.Text = Item.hit_points
    TxtHSpd.Text = Item.missile_speed
    TxtMV.Text = Item.speed_rating
    TxtHArmor.Text = Item.body_armor
    TxtCharge.Text = Item.thrust_damage
    TxtHDiff.Text = Item.difficulty
    TxtSc.Text = Item.weapon_length
ElseIf IsShield(Item.itmType) Then
    TxtSWei.Text = Format(Item.weight, "0.00")
    TxtSSpdR.Text = Item.speed_rating
    TxtSHP.Text = Item.hit_points
    TxtSScale.Text = Item.weapon_length
    TxtArmor.Text = Item.body_armor
    txtSDiff.Text = Item.difficulty
ElseIf IsArmor(Item.itmType) Then
    TxtArWei.Text = Format(Item.weight, "0.00")
    TxtArDiff.Text = Item.difficulty
    TxtHA.Text = Item.head_armor
    TxtBA.Text = Item.body_armor
    TxtFA.Text = Item.leg_armor
ElseIf IsGood(Item.itmType) Or IsAnimal(Item.itmType) Or IsBook(Item.itmType) Then
    TxtGWei.Text = Format(Item.weight, "0.00")
    TxtGMax.Text = Item.max_ammo
    
    If IsFood(Item.itmType) Then
           'TxtGMax.Enabled = True
           TxtQty.Enabled = True
           'Label33.ForeColor = &H80000012
           Label35.ForeColor = &H80000012
           'TxtGMax.Text = Item.max_ammo
           TxtQty.Text = Item.head_armor
    Else
           'TxtGMax.Enabled = False
           TxtQty.Enabled = False
           'Label33.ForeColor = &H80000011
           Label35.ForeColor = &H80000011
           'TxtGMax.Text = ""
           TxtQty.Text = ""
    End If
End If

Case 3
'ģ��
LoadModelList Item
If Item.nmdl > 0 Then
    LstModel.ListIndex = 0
    LoadCBixmesh LstModel.ListIndex
    LoadLstMIMod LstModel.ListIndex
End If

'��
OptAtt(GetAttachment(Item.itmType)).Value = True

If IsWeapon(Item.itmType) Or IsShield(Item.itmType) Then
   For H = 0 To 3
     OptAtt(H).Enabled = True
   Next H
   OptAtt(4).Enabled = False
ElseIf IsArmor(Item.itmType) Then
   For H = 1 To 3
     OptAtt(H).Enabled = False
   Next H
   OptAtt(4).Enabled = True
   OptAtt(0).Enabled = True
Else
   For H = 0 To OptAtt.UBound
     OptAtt(H).Enabled = False
   Next H

End If

'����

If IsWeapon(Item.itmType) Or IsAmmo(Item.itmType) Then
     LoadLstAction Item.Action
     LstAction.Enabled = True

Else
     For i = 0 To LstAction.ListCount - 1
          LstAction.Selected(i) = False
     Next i
     LstAction.Enabled = False
     'LstItc.Enabled = False

End If

Case 4
'������
TrgCnt.InputTrg Item.dbName, "itm", Item.Trigger()

End Select

CustomActive = True
Loading = False
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("frmItem", "LoadItemInfo", Err.Number, Err.Description)
    CustomActive = True
    Loading = False
End Sub


'*************************************************************************
'**�� �� ����InitLstFaction
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-07 13:27:55
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub InitLstFaction()
Dim n As Long

LstFaction.Clear
For n = 0 To UBound(Factions)
      LstFaction.AddItem Factions(n).csvName
Next n

End Sub

Private Sub TabType_Click()
Dim Idx As Integer
Idx = TabType.SelectedItem.Index
If Idx = 1 Then
   LoadItemsList
   enableCMDs True
Else
   loadItemsListbyType (TabType.Tabs(Idx).Key)
   enableCMDs False
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If CustomActive Then

If Index < 2 Then
  Text1(Index).Text = Replace(Text1(Index).Text, " ", "_")
End If

Select Case Index
    Case 0
        CurrentItm.dbName = Text1(Index).Text
    Case 1
        CurrentItm.disname = Text1(Index).Text
    Case 2
        CurrentItm.csvName = Text1(Index).Text
    Case 3
        CurrentItm.csvName_pl = Text1(Index).Text
End Select

End If
End Sub


'*************************************************************************
'**�� �� ����LoadLstFlags
'**��    �룺-(String)itmType
'**��    ����-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-05 17:09:54
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadLstFlags(itmType As String)
Dim tI As Integer64b, i As Byte
         tI = StrToI64(itmType)
    With LstFlags
        For i = 0 To LstFlags.ListCount - 1
           If i <= 26 Then
              LstFlags.Selected(i) = ChkBit64b(tI, i + 12)
           Else
              LstFlags.Selected(i) = ChkBit64b(tI, i + 17)
           End If
        Next i
    End With
End Sub

Private Sub Timer_KillTip_Timer()
    frmTip.HideTip
    Timer_KillTip.Enabled = False
End Sub

Private Sub Timer_MousePos_Timer()
Dim ItemIndex As Long, ReturnValue As Long, MouseOut As Boolean
Dim MPos As POINTAPI
Static PosOld As POINTAPI
Dim LVRect As RECT
'Dim Hwnd As Long

ReturnValue = GetCursorPos(MPos)   '��ȡ������λ��
ReturnValue = GetWindowRect(LstItems.hWnd, LVRect)   '��ȡlistview����
MouseOut = MPos.X < LVRect.Left Or MPos.X > LVRect.Right - 300 / Screen.TwipsPerPixelX Or MPos.Y < LVRect.Top + 250 / Screen.TwipsPerPixelX Or MPos.Y > LVRect.Bottom - 300 / Screen.TwipsPerPixelY

If MouseOut Then
     UnLoad frmInfo
     PosOld.X = 0
     PosOld.Y = 0
     Timer_MousePos.Enabled = False
Else
  If LstItemMPos.X <> PosOld.X Or LstItemMPos.Y <> PosOld.Y Then
  '------------------------���listview�����ָλ��ItemIndex----------------------------------
   'Hwnd = GetActiveWindow()
   'If Hwnd = frmMain.Hwnd Or Hwnd = frmItems.Hwnd Or Hwnd = frmInfo.Hwnd Then
       'LstItems.SetFocus
   'End If
       
   ItemIndex = GetListViewItemIndexUnderMousePointer(LstItems, LstItemMPos.X, LstItemMPos.Y)
   If ItemIndex > 0 Then
      frmInfo.ShowfrmInfo MPos.X * Screen.TwipsPerPixelX + 300, MPos.Y * Screen.TwipsPerPixelY + 300
      frmInfo.LoadItemInfo Val(LstItems.ListItems(ItemIndex).Text)
      frmInfo.ZOrder
      LstItems.SetFocus
   Else
      UnLoad frmInfo
   End If
  '-------------------------------------------------------------------------------------------
   PosOld = LstItemMPos
  End If
End If

End Sub

Private Sub trgPaste_Click()

End Sub

Private Sub TVTrigger_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

'*************************************************************************
'**�� �� ����GettiOnStr
'**��    �룺-(Double)ti_On
'**��    ����-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-13 21:36:37
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Function GettiOnStr(ti_On As Double) As String
Dim i As Integer

For i = 0 To UBound(tiOn)
    If CDbl(tiOn(i).X) = ti_On Then
       GettiOnStr = tiOn(i).Y
       Exit For
    End If
Next i

End Function

'*************************************************************************
'**�� �� ����LoadLstAction
'**��    �룺-(String)Action
'**��    ����-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-11 11:07:31
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub LoadLstAction(Action As String)
Dim tI As Integer64b, i As Integer, tI2 As Integer64b
tI = StrToI64(Action)

For i = 0 To LstAction.ListCount - 1
    LstAction.Selected(i) = False
Next i

'��mask
For i = 0 To 32
    If IsZero64b(And64b(tI, HexStrToI64(Itcf(i).X))) = False Then
          LstAction.Selected(i) = True
    End If
Next i

'shoot
tI2 = And64b(tI, HexStrToI64(itcf_Shoot_mask))
For i = 32 To 41
    If I64ToBinStr(tI2) = HexToBin(Itcf(i).X) Then
          LstAction.Selected(i) = True
    End If
Next i

'carry
tI2 = And64b(tI, HexStrToI64(itcf_Carry_mask))
For i = 42 To 64
    If I64ToBinStr(tI2) = HexToBin(Itcf(i).X) Then
          LstAction.Selected(i) = True
    End If
Next i

'reload
tI2 = And64b(tI, HexStrToI64(itcf_Reload_mask))
For i = 64 To 66
    If I64ToBinStr(tI2) = HexToBin(Itcf(i).X) Then
          LstAction.Selected(i) = True
    End If
Next i

End Sub

Private Sub LoadImodCombines(Prefix As String)
Dim i As Integer, tB As Boolean

tB = CustomActive
CustomActive = False
OptIMod(17).Value = True

For i = 0 To UBound(IModC)

If Prefix = IModC(i).X Then
      OptIMod(i).Value = True
End If

Next i
CustomActive = tB

End Sub

Private Sub Timer1_Timer()
Dim D As Long, DT As Integer, tI As Integer64b, i As Integer, Index As Integer

LbTest.Caption = Timer_MousePos.Enabled

End Sub

Private Sub TriggersEditor1_Validate(Cancel As Boolean)
TrgCnt.OutputTrg CurrentItm.Trigger()
'MsgBox LBound(CurrentItm.Trigger)
If LBound(CurrentItm.Trigger) = 0 Then
  CurrentItm.TriggerCount = 0
Else
  CurrentItm.TriggerCount = UBound(CurrentItm.Trigger)
End If
End Sub

Private Sub TxtAb_Change()
CommonTextChange TxtAb, CurrentItm.abundance, 255
End Sub

Private Sub TxtAcc_Change()
CommonTextChange TxtAcc, CurrentItm.leg_armor
End Sub

Private Sub TxtADam_Change()
CommonDamTextChange TxtADam, OptADamR, CurrentItm.thrust_damage
End Sub

Private Sub TxtAmLen_Change()
CommonTextChange TxtAmLen, CurrentItm.weapon_length
End Sub

Private Sub TxtAmMax_Change()
CommonTextChange TxtAmMax, CurrentItm.max_ammo
End Sub

Private Sub TxtAmWei_Change()
CommonTextChange TxtAmWei, CurrentItm.weight
End Sub

Private Sub TxtArDiff_Change()
CommonTextChange TxtArDiff, CurrentItm.difficulty
End Sub

Private Sub TxtArmor_Change()
CommonTextChange TxtArmor, CurrentItm.body_armor
End Sub

Private Sub TxtArWei_Change()
CommonTextChange TxtArWei, CurrentItm.weight
End Sub

Private Sub TxtBA_Change()
CommonTextChange TxtBA, CurrentItm.body_armor
End Sub

Private Sub txtCalc_Change()

If CustomActive Then
If OptHex(0).Value = True Then
      CalcText_De = I64toStrNZ(StrToI64(txtCalc.Text))
ElseIf OptHex(1).Value = True Then
      CalcText_De = I64toStrNZ(HexStrToI64(txtCalc.Text))
ElseIf OptHex(2).Value = True Then
      CalcText_De = I64toStrNZ(HexStrToI64(BinToHex(txtCalc.Text)))
End If
End If

End Sub

Private Sub TxtCharge_Change()
CommonTextChange TxtCharge, CurrentItm.thrust_damage
End Sub

Private Sub TxtDamR_Change()
CommonDamTextChange TxtDamR, OptDamR, CurrentItm.thrust_damage
End Sub

Private Sub TxtDiff_Change()
CommonTextChange TxtDiff, CurrentItm.difficulty
End Sub

Private Sub TxtFA_Change()
CommonTextChange TxtFA, CurrentItm.leg_armor
End Sub

Private Sub TxtGMax_Change()
CommonTextChange TxtGMax, CurrentItm.max_ammo
End Sub

Private Sub TxtGWei_Change()
CommonTextChange TxtGWei, CurrentItm.weight
End Sub

Private Sub TxtHA_Change()
CommonTextChange TxtHA, CurrentItm.head_armor
End Sub

Private Sub TxtHArmor_Change()
CommonTextChange TxtHArmor, CurrentItm.body_armor
End Sub

Private Sub TxtHDiff_Change()
CommonTextChange TxtHDiff, CurrentItm.difficulty
End Sub

Private Sub TxtHP_Change()
CommonTextChange TxtHP, CurrentItm.hit_points
End Sub

Private Sub TxtHSpd_Change()
CommonTextChange TxtHSpd, CurrentItm.missile_speed
End Sub

Private Sub TxtLen_Change()
CommonTextChange TxtLen, CurrentItm.weapon_length
End Sub

Private Sub txtMax_Change()
CommonTextChange TxtMax, CurrentItm.max_ammo
End Sub

Private Sub TxtModName_Change()
Dim Index As Integer
If CustomActive Then

Index = LstModel.ListIndex + 1
If Index >= 1 Then
    LstModel.List(LstModel.ListIndex) = TxtModName.Text
End If

End If
End Sub

Private Sub TxtModName_Validate(Cancel As Boolean)
Dim Index As Integer
If CustomActive Then

   CustomActive = False
    TxtModName.Text = Replace(TxtModName.Text, " ", "_")
   CustomActive = True
   
    Index = LstModel.ListIndex + 1
    
    If Index >= 1 Then
      CurrentItm.mdlname(Index) = TxtModName.Text
      LstModel.List(LstModel.ListIndex) = TxtModName.Text
    End If
End If
End Sub

Private Sub TxtMV_Change()
CommonTextChange TxtMV, CurrentItm.speed_rating
End Sub

Private Sub TxtPrice_Change()
CommonTextChange TxtPrice, CurrentItm.price
End Sub

Private Sub TxtQty_Change()
CommonTextChange TxtQty, CurrentItm.head_armor
End Sub

Private Sub TxtRDiff_Change()
CommonTextChange TxtRDiff, CurrentItm.difficulty
End Sub

Private Sub TxtRLen_Change()
CommonTextChange TxtRLen, CurrentItm.weapon_length
End Sub

Private Sub TxtRWei_Change()
CommonTextChange TxtRWei, CurrentItm.weight
End Sub

Private Sub TxtSc_Change()
CommonTextChange TxtSc, CurrentItm.weapon_length
End Sub

Private Sub txtSDiff_Change()
CommonTextChange txtSDiff, CurrentItm.difficulty
End Sub

Private Sub TxtSHP_Change()
CommonTextChange TxtSHP, CurrentItm.hit_points
End Sub

Private Sub TxtSpd_Change()
CommonTextChange TxtSpd, CurrentItm.speed_rating
End Sub

Private Sub TxtSpdR_Change()
CommonTextChange TxtSpdR, CurrentItm.speed_rating
End Sub

Private Sub TxtSScale_Change()
CommonTextChange TxtSScale, CurrentItm.weapon_length
End Sub

Private Sub TxtSSpd_Change()
CommonTextChange TxtSSpd, CurrentItm.missile_speed
End Sub

Private Sub TxtSSpdR_Change()
CommonTextChange TxtSSpdR, CurrentItm.speed_rating
End Sub

Private Sub TxtSWei_Change()
CommonTextChange TxtSWei, CurrentItm.weight
End Sub

Private Sub TxtSwing_Change()
CommonDamTextChange TxtSwing, OptDT, CurrentItm.swing_damage
End Sub

Private Sub TxtThrust_Change()
CommonDamTextChange TxtThrust, OptDT2, CurrentItm.thrust_damage
End Sub

Private Sub TxtWeight_Change()
CommonTextChange TxtWeight, CurrentItm.weight
End Sub

Private Sub AttachFrames(ShowFrame As Integer)

With FrmFlags
      .Top = Frame2(ShowFrame).Top + Frame2(ShowFrame).Height + 100
      .Height = Frame2(0).Height - .Top + Frame2(0).Top
      LstFlags.Height = .Height - 435
End With

End Sub

Private Sub InitPropFrames()

      For n = 1 To Frame2.UBound
          Frame2(n).Visible = False
      Next n
             
      If IsMeleeWeapon(CurrentItm.itmType) Then
          Frame2(1).Visible = True
          AttachFrames 1
      ElseIf IsRangedWeapon(CurrentItm.itmType) Or IsFireArm(CurrentItm.itmType) Then
          Frame2(2).Visible = True
          AttachFrames 2
      ElseIf IsHorse(CurrentItm.itmType) Then
          Frame2(3).Visible = True
          AttachFrames 3
      ElseIf IsAmmo(CurrentItm.itmType) Then
          Frame2(4).Visible = True
          AttachFrames 4
      ElseIf IsShield(CurrentItm.itmType) Then
          Frame2(5).Visible = True
          AttachFrames 5
      ElseIf IsGood(CurrentItm.itmType) Or IsBook(CurrentItm.itmType) Or IsAnimal(CurrentItm.itmType) Then
          Frame2(6).Visible = True
          AttachFrames 6
      ElseIf IsArmor(CurrentItm.itmType) Then
          Frame2(7).Visible = True
          AttachFrames 7
      Else
          AttachFrames 8
      End If

End Sub

Private Sub CommonTextChange(TextBox As TextBox, Destination As Variant, Optional Maxium As Long = 10 ^ 9, Optional Minium As Long = -10 ^ 9)
Dim t As Double
If CustomActive Then
   CustomActive = False
     t = Format(Val(TextBox.Text), "0.00")
     If t < Minium Then
        t = Minium
        TextBox.Text = CStr(t)
     ElseIf t > Maxium Then
        t = Maxium
        TextBox.Text = CStr(t)
     End If
     Destination = t
   CustomActive = True
End If

End Sub

Private Sub CommonDamTextChange(TextBox As TextBox, Opt As Object, Destination As Long)
Dim i As Integer, DT As Integer, t As Double, tB As Boolean
If CustomActive Then
  CustomActive = False
    For i = 0 To Opt.UBound
        If Opt(i).Value = True Then
           DT = i
        End If
    Next i
    
    t = Val(TextBox.Text)
    If t < 0 Then
       t = -t
       TextBox.Text = CStr(t)
    End If
    
    If DT = 0 Then
      If t >= 2 ^ 7 Then t = 2 ^ 7 - 1
      TextBox.Text = CStr(t)
    ElseIf DT = 1 Then
      If t >= (2 ^ 8 - 2 ^ 7) Then t = 2 ^ 8 - 2 ^ 7 - 1
      TextBox.Text = CStr(t)
    ElseIf DT = 2 Then
      If t >= 2 ^ 9 - 2 ^ 8 Then t = 2 ^ 9 - 2 ^ 8 - 1
      TextBox.Text = CStr(t)
    End If
    
    Destination = ExDamage(CLng(t), DT)
  CustomActive = True
End If
End Sub

Private Sub CommoOptDTClick(Index As Integer, TextBox As TextBox, Destination As Long)
Dim t As Long
If CustomActive Then
   t = CLng(Val(TextBox.Text))
   If Index < 2 Then
      If t >= 2 ^ 7 Then t = 2 ^ 7 - 1
      TextBox.Text = t
   End If
   Destination = ExDamage(t, Index)
End If
End Sub

'*************************************************************************
'**�� �� ����enableCMDs
'**��    �룺-(Boolean)enable
'**��    ����-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-03-07 15:53:35
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Private Sub enableCMDs(enable As Boolean)
Dim i As Integer

For i = cCMD.LBound To cCMD.UBound
   cCMD(i).Enabled = enable
Next i

End Sub

Public Sub ReLoadInfo()
LoadItmInfo CurrentItm
End Sub

'####################################### Reserved for future#####################################
'Private Sub InitLstItc()
'Dim i As Integer

'LstItc.Clear
'For i = 0 To UBound(Itc)
'      LstItc.AddItem Itc(i).Y
'Next i

'End Sub

'Private Sub LoadLstItc(Index As Integer)
'Dim i As Integer, tI As Integer64b, tI2 As Integer64b

'tI = StrToI64(itm(Index).Action)

'For i = 0 To UBound(Itc)
'     LstItc.Selected(i) = False
'Next i

'For i = 0 To UBound(Itc)
'    tI2 = And64b(tI, StrToI64(Itc(i).X))
'    If I64ToBinStr(tI2) = I64ToBinStr(StrToI64(Itc(i).X)) Then
'           LstItc.Selected(i) = True
'    End If
'Next i

'End Sub



