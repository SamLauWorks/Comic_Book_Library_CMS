VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Top10ForMember 
   Caption         =   "�Q�j�̨��w��H��]"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   12495
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command9 
      Caption         =   "�Q�j�H��f�~�ֿn���ɦ���"
      Height          =   855
      Left            =   9600
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�Q�j�H��@�~�ֿn���ɦ���"
      Height          =   855
      Left            =   11160
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�Q�j�H��@�~�ֿn�P��q"
      Height          =   855
      Left            =   9600
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�Q�j�H��ѥ��ֿn���ɦ���"
      Height          =   855
      Left            =   9600
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�Q�j�H��ѥ��ֿn�P��q"
      Height          =   855
      Left            =   9600
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�C��Q�j�H��f�~�P�⦸�Ʋέp"
      Height          =   855
      Left            =   11160
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�C��Q�j�H��ѥ����ɦ��Ʋέp"
      Height          =   855
      Left            =   11160
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�C��Q�j�H��ѥ��P��q�έp"
      Height          =   855
      Left            =   11160
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Exit 
      Caption         =   "���}"
      Height          =   975
      Left            =   4920
      TabIndex        =   1
      Top             =   6480
      Width           =   3135
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8493
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "�w��Ө�Q�j�̨��w��H��]!�п�ܥk�誺���s,�Ω���ܫȤ�Q�n�s������T:"
      Height          =   1215
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "Top10ForMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Click()
Menu.Show
Top10ForMember.Hide
End Sub

