VERSION 5.00
Begin VB.Form BookSelectView 
   Caption         =   "�f�~�ήѥ���ƪ��d�ߩM�]�w�ﶵ"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8745
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command1 
      Caption         =   "���y�ƥ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��^�D���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton Book3 
      Caption         =   "�f�~"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton Book 
      Caption         =   "���y"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Book2 
      Caption         =   "���y�`��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "���I��U�C�ݭn�n�d�ߤ��ܧ󪺬������:"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   8775
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   600
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "BookSelectView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Book_Click()
PBView.Refresh
PBView.Show
PBView.PBList.Tab = 0
BookSelectView.Hide
End Sub

Private Sub Book2_Click()
PBView.Refresh
PBView.Show
PBView.PBList.Tab = 1
BookSelectView.Hide
End Sub

Private Sub Book3_Click()
PBView.Refresh
PBView.Show
PBView.PBList.Tab = 2
BookSelectView.Hide
End Sub

Private Sub Command1_Click()
PBView.Refresh
PBView.Show
PBView.PBList.Tab = 3
BookSelectView.Hide
End Sub

Private Sub Command4_Click()
PBView.Refresh
StaffUse.Show
BookSelectView.Hide
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\PIC\MEMBER.jpg")
Image1.Move 0, 0, BookSelectView.ScaleWidth, BookSelectView.ScaleHeight
End Sub
