VERSION 5.00
Begin VB.Form OBOSlecteView 
   Caption         =   "OBOSlecteView"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7320
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton B2 
      Caption         =   "�P��q��"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton B1 
      Caption         =   "�ɾ\�q��"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton B3 
      Caption         =   "�B�~���O"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton B4 
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
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   240
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "���I��U�C�ݭn�d�ߪ������������:"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "OBOSlecteView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub B1_Click()
OBOVIEW.Refresh
OBOVIEW.Show
OBOVIEW.SSTab1.Tab = 0
OBOSlecteView.Hide
End Sub

Private Sub B2_Click()
OBOVIEW.Refresh
OBOVIEW.Show
OBOVIEW.SSTab1.Tab = 1
OBOSlecteView.Hide
End Sub

Private Sub B3_Click()
OBOVIEW.Refresh
OBOVIEW.Show
OBOVIEW.SSTab1.Tab = 2
OBOSlecteView.Hide
End Sub

Private Sub B4_Click()
StaffUse.Show
OBOSlecteView.Hide
End Sub
