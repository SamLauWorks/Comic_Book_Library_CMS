VERSION 5.00
Begin VB.Form StaffUse 
   Caption         =   "¾���M��"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   12795
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Member 
      BackColor       =   &H00FFFF00&
      Caption         =   "�|����ƪ��d�ߩM�]�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      MaskColor       =   &H80000003&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   7
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "�q��M���ɸ�ƪ��d��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8040
      MaskColor       =   &H80000003&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Incomes 
      BackColor       =   &H00FFFF00&
      Caption         =   "�d�\�C�ѩΨC�몺�`��~�B"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2040
      MaskColor       =   &H80000003&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   5
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton DataBorrow 
      BackColor       =   &H0000FF00&
      Caption         =   "�P��,���ɤ��k�٥\��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   4
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton Search 
      BackColor       =   &H00FFFF00&
      Caption         =   "�f�~�ήѥ���ƪ��d�ߩM�]�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      MaskColor       =   &H80000003&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton Top10 
      BackColor       =   &H00FFFF00&
      Caption         =   "�ˬd�Q�j�̨��w��H��]"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8040
      MaskColor       =   &H80000003&
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton About 
      Caption         =   "�w���έp�P��M�H��]"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5040
      TabIndex        =   1
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "�n�X�M���}"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "���z���e���t��(���e�αM��) V1.02(¾���M�ε���)"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   45.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12495
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   600
      Top             =   4920
      Width           =   1335
   End
End
Attribute VB_Name = "StaffUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
OBOSlecteView.Show
StaffUse.Hide
End Sub

Private Sub DataBorrow_Click()
COMICSELECTVIEW.Show
StaffUse.Hide
End Sub

Private Sub Exit_Click()
Menu.Show
StaffUse.Hide
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\PIC\dog005.jpg")
Image1.Move 0, 0, StaffUse.ScaleWidth, StaffUse.ScaleHeight
End Sub

Private Sub Incomes_Click()
IncomeView.Show
StaffUse.Hide
End Sub

Private Sub Member_Click()
MView.Show
StaffUse.Hide
End Sub

Private Sub Search_Click()
BookSelectView.Show
StaffUse.Hide
End Sub

Private Sub Top10_Click()
Top10ForStaff.Show
StaffUse.Hide
End Sub
