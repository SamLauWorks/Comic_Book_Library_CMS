VERSION 5.00
Begin VB.Form ALogin 
   Caption         =   "管理員登入"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      Caption         =   "登入"
      Height          =   855
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox PW 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox ACC 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "管理員密碼"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label ID 
      Caption         =   "管理員帳號:"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "ALogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As String
Private Sub Command1_Click()
'If ACC.Text = a And PW.Text = b Then
'ACC.Text = "登入成功!"
StaffUse.Show
ALogin.Hide
Menu.Hide
'Else
'ACC.Text = "登入失敗!"
'End If
End Sub

Private Sub Form_Activate()
a = "A1111111"
b = "Y2222222"
End Sub

Private Sub Form_Unload(Cancel As Integer)
ALogin.Hide
End Sub
