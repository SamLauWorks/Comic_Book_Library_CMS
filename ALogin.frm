VERSION 5.00
Begin VB.Form ALogin 
   Caption         =   "�޲z���n�J"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command1 
      Caption         =   "�n�J"
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
      Caption         =   "�޲z���K�X"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label ID 
      Caption         =   "�޲z���b��:"
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
'ACC.Text = "�n�J���\!"
StaffUse.Show
ALogin.Hide
Menu.Hide
'Else
'ACC.Text = "�n�J����!"
'End If
End Sub

Private Sub Form_Activate()
a = "A1111111"
b = "Y2222222"
End Sub

Private Sub Form_Unload(Cancel As Integer)
ALogin.Hide
End Sub
