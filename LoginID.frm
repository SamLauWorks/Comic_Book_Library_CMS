VERSION 5.00
Begin VB.Form LoginID 
   Caption         =   "請輸入你的會員編號"
   ClientHeight    =   2040
   ClientLeft      =   5985
   ClientTop       =   4770
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2040
   ScaleWidth      =   6000
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton checks 
      Caption         =   "登入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox ID 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入你的會員編號:(例如:M0000001)"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "LoginID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim stuRec As String
Private Sub checks_Click()
'If stuRec = "" Then
'End If
stuRec = FindFirst
If stuRec = "True" Then
MEMBERINFO.Show
Menu.Hide
LoginID.Hide
Menu.Enabled = True
MEMBERINFO.Text1(0).Text = ID.Text
ID.Text = ""
rs.Close
cn.Close
Else
MsgBox "你的會員編號錯誤!"
End If
End Sub

Private Sub Form_Activate()
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "Member", cn, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
Menu.Enabled = True
rs.Close
cn.Close
End Sub

Function FindFirst()
On Error Resume Next
rs.MoveFirst
rs.Find "MemberID = '" & ID.Text & "'"
FindFirst = Not rs.EOF And Err.Number = 0
End Function
