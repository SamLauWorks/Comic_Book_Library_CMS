VERSION 5.00
Begin VB.Form MView 
   Caption         =   "會員資料的查詢和設定界面"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   10335
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      Caption         =   "回到主菜單"
      Height          =   855
      Left            =   3480
      TabIndex        =   18
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "會員資訊"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   1
         Left            =   5280
         TabIndex        =   16
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   2
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   3
         Left            =   4080
         TabIndex        =   14
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   4
         Left            =   1320
         TabIndex        =   13
         Top             =   1440
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   5
         Left            =   1320
         TabIndex        =   12
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   6
         Left            =   4800
         TabIndex        =   11
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   7
         Left            =   1320
         TabIndex        =   10
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   8
         Left            =   3840
         TabIndex        =   9
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   9
         Left            =   6840
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton BookEdit 
         Caption         =   "變更會員資料"
         Height          =   735
         Left            =   3840
         TabIndex        =   7
         Top             =   4560
         Width           =   2415
      End
      Begin VB.CommandButton MoveOne 
         Caption         =   "移到第一項的會員資料"
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton Perv 
         Caption         =   "移到上一項的會員資料"
         Height          =   735
         Left            =   2640
         TabIndex        =   5
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton Nexts 
         Caption         =   "移到下一項的會員資料"
         Height          =   735
         Left            =   5160
         TabIndex        =   4
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton MoveLast 
         Caption         =   "移到最後一項的會員資料"
         Height          =   735
         Left            =   7440
         TabIndex        =   3
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CommandButton ADDS 
         Caption         =   "新增會員資料"
         Height          =   735
         Left            =   1440
         TabIndex        =   2
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton DELS 
         Caption         =   "刪除會員資料"
         Height          =   735
         Left            =   6480
         TabIndex        =   1
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "電話號碼1:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "電話號碼2:"
         Height          =   375
         Left            =   3840
         TabIndex        =   27
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "狀態:"
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "所欠罰款:"
         Height          =   375
         Left            =   3000
         TabIndex        =   25
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "可借閱冊數:"
         Height          =   375
         Left            =   5760
         TabIndex        =   24
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "地址:"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "出生日期:"
         Height          =   375
         Left            =   2880
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "性別:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "會員姓名:"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "會員卡片編號:"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "MView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub ADDS_Click()
rs.AddNew
  For i = 0 To Text1.Count - 1
        If IsNull(Text1(i).Text) Then
        rs.Fields(i) = ""
        Else
        rs.Fields(i) = Text1(i).Text
        End If
        Next i
rs.Update
End Sub

Private Sub BookEdit_Click()
For i = 1 To rs.Fields.Count - 1
        If IsNull(rs.Fields(i)) Then
        rs.Fields(i) = ""
        rs.Update
        Else
        rs.Fields(i) = Text1(i).Text
        rs.Update
        End If
        Next i
End Sub

Private Sub Command1_Click()
StaffUse.Show
MView.Hide
End Sub

Private Sub DELS_Click()
rs.Delete
rs.MoveNext
If rs.EOF Then
   rs.MoveLast
End If
For i = 0 To rs.Fields.Count - 1
        If IsNull(rs.Fields(i)) Then
        Text1(i).Text = ""
        Else
        Text1(i).Text = rs.Fields(i)
        End If
Next i
End Sub

Private Sub Form_Load()
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "Member", cn, adOpenKeyset, adLockOptimistic
  For i = 0 To rs.Fields.Count - 1
        If IsNull(rs.Fields(i)) Then
        Text1(i).Text = ""
        Else
        Text1(i).Text = rs.Fields(i)
        End If
  Next i
End Sub
Private Sub MoveLast_Click()
rs.MoveLast
  For i = 0 To rs.Fields.Count - 1
  If IsNull(rs.Fields(i)) Then
  Text1(i).Text = ""
  Else
  Text1(i).Text = rs.Fields(i)
  End If
  Next i
End Sub

Private Sub MoveOne_Click()
rs.MoveFirst
  For i = 0 To rs.Fields.Count - 1
  If IsNull(rs.Fields(i)) Then
  Text1(i).Text = ""
  Else
  Text1(i).Text = rs.Fields(i)
  End If
  Next i
End Sub

Private Sub Nexts_Click()
rs.MoveNext
If rs.EOF Then
   rs.MoveLast
End If
For i = 0 To rs.Fields.Count - 1
        If IsNull(rs.Fields(i)) Then
        Text1(i).Text = ""
        Else
        Text1(i).Text = rs.Fields(i)
        End If
Next i
End Sub

Private Sub Perv_Click()
rs.MovePrevious
If rs.BOF Then
   rs.MoveFirst
End If
For i = 0 To rs.Fields.Count - 1
        If IsNull(rs.Fields(i)) Then
        Text1(i).Text = ""
        Else
        Text1(i).Text = rs.Fields(i)
        End If
Next i
End Sub
