VERSION 5.00
Begin VB.Form AddBC 
   Caption         =   "增加書籍總集"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6060
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Clears 
      Caption         =   "清除所輸入的資料"
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton ADDS 
      Caption         =   "確定"
      Height          =   735
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "離開"
      Height          =   735
      Left            =   4200
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox BC3 
      Height          =   300
      ItemData        =   "AddBC.frx":0000
      Left            =   1440
      List            =   "AddBC.frx":000A
      TabIndex        =   0
      Text            =   "BC3"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label RULES 
      Caption         =   "Label1"
      Height          =   1575
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "*書籍總集編號:"
      Height          =   495
      Index           =   0
      Left            =   3960
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "*書籍總集名稱:"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "*書籍總集狀態:"
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "AddBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub ADDS_Click()
Dim Y As String
Y = ""
MsgBox "是否新增該書籍總集？", vbYesNo, "新增書籍總集"
If vbYes Then
'Checking
If Text2(1).Text = "" Then
Label3(1).BackColor = vbRed
Y = Y & vbCrLf & "書籍總集沒有名稱!"
End If
If Text2(2).Text = "" Then
Label3(2).BackColor = vbRed
Y = Y & vbCrLf & "書籍總集狀態不可空白!"
End If
If Text2(2).Text <> "全本" Or Text2(2).Text <> "連載" Then
Label3(2).BackColor = vbRed
Y = Y & vbCrLf & "書籍總集狀態只能填入指定的項目!"
End If
If Y = "" Then
Call CreateNewID
rs2.AddNew
  For i = 0 To Text1.Count - 1
        If IsNull(Text2(i).Text) Then
        rs.Fields(i) = ""
        Else
        rs.Fields(i) = Text2(i).Text
        End If
        Next i
rs2.Update
MsgBox "該書籍總集已新增!"
End If
End Sub

Private Sub Command1_Click()
AddBC.Hide
cn.Close
End Sub

Private Sub Form_Activate()
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "Book", cn, adOpenKeyset, adLockOptimistic
  RULES.Caption = "注意:" & vbCrLf & "1.書籍總集編號為系統自動生成." & vbCrLf & "2.帶有*的相關資料必須輸入資料" & vbCrLf & "3.書籍總集狀態只能輸入: 連載 或 全本" & vbCrLf & "4.如果有資料輸入錯誤,提示視窗將會彈出,錯誤的相關欄位將會變成紅色."
End Sub
Function CreateNewID()
Dim newID As String
Dim CID As Integer
Dim fail As Integer
rs.MoveLast
newID = rs.Fields(0)
CID = CInt(Right(rs.Fields(0), 7)) + 1
fail = 8 - Len(CStr(CID))
newID = ""
newID = Trim(newID + "C" + String(fail - 1, "0") + CStr(CID))
Text2(0).Text = Trim(newID)
End Function
