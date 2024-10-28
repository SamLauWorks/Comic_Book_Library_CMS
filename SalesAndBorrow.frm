VERSION 5.00
Begin VB.Form SalesAndBorrow 
   ClientHeight    =   7950
   ClientLeft      =   3855
   ClientTop       =   2895
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   12945
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton FINEPAY 
      Caption         =   "罰款歸還(只適用於欠繳罰款的會員)"
      Height          =   855
      Left            =   8520
      TabIndex        =   28
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox BCID 
      Height          =   495
      Left            =   9120
      TabIndex        =   25
      Text            =   "Text4"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton FINESAVE 
      Caption         =   "罰款拖欠"
      Height          =   855
      Left            =   7200
      TabIndex        =   24
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Exit 
      Caption         =   "離開"
      Height          =   735
      Left            =   5280
      TabIndex        =   23
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox STA 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3840
      Width           =   3255
   End
   Begin VB.TextBox FIN 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   4560
      Width           =   3255
   End
   Begin VB.TextBox BNS 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton RenewMode 
      Caption         =   "漫畫歸還模式"
      Height          =   855
      Left            =   11040
      TabIndex        =   16
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Payment 
      Caption         =   "付款"
      Height          =   615
      Left            =   1200
      TabIndex        =   13
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton MoveDOWN 
      Caption         =   "下移"
      Height          =   735
      Left            =   4200
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton MoveUP 
      Caption         =   "上移"
      Height          =   735
      Left            =   4200
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox OrderID 
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton BorrowMode 
      Caption         =   "漫畫租借模式"
      Height          =   855
      Left            =   11040
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton SalesMode 
      Caption         =   "銷售模式"
      Height          =   1095
      Left            =   11040
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Scanning 
      Caption         =   "檢查編號"
      Height          =   855
      Left            =   11520
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Codes 
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton Del 
      Caption         =   "刪除"
      Height          =   735
      Left            =   4200
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Edit 
      Caption         =   "更改"
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ListBox OrderList 
      Height          =   3300
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   3975
   End
   Begin VB.PictureBox MemberInfo 
      Height          =   2175
      Left            =   6720
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "今天日期"
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label BCN 
      Caption         =   "複本編號:"
      Height          =   495
      Left            =   7920
      TabIndex        =   27
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label IDTYPE 
      Caption         =   "貨品編號"
      Height          =   495
      Left            =   7920
      TabIndex        =   26
      Top             =   0
      Width           =   975
   End
   Begin VB.Label ST 
      Caption         =   "狀態:"
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label FI 
      Caption         =   "所欠罰款:"
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label BN 
      Caption         =   "可借閱冊數:"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "費用總和:"
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "清單列表:"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label OrderNames 
      Caption         =   "清單編號:"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "SalesAndBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BorrowMode_Click()
IDTYPE.Caption = "書本編號:"
OrderNames.Caption = "租借清單編號:"
BCN.Visible = True
BCID.Visible = True
MemberInfo.Visible = True
ST.Visible = True
STA.Visible = True
FI.Visible = True
FIN.Visible = True
BN.Visible = True
BNS.Visible = True
FINESAVE.Visible = False
FINEPAY.Value = True
End Sub

Private Sub Exit_Click()
StaffUse.Show
SalesAndBorrow.Hide
End Sub

Private Sub RenewMode_Click()
IDTYPE.Caption = "書本編號:"
OrderNames.Caption = "租借清單編號:"
BCN.Visible = True
BCID.Visible = True
MemberInfo.Visible = True
ST.Visible = True
STA.Visible = True
FI.Visible = True
FIN.Visible = True
BN.Visible = True
BNS.Visible = True
FINESAVE.Visible = True
FINEPAY.Value = True
End Sub

Private Sub SalesMode_Click()
IDTYPE.Caption = "貨品或書本編號:"
OrderNames.Caption = "清單編號"
BCN.Visible = False
BCID.Visible = False
MemberInfo.Visible = False
ST.Visible = False
STA.Visible = False
FI.Visible = False
FIN.Visible = False
BN.Visible = False
BNS.Visible = False
FINESAVE.Visible = False
FINEPAY.Value = False
End Sub
