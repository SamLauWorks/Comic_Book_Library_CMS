VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   Caption         =   "漫畫屋支援"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   12645
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox P4 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   7080
      ScaleHeight     =   1755
      ScaleWidth      =   3675
      TabIndex        =   10
      Top             =   4320
      Width           =   3735
   End
   Begin VB.PictureBox P3 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   2400
      ScaleHeight     =   1755
      ScaleWidth      =   3675
      TabIndex        =   9
      Top             =   4320
      Width           =   3735
   End
   Begin VB.PictureBox P2 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   7080
      ScaleHeight     =   1635
      ScaleWidth      =   3675
      TabIndex        =   8
      Top             =   1680
      Width           =   3735
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   2400
      ScaleHeight     =   1635
      ScaleWidth      =   3675
      TabIndex        =   7
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton DataBorrow 
      Caption         =   "會員借閱查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   3360
      Width           =   3735
   End
   Begin VB.CommandButton Search 
      Caption         =   "貨品或書本相關查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      TabIndex        =   4
      Top             =   3360
      Width           =   3753
   End
   Begin VB.CommandButton Top10 
      Caption         =   "十大最受歡迎人氣榜"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      TabIndex        =   3
      Top             =   6120
      Width           =   3735
   End
   Begin VB.CommandButton About 
      Caption         =   "關於漫畫屋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      TabIndex        =   2
      Top             =   6120
      Width           =   3753
   End
   Begin VB.CommandButton Command2 
      Caption         =   "管理員專用"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      TabIndex        =   1
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "離開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1280
      Left            =   5520
      TabIndex        =   0
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "智慧漫畫店系統(漫畫屋專用) V1.02"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   45.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12495
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1695
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
ALogin.Refresh
ALogin.Show
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\PIC\8f816ff7.jpg")
Image1.Move 0, 0, Menu.ScaleWidth, Menu.ScaleHeight
P1.Picture = LoadPicture(App.Path & "\PIC\B1.jpg")
P1.PaintPicture P1.Picture, 0, 0, P1.ScaleWidth, P1.ScaleHeight
P2.Picture = LoadPicture(App.Path & "\PIC\B3.jpg")
P2.PaintPicture P2.Picture, 0, 0, P2.ScaleWidth, P2.ScaleHeight
P3.Picture = LoadPicture(App.Path & "\PIC\B4.jpg")
P3.PaintPicture P3.Picture, 0, 0, P3.ScaleWidth, P3.ScaleHeight
P4.Picture = LoadPicture(App.Path & "\PIC\B2.jpg")
P4.PaintPicture P4.Picture, 0, 0, P4.ScaleWidth, P4.ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub DataBorrow_Click()
LoginID.Refresh
LoginID.Show
Menu.Enabled = False
End Sub

Private Sub Search_Click()
LoginID.Refresh
SearchForMember.Show
Menu.Hide
End Sub

Private Sub Top10_Click()
LoginID.Refresh
Top10ForMember.Show
Menu.Hide
End Sub

