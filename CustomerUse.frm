VERSION 5.00
Begin VB.Form CustomerUse 
   Caption         =   "客戶或會員專用"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   12375
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Exit 
      Caption         =   "離開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      TabIndex        =   4
      Top             =   5400
      Width           =   2175
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
      Height          =   1455
      Left            =   6840
      TabIndex        =   3
      Top             =   3840
      Width           =   2655
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
      Height          =   1575
      Left            =   6840
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
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
      Height          =   1575
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
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
      Height          =   1455
      Left            =   2760
      TabIndex        =   0
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "請選擇以下的功能,並按下相關功能的按鈕:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   26.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "CustomerUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataBorrow_Click()
LoginID.Show
CustomerUse.Enabled = False
End Sub

Private Sub Exit_Click()
Menu.Show
CustomerUse.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Search_Click()
SearchForMember.Show
CustomerUse.Hide
End Sub

Private Sub Top10_Click()
Top10ForMember.Show
CustomerUse.Hide
End Sub
