VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MEMBERINFO 
   Caption         =   "會員相關資訊及已借閱書本查詢"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11895
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame3 
      Caption         =   "會員所有借出書籍一覽:(下方的表格沒有顯示資料,該會員沒有借閱任何書籍)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   11655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MEMBERBORROWLIST 
         Height          =   2415
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   4260
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "會員資訊:(漫畫屋)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7215
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000C000&
         Caption         =   "可借閱冊數:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000C000&
         Caption         =   "所欠罰款:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000C000&
         Caption         =   "狀態:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
         Caption         =   "會員編號:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "會員資訊:(個人資料)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   9135
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FFFF&
         Caption         =   "電話號碼1:*"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FFFF&
         Caption         =   "電話號碼2:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "地址:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         Caption         =   "出生日期:*"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "性別:*@"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Caption         =   "會員姓名:*"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Exit 
      Caption         =   "登出及離開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1320
      Top             =   7440
      Width           =   1935
   End
End
Attribute VB_Name = "MEMBERINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub Exit_Click()
Menu.Show
MEMBERINFO.Hide
cn.Close
End Sub

Private Sub Form_Activate()
Image1.Picture = LoadPicture(App.Path & "\PIC\AS.jpg")
Image1.Move 0, 0, MEMBERINFO.ScaleWidth, MEMBERINFO.ScaleHeight
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "Member", cn, adOpenKeyset, adLockOptimistic
On Error Resume Next
rs.MoveFirst
rs.Find "MemberID = '" & Text1(0).Text & "'"
FindFirst = Not rs.EOF And Err.Number = 0
For i = 0 To rs.RecordCount - 1
Text1(i).Text = rs.Fields(i)
Next i
rs2.Open "Select Book.BookID as 書籍編號, BookCopy.CopyN as 書籍複本編號, Book.BookName1 as 書籍名稱, BorrowOrderList.BorrowDate as 借出日期, BorrowOrderList.DueDate as 歸還日期 FROM Book, BookCopy, BorrowOrder, BorrowOrderList,Member WHERE BorrowOrderList.BOID = BorrowOrder.BOID AND Book.BookID = BookCopy.BookID AND BookCopy.BookID = BorrowOrderList.BookID AND BookCopy.CopyN = BorrowOrderList.CopyN AND BorrowOrder.MemberID = Member.MemberID AND BorrowOrderList.status = 'B' AND Member.MemberID = '" & Text1(0).Text & "'", cn, adOpenKeyset, adLockOptimistic
Set MEMBERBORROWLIST.DataSource = rs2
rs2.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
