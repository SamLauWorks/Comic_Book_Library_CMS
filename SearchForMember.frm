VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SearchForMember 
   BackColor       =   &H00FFFFFF&
   Caption         =   "貨品或書本相關查詢"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   12765
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "請選取查詢類別,選取關鍵字類型並在相應的欄中輸入關鍵字:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12735
      Begin VB.CheckBox FF 
         BackColor       =   &H0000FFFF&
         Caption         =   "已完結"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   20
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "搜尋"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11040
         Style           =   1  '圖片外觀
         TabIndex        =   19
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox A 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   18
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox BB 
         BackColor       =   &H0000FFFF&
         Caption         =   "書本/書本總集/貨品名稱"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   17
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CheckBox CC 
         BackColor       =   &H0000FFFF&
         Caption         =   "作者名稱"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   16
         Top             =   1680
         Width           =   4095
      End
      Begin VB.CheckBox DD 
         BackColor       =   &H0000FFFF&
         Caption         =   "書本/貨品價錢"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   15
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox B 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   14
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox C 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   13
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox D 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   12
         Top             =   2280
         Width           =   3615
      End
      Begin VB.CheckBox AA 
         BackColor       =   &H0000FFFF&
         Caption         =   "書本/書本總集/貨品編號"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   11
         Top             =   480
         Width           =   4095
      End
      Begin VB.CheckBox EE 
         BackColor       =   &H0000FF00&
         Caption         =   "只顯示尚未借出的書本"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   3855
      End
      Begin VB.OptionButton OO5 
         BackColor       =   &H0000FF00&
         Caption         =   "貨品(銷售資訊)"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   3015
      End
      Begin VB.OptionButton OO2 
         BackColor       =   &H0000FF00&
         Caption         =   "書本總集(銷售資訊)"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton OO1 
         BackColor       =   &H0000FF00&
         Caption         =   "書本總集(借閱資訊)"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton OO4 
         BackColor       =   &H0000FF00&
         Caption         =   "書本(借閱資訊)"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   2655
      End
      Begin VB.OptionButton OO3 
         BackColor       =   &H0000FF00&
         Caption         =   "書本(銷售資訊)"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2655
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid SEARCHRESULT 
      Height          =   3375
      Left            =   0
      TabIndex        =   6
      Top             =   4200
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   16777215
      BackColorBkg    =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "離開"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "0筆"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "查詢結果:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   3600
      Width           =   1815
   End
End
Attribute VB_Name = "SearchForMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub AA_Click()
If AA.Value Then
    A.Enabled = True
Else
    A.Enabled = False
End If
End Sub

Private Sub BB_Click()
If BB.Value Then
    B.Enabled = True
Else
    B.Enabled = False
End If
End Sub

Private Sub CC_Click()
If CC.Value Then
    C.Enabled = True
Else
    C.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Dim CIDD As String
Dim CIDN As String
Dim BA As String
Dim AP As String
Dim BCS As String
Dim CBS As String
Dim SQLS As String
    If OO1.Value = True Then
    If AA.Value Then
        CIDD = "(BookCollection.CollectionID like '%" & A.Text & "%' or Book.BookID like '%" & A.Text & "%')"
        Else
        CIDD = "BookCollection.CollectionID like '%'"
    End If
    If BB.Value Then
        CIDN = "(BookCollection.CollectionBookName like '%" & B.Text & "%' or Book.BookName1 like '%" & B.Text & "%' or Book.BookName2 like '%" & B.Text & "%')"
        Else
        CIDN = "BookCollection.CollectionBookName like '%'"
    End If
    If CC.Value Then
        BA = "(Book.Author1 like '%" & C.Text & "%' or Book.Author2 like '%" & C.Text & "%')"
        Else
        BA = "Book.Author1 like '%'"
    End If
    If DD.Value Then
        AP = "Book.Price like '%" & D.Text & "%'"
        Else
        AP = "Book.Price like '%'"
    End If
    If EE.Value Then
        BCS = "HAVING BookCopy.Status = '架上'"
        Else
        BCS = " "
    End If
    If FF.Value Then
        CBS = "BookCollection.status = '全本'"
        Else
        CBS = "BookCollection.status like '%'"
    End If
    SQLS = "SELECT BookCollection.CollectionID, BookCollection.CollectionBookName, Book.BookID, BookCopy.CopyN, BookCopy.Status, Book.BookName1, Book.BookName2, Book.Author1, Book.Author2, Book.BookType, Book.ContentType1, Book.ContentType2, Book.PublishingHouse, Book.Year, Book.Price, Book.Episode From BookCollection, BookCopy, Book WHERE BookCollection.CollectionID=Book.CollectionID AND BookCopy.BookID=Book.BookID AND Book.Status='N' AND " & CIDD & " AND " & CIDN & " AND " & BA & " AND " & AP & " AND " & CBS & " GROUP BY BookCollection.CollectionID, BookCollection.CollectionBookName, Book.BookID, BookCopy.CopyN, BookCopy.Status, Book.BookName1, Book.BookName2, Book.Author1, Book.Author2, Book.BookType, Book.ContentType1, Book.ContentType2, Book.PublishingHouse, Book.Year, Book.Price, Book.Episode " & BCS & " ORDER BY 1, 3, 4;"
    rs.Open SQLS, cn, adOpenKeyset, adLockOptimistic
    Set SEARCHRESULT.DataSource = rs
    rs.Close
    Label3.Caption = CStr(SEARCHRESULT.Rows - 1) + "筆"
    End If
    If OO2.Value = True Then
    If AA.Value Then
        CIDD = "(BookCollection.CollectionID like '%" & A.Text & "%' or Book.BookID like '%" & A.Text & "%')"
        Else
        CIDD = "BookCollection.CollectionID like '%'"
    End If
    If BB.Value Then
        CIDN = "(BookCollection.CollectionBookName like '%" & B.Text & "%' or Book.BookName1 like '%" & B.Text & "%' or Book.BookName2 like '%" & B.Text & "%')"
        Else
        CIDN = "BookCollection.CollectionBookName like '%'"
    End If
    If CC.Value Then
        BA = "(Book.Author1 like '%" & C.Text & "%' or Book.Author2 like '%" & C.Text & "%')"
        Else
        BA = "Book.Author1 like '%'"
    End If
    If DD.Value Then
        AP = "Book.Price like '%" & D.Text & "%'"
        Else
        AP = "Book.Price like '%'"
    End If
    If FF.Value Then
        CBS = "BookCollection.status = '全本'"
        Else
        CBS = "BookCollection.status like '%'"
    End If
    SQLS = "SELECT BookCollection.CollectionID, BookCollection.CollectionBookName, Book.BookID, BookCopy.CopyN, BookCopy.Status, Book.BookName1, Book.BookName2, Book.Author1, Book.Author2, Book.BookType, Book.ContentType1, Book.ContentType2, Book.PublishingHouse, Book.Year, Book.Price, Book.Episode From BookCollection, BookCopy, Book WHERE BookCollection.CollectionID=Book.CollectionID AND BookCopy.BookID=Book.BookID AND Book.Status='N' AND " & CIDD & " AND " & CIDN & " AND " & BA & " AND " & AP & " AND " & CBS & " GROUP BY BookCollection.CollectionID, BookCollection.CollectionBookName, Book.BookID, BookCopy.CopyN, BookCopy.Status, Book.BookName1, Book.BookName2, Book.Author1, Book.Author2, Book.BookType, Book.ContentType1, Book.ContentType2, Book.PublishingHouse, Book.Year, Book.Price, Book.Episode  ORDER BY 1, 3, 4;"
    rs.Open SQLS, cn, adOpenKeyset, adLockOptimistic
    Set SEARCHRESULT.DataSource = rs
    rs.Close
    End If
    If OO3.Value = True Then
    If AA.Value Then
        CIDD = "(BookCollection.CollectionID like '%" & A.Text & "%' or Book.BookID like '%" & A.Text & "%')"
        Else
        CIDD = "BookCollection.CollectionID like '%'"
    End If
    If BB.Value Then
        CIDN = "(BookCollection.CollectionBookName like '%" & B.Text & "%' or Book.BookName1 like '%" & B.Text & "%' or Book.BookName2 like '%" & B.Text & "%')"
        Else
        CIDN = "Book.BookName1 like '%'"
    End If
    If CC.Value Then
        BA = "(Book.Author1 like '%" & C.Text & "%' or Book.Author2 like '%" & C.Text & "%')"
        Else
        BA = "Book.Author1 like '%'"
    End If
    If DD.Value Then
        AP = "Book.Price like '%" & D.Text & "%'"
        Else
        AP = "Book.Price like '%'"
    End If
    SQLS = "SELECT BookCollection.CollectionID, BookCollection.CollectionBookName, Book.BookID, Book.BookName1, Book.BookName2, Book.Author1, Book.Author2, Book.BookType, Book.ContentType1, Book.ContentType2, Book.PublishingHouse, Book.Year, Book.Price, Book.Episode From BookCollection, Book WHERE BookCollection.CollectionID=Book.CollectionID AND Book.Status='N' AND " & CIDD & " AND " & CIDN & " AND " & BA & " AND " & AP & " GROUP BY BookCollection.CollectionID, BookCollection.CollectionBookName, Book.BookID, Book.BookName1, Book.BookName2, Book.Author1, Book.Author2, Book.BookType, Book.ContentType1, Book.ContentType2, Book.PublishingHouse, Book.Year, Book.Price, Book.Episode  ORDER BY 1, 3, 4;"
    rs.Open SQLS, cn, adOpenKeyset, adLockOptimistic
    Set SEARCHRESULT.DataSource = rs
    rs.Close
    End If
    If OO4.Value = True Then
    If AA.Value Then
        CIDD = "(BookCollection.CollectionID like '%" & A.Text & "%' or Book.BookID like '%" & A.Text & "%')"
        Else
        CIDD = "Book.BookID like '%'"
    End If
    If BB.Value Then
        CIDN = "(BookCollection.CollectionBookName like '%" & B.Text & "%' or Book.BookName1 like '%" & B.Text & "%' or Book.BookName2 like '%" & B.Text & "%')"
        Else
        CIDN = "Book.BookName1 like '%'"
    End If
    If CC.Value Then
        BA = "(Book.Author1 like '%" & C.Text & "%' or Book.Author2 like '%" & C.Text & "%')"
        Else
        BA = "Book.Author1 like '%'"
    End If
    If DD.Value Then
        AP = "Book.Price like '%" & D.Text & "%'"
        Else
        AP = "Book.Price like '%'"
    End If
    If EE.Value Then
        BCS = "HAVING BookCopy.Status = '架上'"
        Else
        BCS = " "
    End If
    End If
    If OO5.Value = True Then
    If AA.Value Then
        CIDD = "ProductID like '%" & A.Text & "%'"
        Else
        CIDD = "ProductID like '%'"
    End If
    If BB.Value Then
        CIDN = "ProductName like '%" & B.Text & "%'"
        Else
        CIDN = "ProductName like '%'"
    End If
    If DD.Value Then
        AP = "Price like '%" & D.Text & "%'"
        Else
        AP = "Price like '%'"
    End If
    SQLS = "SELECT * From Product Where " & CIDD & " AND  " & CIDN & " AND " & AP & ""
    rs.Open SQLS, cn, adOpenKeyset, adLockOptimistic
    Set SEARCHRESULT.DataSource = rs
    rs.Close
    Else
    MsgBox "搜尋結果:沒有任何一項符合關鍵字!"
    End If
End Sub

Private Sub DD_Click()
If DD.Value Then
    D.Enabled = True
Else
    D.Enabled = False
End If
End Sub

Private Sub Exit_Click()
Menu.Show
SearchForMember.Hide
cn.Close
End Sub

Private Sub Form_Activate()
 cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
 SEARCHRESULT.ColWidth(0) = 150 * 3
End Sub

Private Sub OO1_Click()
If OO1.Value Then
    EE.Enabled = True
    FF.Enabled = True
    CC.Enabled = True
    C.Enabled = True
End If
End Sub

Private Sub OO3_Click()
If OO3.Value Then
EE.Enabled = False
FF.Enabled = False
CC.Enabled = True
End If
End Sub

Private Sub OO4_Click()
If OO4.Value Then
EE.Enabled = True
FF.Enabled = False
CC.Enabled = True
End If
End Sub

Private Sub OO2_Click()
If OO2.Value Then
EE.Enabled = False
FF.Enabled = True
End If
End Sub

Private Sub OO5_Click()
If OO5.Value Then
EE.Enabled = False
FF.Enabled = False
CC.Enabled = False
End If
End Sub
