VERSION 5.00
Begin VB.Form COMICSELECTVIEW 
   Caption         =   "銷售,租借及歸還功能選擇"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7920
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton ReturnBook 
      Caption         =   "歸還書本"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton BorrowBook 
      Caption         =   "租借書本"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton ProductSell 
      Caption         =   "貨品銷售"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton TurnBack 
      Caption         =   "返回主菜單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   720
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "請點選下列需要使用的功能:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "COMICSELECTVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BorrowBook_Click()
COMICS.Refresh
COMICS.Show
COMICS.SSTab1.Tab = 0
COMICSELECTVIEW.Hide
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\PIC\2233003_170028053_2.jpg")
Image1.Move 0, 0, COMICSELECTVIEW.ScaleWidth, COMICSELECTVIEW.ScaleHeight
End Sub

Private Sub ProductSell_Click()
COMICS.Refresh
COMICS.Show
COMICS.SSTab1.Tab = 2
COMICSELECTVIEW.Hide
End Sub

Private Sub ReturnBook_Click()
COMICS.Refresh
COMICS.Show
COMICS.SSTab1.Tab = 1
COMICSELECTVIEW.Hide
End Sub

Private Sub TurnBack_Click()
StaffUse.Show
COMICSELECTVIEW.Hide
End Sub
