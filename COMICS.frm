VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form COMICS 
   Caption         =   "銷售,租借及歸還功能介面"
   ClientHeight    =   10275
   ClientLeft      =   495
   ClientTop       =   1200
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   13545
   StartUpPosition =   2  '螢幕中央
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10200
      Top             =   9120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "離開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   2
      Top             =   9000
      Width           =   3015
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   15478
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   8421376
      ForeColor       =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "書本借閱模式"
      TabPicture(0)   =   "COMICS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label16"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label17"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "BOID"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DELBO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "NowDate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Due"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MSHFlexGrid1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "PAYMENT"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "MSHFlexGrid3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "A"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "書本歸還模式"
      TabPicture(1)   =   "COMICS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "銷售模式"
      TabPicture(2)   =   "COMICS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "FINDBP"
      Tab(2).Control(2)=   "Command8"
      Tab(2).Control(3)=   "FINISHORDER"
      Tab(2).Control(4)=   "BPTOTALSELL"
      Tab(2).Control(5)=   "ORDERLIST"
      Tab(2).Control(6)=   "NowDate3"
      Tab(2).Control(7)=   "OID"
      Tab(2).Control(8)=   "Picture2"
      Tab(2).Control(9)=   "Label34"
      Tab(2).Control(10)=   "Label33"
      Tab(2).Control(11)=   "Label32"
      Tab(2).ControlCount=   12
      Begin VB.Frame Frame5 
         Caption         =   "書籍/貨品編號:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -66000
         TabIndex        =   67
         Top             =   3840
         Width           =   4335
         Begin VB.CommandButton TAKEORDER 
            Caption         =   "確認"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2760
            TabIndex        =   69
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox INPUTBP 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MaxLength       =   8
            TabIndex        =   68
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label35 
            BackColor       =   &H00FFC0C0&
            Caption         =   "書籍/貨品編號:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "會員登錄:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   9000
         TabIndex        =   63
         Top             =   3840
         Width           =   4095
         Begin VB.CommandButton CM 
            Caption         =   "確認"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   65
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox ID 
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
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   64
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "會員編號:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "輸入租借書本:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   9000
         TabIndex        =   57
         Top             =   5400
         Width           =   4095
         Begin VB.TextBox BID 
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
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   60
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton CB 
            Caption         =   "確認"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1320
            TabIndex        =   59
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox CPID 
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
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   58
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "書籍編號:"
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
            TabIndex        =   62
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFC0C0&
            Caption         =   "書籍複本編號:"
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
            TabIndex        =   61
            Top             =   1080
            Width           =   2175
         End
      End
      Begin VB.Frame A 
         Caption         =   "會員租借項目:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   9240
         TabIndex        =   51
         Top             =   720
         Width           =   4095
         Begin VB.CommandButton Command5 
            Caption         =   "罰款清還"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2280
            TabIndex        =   56
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   8
            Left            =   1920
            TabIndex        =   53
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   9
            Left            =   1920
            TabIndex        =   52
            Top             =   600
            Width           =   1695
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
            Height          =   495
            Left            =   120
            TabIndex        =   55
            Top             =   1320
            Width           =   1815
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
            Height          =   495
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   1815
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FINDBP 
         Height          =   975
         Left            =   -64200
         TabIndex        =   50
         Top             =   7440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1720
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton Command8 
         Caption         =   "刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74760
         TabIndex        =   49
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CommandButton FINISHORDER 
         Caption         =   "付款"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -66000
         TabIndex        =   47
         Top             =   7200
         Width           =   1335
      End
      Begin VB.TextBox BPTOTALSELL 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67800
         TabIndex        =   46
         Text            =   "0"
         Top             =   7320
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid ORDERLIST 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   45
         Top             =   3840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.TextBox NowDate3 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -69000
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox OID 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   40
         Top             =   3360
         Width           =   2175
      End
      Begin VB.PictureBox Picture2 
         Height          =   2175
         Left            =   -74880
         ScaleHeight     =   2115
         ScaleWidth      =   13155
         TabIndex        =   39
         Top             =   840
         Width           =   13215
      End
      Begin VB.Frame Frame1 
         Caption         =   "會員資訊:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   9015
         Begin VB.TextBox Text1 
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
            TabIndex        =   23
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   22
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   21
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   20
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   19
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox Text1 
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
            Left            =   5520
            TabIndex        =   18
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox Text1 
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
            Left            =   5520
            TabIndex        =   17
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Text1 
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
            Left            =   5520
            TabIndex        =   16
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000FFFF&
            Caption         =   "電話號碼1:"
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
            TabIndex        =   31
            Top             =   960
            Width           =   1695
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
            TabIndex        =   30
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackColor       =   &H0000FFFF&
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
            TabIndex        =   29
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label11 
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
            TabIndex        =   28
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFF00&
            Caption         =   "出生日期:"
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
            TabIndex        =   27
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFF00&
            Caption         =   "性別:"
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
            TabIndex        =   26
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "會員姓名:"
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
            TabIndex        =   25
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFF00&
            Caption         =   "會員卡片編號:"
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
            TabIndex        =   24
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   8175
         Left            =   -74880
         ScaleHeight     =   8115
         ScaleWidth      =   13155
         TabIndex        =   14
         Top             =   720
         Width           =   13215
         Begin VB.CommandButton CLEARLIST 
            Caption         =   "清除暫存資料"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8880
            TabIndex        =   100
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Frame Frame7 
            Caption         =   "會員資訊:"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   0
            TabIndex        =   81
            Top             =   0
            Width           =   9015
            Begin VB.TextBox REMEID 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   5520
               TabIndex        =   97
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox REMEID 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   6
               Left            =   5520
               TabIndex        =   96
               Top             =   1440
               Width           =   1695
            End
            Begin VB.TextBox REMEID 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   5520
               TabIndex        =   95
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox REMEID 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   5520
               TabIndex        =   94
               Top             =   1920
               Width           =   3375
            End
            Begin VB.TextBox REMEID 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   2160
               TabIndex        =   93
               Top             =   1920
               Width           =   1695
            End
            Begin VB.TextBox REMEID 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   2160
               TabIndex        =   92
               Top             =   1440
               Width           =   1695
            End
            Begin VB.TextBox REMEID 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   2160
               TabIndex        =   91
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox REMEID 
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   2160
               TabIndex        =   90
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label25 
               BackColor       =   &H00FFFF00&
               Caption         =   "會員卡片編號:"
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
               TabIndex        =   89
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFFF00&
               Caption         =   "會員姓名:"
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
               Index           =   1
               Left            =   120
               TabIndex        =   88
               Top             =   960
               Width           =   2055
            End
            Begin VB.Label Label24 
               BackColor       =   &H00FFFF00&
               Caption         =   "性別:"
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
               TabIndex        =   87
               Top             =   1440
               Width           =   2055
            End
            Begin VB.Label Label23 
               BackColor       =   &H00FFFF00&
               Caption         =   "出生日期:"
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
               TabIndex        =   86
               Top             =   1920
               Width           =   2055
            End
            Begin VB.Label Label22 
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
               TabIndex        =   85
               Top             =   1920
               Width           =   1695
            End
            Begin VB.Label Label21 
               BackColor       =   &H0000FFFF&
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
               TabIndex        =   84
               Top             =   480
               Width           =   1695
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
               Index           =   1
               Left            =   3840
               TabIndex        =   83
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label20 
               BackColor       =   &H0000FFFF&
               Caption         =   "電話號碼1:"
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
               TabIndex        =   82
               Top             =   960
               Width           =   1695
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "會員租借項目:"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   9120
            TabIndex        =   77
            Top             =   0
            Width           =   3975
            Begin VB.TextBox REMEID 
               Height          =   495
               Index           =   9
               Left            =   1920
               TabIndex        =   99
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox REMEID 
               Height          =   495
               Index           =   8
               Left            =   1920
               TabIndex        =   98
               Top             =   1320
               Width           =   1815
            End
            Begin VB.CommandButton RETURNFEE 
               Caption         =   "罰款清還"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   2280
               TabIndex        =   78
               Top             =   1800
               Width           =   1455
            End
            Begin VB.Label Label19 
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
               Height          =   495
               Left            =   120
               TabIndex        =   80
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label18 
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
               Height          =   495
               Left            =   120
               TabIndex        =   79
               Top             =   1320
               Width           =   1815
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "輸入租借書本:"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   8880
            TabIndex        =   71
            Top             =   3240
            Width           =   4095
            Begin VB.CommandButton RETURNBP 
               Caption         =   "確認"
               BeginProperty Font 
                  Name            =   "新細明體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   1080
               TabIndex        =   76
               Top             =   1560
               Width           =   1455
            End
            Begin VB.TextBox CPID2 
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
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   75
               Top             =   1080
               Width           =   1575
            End
            Begin VB.TextBox BID2 
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
               Left            =   1800
               MaxLength       =   8
               TabIndex        =   74
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label Label37 
               BackColor       =   &H00FFC0C0&
               Caption         =   "書籍複本編號:"
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
               TabIndex        =   73
               Top             =   1080
               Width           =   2175
            End
            Begin VB.Label Label36 
               BackColor       =   &H00FFC0C0&
               Caption         =   "書籍編號:"
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
               TabIndex        =   72
               Top             =   480
               Width           =   1695
            End
         End
         Begin VB.Timer Timer1 
            Left            =   2040
            Top             =   7680
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid SEARCHMEMBER 
            Height          =   975
            Left            =   11400
            TabIndex        =   44
            Top             =   5760
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1720
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid SEARCHLIST 
            Height          =   975
            Left            =   11280
            TabIndex        =   36
            Top             =   6840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1720
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.TextBox FROMBO 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox NowDate2 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   2640
            Width           =   2055
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MEMBERBOOKLIST 
            Height          =   3855
            Left            =   120
            TabIndex        =   37
            Top             =   3600
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   6800
            _Version        =   393216
            Cols            =   5
            FillStyle       =   1
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
            _Band(0).Cols   =   5
         End
         Begin VB.Label Label28 
            Caption         =   "該會員未還書籍列表:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   3120
            Width           =   3495
         End
         Begin VB.Label Label30 
            Caption         =   "借閱訂單編號:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   2640
            Width           =   2415
         End
         Begin VB.Label Label29 
            Caption         =   "今天日期:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   34
            Top             =   2640
            Width           =   1695
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   855
         Left            =   11400
         TabIndex        =   13
         Top             =   7800
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox PAYMENT 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   11
         Text            =   "0"
         Top             =   8040
         Width           =   1935
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   3960
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   5
         FillStyle       =   1
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
         _Band(0).Cols   =   5
      End
      Begin VB.TextBox Due 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox NowDate 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton DELBO 
         Caption         =   "刪除"
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
         Left            =   120
         TabIndex        =   5
         Top             =   7920
         Width           =   1455
      End
      Begin VB.TextBox BOID 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "付款"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         TabIndex        =   1
         Top             =   8040
         Width           =   1335
      End
      Begin VB.Label Label34 
         Caption         =   "總價格:$"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   20.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69720
         TabIndex        =   48
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Label Label33 
         Caption         =   "今天日期:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70680
         TabIndex        =   43
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "訂單編號:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   42
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "總價格:$"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   20.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   12
         Top             =   8040
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "還書日期:"
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
         Left            =   8040
         TabIndex        =   8
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "今天日期:"
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
         Left            =   4440
         TabIndex        =   6
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "借閱訂單編號:"
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
         TabIndex        =   3
         Top             =   3360
         Width           =   2055
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   465
      Left            =   0
      TabIndex        =   101
      Top             =   9720
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   23680
            Text            =   "時間:"
            TextSave        =   "時間:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "COMICS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Private Sub CB_Click()
Dim stuRec2 As String
Dim BOOKFIND As String
Dim ss As Integer
Dim A As Integer
Dim B As Integer
Dim CHECKHAVE As Boolean
CHECKHAVE = False
BOOKFIND = "Select BookCopy.BookID, BookCopy.CopyN, Book.BookName1,BookCopy.Status From BookCopy, Book Where BookCopy.BookID = '" & BID.Text & "' And BookCopy.CopyN = '" & CPID.Text & "' And BookCopy.BookID = Book.BookID "
rs2.Open BOOKFIND, cn, adOpenKeyset, adLockOptimistic
Set MSHFlexGrid3.DataSource = rs2
rs2.Close
If MSHFlexGrid3.Row = 0 Then
MsgBox "查無此書."
BID.Text = ""
CPID.Text = ""
End If
If Text1(9).Text = "0" Then
MsgBox "這位會員已經達到可借出書本的上限,無法借出更多的書本."
End If
If MSHFlexGrid3.Row <> 0 And Text1(9).Text <> "0" Then
    If MSHFlexGrid3.TextMatrix(1, 4) = "B" Then
        MsgBox "此書已被借出."
        BID.Text = ""
        CPID.Text = ""
    End If
    If MSHFlexGrid3.TextMatrix(1, 4) = "R" Then
        For ss = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(ss, 1) = MSHFlexGrid3.TextMatrix(1, 1) And MSHFlexGrid1.TextMatrix(ss, 2) = MSHFlexGrid3.TextMatrix(1, 2) Then
            MsgBox "此書已被加入訂單中!"
            CHECKHAVE = True
            Exit For
            End If
        Next ss
            If CHECKHAVE = False Then
            MSHFlexGrid1.AddItem "" & vbTab & MSHFlexGrid3.TextMatrix(1, 1) & vbTab & MSHFlexGrid3.TextMatrix(1, 2) & vbTab & MSHFlexGrid3.TextMatrix(1, 3) & vbTab & "8"
            A = CInt(PAYMENT.Text) + 8
            B = CInt(Text1(9).Text) - 1
            Text1(9).Text = CStr(B)
            PAYMENT.Text = CStr(A)
            End If
    End If
BID.Text = ""
CPID.Text = ""
Call CheckDate
End If
End Sub
Function CheckDate()
If MSHFlexGrid1.Rows < 6 Then
Due.Text = Format(Now + 3, "dd/mm/yyyy")
End If
If MSHFlexGrid1.Rows > 6 Then
Due.Text = Format(Now + 4, "dd/mm/yyyy")
End If
If MSHFlexGrid1.Rows > 11 Then
Due.Text = Format(Now + 5, "dd/mm/yyyy")
End If
If MSHFlexGrid1.Rows > 16 Then
Due.Text = Format(Now + 6, "dd/mm/yyyy")
End If
If MSHFlexGrid1.Rows = 21 Then
Due.Text = Format(Now + 7, "dd/mm/yyyy")
End If
End Function

Private Sub CLEARLIST_Click()
Dim i As Integer
For i = 0 To REMEID.Count - 1
REMEID(i).Text = ""
Next i
RETURNFEE.Enabled = False

End Sub

Private Sub CM_Click()
Dim A As Integer
MSHFlexGrid1.Rows = 2
rs2.Open "Select * From Member Where MemberID = '" & ID.Text & "'", cn, adOpenKeyset, adLockOptimistic
Set MSHFlexGrid3.DataSource = rs2
rs2.Close
If MSHFlexGrid3.Row = 0 Then
    MsgBox "你所輸入的會員編號錯誤!"
End If
If MSHFlexGrid3.Row <> 0 Then
    For A = 0 To 9
    Text1(A).Text = MSHFlexGrid3.TextMatrix(1, A + 1)
    Text1(A).Locked = True
    Next A
End If
If CInt(Text1(8).Text) <> 0 Then
    Command5.Enabled = True
    MsgBox "這位會員有罰款未還,需要還清罰款才能借書喔!"
End If
If CInt(Text1(9).Text) = 0 Then
    MsgBox "這位會員借出的書本已滿,需要歸還書本才能借書喔!"
End If
If CInt(Text1(8).Text) = 0 Or CInt(Text1(9).Text) <> 0 Then
  CPID.Enabled = True
  BID.Enabled = True
  CB.Enabled = True
End If
End Sub

Private Sub Command3_Click()
StaffUse.Show
COMICS.Hide
cn.Close
End Sub

Private Sub Command4_Click()
Dim ss As Integer
Dim BB As Integer
Dim LID As Integer
Dim LS As String
LID = 1
cn.Execute ("insert into BorrowOrder(BOID,MemberID,BODate,TotalFee) values ('" & BOID.Text & "','" & Text1(0).Text & "','" & NowDate.Text & "','" & PAYMENT.Text & "')")
For ss = 2 To MSHFlexGrid1.Rows - 1
Call CreateBOL(LID, LS)
cn.Execute ("insert into BorrowOrderList(BOID,BOLN,BookID,CopyN,BorrowDate,DueDate,Fee,status) values ('" & BOID.Text & "','" & LS & "','" & MSHFlexGrid1.TextMatrix(ss, 1) & "','" & MSHFlexGrid1.TextMatrix(ss, 2) & "','" & NowDate.Text & "','" & Due.Text & "','" & MSHFlexGrid1.TextMatrix(ss, 4) & "', 'B')")
cn.Execute ("Update BookCopy SET Status = 'B' where BookID = '" & MSHFlexGrid1.TextMatrix(ss, 1) & "' And CopyN = '" & MSHFlexGrid1.TextMatrix(ss, 2) & "' ")
LID = LID + 1
Next ss
cn.Execute ("Update Member SET BorrowNumber = '" & Text1(9).Text & "' where MemberID = '" & Text1(0).Text & "'")
MsgBox "多謝惠顧!"
Open "d:\" & BOID.Text & ".txt" For Output As #1
Print #1, "=================================漫畫屋====================================="
Print #1, "銅鑼灣謝斐道490-496號金利文廣場7樓 電話:2295 6792 營業時間:下午1時至晚上11時"
Print #1, "借閱訂單編號:", BOID.Text, "        ", "會員編號:", Text1(0).Text
Print #1, "----------------------------------------------------------------------------"
Print #1, "書籍借出日期:", NowDate.Text, "--------------------還書日期:", Due.Text
Print #1, MSHFlexGrid1.TextMatrix(0, 1), MSHFlexGrid1.TextMatrix(0, 2), MSHFlexGrid1.TextMatrix(0, 3), MSHFlexGrid1.TextMatrix(0, 4)
For BB = 2 To MSHFlexGrid1.Rows - 1
Print #1, MSHFlexGrid1.TextMatrix(BB, 1), MSHFlexGrid1.TextMatrix(BB, 2), MSHFlexGrid1.TextMatrix(BB, 3), MSHFlexGrid1.TextMatrix(BB, 4)
Next BB
Print #1, " "
Print #1, " "
Print #1, "============================================================================"
Print #1, "                                                       總價格:$", PAYMENT.Text
Print #1, "--------------------------------多謝惠顧!-----------------------------------"
Close #1
MSHFlexGrid1.Rows = 2
PAYMENT.Text = "0"
For BB = 0 To Text1.Count - 1
Text1(BB).Text = " "
Next BB
CPID.Enabled = False
BID.Enabled = False
CB.Enabled = False
Call CheckDate
End Sub

Private Sub Command5_Click()
Dim A As String
A = ""
cn.Execute ("Update Member SET Fee = 0 where MemberID = '" & Text1(0).Text & "'")
Text1(8).Text = "0"
If Text1(9).Text <> "0" And Text1(8).Text = "0" Then
cn.Execute ("Update Member SET Status = 'Y' where MemberID = '" & Text1(0).Text & "'")
Text1(7).Text = "Y"
Call CreateEID(A)
cn.Execute ("insert into ExtraFee(EID,EType,MemberID,Payment,PaymentDate) values ('" & A & "','2','" & Text1(0).Text & "','" & Text1(8).Text & "','" & NowDate.Text & "')")
MsgBox "罰款已付!該會員可以借書喔"
Command5.Enabled = False
Else
MsgBox "罰款已付!可是會員的借書數量到達上限,無法使用借書服務"
Command5.Enabled = False
End If
End Sub

Private Sub Command8_Click()
If ORDERLIST.Row = 0 Or ORDERLIST.Row = 1 Then
MsgBox "此列不能刪除!"
Else
BPTOTALSELL.Text = CStr(CInt(BPTOTALSELL.Text) - CInt(ORDERLIST.TextMatrix(ORDERLIST.Row, 5)))
ORDERLIST.RemoveItem (ORDERLIST.Row)
MsgBox "此列己刪除!"
End If
End Sub

Private Sub DELBO_Click()
If MSHFlexGrid1.Row = 0 Or MSHFlexGrid1.Row = 1 Then
MsgBox "此列不能刪除!"
Else
MSHFlexGrid1.RemoveItem (MSHFlexGrid1.Row)
Text1(9).Text = CInt(Text1(9).Text) + 1
PAYMENT.Text = CInt(PAYMENT.Text) - 8
MsgBox "此列己刪除!"
End If
End Sub

Private Sub FINISHORDER_Click()
Dim ONN As Integer
Dim LIDS As Integer
Dim LS As String
If ORDERLIST.Rows = 2 Then
MsgBox "這位客人還沒有在清單加入想要購買的東西喔!"
Else
LS = ""
LIDS = 1
cn.Execute ("insert into SellOrder(OID,TotalPrice,OrderDate) values ('" & OID.Text & "','" & BPTOTALSELL.Text & "','" & NowDate3.Text & "')")
For ONN = 2 To ORDERLIST.Rows - 1
Call CreateBOL(LIDS, LS)
If Left(ORDERLIST.TextMatrix(ONN, 1), 1) = "B" Then
cn.Execute ("insert into BookOrderList(OID,OLN,BookID,Price,OrderN,TotalPrice) values ('" & OID.Text & "','" & LS & "','" & ORDERLIST.TextMatrix(ONN, 1) & "','" & ORDERLIST.TextMatrix(ONN, 3) & "','" & ORDERLIST.TextMatrix(ONN, 4) & "','" & ORDERLIST.TextMatrix(ONN, 5) & "')")
End If
If Left(ORDERLIST.TextMatrix(ONN, 1), 1) = "P" Then
cn.Execute ("insert into ProductOrderList(OID,OLN,ProductID,Price,OrderN,TotalPrice) values ('" & OID.Text & "','" & LS & "','" & ORDERLIST.TextMatrix(ONN, 1) & "','" & ORDERLIST.TextMatrix(ONN, 3) & "','" & ORDERLIST.TextMatrix(ONN, 4) & "','" & ORDERLIST.TextMatrix(ONN, 5) & "')")
End If
LID = LID + 1
Next ONN
MsgBox "多謝惠顧!"
Open "d:\" & OID.Text & ".txt" For Output As #1
Print #1, "=================================漫畫屋====================================="
Print #1, "銅鑼灣謝斐道490-496號金利文廣場7樓 電話:2295 6792 營業時間:下午1時至晚上11時"
Print #1, "訂單編號:", OID.Text
Print #1, "----------------------------------------------------------------------------"
Print #1, ORDERLIST.TextMatrix(0, 1), ORDERLIST.TextMatrix(0, 2), ORDERLIST.TextMatrix(0, 3), ORDERLIST.TextMatrix(0, 4), ORDERLIST.TextMatrix(0, 5)
For ONN = 2 To ORDERLIST.Rows - 1
Print #1, ORDERLIST.TextMatrix(ONN, 1), "      ", ORDERLIST.TextMatrix(ONN, 2), "      ", ORDERLIST.TextMatrix(ONN, 3), "      ", ORDERLIST.TextMatrix(ONN, 4), "      ", ORDERLIST.TextMatrix(ONN, 5)
Next ONN
Print #1, " "
Print #1, " "
Print #1, "============================================================================"
Print #1, "                                                       總價格:$", BPTOTALSELL.Text
Print #1, "--------------------------------多謝惠顧!-----------------------------------"
Close #1
ORDERLIST.Rows = 2
Call CreateOID
BPTOTALSELL.Text = "0"
End If
End Sub

Private Sub Form_Activate()
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "Member", cn, adOpenKeyset, adLockOptimistic
  NowDate.Locked = True
  NowDate.BackColor = vbBlue
  NowDate.ForeColor = vbYellow
  Due.BackColor = vbRed
  Due.ForeColor = vbYellow
  NowDate.Text = Format(Now, "dd/mm/yyyy")
  NowDate2.Text = Format(Now, "dd/mm/yyyy")
  NowDate3.Text = NowDate2.Text
  Due.Text = Format(Date + 3, "dd/mm/yyyy")
  CPID.Enabled = False
  BID.Enabled = False
  CB.Enabled = False
  MSHFlexGrid1.ColWidth(0) = 150 * 3
  MSHFlexGrid1.ColWidth(1) = 500 * 3
  MSHFlexGrid1.ColWidth(2) = 700 * 3
  MSHFlexGrid1.TextMatrix(0, 1) = "書籍編號"
  MSHFlexGrid1.TextMatrix(0, 2) = "書籍複本編號"
  MSHFlexGrid1.TextMatrix(0, 3) = "書籍名稱"
  MSHFlexGrid1.TextMatrix(0, 4) = "租借費用"
  ORDERLIST.ColWidth(0) = 150 * 3
  ORDERLIST.ColWidth(1) = 500 * 3
  ORDERLIST.ColWidth(2) = 600 * 3
  ORDERLIST.TextMatrix(0, 1) = "書籍/貨品編號"
  ORDERLIST.TextMatrix(0, 2) = "書籍/貨品名稱"
  ORDERLIST.TextMatrix(0, 3) = "單價"
  ORDERLIST.TextMatrix(0, 4) = "數量"
  ORDERLIST.TextMatrix(0, 5) = "總值"
  MEMBERBOOKLIST.ColWidth(0) = 150 * 3
  Call CreateBOP
  Call CreateOID
  Command5.Enabled = False
  RETURNFEE.Enabled = False
End Sub
Function CreateOID()
Dim newID As String
Dim CID As Integer
Dim fail As Integer
rs4.Open "SellOrder", cn, adOpenKeyset, adLockOptimistic
If rs4.BOF = True Then
CID = 1
Else
rs4.MoveLast
newID = rs4.Fields(0)
CID = CInt(Right(rs4.Fields(0), 7)) + 1
End If
fail = 8 - Len(CID)
newID = ""
newID = Trim(newID + "O" + String(fail, "0") + CStr(CID))
OID.Text = Trim(newID)
rs4.Close
End Function
Function CreateBOP()
Dim newID As String
Dim CID As Integer
Dim fail As Integer
rs4.Open "BorrowOrder", cn, adOpenKeyset, adLockOptimistic
If rs4.BOF = True Then
CID = 1
Else
rs4.MoveLast
newID = rs4.Fields(0)
CID = CInt(Right(rs4.Fields(0), 6)) + 1
End If
fail = 6 - Len(CID)
newID = ""
newID = Trim(newID + "BO" + String(fail, "0") + CStr(CID))
BOID.Text = Trim(newID)
rs4.Close
End Function
Function CreateBOL(A As Integer, B As String)
Dim fail As Integer
fail = 2 - Len(CStr(A))
B = ""
B = Trim(B + "L" + String(fail, "0") + CStr(A))
End Function
Function CreateEID(NE As String)
Dim newID As String
Dim CID As Integer
Dim fail As Integer
rs4.Open "ExtraFee", cn, adOpenKeyset, adLockOptimistic
If rs4.BOF = True Then
CID = 1
Else
rs4.MoveLast
newID = rs4.Fields(0)
CID = CInt(Right(rs4.Fields(0), 2)) + 1
End If
fail = 8 - Len(CID)
newID = ""
newID = Trim(newID + "E" + String(fail, "0") + CStr(CID))
NE = Trim(newID)
rs4.Close
End Function
Private Sub RETURNBP_Click()
Dim RETURNING As String
Dim CHECKDATA As Long
Dim TRUEDATA As Long
Dim PCOUNT As Long
RETURNING = "Select BorrowOrderList.BookID, BorrowOrderList.CopyN, BorrowOrderList.BorrowDate, BorrowOrderList.DueDate, BorrowOrderList.BOID, BorrowOrder.MemberID ,Book.BookName1 From BorrowOrderList, Book, BorrowOrder where BorrowOrderList.BookID = '" & BID2.Text & "' and Book.BookID = BorrowOrderList.BookID And BorrowOrderList.CopyN = '" & CPID2.Text & "' And BorrowOrderList.status = 'B' and BorrowOrderList.BOID = BorrowOrder.BOID"
rs4.Open RETURNING, cn, adOpenKeyset, adLockOptimistic
Set SEARCHLIST.DataSource = rs4
rs4.Close
If SEARCHLIST.Row = 0 Then
MsgBox "查無此書."
Else
rs4.Open "Select * from Member where MemberID = '" & SEARCHLIST.TextMatrix(1, 6) & "'", cn, adOpenKeyset, adLockOptimistic
Set SEARCHMEMBER.DataSource = rs4
rs4.Close
FROMBO.Text = SEARCHLIST.TextMatrix(1, 5)
For EE = 0 To REMEID.Count - 1
REMEID(EE).Text = SEARCHMEMBER.TextMatrix(1, EE + 1)
REMEID(EE).Locked = True
Next EE
CHECKDATA = DateDiff("d", SEARCHLIST.TextMatrix(1, 3), SEARCHLIST.TextMatrix(1, 4))
TRUEDATA = DateDiff("d", SEARCHLIST.TextMatrix(1, 3), NowDate2.Text)
PCOUNT = TRUEDATA - CHECKDATA
If PCOUNT > 0 Then
REMEID(8).Text = CStr(CInt(REMEID(8).Text + PCOUNT * 3))
MsgBox "此書已過期!請歸還罰款!"
cn.Execute ("UPDATE Member SET Fee =  Fee + '" & CInt(REMEID(9).Text) & "'  where MemberID ='" & SEARCHLIST.TextMatrix(1, 6) & "'")
RETURNFEE.Enabled = False
Else
End If
cn.Execute ("UPDATE BorrowOrderList SET status = 'R' where BookID = '" & SEARCHLIST.TextMatrix(1, 1) & "' and CopyN = '" & SEARCHLIST.TextMatrix(1, 2) & "' and status = 'B'")
cn.Execute ("UPDATE BookCopy SET Status = 'R' where BookID ='" & SEARCHLIST.TextMatrix(1, 1) & "' and CopyN = '" & SEARCHLIST.TextMatrix(1, 2) & "' and status = 'B'")
cn.Execute ("UPDATE Member SET BorrowNumber = BorrowNumber + 1 where MemberID = '" & SEARCHLIST.TextMatrix(1, 6) & "'")
REMEID(9).Text = CStr(CInt(REMEID(9).Text + 1))
rs4.Open "Select Book.BookID, BookCopy.CopyN, Book.BookName1, BorrowOrderList.BorrowDate, BorrowOrderList.DueDate FROM Book, BookCopy, BorrowOrder, BorrowOrderList,Member WHERE BorrowOrderList.BOID = BorrowOrder.BOID AND Book.BookID = BookCopy.BookID AND BookCopy.BookID = BorrowOrderList.BookID AND BookCopy.CopyN = BorrowOrderList.CopyN AND BorrowOrder.MemberID = Member.MemberID AND BorrowOrderList.status = 'B' AND Member.MemberID = '" & SEARCHLIST.TextMatrix(1, 6) & "'", cn, adOpenKeyset, adLockOptimistic
Set MEMBERBOOKLIST.DataSource = rs4
rs4.Close
MsgBox "還書成功!"
End If
BID2.Text = ""
CPID2.Text = ""
End Sub

Private Sub RETURNFEE_Click()
cn.Execute ("Update Member SET Fee = 0 where MemberID = '" & REMEID(0).Text & "'")
REMEID(8).Text = "0"
If REMEID(9).Text <> "0" And REMEID(8).Text = "0" Then
cn.Execute ("Update Member SET Status = 'Y' where MemberID = '" & REMEID(0).Text & "'")
Text1(7).Text = "Y"
Call CreateEID(A)
cn.Execute ("insert into ExtraFee(EID,EType,MemberID,Payment,PaymentDate) values ('" & A & "','2','" & REMEID(0).Text & "','" & REMEID(8).Text & "','" & NowDate.Text & "')")
MsgBox "罰款已付!該會員可以借書喔"
RETURNFEE.Enabled = False
End If
End Sub



Private Sub TAKEORDER_Click()
Dim i As Integer
Dim CHECKHAVE As Boolean
CHECKHAVE = False
If Left(INPUTBP.Text, 1) = "B" Then
rs4.Open "Select BookID, BookName1, Price From Book where BookID = '" & INPUTBP.Text & "' ", cn, adOpenKeyset, adLockOptimistic
Set FINDBP.DataSource = rs4
rs4.Close
    If FINDBP.Row = 0 Then
        MsgBox "查無此書!"
    End If
        If FINDBP.Row <> 0 Then
            For i = 1 To ORDERLIST.Rows - 1
            If ORDERLIST.TextMatrix(i, 1) = FINDBP.TextMatrix(1, 1) Then
                ORDERLIST.TextMatrix(i, 4) = CInt(ORDERLIST.TextMatrix(i, 4)) + 1
                ORDERLIST.TextMatrix(i, 5) = CInt(ORDERLIST.TextMatrix(i, 5)) + CInt(FINDBP.TextMatrix(1, 3))
                CHECKHAVE = True
                BPTOTALSELL.Text = CInt(BPTOTALSELL.Text) + CInt(FINDBP.TextMatrix(1, 3))
                Exit For
                End If
            Next i
            If CHECKHAVE = False Then
            ORDERLIST.AddItem "->" & vbTab & FINDBP.TextMatrix(1, 1) & vbTab & FINDBP.TextMatrix(1, 2) & vbTab & FINDBP.TextMatrix(1, 3) & vbTab & "1" & vbTab & FINDBP.TextMatrix(1, 3)
            BPTOTALSELL.Text = CInt(BPTOTALSELL.Text) + CInt(FINDBP.TextMatrix(1, 3))
            End If
        End If
End If
If Left(INPUTBP.Text, 1) = "P" Then
rs4.Open "Select ProductID, ProductName, Price From Product where ProductID = '" & INPUTBP.Text & "' ", cn, adOpenKeyset, adLockOptimistic
Set FINDBP.DataSource = rs4
rs4.Close
    If FINDBP.Row = 0 Then
        MsgBox "查無此貨品!"
    End If
        If FINDBP.Row <> 0 Then
            For i = 1 To ORDERLIST.Rows - 1
            If ORDERLIST.TextMatrix(i, 1) = FINDBP.TextMatrix(1, 1) Then
                ORDERLIST.TextMatrix(i, 4) = CInt(ORDERLIST.TextMatrix(i, 4)) + 1
                ORDERLIST.TextMatrix(i, 5) = CInt(ORDERLIST.TextMatrix(i, 5)) + CInt(FINDBP.TextMatrix(1, 3))
                CHECKHAVE = True
                BPTOTALSELL.Text = CInt(BPTOTALSELL.Text) + CInt(FINDBP.TextMatrix(1, 3))
                Exit For
                End If
            Next i
            If CHECKHAVE = False Then
            ORDERLIST.AddItem "->" & vbTab & FINDBP.TextMatrix(1, 1) & vbTab & FINDBP.TextMatrix(1, 2) & vbTab & FINDBP.TextMatrix(1, 3) & vbTab & "1" & vbTab & FINDBP.TextMatrix(1, 3)
            BPTOTALSELL.Text = CInt(BPTOTALSELL.Text) + CInt(FINDBP.TextMatrix(1, 3))
            End If
        End If
End If
If Left(INPUTBP.Text, 1) <> "B" And Left(INPUTBP.Text, 1) <> "P" Then
MsgBox "輸入錯誤!"
End If
End Sub

Private Sub Timer2_Timer()
StatusBar1.Panels(1).Text = "時間:" + Format(Now, "Medium Time")
End Sub
