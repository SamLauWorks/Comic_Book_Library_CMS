VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PBView 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10365
   ClientLeft      =   2565
   ClientTop       =   2325
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   13395
   StartUpPosition =   2  '螢幕中央
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   615
      Left            =   8520
      Top             =   9600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\vb6\FYP(YEAR4)\BookD.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\vb6\FYP(YEAR4)\BookD.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"PBView.frx":0000
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   615
      Left            =   0
      Top             =   9480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\vb6\FYP(YEAR4)\BookD.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\vb6\FYP(YEAR4)\BookD.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"PBView.frx":0101
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   2040
      Top             =   9480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\vb6\FYP(YEAR4)\BookD.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\vb6\FYP(YEAR4)\BookD.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"PBView.frx":0197
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   10800
      Top             =   9600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\vb6\FYP(YEAR4)\BookD.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\vb6\FYP(YEAR4)\BookD.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"PBView.frx":0221
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回到主菜單"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   1
      Top             =   9480
      Width           =   3135
   End
   Begin TabDlg.SSTab PBList 
      Height          =   9015
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   794
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "書籍資訊"
      TabPicture(0)   =   "PBView.frx":03C6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "書籍總集資訊"
      TabPicture(1)   =   "PBView.frx":03E2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "貨品資訊"
      TabPicture(2)   =   "PBView.frx":03FE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "書籍複本資訊"
      TabPicture(3)   =   "PBView.frx":041A
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "書籍資訊:(*:該資料項目必須輸入(格子變成白色時才能輸入資料))"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Left            =   -74880
         TabIndex        =   25
         Top             =   600
         Width           =   13095
         Begin MSDataGridLib.DataGrid BOOKDG 
            Bindings        =   "PBView.frx":0436
            Height          =   3015
            Left            =   120
            TabIndex        =   61
            Top             =   5040
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   5318
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   29
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3076
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3076
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame6 
            Caption         =   "書籍資訊控制項:(或點選下方的列表指標瀏覽書籍資訊)"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   60
            Top             =   2880
            Width           =   12855
            Begin VB.CommandButton MoveOne 
               Caption         =   "第一項書籍"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   70
               Top             =   1200
               Width           =   2055
            End
            Begin VB.CommandButton MovePerv 
               Caption         =   "上一項書籍"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2400
               TabIndex        =   69
               Top             =   1200
               Width           =   2175
            End
            Begin VB.CommandButton MoveNext 
               Caption         =   "下一項書籍"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4800
               TabIndex        =   68
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton MoveLast 
               Caption         =   "最後一項書籍"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   7200
               TabIndex        =   67
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton DELETEBOOK 
               Caption         =   "刪除書籍"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   7800
               TabIndex        =   66
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton ADDBOOK 
               Caption         =   "新增書籍"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   65
               Top             =   360
               Width           =   2055
            End
            Begin VB.CommandButton EDITBOOK 
               Caption         =   "修改書籍"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2400
               TabIndex        =   64
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton CANCEL1 
               Caption         =   "取消"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   6360
               TabIndex        =   63
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton SAVEBOOK 
               Caption         =   "儲存"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4800
               TabIndex        =   62
               Top             =   360
               Width           =   1455
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid BCOPYFIND 
               Height          =   615
               Left            =   10200
               TabIndex        =   85
               Top             =   1200
               Visible         =   0   'False
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   1085
               _Version        =   393216
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label BOOKCOUNT 
               BackColor       =   &H000000FF&
               Height          =   375
               Left            =   10320
               TabIndex        =   71
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   0
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "X"
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   1
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   2
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   3
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   4
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox Text1 
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
            Index           =   5
            Left            =   9960
            TabIndex        =   39
            Top             =   7440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
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
            Index           =   6
            Left            =   10680
            TabIndex        =   38
            Top             =   7440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text1 
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
            Index           =   7
            Left            =   11280
            TabIndex        =   37
            Top             =   6120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   8
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1800
            Width           =   4095
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   9
            Left            =   7440
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   35
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Text1 
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
            Index           =   11
            Left            =   11880
            TabIndex        =   33
            Top             =   5640
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   12
            Left            =   10080
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   32
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000A&
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
            Index           =   13
            Left            =   10800
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   31
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox BN 
            BackColor       =   &H8000000A&
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   2280
            Width           =   2295
         End
         Begin VB.ComboBox B5 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "PBView.frx":044B
            Left            =   8760
            List            =   "PBView.frx":0458
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "B5"
            Top             =   840
            Width           =   1815
         End
         Begin VB.ComboBox B6 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "PBView.frx":046E
            Left            =   2160
            List            =   "PBView.frx":049C
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "B6"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox B7 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "PBView.frx":04F4
            Left            =   5520
            List            =   "PBView.frx":0522
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "B7"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.ComboBox B12 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "標楷體"
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
            TabIndex        =   26
            Text            =   "B12"
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*書籍編號:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*書籍名稱:"
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
            Index           =   1
            Left            =   3960
            TabIndex        =   58
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "書籍別名:"
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
            Index           =   2
            Left            =   8520
            TabIndex        =   57
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*作者1:"
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
            Index           =   3
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "作者2:"
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
            Index           =   4
            Left            =   3720
            TabIndex        =   55
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*書籍類別:"
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
            Index           =   5
            Left            =   6960
            TabIndex        =   54
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*內容類別1:"
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
            Index           =   6
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "內容類別2:"
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
            Index           =   7
            Left            =   3720
            TabIndex        =   52
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*出版社:"
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
            Index           =   8
            Left            =   120
            TabIndex        =   51
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*出版年份:"
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
            Index           =   9
            Left            =   5640
            TabIndex        =   50
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*價錢:"
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
            Index           =   10
            Left            =   7200
            TabIndex        =   49
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "書籍總集編號:"
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
            Index           =   11
            Left            =   120
            TabIndex        =   48
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "集數:"
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
            Index           =   12
            Left            =   9000
            TabIndex        =   47
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "*建立日期:"
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
            Index           =   13
            Left            =   9000
            TabIndex        =   46
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFF80&
            Caption         =   "書籍總集名稱:"
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
            Left            =   4200
            TabIndex        =   45
            Top             =   2280
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "書籍複本資訊:(*:該資料項目必須輸入(格子變成白色時才能輸入資料)"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   13095
         Begin VB.TextBox BPT 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   12120
            TabIndex        =   121
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox BCOPY 
            BackColor       =   &H8000000C&
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
            Index           =   2
            Left            =   2760
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   119
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox BCOPY 
            BackColor       =   &H8000000C&
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
            Index           =   1
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   118
            Top             =   480
            Width           =   6735
         End
         Begin VB.Frame Frame4 
            Caption         =   "書籍資訊控制項:(或點選下方的列表指標瀏覽書籍資訊)"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   106
            Top             =   1680
            Width           =   12855
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid BCOPYCOUNT 
               Height          =   1215
               Left            =   10440
               TabIndex        =   122
               Top             =   720
               Visible         =   0   'False
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   2143
               _Version        =   393216
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.TextBox BCOPY 
               BackColor       =   &H8000000C&
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
               Index           =   3
               Left            =   9600
               Locked          =   -1  'True
               TabIndex        =   120
               Top             =   1320
               Width           =   975
            End
            Begin VB.CommandButton SAVEBCOPY 
               Caption         =   "儲存"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4800
               TabIndex        =   115
               Top             =   360
               Width           =   1455
            End
            Begin VB.CommandButton CANCEL4 
               Caption         =   "取消"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   6360
               TabIndex        =   114
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton EDITBCOPY 
               Caption         =   "修改書籍複本"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2400
               TabIndex        =   113
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton ADDBCOPY 
               Caption         =   "新增書籍複本"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   112
               Top             =   360
               Width           =   2055
            End
            Begin VB.CommandButton DELETEBCOPY 
               Caption         =   "刪除書籍複本"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   7800
               TabIndex        =   111
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton MoveLast4 
               Caption         =   "最後一項書籍複本"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   7200
               TabIndex        =   110
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton MoveNext4 
               Caption         =   "下一項書籍複本"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4800
               TabIndex        =   109
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton MovePerv4 
               Caption         =   "上一項書籍複本"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2400
               TabIndex        =   108
               Top             =   1200
               Width           =   2175
            End
            Begin VB.CommandButton MoveOne4 
               Caption         =   "第一項書籍複本"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   107
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label BOOKCOUNT4 
               BackColor       =   &H000000FF&
               Height          =   375
               Left            =   10080
               TabIndex        =   116
               Top             =   240
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin VB.TextBox BCOPY 
            BackColor       =   &H8000000C&
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
            Index           =   0
            Left            =   2760
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   100
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox BC2 
            BackColor       =   &H8000000C&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "PBView.frx":057A
            Left            =   6960
            List            =   "PBView.frx":0584
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   1080
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid BCOPYDG 
            Bindings        =   "PBView.frx":0594
            Height          =   4215
            Left            =   120
            TabIndex        =   117
            Top             =   3840
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   7435
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   29
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3076
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3076
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000004&
            Caption         =   "B:借出 R:架上"
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
            Left            =   8160
            TabIndex        =   123
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000FFFF&
            Caption         =   "書籍名稱:"
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
            Index           =   14
            Left            =   4440
            TabIndex        =   105
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label8 
            BackColor       =   &H0000FFFF&
            Caption         =   "*書籍複本編號:"
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
            Left            =   240
            TabIndex        =   104
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000FFFF&
            Caption         =   "*書籍編號:"
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
            Left            =   240
            TabIndex        =   103
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label7 
            BackColor       =   &H0000FFFF&
            Caption         =   "該書籍複本總數:"
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
            Left            =   9480
            TabIndex        =   102
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000FFFF&
            Caption         =   "*書籍複本狀態:"
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
            Left            =   4440
            TabIndex        =   101
            Top             =   1080
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "貨品資訊:(*:該資料項目必須輸入(格子變成白色時才能輸入資料)"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Left            =   -74880
         TabIndex        =   13
         Top             =   540
         Width           =   13095
         Begin VB.Frame Frame9 
            Caption         =   "書籍資訊控制項:(或點選下方的列表指標瀏覽借閱訂單的項列)"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   86
            Top             =   1560
            Width           =   12855
            Begin VB.TextBox Text3 
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
               Index           =   3
               Left            =   9840
               TabIndex        =   98
               Top             =   1320
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.CommandButton SAVEP 
               Caption         =   "儲存"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4800
               TabIndex        =   95
               Top             =   360
               Width           =   1455
            End
            Begin VB.CommandButton CANCEL3 
               Caption         =   "取消"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   6360
               TabIndex        =   94
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton EDITP 
               Caption         =   "修改貨品"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2400
               TabIndex        =   93
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton ADDP 
               Caption         =   "新增貨品"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   92
               Top             =   360
               Width           =   2055
            End
            Begin VB.CommandButton DELETEP 
               Caption         =   "刪除貨品"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   7800
               TabIndex        =   91
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton MoveLast3 
               Caption         =   "最後一項書籍資訊"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   7200
               TabIndex        =   90
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton MoveNext3 
               Caption         =   "下一項的貨品"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4800
               TabIndex        =   89
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton MovePerv3 
               Caption         =   "上一項貨品"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2400
               TabIndex        =   88
               Top             =   1200
               Width           =   2175
            End
            Begin VB.CommandButton MoveOne3 
               Caption         =   "第一項貨品"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   87
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label BOOKCOUNT3 
               BackColor       =   &H000000FF&
               Height          =   375
               Left            =   10320
               TabIndex        =   96
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin VB.ComboBox P1 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "PBView.frx":05A9
            Left            =   6120
            List            =   "PBView.frx":05B6
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000A&
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
            Index           =   1
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   480
            Width           =   6015
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000A&
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
            Index           =   0
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000A&
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
            Index           =   2
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   15
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000A&
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
            Index           =   4
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   960
            Width           =   2295
         End
         Begin MSDataGridLib.DataGrid PDG 
            Bindings        =   "PBView.frx":05D4
            Height          =   4335
            Left            =   120
            TabIndex        =   97
            Top             =   3840
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   7646
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   29
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3076
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3076
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label11 
            BackColor       =   &H0080C0FF&
            Caption         =   "*貨品名稱:"
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
            Left            =   4320
            TabIndex        =   2
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackColor       =   &H0080C0FF&
            Caption         =   "*貨品編號:"
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
            TabIndex        =   21
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0080C0FF&
            Caption         =   "*價錢:"
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
            TabIndex        =   20
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackColor       =   &H0080C0FF&
            Caption         =   "*貨品類別:"
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
            Left            =   4320
            TabIndex        =   19
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackColor       =   &H0080C0FF&
            Caption         =   "*新增日期:"
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
            Left            =   8040
            TabIndex        =   18
            Top             =   960
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "書籍總集資訊:(*:該資料項目必須輸入(格子變成白色時才能輸入資料)"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8055
         Left            =   -74880
         TabIndex        =   4
         Top             =   540
         Width           =   13095
         Begin VB.Frame Frame8 
            Caption         =   "書籍總集資訊控制項:(或點選左方的列表指標瀏覽書籍總集)"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   2880
            TabIndex        =   74
            Top             =   1560
            Width           =   9855
            Begin VB.CommandButton SAVEBC 
               Caption         =   "儲存"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4440
               TabIndex        =   83
               Top             =   360
               Width           =   1455
            End
            Begin VB.CommandButton CANCEL2 
               Caption         =   "取消"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   6000
               TabIndex        =   82
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton EDITBC 
               Caption         =   "修改書籍資訊"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2160
               TabIndex        =   81
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton ADDBC 
               Caption         =   "新增書籍總集"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   80
               Top             =   360
               Width           =   1815
            End
            Begin VB.CommandButton DELETEBC 
               Caption         =   "刪除書籍總集"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   7440
               TabIndex        =   79
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton MoveLast2 
               Caption         =   "最後一項書籍總集"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   6960
               TabIndex        =   78
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton MoveNext2 
               Caption         =   "下一項書籍總集"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4560
               TabIndex        =   77
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton MovePerv2 
               Caption         =   "上一項書籍總集"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2280
               TabIndex        =   76
               Top             =   1200
               Width           =   2175
            End
            Begin VB.CommandButton MoveOne2 
               Caption         =   "第一項書籍總集"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   12
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   75
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label BOOKCOUNT2 
               BackColor       =   &H000000FF&
               Height          =   375
               Left            =   9360
               TabIndex        =   84
               Top             =   1320
               Visible         =   0   'False
               Width           =   375
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "借閱訂單列表:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6015
            Left            =   120
            TabIndex        =   72
            Top             =   480
            Width           =   2655
            Begin MSDataGridLib.DataGrid BCDG 
               Bindings        =   "PBView.frx":05E9
               Height          =   5175
               Left            =   120
               TabIndex        =   73
               Top             =   480
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   9128
               _Version        =   393216
               AllowUpdate     =   0   'False
               HeadLines       =   1
               RowHeight       =   29
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "標楷體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "標楷體"
                  Size            =   18
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   3076
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   3076
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.ComboBox BC3 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "PBView.frx":05FE
            Left            =   10560
            List            =   "PBView.frx":0608
            TabIndex        =   22
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
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
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Text            =   "Text2"
            Top             =   6600
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
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
            Index           =   1
            Left            =   5520
            TabIndex        =   6
            Top             =   1080
            Width           =   6495
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
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
            Index           =   0
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   480
            Width           =   2415
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid BookList 
            Height          =   3735
            Left            =   3000
            TabIndex        =   8
            Top             =   4080
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   6588
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label4 
            Caption         =   "書籍總集內的書籍列表:(該資料變更時,其他書籍資料會同步更新)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   12
            Top             =   3720
            Width           =   9615
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000FF00&
            Caption         =   "*書籍總集狀態:"
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
            Index           =   2
            Left            =   8040
            TabIndex        =   11
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000FF00&
            Caption         =   "*書籍總集名稱:"
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
            Index           =   1
            Left            =   3000
            TabIndex        =   10
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000FF00&
            Caption         =   "*書籍總集編號:"
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
            Index           =   0
            Left            =   3000
            TabIndex        =   9
            Top             =   480
            Width           =   2535
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "如果需要查詢或變更其他類別的資料,請按到指定的書籤目錄(例如下方的""貨品資訊"")"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "PBView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub ADDBC_Click()
  For i = 1 To 2
  Text2(i).Text = ""
  Text2(i).Locked = False
  Text2(i).BackColor = vbWhite
  Next i
  Text2(0).Text = "自動生成"
  BC3.Text = ""
  BC3.Locked = False
  BC3.BackColor = vbWhite
  MoveOne2.Enabled = False
  MoveNext2.Enabled = False
  MovePerv2.Enabled = False
  MoveLast2.Enabled = False
  SAVEBC.Enabled = True
  CANCEL2.Enabled = True
  EDITBC.Enabled = False
  DELETEBC.Enabled = False
  ADDBC.Enabled = False
  BCDG.Enabled = False
  BOOKCOUNT2.Caption = 1
End Sub

Private Sub ADDBCOPY_Click()
  BCOPY(3).Text = ""
  BCOPY(3).Locked = False
  BCOPY(3).BackColor = vbWhite
  BCOPY(2).Text = "自動生成"
  BC2.Text = ""
  BC2.Locked = False
  BC2.BackColor = vbWhite
  MoveOne4.Enabled = False
  MoveNext4.Enabled = False
  MovePerv4.Enabled = False
  MoveLast4.Enabled = False
  SAVEBCOPY.Enabled = True
  CANCEL4.Enabled = True
  EDITBCOPY.Enabled = False
  DELETEBCOPY.Enabled = False
  ADDBCOPY.Enabled = False
  BCOPYDG.Enabled = False
  BOOKCOUNT4.Caption = 1
End Sub

Private Sub ADDBOOK_Click()
  For i = 1 To 12
  Text1(i).Text = ""
  Text1(i).Locked = False
  Text1(i).BackColor = vbWhite
  Next i
  Text1(0).Text = "自動生成"
  Text1(13).Text = "自動生成"
  B5.Text = ""
  B5.Locked = False
  B5.BackColor = vbWhite
  B6.Text = ""
  B6.Locked = False
  B6.BackColor = vbWhite
  B7.Text = ""
  B7.Locked = False
  B7.BackColor = vbWhite
  B12.Text = ""
  B12.Locked = False
  B12.BackColor = vbWhite
  MoveOne.Enabled = False
  MoveNext.Enabled = False
  MovePerv.Enabled = False
  MoveLast.Enabled = False
  SAVEBOOK.Enabled = True
  CANCEL1.Enabled = True
  EDITBOOK.Enabled = False
  DELETEBOOK.Enabled = False
  ADDBOOK.Enabled = False
  BOOKDG.Enabled = False
  BOOKCOUNT.Caption = 1
End Sub

Private Sub ADDP_Click()
  For i = 1 To 4
  Text3(i).Text = ""
  Text3(i).Locked = False
  Text3(i).BackColor = vbWhite
  Next i
  Text3(0).Text = "自動生成"
  P1.Text = ""
  P1.Locked = False
  P1.BackColor = vbWhite
  MoveOne3.Enabled = False
  MoveNext3.Enabled = False
  MovePerv3.Enabled = False
  MoveLast3.Enabled = False
  SAVEP.Enabled = True
  CANCEL3.Enabled = True
  EDITP.Enabled = False
  DELETEP.Enabled = False
  ADDP.Enabled = False
  PDG.Enabled = False
  BOOKCOUNT3.Caption = 1
End Sub

Private Sub BCDG_Click()
        For i = 0 To Adodc2.Recordset.Fields.Count - 1
        If IsNull(Adodc2.Recordset.Fields(i)) Then
        Text2(i).Text = ""
        Else
        Text2(i).Text = Adodc2.Recordset.Fields(i)
        End If
        Next i
        BC3.Text = Text2(2).Text
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
        rs.Open "SELECT book.BookID AS 書籍編號, book.BookName1 AS 書籍名稱, book.BookName2 AS 書籍別名, book.Author1 AS 作者1, book.Author2 AS 作者2, book.BookType AS 書籍類別, book.ContentType1 AS 內容類別1, book.ContentType2 AS 內容類別2, book.PublishingHouse AS 出版社, book.Year AS 出版年份, book.Price AS 價錢, book.CollectionID AS 書籍總集編號, book.Episode AS 集數, book.CreateDate AS 建立日期 FROM book where Status = 'N' and CollectionID = '" & Text2(0).Text & "'", cn, adOpenKeyset, adLockOptimistic
        Set BookList.Recordset = rs
        rs.Close
        cn.Close
End Sub

Private Sub BCOPYDG_Click()
  If Adodc4.Recordset.RecordCount > 0 Then
  For i = 0 To BCOPY.UBound
        If IsNull(Adodc4.Recordset.Fields(i)) Then
        BCOPY(i).Text = ""
        Else
        BCOPY(i).Text = Adodc4.Recordset.Fields(i)
        End If
  Next i
  BC2.Text = BCOPY(3).Text
  End If
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "SELECT count(*) as a FROM BookCopy Where BookID = '" & BCOPY(0).Text & "'", cn, adOpenKeyset, adLockOptimistic
  Set BCOPYCOUNT.DataSource = rs
  rs.Close
  cn.Close
  BPT.Text = BCOPYCOUNT.TextMatrix(1, 1)
End Sub

Private Sub BOOKDG_Click()
  If Adodc1.Recordset.RecordCount > 0 Then
  For i = 0 To Text1.UBound
        If IsNull(Adodc1.Recordset.Fields(i)) Then
        Text1(i).Text = ""
        Else
        Text1(i).Text = Adodc1.Recordset.Fields(i)
        End If
  Next i
  B5.Text = Text1(5).Text
  B6.Text = Text1(6).Text
  B7.Text = Text1(7).Text
  B12.Text = Text1(11).Text
  End If
End Sub

Private Sub CANCEL1_Click()
  For i = 1 To 12
  Text1(i).Locked = True
  Text1(i).BackColor = &H8000000A
  Next i
  B5.Locked = True
  B5.BackColor = &H8000000A
  B6.Locked = True
  B6.BackColor = &H8000000A
  B7.Locked = True
  B7.BackColor = &H8000000A
  B12.Locked = True
  B12.BackColor = &H8000000A
  MoveOne.Enabled = True
  MoveNext.Enabled = True
  MovePerv.Enabled = True
  MoveLast.Enabled = True
  SAVEBOOK.Enabled = False
  CANCEL1.Enabled = False
  EDITBOOK.Enabled = True
  DELETEBOOK.Enabled = True
  ADDBOOK.Enabled = True
  BOOKDG.Enabled = True
  Call BOOKDG_Click
End Sub

Private Sub CANCEL2_Click()
  For i = 1 To 2
  Text2(i).Locked = True
  Text2(i).BackColor = &H8000000A
  Next i
  BC3.Locked = True
  BC3.BackColor = &H8000000A
  MoveOne2.Enabled = True
  MoveNext2.Enabled = True
  MovePerv2.Enabled = True
  MoveLast2.Enabled = True
  SAVEBC.Enabled = False
  CANCEL2.Enabled = False
  EDITBC.Enabled = True
  DELETEBC.Enabled = True
  ADDBC.Enabled = True
  BCDG.Enabled = True
Call BCDG_Click
End Sub

Private Sub CANCEL3_Click()
  For i = 1 To 4
  Text3(i).Locked = True
  Text3(i).BackColor = &H8000000A
  Next i
  BC3.Locked = True
  BC3.BackColor = &H8000000A
  MoveOne3.Enabled = True
  MoveNext3.Enabled = True
  MovePerv3.Enabled = True
  MoveLast3.Enabled = True
  SAVEP.Enabled = False
  CANCEL3.Enabled = False
  EDITP.Enabled = True
  DELETEP.Enabled = True
  ADDP.Enabled = True
  PDG.Enabled = True
Call PDG_Click
End Sub

Private Sub CANCEL4_Click()
  For i = 1 To 3
  BCOPY(i).Locked = True
  BCOPY(i).BackColor = &H8000000A
  Next i
  BC2.Locked = True
  BC2.BackColor = &H8000000A
  MoveOne4.Enabled = True
  MoveNext4.Enabled = True
  MovePerv4.Enabled = True
  MoveLast4.Enabled = True
  SAVEBCOPY.Enabled = False
  CANCEL4.Enabled = False
  EDITBCOPY.Enabled = True
  DELETEBCOPY.Enabled = True
  ADDBCOPY.Enabled = True
  BCOPYDG.Enabled = True
Call BCOPYDG_Click
End Sub

Private Sub Command1_Click()
StaffUse.Show
PBView.Hide
End Sub

Private Sub DELETEBCOPY_Click()
MsgBox "是否刪除該書籍複本？", vbYesNo, "刪除書籍"
If vbYes Then
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
    If BC2.Text = "R" Then
            cn.Execute ("Update BookCopy Set SStatus = 'Y' where BookID = '" & BCOPY(0).Text & "' and CopyN = '" & BCOPY(2).Text & "'")
            cn.Close
            Adodc4.Refresh
            If Adodc4.Recordset.EOF Then
            Adodc4.Recordset.MoveLast
            End If
    MsgBox "該書籍複本已被刪除!"
    End If
        If BC2.Text = "B" Then
            MsgBox "該書籍複本仍未歸還,無法刪除!"
            cn.Close
        End If
End If
End Sub

Private Sub DELETEBOOK_Click()
MsgBox "是否刪除該書籍？" & vbCrLf & "如果刪除該書籍,相關複本的資料將會被刪除.", vbYesNo, "刪除書籍"
If vbYes Then
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
    rs.Open "Select * from BookCopy,Book where Book.BookID = BookCopy.BookID and BookCopy.status = 'B' and Book.BookID = '" & Text1(0).Text & "' ", cn, adOpenKeyset, adLockOptimistic
    Set BCOPYFIND.DataSource = rs
    rs.Close
    If BCOPYFIND.Row = 0 Then
            cn.Execute ("Update Book Set Status = 'Y' where BookID = '" & Text1(0).Text & "'")
            cn.Execute ("Update BookCopy Set SStatus = 'Y' where BookID = '" & Text1(0).Text & "'")
            cn.Close
            Adodc1.Refresh
            Adodc1.Recordset.MoveNext
            If Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveLast
            End If
    MsgBox "該書籍和相關複本已被刪除!"
    End If
        If BCOPYFIND.Row <> 0 Then
            MsgBox "該書籍的相關複本仍未歸還,無法刪除!"
            cn.Close
        End If
End If
End Sub

Private Sub DELETEP_Click()
MsgBox "是否刪除該貨品？", vbYesNo, "刪除貨品"
If vbYes Then
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
    If BC2.Text = "R" Then
            cn.Execute ("Update Product Set Status = 'Y' where ProductID = '" & Text3(0).Text & "'")
            cn.Close
            Adodc3.Refresh
            If Adodc3.Recordset.EOF Then
            Adodc3.Recordset.MoveLast
            End If
    MsgBox "該貨品已被刪除!"
    End If
End If
End Sub

Private Sub EDITBC_Click()
  For i = 1 To 2
  Text2(i).Locked = False
  Text2(i).BackColor = vbWhite
  Next i
  BC3.Locked = False
  BC3.BackColor = vbWhite
  MoveOne2.Enabled = False
  MoveNext2.Enabled = False
  MovePerv2.Enabled = False
  MoveLast2.Enabled = False
  SAVEBC.Enabled = True
  CANCEL2.Enabled = True
  EDITBC.Enabled = False
  DELETEBC.Enabled = False
  ADDBC.Enabled = False
  BCDG.Enabled = False
  BOOKCOUNT2.Caption = 2
End Sub

Private Sub EDITBCOPY_Click()
  BCOPY(3).Text = ""
  BCOPY(3).Locked = False
  BCOPY(3).BackColor = vbWhite
  BCOPY(2).Text = "自動生成"
  BC2.Text = ""
  BC2.Locked = False
  BC2.BackColor = vbWhite
  MoveOne4.Enabled = False
  MoveNext4.Enabled = False
  MovePerv4.Enabled = False
  MoveLast4.Enabled = False
  SAVEBCOPY.Enabled = True
  CANCEL4.Enabled = True
  EDITBCOPY.Enabled = False
  DELETEBCOPY.Enabled = False
  ADDBCOPY.Enabled = False
  BCOPYDG.Enabled = False
  BOOKCOUNT4.Caption = 2
End Sub

Private Sub EDITBOOK_Click()
  For i = 1 To 12
  Text1(i).Locked = False
  Text1(i).BackColor = vbWhite
  Next i
  B5.Locked = False
  B5.BackColor = vbWhite
  B6.Locked = False
  B6.BackColor = vbWhite
  B7.Locked = False
  B7.BackColor = vbWhite
  B12.Locked = False
  B12.BackColor = vbWhite
  MoveOne.Enabled = False
  MoveNext.Enabled = False
  MovePerv.Enabled = False
  MoveLast.Enabled = False
  SAVEBOOK.Enabled = True
  CANCEL1.Enabled = True
  EDITBOOK.Enabled = False
  DELETEBOOK.Enabled = False
  ADDBOOK.Enabled = False
  BOOKDG.Enabled = False
  BOOKCOUNT.Caption = 2
End Sub

Private Sub EDITP_Click()
  For i = 1 To 4
  Text3(i).Locked = False
  Text3(i).BackColor = vbWhite
  Next i
  P1.Locked = False
  P1.BackColor = vbWhite
  MoveOne3.Enabled = False
  MoveNext3.Enabled = False
  MovePerv3.Enabled = False
  MoveLast3.Enabled = False
  SAVEP.Enabled = True
  CANCEL3.Enabled = True
  EDITP.Enabled = False
  DELETEP.Enabled = False
  ADDP.Enabled = False
  PDG.Enabled = False
  BOOKCOUNT3.Caption = 2
End Sub

Private Sub Form_Load()
 Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
   Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
    Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
 BookList.ColWidth(0) = 150 * 3
 Call BOOKDG_Click
 Call BCDG_Click
 Call PDG_Click
 Call BCOPYDG_Click
End Sub

Private Sub MoveLast_Click()
Adodc1.Recordset.MoveLast
Call BOOKDG_Click
End Sub

Private Sub MoveLast2_Click()
Adodc2.Recordset.MoveLast
Call BCDG_Click
End Sub

Private Sub MoveLast3_Click()
Adodc3.Recordset.MoveLast
Call PDG_Click
End Sub

Private Sub MoveLast4_Click()
Adodc4.Recordset.MoveLast
Call BCOPYDG_Click
End Sub

Private Sub MoveNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
   Adodc1.Recordset.MoveLast
End If
Call BOOKDG_Click
End Sub

Private Sub MoveNext2_Click()
Adodc2.Recordset.MoveNext
If Adodc2.Recordset.EOF Then
   Adodc2.Recordset.MoveLast
End If
Call BCDG_Click
End Sub

Private Sub MoveNext3_Click()
Adodc3.Recordset.MoveNext
If Adodc3.Recordset.EOF Then
   Adodc3.Recordset.MoveLast
End If
Call PDG_Click
End Sub

Private Sub MoveNext4_Click()
Adodc3.Recordset.MoveNext
If Adodc3.Recordset.EOF Then
   Adodc3.Recordset.MoveLast
End If
Call BCOPYDG_Click
End Sub

Private Sub MoveOne_Click()
Adodc1.Recordset.MoveFirst
Call BOOKDG_Click
End Sub

Private Sub MoveOne2_Click()
Adodc2.Recordset.MoveFirst
Call BCDG_Click
End Sub

Private Sub MoveOne3_Click()
Adodc3.Recordset.MoveFirst
Call PDG_Click
End Sub

Private Sub MoveOne4_Click()
Adodc4.Recordset.MoveFirst
Call BCOPYDG_Click
End Sub

Private Sub MovePerv_Click()
Adodc1.Recordset.MovePrevious

If Adodc1.Recordset.BOF Then
   Adodc1.Recordset.MoveFirst
End If
Call BOOKDG_Click
End Sub

Private Sub MovePerv2_Click()
Adodc2.Recordset.MovePrevious
If Adodc2.Recordset.BOF Then
   Adodc2.Recordset.MoveFirst
End If
Call BCDG_Click
End Sub

Private Sub MovePerv3_Click()
Adodc3.Recordset.MovePrevious
If Adodc3.Recordset.BOF Then
   Adodc3.Recordset.MoveFirst
End If
Call PDG_Click
End Sub

Private Sub MovePerv4_Click()
Adodc4.Recordset.MovePrevious
If Adodc4.Recordset.BOF Then
   Adodc4.Recordset.MoveFirst
End If
Call BCOPYDG_Click
End Sub

Private Sub PDG_Click()
  If Adodc3.Recordset.RecordCount > 0 Then
  For i = 0 To Text3.UBound
        If IsNull(Adodc3.Recordset.Fields(i)) Then
        Text3(i).Text = ""
        Else
        Text3(i).Text = Adodc3.Recordset.Fields(i)
        End If
  Next i
  P1.Text = Text3(3).Text
  End If
End Sub

Private Sub SAVEBC_Click()
Dim Y As String
Dim X As String
Y = ""
If BOOKCOUNT2.Caption = "1" Then
Text2(2).Text = BC3.Text
    If Text2(1).Text = "" Then
    Y = Y & vbCrLf & "書籍總集沒有名稱!"
    End If
    If BC3.Text = "" Then
    Y = Y & vbCrLf & "書籍總集狀態不可空白!"
    End If
    'If Trim(BC3.Text) <> "全本" Or Trim(BC3.Text) <> "連載" Then
    'Y = Y & vbCrLf & "書籍總集狀態只能填入指定的項目!"
    'End If
        If Y = "" Then
        Call CreateBCID(X)
        Text2(0).Text = X
            cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
            rs.Open "BookCollection", cn, adOpenKeyset, adLockOptimistic
        rs.AddNew
        For i = 0 To 2
        rs.Fields(i) = Text2(i).Text
        Next i
        rs.Fields(3) = "N"
        rs.Update
        rs.Close
        cn.Close
        Adodc2.Refresh
        Adodc2.Recordset.MoveLast
    MsgBox "該書籍總集已新增!"
    Call CANCEL2_Click
    Call BCDG_Click
    Else
    MsgBox Y
    End If
    End If
If BOOKCOUNT2.Caption = "2" Then
Text2(2).Text = BC3.Text
    If Text2(1).Text = "" Then
    Y = Y & vbCrLf & "書籍總集沒有名稱!"
    End If
    If Text2(2).Text = "" Then
    Y = Y & vbCrLf & "書籍總集狀態不可空白!"
    End If
    'If Trim(BC3.Text) <> "全本" Or Trim(BC3.Text) <> "連載" Then
    'Y = Y & vbCrLf & "書籍總集狀態只能填入指定的項目!"
    'End If
    If Y = "" Then
        For i = 1 To 2
    If IsNull(Adodc2.Recordset.Fields(i)) Then
    Adodc2.Recordset.Fields(i) = ""
    Adodc2.Recordset.Update
    Else
    Adodc2.Recordset.Fields(i) = Text2(i).Text
    Adodc2.Recordset.Update
    End If
    Next i
MsgBox "該書籍總集變更成功!"
Call CANCEL2_Click
Call BCDG_Click
Else
MsgBox Y
End If
End If
End Sub

Private Sub SAVEBOOK_Click()
Dim X As String
If BOOKCOUNT.Caption = "1" Then
Dim Y As String
If Text1(1).Text = "" Then
X = X & vbCrLf & "請輸入書籍名稱!"
End If
If Text1(3).Text = "" Then
X = X & vbCrLf & "請輸入作者名稱!"
End If
If B5.Text = "" Then
X = X & vbCrLf & "請輸入書籍類別!"
End If
If B6.Text = "" Then
X = X & vbCrLf & "請輸入內容類別!"
End If
If Text1(8).Text = "" Then
X = X & vbCrLf & "請輸入出版社!"
End If
If Text1(9).Text = "" Then
X = X & vbCrLf & "請輸入出版年份!"
End If
If Val(Text1(9).Text) > 2013 Or Val(Text1(9).Text) < 1900 Then
X = X & vbCrLf & "請輸入正確的出版年份!"
End If
If Text1(10).Text = "" Then
X = X & vbCrLf & "請輸入價錢!"
End If
If IsNumeric(Text1(10).Text) = False Then
X = X & vbCrLf & "請輸入正確的價錢!"
End If

If X = "" Then
Call CreateBID(Y)
Text1(0).Text = Y
Text1(5).Text = B5.Text
Text1(6).Text = B6.Text
Text1(7).Text = B7.Text
Text1(11).Text = B12.Text
Text1(13).Text = Format(Now, "dd/mm/yyyy")
If i = 12 And Text1(12).Text = "" Then
rs.Fields(i) = "0"
End If
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
rs.Open "Book", cn, adOpenKeyset, adLockOptimistic
rs.AddNew
For i = 0 To 13
If i = 12 Then
rs.Fields(i) = Val(Text1(i).Text)
Else
rs.Fields(i) = Text1(i).Text
End If
Next i
rs.Fields(14) = "N"
rs.Update
rs.Close
MsgBox "該書籍新增成功!"
Adodc1.Refresh
Adodc1.Recordset.MoveLast
Call CANCEL1_Click
Call BOOKDG_Click
Else
MsgBox X
End If
End If


If BOOKCOUNT.Caption = "2" Then
If Text1(1).Text = "" Then
X = X & vbCrLf & "請輸入書籍名稱!"
End If
If Text1(3).Text = "" Then
X = X & vbCrLf & "請輸入作者名稱!"
End If
If B5.Text = "" Then
X = X & vbCrLf & "請輸入內容名稱!"
End If
If Text1(8).Text = "" Then
X = X & vbCrLf & "請輸入出版社!"
End If
If Text1(9).Text = "" Then
X = X & vbCrLf & "請輸入出版年份!"
End If
If Val(Text1(9).Text) > 2013 Or Val(Text1(9).Text) < 1900 Then
X = X & vbCrLf & "請輸入正確的出版年份!"
End If
If Text1(10).Text = "" Then
X = X & vbCrLf & "請輸入價錢!"
End If
If IsNumeric(Text1(10).Text) = False Then
X = X & vbCrLf & "請輸入正確的價錢!"
End If

If X = "" Then
For i = 1 To 12
If IsNull(Adodc1.Recordset.Fields(i)) Then
    Adodc1.Recordset.Fields(i) = ""
    Adodc1.Recordset.Update
    Else
    Adodc1.Recordset.Fields(i) = Text1(i).Text
    Adodc1.Recordset.Update
    End If
    Next i
MsgBox "該書籍資料變更成功!"
Call CANCEL1_Click
Call BOOKDG_Click
Else
MsgBox X
End If
End If
End Sub
Function CreateBID(XY As String)
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
    rs.Open "Book", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
    rs.MoveLast
    XY = "B" + CStr(Format(CInt(Right(rs.Fields(0), 7)) + 1, "0000000"))
    Else
    XY = "B0000001"
    End If
    rs.Close
    cn.Close
End Function

Function CreateBCID(XY As String)
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
    rs.Open "BookCollection", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
    rs.MoveLast
    XY = "C" + CStr(Format(CInt(Right(rs.Fields(0), 7)) + 1, "0000000"))
    Else
    XY = "C0000001"
    End If
    rs.Close
    cn.Close
End Function

