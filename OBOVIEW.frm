VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form OBOVIEW 
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   13080
   StartUpPosition =   2  '螢幕中央
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   465
      Left            =   0
      TabIndex        =   43
      Top             =   8880
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11404
            Text            =   "日期:"
            TextSave        =   "日期:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11404
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10200
      Top             =   8520
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   2280
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "Select OID as 訂單編號,TotalPrice as 總收入,OrderDate as 建立日期 from SellOrder"
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
   Begin VB.CommandButton EXIT 
      Caption         =   "離開"
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
      Left            =   5040
      TabIndex        =   12
      Top             =   8040
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   120
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "Select BOID as 借閱訂單編號,MemberID as 會員編號,TotalFee as 總收入,BODate as 建立日期 from BorrowOrder"
      Caption         =   "Adodc1"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12938
      _Version        =   393216
      Tab             =   2
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "借閱訂單"
      TabPicture(0)   =   "OBOVIEW.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "銷售訂單"
      TabPicture(1)   =   "OBOVIEW.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "額外收費"
      TabPicture(2)   =   "OBOVIEW.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame9 
         Caption         =   "額外收費詳情:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   20.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   12735
         Begin VB.PictureBox P1 
            Height          =   5775
            Left            =   120
            ScaleHeight     =   5715
            ScaleWidth      =   2475
            TabIndex        =   40
            Top             =   480
            Width           =   2535
         End
         Begin VB.Frame Frame11 
            Caption         =   "額外收費列表:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5895
            Left            =   2760
            TabIndex        =   39
            Top             =   480
            Width           =   9855
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid DG5 
               Height          =   5295
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   9340
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "標楷體"
                  Size            =   18
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "標楷體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "訂單詳情:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   20.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   -74880
         TabIndex        =   21
         Top             =   720
         Width           =   12735
         Begin VB.TextBox OI 
            BackColor       =   &H80000003&
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
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   6000
            Width           =   1935
         End
         Begin VB.TextBox OI 
            BackColor       =   &H80000003&
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
            Left            =   10200
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   6000
            Width           =   1935
         End
         Begin VB.TextBox OI 
            BackColor       =   &H80000003&
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
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   480
            Width           =   1935
         End
         Begin VB.Frame Frame8 
            Caption         =   "訂單內容:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   2880
            TabIndex        =   29
            Top             =   2400
            Width           =   9735
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid DG4 
               Height          =   2895
               Left            =   120
               TabIndex        =   30
               Top             =   480
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   5106
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "標楷體"
                  Size            =   18
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "標楷體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "訂單列表:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   2655
            Begin MSDataGridLib.DataGrid DG3 
               Bindings        =   "OBOVIEW.frx":0054
               Height          =   5175
               Left            =   120
               TabIndex        =   28
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
         Begin VB.Frame Frame6 
            Caption         =   "訂單控制項:(或點選左方的列表指標瀏覽訂單的項列)"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   2880
            TabIndex        =   22
            Top             =   1200
            Width           =   9735
            Begin VB.CommandButton MoveOne2 
               Caption         =   "第一項訂單"
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
               TabIndex        =   26
               Top             =   360
               Width           =   2055
            End
            Begin VB.CommandButton MovePerv2 
               Caption         =   "上一項訂單"
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
               TabIndex        =   25
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton MoveNext2 
               Caption         =   "下一項的訂單"
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
               TabIndex        =   24
               Top             =   360
               Width           =   2295
            End
            Begin VB.CommandButton MoveLast2 
               Caption         =   "最後一項訂單"
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
               TabIndex        =   23
               Top             =   360
               Width           =   2295
            End
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFF00&
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
            Left            =   3000
            TabIndex        =   34
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFF00&
            Caption         =   "訂單的總收入:"
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
            Left            =   7920
            TabIndex        =   33
            Top             =   6000
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFF00&
            Caption         =   "訂單建立日期:"
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
            Left            =   2880
            TabIndex        =   32
            Top             =   6000
            Width           =   2295
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFF00&
            Caption         =   "元"
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
            Left            =   12120
            TabIndex        =   31
            Top             =   6000
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "借閱訂單詳情:"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   20.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   -74880
         TabIndex        =   1
         Top             =   720
         Width           =   12855
         Begin VB.Frame Frame5 
            Caption         =   "借閱訂單控制項:(或點選左方的列表指標瀏覽借閱訂單的項列)"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   2880
            TabIndex        =   15
            Top             =   1200
            Width           =   9735
            Begin VB.CommandButton MoveLast 
               Caption         =   "最後一項借閱訂單"
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
               TabIndex        =   19
               Top             =   360
               Width           =   2295
            End
            Begin VB.CommandButton MoveNext 
               Caption         =   "下一項的借閱訂單"
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
               TabIndex        =   18
               Top             =   360
               Width           =   2295
            End
            Begin VB.CommandButton MovePerv 
               Caption         =   "上一項借閱訂單"
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
               TabIndex        =   17
               Top             =   360
               Width           =   2175
            End
            Begin VB.CommandButton MoveOne 
               Caption         =   "第一項借閱訂單"
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
               TabIndex        =   16
               Top             =   360
               Width           =   2055
            End
         End
         Begin VB.Frame Frame4 
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
            Height          =   5775
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   2655
            Begin MSDataGridLib.DataGrid DG1 
               Bindings        =   "OBOVIEW.frx":0069
               Height          =   5175
               Left            =   120
               TabIndex        =   14
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
         Begin VB.Frame Frame3 
            Caption         =   "借閱訂單內容:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   18
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   2880
            TabIndex        =   10
            Top             =   2400
            Width           =   9735
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid DG2 
               Height          =   2895
               Left            =   120
               TabIndex        =   11
               Top             =   480
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   5106
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "標楷體"
                  Size            =   18
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "標楷體"
                  Size            =   15.75
                  Charset         =   136
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.TextBox BOI 
            BackColor       =   &H80000003&
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
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   6000
            Width           =   1935
         End
         Begin VB.TextBox BOI 
            BackColor       =   &H80000003&
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
            Left            =   10200
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   6000
            Width           =   1935
         End
         Begin VB.TextBox BOI 
            BackColor       =   &H80000003&
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
            Left            =   10680
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox BOI 
            BackColor       =   &H80000003&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFF00&
            Caption         =   "元"
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
            Left            =   12120
            TabIndex        =   20
            Top             =   6000
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFF00&
            Caption         =   "訂單建立日期:"
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
            Left            =   2880
            TabIndex        =   9
            Top             =   6000
            Width           =   2655
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "訂單的總收入:"
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
            Left            =   7920
            TabIndex        =   7
            Top             =   6000
            Width           =   2295
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF00&
            Caption         =   "持有訂單的會員編號:"
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
            Left            =   7320
            TabIndex        =   5
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label25 
            BackColor       =   &H00FFFF00&
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
            Left            =   3000
            TabIndex        =   3
            Top             =   480
            Width           =   2295
         End
      End
   End
   Begin VB.Label Label8 
      Caption         =   "如果需要查詢其他類別的資料,請按到指定的書籤目錄(例如下方的""額外收費"")"
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
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "OBOVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub DG1_Click()
  If Adodc1.Recordset.RecordCount > 0 Then
  For i = 0 To BOI.UBound
    BOI(i) = DG1.Columns(i).Text
  Next i
  Call FindList
  For i = 1 To DG2.Rows - 1
  If DG2.TextMatrix(i, 8) = "B" Then
  DG2.TextMatrix(i, 8) = "未歸還"
  End If
  If DG2.TextMatrix(i, 8) = "R" Then
  DG2.TextMatrix(i, 8) = "已歸還"
  End If
  Next i
  Else
  End If
End Sub

Private Sub DG3_Click()
  If Adodc2.Recordset.RecordCount > 0 Then
  For i = 0 To OI.UBound
    OI(i) = DG3.Columns(i).Text
  Next i
  Call FindOList
  End If
End Sub

Private Sub Exit_Click()
StaffUse.Show
OBOVIEW.Hide
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb;Persist Security Info=False"
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb;Persist Security Info=False"
Call DG1_Click
Call DG3_Click
DG2.ColWidth(0) = 150 * 3
DG2.ColWidth(1) = 710 * 3
DG2.ColWidth(3) = 710 * 3
DG4.ColWidth(0) = 150 * 3
DG4.ColWidth(1) = 710 * 3
DG4.ColWidth(3) = 710 * 3
DG5.ColWidth(0) = 150 * 3
DG5.ColWidth(1) = 710 * 3
DG5.ColWidth(2) = 710 * 3
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "SELECT ExtraFee.EID AS 額外收費編號, ExtraFee.ETYPE AS 額外收費類型, ExtraFee.MemberID AS 會員編號, ExtraFee.Payment AS 收取費用, ExtraFee.PaymentDate AS 建立日期 FROM ExtraFee ORDER BY ExtraFee.EID ASC", cn, adOpenKeyset, adLockOptimistic
  Set DG5.DataSource = rs
  rs.Close
  cn.Close
For i = 1 To DG5.Rows - 1
If DG5.TextMatrix(i, 2) = "1" Then
DG5.TextMatrix(i, 2) = "罰款"
End If
If DG5.TextMatrix(i, 2) = "2" Then
DG5.TextMatrix(i, 2) = "會員註冊"
End If
Next i
End Sub

Private Sub MoveLast_Click()
Adodc1.Recordset.MoveLast
For i = 0 To Adodc1.Recordset.Fields.Count - 1
   BOI(i).Text = Adodc1.Recordset.Fields(i)
Next i
Call FindList
End Sub

Private Sub MoveLast2_Click()
Adodc2.Recordset.MoveLast
For i = 0 To Adodc2.Recordset.Fields.Count - 1
   OI(i).Text = Adodc2.Recordset.Fields(i)
Next i
Call FindOList
End Sub

Private Sub MoveNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
   Adodc1.Recordset.MoveLast
End If
For i = 0 To Adodc1.Recordset.Fields.Count - 1
   BOI(i).Text = Adodc1.Recordset.Fields(i)
Next i
Call FindList
End Sub

Private Sub MoveNext2_Click()
Adodc2.Recordset.MoveNext
If Adodc2.Recordset.EOF Then
   Adodc2.Recordset.MoveLast
End If
For i = 0 To Adodc2.Recordset.Fields.Count - 1
   OI(i).Text = Adodc2.Recordset.Fields(i)
Next i
Call FindOList
End Sub

Private Sub MoveOne_Click()
Adodc1.Recordset.MoveFirst
For i = 0 To Adodc1.Recordset.Fields.Count - 1
   BOI(i).Text = Adodc1.Recordset.Fields(i)
Next i
Call FindList
End Sub

Private Sub MoveOne2_Click()
Adodc2.Recordset.MoveFirst
For i = 0 To Adodc2.Recordset.Fields.Count - 1
   OI(i).Text = Adodc2.Recordset.Fields(i)
Next i
Call FindOList
End Sub

Private Sub MovePerv_Click()
Adodc1.Recordset.MovePrevious

If Adodc1.Recordset.BOF Then
   Adodc1.Recordset.MoveFirst
End If

For i = 0 To Adodc1.Recordset.Fields.Count - 1
   BOI(i).Text = Adodc1.Recordset.Fields(i)
Next i
Call FindList
End Sub
Function FindList()
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open ("SELECT BorrowOrderList.BOLN as 借閱訂單項目編號,BorrowOrderList.BookID as 書籍編號,BorrowOrderList.CopyN as 書籍複本編號, Book.BookName1 as 書籍名稱,BorrowOrderList.BorrowDate as 借出日期,BorrowOrderList.DueDate as 歸還日期,BorrowOrderList.Fee as 借閱費用,BorrowOrderList.status as 書籍狀態 from BorrowOrderList ,Book Where BorrowOrderList.BookID = Book.BookID and BorrowOrderList.BOID =  '" & BOI(0).Text & "' ORDER BY BorrowOrderList.BOLN ASC;"), cn, adOpenKeyset, adLockOptimistic
  Set DG2.DataSource = rs
  rs.Close
  cn.Close
End Function
Function FindOList()
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "SELECT ProductOrderList.OLN,ProductOrderList.ProductID,Product.ProductName,ProductOrderList.Price,ProductOrderList.OrderN,ProductOrderList.TotalPrice From ProductOrderList,Product,SellOrder Where ProductOrderList.ProductID=Product.ProductID and  ProductOrderList.OID = SellOrder.OID and  SellOrder.OID =  '" & OI(0).Text & "' UNION ALL SELECT BookOrderList.OLN,BookOrderList.BookID,Book.BookName1,BookOrderList.Price,BookOrderList.OrderN,BookOrderList.TotalPrice From BookOrderList,Book,SellOrder  Where BookOrderList.BookID =Book.BookID and  BookOrderList.OID = SellOrder.OID and  SellOrder.OID = '" & OI(0).Text & "';", cn, adOpenKeyset, adLockOptimistic
  Set DG4.DataSource = rs
  rs.Close
  cn.Close
End Function

Private Sub MovePerv2_Click()
Adodc2.Recordset.MovePrevious
If Adodc2.Recordset.BOF Then
   Adodc2.Recordset.MoveFirst
End If
For i = 0 To Adodc2.Recordset.Fields.Count - 1
   OI(i).Text = Adodc2.Recordset.Fields(i)
Next i
Call FindList
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(1).Text = "日期:" + Format(Now, "dd/mm/yyyy")
StatusBar1.Panels(2).Text = "時間:" + Format(Now, "Medium Time")
End Sub
