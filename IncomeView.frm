VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form IncomeView 
   Caption         =   "�˵��C�ѩΨC�몺�`��~�B"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   11970
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox IT 
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0"
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FFFF&
      Caption         =   "�������J�ӷ�(�i�h��X�U�C�����s���V�ܬ�����T):"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   11775
      Begin VB.OptionButton O1 
         BackColor       =   &H0000FFFF&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton O4 
         BackColor       =   &H0000FFFF&
         Caption         =   "��L�O��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6000
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton O3 
         BackColor       =   &H0000FFFF&
         Caption         =   "�P��q��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton O2 
         BackColor       =   &H0000FFFF&
         Caption         =   "�ɮѭq��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�T�w"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10320
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "�������:"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4440
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
      Begin VB.ComboBox DDLIST 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox MMLIST 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox YYLIST 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton YY 
         BackColor       =   &H0000FF00&
         Caption         =   "�~��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton MM 
         BackColor       =   &H0000FF00&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton DD 
         BackColor       =   &H0000FF00&
         Caption         =   "���w���"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "�~"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   22
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   9
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   1080
         Width           =   495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid INCOMEVIEWS 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5106
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�з���"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Exit 
      Caption         =   "���}"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   0
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Label DS 
      Caption         =   "Label9"
      Height          =   735
      Left            =   9000
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label MS 
      Caption         =   "Label8"
      Height          =   615
      Left            =   9000
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label YS 
      Caption         =   "Label7"
      Height          =   495
      Left            =   9000
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "�Х���ܤ������,�M���J��������ë��U�T�w�ӷj�����w��������J"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   11895
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  '�z��
      Caption         =   "�`���J/�ӳ������J:"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   7560
      Width           =   3375
   End
End
Attribute VB_Name = "IncomeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
YS.Caption = YYLIST.Text
MS.Caption = MMLIST.Text
DS.Caption = DDLIST.Text
If YYLIST.Text = "" Then
YS.Caption = Format(Now, "YYYY")
End If
If MMLIST.Text = "" Then
MS.Caption = Format(Now, "MM")
End If
If DDLIST.Text = "" Then
DS.Caption = Format(Now, "DD")
End If
If YY.Value Then
'YS.Caption = YYLIST.Text
    If O1.Value Then
    rs.Open "SELECT BOID as �����q��s��,TotalFee as �q�榬�J,BODate as �إߤ�� FROM BorrowOrder WHERE YEAR(BODate) = " & CInt(YS.Caption) & " UNION SELECT OID,TotalPrice,OrderDate FROM SellOrder WHERE YEAR(OrderDate) = " & CInt(YS.Caption) & " UNION SELECT EID,Payment,PaymentDate FROM ExtraFee WHERE YEAR(PaymentDate) = " & CInt(YS.Caption) & ";", cn, adOpenKeyset, adLockOptimistic
    Set INCOMEVIEWS.DataSource = rs
    rs.Close
    End If
    If O2.Value Then
    rs.Open "SELECT BOID as �ɾ\�q��s��,TotalFee as �q�榬�J,BODate as �إߤ�� FROM BorrowOrder WHERE  YEAR(BODate) = " & CInt(YS.Caption) & ";", cn, adOpenKeyset, adLockOptimistic
    Set INCOMEVIEWS.DataSource = rs
    rs.Close
    End If
    If O3.Value Then
    rs.Open "SELECT OID as �q��s��,TotalPrice as �q�榬�J,OrderDate as �إߤ�� FROM SellOrder WHERE YEAR(OrderDate) = " & CInt(YS.Caption) & ";", cn, adOpenKeyset, adLockOptimistic
    Set INCOMEVIEWS.DataSource = rs
    rs.Close
    End If
    If O4.Value Then
    rs.Open "SELECT EID as ��L���O�s��,Payment as �q�榬�J,PaymentDate as �إߤ�� FROM ExtraFee WHERE YEAR(PaymentDate) = " & CInt(YS.Caption) & ";", cn, adOpenKeyset, adLockOptimistic
    Set INCOMEVIEWS.DataSource = rs
    rs.Close
    End If
End If
If MM.Value Then
'YS.Caption = YYLIST.Text
'MS.Caption = MMLIST.Text
    If O1.Value Then
      rs.Open "SELECT BOID as �����q��s��,TotalFee as �q�榬�J,BODate as �إߤ�� FROM BorrowOrder WHERE YEAR(BODate) = " & CInt(YS.Caption) & " AND MONTH(BODate) = " & CInt(MS.Caption) & " UNION SELECT OID,TotalPrice,OrderDate FROM SellOrder WHERE YEAR(OrderDate) = " & CInt(YS.Caption) & "AND MONTH(OrderDate) = " & CInt(MS.Caption) & " UNION SELECT EID,Payment,PaymentDate FROM ExtraFee WHERE YEAR(PaymentDate) = " & CInt(YS.Caption) & "AND MONTH(PaymentDate) = " & CInt(MS.Caption) & " ;", cn, adOpenKeyset, adLockOptimistic
        Set INCOMEVIEWS.DataSource = rs
      rs.Close
    End If
    If O2.Value Then
    rs.Open "SELECT BOID as �ɾ\�q��s��,TotalFee as �q�榬�J,BODate as �إߤ�� FROM BorrowOrder WHERE YEAR(BODate) = " & CInt(YS.Caption) & " AND MONTH(BODate) = " & CInt(MS.Caption) & ";", cn, adOpenKeyset, adLockOptimistic
        Set INCOMEVIEWS.DataSource = rs
      rs.Close
    End If
    If O3.Value Then
      rs.Open "SELECT OID as �q��s��,TotalPrice as �q�榬�J,OrderDate as �إߤ�� FROM SellOrder WHERE YEAR(OrderDate) = " & CInt(YS.Caption) & "AND MONTH(OrderDate) = " & CInt(MS.Caption) & ";", cn, adOpenKeyset, adLockOptimistic
        Set INCOMEVIEWS.DataSource = rs
      rs.Close
    End If
    If O4.Value Then
        rs.Open "SELECT EID as ��L���O�s��,Payment as �q�榬�J,PaymentDate as �إߤ�� FROM ExtraFee WHERE YEAR(PaymentDate) = " & CInt(YS.Caption) & "AND MONTH(PaymentDate) = " & CInt(MS.Caption) & ";", cn, adOpenKeyset, adLockOptimistic
        Set INCOMEVIEWS.DataSource = rs
      rs.Close
    End If
End If
If DD.Value Then
'YS.Caption = YYLIST.Text
'MS.Caption = MMLIST.Text
'DS.Caption = DDLIST.Text
    If O1.Value Then
      rs.Open "SELECT BOID,TotalFee as �����q��s��,BODate FROM BorrowOrder WHERE BODate  = #  " & CInt(DS.Caption) & " / " & CInt(MS.Caption) & " / " & CInt(YS.Caption) & "# UNION SELECT OID,TotalPrice,OrderDate FROM SellOrder WHERE OrderDate  = #  " & CInt(DS.Caption) & " / " & CInt(MS.Caption) & " / " & CInt(YS.Caption) & "# UNION SELECT EID,Payment,PaymentDate FROM ExtraFee WHERE PaymentDate = #  " & CInt(DS.Caption) & " / " & CInt(MS.Caption) & " / " & CInt(YS.Caption) & "#", cn, adOpenKeyset, adLockOptimistic
        Set INCOMEVIEWS.DataSource = rs
      rs.Close
    End If
    If O2.Value Then
      rs.Open "SELECT BOID as �ɾ\�q��s��,TotalFee as �q�榬�J,BODate as �إߤ�� FROM BorrowOrder WHERE BODate  = #  " & CInt(DS.Caption) & " / " & CInt(MS.Caption) & " / " & CInt(YS.Caption) & "#;", cn, adOpenKeyset, adLockOptimistic
        Set INCOMEVIEWS.DataSource = rs
      rs.Close
    End If
    If O3.Value Then
      rs.Open "SELECT OID as �q��s��,TotalPrice as �q�榬�J,OrderDate as �إߤ�� FROM SellOrder WHERE OrderDate  = #  " & CInt(DS.Caption) & " / " & CInt(MS.Caption) & " / " & CInt(YS.Caption) & "#;", cn, adOpenKeyset, adLockOptimistic
        Set INCOMEVIEWS.DataSource = rs
      rs.Close
    End If
    If O4.Value Then
      rs.Open "SELECT  EID as ��L���O�s��,Payment as �q�榬�J,PaymentDate as �إߤ�� FROM ExtraFee WHERE PaymentDate = #  " & CInt(DS.Caption) & " / " & CInt(MS.Caption) & " / " & CInt(YS.Caption) & "#;", cn, adOpenKeyset, adLockOptimistic
        Set INCOMEVIEWS.DataSource = rs
      rs.Close
    End If
End If
End Sub

Private Sub DD_Click()
If DD.Value = True Then
YYLIST.Enabled = True
MMLIST.Enabled = True
DDLIST.Enabled = True
End If
If DD.Value = False Then
YYLIST.Enabled = False
MMLIST.Enabled = False
DDLIST.Enabled = False
End If
End Sub

Private Sub Exit_Click()
StaffUse.Show
IncomeView.Hide
End Sub

Private Sub Form_Load()
Dim AA As Integer
AA = Format(Now, "yyyy")
INCOMEVIEWS.ColWidth(0) = 150 * 3
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BookD.mdb"
  rs.Open "SELECT BOID,TotalFee,BODate FROM BorrowOrder Union SELECT EID,Payment,PaymentDate FROM ExtraFee Union SELECT OID,TotalPrice,OrderDate FROM SellOrder", cn, adOpenKeyset, adLockOptimistic
  Set INCOMEVIEWS.DataSource = rs
  rs.Close
  IncomeView.Picture = LoadPicture(App.Path & "\PIC\papab101.jpg")
For i = 2011 To AA
YYLIST.AddItem (i)
Next i
For i = 1 To 12
MMLIST.AddItem (i)
Next i
For i = 1 To 31
DDLIST.AddItem (i)
Next i
Dim B As Integer
Dim QQ As Integer
For X = 1 To INCOMEVIEWS.Row - 1
QQ = QQ + Val(INCOMEVIEWS.TextMatrix(i, 2))
B = QQ
Next X
IT.Text = CStr(B)
End Sub

Private Sub MM_Click()
If MM.Value = True Then
YYLIST.Enabled = True
MMLIST.Enabled = True
End If
If MM.Value = False Then
YYLIST.Enabled = False
MMLIST.Enabled = False
End If
End Sub

Private Sub YY_Click()
If YY.Value = True Then
YYLIST.Enabled = True
End If
If YY.Value = False Then
YYLIST.Enabled = False
End If
End Sub

'SELECT BOID,TotalFee,BODate FROM BorrowOrder WHERE MONTH(BODate) = 4
'Union All
'SELECT OID,TotalPrice,OrderDate FROM SellOrder WHERE DAY(OrderDate) = 2
'UNION ALL SELECT EID,Payment,PaymentDate FROM ExtraFee WHERE PaymentDate = #26/4/2013#;

