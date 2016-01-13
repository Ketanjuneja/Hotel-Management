VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   Caption         =   "BILL"
   ClientHeight    =   9510
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   9510
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   1095
      Left            =   12120
      Top             =   2400
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1931
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=WS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "WS"
      OtherAttributes =   ""
      UserName        =   "HOTEL"
      Password        =   "1234"
      RecordSource    =   "ROOM"
      Caption         =   "Adodc7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   855
      Left            =   12480
      Top             =   3960
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "HOTEL"
      Password        =   "1234"
      RecordSource    =   "GUEST"
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   19560
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8040
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      TabIndex        =   8
      Top             =   6480
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      TabIndex        =   4
      Top             =   4800
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   975
      Left            =   17160
      Top             =   8040
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "HOTEL"
      Password        =   "1234"
      RecordSource    =   "SELECT * FROM ALLOCATE"
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   18480
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "HOTEL"
      Password        =   "1234"
      RecordSource    =   "SELECT * FROM ROOM"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   855
      Left            =   17280
      Top             =   6960
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "HOTEL"
      Password        =   "1234"
      RecordSource    =   "select * FROM food_bill"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6240
      TabIndex        =   3
      Top             =   3000
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   17520
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=HOTEL;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "HOTEL"
      Password        =   "1234"
      RecordSource    =   "select * from FOOD_BILL where CUSt_ID=0"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "GET BILL"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   4335
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form6.frx":652E0
      Height          =   2415
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   65280
      ForeColor       =   255
      HeadLines       =   1
      RowHeight       =   33
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form6.frx":652F5
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   65280
      ForeColor       =   255
      HeadLines       =   1
      RowHeight       =   33
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL BILL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   6720
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ROOM BILL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RESTAURANT BILL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   5775
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Text3.Text = Val(Text1.Text) + Val(Text2.Text)
End Sub

Private Sub Command2_Click()
'DataEnvironment1.rsCommand1.DataSource=
DataReport1.Show

End Sub

Private Sub Command3_Click()
Adodc1.Refresh
While (Adodc1.Recordset.EOF <> True)
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
Wend
Adodc3.Recordset.MoveFirst
While (Adodc3.Recordset.EOF <> True)
Adodc7.CommandType = adCmdText
Adodc7.RecordSource = " SELECT * FROM ROOM WHERE ROOM_NO=" & Adodc3.Recordset.Fields(0)

Adodc7.Refresh

Adodc7.Recordset.Fields("AVAILABILITY").Value = 1
Adodc7.Recordset.Update


Adodc3.Recordset.Delete
Adodc3.Recordset.MoveNext
Wend
Form4.Show
Unload Form6


End Sub

Private Sub Form_Load()
Dim a As Integer
a = Val(InputBox("enter cust id:"))
Adodc5.CommandType = adCmdText
Adodc5.RecordSource = "select * from GUEST WHERE CUST_ID=" & a
Adodc5.Refresh

If (Adodc5.Recordset.EOF = True) Then
MsgBox ("customer id not valid")
Unload Form6
Form2.Show
Else
Adodc5.Recordset.MoveFirst
While (Adodc5.Recordset.EOF <> True)
'MsgBox (Adodc5.Recordset.Fields(0).Value)
Adodc5.Recordset.Delete
Adodc5.Recordset.MoveNext
Wend
End If
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select sum(PRICE) from FOOD_BILL group by cust_id having CUst_ID=" & a
Adodc2.Refresh
If (Adodc2.Recordset.EOF <> True) Then
Text1.Text = Adodc2.Recordset.Fields(0)

Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from FOOD_BILL where CUst_ID=" & a
Adodc1.Refresh
Else
Text1.Text = 0
DataGrid1.Visible = False
End If

Adodc4.CommandType = adCmdText
Adodc4.RecordSource = "select sum(PRICE) from ALLOCATE group by cust_id having CUst_ID=" & a
Adodc4.Refresh
If (Adodc4.Recordset.EOF <> True) Then
Text2.Text = Adodc4.Recordset.Fields(0)

Adodc3.CommandType = adCmdText
Adodc3.RecordSource = "select * from ALLOCATE where CUst_ID=" & a
Adodc3.Refresh
Else
Text2.Text = 0
DataGrid2.Visible = False
End If



'Adodc1.CommandType = adCmdText
'Adodc1.Refresh
'Adodc1.Recordset.MoveFirst
'While (Adodc1.Recordset.EOF <> True)
'If (Adodc1.Recordset.Fields(0) = a) Then
'  DataGrid2.Row = Adodc1.Recordset.
' If
'Adodc1.Recordset.MoveNext
'Wend
Form6.Width = 1280
Form6.Height = 800

End Sub

