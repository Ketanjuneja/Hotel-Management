VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "MENU CARD"
   ClientHeight    =   9345
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   9345
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "<=BACK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8040
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   5880
      Top             =   8160
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
   Begin VB.TextBox Text7 
      BackColor       =   &H00FF8080&
      DataField       =   "ITEM_ID"
      DataSource      =   "Adodc3"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   17280
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FF8080&
      DataField       =   "ITEM_NAME"
      DataSource      =   "Adodc3"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   17280
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF8080&
      DataField       =   "ITEM_NAME"
      DataSource      =   "Adodc1"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   11880
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF8080&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   17280
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   975
      Left            =   8760
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1720
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
      RecordSource    =   "FOOD_BILL"
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
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF8080&
      DataField       =   "PRICE"
      DataSource      =   "Adodc1"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   11880
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   6960
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      DataField       =   "ITEM_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   5640
      Width           =   6255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11160
      Top             =   5400
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1296
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
      RecordSource    =   "MENU_CARD"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":12F3DB
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483645
      HeadLines       =   1
      RowHeight       =   43
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "ITEM_ID"
         Caption         =   "ITEM_ID"
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
         DataField       =   "PRICE"
         Caption         =   "PRICE"
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
      BeginProperty Column02 
         DataField       =   "ITEM_NAME"
         Caption         =   "ITEM_NAME"
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
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4155.024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER QUANTITY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER ITEM ID"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   3615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim a, B, c, d, f, g As Integer
    Dim e As String
    
    
    
    Private Sub Command1_Click()
    Adodc3.Recordset.MoveLast
    If (Adodc3.Recordset.EOF <> True) Then
    Adodc3.Recordset.AddNew
    Adodc3.Recordset.Fields("CUST_ID") = a
    Adodc3.Recordset.Fields("ITEM_ID") = Val(Text1.Text)
    Adodc3.Recordset.Fields("PRICE") = d
    Adodc3.Recordset.Fields("QUANTITY") = Val(Text2.Text)
    Adodc3.Recordset.Fields("ITEM_NAME") = Text5.Text
    Adodc3.Recordset.Save
     
     Text1.Text = ""
    Text2.Text = ""
     
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    End Sub
    
    Private Sub Command2_Click()
    Form4.Show
    Unload Form3
    
    
    End Sub
    
    Private Sub Form_Load()
    a = Val(InputBox("enter the customer id"))
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select CUST_ID from GUEST WHERE CUST_ID=" & a
    Adodc2.Refresh
    
    If (Adodc2.Recordset.EOF <> False) Then
    MsgBox ("customer id not valid")
    Unload Form3
    Form2.Show
    End If
    End Sub
    
    Private Sub Text1_Change()
    g = Val(Text1.Text)
    
    End Sub
    
    Private Sub Text2_Change()
    c = Val(Text2.Text)
    d = B * c
    
    
     Text4.Text = d
     Text6.Text = e
     Text7.Text = g
    
    End Sub
    
    Private Sub Text2_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 47 And KeyAscii < 58) Then
    
    Else
    MsgBox ("Invalid")
    Text2.Text = ""
    End If
    End Sub
    
    Private Sub Text3_Change()
    B = Val(Text3.Text)
    
    Text3.Visible = False
    
    End Sub
    
    Private Sub Text5_Change()
    Text5.Visible = False
    e = Text5.Text
    
    End Sub
    
