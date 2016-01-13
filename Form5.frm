VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "ALLOCATION"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16695
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   9090
   ScaleWidth      =   16695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "ALLOCATE"
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   735
      Left            =   4440
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "GUEST"
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
      Height          =   615
      Left            =   16680
      Top             =   5040
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
      RecordSource    =   "ALLOCATE"
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
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   17040
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   17040
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   17160
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   6720
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      DataField       =   "AVAILABILITY"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11520
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      DataField       =   "PRICE"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11520
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   13560
      Top             =   6000
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "ROOM"
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
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      DataField       =   "ROOM_NO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   5760
      Width           =   4095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form5.frx":1E85F
      Height          =   4935
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   8705
      _Version        =   393216
      BackColor       =   -2147483645
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   33
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
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
      Caption         =   "NUMBER OF DAYS:-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   6600
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ROOM NUMBER:-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   5640
      Width           =   4335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, B, c, d, f, X As Integer
Dim k As String
Private Sub Command1_Click()



f = 0

If c = 1 Then

Text3.Text = f
Adodc1.Recordset.Update
Adodc1.Recordset.Save
Adodc2.Refresh

Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = Val(Text6.Text)
Adodc2.Recordset.Fields(1) = a
Adodc2.Recordset.Fields(2) = Val(Text5.Text)
Adodc2.Recordset.Fields(3) = Val(Text4.Text)
Adodc2.Recordset.Save
Else
MsgBox ("room is not available")
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""



End Sub

Private Sub Command2_Click()
Form4.Show
Unload Form5


End Sub

Private Sub Form_Load()
Adodc1.Visible = False
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

a = Val(InputBox("enter the customer id"))
MsgBox a
Text7.Text = a

Adodc3.CommandType = adCmdText
Adodc3.RecordSource = "select CUST_ID from GUEST WHERE CUST_ID=" & a
Adodc3.Refresh
If (Adodc3.Recordset.EOF = True) Then
MsgBox ("customer id not valid")
Unload Form5
Form2.Show
End If

Text3.Visible = False
Text2.Visible = True

End Sub

Private Sub Text1_Change()
k = Text1.Text

End Sub

Private Sub Text2_Change()
X = Val(Text2.Text)


End Sub

Private Sub Text3_Change()
c = Val(Text3.Text)



End Sub

Private Sub Text4_Change()
d = Val(Text4.Text)
X = X * d
Text5.Text = X
Text6.Text = k


End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Then

Else
MsgBox ("Invalid")
Text2.Text = ""
End If
End Sub
