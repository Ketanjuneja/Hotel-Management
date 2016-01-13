VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "REGISTRATION FORM"
   ClientHeight    =   9720
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   Picture         =   "form2.frx":0000
   ScaleHeight     =   9720
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer6 
      Interval        =   1200
      Left            =   16080
      Top             =   6480
   End
   Begin VB.Timer Timer5 
      Interval        =   1200
      Left            =   10680
      Top             =   6600
   End
   Begin VB.Timer Timer4 
      Interval        =   1200
      Left            =   16560
      Top             =   3240
   End
   Begin VB.Timer Timer3 
      Interval        =   1200
      Left            =   10680
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Interval        =   1200
      Left            =   16320
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   10560
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "REGISTRATION"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Caption         =   "NEXT=>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7920
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   13560
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   13560
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4920
         UseMaskColor    =   -1  'True
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4920
         UseMaskColor    =   -1  'True
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFF00&
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   4200
         TabIndex        =   4
         Top             =   3000
         Width           =   5415
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   4200
         TabIndex        =   2
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   1
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.Image Image12 
      Height          =   2910
      Left            =   14520
      Picture         =   "form2.frx":12F3DB
      Top             =   6960
      Width           =   3885
   End
   Begin VB.Image Image11 
      Height          =   2805
      Left            =   10200
      Picture         =   "form2.frx":132BA9
      Top             =   6960
      Width           =   4050
   End
   Begin VB.Image Image10 
      Height          =   2745
      Left            =   14520
      Picture         =   "form2.frx":135AA8
      Top             =   3720
      Width           =   4125
   End
   Begin VB.Image Image9 
      Height          =   2910
      Left            =   10200
      Picture         =   "form2.frx":1388CF
      Top             =   3600
      Width           =   3885
   End
   Begin VB.Image Image8 
      Height          =   2715
      Left            =   14880
      Picture         =   "form2.frx":13B4CF
      Top             =   360
      Width           =   4185
   End
   Begin VB.Image Image7 
      Height          =   2505
      Left            =   10320
      Picture         =   "form2.frx":13D23E
      Top             =   360
      Width           =   4515
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOMODATION"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10320
      TabIndex        =   9
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUISINES"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10440
      TabIndex        =   8
      Top             =   3120
      Width           =   8415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AMENITIES"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      Top             =   6480
      Width           =   8175
   End
   Begin VB.Image Image6 
      Height          =   2760
      Left            =   14400
      Picture         =   "form2.frx":13F836
      Top             =   6960
      Width           =   4110
   End
   Begin VB.Image Image5 
      Height          =   2745
      Left            =   14760
      Picture         =   "form2.frx":14198C
      Top             =   360
      Width           =   4140
   End
   Begin VB.Image Image4 
      Height          =   2670
      Left            =   10200
      Picture         =   "form2.frx":1434C2
      Top             =   6960
      Width           =   4260
   End
   Begin VB.Image Image3 
      Height          =   2910
      Left            =   14520
      Picture         =   "form2.frx":145A75
      Top             =   3600
      Width           =   3885
   End
   Begin VB.Image Image2 
      Height          =   2520
      Left            =   10080
      Picture         =   "form2.frx":147D39
      Top             =   360
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   2790
      Left            =   10200
      Picture         =   "form2.frx":149A97
      Top             =   3600
      Width           =   4065
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOB As String

Private Sub Command1_Click()
Adodc1.Recordset.MoveLast
Adodc1.Recordset.AddNew

Adodc1.Recordset.Fields(1) = Text1.Text
Adodc1.Recordset.Fields(2) = Val(Text2.Text)
Adodc1.Recordset.Save
MOB = Text1.Text
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select CUST_ID from GUEST where NAME= '" + MOB + "'"
Adodc2.Refresh
If (Adodc2.Recordset.EOF <> True) Then

Dim B
 Set B = Adodc2.Recordset.Fields(0)
MsgBox ("CUSTOMER ID IS " & B)
End If


End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()


Form4.Show
Unload Form2


End Sub

Private Sub Form_Load()
Adodc1.Visible = False
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""

Timer1.Enabled = True
Image2.Visible = True
Image7.Visible = False
Timer2.Enabled = True
Image8.Visible = True
Image5.Visible = False
Timer3.Enabled = True
Image9.Visible = True
Image1.Visible = False
Timer4.Enabled = True
Image10.Visible = True
Image3.Visible = False
Timer5.Enabled = True
Image11.Visible = True
Image4.Visible = True
Timer5.Enabled = True
Image6.Visible = True
Image12.Visible = False





End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If ((KeyAscii > 65 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123)) Then

Else
MsgBox ("Invalid")
Text1.Text = ""

End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Then

Else
MsgBox ("Invalid")
Text2.Text = ""
End If
End Sub

Private Sub Timer1_Timer()


If Image2.Visible = True Then
Image2.Visible = False
Image7.Visible = True

ElseIf Image7.Visible = True Then
Image7.Visible = False
Image2.Visible = True

ElseIf Image2.Visible = True Then
Image2.Visible = False
Image7.Visible = True

ElseIf Image7.Visible = True Then
Image7.Visible = False
Image2.Visible = True







End If






End Sub

Private Sub Timer2_Timer()

If Image8.Visible = True Then
Image8.Visible = False
Image5.Visible = True

ElseIf Image5.Visible = True Then
Image5.Visible = False
Image8.Visible = True

ElseIf Image8.Visible = True Then
Image5.Visible = False
Image8.Visible = True

ElseIf Image8.Visible = True Then
Image8.Visible = False
Image5.Visible = True
End If

End Sub

Private Sub Timer3_Timer()
If Image9.Visible = True Then
Image9.Visible = False
Image1.Visible = True

ElseIf Image1.Visible = True Then
Image1.Visible = False
Image9.Visible = True

ElseIf Image9.Visible = True Then
Image9.Visible = False
Image1.Visible = True

ElseIf Image1.Visible = True Then
Image1.Visible = False
Image9.Visible = True
End If
End Sub

Private Sub Timer4_Timer()
If Image10.Visible = True Then
Image10.Visible = False
Image3.Visible = True

ElseIf Image3.Visible = True Then
Image3.Visible = False
Image10.Visible = True

ElseIf Image10.Visible = True Then
Image10.Visible = False
Image3.Visible = True

ElseIf Image3.Visible = True Then
Image3.Visible = False
Image10.Visible = True

End If
End Sub

Private Sub Timer5_Timer()
If Image11.Visible = True Then
Image11.Visible = False
Image4.Visible = True

ElseIf Image4.Visible = True Then
Image4.Visible = False
Image11.Visible = True

ElseIf Image11.Visible = True Then
Image11.Visible = False
Image4.Visible = True

ElseIf Image4.Visible = True Then
Image4.Visible = False
Image11.Visible = True
End If

End Sub

Private Sub Timer6_Timer()
If Image6.Visible = True Then
Image6.Visible = False
Image12.Visible = True

ElseIf Image12.Visible = True Then
Image12.Visible = False
Image6.Visible = True

ElseIf Image6.Visible = True Then
Image6.Visible = False
Image12.Visible = True

ElseIf Image12.Visible = True Then
Image12.Visible = False
Image6.Visible = True
End If
End Sub
