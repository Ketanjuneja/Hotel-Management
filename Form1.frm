VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LOGIN FORM"
   ClientHeight    =   8175
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   FillColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8175
   ScaleMode       =   0  'User
   ScaleWidth      =   690.262
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   5895
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "HOTEL TRANSYLVANIA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
If (Text1.Text = "admin" And Text2.Text = "123") Then
Form4.Show
Unload Form1

Else
MsgBox ("ERROR IN LOGIN")
Text1.Text = " "
Text2.Text = " "
End If

End Sub

