VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   Caption         =   "HUB"
   ClientHeight    =   10200
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   1.65913e7
   ScaleMode       =   0  'User
   ScaleWidth      =   6.24318e5
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "BILL"
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
      Left            =   6720
      Picture         =   "Form4.frx":652E0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      UseMaskColor    =   -1  'True
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   9720
      Picture         =   "Form4.frx":CA5C0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   1080
      MaskColor       =   &H0080FFFF&
      Picture         =   "Form4.frx":12BC3D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   7575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RESTAURANT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   11160
      TabIndex        =   4
      Top             =   1920
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LODGING"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Integer
a = MsgBox("ARE YOU ALREADY REGISTERED?", vbYesNo)
If a = vbYes Then
Form5.Show
Unload Form4

Else
Form2.Show
Unload Form4

End If

End Sub

Private Sub Command2_Click()
a = MsgBox("ARE YOU ALREADY REGISTERED OR NOT?", vbYesNo)
If a = vbYes Then
Form3.Show
Unload Form4

Else
Form2.Show
Unload Form4

End If
End Sub

Private Sub Command3_Click()
Form6.Show
Unload Form4


End Sub

