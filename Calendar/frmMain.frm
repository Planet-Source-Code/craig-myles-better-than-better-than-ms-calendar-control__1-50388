VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Date Manipulation by Bill Hamilton and Craig Myles"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Lose focus by clicking here"
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   285
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "2. Use the command button to select a date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "1. Try typing a date DD/MM/YY and losing the text focus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim sDate As String
    
    sDate = GetOneDate(Me.Top + Text1.Top + Text1.Height, Me.Left + Text1.Left + Text1.Width)
    If sDate <> "" Then
        Text1.Text = sDate
    End If
End Sub

Private Sub Text1_DblClick()
    Command1_Click
End Sub

Private Sub Text1_LostFocus()
    Dim sDate As String
    Dim sTip As String
    
    sDate = Text1.Text
    sTip = dateCheck(sDate, "Date Manipulation by Bill Hamilton and Craig Myles")
    Text1.Text = sDate
End Sub
