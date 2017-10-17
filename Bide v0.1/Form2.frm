VERSION 5.00
Begin VB.Form Form_Table 
   Caption         =   "Form2"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   ScaleHeight     =   2625
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox Check4 
         Caption         =   "table-condensed"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "table-hover"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "table-bordered"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "table-striped"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_ReturnValue As String

Private Sub Command1_Click()
m_ReturnValue = tableSnippet
Unload Me
End Sub

Private Function tableSnippet() As String
Dim s As String
s = s & "  <table class=""table " & IIf(Check1.Value = 1, Check1.Caption & " ", "") & IIf(Check2.Value = 1, Check2.Caption & " ", "") & _
IIf(Check3.Value = 1, Check3.Caption & " ", "") & IIf(Check4.Value = 1, Check4.Caption & " ", "") & """>" & vbCrLf
s = s & "    <thead>" & vbCrLf
s = s & "      <tr>" & vbCrLf
s = s & "        <th>Firstname</th>" & vbCrLf
s = s & "        <th>Lastname</th>" & vbCrLf
s = s & "        <th>Email</th>" & vbCrLf
s = s & "      </tr>" & vbCrLf
s = s & "    </thead>" & vbCrLf
s = s & "    <tbody>" & vbCrLf
s = s & "      <tr>" & vbCrLf
s = s & "        <td>John</td>" & vbCrLf
s = s & "        <td>Doe</td>" & vbCrLf
s = s & "        <td>john@example.com</td>" & vbCrLf
s = s & "      </tr>" & vbCrLf
s = s & "      <tr>" & vbCrLf
s = s & "        <td>Mary</td>" & vbCrLf
s = s & "        <td>Moe</td>" & vbCrLf
s = s & "        <td>mary@example.com</td>" & vbCrLf
s = s & "      </tr>" & vbCrLf
s = s & "      <tr>" & vbCrLf
s = s & "        <td>July</td>" & vbCrLf
s = s & "        <td>Dooley</td>" & vbCrLf
s = s & "        <td>july@example.com</td>" & vbCrLf
s = s & "      </tr>" & vbCrLf
s = s & "    </tbody>" & vbCrLf
s = s & "  </table>" & vbCrLf
tableSnippet = s
End Function
