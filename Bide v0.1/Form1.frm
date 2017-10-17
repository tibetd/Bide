VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form_Main 
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   2745
   ClientTop       =   1095
   ClientWidth     =   13125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   13125
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":50DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1FADE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   1058
      ButtonWidth     =   1323
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Startup"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Uglify"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Beautify"
            ImageIndex      =   2
            Style           =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Mode1"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Mode2"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Mode3"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   8235
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3975
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7011
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      FileName        =   "Startup.rtf"
      TextRTF         =   $"Form1.frx":21230
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   13095
      ExtentX         =   23098
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2775
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         ItemData        =   "Form1.frx":2153B
         Left            =   240
         List            =   "Form1.frx":2159F
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private formLoaded As Boolean
Private beautifymode As Integer
Dim WithEvents objDoc As MSHTMLCtl.HTMLDocument
Attribute objDoc.VB_VarHelpID = -1
Dim WithEvents objWind As MSHTMLCtl.HTMLWindow2
Attribute objWind.VB_VarHelpID = -1

Function tableSnippet() As String
    Form_Table.Show vbModal, Me
    'MsgBox Form_Table.m_ReturnValue  ' delete after stable
    tableSnippet = Form_Table.m_ReturnValue
    Set Form_Table = Nothing
End Function

Private Sub List1_DblClick()
    Dim strToInsert As String
    Select Case List1.List(List1.ListIndex)
        Case "Container": strToInsert = containerSnippet()
        Case "Tables": strToInsert = tableSnippet()
    Case Else

    End Select
    RichTextBox1.Text = Left(RichTextBox1.Text, RichTextBox1.SelStart) & strToInsert & _
    Mid(RichTextBox1.Text, RichTextBox1.SelStart + 1, Len(RichTextBox1.Text))
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then Form_Load
If Button.Index = 3 Then RichTextBox1.Text = Uglify(RichTextBox1.Text)
If Button.Index = 4 Then RichTextBox1.Text = Beautify(RichTextBox1.Text, beautifymode)
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
beautifymode = ButtonMenu.Index
Toolbar1.buttons(4).Caption = "Beautify " & ButtonMenu.Index
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Set objDoc = WebBrowser1.document
    Set objWind = objDoc.parentWindow
End Sub
Private Sub writeHTML(s As String)
WebBrowser1.navigate "about:blank" 'Required
WebBrowser1.document.open
WebBrowser1.Silent = True
WebBrowser1.document.write (s)
WebBrowser1.document.Close
End Sub
Private Sub Form_Load()
    formLoaded = False
    'RichTextBox1.LoadFile ("Startup.rtf")
    writeHTML (RichTextBox1.Text)
    formLoaded = True
End Sub
Private Sub RichTextBox1_Change()
    If formLoaded Then
        objDoc.Close
        WebBrowser1.Silent = True
        WebBrowser1.document.write (RichTextBox1.Text)
    End If
End Sub
