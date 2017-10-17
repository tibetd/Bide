VERSION 5.00
Begin VB.Form Form_Grid 
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4815
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   11
         Left            =   360
         TabIndex        =   89
         Top             =   5640
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   88
         Top             =   5160
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   87
         Top             =   4680
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   86
         Top             =   4200
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   85
         Top             =   3720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   84
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   83
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   82
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   81
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   80
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   79
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   78
         Top             =   360
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   360
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   69
         Top             =   240
         Width           =   2775
         Begin VB.OptionButton Option1 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option1 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   72
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   71
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option1 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   70
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   64
         Top             =   720
         Width           =   2775
         Begin VB.OptionButton Option2 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   68
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   67
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option2 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   66
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   65
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   840
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   58
         Top             =   1200
         Width           =   2775
         Begin VB.OptionButton Option3 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option3 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   61
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   60
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option3 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   59
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   3
         Left            =   840
         TabIndex        =   53
         Top             =   1680
         Width           =   2775
         Begin VB.OptionButton Option4 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   57
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton Option4 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   56
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option4 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   55
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   4
         Left            =   840
         TabIndex        =   48
         Top             =   2160
         Width           =   2775
         Begin VB.OptionButton Option5 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   51
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option5 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   50
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   49
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   5
         Left            =   840
         TabIndex        =   43
         Top             =   2640
         Width           =   2775
         Begin VB.OptionButton Option6 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   47
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton Option6 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   46
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option6 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   45
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option6 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   6
         Left            =   840
         TabIndex        =   38
         Top             =   3120
         Width           =   2775
         Begin VB.OptionButton Option7 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   42
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton Option7 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   41
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option7 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   40
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option7 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   7
         Left            =   840
         TabIndex        =   33
         Top             =   3600
         Width           =   2775
         Begin VB.OptionButton Option8 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option8 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   36
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option8 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   35
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option8 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   34
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   8
         Left            =   840
         TabIndex        =   28
         Top             =   4080
         Width           =   2775
         Begin VB.OptionButton Option9 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   32
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton Option9 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   31
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option9 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   30
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option9 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   9
         Left            =   840
         TabIndex        =   23
         Top             =   4560
         Width           =   2775
         Begin VB.OptionButton Option10 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option10 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   26
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option10 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   25
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option10 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   24
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   10
         Left            =   840
         TabIndex        =   18
         Top             =   5040
         Width           =   2775
         Begin VB.OptionButton Option11 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   22
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton Option11 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   21
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option11 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   20
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option11 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Index           =   11
         Left            =   840
         TabIndex        =   13
         Top             =   5520
         Width           =   2775
         Begin VB.OptionButton Option12 
            Caption         =   "xs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option12 
            Caption         =   "md"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   16
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton Option12 
            Caption         =   "lg"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   15
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton Option12 
            Caption         =   "sm"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   14
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4200
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4680
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   5160
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   5640
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Left :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   76
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Total number of columns should add up to 12. "
      Height          =   255
      Left            =   120
      TabIndex        =   75
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form_Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_ReturnValue As String
Private optchoice(12) As String

Private Sub Check1_Click(Index As Integer)
Combo1(Index).Enabled = IIf(Check1(Index).Value = 1, True, False)
End Sub

Private Sub Command1_Click()
m_ReturnValue = mySnippet
Unload Me
End Sub

Private Function mySnippet() As String
Dim s As String
Dim k As Integer
s = s & "<div class=""row"">" & vbCrLf
For k = 0 To 11
    If Combo1(k).Enabled Then
        s = s & "<div class=""col-" & optchoice(k) & "-" & Combo1.Item(k).Text & """>  </div>" & vbCrLf
    End If
Next
s = s & "</div>" & vbCrLf
mySnippet = s
End Function

Private Sub Combo1_Click(Index As Integer)
Dim k As Integer
Dim t As Integer
t = 0
For k = 0 To 11
    If (Combo1.Item(k).Enabled And (Combo1.Item(k).Text <> "")) Then t = t + Int(Combo1.Item(k).Text)
Next
If t < 12 Then
    Label1.Caption = "Total number of columns should add up to 12."
    Label2.Caption = 12 - t
End If
If t = 12 Then
    Label1.Caption = "Column number satisfied. Press Apply buttun."
    Label2.Caption = 0
End If
If t > 12 Then
    MsgBox "You have exceeded the allowable column count(12). Try again.."
    Combo1.Item(Index).Text = 1
End If
End Sub

Private Sub Form_Load()
Option1_Click (0)
Option2_Click (0)
Option3_Click (0)
Option4_Click (0)
Option5_Click (0)
Option6_Click (0)
Option7_Click (0)
Option8_Click (0)
Option9_Click (0)
Option10_Click (0)
Option11_Click (0)
Option12_Click (0)
End Sub

Private Sub Option1_Click(Index As Integer)
optchoice(0) = Option1.Item(Index).Caption
End Sub
Private Sub Option2_Click(Index As Integer)
optchoice(1) = Option1.Item(Index).Caption
End Sub
Private Sub Option3_Click(Index As Integer)
optchoice(2) = Option1.Item(Index).Caption
End Sub
Private Sub Option4_Click(Index As Integer)
optchoice(3) = Option1.Item(Index).Caption
End Sub
Private Sub Option5_Click(Index As Integer)
optchoice(4) = Option1.Item(Index).Caption
End Sub
Private Sub Option6_Click(Index As Integer)
optchoice(5) = Option1.Item(Index).Caption
End Sub
Private Sub Option7_Click(Index As Integer)
optchoice(6) = Option1.Item(Index).Caption
End Sub
Private Sub Option8_Click(Index As Integer)
optchoice(7) = Option1.Item(Index).Caption
End Sub
Private Sub Option9_Click(Index As Integer)
optchoice(8) = Option1.Item(Index).Caption
End Sub
Private Sub Option10_Click(Index As Integer)
optchoice(9) = Option1.Item(Index).Caption
End Sub
Private Sub Option11_Click(Index As Integer)
optchoice(10) = Option1.Item(Index).Caption
End Sub
Private Sub Option12_Click(Index As Integer)
optchoice(11) = Option1.Item(Index).Caption
End Sub
