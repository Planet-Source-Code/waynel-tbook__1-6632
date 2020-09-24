VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  InfoE"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   ClipControls    =   0   'False
   Icon            =   "NewBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "NewBook.frx":030A
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3615
   ScaleWidth      =   8685
   Begin VB.CheckBox Check1 
      Caption         =   "Delete after due date ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   110
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   3960
      Picture         =   "NewBook.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   107
      ToolTipText     =   " Edit current record "
      Top             =   1080
      Width           =   555
   End
   Begin VB.CommandButton Command13 
      Height          =   375
      Left            =   3960
      Picture         =   "NewBook.frx":0F16
      Style           =   1  'Graphical
      TabIndex        =   106
      ToolTipText     =   " Exit program "
      Top             =   3120
      Width           =   555
   End
   Begin VB.CommandButton Command8 
      Height          =   375
      Left            =   3960
      Picture         =   "NewBook.frx":1258
      Style           =   1  'Graphical
      TabIndex        =   105
      ToolTipText     =   " Delete current record "
      Top             =   1560
      Width           =   555
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   3960
      Picture         =   "NewBook.frx":159A
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   " Add a record "
      Top             =   600
      Width           =   555
   End
   Begin VB.CommandButton Command10 
      Height          =   375
      Left            =   7440
      Picture         =   "NewBook.frx":18DC
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   " Save "
      Top             =   120
      Width           =   475
   End
   Begin VB.CommandButton Command11 
      Height          =   375
      Left            =   7920
      Picture         =   "NewBook.frx":1C1E
      Style           =   1  'Graphical
      TabIndex        =   102
      ToolTipText     =   " Cancel "
      Top             =   120
      Width           =   475
   End
   Begin VB.CommandButton Command9 
      Height          =   375
      Left            =   3960
      Picture         =   "NewBook.frx":1F60
      Style           =   1  'Graphical
      TabIndex        =   101
      ToolTipText     =   " About "
      Top             =   120
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4920
      Top             =   2040
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   3375
      Left            =   4680
      ScaleHeight     =   3315
      ScaleWidth      =   3795
      TabIndex        =   86
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         MaxLength       =   18
         TabIndex        =   92
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         MaxLength       =   15
         TabIndex        =   93
         Top             =   2880
         Width           =   1935
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2880
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   98
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Register"
         Height          =   375
         Left            =   2640
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   94
         Top             =   1560
         Width           =   855
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         TabIndex        =   89
         Top             =   840
         Width           =   3495
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "User :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   720
            TabIndex        =   90
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "NewBook.frx":22A2
         ScaleHeight     =   615
         ScaleWidth      =   1575
         TabIndex        =   87
         ToolTipText     =   " SoftWhere by EnyaW "
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "SoftWhere by EnyaW "
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   3720
         TabIndex        =   100
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2880
         MouseIcon       =   "NewBook.frx":51C4
         MousePointer    =   99  'Custom
         TabIndex        =   99
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ver 1.0"
         Height          =   255
         Left            =   1680
         TabIndex        =   88
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   295
      Left            =   3120
      TabIndex        =   85
      ToolTipText     =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   120
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5106
      _Version        =   327680
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Address"
      TabPicture(0)   =   "NewBook.frx":54CE
      Tab(0).ControlCount=   9
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text6"
      Tab(0).Control(0).Enabled=   -1  'True
      Tab(0).Control(1)=   "Text1(3)"
      Tab(0).Control(1).Enabled=   -1  'True
      Tab(0).Control(2)=   "Text1(2)"
      Tab(0).Control(2).Enabled=   -1  'True
      Tab(0).Control(3)=   "Text1(1)"
      Tab(0).Control(3).Enabled=   -1  'True
      Tab(0).Control(4)=   "Text1(0)"
      Tab(0).Control(4).Enabled=   -1  'True
      Tab(0).Control(5)=   "Label17"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5(0)"
      Tab(0).Control(8).Enabled=   0   'False
      TabCaption(1)   =   "Phone"
      TabPicture(1)   =   "NewBook.frx":54EA
      Tab(1).ControlCount=   10
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2(4)"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "Text2(3)"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "Text2(2)"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "Text2(1)"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "Text2(0)"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "Label6(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label6(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label6(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label6(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label6(0)"
      Tab(1).Control(9).Enabled=   0   'False
      TabCaption(2)   =   "Event"
      TabPicture(2)   =   "NewBook.frx":5506
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Combo2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Picture3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Option1(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Option1(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Option1(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         MaxLength       =   52
         MultiLine       =   -1  'True
         TabIndex        =   108
         Top             =   2160
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Year"
         Height          =   315
         Index           =   0
         Left            =   2640
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Month"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   24
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "None"
         Height          =   315
         Index           =   2
         Left            =   2640
         TabIndex        =   23
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2880
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "NewBook.frx":5522
         Left            =   240
         List            =   "NewBook.frx":555C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -74160
         MaxLength       =   25
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -74160
         MaxLength       =   20
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -74160
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -74160
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -74160
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -72360
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -74160
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -74160
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -74160
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label17 
         Caption         =   "Notes"
         Height          =   255
         Left            =   -74760
         TabIndex        =   109
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   20
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Cell"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   19
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Fax"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Phone2"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Phone1"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Add3"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Add2"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Add1"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   20
      TabIndex        =   26
      Text            =   "Text3"
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   600
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2655
      Left            =   4680
      ScaleHeight     =   2595
      ScaleWidth      =   3645
      TabIndex        =   28
      Top             =   600
      Width           =   3705
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         ToolTipText     =   " Next Month "
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         ToolTipText     =   " Next Year "
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   " Last Month "
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   " Last Year "
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Today"
         Height          =   255
         Left            =   1200
         TabIndex        =   29
         ToolTipText     =   " Current Date "
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   83
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "42"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   3000
         TabIndex        =   82
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "38"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   1080
         TabIndex        =   81
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "39"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   1560
         TabIndex        =   80
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "41"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   2520
         TabIndex        =   79
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "40"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   2040
         TabIndex        =   78
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "37"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   600
         TabIndex        =   77
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   76
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "35"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   3000
         TabIndex        =   75
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "34"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   2520
         TabIndex        =   74
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "33"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   2040
         TabIndex        =   73
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   1560
         TabIndex        =   72
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "31"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   1080
         TabIndex        =   71
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   600
         TabIndex        =   70
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "29"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   69
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "28"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   3000
         TabIndex        =   68
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "27"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   2520
         TabIndex        =   67
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "26"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   2040
         TabIndex        =   66
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   1560
         TabIndex        =   65
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   1080
         TabIndex        =   64
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "23"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   63
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "22"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "21"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   3000
         TabIndex        =   61
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   2520
         TabIndex        =   60
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "19"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   2040
         TabIndex        =   59
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "18"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   1560
         TabIndex        =   58
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "17"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   1080
         TabIndex        =   57
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   600
         TabIndex        =   56
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   54
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "13"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   2520
         TabIndex        =   53
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   2040
         TabIndex        =   52
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   51
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   50
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   49
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   47
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   46
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   45
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   44
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   43
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   42
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   39
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   38
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   37
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thu"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   36
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   35
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   34
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8190
      TabIndex        =   97
      Top             =   4680
      Width           =   8250
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   6480
      TabIndex        =   84
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Reminders:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4800
      TabIndex        =   22
      Top             =   120
      Width           =   1455
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8454143
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":5620
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":6272
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":6EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":7B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":8768
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":93BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":A00C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":AC5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":B8B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":C502
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":D154
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":DDA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":E9F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":F64A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":1029C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":10EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":11B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":12792
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":133E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":14036
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewBook.frx":14C88
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Thanks to all the people whose programs gave me ideas
'  and ways to make my code possible.
'
' Special thanks to MarkB - A great programmer !
' 000315

Private Sub Form_Load()
Dim f As Integer

Open App.Path & " \dbData._db" For Random As #1 Len = 239
Get #1, 1, hd

If hd.hdLast1 = 0 Then
   Me.Left = (Screen.Width - Me.Width) / 4
   Me.Top = (Screen.Height - Me.Height) / 2
Else
   Me.Top = hd.hdLast1
   Me.Left = hd.hdLast2
End If

If hd.hdTotal = 0 Then TempData
   dbTotal = hd.hdTotal: dbCurrent = 1
   TempId = DriveInfo
   
If TempId <> DeCode(hd.hdDiskID) Or DeCode(hd.hdRegOK) = "NoDrive" Then
   hd.hdUser = EnCode("Unregistered!")
   Put #1, 1, hd
If dbTotal > 5 Then dbTotal = 5
End If

For f = 1 To dbTotal
   dbCurrent = f
   Get #1, dbCurrent + 1, db
   Combo1.AddItem db.dbName
Next

UpDateRecord

FillList
dbCurrent = 1
LockAll
Command12_Click

End Sub

Public Sub RefreshList()

Get #1, dbCurrent + 1, db
Text1(0).Text = Trim(db.dbAdd1): Text1(1).Text = Trim(db.dbAdd2)
 Text1(2).Text = Trim(db.dbAdd3): Text1(3).Text = Trim(db.dbAdd4)
  Text2(0) = Trim(db.dbTel1): Text2(1) = Trim(db.dbTel2)
Text2(2) = Trim(db.dbTel3): Text2(3) = Trim(db.dbTel4)
 Text2(4) = Trim(db.dbTel5)
  Text3.Text = Trim(db.dbName)
Text6.Text = db.dbNotes
Check1.Value = Val(db.dbDelEv)

Combo2.ListIndex = Val(db.dbEvent)

For f = 0 To 2
If f = Val(db.dbWhen) Then
   Option1(f).Enabled = True
   Option1(f).Value = True
Else
   Option1(f).Enabled = False
   Option1(f).Value = False
End If
Next

If DeCode(Trim(hd.hdUser)) = "Unregistered!" Then
a$ = " InfoE - Unregistered! "
Else
a$ = " InfoE "
End If
Form1.Caption = a$

TempPic = Val(db.dbEvent) + 1
Picture3.Picture = ImageList1.ListImages(TempPic).Picture

Label4.Caption = Format(db.dbDate, "dd mmmm yyyy")

End Sub

Private Sub Combo1_Click()
dbCurrent = Combo1.ListIndex + 1
RefreshList
End Sub

Private Sub Command10_Click()

If Trim(Text3.Text) = "" Then
   MsgBox "Please enter a Name !", vbInformation + vbOKOnly, " Oops !"
   Text3.SetFocus
Exit Sub
End If

Command12.Caption = "<<"
 Command9.Visible = True
Check1.Visible = False
  
LockAll

db.dbName = Trim(Text3.Text)
 db.dbAdd1 = Trim(Text1(0).Text): db.dbAdd2 = Trim(Text1(1).Text)
  db.dbAdd3 = Trim(Text1(2).Text): db.dbAdd4 = Trim(Text1(3).Text)
db.dbTel1 = Trim(Text2(0)): db.dbTel2 = Trim(Text2(1))
 db.dbTel3 = Trim(Text2(2)): db.dbTel4 = Trim(Text2(3))
  db.dbTel5 = Trim(Text2(4))
db.dbEvent = Combo2.ListIndex
 db.dbDate = NewDate
db.dbNotes = Trim(Text6.Text)
db.dbDelEv = Trim("0" & Check1.Value)

For f = 0 To 2
   If Option1(f) = True Then db.dbWhen = Str$(f)
Next

   If TempEdit = 1 Then GoTo EditSkip
   
dbTotal = dbTotal + 1: dbCurrent = dbTotal

EditSkip:

TempEdit = 0
DataWrite
UpDateList
RefreshList
FillList

End Sub

Private Sub Command11_Click()
Form1.Width = 4680
Command9.Visible = True
Check1.Visible = False
RefreshList
LockAll
End Sub

Private Sub Command13_Click()
   Unload Me
End Sub

Private Sub Command6_Click()
Get #1, 1, hd

If Trim(DeCode(hd.hdUser)) = "Unregistered!" Or DeCode(hd.hdRegOK) = "NoDrive ?" Then
 If dbTotal >= 5 Then
 dbTotal = 5
 Command10.Visible = False
     Command11.Visible = False
 Command12.ToolTipText = " Close reminder list "
  Call Command9_Click
  Exit Sub
End If
 End If
 
OpenAll
NewRecord
Form1.Caption = " Add a record     Today - " & Format(Date$, "dd mmmm yyyy")
Check1.Visible = True
DrawCal (Date$)

End Sub

Private Sub Command7_Click()
Check1.Visible = True
dbCurrent = Combo1.ListIndex + 1
Get #1, dbCurrent + 1, db

Form1.Caption = " Edit record     Today - " & Format(Date$, "dd mmmm yyyy")
OpenAll
Text3.Text = Trim(db.dbName)
Text6.Text = Trim(db.dbNotes)
DrawCal (db.dbDate)
TempEdit = 1
End Sub

Private Sub Command8_Click()
Dim f As Integer

If Command12.Caption = "<<" Then
Command12.Caption = ">>"
   Form1.Width = 4680
     Command10.Visible = True
   Command11.Visible = True
    Picture4.Visible = False
     Command12.ToolTipText = " Open reminder list "
     End If

If dbTotal <= 1 Then
   MsgBox "You can`t DELETE the last record !", vbInformation + vbOKOnly, " Delete Error !"
   Exit Sub
End If
Form1.Caption = " Delete record "

If MsgBox("Do you want to DELETE this record ?", vbQuestion + vbYesNo, " Delete record ?") = vbNo Then
   Exit Sub
End If

DelRecord

Combo1.Clear

Get #1, 1, hd

If hd.hdTotal = 0 Then TempData

dbTotal = hd.hdTotal: dbCurrent = 1
TempId = DriveInfo

If TempId <> DeCode(hd.hdDiskID) Then
 If dbTotal > 5 Then dbTotal = 5
End If

For f = 1 To dbTotal
   dbCurrent = f
   Get #1, dbCurrent + 1, db
   Combo1.AddItem db.dbName
Next

FillList

dbCurrent = 1
RefreshList
Combo1.ListIndex = dbCurrent - 1
End Sub

Private Sub Command12_Click()

If Command12.Caption = ">>" Then
   Form1.Width = 8745
     Command12.Caption = "<<"
   Command12.ToolTipText = " Close reminder list "
    Command10.Visible = False
     Command11.Visible = False
Else
   Command12.Caption = ">>"
   Form1.Width = 4680
    
     Command10.Visible = True
   Command11.Visible = True
    Picture4.Visible = False
     Command12.ToolTipText = " Open reminder list "
End If

End Sub

Private Sub Label1_Click(Index As Integer)
Dim f As Integer

If Index < DayOfWeek Then GoTo CheckMNeg
 If Index - DayOfWeek + 1 > LastDay Then GoTo CheckMPos

For f = DayOfWeek To LastDay + DayOfWeek - 1
   Label1(f).BackColor = &HC0C0FF
Next

Label1(Index).BackColor = &H8080FF
TempDate = Left$(NewDate, 8) & Index - DayOfWeek + 1
NewDate = TempDate
Label3.Caption = Format(NewDate, "dd mmmm yyyy")
 Label4.Caption = Label3.Caption

Exit Sub

CheckMNeg:
a = Val(Label1(Index).Caption)
TempDate = DateAdd("m", -1, NewDate)
TempDate = Left(TempDate, 8) & a
 DrawCal (TempDate)
Exit Sub

CheckMPos:
a = Val(Label1(Index).Caption)
TempDate = DateAdd("m", 1, NewDate)
TempDate = Left(TempDate, 8) & a
 DrawCal (TempDate)

End Sub

Public Sub DrawCal(TempDate)
Dim f As Integer

Label4.Caption = Format(TempDate, "dd mmmm yyyy")
List1.Visible = False
Command9.Visible = False
Form1.Width = 8745

NewDate = Format(TempDate, "yyyy-mm-dd")
StartDate = DateSerial(Year(NewDate), Month(NewDate), 1)
LastDay = DateDiff("d", StartDate, DateAdd("m", 1, StartDate))
DayOfWeek = WeekDay(StartDate) - 1: TDay = 0

For f = 0 To 41
   Label1(f).ForeColor = &H0
   Label1(f).BackColor = &HC0C0FF
   Label1(f).FontBold = False
   Label1(f).FontUnderline = False
    
If f >= DayOfWeek Then
   TDay = TDay + 1
If TDay > LastDay Then GoTo Skip
   Label1(f).Caption = TDay

If TDay = Day(Now) And Val(Mid$(NewDate, 6, 2)) = Month(Now) Then
   Label1(f).ForeColor = &HC00000
   Label1(f).FontBold = True
   Label1(f).FontUnderline = True
   End If

If TDay = Day(NewDate) Then Label1(f).BackColor = &H8080FF

End If

Skip:
Next

LastM = Abs(DateDiff("d", StartDate, DateAdd("m", -1, StartDate)))
For f = DayOfWeek - 1 To 0 Step -1
   Label1(f).Caption = LastM
   Label1(f).BackColor = &HC0E0FF
   LastM = LastM - 1
Next

For f = LastDay + DayOfWeek To 41
   NewM = NewM + 1
   Label1(f).Caption = NewM
   Label1(f).BackColor = &HC0E0FF
Next

Label3.Caption = Format(NewDate, "dd mmmm yyyy")

End Sub

Private Sub Command1_Click()
TempDate = DateAdd("yyyy", -1, NewDate)
 DrawCal (TempDate)
End Sub

Private Sub Command4_Click()
TempDate = DateAdd("yyyy", 1, NewDate)
 DrawCal (TempDate)
End Sub

Private Sub Command2_Click()
TempDate = DateAdd("m", -1, NewDate)
 DrawCal (TempDate)
End Sub

Private Sub Command3_Click()
TempDate = DateAdd("m", 1, NewDate)
 DrawCal (TempDate)
End Sub

Private Sub Command5_Click()
NewDate = Format(Date, "yyyy-mm-dd")
 DrawCal (Date)
End Sub

Private Sub Combo2_Click()
TempPic = Combo2.ListIndex + 1
Picture3.Picture = ImageList1.ListImages(TempPic).Picture
End Sub

Private Sub Command9_Click()

If Picture4.Visible = True Then
If Text5.Visible = True Then Text5.SetFocus
Exit Sub
End If

Timer1.Enabled = True
 Text5.Text = ""
Command12.Caption = "<<"
 Command10.Visible = False
  Command11.Visible = False
                   
Picture4.Visible = True
Label10.Caption = DeCode(Trim(hd.hdUser))

If DeCode(Trim(hd.hdUser)) = "Unregistered!" Then
   Picture6.ToolTipText = " Send E-Mail to Register free ! "
    Picture6.Picture = ImageList1.ListImages(19).Picture
   Command14.Caption = " Register "
   Label13.Visible = True
    Label14.Visible = True
   Text4.Visible = True: Text4.Text = DeCode(Trim(hd.hdUser))
    Text5.Visible = True
     Text5.SetFocus
Else
   Picture6.ToolTipText = " Thanks for Registering ! "
   Picture6.Picture = ImageList1.ListImages(20).Picture
   Command14.Caption = "OK"
   Label13.Visible = False
    Label14.Visible = False
   Text4.Visible = False
    Text5.Visible = False
End If

Form1.Width = 8745

End Sub

Public Sub TempData()

db.dbName = "Add Name"
 db.dbAdd1 = "Add Add1": db.dbAdd2 = "Add Add2": db.dbAdd3 = "Add Add3"
  db.dbAdd4 = "1234"
db.dbTel1 = "Add Tel1": db.dbTel2 = "Add Tel2"
 db.dbTel3 = "Add Fax": db.dbTel4 = "Add Cell"
  db.dbTel5 = "Add E-Mail.co.za"
db.dbEvent = "01": db.dbDate = Date
 db.dbWhen = "00"
 db.dbNotes = "Extra notes,comments or ideas."
  db.dbDelEv = "00"
  
  dbTotal = 1: dbCurrent = 1
Put #1, 2, db
hd.hdDiskID = EnCode("23ED-911F")
 hd.hdUser = EnCode("Unregistered!")
  hd.hdRegOK = EnCode("NoDrive ?")
   hd.hdSdate = EnCode(Date$)
WriteHeader
End Sub

Public Sub DataWrite()
Put #1, dbCurrent + 1, db
WriteHeader
SortData
End Sub

Public Sub WriteHeader()
hd.hdDate = EnCode(Date$)
 hd.hdTime = EnCode(Time)
  hd.hdTotal = dbTotal
Put #1, 1, hd
End Sub

Public Sub SortData()
Dim i As Integer, Boundry As Integer, Center As Integer

Dim db1 As dbRecord, db2 As dbRecord

Center = dbTotal / 2
Do While Center > 0
Boundry = (dbTotal + 1) - Center
Do
Flag = 0
For i = 2 To Boundry
Get #1, i, db1: Get #1, i + Center, db2

If UCase(Trim(db1.dbName)) > UCase(Trim(db2.dbName)) Then
   Put #1, i, db2: Put #1, i + Center, db1
   Flag = i
End If
Next i

Boundry = Flag - Center
Loop While Flag
Center = Center \ 2
Loop
dbCurrent = 1
RefreshList

End Sub

Public Sub UpDateList()
Dim f As Integer

Combo1.Clear
For f = 1 To dbTotal
   dbCurrent = f
   Get #1, dbCurrent + 1, db
   Combo1.AddItem db.dbName
Next

End Sub

Private Sub Label15_Click()
Dim RetVal

 a = UCase(App.Path)
 
 If DeCode(Trim(hd.hdRegOK)) = "NoDrive ?" Then
   RetVal = Shell("start mailto:softeny@netactive.co.za?&cc=""&subject=Registration Request&body=Please send me your Registration Name and a copy of dbData._db found in " & a, vbNormal)
 Else
   RetVal = Shell("start mailto:softeny@netactive.co.za?&cc=""&subject=TBook Questions?&body=Thanks for Registering.", vbNormal)
 End If
 
End Sub

Private Sub List1_Click()

dbCurrent = Val(Right$(List1.Text, 2))
Combo1.ListIndex = dbCurrent - 1
TempDb = Format(db.dbDate, "mm-dd-") & Right$(Date$, 4)

DaysNow = DateDiff("d", Now, TempDb)
 DayA = WeekDay(Date$, vbMonday)
  DayB = WeekDay(TempDb, vbMonday)
WeekA = DateDiff("w", Now, vbMonday) - DateDiff("w", TempDb, vbMonday)

If DaysNow = -1 Then
   Label9.Caption = "Yesterday"
ElseIf DaysNow < -1 Then
   Label9.Caption = "Passed"
ElseIf DaysNow = 0 Then
   Label9.Caption = "Today"
ElseIf DaysNow = 1 Then
   Label9.Caption = "Tomorrow"
ElseIf WeekA = 0 Then
   Label9.Caption = "This Week"
ElseIf WeekA = 1 Then
   Label9.Caption = "Next Week"
Else
   Label9.Caption = "This Month"
End If

End Sub

Public Sub FillList()
Dim f As Integer

List1.Clear
For f = 1 To dbTotal
   dbCurrent = f: Togo = ""
   Get #1, dbCurrent + 1, db
   DaysNow = DateDiff("d", Now, Format(db.dbDate, "mm-dd"))
   BirthYear = DateDiff("yyyy", Now, db.dbDate)
   
If BirthYear >= 0 And BirthYear <= 9 Then BirthYear = " 0" & Trim(Str(BirthYear))
 If BirthYear <= -1 And BirthYear >= -9 Then BirthYear = "-0" & Abs(BirthYear)
  If BirthYear >= 0 Then BirthYear = " " & Abs(BirthYear)
   If BirthYear = 0 Then BirthYear = " 00"
 
 If Month(db.dbDate) = Month(Now) Then
 If Val(db.dbWhen) <= 1 And DaysNow >= -1 Then
 If DaysNow >= 10 Or DaysNow <= -1 Then
   Togo = Trim(Str$(DaysNow))
 ElseIf DaysNow <= 9 And DaysNow >= 0 Then
 Togo = "0" & Trim(Str(DaysNow))
 End If
   GoTo AddList
ElseIf Val(db.dbWhen) <= 1 And DaysNow <= 0 Then
   Togo = " *"
   GoTo AddList
End If
 End If

GoTo SkipNext
AddList:
List1.AddItem Format(db.dbDate, "dd-mm") & " " & db.dbName & " " & Togo & " " & BirthYear & "    " & Str(f)
SkipNext:
Next

FirstList

End Sub

Public Sub FirstList()
Dim f As Integer

TempList = 0
If List1.ListCount = 0 Then
Label9.Caption = "None"
Combo1.ListIndex = 0
Exit Sub
End If

For f = 0 To List1.ListCount - 1
If TempList = 1 Then GoTo JumpList
 If Val(Left$(List1.List(f), 2)) = Day(Now) Then
   a = f
   TempList = 1
End If
Next

JumpList:
   List1.ListIndex = a
End Sub

Private Sub Command14_Click()
Dim RealReg&, TempReg, RealID

If Command14.Caption = "OK" Then
   Picture4.Visible = False
Exit Sub
End If

If Text4.Text = "" Or Text5.Text = "" Then
   Picture4.Visible = False
   Exit Sub
End If

RealReg& = 0
TempReg = Format(DeCode(hd.hdSdate), "dd mmmm yyyy")
For f = 1 To Len(TempReg)
   RealReg& = RealReg& + Asc(Mid$(TempReg, f, 1)) * Asc(Left(Text4.Text, 1)) + 12345
Next

RealID = DiskInfo
If Text5.Text <> Trim(Str(RealReg&)) Then
   Picture4.Visible = False
Exit Sub
Else
   Picture4.Visible = False
hd.hdDiskID = EnCode(DriveInfo)
 hd.hdUser = EnCode(Trim(Text4.Text))
  hd.hdRegOK = EnCode("Drive C:\")
Put #1, 1, hd
End If
   Form1.Caption = " InfoE "
   End Sub

Private Sub List1_DblClick()
Form2.Show 1
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then Form2.Show
Pm = Int(Y)
End Sub

Private Sub Timer1_Timer()

If cl >= 10 Then cl = 2
Label16.ForeColor = QBColor(cl)

If Picture4.Visible = True Then
   Timer1.Enabled = True
Else
   Timer1.Enabled = False
   Label16.Left = 3800
End If

Label16.Left = Label16.Left - 30
cl = cl + 2
If Label16.Left <= -2500 Then
   Label16.Left = 3800
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
hd.hdLast1 = Form1.Top
 hd.hdLast2 = Form1.Left
Put #1, 1, hd
Close #1
   Unload Me
End Sub


Public Sub DelRecord()
Dim dbCurrent As Integer, TempCurrent As Integer

Open App.Path & "\dbData.bak" For Random As #2 Len = 239
dbCurrent = 1: TempCurrent = 1
Do While dbCurrent < dbTotal + 1
If dbCurrent <> Combo1.ListIndex + 1 Then
   Get #1, dbCurrent + 1, db
   Put #2, TempCurrent + 1, db
   TempCurrent = TempCurrent + 1
End If
dbCurrent = dbCurrent + 1
Loop
dbTotal = dbTotal - 1
WriteHeader
Put #2, 1, hd
Close #1: Close #2

SourceF = App.Path & "\dbData.bak"
 DestnaF = App.Path & "\dbData._db"
  FileCopy SourceF, DestnaF
  
Open App.Path & "\dbData._db" For Random As #1 Len = 239

End Sub

Public Sub UpDateRecord()
Dim f As Integer

Get #1, 1, hd
dbTotal = hd.hdTotal: dbCurrent = 1

ReStart:
If dbTotal <= 1 Then Exit Sub

For f = 1 To dbTotal
Get #1, f + 1, db
DelDate = DateDiff("d", Now, db.dbDate)

If db.dbDelEv = "01" And DelDate <= -2 Then
Combo1.ListIndex = f - 1

DelRecord
GoTo ReStart
End If
Next f

Combo1.Clear

For f = 1 To dbTotal
   dbCurrent = f
   Get #1, dbCurrent + 1, db
   Combo1.AddItem db.dbName
Next

Combo1.ListIndex = 0

End Sub
