VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Calendar"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   3705
      TabIndex        =   0
      Top             =   0
      Width           =   3705
      Begin VB.CommandButton Command10 
         Height          =   375
         Left            =   3120
         Picture         =   "NewBookB.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "  Close "
         Top             =   1800
         Width           =   475
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   3135
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
         TabIndex        =   49
         Top             =   0
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
         TabIndex        =   48
         Top             =   0
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
         TabIndex        =   47
         Top             =   0
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
         TabIndex        =   46
         Top             =   0
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
         TabIndex        =   45
         Top             =   0
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
         TabIndex        =   44
         Top             =   0
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
         TabIndex        =   43
         Top             =   0
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
         TabIndex        =   42
         Top             =   240
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
         TabIndex        =   41
         Top             =   240
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
         TabIndex        =   40
         Top             =   240
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
         TabIndex        =   39
         Top             =   240
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
         TabIndex        =   38
         Top             =   240
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
         TabIndex        =   37
         Top             =   240
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
         TabIndex        =   36
         Top             =   240
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
         TabIndex        =   35
         Top             =   480
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
         TabIndex        =   34
         Top             =   480
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
         TabIndex        =   33
         Top             =   480
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
         TabIndex        =   32
         Top             =   480
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
         TabIndex        =   31
         Top             =   480
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
         TabIndex        =   30
         Top             =   480
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
         TabIndex        =   29
         Top             =   480
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
         TabIndex        =   28
         Top             =   720
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
         TabIndex        =   27
         Top             =   720
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
         TabIndex        =   26
         Top             =   720
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
         TabIndex        =   25
         Top             =   720
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
         TabIndex        =   24
         Top             =   720
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
         TabIndex        =   23
         Top             =   720
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
         TabIndex        =   22
         Top             =   720
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
         TabIndex        =   21
         Top             =   960
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
         TabIndex        =   20
         Top             =   960
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
         TabIndex        =   19
         Top             =   960
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
         TabIndex        =   18
         Top             =   960
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
         TabIndex        =   17
         Top             =   960
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
         TabIndex        =   16
         Top             =   960
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
         TabIndex        =   15
         Top             =   960
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
         TabIndex        =   14
         Top             =   1200
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
         TabIndex        =   13
         Top             =   1200
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
         TabIndex        =   12
         Top             =   1200
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
         TabIndex        =   11
         Top             =   1200
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
         TabIndex        =   10
         Top             =   1200
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
         TabIndex        =   9
         Top             =   1200
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
         TabIndex        =   8
         Top             =   1200
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
         TabIndex        =   7
         Top             =   1440
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
         TabIndex        =   6
         Top             =   1440
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
         TabIndex        =   5
         Top             =   1440
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
         TabIndex        =   4
         Top             =   1440
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
         TabIndex        =   3
         Top             =   1440
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
         TabIndex        =   2
         Top             =   1440
         Width           =   495
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
         TabIndex        =   1
         Top             =   1440
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()

 Me.Top = Form1.List1.Height / 8 + Form1.Top + Pm + 900
   Me.Left = Form1.Left + 4800
   
  Get #1, dbCurrent + 1, db
DrawCal2 (Year(Now) & Right(Format(db.dbDate, "yyyy-mm-dd"), 6))
 End Sub

Public Sub DrawCal2(TempDate)
Dim f As Integer

If db.dbDelEv = "00" Then
Label4.Caption = "Delete one day after due date - [ NO ]"
Else
Label4.Caption = "Delete one day after due date - [ YES ]"
End If

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

End Sub

