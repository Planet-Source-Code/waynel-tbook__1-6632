VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " InfoE"
   ClientHeight    =   2655
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4110
   Icon            =   "NewBook2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hd As dbHeader

Private Sub Form_Load()
Open App.Path & " \dbData._db" For Random As #2 Len = 185
Get #2, 1, hd
Text1.Text = hd.hdUser: Text2.Text = hd.hdDiskID

If Trim(hd.hdUser) = "Unregistered" Then
   Picture2.Picture = Form1.ImageList1.ListImages(19).Picture
   Label4.Caption = "Unregistered !"
   Command1.Visible = True: Command2.Caption = "Cancel"
   Form2.Height = 3030
Else
   Command1.Visible = False: Command2.Caption = "OK"
   Label4.Caption = Trim(hd.hdUser)
   Picture2.Picture = Form1.ImageList1.ListImages(20).Picture
   Form2.Height = 1845
End If

Form1.Enabled = False
Me.Top = Form1.Top + 600
Me.Left = Form1.Left - 340

TempId = DriveInfo
End Sub

Private Sub Command1_Click()
Dim RegNum&, TempReg, NewID

NewID = DriveInfo

If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Then
   Unload Me
   Exit Sub
End If

TempReg = Format(Date$, "dd mmmm yyyy")

For f = 1 To Len(TempReg)
   RegNum& = RegNum& + Asc(Mid$(TempReg, f, 1)) * Asc(Left$(Text1.Text, 1)) + 12345
Next f

T$ = Trim(Str$(RegNum&))
If Trim(Text2.Text) <> T$ Then
   Unload Me
Else
   Form1.Caption = " InfoE "
   hd.hdDiskID = Trim(NewID)
   hd.hdUser = Trim(Text1.Text)
   Put #2, 1, hd
End If
   Unload Me
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #2
Form1.Enabled = True
End Sub

Private Sub Label6_Click()
Dim RegNum&, TempReg

TempReg = Format(Date$, "dd mmmm yyyy")
If Trim(Text1.Text) = "" Then Exit Sub
If Left(Text1.Text, 1) = "`" Then
For f = 1 To Len(TempReg)
   RegNum& = RegNum& + Asc(Mid$(TempReg, f, 1)) * Asc(Mid$(Text1.Text, 2, 1)) + 12345
Next f

Text2.Text = Trim(Str$(RegNum&))
Text1.Text = Right(Text1.Text, Len(Text1.Text) - 1)

End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Trim(hd.hdUser) = "Unregistered" Then
   Picture2.ToolTipText = " Please Register ! "
Else
   Picture2.ToolTipText = " Thanks for Registering ! "
End If
End Sub
