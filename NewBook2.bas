Attribute VB_Name = "Module2"
Public Sub LockAll()

With Form1
   .Label8.Visible = True
   .Label9.Visible = True
   .Combo1.Visible = True
   .Combo2.Locked = True
   .List1.Visible = True
   .Command6.Enabled = True
   .Command7.Enabled = True
   .Command8.Enabled = True
   .Command12.Visible = True
   .Command10.Visible = False
   .Command11.Visible = False
   
   .Text1(0).Locked = True: .Text1(1).Locked = True
   .Text1(2).Locked = True: .Text1(3).Locked = True
   .Text2(0).Locked = True: .Text2(1).Locked = True
   .Text2(2).Locked = True: .Text2(3).Locked = True
   .Text2(4).Locked = True: .Text6.Locked = True
End With

End Sub

Public Sub OpenAll()
With Form1
   .Picture4.Visible = False
   .Command12.Caption = ">>"
   .Command10.Visible = True
   .Command11.Visible = True
   .Label8.Visible = False
   .Label9.Visible = False
   .Combo1.Visible = False
   
For f = 0 To 2
   .Option1(f).Enabled = True
Next

   .Combo1.Visible = False
   .Combo2.Locked = False
   .List1.Visible = False
   .Command6.Enabled = False
   .Command7.Enabled = False
   .Command8.Enabled = False
   .Command12.Visible = False

   .Text1(0).Locked = False: .Text1(1).Locked = False
   .Text1(2).Locked = False: .Text1(3).Locked = False
   .Text2(0).Locked = False: .Text2(1).Locked = False
   .Text2(2).Locked = False: .Text2(3).Locked = False
   .Text2(4).Locked = False: .Text6.Locked = False
   .Text3.Locked = False
End With

End Sub

Public Sub NewRecord()
With Form1
   .Text3.Text = ""
   .Text1(0).Text = "": .Text1(1).Text = ""
   .Text1(2).Text = "": .Text1(3).Text = ""
   .Text2(0).Text = "": .Text2(1).Text = "": .Text2(2).Text = ""
   .Text2(3).Text = "": .Text2(4).Text = ""
   .Text6.Text = "": .Check1.Value = 0
      
   .Text3.SetFocus
   .Combo2.ListIndex = 0: .Option1(0) = True
End With

End Sub
