Attribute VB_Name = "Module1"
Type dbHeader
 hdTotal As Long
 hdTime As String * 8
 hdDate As String * 10
 hdDiskID As String * 9
 hdUser As String * 20
 hdLast1 As Long
 hdLast2 As Long
 hdRegOK As String * 9
 hdSdate As String * 10
 hdNull As String * 161
End Type

Type dbRecord
 dbName As String * 20
 dbAdd1 As String * 20
 dbAdd2 As String * 20
 dbAdd3 As String * 20
 dbAdd4 As String * 6
 dbTel1 As String * 15
 dbTel2 As String * 15
 dbTel3 As String * 15
 dbTel4 As String * 15
 dbTel5 As String * 25
 dbDate As String * 10
 dbEvent As String * 2
 dbWhen As String * 2
 dbNotes As String * 52
 dbDelEv As String * 2
 End Type
 
Global hd As dbHeader, db As dbRecord
Global dbTotal, dbCurrent
Global NewDate, TempDate, DayOfWeek, LastDay, TempEdit, cl, Pm

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" ( _
 ByVal lpRootPathName As String, _
 ByVal lpVolumeNameBuffer As String, _
 ByVal nVolumeNameSize As Long, _
 lpVolumeSerialNumber As Long, _
 lpMaximumComponentLength As Long, _
 lpFileSystemFlags As Long, _
 ByVal lpFileSystemNameBuffer As String, _
 ByVal nFileSystemNameSize As Long) As Long
Public Function DriveInfo()

VolumeNum = Space$(15): ResStr = Space$(32)
RetVal = GetVolumeInformation("C:\", VolumeNum, Len(VolumeNum), _
 DiskId, 0, 0, ResStr, Len(ResStr))

TempId = Right(String(8, "0") + Hex$(DiskId), 8)
DriveInfo = Left(TempId, 4) + "-" + Right$(TempId, 4)
End Function

Function EnCode(TempTxT)
Dim f As Integer, Temp As String

KeyCode = 5
For f = 1 To Len(TempTxT)
                       
X = Asc(Mid(TempTxT, f, 1))
X = X + KeyCode
If X = 256 Then X = 1
   Temp = Temp & Chr(X)
Next
   EnCode = Temp
End Function

Function DeCode(TempTxT)
Dim f As Integer, Temp As String

KeyCode = 5
For f = 1 To Len(TempTxT)
                       
X = Asc(Mid(TempTxT, f, 1))
X = X - KeyCode
If X = 0 Then X = 255
   Temp = Temp & Chr(X)
Next
   DeCode = Temp
End Function


