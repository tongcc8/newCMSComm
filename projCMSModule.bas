Attribute VB_Name = "projCMSModule"
Option Explicit


'API declarations
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$) As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Long, ByVal FileName$) As Long

Public Lo
Public fMainForm As frmMain


Public Sub ReadFromFile(ByVal fP As Form)
'read .ini file
Dim TextInfo$   'holds text from INI file
Dim res         'holds results
Dim i, j As Integer
Dim iniFName$

  
  'ini file name
  iniFName$ = "projCMSComm.ini"

  'change cursor
  Screen.MousePointer = vbHourglass
  
  DoEvents
  
  'read no of sensor
  'TextInfo$ = Space(80)
  'res = GetPrivateProfileString("NOSENSOR", "No", "", TextInfo$, 100, iniFName)
  'fP.txtMonitor.SelText = fP.txtMonitor.SelText + "No of Sensor: " & Left$(TextInfo$, res) + vbCrLf
  'NoOfRadar = Val(Left$(TextInfo$, res))
  
  'read radar name
  'ReDim RadarName(1 To NoOfRadar, 2) As String
  'For i = 1 To NoOfRadar
  '  TextInfo$ = Space(80)
  '  res = GetPrivateProfileString("SENSOR", "Name" & Trim$(Str$(i)), "", TextInfo$, 100, iniFName)
  '  fP.txtMonitor.SelText = fP.txtMonitor.SelText + "Radar " & Trim$(Str$(i)) & " : " & Left$(TextInfo$, res) + vbCrLf
  '
  '  j = InStr(1, Left$(TextInfo$, res), ",", vbBinaryCompare)
  '  If j > 0 Then
  '    'RadarName(i, 1) = Mid$(Left$(TextInfo$, res), 1, j - 1)
  '    RadarName(i, 2) = Trim$(Mid$(Left$(TextInfo$, res), j + 1))
  '    If Val(RadarName(i, 2)) < 1 Then
  '      NoOfFilter = NoOfFilter + 1
  '    End If
  '  Else
  '    RadarName(i, 1) = ""
  '    RadarName(i, 2) = ""
  '  End If
  'Next
  
  
  'change cursor
  Screen.MousePointer = vbDefault
  
End Sub



Sub main()

  'change working directory to the directory wher the application was executed.
  ChDrive App.Path
  ChDir App.Path

  Set fMainForm = New frmMain
  fMainForm.Show
  
End Sub
