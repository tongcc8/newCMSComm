Attribute VB_Name = "projCSMCommMod"
'update record
'Date: Jan-2001, 28-Dec-2001
'
'
Option Explicit
Option Base 1

'API declarations
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_CLOSE = &H10
Public g_COMM_CMSLocalPortNo    'COMM to CMS local port no, 51200 for HITT
Public g_COMM_OPCLocalPortNo    'COMM to OPC Local Port No
Public g_COMM_CMSTimeSend       'time to send alarm to CMS
Public g_CMSIP                  'CMSIPAddr 10.15.15.43
Public g_OPCIP                  'OPCIPAddr 10.15.15.63
Public g_OPC_COMMLocalPortNo    'OPC to COMM Local Port No
Public g_NoSite                 'no of site
Public g_NoSubSys               'no of subsys
Public g_No_CMSAlm As Integer   'no of CMS Alarm read from .ini file

Public Const MaxSite = 50       'max no of site allow
Public Const MaxSubSys = 50     'max no of subsysstem allow

Public g_SiteIni(MaxSite, 2)          '1 site inital
                                      '2 site description
Public g_SubSysIni(MaxSubSys, 2)      '1 subsys inital
                                      '2 subsys description

Public g_DefAlarm(MaxSite * MaxSubSys, 4)   'Default alarm read from .ini file
                                            '1 site, 2 subsys, 3 status, 4 excel row
Public g_CMSAlarm(MaxSite * MaxSubSys, 3)   'realtime alarm read from CMS Server
                                            '1 site, 2 subsys, 3 status
Public g_SCADAAlarm(MaxSite * MaxSubSys, 3) 'alarm read from SCADA Server
                                            '1 site, 2 subsys, 3 status

Public g_iniFileError As Boolean      'ini file error, if error, connection will be disable

Public fMainForm As frmMain


Public Sub ReadFromFile(ByVal fP As Form)
'read .ini file
Dim TextInfo$   'holds text from INI file
Dim res         'holds results
Dim i, j As Integer
Dim iniFName$
Dim s As String
Dim errMsg As String

  'ini file name
  iniFName$ = App.Path & "\projCMSComm.ini"

  'change cursor
  Screen.MousePointer = vbHourglass
  g_iniFileError = False       'default ini File Error
  errMsg = ""
  
  DoEvents
  
  'read COMM_CMS Local port no, default 55555
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("COMMIP", "COMM_CMSLocalPort", "", TextInfo$, 100, iniFName)
  If res = 0 Then
    errMsg = "COMM for CMS Local port read error"
    GoTo iniError
  End If
  g_COMM_CMSLocalPortNo = Val(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: local port number is " & Trim(Str(g_COMM_CMSLocalPortNo))
  
  'read COMM_OPC Local port no, default 55556
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("COMMIP", "COMM_OPCLocalPort", "", TextInfo$, 100, iniFName)
  If res = 0 Then
    errMsg = "COMM for OPC Local port read error"
    GoTo iniError
  End If
  g_COMM_OPCLocalPortNo = Val(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: local port number is " & Trim(Str(g_COMM_OPCLocalPortNo))
  
  'read CMS IP Address
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("CMSIP", "CMSIPAddr", "", TextInfo$, 100, iniFName)
  If res = 0 Then
    errMsg = "CMS IP read error"
    GoTo iniError
  End If
  g_CMSIP = Trim(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: CMS IP Address is " & g_CMSIP
  
  'read OPC IP Address
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("OPCIP", "OPCIPAddr", "", TextInfo$, 100, iniFName)
  If res = 0 Then
    errMsg = "OPC IP read error"
    GoTo iniError
  End If
  g_OPCIP = Trim(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: CMS IP Address is " & g_OPCIP
  
  'read OPC_COMM Local port no, default 55557
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("OPCIP", "OPC_COMMLocalPort", "", TextInfo$, 100, iniFName)
  If res = 0 Then
    errMsg = "OPC for COMM Local port read error"
    GoTo iniError
  End If
  g_OPC_COMMLocalPortNo = Val(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: local port number is " & Trim(Str(g_OPC_COMMLocalPortNo))
    
  'read time to send alarm to CMS Server
  TextInfo$ = Space(80)
  If res = 0 Then
    errMsg = "OPC for COMM Local port read error"
    GoTo iniError
  End If
  res = GetPrivateProfileString("TIMESEND", "COMM_CMSTime", "", TextInfo$, 100, iniFName)
  g_COMM_CMSTimeSend = Val(Left$(TextInfo$, res)) * 1000
  fP.DispMsg "Read initial file: time to send alarm to CMS is " & Trim(Str(g_COMM_CMSTimeSend / 1000)) & "sec"
  
  'read site no
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("NO_SITE", "No", "", TextInfo$, 100, iniFName)
  If res = 0 Then
    errMsg = "No of site read error"
    GoTo iniError
  End If
  g_NoSite = Val(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: no of site is " & Trim(Str(g_NoSite))
  
  'read site inital
  For i = 1 To g_NoSite
    If i <= 9 Then
      s = "0" & Trim(Str(i))
    Else
      s = Trim(Str(i))
    End If
    
    TextInfo$ = Space(80)
    res = GetPrivateProfileString("SITE", "Name" & s, "", TextInfo$, 20, iniFName)
    If res = 0 Then
      errMsg = "Site initial read error"
      GoTo iniError
    End If
    g_SiteIni(i, 1) = Mid(Left$(TextInfo$, res), 1, 3)
    g_SiteIni(i, 2) = Mid(Left$(TextInfo$, res), 4)
    fP.DispMsg "Read initial file: site name is " & g_SiteIni(i, 1) & " " & g_SiteIni(i, 2)
  Next i
  
  'read no of subsys
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("NO_SUBSYS", "No", "", TextInfo$, 100, iniFName)
  If res = 0 Then
    errMsg = "No of Subsystem read error"
    GoTo iniError
  End If
  g_NoSubSys = Val(Left$(TextInfo$, res))
  fP.DispMsg "Read initial file: no of subsys is " & Trim(Str(g_NoSubSys))
  
  'read subsys initial
  For i = 1 To g_NoSubSys
    If i <= 9 Then
      s = "0" & Trim(Str(i))
    Else
      s = Trim(Str(i))
    End If
    
    TextInfo$ = Space(80)
    res = GetPrivateProfileString("SUBSYS", "Name" & s, "", TextInfo$, 20, iniFName)
    If res = 0 Then
      errMsg = "Subsystem initial read error"
      GoTo iniError
    End If
    g_SubSysIni(i, 1) = Mid(Left$(TextInfo$, res), 1, 3)
    g_SubSysIni(i, 2) = Mid(Left$(TextInfo$, res), 4)
    fP.DispMsg "Read initial file: subsys name is " & g_SubSysIni(i, 1) & " " & g_SubSysIni(i, 2)
  Next i
  
    
  'read CMS TO SCADA ALARM
  Dim p, m, k As Integer
  Dim sINI As String
  TextInfo$ = Space(80)
  res = GetPrivateProfileString("CMS_TO_SCADA", "No", "", TextInfo$, 100, iniFName)
  If res = 0 Then
    errMsg = "CMS_TO_SCADA read error"
    GoTo iniError
  End If
  p = Val(Left$(TextInfo$, res))
  m = 1
  g_No_CMSAlm = 1
  For i = 1 To p
    If i <= 9 Then
      s = "0" & Trim(Str(i))
    Else
      s = Trim(Str(i))
    End If
    
    TextInfo$ = Space(80)
    res = GetPrivateProfileString("CMS", "Name" & s, "", TextInfo$, 80, iniFName)
    If res = 0 Then
      errMsg = "CMS alarm read error"
      GoTo iniError
    End If
    sINI = Mid(Left$(TextInfo$, res), 1, 3)
    j = Val(Mid(Left$(TextInfo$, res), 5, 2))
    For k = 1 To j
      g_DefAlarm(m + k - 1, 1) = sINI
      g_DefAlarm(m + k - 1, 2) = Mid(Left$(TextInfo$, res), 8 + (k - 1) * 4, 3)
      g_DefAlarm(m + k - 1, 3) = "0"
      g_DefAlarm(m + k - 1, 4) = g_No_CMSAlm
      g_No_CMSAlm = g_No_CMSAlm + 1
    Next
    m = m + j
  Next i
  
  g_No_CMSAlm = g_No_CMSAlm - 1       'default no of alarm will be sent from CMS
  fP.DispMsg "No of Status from CMS are " & Trim(Str(g_No_CMSAlm))
  For i = 1 To g_No_CMSAlm
    fP.DispMsg g_DefAlarm(i, 1) & " " & g_DefAlarm(i, 2)
  Next
  
  'change cursor
  Screen.MousePointer = vbDefault
  Exit Sub
  
iniError:
  g_iniFileError = True
  fP.DispMsg ".ini File read error: " & errMsg & ", Program failed to start"
  
End Sub
Sub Main()
Dim i As Integer

  'variable init
  For i = 1 To MaxSite * MaxSubSys
    g_CMSAlarm(i, 1) = "XXX"
    g_CMSAlarm(i, 2) = "XXX"
    g_CMSAlarm(i, 3) = "1"        'default alarm state
    g_SCADAAlarm(i, 1) = "XXX"
    g_SCADAAlarm(i, 2) = "XXX"
    g_SCADAAlarm(i, 3) = "1"      'default alarm state
  Next
  
  
  
  'change working directory to the directory wher the application was executed.
  ChDrive App.Path
  ChDir App.Path
  
  'ChDrive "D\ProjMD\CMSProj\CMSComm"
  'ChDir "D:\ProjMD\CMSProj\CMSComm"
  
  'load frmMain
  Set fMainForm = New frmMain
  Load fMainForm
  fMainForm.Show
  
End Sub

