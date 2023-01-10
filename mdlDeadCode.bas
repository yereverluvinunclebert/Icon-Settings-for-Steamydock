Attribute VB_Name = "mdlDeadCode"
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_CURRENT_USER = &H80000001
'Public Const REG_SZ = 1                          ' Unicode nul terminated string

'Public rdSettingsFile As String
'Public dockSettingsFile As String
'Public origSettingsFile As String
'Public toolSettingsFile  As String
'Public WindowsVer As String
'Public requiresAdmin As Boolean
'Public rdAppPath As String
'Public RDinstalled As String
'Public RD86installed As String
'Public rocketDockInstalled As Boolean
'Public RDregistryPresent As Boolean
'Public rDCustomIconFolder As String ' .NET

'Public rDGeneralReadConfig As String
'Public rDGeneralWriteConfig As String

'Public Enum eSpecialFolders
'  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
'  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
'  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
'  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
'End Enum

''API Function to read information from INI File
'Public Declare Function GetPrivateProfileString Lib "kernel32" _
'    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
'    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
'    , ByVal lpFileName As String) As Long
'
''API Function to write information to the INI File
'Private Declare Function WritePrivateProfileString Lib "kernel32" _
'    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
'    , ByVal lpString As Any, ByVal lpFileName As String) As Long
    
'Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Public Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

'Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

'Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
'Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
'Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
'Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
    
'Public sFilename  As String
'Public sFileName2  As String
'Public sTitle  As String
'Public sCommand  As String
'Public sArguments  As String
'Public sWorkingDirectory  As String
'Public sShowCmd  As String
'Public sOpenRunning  As String
'Public sIsSeparator  As String
'Public sUseContext  As String
'Public sDockletFile  As String

''Get the INI Setting from the File
''---------------------------------------------------------------------------------------
'' Procedure : GetINISetting
'' Author    : beededea
'' Date      : 05/07/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function GetINISetting(ByVal sHeading As String, ByVal sKey As String, ByRef sINIFileName As String) As String
'   On Error GoTo GetINISetting_Error
'    Const cparmLen = 500 ' maximum no of characters allowed in the returned string
'    Dim sReturn As String * cparmLen
'    Dim sDefault As String * cparmLen
'    Dim lLength As Long
'
'    lLength = GetPrivateProfileString(sHeading, sKey, sDefault, sReturn, cparmLen, sINIFileName)
'    GetINISetting = mid$(sReturn, 1, lLength)
'
'   On Error GoTo 0
'   Exit Function
'
'GetINISetting_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetINISetting of Module Module2"
'End Function

''Save INI Setting in the File
''---------------------------------------------------------------------------------------
'' Procedure : PutINISetting
'' Author    : beededea
'' Date      : 05/07/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)
'
'   On Error GoTo PutINISetting_Error
'
'    Dim aLength As Long
'
'    aLength = WritePrivateProfileString(sHeading, sKey, sSetting, sINIFileName)
'
'   On Error GoTo 0
'   Exit Sub
'
'PutINISetting_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutINISetting of Module Module2"
'End Sub
'Sub findIconMax()
'
'    rdIconMaximum = GetINISetting("Software\RocketDock\Icons", "count", rdAppPath & "\SETTINGS.INI")
'
'    'Reads a INI File (SETTINGS.INI)
'    'For useloop = 0 To 500 ' the current maximum
'        'sFileName(useloop) = GetINISetting("Software\RocketDock\Icons", useloop & "-FileName", rdAppPath & "\SETTINGS.INI")
'        'If sFileName(useloop) = "" Then
'        '    Exit Sub
'        '    rdIconMaximum = useloop ' obtain the number of the last icon in the settings file
'        'End If
'
'End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : writeIconSettingsIni
'' Author    : beededea
'' Date      : 10/05/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Sub writeIconSettingsIni(location As String, iconNumberToWrite As Integer, settingsFile As String)
'    'Writes an .INI File (SETTINGS.INI)
'
'   On Error GoTo writeIconSettingsIni_Error
'   If debugflg = 1 Then Debug.Print "%writeIconSettingsIni"
'
'        sFilenameCheck = sFilename  ' debug 01
'
'        PutINISetting location, iconNumberToWrite & "-FileName", sFilename, settingsFile
'        PutINISetting location, iconNumberToWrite & "-FileName2", sFileName2, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Title", sTitle, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Command", sCommand, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Arguments", sArguments, settingsFile
'        PutINISetting location, iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory, settingsFile
'        PutINISetting location, iconNumberToWrite & "-ShowCmd", sShowCmd, settingsFile
'        PutINISetting location, iconNumberToWrite & "-OpenRunning", sOpenRunning, settingsFile
'        PutINISetting location, iconNumberToWrite & "-IsSeparator", sIsSeparator, settingsFile
'        PutINISetting location, iconNumberToWrite & "-UseContext", sUseContext, settingsFile
'        PutINISetting location, iconNumberToWrite & "-DockletFile", sDockletFile, settingsFile
'
'        'test to see if the icon path has been truncated
'        sFilename = GetINISetting(location, iconNumberToWrite & "-FileName", settingsFile)
'        If sFilenameCheck <> "" Then
'            If sFilename <> sFilenameCheck Then
'                MsgBox " that strange truncated filename bug encountered, check " & settingsFile & " now and look for " & sFilenameCheck
'            End If
'        End If
'
'   On Error GoTo 0
'   Exit Sub
'
'writeIconSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeIconSettingsIni of Module Module2"
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : writeSettingsIni
' Author    : beededea
' Date      : 21/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Sub writeSettingsIni(ByVal iconNumberToWrite As Integer)
'    'Writes an .INI File (SETTINGS.INI)
'
'    'E:\Program Files (x86)\RocketDock\Icons\Steampunk_Clockwerk_Kubrick
'    ' determine relative path TODO
'    ' Icons\Steampunk_Clockwerk_Kubrick
'
'   On Error GoTo writeSettingsIni_Error
'   If debugflg = 1 Then DebugPrint "%writeSettingsIni"
'
'        sFilenameCheck = sFilename  ' debug 01
'
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", sFilename, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", sFileName2, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Title", sTitle, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Command", sCommand, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", sArguments, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", sShowCmd, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", sOpenRunning, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", sIsSeparator, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", sUseContext, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", sDockletFile, rdSettingsFile
'
'
'        sFilename = GetINISetting("Software\RocketDock\Icons", iconNumberToWrite & "-FileName", rdSettingsFile)
'        If sFilenameCheck <> "" Then
'            If sFilename <> sFilenameCheck Then
'                MsgBox " that strange filename bug encountered, check rdSettings.ini now and look for " & sFilenameCheck
'            End If
'        End If
'
'   On Error GoTo 0
'   Exit Sub
'
'writeSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeSettingsIni of Module Module2"
'
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : removeSettingsIni
'' Author    : beededea
'' Date      : 21/09/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub removeSettingsIni(ByVal iconNumberToWrite As Integer)
'
'    'removes data from the ini file at the given location
'
'   On Error GoTo removeSettingsIni_Error
'   If debugflg = 1 Then DebugPrint "%removeSettingsIni"
'
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Title", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Command", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", vbNullString, rdSettingsFile
'
'   On Error GoTo 0
'   Exit Sub
'
'removeSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure removeSettingsIni of Module Module2"
'
'End Sub



