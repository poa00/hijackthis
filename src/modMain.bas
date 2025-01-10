Attribute VB_Name = "modMain"
'[modMain.bas]

'
' Core check / Fix Engine
'
' (part 1: R0-R4 / F0-F1 / O1 - O24)
' (part 2: see modMain_2.bas)

' Generic format of log line: OX-32 - HKLM\..\Key: [param] = data (sFile), (other marks), FormatSign, (disabled), Checksum

' (c) Fork copyrights:
'
' R4 by Alex Dragokas
' O1 hosts.ics / DNSApi hijackers by Alex Dragokas
' O4 MSconfig and full rework by Alex Dragokas
' O5 Applet by Alex Dragokas
' O7 IPSec / TroubleShooting / Certificates / AppLocker / KnownFolder by Alex Dragokas
' O17 Policy Scripts / DHCP DNS by Alex Dragokas
' O21 ShellIconOverlayIdentifiers / ShellExecuteHooks by Alex Dragokas
' O22 Tasks (Vista+) / .job / BITS by Alex Dragokas
'
' Everything else is based on Merijn Bellekom work

'
' List of all sections:
'

'R0 - Changed Registry value (MSIE)
'R1 - Created Registry value
'R2 - Created Registry key
'R3 - Created extra value in regkey where only one should be
'R4 - IE SearchScopes, DefaultScope
'F0 - Changed inifile value (system.ini)
'F1 - Created inifile value (win.ini)
'B - Browsers
'O1 - Hosts / hosts.ics / DNSApi hijackers
'O2 - BHO (IE Browser Helper Objects)
'O3 - IE Toolbar
'O4 - Reg. autorun entry / msconfig disabled items
'O5 - Control.ini IE Options block / Applet
'O6 - IE Policy: IE Options/Control Panel block
'O7 - Policies / IPSec / TroubleShooting / Certificates / AppLocker / KnownFolder
'O8 - IE Context menu item
'O9 - IE Tools menu item/button
'O10 - Winsock hijack
'O11 - IE Advanced Options group
'O12 - IE Plugin
'O13 - IE DefaultPrefix hijack
'O14 - IERESET.INF hijack
'O15 - Trusted Zone autoadd
'O16 - Downloaded Program Files
'O17 - Domain hijacks / DHCP DNS
'O18 - Protocol & Filter enum & Ports
'O19 - User style sheet hijack
'O20 - AppInit_DLLs registry value + Winlogon Notify subkeys
'O21 - ShellServiceObjectDelayLoad / ShellIconOverlayIdentifiers / ShellExecuteHooks enum
'O22 - SharedTaskScheduler enum / .job / BITS
'O23 - Windows Services
'O24 - Active desktop components
'O25 - Windows Management Instrumentation (WMI) event consumers
'O26 - Image File Execution Options (IFEO) / Tools Hijack
'O27 - Account & Remote Desktop Protocol

'Memo:
'If you have added a new section, you must also:
' - add check procedure CheckOxxItem to (modMain => StartScan)
' - and prefix to Fix module (frmMain => cmdFix_Click)
' - append progressbar max value (frmMain => FormStart_Stage1 => g_HJT_Items_Count)
' - Append label in (modMain => StartScan), see UpdateProgressBar "label"
' - Implement prefix in (modMain => UpdateProgressBar)
' - add translation strings: in # 31, after # 261, 435 (update all files: _Lang_*.lng)
' - add translation index (ModTranslation => GetTranslationIndex_HelpSection)
' - Add section prefix (frmMain => chkHelp_Click)
' - increase (modGlobals => LAST_CHECK_OTHER_SECTION_NUMBER) const (if required)
' - Update sections sort order if required (modMain => SortSectionsOfResultList_Ex)

'Next possible methods:
'* SearchAccurates 'URL' method in a InitPropertyBag (??)
'* HKLM\..\CurrentVersion\ModuleUsage
'* HKLM\..\Internet Explorer\SafeSites (searchaccurate)

'New command line keys:
'
'/noBackup - ��������� �������� ��������� ����� �� ����� �����
'/install /autostart d:X - ��������� � ����������� ������� � �������� �� ������ ����� � ��������� � X ���.
'/instDir:"PATH" - PATH: ���� � �����, ���� ������������ ��������� (�� ���������: "%ProgramFiles(x86)%\HiJackThis Fork").
'/tool:Autoruns - ������ �������� ����� SysInternals Autoruns
'/tool:Executed - ������ �������� ����� NirSoft ExecutedProgramList
'/tool:LastActivity - ������ �������� ����� NirSoft LastActivityView (������ *.exe � *.dll �����)
'/tool:ServiWin - ������ �������� ����� NirSoft ServiWin
'/tool:TaskScheduler - ������ �������� ����� NirSoft TaskSchedulerView
'/sigcheck - ���������� �� �������� ������� Microsoft
'/vtcheck - ���������� �������� ������ �� VT (����� AutoRuns)
'/autofix:vt - ��������� �������� ������ � ��������� �������� VT
'/delmode:pending - �������� ������ ������ � ���������� ������ (����� ������������)
'/delmode:disable - ����������� ��������� ������ ����� ���������� (����� �� ����� �������, ���� ��� ��������)
'/reboot - ������������� �������, ���� ���� ������������� ����� �� ��������
'/addfirewall - �������� AutoRuns � ���������� �����������
'/fixHosts - ��������� ������� Hosts � ����� ���� ������������� (����� ������� �������� �� VT)
'/FixO4 - �������� ��������� ����������� (����� ������� Run* � ����� "������������")
'/FixPolicy - ����� ������� TaskMgr, Regedit, Explorer, TaskBar
'/FixCert - ������ ���������� �� ����� ����������� ���
'/FixIpSec - �������� ������� IP Security (���������� IP / �������� ���� � ������)
'/FixEnvVar - ����������� ������������ �������� ���������� ��������� (%PATH%, %TEMP%)
'/FixO20 - ����� ������ WinLogon � App Init DLLs
'/FixTasks - �������� ������� ������������
'/FixServices - �������� �����
'/FixWMIJob - �������� ������� WMI
'/FixIFEO - �������� ������ �� ��������� ���������, ����������� ������ ��
'/Disinfect - ��������� ��� ��������� ���� �����
'/FreezeProcess - ���������� ��� ��������� �������� ����� ����������� ������
'/LockPoints - ������������� ����� ����������� (��������� ����� WMI, tasks, ���������� ���� �� ������ ������ � ������ � ����� �����������)
'/rawIgnoreList - ������������ ������ ������������� (����� ������) � ���� �������� ������ ������ .\whitelists.txt
'/Area:None - ��������� ���������� ��������� ������������ HiJackThis.
'/noShortcuts - ��������� �������� ������� ��� ��������� ����� /install
'/! - ������������� ������� ������. ���, ��� ��������� �����: ������������ � �������� ������ ��� ��������� HJT � ���������� (����������� �������) ������ ����������� /startupscan.


Option Explicit

Public Enum ENUM_REG_HIVE_FIX
    HKCR_FIX = 1
    HKCU_FIX = 2
    HKLM_FIX = 4
    HKU_FIX = 8
End Enum
#If False Then
    Dim HKCR_FIX, HKCU_FIX, HKLM_FIX, HKU_FIX
#End If

Public Enum ENUM_REG_REDIRECTION
    REG_REDIRECTED = -1
    REG_NOTREDIRECTED = 0
    REG_REDIRECTION_BOTH = 1
    [_REG_REDIRECTION_NOT_DEFINED] = -2
End Enum
#If False Then
    Dim REG_REDIRECTED, REG_NOTREDIRECTED, REG_REDIRECTION_BOTH, _REG_REDIRECTION_NOT_DEFINED
#End If

Public Enum ENUM_REG_VALUE_TYPE_RESTORE
    REG_RESTORE_SAME = -1&
    REG_RESTORE_SZ = 1&
    REG_RESTORE_EXPAND_SZ = 2&
    REG_RESTORE_BINARY = 3&
    REG_RESTORE_DWORD = 4&
    'REG_RESTORE_LINK = 6&
    REG_RESTORE_MULTI_SZ = 7&
    REG_RESTORE_QWORD = 8&
End Enum
#If False Then
    Dim REG_RESTORE_SAME, REG_RESTORE_SZ, REG_RESTORE_EXPAND_SZ, REG_RESTORE_DWORD, REG_RESTORE_MULTI_SZ, REG_RESTORE_QWORD
#End If

Public Enum ENUM_CURE_BASED
    FILE_BASED = 1          ' if need to cure .File()
    REGISTRY_BASED = 2      ' if need to cure .Reg()
    INI_BASED = 4           ' if need to cure ini-file in .reg()
    PROCESS_BASED = 8       ' if need to kill/freeze a process
    SERVICE_BASED = 16      ' if need to delete/restore service .ServiceName
    CUSTOM_BASED = 32       ' individual rule, based on .Custom() settings
    COMMANDLINE_BASED = 64  ' if need to run CMD command in .CommandLine()
    TASK_BASED = 128        ' if need to change .Task() state
End Enum
#If False Then
    Dim FILE_BASED, REGISTRY_BASED, INI_BASED, PROCESS_BASED, SERVICE_BASED, CUSTOM_BASED
#End If

Public Enum ENUM_COMMON_ACTION_BASED
    USE_FEATURE_DISABLE = &H10000
End Enum

Public Enum ENUM_REG_ACTION_BASED
    REMOVE_KEY = 1&
    REMOVE_VALUE = 2&
    RESTORE_VALUE = 4&      ' just overwrites the value with provided one
    RESTORE_VALUE_INI = 8&
    REMOVE_VALUE_INI = &H10&
    REPLACE_VALUE = &H20&   ' doing a replace of existing value if it is match the template defined
    APPEND_VALUE_NO_DOUBLE = &H40&
    REMOVE_VALUE_IF_EMPTY = &H80&
    REMOVE_KEY_IF_NO_VALUES = &H100&
    TRIM_VALUE = &H200&
    BACKUP_KEY = &H400&
    BACKUP_VALUE = &H800&
    JUMP_KEY = &H1000&
    JUMP_VALUE = &H2000&
    RESTORE_KEY_PERMISSIONS = &H4000&
    RESTORE_KEY_PERMISSIONS_RECURSE = &H8000&
    USE_FEATURE_DISABLE_REG = &H10000
    CREATE_KEY = &H20000
End Enum
#If False Then
    Dim REMOVE_KEY, REMOVE_VALUE, RESTORE_VALUE, RESTORE_VALUE_INI, REMOVE_VALUE_INI, REPLACE_VALUE
    Dim APPEND_VALUE_NO_DOUBLE, REMOVE_VALUE_IF_EMPTY, REMOVE_KEY_IF_NO_VALUES, TRIM_VALUE, BACKUP_KEY, BACKUP_VALUE
    Dim JUMP_KEY, JUMP_VALUE, RESTORE_KEY_PERMISSIONS, RESTORE_KEY_PERMISSIONS_RECURSE, USE_FEATURE_DISABLE_REG, CREATE_KEY
#End If

Public Enum ENUM_FILE_ACTION_BASED
    REMOVE_FILE = 1
    REMOVE_FOLDER = 2
    RESTORE_FILE = 4   'not used yet
    RESTORE_FILE_SFC = 8
    UNREG_DLL = &H10&
    BACKUP_FILE = &H20&
    JUMP_FILE = &H40&
    JUMP_FOLDER = &H80&
    CREATE_FOLDER = &H100&
    USE_FEATURE_DISABLE_FILE = &H10000
End Enum
#If False Then
    Dim REMOVE_FILE, REMOVE_FOLDER, RESTORE_FILE, RESTORE_FILE_SFC, UNREG_DLL, BACKUP_FILE, JUMP_FILE, JUMP_FOLDER, CREATE_FOLDER
    Dim USE_FEATURE_DISABLE_FILE
#End If

Public Enum ENUM_PROCESS_ACTION_BASED
    KILL_PROCESS = 1
    FREEZE_PROCESS = 2
    FREEZE_OR_KILL_PROCESS = 4
    CLOSE_PROCESS = 8
    CLOSE_OR_KILL_PROCESS = 16
    USE_FEATURE_DISABLE_PROCESS = &H10000
End Enum
#If False Then
    Dim KILL_PROCESS, FREEZE_PROCESS, FREEZE_OR_KILL_PROCESS
    Dim USE_FEATURE_DISABLE_PROCESS
#End If

Public Enum ENUM_SERVICE_ACTION_BASED
    DELETE_SERVICE = 1
    RESTORE_SERVICE = 2 ' not yet implemented
    DISABLE_SERVICE = 4
    ENABLE_SERVICE = 8
    MANUAL_SERVICE = 16
    STOP_SERVICE = 32
    START_SERVICE = 64
    USE_FEATURE_DISABLE_SERVICE = &H10000
End Enum
#If False Then
    Dim DELETE_SERVICE, RESTORE_SERVICE, DISABLE_SERVICE, ENABLE_SERVICE, MANUAL_SERVICE, STOP_SERVICE, START_SERVICE
    Dim USE_FEATURE_DISABLE_SERVICE
#End If

Public Enum ENUM_TASK_ACTION_BASED
    ENABLE_TASK = 1
    DISABLE_TASK = 2
End Enum
#If False Then
    Dim ENABLE_TASK, DISABLE_TASK
#End If

Public Enum ENUM_CUSTOM_ACTION_BASED
    CUSTOM_ACTION_O25 = 2 ^ 0
    CUSTOM_ACTION_SPECIFIC = 2 ^ 1
    CUSTOM_ACTION_BITS = 2 ^ 2
    CUSTOM_ACTION_APPLOCKER = 2 ^ 3
    CUSTOM_ACTION_REMOVE_GROUP_MEMBERSHIP = 2 ^ 4
    CUSTOM_ACTION_FIREWALL_RULE = 2 ^ 5
End Enum
#If False Then
    Dim CUSTOM_ACTION_O25, CUSTOM_ACTION_SPECIFIC, CUSTOM_ACTION_BITS, CUSTOM_ACTION_APPLOCKER, CUSTOM_ACTION_REMOVE_GROUP_MEMBERSHIP
    Dim CUSTOM_ACTION_FIREWALL_RULE
#End If

Public Enum ENUM_COMMANDLINE_ACTION_BASED
    COMMANDLINE_RUN = 1
    COMMANDLINE_POWERSHELL = 2
End Enum
#If False Then
    Dim COMMANDLINE_RUN, COMMANDLINE_POWERSHELL
#End If

Public Type FIX_REG_KEY
    IniFile         As String
    Hive            As ENUM_REG_HIVE
    Key             As String
    Param           As String
    ParamType       As ENUM_REG_VALUE_TYPE_RESTORE
    DefaultData     As Variant
    Redirected      As Boolean  'is key under Wow64
    ActionType      As ENUM_REG_ACTION_BASED
    ReplaceDataWhat As String
    ReplaceDataInto As String
    TrimDelimiter   As String
    DateM           As Date
    SD              As String
End Type

Public Type FIX_FILE
    path            As String
    Arguments       As String
    GoodFile        As String
    ActionType      As ENUM_FILE_ACTION_BASED
End Type

Private Type FIX_PROCESS
    PathOrName      As String
    pid             As Long
    ActionType      As ENUM_PROCESS_ACTION_BASED
End Type

Private Type FIX_SERVICE
    ImagePath       As String
    DllPath         As String
    serviceName     As String
    ServiceDisplay  As String
    ForceMicrosoft  As Boolean
    RunState        As SERVICE_STATE
    ActionType      As ENUM_SERVICE_ACTION_BASED
End Type

Private Type FIX_TASK
    TaskPath        As String
    ActionType      As ENUM_TASK_ACTION_BASED
End Type

Public Type FIX_CUSTOM
    ActionType      As ENUM_CUSTOM_ACTION_BASED
    Name            As String
    id              As String
    URL             As String
    TargetOrUser    As String
    CommandLine     As String
End Type

Public Type FIX_COMMANDLINE
    ActionType      As ENUM_COMMANDLINE_ACTION_BASED
    Executable      As String
    Arguments       As String
    Style           As SHOWWINDOW_FLAGS
    Wait            As Boolean
    TimeoutMs       As Long
End Type

Public Enum JUMP_ENTRY_TYPE
    JUMP_ENTRY_FILE = 1
    JUMP_ENTRY_REGISTRY = 2
End Enum

Private Type JUMP_ENTRY
    File()          As FIX_FILE
    Registry()      As FIX_REG_KEY
    Type            As JUMP_ENTRY_TYPE
End Type

Private Type O25_ActiveScriptConsumer_Entry
    File      As String
    Text      As String
    Engine    As String
End Type

Private Type O25_CommandLineConsumer_Entry
    ExecPath        As String
    WorkDir         As String
    CommandLine     As String
    Interactive     As Boolean
End Type

Public Enum O25_TIMER_TYPE
    O25_TIMER_ABSOLUTE = 1
    O25_TIMER_INTERVAL
End Enum

Private Type O25_Timer_Entry
    Type            As O25_TIMER_TYPE
    className       As String
    id              As String
    Interval        As Long 'for O25_TIMER_INTERVAL
    EventDateTime   As Date 'for O25_TIMER_ABSOLUTE
End Type

Public Enum O25_CONSUMER_TYPE
    O25_CONSUMER_ACTIVE_SCRIPT = 1
    O25_CONSUMER_COMMAND_LINE
End Enum

Private Type O25_Consumer_Entry
    Name        As String
    NameSpace   As String
    path        As String
    Type        As O25_CONSUMER_TYPE
    Script      As O25_ActiveScriptConsumer_Entry
    Cmd         As O25_CommandLineConsumer_Entry
    KillTimeout As Long
End Type

Private Type O25_Filter_Entry
    Name      As String
    NameSpace As String
    path      As String
    query     As String
End Type

Public Type O25_ENTRY
    Filter      As O25_Filter_Entry
    Timer       As O25_Timer_Entry
    Consumer    As O25_Consumer_Entry
End Type

Public Enum FIX_ITEM_STATE
    ITEM_STATE_DISABLED
    ITEM_STATE_ENABLED
End Enum

Public Type SCAN_RESULT
    HitLineA        As String
    HitLineW        As String
    Section         As String
    Alias           As String
    Name            As String
    State           As FIX_ITEM_STATE
    Reg()           As FIX_REG_KEY
    File()          As FIX_FILE
    Process()       As FIX_PROCESS
    Service()       As FIX_SERVICE
    Task()          As FIX_TASK
    Custom()        As FIX_CUSTOM
    CommandLine()   As FIX_COMMANDLINE
    Jump()          As JUMP_ENTRY
    CureType        As ENUM_CURE_BASED
    O25             As O25_ENTRY
    NoNeedBackup    As Boolean          'if no backup required / or impossible
    Reboot          As Boolean
    ForceMicrosoft  As Boolean
    FixAll          As Boolean
    SignResult      As SignResult_TYPE
End Type

Type TYPE_PERFORMANCE
    StartTime       As Long ' time the program started its working
    EndTime         As Long ' time the program finished its working
    MAX_TimeOut     As Long ' maximum time (mm) allowed for program to run the scanning
End Type

Private Type TASK_WHITELIST_ENTRY
    OSver       As Single
    path        As String
    RunObj      As String
    Args        As String
End Type

Private Type DICTIONARIES
    TaskWL_ID           As clsTrickHashTable
    dSafeProtocols      As clsTrickHashTable
    dSafeFilters        As clsTrickHashTable
    dLoLBin             As clsTrickHashTable
    dLoLBin_Protected   As clsTrickHashTable
    dSafeSvcPath        As clsTrickHashTable
    dSafeSvcFilename    As clsTrickHashTable
    DriverMapped        As clsTrickHashTable
End Type

Private Type IPSEC_FILTER_RECORD    '36 bytes
    Mirrored       As Byte
    Unknown1(2)    As Byte
    IP1(3)         As Byte
    IPTypeFlag1    As Long
    IP2(3)         As Byte
    IPTypeFlag2    As Long
    Unknown2       As Long
    ProtocolType   As Byte
    Unknown3(2)    As Byte
    PortNum1       As Integer
    PortNum2       As Integer
    Unknown4       As Byte
    DynPacketType  As Byte
    Unknown5(1)    As Byte
End Type

Private Type MY_PROC_LOG
    ProcName    As String
    Number      As Long
    IsMicrosoft As Boolean
    EDS_issued  As String
End Type

Private Type CERTIFICATE_BLOB_PROPERTY
    PropertyId As Long
    Reserved As Long
    Length As Long
    Data() As Byte
End Type

Private Type FONT_PROPERTY
    Bold        As Boolean
    Italic      As Boolean
    Underline   As Boolean
    Size        As Long
End Type

Private Enum APPLOCKER_RULE_TYPE
    APPLOCKER_RULE_UNKNOWN
    APPLOCKER_RULE_FILE_PATH
    APPLOCKER_RULE_FILE_HASH
    APPLOCKER_RULE_FILE_PUBLISHER
End Enum

Private Type APPLOCKER_HASH_RULE_DATA
    FileName As String
    FileLength As String
    Hash As String
End Type

Private Declare Sub OutputDebugStringA Lib "kernel32.dll" (ByVal lpOutputString As String)
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

Private HitSorted()     As String

Public gProcess()           As MY_PROC_ENTRY
Public g_TasksWL()          As TASK_WHITELIST_ENTRY
Public oDict                As DICTIONARIES

Public oDictFileExist       As clsTrickHashTable
Private dFontDefault        As clsTrickHashTable
Private aFontDefProp()      As FONT_PROPERTY

Public Scan()   As SCAN_RESULT    '// Dragokas. Used instead of parsing lines from result screen directly (like it was in original HJT 2.0.5).
                                  '// User type structures of arrays is filled together - using method frmMain.lstResults.AddItem
                                  '// It is much efficiently and have Unicode support (native vb6 ListBox is ANSI only,
                                  '// until we finally replaced it with Krool's CommonControls).

Public Perf     As TYPE_PERFORMANCE

Public OSver    As clsOSInfo
Public Proc     As clsProcess
Public cMath    As clsMath
Public cDrives  As clsDrives

Private oDictProcAvail As clsTrickHashTable

Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal pszStrPtr As Long, ByVal Length As Long) As String

'maps scan result string from ListBox to SCAN_RESULT structure associated with it
'previously it is used to find appropriate mapping beetween Ansi -> Unicode
'atm, leave it as is just in case some possible distortions in listbox
Public Function GetScanResults(HitLineA As String, result As SCAN_RESULT, Optional out_idx As Long) As Boolean
    Dim i As Long
    For i = 1 To UBound(Scan)
        If HitLineA = Scan(i).HitLineA Then
            result = Scan(i)
            out_idx = i
            GetScanResults = True
            Exit Function
        End If
    Next
    'Cannot find appropriate cure item for:, "Error"
    MsgBoxW Translate(592) & vbCrLf & HitLineA, vbCritical, Translate(591)
End Function

'.HitLineA => .HitLineW
Public Function MapHitlineAnsiToUnicode(HitLineA As String, Optional out_idx As Long) As String
    Dim i As Long
    For i = 1 To UBound(Scan)
        If HitLineA = Scan(i).HitLineA Then
            out_idx = i
            MapHitlineAnsiToUnicode = Scan(i).HitLineW
            Exit Function
        End If
    Next
End Function

Public Function RemoveFromScanResults(HitLineA As String) As Boolean
    Dim i As Long, j As Long
    Dim result As SCAN_RESULT
    For i = 1 To UBound(Scan)
        If HitLineA = Scan(i).HitLineA Then
            For j = i + 1 To UBound(Scan)
                Scan(j - 1) = Scan(j)
            Next
            If UBound(Scan) > 0 Then
                ReDim Preserve Scan(UBound(Scan) - 1)
            Else
                Scan(0) = result
            End If
            Exit For
        End If
    Next
End Function

Private Function FindHitLineIndex(sHitLineW As String) As Long

    Dim i As Long
    For i = 1 To UBound(Scan)
        If StrComp(Scan(i).HitLineW, sHitLineW, vbTextCompare) = 0 Then
            FindHitLineIndex = i
            Exit For
        End If
    Next

End Function

' it add Unicode SCAN_RESULT structure to shared array
Public Sub AddToScanResults( _
    result As SCAN_RESULT, _
    Optional ByVal DoNotAddToListBox As Boolean, _
    Optional DontClearResults As Boolean, _
    Optional DoNotDuplicate As Boolean)
    
    Dim bFirstWarning As Boolean
    Dim bAddedToList As Boolean
    Dim iStage As Long
    
    On Error GoTo ErrorHandler:
    
    Const SelLastAdded As Boolean = False
    
    'result.HitLineW = ScreenHitLine(result.HitLineW)
    'moved to => IsOnIgnoreList
    
    If DoNotDuplicate Then
        iStage = 1
        If UBound(Scan) > 0 Then
            iStage = 2
            Dim idx As Long
            idx = FindHitLineIndex(result.HitLineW)
            If idx <> 0 Then
                iStage = 3
                ConcatScanResults Scan(idx), result
                GoTo Finalize
            End If
        End If
    End If
    
    If bAutoLogSilent And Not g_bFixArg Then
        DoNotAddToListBox = True
    Else
        DoEvents
    End If
    If Not DoNotAddToListBox Then
        'checking if one of sections planned to be contains more then 50 entries -> block such attempt
        iStage = 4
        If Not SectionOutOfLimit(result.Section, bFirstWarning) Then
            bAddedToList = True
            'Commented (no difference)
            'LockWindowUpdate frmMain.lstResults.hwnd
            iStage = 5
            frmMain.lstResults.AddItem LimitHitLineLength(result.HitLineW, LIMIT_CHARS_COUNT_FOR_LISTLINE)
            'LockWindowUpdate 0&
            'Unicode to ANSI mapping (dirty hack)
            iStage = 6
            result.HitLineA = frmMain.lstResults.List(frmMain.lstResults.ListCount - 1)
            
            'select the last added line
            If SelLastAdded Then
                iStage = 7
                frmMain.lstResults.ListIndex = frmMain.lstResults.ListCount - 1
            End If
        Else
            If bFirstWarning Then
                'LockWindowUpdate frmMain.lstResults.hwnd
                iStage = 8
                frmMain.lstResults.AddItem result.Section & " - Too many entries ( > 250 )" '=> look Const LIMIT
                'LockWindowUpdate 0&
                iStage = 9
                AppendErrorLogCustom result.Section & " - Too many entries ( > 250 )"
                If SelLastAdded Then
                    iStage = 10
                    frmMain.lstResults.ListIndex = frmMain.lstResults.ListCount - 1
                End If
            End If
        End If
    End If
    iStage = 11
    ReDim Preserve Scan(UBound(Scan) + 1)
    iStage = 12
    Scan(UBound(Scan)) = result
    
    If (bDebugMode Or bDebugToFile) Then
        iStage = 13
        AppendErrorLogCustom "NEW DETECTION: " & result.HitLineW
        
        If bAddedToList Then
            AppendErrorLogCustom "[lstResults] Added to main frame."
        Else
            AppendErrorLogCustom "[lstResults] [NOT] Added to main frame."
        End If
    End If
    
Finalize:
    
    'Erase Result struct
    If Not DontClearResults Then
        iStage = 14
        EraseScanResults result
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddToScanResults", result.Section, result.HitLineW, "Stage: " & iStage
    If inIDE Then Stop: Resume Next
End Sub

Public Sub EraseScanResults(result As SCAN_RESULT)
    Dim EmptyResult As SCAN_RESULT
    result = EmptyResult
End Sub

'// Increase number of sections +1 and returns TRUE, if total number > LIMIT
Private Function SectionOutOfLimit(p_Section As String, Optional bFirstWarning As Boolean, Optional bErase As Boolean) As Long
    Dim LIMIT As Long: LIMIT = 250&
    If p_Section = "O23" Then LIMIT = 750
    
    Static Section As String
    Static Num As Long
    
    If bErase = True Then
        Section = vbNullString
        Num = 0
        Exit Function
    End If
    
    If p_Section = Section Then
        Num = Num + 1
        If Num > LIMIT Then
            If Num = LIMIT + 1 Then
                bFirstWarning = True
            End If
            SectionOutOfLimit = True
        End If
    Else
        Section = p_Section
        Num = 1
    End If
End Function

Public Sub AddToScanResultsSimple(Section As String, HitLine As String, Optional DoNotAddToListBox As Boolean)
    Dim result As SCAN_RESULT
    With result
        .Section = Section
        .HitLineW = HitLine
    End With
    AddToScanResults result, DoNotAddToListBox
End Sub

Private Sub ConcatScanFile(Dst() As FIX_FILE, Src() As FIX_FILE)
    Dim i As Long
    If AryPtr(Src) Then
        If AryPtr(Dst) Then
            For i = 0 To UBound(Src)
                If Not InArrayResultFile(Dst, Src(i)) Then
                    ReDim Preserve Dst(UBound(Dst) + 1)
                    Dst(UBound(Dst)) = Src(i)
                End If
            Next
        Else
            Dst = Src
        End If
    End If
End Sub

Private Sub ConcatScanRegistry(Dst() As FIX_REG_KEY, Src() As FIX_REG_KEY)
    Dim i As Long
    If AryPtr(Src) Then
        If AryPtr(Dst) Then
            For i = 0 To UBound(Src)
                If Not InArrayResultRegistry(Dst, Src(i)) Then
                    ReDim Preserve Dst(UBound(Dst) + 1)
                    Dst(UBound(Dst)) = Src(i)
                End If
            Next
        Else
            Dst = Src
        End If
    End If
End Sub

Private Sub ConcatScanProcess(Dst() As FIX_PROCESS, Src() As FIX_PROCESS)
    Dim i As Long
    If AryPtr(Src) Then
        If AryPtr(Dst) Then
            For i = 0 To UBound(Src)
                If Not InArrayResultProcess(Dst, Src(i)) Then
                    ReDim Preserve Dst(UBound(Dst) + 1)
                    Dst(UBound(Dst)) = Src(i)
                End If
            Next
        Else
            Dst = Src
        End If
    End If
End Sub

Private Sub ConcatScanService(Dst() As FIX_SERVICE, Src() As FIX_SERVICE)
    
    Dim i As Long
    If AryPtr(Src) Then
        If AryPtr(Dst) Then
            For i = 0 To UBound(Src)
                If Not InArrayResultService(Dst, Src(i)) Then
                    ReDim Preserve Dst(UBound(Dst) + 1)
                    Dst(UBound(Dst)) = Src(i)
                End If
            Next
        Else
            Dst = Src
        End If
    End If
End Sub

Private Sub ConcatScanTask(Dst() As FIX_TASK, Src() As FIX_TASK)
    
    Dim i As Long
    If AryPtr(Src) Then
        If AryPtr(Dst) Then
            For i = 0 To UBound(Src)
                If Not InArrayResultTask(Dst, Src(i)) Then
                    ReDim Preserve Dst(UBound(Dst) + 1)
                    Dst(UBound(Dst)) = Src(i)
                End If
            Next
        Else
            Dst = Src
        End If
    End If
End Sub

Private Sub ConcatScanCommandline(Dst() As FIX_COMMANDLINE, Src() As FIX_COMMANDLINE)
    
    Dim i As Long
    If AryPtr(Src) Then
        If AryPtr(Dst) Then
            For i = 0 To UBound(Src)
                If Not InArrayResultCommandline(Dst, Src(i)) Then
                    ReDim Preserve Dst(UBound(Dst) + 1)
                    Dst(UBound(Dst)) = Src(i)
                End If
            Next
        Else
            Dst = Src
        End If
    End If
End Sub

Private Sub ConcatScanCustom(Dst() As FIX_CUSTOM, Src() As FIX_CUSTOM)
    
    Dim i As Long
    If AryPtr(Src) Then
        If AryPtr(Dst) Then
            For i = 0 To UBound(Src)
                If Not InArrayResultCustom(Dst, Src(i)) Then
                    ReDim Preserve Dst(UBound(Dst) + 1)
                    Dst(UBound(Dst)) = Src(i)
                End If
            Next
        Else
            Dst = Src
        End If
    End If
End Sub

Public Sub ConcatJumpArray(Dst() As JUMP_ENTRY, Src() As JUMP_ENTRY)
    
    Dim i As Long
    
    If AryPtr(Src) Then
        If AryPtr(Dst) Then
            For i = 0 To UBound(Src)
                If Src(i).Type = JUMP_ENTRY_FILE Then
                
                    If Not InArrayJumpFile(Dst, Src(i).File(0)) Then 'jumps file array always has exactly 1 item
                        ReDim Preserve Dst(UBound(Dst) + 1)
                        Dst(UBound(Dst)) = Src(i)
                    End If
                
                ElseIf Src(i).Type = JUMP_ENTRY_REGISTRY Then
                
                    If Not InArrayJumpRegistry(Dst, Src(i).Registry(0)) Then 'jumps registry array always has exactly 1 item
                        ReDim Preserve Dst(UBound(Dst) + 1)
                        Dst(UBound(Dst)) = Src(i)
                    End If
                
                End If
            Next
        Else
            Dst = Src
        End If
    End If
    
End Sub

'
' State / NoNeedBackup / O25 / HitLineA are not included.
'
Public Sub ConcatScanResults(Dst As SCAN_RESULT, Src As SCAN_RESULT)
    
    If Src.CureType And FILE_BASED Then ConcatScanFile Dst.File, Src.File
    If Src.CureType And (REGISTRY_BASED Or INI_BASED) Then ConcatScanRegistry Dst.Reg, Src.Reg
    If Src.CureType And PROCESS_BASED Then ConcatScanProcess Dst.Process, Src.Process
    If Src.CureType And SERVICE_BASED Then ConcatScanService Dst.Service, Src.Service
    If Src.CureType And TASK_BASED Then ConcatScanTask Dst.Task, Src.Task
    If Src.CureType And CUSTOM_BASED Then ConcatScanCustom Dst.Custom, Src.Custom
    If Src.CureType And COMMANDLINE_BASED Then ConcatScanCommandline Dst.CommandLine, Src.CommandLine
    If AryPtr(Src.Jump) Then ConcatJumpArray Dst.Jump, Src.Jump
    
    With Dst
        If Len(.Alias) = 0 Then .Alias = Src.Alias
        If Len(.HitLineW) = 0 Then .HitLineW = Src.HitLineW
        If Len(.Name) = 0 Then .Name = Src.Name
        If Len(.Section) = 0 Then .Section = Src.Section
        
        .Reboot = .Reboot Or Src.Reboot
        .ForceMicrosoft = .ForceMicrosoft Or Src.ForceMicrosoft
        .FixAll = .FixAll Or Src.FixAll
        .CureType = .CureType Or Src.CureType
        
    End With
    
End Sub

Public Function InArrayJumpFile(Jump() As JUMP_ENTRY, Item As FIX_FILE) As Boolean
    Dim i As Long
    If AryPtr(Jump) Then
        For i = 0 To UBound(Jump)
            If InArrayResultFile(Jump(i).File, Item) Then
                InArrayJumpFile = True
                Exit For
            End If
        Next
    End If
End Function

Public Function InArrayJumpRegistry(Jump() As JUMP_ENTRY, Item As FIX_REG_KEY) As Boolean
    Dim i As Long
    If AryPtr(Jump) Then
        For i = 0 To UBound(Jump)
            If InArrayResultRegistry(Jump(i).Registry, Item) Then
                InArrayJumpRegistry = True
                Exit For
            End If
        Next
    End If
End Function

Public Function InArrayResultFile(FileArray() As FIX_FILE, Item As FIX_FILE) As Boolean
    Dim i As Long
    If AryPtr(FileArray) Then
        For i = 0 To UBound(FileArray)
            With FileArray(i)
                If .ActionType = Item.ActionType Then
                    If .path = Item.path Then
                        If .GoodFile = Item.GoodFile Then
                            InArrayResultFile = True
                            Exit For
                        End If
                    End If
                End If
            End With
        Next
    End If
End Function

Public Function InArrayResultRegistry(KeyArray() As FIX_REG_KEY, Item As FIX_REG_KEY) As Boolean
    Dim i As Long
    If AryPtr(KeyArray) Then
        For i = 0 To UBound(KeyArray)
            With KeyArray(i)
                If Item.Param = .Param Then
                    If Item.Key = .Key Then
                        If Item.Hive = .Hive And Item.ActionType = .ActionType Then
                            If Item.Redirected = .Redirected And Item.ParamType = .ParamType Then
                                If Item.DefaultData = .DefaultData Then
                                    If Item.IniFile = .IniFile Then
                                        If Item.ReplaceDataWhat = .ReplaceDataWhat Then
                                            If Item.ReplaceDataInto = .ReplaceDataInto Then
                                                If Item.TrimDelimiter = .TrimDelimiter Then
                                                    InArrayResultRegistry = True
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next
    End If
End Function

Public Function InArrayResultProcess(ProcessArray() As FIX_PROCESS, Item As FIX_PROCESS) As Boolean
    Dim i As Long
    If AryPtr(ProcessArray) Then
        For i = 0 To UBound(ProcessArray)
            With ProcessArray(i)
                If Item.ActionType = .ActionType Then
                    If Item.PathOrName = .PathOrName Then
                        If Item.pid = .pid Then
                            InArrayResultProcess = True
                            Exit For
                        End If
                    End If
                End If
            End With
        Next
    End If
End Function

Public Function InArrayResultService(ServiceArray() As FIX_SERVICE, Item As FIX_SERVICE) As Boolean
    Dim i As Long
    If AryPtr(ServiceArray) Then
        For i = 0 To UBound(ServiceArray)
            With ServiceArray(i)
                If Item.ActionType = .ActionType Then
                    If Item.RunState = .RunState Then
                        If Item.ImagePath = .ImagePath Then
                            If Item.DllPath = .DllPath Then
                                If Item.ServiceDisplay = .ServiceDisplay Then
                                    If Item.serviceName = .serviceName Then
                                        If Item.ForceMicrosoft = .ForceMicrosoft Then
                                            InArrayResultService = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next
    End If
End Function

Public Function InArrayResultTask(TaskArray() As FIX_TASK, Item As FIX_TASK) As Boolean
    Dim i As Long
    If AryPtr(TaskArray) Then
        For i = 0 To UBound(TaskArray)
            With TaskArray(i)
                If Item.TaskPath = .TaskPath Then
                    
                    InArrayResultTask = True
                    Exit For
                End If
            End With
        Next
    End If
End Function

Public Function InArrayResultCommandline(CommandlineArray() As FIX_COMMANDLINE, Item As FIX_COMMANDLINE) As Boolean
    Dim i As Long
    If AryPtr(CommandlineArray) Then
        For i = 0 To UBound(CommandlineArray)
            With CommandlineArray(i)
                If Item.ActionType = .ActionType Then
                    If Item.Executable = .Executable Then
                        If Item.Arguments = .Arguments Then
                            InArrayResultCommandline = True
                            Exit For
                        End If
                    End If
                End If
            End With
        Next
    End If
End Function

Public Function InArrayResultCustom(CustomArray() As FIX_CUSTOM, Item As FIX_CUSTOM) As Boolean
    Dim i As Long
    If AryPtr(CustomArray) Then
        For i = 0 To UBound(CustomArray)
            With CustomArray(i)
                If Item.ActionType = .ActionType Then
                    If Item.Name = .Name Then
                        If Item.id = .id Then
                            If Item.TargetOrUser = .TargetOrUser Then
                                If Item.CommandLine = .CommandLine Then
                                    InArrayResultCustom = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next
    End If
End Function

Public Sub GetHosts()
    If bIsWinNT Then
        g_HostsFile = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DataBasePath") & "\hosts"
    Else
        g_HostsFile = sWinDir & "\hosts"
    End If
End Sub

Public Sub LoadDatabase()
    On Error GoTo ErrorHandler:
    
    Static bInit As Boolean
    If bInit Then Exit Sub
    bInit = True
    
    AppendErrorLogCustom "LoadDatabase - Begin"
    
    Dim i As Long

    '=== LOAD REGVALS ===
    'syntax:
    '  regkey,regvalue,resetdata,baddata
    '  |      |        |          |
    '  |      |        |          data that shouldn't be (never used)
    '  |      |        R0 - data to reset to
    '  |      R1 - value to check
    '  R2 - regkey to check
    '
    'when empty:
    'R0 - everything is considered bad (always used), change to resetdata
    'R1 - value being present is considered bad, delete value
    'R2 - key being present is considered bad, delete key (not used)
    
    sRegVals = LoadEncryptedResFileAsArray("database\R_Section.txt", 108)
    
    'Only short hive names permitted here !
    ArrayAddStr sRegVals, "HKLM\SOFTWARE\Clients\StartMenuInternet\IEXPLORE.EXE\shell\open\command,(Default)," & _
            IIf(bIsWin64, "%ProgramW6432%", "%ProgramFiles%") & "\Internet Explorer\iexplore.exe" & _
            IIf(bIsWin64, "|%ProgramFiles(x86)%\Internet Explorer\iexplore.exe", vbNullString) & _
            "|" & """" & "%ProgramFiles%\Internet Explorer\iexplore.exe" & """" & _
            IIf(OSver.MajorMinor <= 5, "|", vbNullString) & _
            "|iexplore.exe" & _
            ","
    
    For i = 0 To UBound(sRegVals)
        sRegVals(i) = EnvironW(sRegVals(i))
    Next i
    
    
    '=== LOAD FILEVALS ===
    'syntax:
    ' inifile,section,value,resetdata,baddata
    ' |       |       |     |         |
    ' |       |       |     |         5) data that shouldn't be (never used)
    ' |       |       |     4) data to reset to
    ' |       |       |        (delete all if empty)
    ' |       |       3) value to check
    ' |       2) section to check
    ' 1) file to check
    
    Dim colFileVals As Collection
    Set colFileVals = New Collection
    
    'F0, F2 - if value modified
    'F1, F3 - if param. created
    
    With colFileVals
        .Add "system.ini;boot;Shell;explorer.exe;"        'F0 (boot;Shell)
        .Add "win.ini;windows;load;;"                     'F1 (windows;load)
        .Add "win.ini;windows;run;;"                      'F1 (windows;run)
        '\Software\Microsoft\Windows NT\CurrentVersion\WinLogon   'F2 (boot;Shell)
        .Add "REG:system.ini;boot;Shell;explorer.exe|%WINDIR%\explorer.exe;"
        '\Software\Microsoft\Windows NT\CurrentVersion\WinLogon   'F2 (boot;UserInit)
        .Add "REG:system.ini;boot;UserInit;%WINDIR%\System32\userinit.exe|userinit.exe|userinit;"
        '\Software\Microsoft\Windows NT\CurrentVersion\Windows    'F3 (windows;load)
        .Add "REG:win.ini;windows;load;;"
        '\Software\Microsoft\Windows NT\CurrentVersion\Windows    'F3 (windows;run)
        .Add "REG:win.ini;windows;run;;"
    End With
    ReDim sFileVals(colFileVals.Count - 1)
    For i = 1 To colFileVals.Count
        sFileVals(i - 1) = EnvironW(colFileVals.Item(i))
    Next

    '//TODO:
    '
    'What are ShellInfrastructure, VMApplet under winlogon ?
    'there are also 2 dll-s that may be interesting under \Windows (NaturalInputHandler, IconServiceLib)

    
    'LOAD R4 SEARCH URLS
    
    'HKCU\Software\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Microsoft\Internet Explorer\SearchScopes
    
    Dim aHives() As String, sHive$, j&
    
    GetHives aHives, addService:=False
    
    With cReg4vals
        '.Add "HKCU,DisplayName,Bing"
        '.Add "HKCU,FaviconURLFallback,http://www.bing.com/favicon.ico"
        
        If OSver.MajorMinor >= 6.3 Then 'Win 8.1
            '.Add "HKCU,FaviconURL,"
            '.Add "HKCU,NTLogoPath,"
            '.Add "HKCU,NTLogoURL,"
            '.Add "HKCU,NTSuggestionsURL,"
            '.Add "HKCU,NTTopResultURL,"
            '.Add "HKCU,NTURL,"
            .Add "HKCU,SuggestionsURL,"
            .Add "HKCU,TopResultURL,"
        
            .Add "HKCU,SuggestionsURLFallback,http://api.bing.com/qsml.aspx?query={searchTerms}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IE11SS&market={language}"
            .Add "HKCU,TopResultURLFallback,http://www.bing.com/search?q={searchTerms}&src=IE-TopResult&FORM=IE11TR"
            .Add "HKCU,URL,http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IE11SR"
            
        Else
        
            '.Add "HKCU,FaviconURL,http://www.bing.com/favicon.ico"
            '.Add "HKCU,NTLogoPath," & AppDataLocalLow & "\Microsoft\Internet Explorer\Services\"
            '.Add "HKCU,NTLogoURL,http://go.microsoft.com/fwlink/?LinkID=403856&language={language}&scale={scalelevel}&contrast={contrast}"
            '.Add "HKCU,NTSuggestionsURL,http://api.bing.com/qsml.aspx?query={searchTerms}&market={language}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IENTSS"
            '.Add "HKCU,NTTopResultURL,http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTTR"
            '.Add "HKCU,NTURL,http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTSR"
            
            .Add "HKCU,SuggestionsURL,http://api.bing.com/qsml.aspx?query={searchTerms}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IESS02&market={language}"
            .Add "HKCU,TopResultURL,http://www.bing.com/search?q={searchTerms}&src=IE-TopResult&FORM=IETR02"
            .Add "HKCU,SuggestionsURLFallback,http://api.bing.com/qsml.aspx?query={searchTerms}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IESS02&market={language}"
            .Add "HKCU,TopResultURLFallback,http://www.bing.com/search?q={searchTerms}&src=IE-TopResult&FORM=IETR02"
            .Add "HKCU,URL,http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IESR02"
        End If
        
        '.Add "HKLM,,Bing"
        '.Add "HKLM,DisplayName,@ieframe.dll,-12512|Bing"
        .Add "HKLM,URL,http://www.bing.com/search?q={searchTerms}&FORM=IE8SRC"
    End With
    
    For i = 1 To cReg4vals.Count  ' append HKU hive
        sHive = Left$(cReg4vals.Item(i), 4)
        If sHive = "HKCU" Then
            For j = 0 To UBound(aHives)
                If Left$(aHives(j), 3) = "HKU" Then
                    cReg4vals.Add Replace$(cReg4vals.Item(i), "HKCU", aHives(j), 1, 1)
                End If
            Next
        End If
    Next
    
    
    ' === LOAD SAFE O5 CONTROL PANEL DISABLED ITEMS
    
    If OSver.IsWindows10OrGreater Then
        '"appwiz.cpl|bthprops.cpl|desk.cpl|Firewall.cpl|hdwwiz.cpl|inetcpl.cpl|intl.cpl|irprops.cpl|joy.cpl|main.cpl|mmsys.cpl|ncpa.cpl|powercfg.cpl|sysdm.cpl|tabletpc.cpl|telephon.cpl|timedate.cpl"
        sSafeO5Items_HKLM = Caes_Decode("bsuDrK.rGE|ySISWVY^.Ra_|[^nh.dsq|OtEtNtGI.DSQ|QOdfZm.Zig|hohyjyw.rGE|FMUO.JYW|Xccgfin.bqo|qxJ.rGE|JzJQ.JYW|\^fnj.^mk|qhwj.pEC|KLVFUHMP.P_]|hpl_j.dsq|CloAvMKz.DSQ|]PYTa[de.^mk|wntnonIv.xMK")
        '"appwiz.cpl|bthprops.cpl|desk.cpl|Firewall.cpl|hdwwiz.cpl|inetcpl.cpl|intl.cpl|irprops.cpl|joy.cpl|main.cpl|mmsys.cpl|ncpa.cpl|powercfg.cpl|sysdm.cpl|telephon.cpl|timedate.cpl"
        sSafeO5Items_HKLM_32 = Caes_Decode("bsuDrK.rGE|ySISWVY^.Ra_|[^nh.dsq|OtEtNtGI.DSQ|QOdfZm.Zig|hohyjyw.rGE|FMUO.JYW|Xccgfin.bqo|qxJ.rGE|JzJQ.JYW|\^fnj.^mk|qhwj.pEC|KLVFUHMP.P_]|hpl_j.dsq|CpytGAJK.DSQ|]TZTUTi\.^mk")
    ElseIf OSver.IsWindows8OrGreater Then
        '"sysdm.cpl|inetcpl.cpl|ncpa.cpl|tabletpc.cpl|joy.cpl|powercfg.cpl|Firewall.cpl|telephon.cpl|irprops.cpl|intl.cpl|timedate.cpl|hdwwiz.cpl|mmsys.cpl|desk.cpl|main.cpl|appwiz.cpl|bthprops.cpl"
        sSafeO5Items_HKLM = Caes_Decode("tBxkv.pEC|DKDUFUS.N][|aXgZ.`om|yhkwrIGv.zOM|OVb.P_]|efp`obgj.jyw|UzKzTzMO.JYW|cV_Zgajk.dsq|rCCGFIN.BQO|PW_Y.Tca|mdjdedyl.nCA|AyTVJ].JYW|\^fnj.^mk|gjzt.pEC|HxHO.HWU|N_aj^q.^mk|eyoyCBEJ.xMK")
        '"sysdm.cpl|inetcpl.cpl|ncpa.cpl|Firewall.cpl|telephon.cpl|powercfg.cpl|irprops.cpl|joy.cpl|intl.cpl|timedate.cpl|hdwwiz.cpl|mmsys.cpl|main.cpl|desk.cpl|appwiz.cpl|bthprops.cpl"
        sSafeO5Items_HKLM_32 = Caes_Decode("tBxkv.pEC|DKDUFUS.N][|aXgZ.`om|KpApJpCE.zOM|YLUP]W`a.Zig|opzjylqt.tIG|HSSWVY^.Ra_|aht.bqo|pwEy.tIG|SJPJKJ_R.Tca|a_tvjC.jyw|BDLTP.DSQ|VLV].Vec|_brl.hwu|nEGPDW.DSQ|K_U_cbej.^mk")
    ElseIf (OSver.MajorMinor = 6.1 And OSver.IsServer) Then '2008 Server R2
        '"hdwwiz.cpl|telephon.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|inetcpl.cpl|joy.cpl|mmsys.cpl|Firewall.cpl|powercfg.cpl|intl.cpl|timedate.cpl|main.cpl|tabletpc.cpl|infocardcpl.cpl"
        sSafeO5Items_HKLM = Caes_Decode("igBDrK.rGE|QDMHUOXY.Ra_|XikthA.hwu|ArGt.zOM|X`\OZ.Tca|]`pj.fus|tAtKvKI.DSQ|SZf.Tca|fhpxt.hwu|SxIxRxKM.HWU|]^hXgZ_b.bqo|pwEy.tIG|SJPJKJ_R.Tca|f\fm.fus|EnqCxOMB.FUS|T[U`VVi]^mk.fus")
        '"hdwwiz.cpl|telephon.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|inetcpl.cpl|joy.cpl|mmsys.cpl|Firewall.cpl|powercfg.cpl|intl.cpl|timedate.cpl|main.cpl|tabletpc.cpl|infocardcpl.cpl"
        sSafeO5Items_HKLM_32 = Caes_Decode("igBDrK.rGE|QDMHUOXY.Ra_|XikthA.hwu|ArGt.zOM|X`\OZ.Tca|]`pj.fus|tAtKvKI.DSQ|SZf.Tca|fhpxt.hwu|SxIxRxKM.HWU|]^hXgZ_b.bqo|pwEy.tIG|SJPJKJ_R.Tca|f\fm.fus|EnqCxOMB.FUS|T[U`VVi]^mk.fus")
    ElseIf OSver.IsWindows7OrGreater Then
        '"hdwwiz.cpl|telephon.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|inetcpl.cpl|joy.cpl|mmsys.cpl|Firewall.cpl|powercfg.cpl|intl.cpl|timedate.cpl|main.cpl|collab.cpl|irprops.cpl|tabletpc.cpl|infocardcpl.cpl|bthprops.cpl"
        sSafeO5Items_HKLM = Caes_Decode("igBDrK.rGE|QDMHUOXY.Ra_|XikthA.hwu|ArGt.zOM|X`\OZ.Tca|]`pj.fus|tAtKvKI.DSQ|SZf.Tca|fhpxt.hwu|SxIxRxKM.HWU|]^hXgZ_b.bqo|pwEy.tIG|SJPJKJ_R.Tca|f\fm.fus|nBACtw.BQO|P[[_^af.Zig|sbeqlCAp.tIG|HOITJJ]QRa_.Zig|aukuyxAF.tIG")
        '"hdwwiz.cpl|telephon.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|inetcpl.cpl|joy.cpl|mmsys.cpl|Firewall.cpl|powercfg.cpl|intl.cpl|timedate.cpl|main.cpl|collab.cpl|irprops.cpl|infocardcpl.cpl|bthprops.cpl"
        sSafeO5Items_HKLM_32 = Caes_Decode("igBDrK.rGE|QDMHUOXY.Ra_|XikthA.hwu|ArGt.zOM|X`\OZ.Tca|]`pj.fus|tAtKvKI.DSQ|SZf.Tca|fhpxt.hwu|SxIxRxKM.HWU|]^hXgZ_b.bqo|pwEy.tIG|SJPJKJ_R.Tca|f\fm.fus|nBACtw.BQO|P[[_^af.Zig|hoitjjCqrGE.zOM|G[Q[_^af.Zig")
    ElseIf OSver.IsWindowsVistaOrGreater Then
        '"hdwwiz.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|Firewall.cpl|powercfg.cpl|infocardcpl.cpl|bthprops.cpl"
        sSafeO5Items_HKLM = Caes_Decode("igBDrK.rGE|xOQZNa.N][|aXgZ.`om|xFBoz.tIG|CFVP.L[Y|q\g\p\ik.fus|ABLvKxCF.FUS|T[U`VVi]^mk.fus|mGwGKJMR.FUS")
        '"ncpa.cpl|sysdm.cpl|desk.cpl|hdwwiz.cpl|Firewall.cpl|powercfg.cpl|appwiz.cpl|infocardcpl.cpl|bthprops.cpl"
        sSafeO5Items_HKLM_32 = Caes_Decode("ofuh.nCA|LTPCN.HWU|QTd^.Zig|gezBpI.pEC|[FQFZFSU.P_]|efp`obgj.jyw|pGIRFY.FUS|T[U`VVi]^mk.fus|mGwGKJMR.FUS")
    ElseIf OSver.IsWindowsXPOrGreater Then
        '"ncpa.cpl|odbccp32.cpl"
        sSafeO5Items_HKU = Caes_Decode("ofuh.nCA|HyyBDS45.N][")
        '"speech.cpl|infocardcpl.cpl"
        sSafeO5Items_HKLM = Caes_Decode("tsjlls.rGE|FMGRHH[OP_].Xge")
    Else
        '"ncpa.cpl|odbccp32.cpl"
        sSafeO5Items_HKU = Caes_Decode("ofuh.nCA|HyyBDS45.N][")
    End If
    
    
    ' === LOAD NONSTANDARD-BUT-SAFE-DOMAINS LIST === (R0, R1, O15)
    
    aSafeRegDomains = LoadEncryptedResFileAsArray("database\SafeDomains.txt", 109)
    
    
    ' === LOAD PROTOCOL SAFELIST === (O18)
    
    '//TODO: O18 - add file path checking to database
    Set oDict.dSafeProtocols = LoadEncryptedResFileAsDictionary("database\SafeProtocols.txt", 110, ",", False)
    
    
    ' === LOAD FILTER SAFELIST === (O18)
    
    Set oDict.dSafeFilters = LoadEncryptedResFileAsDictionary("database\SafeFilters.txt", 111, ",", False)


    'LOAD APPINIT_DLLS SAFELIST (O20)
    
    '*aakah.dll*akdllnt.dll*ROUSRNT.DLL*ssohook*KATRACK.DLL*APITRAP.DLL*UmxSbxExw.dll*sockspy.dll*scorillont.dll*wbsys.dll*NVDESK32.DLL*hplun.dll*mfaphook.dll*PAVWAIT.DLL*OCMAPIHK.DLL*MsgPlusLoader.dll*IconCodecService.dll*wl_hook.dll*Google\GOOGLE~1\GOEC62~1.DLL*adialhk.dll*wmfhotfix.dll*interceptor.dll*qaphooks.dll*RMProcessLink.dll*msgrmate.dll*wxvault.dll*ctu33.dll*ati2evxx.dll*vsmvhk.dll*
    sSafeAppInit = Caes_Decode("*dfrjs.sCE*xJEOQU].Q[]*GFNNOMU.ISU*FHFAJLJ*h`uufjt.qAC*vMHUUFW.OY[*hHUlC[d_`.Q[]*hf\fpoz.isu*FrFKDIKPQY.MWY*hUhpl.akm*S]MP`Z23.[eg*MWU`[.U_a*fa^oirtr.oyA*cVmp\fs.gqs*zpBrIDEJ.GQS*XFvaEPPePDIL[.Q[]*xZhiznehhZnCIxtx.AKM*\SHS\^\.Yce*DnpjqleR\^X_Z~0wdnff75~8.sCE*^cjdqot.qAC*RJEIRYMRc.S]_*`gobqdhuAxC.sCE*NzQKTVT^.S]_*IFKondhxzUtAz.wGI*NVLYVLaT.Wac*ruubxqA.oyA*vOR46.ISU*NcZ3Zmqs.cmo*CBxIwB.yIK*")
    
    
    'LOAD SSODL SAFELIST (O21)
    
    Dim colSafeSSODL As New Collection
    With colSafeSSODL
        .Add "{E6FB5E20-DE35-11CF-9C87-00AA005127ED}"  'WebCheck: C:\WINDOWS\System32\webcheck.dll (WinAll)
        .Add "{35CEC8A3-2BE6-11D2-8773-92E220524153}"  'SysTray: C:\WINDOWS\System32\stobject.dll (Win2k/XP)
        .Add "{7849596a-48ea-486e-8937-a2a3009f31a9}"  'PostBootReminder: C:\WINDOWS\system32\SHELL32.dll (WinXP)
        .Add "{fbeb8a05-beee-4442-804e-409d6c4515e9}"  'CDBurn: C:\WINDOWS\system32\SHELL32.dll (WinXP)
        .Add "{11566B38-955B-4549-930F-7B7482668782}"  'AUHook: C:\WINDOWS\SYSTEM\AUHOOK.DLL (WinME)
        .Add "{7007ACCF-3202-11D1-AAD2-00805FC1270E}"  'Network.ConnectionTray: C:\WINNT\system32\NETSHELL.dll (Win2k)
        .Add "{e57ce738-33e8-4c51-8354-bb4de9d215d1}"  'UPnPMonitor: C:\WINDOWS\SYSTEM\UPNPUI.DLL (WinME/XP)
        .Add "{BCBCD383-3E06-11D3-91A9-00C04F68105C}"  'AUHook: C:\WINDOWS\SYSTEM\AUHOOK.DLL (WinME)
        .Add "{F5DF91F9-15E9-416B-A7C3-7519B11ECBFC}"  '0aMCPClient: C:\Program Files\StarDock\MCPCore.dll
        .Add "{AAA288BA-9A4C-45B0-95D7-94D524869DB5}"  'WPDShServiceObj   WPDShServiceObj.dll Windows Portable Device Shell Service Object
        .Add "{1799460C-0BC8-4865-B9DF-4A36CD703FF0}" 'IconPackager Repair  iprepair.dll    Stardock\Object Desktop\ ThemeManager
        .Add "{6D972050-A934-44D7-AC67-7C9E0B264220}" 'EnhancedDialog   enhdlginit.dll  EnhancedDialog by Stardock
    End With
    'BE AWARE: SHELL32.dll - sometimes this file is patched (e.g. seen in Simplix)
    aSafeSSODL = ConvertCollectionToArray(colSafeSSODL)
    
    
    'LOAD SIOI SAFELIST (O21)
    
    Dim colSafeSIOI As Collection
    Set colSafeSIOI = New Collection
    
    With colSafeSIOI
'        .Add "{D9144DCD-E998-4ECA-AB6A-DCD83CCBA16D}"  'EnhancedStorageShell: C:\Windows\system32\EhStorShell.dll (Win7)
'        .Add "{4E77131D-3629-431c-9818-C5679DC83E81}"  'Offline Files: C:\Windows\System32\cscui.dll (Win7)
'        .Add "{08244EE6-92F0-47f2-9FC9-929BAA2E7235}"  'SharingPrivate: C:\Windows\system32\ntshrui.dll (Win7)
'        .Add "{750fdf0e-2a26-11d1-a3ea-080036587f03}"  'Offline Files: C:\WINDOWS\System32\cscui.dll (WinXP)
'        .Add "{fbeb8a05-beee-4442-804e-409d6c4515e9}"  'CDBurn: C:\WINDOWS\system32\SHELL32.dll (WinXP)
'        .Add "{7849596a-48ea-486e-8937-a2a3009f31a9}"  'PostBootReminder: C:\WINDOWS\system32\SHELL32.dll (WinXP)
'        .Add "{0CA2640D-5B9C-4c59-A5FB-2DA61A7437CF}" 'StorageProviderError: C:\Windows\System32\shell32.dll, C:\Windows\SysWOW64\shell32.dll (Win 8.1)
'        .Add "{0A30F902-8398-4ee8-86F7-4CFB589F04D1}" 'StorageProviderSyncing: C:\Windows\System32\shell32.dll, C:\Windows\SysWOW64\shell32.dll (Win 8.1)

        '<SysRoot>\system32\EhStorShell.dll
        .Add Caes_Decode("<VDz[zBI>oNVRUHR67GlWDgdiLcbkm.isu")
        '<SysRoot>\system32\cscui.dll
        .Add Caes_Decode("<VDz[zBI>oNVRUHR67GPbTh^.]gi")
        '<SysRoot>\system32\ntshrui.dll
        .Add Caes_Decode("<VDz[zBI>oNVRUHR67G[cd[glb.akm")
        '<SysRoot>\system32\SHELL32.dll
        .Add Caes_Decode("<VDz[zBI>oNVRUHR67GzqpyA23.akm")
        '<SysRoot>\SysWOW64\shell32.dll
        .Add Caes_Decode("<VDz[zBI>ohVRrlv99G`WV_a23.akm")
        
        '.Add "<SysRoot>\system32\mscoree.dll" 'adware
    End With
    ReDim aSafeSIOI(colSafeSIOI.Count - 1)
    For i = 1 To colSafeSIOI.Count
        aSafeSIOI(i - 1) = Replace(colSafeSIOI.Item(i), "<SysRoot>", sWinDir, 1, -1, vbTextCompare)
    Next
    
    
    'LOAD ShellExecuteHooks (SEH) SAFELIST (O21)
    
    Dim colSafeSEH As Collection
    Set colSafeSEH = New Collection
    
    With colSafeSEH
        '<SysRoot>\system32\shell32.dll
        .Add Caes_Decode("<VDz[zBI>oNVRUHR67G`WV_a23.akm")
    End With
    ReDim aSafeSEH(colSafeSEH.Count - 1)
    For i = 1 To colSafeSEH.Count
        aSafeSEH(i - 1) = Replace(colSafeSEH.Item(i), "<SysRoot>", sWinDir, 1, -1, vbTextCompare)
    Next
    
    
    'LOAD DEBUGGER SAFELIST (O26)
    
    '*vrfcore.dll*vfbasics.dll*vfcompat.dll*vfluapriv.dll*vfprint.dll*vfnet.dll*vfntlmless.dll*vfnws.dll*vfcuzz.dll*
    sSafeIfeVerifier = Caes_Decode("*ywmlzEt.wGI*WIGH\TPb.Wac*qcbppuhC.qAC*QCKVDUYRa.S]_*m_kohow.kuw*KwGzQ.EOQ*_Q[c]`a\ln.cmo*CoyJH.wGI*WIH\ce.S]_*")
    
    
    'LOAD SAFE DNS LIST (O17)
    
    'These are checked with nslookup
    'https://www.comss.ru/list.php?c=securedns
    
    Set colSafeDNS = LoadEncryptedResFileAsCollection("database\SafeDNS.txt", 112, ",")
    
    
    'LOAD DISALLOWED CERTIFICATES (O7)
    
    Set colDisallowedCert = LoadEncryptedResFileAsCollection("database\DisallowedCert.txt", 113, ";")
    
    
    'LOAD WINDOWS SERVICE (O23 - Service)
    'Note: this list is only used to improve scan speed
    
    Set oDict.dSafeSvcPath = LoadEncryptedResFileAsDictionary("database\ServicePath.txt", 115, ",", True) 'Arguments delimiter is |
    Set oDict.dSafeSvcFilename = LoadEncryptedResFileAsDictionary("database\ServiceFilename.txt", 116, vbNullString, False)
    
    
    'LOAD MAPPED DRIVERS (O23 - Drivers)
    
    Set oDict.DriverMapped = LoadEncryptedResFileAsDictionary("database\DriverMapped.txt", 117, vbNullString, True)
    
    
    AppendErrorLogCustom "LoadDatabase - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "LoadDatabase"
    If inIDE Then Stop: Resume Next
End Sub

Private Function IsScanRequired(idSection As ID_SECTION) As Boolean
    With g_ScanFilter
        If .DoInclude Then
            If .Inclusion(idSection) Then
                IsScanRequired = True
            Else
                Exit Function
            End If
        End If
        If .DoExclude Then
            If .Exclusion(idSection) Then
                Exit Function
            End If
        End If
        IsScanRequired = True
    End With
End Function

Public Sub StartScan()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "StartScan - Begin"
    
    If bDebugToFile Then
        If g_hDebugLog = 0 Then OpenDebugLogHandle
    End If
    
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    
    bScanMode = True
    oDictFileExist.RemoveAll
    
    SetPriorityAllThreads GetCurrentProcess(), THREAD_PRIORITY_HIGHEST
    
    frmMain.txtNothing.ZOrder 1
    frmMain.txtNothing.Visible = False
    
    'frmMain.shpBackground.Tag = iItems
    SetProgressBar g_HJT_Items_Count   'R + F + O26
    
    If Not bAutoLogSilent Then ' Do not touch!
        Call GetProcesses(gProcess)
    Else
        If AryPtr(gProcess) = 0 Then
            Call GetProcesses(gProcess)
        End If
    End If
    
    LoadDatabase
    
    Dim i&
    'load ignore list
    IsOnIgnoreList vbNullString
    
    frmMain.lstResults.Clear
    AppendErrorLogCustom "[lstResults] Cleared."
    
    'Disabled - too controversial decision: Win10+ has too many catalogues;
    '   also, cannot extract the real signer of 3-rd party app: certificate context is pointed to catalogue itself which is signed by Microsoft
    'Loading security catalogue to cache
    'UpdateProgressBar "SecCat"
    'Dim SignResult As SignResult_TYPE
    'SignVerify vbNullString, SV_EnableAllTagsPrecache, SignResult
    
    If IsScanRequired(ID_SECTION_O22) Then
        EnumBITS_Stage1 ' Speed hack. Run process in advance, get results at the very end.
    End If
    
    'Registry
    
    If IsScanRequired(ID_SECTION_R) Then
        UpdateProgressBar "R"
        For i = 0 To UBound(sRegVals)
            ProcessRuleReg sRegVals(i)
        Next i
        
        CheckR3Item
        CheckR4Item
    End If
    
    If IsScanRequired(ID_SECTION_F) Then
        UpdateProgressBar "F"
        'File
        For i = 0 To UBound(sFileVals)
            If Len(sFileVals(i)) <> 0 Then
                CheckFileItems sFileVals(i)
            End If
        Next i
    End If
    
    'Netscape/Mozilla stuff
    'CheckNetscapeMozilla        'N1-4
    
    If IsScanRequired(ID_SECTION_O7) Then
        Dim sWalletAddr As String
        Dim sClipPrevText As String
        sWalletAddr = GenWalletAddressETH()
        sClipPrevText = ClipboardGetText()
        ClipboardSetText sWalletAddr
    End If
    
    If IsScanRequired(ID_SECTION_B) Then
        UpdateProgressBar "B"
        CheckBrowsersItem
    End If
    
    'Other scans
    If IsScanRequired(ID_SECTION_O1) Then
        UpdateProgressBar "O1"
        CheckO1Item 'Hosts
        CheckO1Item_ICS
        CheckO1Item_DNSApi
    End If
    
    If IsScanRequired(ID_SECTION_O2) Then
        UpdateProgressBar "O2"
        CheckO2Item 'BHO
    End If
    
    If IsScanRequired(ID_SECTION_O3) Then
        UpdateProgressBar "O3"
        CheckO3Item 'toolbars
    End If
    
    If IsScanRequired(ID_SECTION_O4) Then
        UpdateProgressBar "O4"
        CheckO4Item 'Autorun
    End If
    
    If IsScanRequired(ID_SECTION_O5) Then
        UpdateProgressBar "O5"
        CheckO5Item 'Control panel
    End If
    
    If IsScanRequired(ID_SECTION_O6) Then
        UpdateProgressBar "O6"
        CheckO6Item 'IE Policy
    End If
    
    If IsScanRequired(ID_SECTION_O7) Then
        UpdateProgressBar "O7"
        CheckO7Item 'OS Policy
        CheckO7Item_Bitcoin sWalletAddr
        If Len(sClipPrevText) <> 0 Then ClipboardSetText sClipPrevText
    End If
        
    If IsScanRequired(ID_SECTION_O8) Then
        UpdateProgressBar "O8"
        CheckO8Item 'IE: Context menu
    End If
    
    If IsScanRequired(ID_SECTION_O9) Then
        UpdateProgressBar "O9"
        CheckO9Item 'IE: Services & Buttons
    End If
        
    If IsScanRequired(ID_SECTION_O10) Then
        UpdateProgressBar "O10"
        CheckO10Item 'LSP
    End If
    
    If IsScanRequired(ID_SECTION_O11) Then
        UpdateProgressBar "O11"
        CheckO11Item 'IE: 'Advanced' tab
    End If
    
    If IsScanRequired(ID_SECTION_O12) Then
        UpdateProgressBar "O12"
        CheckO12Item 'IE: plugins of file ext./MIME types
    End If
    
    If IsScanRequired(ID_SECTION_O13) Then
        UpdateProgressBar "O13"
        CheckO13Item 'URL Prefixes
    End If
    
    If IsScanRequired(ID_SECTION_O14) Then
        UpdateProgressBar "O14"
        CheckO14Item 'IE: IERESET.INF
    End If
    
    If IsScanRequired(ID_SECTION_O15) Then
        UpdateProgressBar "O15"
        CheckO15Item 'Trusted Zone
    End If
    
    If IsScanRequired(ID_SECTION_O16) Then
        UpdateProgressBar "O16"
        CheckO16Item 'Downloaded Program Files
    End If
    
    If IsScanRequired(ID_SECTION_O17) Then
        UpdateProgressBar "O17"
        CheckO17Item 'DNS/DHCP
    End If
    
    If IsScanRequired(ID_SECTION_O18) Then
        UpdateProgressBar "O18"
        CheckO18Item 'Protocols, filters
    End If
    
    If IsScanRequired(ID_SECTION_O19) Then
        UpdateProgressBar "O19"
        CheckO19Item 'User stylesheet
    End If
    
    If IsScanRequired(ID_SECTION_O20) Then
        UpdateProgressBar "O20"
        CheckO20Item 'AppInit_DLLs, Winlogon Notify
    End If
    
    If IsScanRequired(ID_SECTION_O21) Then
        UpdateProgressBar "O21"
        CheckO21Item 'Shell Service Object Delay Load (SSODL), Shell Icon Overlay (SIOI), ShellExecuteHooks (SEH)
    End If
    
    If IsScanRequired(ID_SECTION_O22) Then
        UpdateProgressBar "O22"
        CheckO22Item 'Tasks, BITS Admin
    End If
    
    If IsScanRequired(ID_SECTION_O23) Then
        UpdateProgressBar "O23"
        CheckO23Item 'Services & Drivers
    End If
    
    If IsScanRequired(ID_SECTION_O24) Then
        UpdateProgressBar "O24"
        CheckO24Item 'ActiveX Desktop
    End If
    
    If IsScanRequired(ID_SECTION_O25) Then
        UpdateProgressBar "O25"
        CheckO25Item 'WMI
    End If
    
    If IsScanRequired(ID_SECTION_O26) Then
        UpdateProgressBar "O26"
        CheckO26Item 'Debuggers, Tools hijack
    End If
    
    If IsScanRequired(ID_SECTION_O27) Then
        UpdateProgressBar "O27"
        CheckO27Item 'Account
    End If
    
    If IsScanRequired(ID_SECTION_O22) Then
        EnumBITS_Stage2
    End If
    
    UpdateProgressBar "ProcList"
    
    With frmMain
        If .lstResults.ListCount > 0 Or bAutoLogSilent Then
            .txtNothing.ZOrder 1
            .txtNothing.Visible = False
        Else
            .txtNothing.Visible = True
            .txtNothing.ZOrder 0
        End If
    End With
    
    bScanMode = False
    'SignVerify "", SV_CacheFree, SignResult
    
    SectionOutOfLimit vbNullString, bErase:=True
    
    Dim sEDS_Time   As String
    Dim OSData      As String
    
    If bDebugMode Or bDebugToFile Then
    
        If ObjPtr(OSver) <> 0 Then
                OSData = OSver.Bitness & " " & OSver.OSName & IIf(Len(OSver.Edition) <> 0, " (" & OSver.Edition & ")", vbNullString) & ", " & _
                    OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & ", " & _
                    "Service Pack: " & Replace(OSver.SPVer, ",", ".") & IIf(OSver.IsSafeBoot, " (Safe Boot)", vbNullString)
        End If
    
        sEDS_Time = vbCrLf & vbCrLf & "Logging is finished." & vbCrLf & vbCrLf & AppVerPlusName & vbCrLf & vbCrLf & OSData & vbCrLf & vbCrLf & _
                "Time spent: " & ((GetTickCount() - Perf.StartTime) \ 100) / 10 & " sec." & vbCrLf & vbCrLf & _
                "Whole EDS function: " & Format$(tim(0).GetTime, "##0.000 sec.") & vbCrLf & _
                "CryptCATAdminAcquireContext: " & Format$(tim(1).GetTime, "##0.000 sec.") & vbCrLf & _
                "CryptCATAdminCalcHashFromFileHandle: " & Format$(tim(2).GetTime, "##0.000 sec.") & vbCrLf & _
                "CryptCATAdminEnumCatalogFromHash: " & Format$(tim(3).GetTime, "##0.000 sec.") & vbCrLf & _
                "WinVerifyTrust: " & Format$(tim(4).GetTime, "##0.000 sec.") & vbCrLf & _
                "GetSignerInfo: " & Format$(tim(5).GetTime, "##0.000 sec.") & vbCrLf & _
                "Release: " & Format$(tim(6).GetTime, "##0.000 sec.") & vbCrLf & _
                "CryptCATEnumerateMember: " & Format$(tim(7).GetTime, "##0.000 sec.") & vbCrLf & vbCrLf
        
        AppendErrorLogCustom sEDS_Time
    End If
    
    If bDebugToFile Then
        If g_hDebugLog <> 0 Then
            'Append Header to the end and close debug log file
            Dim b() As Byte
            b = sEDS_Time & vbCrLf & vbCrLf
            PutW_NoLog g_hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
        End If
    End If
    
    SetPriorityAllThreads GetCurrentProcess(), THREAD_PRIORITY_NORMAL
    
    AppendErrorLogCustom "StartScan - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_StartScan"
    bScanMode = False
    SetPriorityAllThreads GetCurrentProcess(), THREAD_PRIORITY_NORMAL
    If inIDE Then Stop: Resume Next
End Sub

Public Sub ResumeHashProgressbar()
    
    g_bCalcHashInProgress = True
    
    With frmMain.lblMD5
        .Visible = True
        .Font.Bold = False
        .Font.Underline = False
    End With
    
    frmMain.shpMD5Background.Visible = True
    frmMain.shpMD5Progress.Visible = True
    
    frmMain.lblInfo(0).Visible = False
    frmMain.lblInfo(1).Visible = False
    
End Sub

Public Sub SetHashProgressBar(lPercent As Long, Optional sText As String)
    On Error GoTo ErrorHandler:
    
    frmMain.shpMD5Progress.Width = frmMain.ScaleWidth * (lPercent / 100)
    frmMain.shpMD5Progress.Tag = lPercent
    
    If Len(sText) <> 0 Then
        frmMain.lblMD5.Caption = sText
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SetHashProgressBar", "Percent: " & lPercent, "Form ScaleWidth: " & frmMain.ScaleWidth
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CloseHashProgressbar()
    
    frmMain.lblMD5.Visible = False
    frmMain.shpMD5Background.Visible = False
    frmMain.shpMD5Progress.Visible = False
    
    g_bCalcHashInProgress = False
    
End Sub

Public Sub SetProgressBarOnFront()
    frmMain.lblStatus.Visible = True
    frmMain.lblStatus.ZOrder 0 'on top
    frmMain.lblInfo(0).Visible = False
    frmMain.lblInfo(1).Visible = False
    frmMain.shpBackground.Visible = True
    frmMain.shpProgress.Visible = True
    frmMain.shpProgress.ZOrder 1
    frmMain.shpBackground.ZOrder 1
End Sub

Public Sub SetProgressBar(lMaxTags As Long)
    
    If g_bCheckSum Then ResumeHashProgressbar
    
    'ProgressBar label settings
    frmMain.lblStatus.Visible = True
    frmMain.lblStatus.Caption = vbNullString
    frmMain.lblStatus.ForeColor = &HFFFF&   'Yellow
    frmMain.lblStatus.ZOrder 0 'on top
    frmMain.lblStatus.Left = 400
    
    'Logo -> off
    frmMain.pictLogo.Visible = False
    
    'results label -> off
    frmMain.lblInfo(0).Visible = False
    'program description label -> off
    frmMain.lblInfo(1).Visible = False
    
    frmMain.shpBackground.Visible = True
    With frmMain.shpProgress
        .Tag = "0"
        .Visible = True
    End With
    frmMain.shpProgress.Width = 255 ' default
    frmMain.shpProgress.ZOrder 1
    frmMain.shpBackground.ZOrder 1
    g_ProgressMaxTags = lMaxTags
End Sub

Public Sub CloseProgressbar(Optional bScanCompeleted As Boolean)
    frmMain.shpBackground.Visible = False
    frmMain.shpProgress.Visible = False
    frmMain.lblStatus.Visible = False
    
    If bScanCompeleted Then
        If frmMain.lstResults.Visible Then
            If g_bVTScanInProgress Or g_bVTScanned Then
                UpdateVTProgressbar bFinished:=g_bVTScanned
            Else
                CloseHashProgressbar
                frmMain.lblInfo(1).Visible = True
            End If
        End If
        If Not TaskBar Is Nothing Then TaskBar.SetProgressState g_HwndMain, TBPF_NOPROGRESS
    End If
End Sub

Public Sub ResumeProgressbar()
    frmMain.shpBackground.Visible = True
    frmMain.shpProgress.Visible = True
    frmMain.lblStatus.Visible = True
    If g_bCheckSum Then
        frmMain.shpMD5Background.Visible = True
        frmMain.shpMD5Progress.Visible = True
        frmMain.lblMD5.Visible = True
    End If
End Sub

Public Sub UpdateProgressBar(Section As String, Optional sAppendText As String)
    On Error GoTo ErrorHandler:
    
    If g_bNoGUI Then Exit Sub
    
    AppendErrorLogCustom "Progressbar - " & Section & " " & sAppendText
    
    Dim lTag As Long
    
    With frmMain
    
        If Not IsNumeric(.shpProgress.Tag) Then .shpProgress.Tag = "0"
        lTag = .shpProgress.Tag
        If Len(sAppendText) = 0 Then lTag = lTag + 1
        .shpProgress.Tag = lTag
        
        Select Case Section
            Case "SecCat": .lblStatus.Caption = Translate(1871)
        
            Case "R", "R0", "R1", "R2", "R3": .lblStatus.Caption = Translate(230) & "..."
            Case "F", "F1", "F2", "F3": .lblStatus.Caption = Translate(231) & "..."
            Case "B": .lblStatus.Caption = Translate(232) & "..."
            Case "O1": .lblStatus.Caption = Translate(233) & "..."
            Case "O2": .lblStatus.Caption = Translate(234) & "..."
            Case "O3": .lblStatus.Caption = Translate(235) & "..."
            Case "O4": .lblStatus.Caption = Translate(236) & "..."
            Case "O5": .lblStatus.Caption = Translate(237) & "..."
            Case "O6": .lblStatus.Caption = Translate(238) & "..."
            Case "O7": .lblStatus.Caption = Translate(239) & "..."
            Case "O7-Cert": .lblStatus.Caption = Translate(264) & "..."
            Case "O7-Trouble": .lblStatus.Caption = Translate(265) & "..."
            Case "O7-ACL": .lblStatus.Caption = Translate(267) & "..."
            Case "O7-IPSec": .lblStatus.Caption = Translate(266) & "..."
            Case "O8": .lblStatus.Caption = Translate(240) & "..."
            Case "O9": .lblStatus.Caption = Translate(241) & "..."
            Case "O10": .lblStatus.Caption = Translate(242) & "..."
            Case "O11": .lblStatus.Caption = Translate(243) & "..."
            Case "O12": .lblStatus.Caption = Translate(244) & "..."
            Case "O13": .lblStatus.Caption = Translate(245) & "..."
            Case "O14": .lblStatus.Caption = Translate(246) & "..."
            Case "O15": .lblStatus.Caption = Translate(247) & "..."
            Case "O16": .lblStatus.Caption = Translate(248) & "..."
            Case "O17": .lblStatus.Caption = Translate(249) & "..."
            Case "O18": .lblStatus.Caption = Translate(250) & "..."
            Case "O19": .lblStatus.Caption = Translate(251) & "..."
            Case "O20": .lblStatus.Caption = Translate(252) & "..."
            Case "O21": .lblStatus.Caption = Translate(253) & "..."
            Case "O22": .lblStatus.Caption = Translate(254) & "..."
            Case "O23": .lblStatus.Caption = Translate(255) & "..."
            Case "O23-D": .lblStatus.Caption = Translate(263) & "..."
            Case "O24": .lblStatus.Caption = Translate(257) & "..."
            Case "O25": .lblStatus.Caption = Translate(258) & "..."
            Case "O26": .lblStatus.Caption = Translate(261) & "..."
            Case "O27": .lblStatus.Caption = Translate(1750) & "..."
            
            Case "ProcList": .lblStatus.Caption = Translate(260) & "..."
            Case "ModuleList": .lblStatus.Caption = Translate(268) & "..."
            Case "Backup":   .lblStatus.Caption = Translate(259) & "...": .shpProgress.Width = 255
            Case "Report":   .lblStatus.Caption = Translate(262) & "..."
            Case "Finish":   .lblStatus.Caption = Translate(256): .shpProgress.Width = .shpBackground.Width + .shpBackground.Left - .shpProgress.Left
        End Select
        
        If Len(sAppendText) <> 0 Then .lblStatus.Caption = .lblStatus.Caption & " - " & sAppendText
        
        Select Case Section
            Case "ProcList": Exit Sub
            Case "Backup": Exit Sub
            Case "Finish": Exit Sub
        End Select
        
        If lTag > g_ProgressMaxTags Then lTag = g_ProgressMaxTags
        
        If g_ProgressMaxTags <> 0 Then
            .shpProgress.Width = .shpBackground.Width * (lTag / g_ProgressMaxTags)  'g_ProgressMaxTags = items to check or fix -1
            SetTaskBarProgressValue frmMain, (lTag / g_ProgressMaxTags)
        End If
        
        '.lblStatus.Refresh
        '.Refresh
    End With
    
    If Not bAutoLogSilent Then DoEvents
    
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "modMain_UpdateProgressBar", "shpProgress.Tag=", frmMain.shpProgress.Tag
    If inIDE Then Stop: Resume Next
End Sub


'CheckR0item
'CheckR1item
'CheckR2item
Private Sub ProcessRuleReg(ByVal sRule$)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ProcessRuleReg - Begin", "Rule: " & sRule
    
    Dim vRule As Variant, iMode&, bIsNSBSD As Boolean, result As SCAN_RESULT
    Dim sHit$, sKey$, sParam$, sData$, sDefDataStrings$, Wow6432Redir As Boolean, UseWow
    Dim bProxyEnabled As Boolean, hHive As ENUM_REG_HIVE
    
    'Registry rule syntax:
    '[regkey],[regvalue],[infected data],[default data]
    '* [regkey]           = "" -> abort - no way man!
    ' * [regvalue]        = "" -> delete entire key
    '  * [default data]   = "" -> delete value
    '   * [infected data] = "" -> any value (other than default) is considered infected
    vRule = Split(sRule, ",")
    
    ' iMode = 0 -> check if value is infected
    ' iMode = 1 -> check if value is present
    ' iMode = 2 -> check if regkey is present
    If CStr(vRule(0)) = vbNullString Then Exit Sub
    If CStr(vRule(3)) = vbNullString Then iMode = 0
    If CStr(vRule(2)) = vbNullString Then iMode = 1
    If CStr(vRule(1)) = vbNullString Then iMode = 2
    
    sKey = vRule(0)
    sParam = vRule(1)
    If sParam = "(Default)" Then sParam = vbNullString
    sDefDataStrings = vRule(2)
       
    'Initialize hives enumerator
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey sKey
    
    Do While HE.MoveNext
    
        Wow6432Redir = HE.Redirected
        sKey = HE.Key
        hHive = HE.Hive
    
        Select Case iMode
        
        Case 0 'check for incorrect value
            If Reg.ValueExists(hHive, sKey, sParam, Wow6432Redir) Then
            
                sData = Reg.GetString(hHive, sKey, sParam, Wow6432Redir)
              
                If Len(sData) <> 0 Then
              
                    sData = UnQuote(EnvironW(sData))
            
                    If Not inArraySerialized(sData, sDefDataStrings, "|", , , 1) Or (Not bHideMicrosoft) Then
                        bIsNSBSD = False
                        If bHideMicrosoft And Not bIgnoreAllWhitelists Then bIsNSBSD = StrBeginWithArray(sData, aSafeRegDomains)
                        If Not bIsNSBSD Then
                            If InStr(1, sData, "%2e", 1) > 0 Then sData = UnEscape(sData) & " " & STR_OBFUSCATED
                            
                            sHit = BitPrefix("R0", HE) & " - " & _
                                HE.KeyAndHivePhysical & ": " & IIf(Len(sParam) = 0, "(default)", "[" & sParam & "]") & _
                                " = " & IIf(Len(sData) <> 0, sData, "(empty)") 'doSafeURLPrefix
                            
                            If Not IsOnIgnoreList(sHit) Then
                                With result
                                    .Section = "R0"
                                    .HitLineW = sHit
                                    AddRegToFix .Reg, RESTORE_VALUE, hHive, sKey, sParam, SplitSafe(sDefDataStrings, "|")(0), CLng(Wow6432Redir)
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults result
                            End If
                        End If
                    End If
                End If
            End If
            
        Case 1  'check for present value
            
            If Reg.ValueExists(hHive, sKey, sParam, Wow6432Redir) Then
            
                sData = Reg.GetString(hHive, sKey, sParam, Wow6432Redir)
              
                If Len(sData) <> 0 Then
            
                    'check if domain is on safe list
                    bIsNSBSD = False
                    If bHideMicrosoft And Not bIgnoreAllWhitelists Then bIsNSBSD = StrBeginWithArray(sData, aSafeRegDomains)
                    'make hit
                    If Not bIsNSBSD Then
                        If InStr(1, sData, "%2e", 1) > 0 Then sData = UnEscape(sData) & " " & STR_OBFUSCATED

                        If sParam = "ProxyServer" Then
                            bProxyEnabled = (Reg.GetDword(hHive, sKey, "ProxyEnable", Wow6432Redir) = 1)
                            
                            sHit = BitPrefix("R1", HE) & " - " & _
                                HE.KeyAndHivePhysical & ": " & IIf(Len(sParam) = 0, "(default)", "[" & sParam & "]") & " = " & _
                                IIf(Len(sData) <> 0, sData, "(empty)") & IIf(bProxyEnabled, " (enabled)", " (disabled)")
                        Else
                            sHit = BitPrefix("R1", HE) & " - " & _
                                HE.KeyAndHivePhysical & ": " & IIf(Len(sParam) = 0, "(default)", "[" & sParam & "]") & " = " & _
                                IIf(Len(sData) <> 0, sData, "(empty)")   'doSafeURLPrefix
                        End If
                    
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "R1"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REMOVE_VALUE, hHive, sKey, sParam, , CLng(Wow6432Redir)
                                If sParam = "ProxyServer" Then
                                    AddRegToFix .Reg, RESTORE_VALUE, hHive, sKey, "ProxyEnable", 0, CLng(Wow6432Redir), REG_RESTORE_DWORD
                                End If
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                End If
            End If
            
        Case 2 'check if regkey is present
            If Reg.KeyExists(hHive, sKey, Wow6432Redir) Then
            
                sHit = BitPrefix("R2", HE) & " - " & HE.KeyAndHivePhysical
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "R2"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, hHive, sKey, , , CLng(Wow6432Redir)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
            End If
        End Select
    Loop
    
    'Set HE = Nothing
    
    AppendErrorLogCustom "ProcessRuleReg - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_ProcessRuleReg", "sRule=", sRule
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixRegItem(sItem$, result As SCAN_RESULT)
    'R0 - HKCU\Software\..\Main,Window Title
    'R1 - HKCU\Software\..\Main,Window Title=MSIE 5.01
    'R2 - HKCU\Software\..\Main
    FixRegistryHandler result
End Sub


'CheckR3item
Public Sub CheckR3Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckR3Item - Begin"

    Dim sURLHook$, hKey&, i&, sName$, sHit$, sCLSID$, sFile$, result As SCAN_RESULT, lret&
    Dim bHookMising As Boolean, sDefHookDll$, sDefHookCLSID$, sHookDll_1$, sHookDll_2$
    
    sURLHook = "Software\Microsoft\Internet Explorer\URLSearchHooks"
    
    sDefHookCLSID = "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}"

    sHookDll_1 = sWinSysDir & "\ieframe.dll"
    sHookDll_2 = sWinSysDir & "\shdocvw.dll"

    If OSver.MajorMinor >= 5.2 Then 'XP x64 +
        sDefHookDll = sHookDll_1
    Else
        sDefHookDll = sHookDll_2
    End If
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_HKCU Or HE_HIVE_HKU, HE_SID_USER Or HE_SID_NO_VIRTUAL, HE_REDIR_NO_WOW
    HE.AddKey sURLHook
    
    Do While HE.MoveNext
        bHookMising = False
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_QUERY_VALUE, hKey) = 0 Then
            i = 0
            sCLSID = String$(MAX_VALUENAME, 0&)
            If RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&) <> 0 Then
                sHit = "R3 - " & HE.HiveNameAndSID & ": Default URLSearchHook is missing"
                bHookMising = True
                RegCloseKey hKey
            End If
        Else
            sHit = "R3 - " & HE.HiveNameAndSID & ": Default URLSearchHook is missing"
            bHookMising = True
        End If
    
        If bHookMising Then
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "R3"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sDefHookCLSID, vbNullString, , REG_RESTORE_SZ
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID, vbNullString, "Microsoft Url Search Hook"
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID & "\InProcServer32", vbNullString, sDefHookDll
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID & "\InProcServer32", "ThreadingModel", "Apartment"
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Loop
    
    HE.Init HE_HIVE_ALL
    HE.AddKey sURLHook
    
    Do While HE.MoveNext
        
        lret = RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey)
        
        If lret = 0 Then
        
          sCLSID = String$(MAX_VALUENAME, 0&)
          i = 0
          Do While 0 = RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&)
            
            sCLSID = TrimNull(sCLSID)
            
            GetFileByCLSID sCLSID, sFile, , HE.Redirected, HE.SharedKey
            
            If Not (sCLSID = sDefHookCLSID And _
                (StrComp(sFile, sHookDll_1, 1) = 0 Or StrComp(sFile, sHookDll_2, 1) = 0)) Or (Not bHideMicrosoft) Then
                
                GetTitleByCLSID sCLSID, sName, HE.Redirected, HE.SharedKey
                
                sHit = BitPrefix("R3", HE) & " - " & HE.HiveNameAndSID & "\..\URLSearchHooks: " & _
                    sName & " - " & sCLSID & " - " & sFile
                    
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "R3"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
            
            i = i + 1
            sCLSID = String$(MAX_VALUENAME, 0&)
          Loop
          RegCloseKey hKey
        End If
    Loop
    
    AppendErrorLogCustom "CheckR3Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckR3Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixR3Item(sItem$, result As SCAN_RESULT)
    'R3 - Shitty search hook - {00000000} - c:\windows\bho.dll"
    'R3 - Default URLSearchHook is missing
    
    FixRegistryHandler result
End Sub

Public Sub CheckR4Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckR4Item - Begin"
    
    'http://ijustdoit.eu/changing-default-search-provider-in-internet-explorer-11-using-group-policies/
    
    'SearchScope
    'R4 - SearchScopes:
    
    Dim result As SCAN_RESULT, sHit$, j&, k&, sURL$, sProvider$, aScopes() As String, sBuf$, sDefScope$
    Dim Param As Variant, aData() As String, sHive$, sParam$, sDefData$
    Dim HE As clsHiveEnum, HEFix As clsHiveEnum
    Set HE = New clsHiveEnum
    Set HEFix = New clsHiveEnum
    
    'Enum custom scopes
    '
    'HKCU\Software\Policies\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Policies\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Microsoft\Internet Explorer\SearchScopes
    'HKCU\Software\Microsoft\Internet Explorer\SearchScopes
    
    HE.Init HE_HIVE_ALL, (HE_SID_ALL And Not HE_SID_SERVICE) Or HE_SID_NO_VIRTUAL, HE_REDIR_NO_WOW
    
    HE.AddKey "Software\Microsoft\Internet Explorer\SearchScopes"
    HE.AddKey "Software\Policies\Microsoft\Internet Explorer\SearchScopes"
    
    HE.Clone HEFix
    
    Dim sLastURL As String
    Dim sParams As String
    
    Do While HE.MoveNext
        
        For j = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aScopes())
            
            sProvider = Reg.GetString(HE.Hive, HE.Key & "\" & aScopes(j), "DisplayName")
            If Len(sProvider) = 0 Then sProvider = STR_NO_NAME
            
            If Left$(sProvider, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sProvider)
                If 0 <> Len(sBuf) Then sProvider = sBuf
            End If
            
            sParams = vbNullString
            sLastURL = vbNullString
            
            For Each Param In Array("URL", "SuggestionsURL_JSON", "SuggestionsURL", "SuggestionsURLFallback", "TopResultURL", "TopResultURLFallback")
            
              sURL = Reg.GetString(HE.Hive, HE.Key & "\" & aScopes(j), CStr(Param))
              
              If Len(sURL) <> 0 Or Reg.ValueExists(HE.Hive, HE.Key & "\" & aScopes(j), CStr(Param)) Then
                
                If Not IsBingScopeKeyPara("URL", sURL) Then
                  
                    With result
                        .Section = "R4"
                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aScopes(j)
                        .CureType = REGISTRY_BASED
                    End With
                    
                    If Len(sLastURL) = 0 Then 'first time?
                        sLastURL = sURL
                        sParams = CStr(Param)
                    Else
                        If sURL = sLastURL Then 'same URL ?
                            'Save several same URLs in one line by adding the list of param names at the end
                            sParams = sParams & IIf(Len(sParams) <> 0, ",", vbNullString) & CStr(Param)
                        Else 'new URL ?
                            'save last result and flush
                            sHit = "R4 - SearchScopes: " & HE.KeyAndHivePhysical & "\" & aScopes(j) & ": [" & sParams & "] = " & sLastURL & " - " & sProvider
                            
                            If Not IsOnIgnoreList(sHit) Then
                                result.HitLineW = sHit
                                GoSub CheckDefaultScope
                                AddToScanResults result, , True
                            End If
                            
                            sLastURL = sURL
                            sParams = CStr(Param)
                        End If
                    End If
                End If
              End If
            Next
            
            If Len(sParams) <> 0 And Len(sLastURL) <> 0 Then
                sHit = "R4 - SearchScopes: " & HE.KeyAndHivePhysical & "\" & aScopes(j) & ": [" & sParams & "] = " & sLastURL & " - " & sProvider
                
                If Not IsOnIgnoreList(sHit) Then
                    result.HitLineW = sHit
                    GoSub CheckDefaultScope
                    AddToScanResults result
                End If
            End If
            
        Next
    Loop
    
    AppendErrorLogCustom "CheckR4Item - End"
    
    'Set HE = Nothing
    Set HEFix = Nothing
    
    Exit Sub
    
CheckDefaultScope:
    
    HEFix.Repeat
    Do While HEFix.MoveNext
        sDefScope = Reg.GetString(HEFix.Hive, HEFix.Key, "DefaultScope")
        If Len(sDefScope) <> 0 Then
            If StrComp(sDefScope, aScopes(j), 1) = 0 Then
                If InStr(1, HEFix.Key, "Policies", 1) <> 0 Then
                    'remove policies
                    AddRegToFix result.Reg, REMOVE_VALUE, HEFix.Hive, HEFix.Key, "DefaultScope"
                Else
                    'reset default scope to bing
                    
                    For k = 1 To cReg4vals.Count
                        aData = Split(cReg4vals.Item(k), ",", 3)
                        sHive = aData(0)
                        sParam = aData(1)
                        sDefData = SplitSafe(aData(2), "|")(0)
        
                        If (HEFix.HiveName = "HKLM" And sHive = "HKLM") Or _
                            (HEFix.HiveName <> "HKLM" And sHive = "HKCU") Then
                
                            AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", sParam, sDefData, , REG_RESTORE_SZ
                        End If
                    Next
            
                    AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key, "DefaultScope", "{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", , REG_RESTORE_SZ
                    AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "DisplayName", "Bing", , REG_RESTORE_SZ
                    If HEFix.HiveName = "HKLM" Then
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", vbNullString, "Bing", , REG_RESTORE_SZ
                    Else
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "FaviconURL", "https://www.bing.com/favicon.ico", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "FaviconURLFallback", "https://www.bing.com/favicon.ico", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTLogoPath", AppDataLocalLow & "\Microsoft\Internet Explorer\Services\", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTLogoURL", "https://go.microsoft.com/fwlink/?LinkID=403856&language={language}&scale={scalelevel}&contrast={contrast}", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTSuggestionsURL", "https://api.bing.com/qsml.aspx?query={searchTerms}&market={language}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IENTSS", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTTopResultURL", "https://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTTR", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTURL", "https://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTSR", , REG_RESTORE_SZ
                    End If
                End If
            End If
        End If
    Loop
    
    Return
ErrorHandler:
    ErrorMsg Err, "CheckR4Item"
    If inIDE Then Stop: Resume Next
End Sub

Private Function IsBingScopeKeyPara(sRegParam As String, sURL As String) As Boolean
    If Len(sURL) = 0 Then Exit Function
    
    If Not bHideMicrosoft Then Exit Function
    
    'Is valid domain
    Dim pos As Long, sPrefix As String
    pos = InStr(sURL, "?")
    If pos = 0 Then Exit Function
    sPrefix = Left$(sURL, pos - 1)
    Select Case sPrefix
    Case "http://search.microsoft.com/results.aspx"
    Case "http://www.bing.com/search"
    Case "http://www.bing.com/as/api/qsml"
    Case "http://api.bing.com/qsml.aspx"
    Case "http://search.live.com/results.aspx"
    Case "http://api.search.live.com/qsml.aspx"
    Case Else
        Exit Function
    End Select

    Dim aKey() As String, aVal() As String, i As Long
    Dim bSearchTermPresent As Boolean
    
    IsBingScopeKeyPara = True
    
    If StrEndWith(sURL, ";") Then sURL = Left$(sURL, Len(sURL) - 1)
    
    Call ParseKeysURL(sURL, aKey, aVal)
    
    Select Case UCase$(sRegParam)
    
        Case "URL", UCase$("SuggestionsURL"), UCase$("SuggestionsURLFallback"), UCase$("TopResultURL"), UCase$("TopResultURLFallback")
            If AryItems(aKey) Then
                For i = 0 To UBound(aKey)
                    Select Case LCase$(aKey(i))
                    Case "q", "query"
                    '{searchTerms}
                        If StrComp(aVal(i), "{searchTerms}", 1) = 0 Then bSearchTermPresent = True
                    
                    Case "src"
                    'IE-SearchBox
                    'IE11TR
                    '{referrer:source?}
                    'IE10TR
                    'src=ie9tr
                    'IE-TopResult 'for TopResultURL
                        If StrBeginWith(aVal(i), "IE") Then
                            If Len(aVal(i)) > 6 Then
                                If StrComp(aVal(i), "IE-SearchBox", 1) = 0 Then
                                ElseIf StrComp(aVal(i), "IE-TopResult", 1) = 0 Then
                                Else
                                    IsBingScopeKeyPara = False
                                End If
                            End If
                        ElseIf StrComp(aVal(i), "{referrer:source?}", 1) = 0 Then
                        Else
                            IsBingScopeKeyPara = False
                        End If
                    
                    Case "form"
                    'IE8SRC
                    'IE10SR
                    'IE11SR
                    'IESR02
                    'IESS02
                    'IE8SSC
                    'IE11SS
                    'IETR02 'for TopResultURL
                    'IE11TR 'for TopResultURL
                    'IE10TR 'for TopResultURL
                    'SKY2DF
                    'PRHPR1
                    'MSERBM
                    'IE8SRC
                    'HPNTDF
                    'MSSEDF
                    'APBTDF
                    'SK216DF
                        If Len(aVal(i)) > 7 Then IsBingScopeKeyPara = False
                    
                      Case "pc"
                    'HRTS
                    'MSERT1
                    'HPNTDF
                    'MSE1
                    'MAPB
                    'MAARJS
                    'MAMIJS;
                    'CMDTDFJS
                        If Len(aVal(i)) > 8 Then IsBingScopeKeyPara = False
                    
                    Case "maxwidth"
                    '{ie:maxWidth}
                        If StrComp(aVal(i), "{ie:maxWidth}", 1) <> 0 Then IsBingScopeKeyPara = False
                    
                    Case "rowheight"
                    '{ie:rowHeight}
                        If StrComp(aVal(i), "{ie:rowHeight}", 1) <> 0 Then IsBingScopeKeyPara = False
                    
                    Case "sectionheight"
                    '{ie:sectionHeight}
                        If StrComp(aVal(i), "{ie:sectionHeight}", 1) <> 0 Then IsBingScopeKeyPara = False
                    
                    Case "market"
                    '{language}
                        If StrComp(aVal(i), "{language}", 1) <> 0 Then IsBingScopeKeyPara = False
                    
                    Case "mkt"
                    Case "setlang"
                    Case "ptag"
                    Case "conlogo"
                    
                    Case vbNullString
                        If Len(aVal(i)) > 0 Then IsBingScopeKeyPara = False
                    
                    Case Else
                        IsBingScopeKeyPara = False
                    End Select
                Next
            End If
        
        Case Else
            IsBingScopeKeyPara = False
    End Select
    
    If Not bSearchTermPresent Then IsBingScopeKeyPara = False
End Function

Public Sub FixR4Item(sItem$, result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    FixRegistryHandler result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixR4Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckFileItems(ByVal sRule$)
    On Error GoTo ErrorHandler:
    
    Dim vRule As Variant, iMode&, sHit$, result As SCAN_RESULT
    Dim sFile$, sSection$, sParam$, sData$, sLegitData$
    Dim sTmp$
    
    AppendErrorLogCustom "CheckFileItems - Begin", "Rule: " & sRule
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    'IniFile rule syntax:
    '[inifile],[section],[value],[default data],[infected data]
    '* [inifile]          = "" -> abort
    ' * [section]         = "" -> abort
    '  * [value]          = "" -> abort
    '   * [default data]  = "" -> delete if found
    '    * [infected data]= "" -> fix if infected
    
    'decrypt rule
    'sRule = Crypt(sRule)
    
    'Checking white list rules
    '1-st token should contains .ini
    'total number of tokens should be 5 (0 to 4)
    
    vRule = Split(sRule, ";")
    If UBound(vRule) <> 4 Or InStr(CStr(vRule(0)), ".ini") = 0 Then
        If Not bAutoLogSilent Then
            MsgBoxW "CheckFileItems: Spelling error or decrypting error for: " & sRule
        End If
        Exit Sub
    End If
    
    '1,2,3 tokens should not be empty
    '4-th token is empty -> check if value is present     (F1)
    '4-th token is present -> check if value is infected  (F0)
    
    'File checking rules:
    '
    'example:
    '--------------
    '1. system.ini    (file)
    '2. boot          (section)
    '3. Shell         (parameter)
    '4. explorer.exe  (data / value)
    
    If CStr(vRule(0)) = vbNullString Then Exit Sub
    If CStr(vRule(1)) = vbNullString Then Exit Sub
    If CStr(vRule(2)) = vbNullString Then Exit Sub
    If CStr(vRule(4)) = vbNullString Then iMode = 0
    If CStr(vRule(3)) = vbNullString Then iMode = 1
    
    sFile = vRule(0)
    sSection = vRule(1)
    sParam = vRule(2)
    sLegitData = vRule(3)
    
    'Registry checking rules (prefix REG: on 1-st token)
    '
    'example:
    '1. REG:system.ini ()
    '2. boot           (section)
    '3. Shell          (parameter)
    '4. explorer.exe   (data / value)
    
    'if 4-th token is empty -> check if value is present, in the Registry      (F3)
    'if 4-th token is present -> check if value is infected, in the Registry   (F2)
    
'    ' adding char "," to each value 'UserInit'
'    If InStr(1, sLegitData, "UserInit", 1) <> 0 Then
'        arr = Split(sLegitData, "|")
'        For i = 0 To UBound(arr)
'            sTmp = sTmp & arr(i) & ",|"
'        Next
'        sTmp = Left$(sTmp, Len(sTmp) - 1)
'        sLegitData = sLegitData & "|" & sTmp
'    End If
    
    If Left$(sFile, 3) = "REG" Then
        'skip Win9x
        If Not bIsWinNT Then Exit Sub
        If CStr(vRule(4)) = vbNullString Then iMode = 2
        If CStr(vRule(3)) = vbNullString Then iMode = 3
    End If
    
    'iMode:
    ' F0 = check if value is infected (file)
    ' F1 = check if value is present (file)
    ' F2 = check if value is infected, in the Registry
    ' F3 = check if value is present, in the Registry
    
    Select Case iMode
        Case 0
            'F0 = check if value is infected (file)
            'sValue = String$(255, " ")
            'GetPrivateProfileString CStr(vRule(1)), CStr(vRule(2)), "", sValue, 255, CStr(vRule(0))
            'sValue = Rtrim$(sValue)
            
            If Not FileExists(sFile) Then
                sFile = FindOnPath(sFile, True)
            End If
            
            sData = IniGetString(sFile, sSection, sParam)
            sData = Trim$(RTrimNull(sData))
            
            If bIsWinNT And Len(sData) <> 0 Then
            
                If Not inArraySerialized(sData, sLegitData, "|", , , vbTextCompare) Or (Not bHideMicrosoft) Then
                    
                    SignVerifyJack sData, result.SignResult
                    sHit = "F0 - " & sFile & ": " & "[" & sSection & "]" & " " & sParam & " = " & sData & FormatSign(result.SignResult)
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "F0"
                            .HitLineW = sHit
                            'system.ini
                            AddIniToFix .Reg, RESTORE_VALUE_INI, sFile, "boot", "shell", SplitSafe(sLegitData, "|")(0)  '"explorer.exe"
                            .CureType = INI_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
            
        Case 1
            'F1 = check if value is present (file)
            'sValue = String$(255, " ")
            'GetPrivateProfileString CStr(vRule(1)), CStr(vRule(2)), "", sValue, 255, CStr(vRule(0))
            'sValue = Rtrim$(sValue)
            
            If Not FileExists(sFile) Then
                sFile = FindOnPath(sFile, True)
            End If
            
            sData = IniGetString(sFile, sSection, sParam)
            sData = Trim$(RTrimNull(sData))
            
            If Len(sData) <> 0 Then
                SignVerifyJack sData, result.SignResult
                sHit = "F1 - " & sFile & ": " & "[" & sSection & "]" & " " & sParam & " = " & sData & FormatSign(result.SignResult)
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "F1"
                        .HitLineW = sHit
                        'win.ini
                        AddIniToFix .Reg, RESTORE_VALUE_INI, sFile, "windows", sParam, vbNullString 'param = 'load' or 'run'
                        .CureType = INI_BASED
                    End With
                    AddToScanResults result
                End If
            End If
            
        Case 2
            'F2 = check if value is infected, in the Registry
            'so far F2 is only reg:Shell and reg:UserInit
            
            HE.Init HE_HIVE_ALL
            HE.AddKey "Software\Microsoft\Windows NT\CurrentVersion\WinLogon"
            
            Do While HE.MoveNext
                
                sData = Reg.GetString(HE.Hive, HE.Key, sParam, HE.Redirected)
                sTmp = sData
                If Right$(sData, 1) = "," Then sTmp = Left$(sTmp, Len(sTmp) - 1)
                
                'Note: HKCU + empty values are allowed
                If (Not inArraySerialized(sTmp, sLegitData, "|", , , vbTextCompare) Or (Not bHideMicrosoft)) And _
                  Not ((HE.Hive = HKCU Or HE.Hive = HKU) And Len(sData) = 0) Then
            
                    'exclude no WOW64 value on Win10 for UserInit
                    If Not (HE.Redirected And OSver.MajorMinor >= 10 And sParam = "UserInit" And Len(sData) = 0) Then
                    If Not (HE.Redirected And sParam = "UserInit" And StrComp(sData, BuildPath(sWinSysDirWow64, "userinit.exe")) = 0) Then
                        
                        SignVerifyJack sData, result.SignResult
                        
                        sHit = BitPrefix("F2", HE) & " - " & HE.HiveNameAndSID & "\..\WinLogon: " & _
                            "[" & sParam & "] = " & sData & FormatSign(result.SignResult)
                            
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "F2"
                                .HitLineW = sHit
                                AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sParam, SplitSafe(sLegitData, "|")(0), HE.Redirected
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                    End If
                End If
            Loop
            
        Case 3
            'F3 = check if value is present, in the Registry
            'this is not really smart when more INIFile items get
            'added, but so far F3 is only reg:load and reg:run
        
            HE.Init HE_HIVE_ALL
            HE.AddKey "Software\Microsoft\Windows NT\CurrentVersion\Windows"
            
            Do While HE.MoveNext
            
                sData = Reg.GetString(HE.Hive, HE.Key, sParam, HE.Redirected)
                If 0 <> Len(sData) Then
                    SignVerifyJack sData, result.SignResult
                    sHit = BitPrefix("F3", HE) & " - " & HE.HiveNameAndSID & "\..\Windows: " & _
                        "[" & sParam & "] = " & sData & FormatSign(result.SignResult)
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "F3"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Loop
    End Select
    
    AppendErrorLogCustom "CheckFileItems - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_ProcessRuleIniFile", "sRule=", sRule
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixFileItem(sItem$, result As SCAN_RESULT)
    'F0 - system.ini: Shell=c:\win98\explorer.exe openme.exe
    'F1 - win.ini: load=hpfsch
    'F2, F3 - registry

    'coding is easy if you cheat :)
    '(c) Dragokas: Cheaters will be punished ^_^
    
    FixRegistryHandler result
End Sub

Public Sub CheckO1Item_DNSApi()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item_DNSApi - Begin"
    
    If OSver.MajorMinor <= 5 Then Exit Sub 'XP+ only
    
    Const MaxSize As Long = 5242880 ' 5 MB.
    
    Dim vFile As Variant, ff As Long, Size As Currency, p As Long, buf() As Byte, sHit As String, result As SCAN_RESULT
    Dim bufExample() As Byte
    Dim bufExample_2() As Byte
    
    bufExample = StrConv(LCase$("\drivers\etc\hosts"), vbFromUnicode)
    bufExample_2 = StrConv(UCase$("\drivers\etc\hosts"), vbFromUnicode)
    
    ToggleWow64FSRedirection False
    
    For Each vFile In Array(sWinDir & "\system32\dnsapi.dll", sWinDir & "\syswow64\dnsapi.dll")
    
        If OSver.IsWin32 And InStr(1, vFile, "syswow64", 1) <> 0 Then Exit For

        If OpenW(CStr(vFile), FOR_READ, ff) Then
            
            Size = LOFW(ff)
            
            If Size > MaxSize Then
                ErrorMsg Err, "modMain_CheckO1Item_DNSApi", "File is too big: " & vFile & " (Allowed: " & MaxSize & " byte max., current is: " & Size & "byte.)"
            ElseIf Size > 0 Then
                
                ReDim buf(Size - 1)
                
                If GetW(ff, 1, , VarPtr(buf(0)), CLng(Size)) Then
                
                    p = InArrSign_NoCase(0, buf, bufExample, bufExample_2)
                    
                    If p = -1 Then                      '//TODO: add isMicrosoftFile() ?
                        ' if signature not found
                        sHit = "O1 - DNSApi: File is patched - " & vFile
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O1"
                                .HitLineW = sHit
                                AddFileToFix .File, RESTORE_FILE_SFC, CStr(vFile)
                                .CureType = FILE_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                End If
            End If
            CloseW ff
        End If
    Next
    ToggleWow64FSRedirection True
    AppendErrorLogCustom "CheckO1Item_DNSApi - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO1Item_DNSApi"
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub

Public Function InArrString(pos As Long, ArrSrc() As Byte, StrExample As String, CompareMethod As VbCompareMethod) As Long
    If CompareMethod = vbBinaryCompare Then
        Dim aExample() As Byte
        aExample = StrConv(StrExample, vbFromUnicode)
        InArrString = InArrSign(pos, ArrSrc, aExample)
    Else
        Dim aLCase() As Byte
        Dim aUCase() As Byte
        aLCase = StrConv(LCase$(StrExample), vbFromUnicode)
        aUCase = StrConv(UCase$(StrExample), vbFromUnicode)
        InArrString = InArrSign_NoCase(pos, ArrSrc, aLCase, aUCase)
    End If
End Function

Private Function InArrSign(pos As Long, ArrSrc() As Byte, ArrEx() As Byte) As Long
    Dim i As Long, j As Long, p As Long, Found As Boolean
    InArrSign = -1
    For i = pos To UBound(ArrSrc) - UBound(ArrEx)
        p = i
        Found = True
        For j = 0 To UBound(ArrEx)
            If ArrSrc(p) <> ArrEx(j) Then Found = False: Exit For
            p = p + 1
        Next
        If Found Then InArrSign = p - UBound(ArrEx) - 1: Exit For
    Next
End Function

Private Function InArrSign_NoCase(pos As Long, ArrSrc() As Byte, ArrEx() As Byte, ArrEx_2() As Byte) As Long
    'ArrEx - all lcase
    'ArrEx_2 - all Ucase
    Dim i As Long, j As Long, p As Long, Found As Boolean
    InArrSign_NoCase = -1
    For i = pos To UBound(ArrSrc) - UBound(ArrEx)
        p = i
        Found = True
        For j = 0 To UBound(ArrEx)
            If ArrSrc(p) <> ArrEx(j) And _
                ArrSrc(p) <> ArrEx_2(j) Then Found = False: Exit For
            p = p + 1
        Next
        If Found Then InArrSign_NoCase = p - UBound(ArrEx) - 1: Exit For
    Next
End Function

Public Sub CheckO1Item_ICS()
    ' hosts.ics
    'https://support.microsoft.com/ru-ru/kb/309642
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item_ICS - Begin"
    
    Dim sHostsFileICS$, sHit$, sHostsFileICS_Default$
    Dim sLines$, sLine As Variant, NonDefaultPath As Boolean, cFileSize As Currency, hFile As Long
    Dim result As SCAN_RESULT
    
    If bIsWin9x Then sHostsFileICS_Default = sWinDir & "\hosts.ics"
    If bIsWinNT Then sHostsFileICS_Default = sWinDir & "\System32\drivers\etc\hosts.ics"
    
    sHostsFileICS = g_HostsFile & ".ics"
    
    If StrComp(sHostsFileICS, sHostsFileICS_Default) <> 0 Then
        NonDefaultPath = True
    End If
    
    If NonDefaultPath Then                              'Note: \System32\drivers\etc is not under Wow6432 redirection
        ToggleWow64FSRedirection False, sHostsFileICS
    End If
    
    cFileSize = FileLenW(sHostsFileICS)
    
    ' Size = 0 or just not exists
    If cFileSize = 0 Then
        ToggleWow64FSRedirection True
        
        If NonDefaultPath Then
            GoTo CheckHostsICS_Default:
        Else
            Exit Sub
        End If
    End If
    
    If OpenW(sHostsFileICS, FOR_READ, hFile, g_FileBackupFlag) Then
        CloseW hFile
        sLines = ReadFileContents(sHostsFileICS, False)
        ToggleWow64FSRedirection True
    Else
    
        sHit = "O1 - Unable to read Hosts.ICS file"
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFileICS
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If

        ToggleWow64FSRedirection True
        If NonDefaultPath Then
            GoTo CheckHostsICS_Default:
        Else
            Exit Sub
        End If
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    For Each sLine In Split(sLines, vbLf)
        sLine = Replace$(sLine, vbTab, " ")
        sLine = Replace$(sLine, vbCr, vbNullString)
        sLine = Trim$(sLine)
        
        If sLine <> vbNullString Then
            If Left$(sLine, 1) <> "#" Then
                Do
                    sLine = Replace$(sLine, "  ", " ")
                Loop Until InStr(sLine, "  ") = 0
                
                sHit = "O1 - Hosts.ICS: " & sLine
                'If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit

                '// TODO: ������� ����� ������ ������ ���������� ��������.
                '������ ��� � ��� ��������� ��������, �� ����� ����� ������ ���������� ����������� ���� ���������������
                '�� ������� ����, � ��������� ������.
                '��� ���� ������������� �������� ���� ������� ������ (�.�. ��� ��� ������ ���� ����� ����� ������� � ������� AddToScanResultsSimple)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, sHostsFileICS
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
                End If

            End If
        End If
    Next
    
CheckHostsICS_Default:
    
    ToggleWow64FSRedirection True

    If Not NonDefaultPath Then Exit Sub
    
    cFileSize = FileLenW(sHostsFileICS_Default)
    
    ' Size = 0 or just not exists
    If cFileSize = 0 Then Exit Sub
    
    If OpenW(sHostsFileICS_Default, FOR_READ, hFile, g_FileBackupFlag) Then
        CloseW hFile
        sLines = ReadFileContents(sHostsFileICS_Default, False)
    Else
        sHit = "O1 - Unable to read Hosts.ICS default file"
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFileICS_Default
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
        
        Exit Sub
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    For Each sLine In Split(sLines, vbLf)
        sLine = Replace$(sLine, vbTab, " ")
        sLine = Replace$(sLine, vbCr, vbNullString)
        sLine = Trim$(sLine)
        
        If sLine <> vbNullString Then
            If Left$(sLine, 1) <> "#" Then
                Do
                    sLine = Replace$(sLine, "  ", " ")
                Loop Until InStr(sLine, "  ") = 0
                
                sHit = "O1 - Hosts.ICS default: " & sLine
                'If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O1", sHit
                
                '// TODO: ������� ����� ������ ������ ���������� ��������.
                '������ ��� � ��� ��������� ��������, �� ����� ����� ������ ���������� ����������� ���� ���������������
                '�� ������� ����, � ��������� ������.
                '��� ���� ������������� �������� ���� ������� ������ (�.�. ��� ��� ������ ���� ����� ����� ������� � ������� AddToScanResultsSimple)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, sHostsFileICS_Default
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
                End If
                
            End If
        End If
    Next
    
    AppendErrorLogCustom "CheckO1Item_ICS - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO1Item_ICS"
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckO1Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item - Begin"
    
    Dim sHit$, i&, HostsDefaultFile$, NonDefaultPath As Boolean
    Dim sLine As Variant, sLines$, cFileSize@
    Dim aHits() As String, j As Long, hFile As Long
    Dim HostsDefaultPath As String
    ReDim aHits(0)
    Dim result As SCAN_RESULT
    
    '// TODO: Add UTF8.
    'http://serverfault.com/questions/452268/hosts-file-ignored-how-to-troubleshoot
    
    Dbg "1"
    
    GetHosts
    
    If bIsWin9x Then HostsDefaultFile = sWinDir & "\hosts"
    If bIsWinNT Then HostsDefaultFile = sWinDir & "\System32\drivers\etc\hosts"
    
    Dbg "2"
    
    If StrComp(g_HostsFile, HostsDefaultFile) <> 0 Then
        'sHit = "O1 - Hosts file is located at: " & sHostsFile
        sHit = "O1 - " & Translate(271) & ": " & g_HostsFile
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                HostsDefaultPath = EnvironUnexpand(GetParentDir(HostsDefaultFile))
                AddRegToFix .Reg, RESTORE_VALUE, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", _
                  HostsDefaultPath, , REG_RESTORE_EXPAND_SZ
                .CureType = REGISTRY_BASED Or CUSTOM_BASED
            End With
            AddToScanResults result
        End If
        NonDefaultPath = True
    End If
    
    Dbg "3"
    
    If NonDefaultPath Then                              'Note: \System32\drivers\etc is not under Wow6432 redirection
        ToggleWow64FSRedirection False, g_HostsFile
    End If
    
    Dbg "4"
    
    cFileSize = FileLenW(g_HostsFile)
    
    If cFileSize = 0 Then
        If NonDefaultPath Then
            'Check default path also
            GoTo CheckHostsDefault:
        Else
            sHit = "O1 - Hosts: is empty"

            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O1"
                    .HitLineW = sHit
                    AddFileToFix .File, BACKUP_FILE, g_HostsFile
                    .CureType = CUSTOM_BASED
                End With
                AddToScanResults result
            End If
            
            ToggleWow64FSRedirection True
            Exit Sub
        End If
    End If
    
    Dbg "5"
    
    If OpenW(g_HostsFile, FOR_READ, hFile, g_FileBackupFlag) Then
        CloseW hFile
        sLines = ReadFileContents(g_HostsFile, False)  'speed up
        ToggleWow64FSRedirection True
    Else
        ToggleWow64FSRedirection True
    
        sHit = "O1 - Hosts: Unable to read Hosts file"
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, g_HostsFile
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
        
        If NonDefaultPath Then
            GoTo CheckHostsDefault:
        Else
            Exit Sub
        End If
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    If Len(Replace$(sLines, vbNullChar, vbNullString)) = 0 Then
    
        sHit = "O1 - Hosts: is damaged (contains NUL characters only)"
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, g_HostsFile
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
        
        If NonDefaultPath Then
            GoTo CheckHostsDefault:
        Else
            Exit Sub
        End If
    End If
    
    i = 0
    
    Dbg "6"
    
    For Each sLine In Split(sLines, vbLf)
            
            'ignore all lines that start with loopback
            '(127.0.0.1), null (0.0.0.0) and private IPs
            '(192.168. / 10.)
            sLine = Replace$(sLine, vbTab, " ")
            sLine = Replace$(sLine, vbCr, vbNullString)
            sLine = Trim$(sLine)
            
            If sLine <> vbNullString Then
                'If InStr(sLine, "127.0.0.1") <> 1 And _
                '   InStr(sLine, "0.0.0.0") <> 1 And _
                '   InStr(sLine, "192.168.") <> 1 And _
                '   InStr(sLine, "10.") <> 1 And _
                '   InStr(sLine, "#") <> 1 And _
                '   Not (bIgnoreSafeDomains And InStr(sLine, "216.239.37.101") > 0) Or _
                '   bIgnoreAllWhitelists Then
                    '216.239.37.101 = google.com
                    
                '::1 - default for Vista
                If (Left$(sLine, 1) <> "#" Or bIgnoreAllWhitelists) And _
                  ((StrComp(sLine, "127.0.0.1       localhost", 1) <> 0 And _
                  StrComp(sLine, "::1             localhost", 1) <> 0 And _
                  StrComp(sLine, "127.0.0.1 localhost", 1) <> 0) Or Not bHideMicrosoft) Then
                  
                    Do
                        sLine = Replace$(sLine, "  ", " ")
                    Loop Until InStr(sLine, "  ") = 0
                    
                    sHit = "O1 - Hosts: " & sLine
                    If Not IsOnIgnoreList(sHit) Then
                        'AddToScanResultsSimple "O1", sHit
                        If UBound(aHits) < i Then ReDim Preserve aHits(UBound(aHits) + 100)
                        aHits(i) = sHit
                        i = i + 1
                    End If
                    
                End If
            End If
    Next

    Dbg "7"

    If i > 0 Then
        If i >= 10 Then
            If Not NonDefaultPath Then
                sHit = "O1 - Hosts: Reset contents to default"
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, g_HostsFile
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        End If
'        'maximum 100 hosts entries
'        If i <= 100 Then
'            For j = 0 To i - 1
'                AddToScanResultsSimple "O1", aHits(j)
'            Next
'        Else
'            sHit = "O1 - Hosts: has " & i & " entries"
'        End If
        For j = 0 To i - 1
        
            'AddToScanResultsSimple "O1", aHits(j), IIf((j < 20) Or (j > i - 1 - 20), False, True)
        
            '// TODO: ������� ����� ������ ������ ���������� ��������.
            '������ ��� � ��� ��������� ��������, �� ����� ����� ������ ���������� ����������� ���� ���������������
            '�� ������� ����, � ��������� ������.
            '��� ���� ������������� �������� ���� ������� ������ (�.�. ��� ��� ������ ���� ����� ����� ������� � ������� AddToScanResultsSimple)
        
            sHit = aHits(j)
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, g_HostsFile
                .CureType = CUSTOM_BASED
            End With
            'limit for first and last 20 entries only to view on results window
            AddToScanResults result, IIf((j < 20) Or (j > i - 1 - 20), False, True)
        Next
    End If
    
    ReDim aHits(0)

CheckHostsDefault:
    'if Hosts was redirected -> checking records on default hosts also. ( Prefix "O1 - Hosts default: " )
    
    i = 0
    
    ToggleWow64FSRedirection True
    
    Dbg "8"
    
    If NonDefaultPath Then
        
        If FileExists(HostsDefaultFile) Then
            
            cFileSize = FileLenW(HostsDefaultFile)
            If cFileSize <> 0 Then

                Dbg "9"

                If OpenW(HostsDefaultFile, FOR_READ, hFile, g_FileBackupFlag) Then
                    CloseW hFile
                    sLines = ReadFileContents(HostsDefaultFile, False)
                Else
                    sHit = "O1 - Hosts default: Unable to read Default Hosts file"

                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O1"
                            .HitLineW = sHit
                            AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                            .CureType = CUSTOM_BASED
                        End With
                        AddToScanResults result
                    End If
                    
                    Exit Sub
                End If
                
                Dbg "10"
                
                sLines = Replace$(sLines, vbCrLf, vbLf)
                
                If Len(Replace$(sLines, vbNullChar, vbNullString)) = 0 Then
    
                    sHit = "O1 - Hosts default: is damaged (contains NUL characters only)"
        
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O1"
                            .HitLineW = sHit
                            AddFileToFix .File, BACKUP_FILE, g_HostsFile
                            .CureType = CUSTOM_BASED
                        End With
                        AddToScanResults result
                    End If
                    
                    Exit Sub
                End If

                For Each sLine In Split(sLines, vbLf)
                
                    sLine = Replace$(sLine, vbTab, " ")
                    sLine = Replace$(sLine, vbCr, vbNullString)
                    sLine = Trim$(sLine)
                    
                    If sLine <> vbNullString Then
                    
                        If (Left$(sLine, 1) <> "#" Or bIgnoreAllWhitelists) And _
                          ((StrComp(sLine, "127.0.0.1       localhost", 1) <> 0 And _
                          StrComp(sLine, "::1             localhost", 1) <> 0) Or Not bHideMicrosoft) Then    '::1 - default for Vista
                            Do
                                sLine = Replace$(sLine, "  ", " ")
                            Loop Until InStr(sLine, "  ") = 0
                    
                            Dbg "11"
                    
                            sHit = "O1 - Hosts default: " & sLine
                            If Not IsOnIgnoreList(sHit) Then
                                'AddToScanResultsSimple "O1", sHit
                                If UBound(aHits) < i Then ReDim Preserve aHits(UBound(aHits) + 100)
                                aHits(i) = sHit
                                i = i + 1
                            End If
                        End If
                    End If
                Next
            End If
        End If
        
        Dbg "12"
        
        If i > 0 Then
            If i >= 10 Then
                sHit = "O1 - Hosts default: Reset contents to default"
                    
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
                End If
            End If
'            'maximum 100 hosts entries
'            If i <= 100 Then
'                 For j = 0 To i - 1
'                    AddToScanResultsSimple "O1", aHits(j)
'                 Next
'            Else
'                sHit = "O1 - Hosts default: has " & i & " entries"
'            End If
            For j = 0 To i - 1
                
                'AddToScanResultsSimple "O1", aHits(j), IIf((j < 20) Or (j > i - 1 - 20), False, True)
                
                '// TODO: ������� ����� ������ ������ ���������� ��������.
                '������ ��� � ��� ��������� ��������, �� ����� ����� ������ ���������� ����������� ���� ���������������
                '�� ������� ����, � ��������� ������.
                '��� ���� ������������� �������� ���� ������� ������ (�.�. ��� ��� ������ ���� ����� ����� ������� � ������� AddToScanResultsSimple)
            
                sHit = aHits(j)
                With result
                    .Section = "O1"
                    .HitLineW = sHit
                    AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                    .CureType = CUSTOM_BASED
                End With
                'limit for first and last 20 entries only to view on results window
                AddToScanResults result, IIf((j < 20) Or (j > i - 1 - 20), False, True)
            Next
        End If
    End If

    AppendErrorLogCustom "CheckO1Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO1Item"
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO1Item(sItem$, result As SCAN_RESULT)
    'O1 - Hijack of auto.search.msn.com etc with Hosts file
    On Error GoTo ErrorHandler:
    Dim sLine As Variant, sHijacker$, i&, iAttr&, ff1%, ff2%, HostsDefaultPath$, sLines$, HostsDefaultFile$, cFileSize@, sHosts$
    Dim sHostsTemp$, bResetHosts As Boolean, aLines() As String, isICS As Boolean
    
    If InStr(1, sItem, "O1 - DNSApi:", 1) <> 0 Then
        FixFileHandler result
        Exit Sub
    End If
    
    If bIsWin9x Then HostsDefaultPath = sWinDir
    If bIsWinNT Then HostsDefaultPath = "%SystemRoot%\System32\drivers\etc"
    
    HostsDefaultFile = EnvironW(HostsDefaultPath & "\" & "hosts")
    
    'If InStr(sItem, "Hosts file is located at") > 0 Then
    If InStr(sItem, Translate(271)) > 0 Then
        Reg.SetExpandStringVal HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", HostsDefaultPath
        GetHosts    'reload var. 'sHostsFile'
        Exit Sub
    End If
    
    If StrComp(sItem, "O1 - Hosts: is empty", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts: Unable to read Hosts file", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts default: Unable to read Default Hosts file", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts: Reset contents to default", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts default: Reset contents to default", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts: is damaged (contains NUL characters only)", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts default: is damaged (contains NUL characters only)", 1) = 0 Then
        bResetHosts = True
    End If
    
    If StrBeginWith(sItem, "O1 - Hosts default: ") Or bResetHosts Then
        
        sHosts = HostsDefaultFile   'default hosts path
    Else
        sHosts = g_HostsFile         'path that may be redirected
    End If
    
    If StrBeginWith(sItem, "O1 - Hosts.ICS: ") Then
        sHosts = g_HostsFile & ".ics"
        isICS = True
    ElseIf StrBeginWith(sItem, "O1 - Hosts.ICS default: ") Then
        sHosts = HostsDefaultFile & ".ics"
        isICS = True
    End If
    
    sHostsTemp = TempCU & "\" & "hosts.new"
    If Not CheckFileAccessWrite_Physically(sHostsTemp) Then
        sHostsTemp = BuildPath(AppPath(), "hosts.new")
    End If
    
    If FileExists(sHostsTemp) Then
        DeleteFileForce sHostsTemp
    End If
    
    If StrComp(GetParentDir(sHosts), sWinDir & "\System32\drivers\etc\hosts", 1) <> 0 Then
        ToggleWow64FSRedirection False
    End If
    
    cFileSize = FileLenW(sHosts)
    
    If cFileSize = 0 Or bResetHosts Then
        'no reset for ICS for now
        If isICS Then GoTo Finalize
        '2.0.7. - Reset Hosts to its default contents
        ff2 = FreeFile()
        Open sHostsTemp For Output As #ff2
            Print #ff2, GetDefaultHostsContents()
        Close #ff2
        GoTo Replace
    End If
    
    'If Not StrBeginWith(sItem, "O1 - Hosts: ") Then Exit Sub
    
    'parse to server name
    ' Example: 127.0.0.1 my.dragokas.com -> var. 'sHijacker' = "my.dragokas.com"
    sHijacker = mid$(sItem, InStr(sItem, ":") + 2)
    sHijacker = Trim$(sHijacker)
    If Not isICS Then
        If InStr(sHijacker, " ") > 0 Then
            Dim sTemp$
            sTemp = mid$(sHijacker, InStr(sHijacker, " ") + 1)
            If 0 <> Len(sTemp) Then sHijacker = sTemp
        End If
    End If
    
    'Reset attributes
    SetFileAttributes StrPtr(sHosts), vbNormal
    
    BackupFile result, sHosts
    
    'read current hosts file
    ff1 = FreeFile()
    Open sHosts For Binary Access Read As #ff1
    sLines = String$(LOF(ff1), 0)
    Get #ff1, , sLines
    Close #ff1
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    'build new hosts file (exclude bad lines)
    ff2 = FreeFile()
    Open sHostsTemp For Output As #ff2
        aLines = Split(sLines, vbLf)
          For i = 0 To UBoundSafe(aLines)
            sLine = aLines(i)
            sLine = Replace$(sLine, vbTab, " ")
            sLine = Replace$(sLine, vbCr, vbNullString)
            Do
                sLine = Replace$(sLine, "  ", " ")
            Loop Until InStr(sLine, "  ") = 0
            If InStr(1, sLine, sHijacker, 1) <> 0 Then
                'don't write line to hosts file
            Else
                'skip last empty line
                If 0 <> Len(sLine) Or (0 = Len(sLine) And i < UBound(aLines)) Then Print #ff2, aLines(i)
            End If
          Next
    Close #ff2
    
Replace:
    If DeleteFileForce(sHosts) Then
        
        If StrComp(GetParentDir(sHosts), sWinDir & "\System32\drivers\etc\hosts", 1) <> 0 Then
            ToggleWow64FSRedirection False
        End If
        
        If 0 = MoveFile(StrPtr(sHostsTemp), StrPtr(sHosts)) Then
            If Err.LastDllError = 5 Then Err.Raise 70
        End If
    Else
        Err.Raise 70
    End If
    
    SetFileStringSD sHosts, "O:SYG:SYD:AI(A;ID;FA;;;SY)(A;ID;FA;;;BA)(A;ID;0x1200a9;;;BU)", False
    
    '//TODO:
    'clear cache
    
    '1. Mozilla Firefox
    '%LocalAppData%\Mozilla\Firefox\Profiles\<Name>\cache2 -> rename to *.bak
    
    '2. Microsoft Internet Explorer
    
    '3. Google Chrome
    
    '4. Yandex Browser
    
    '5.1. Opera Presto
    
    '5.2. (Chromo) Opera
    
    '6. Edge
    '...

Finalize:
    ToggleWow64FSRedirection True
    
    AppendErrorLogCustom "FixO1Item - End"
    Exit Sub
    
ErrorHandler:
    If Err.Number = 70 And Not bSeenHostsFileAccessDeniedWarning Then
        'permission denied
        If Not bAutoLogSilent Then
            MsgBoxW Translate(303), vbExclamation
        End If
'        msgboxW "HiJackThis could not write the selected changes to your " & _
'               "hosts file. The probably cause is that some program is " & _
'               "denying access to it, or that your user account doesn't have " & _
'               "the rights to write to it.", vbExclamation
        bSeenHostsFileAccessDeniedWarning = True
    Else
        ErrorMsg Err, "modMain_FixO1Item", "sItem=", sItem
    End If
    Close #ff1, #ff2
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FlushDNS()
    On Error GoTo ErrorHandler:
    If GetServiceRunState("dnscache") <> SERVICE_RUNNING Then StartService "dnscache"

    'ipconfig.exe
    If Proc.ProcessRun(BuildPath(sSysNativeDir, Caes_Decode("jshvwqvv.xSB")), "/flushdns", , vbHide) Then
        Proc.WaitForTerminate , , , 15000
    End If
    
    RestartService "dnscache"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FlushDNS"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckO2Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO2Item - Begin"
    
    Dim hKey&, i&, sName$, sCLSID$, lpcName&, sFile$, sHit$, BHO_key$, result As SCAN_RESULT
    Dim sProgId$, sProgId_CLSID$, bSafe As Boolean
    
    Dim HEFixKey As clsHiveEnum
    Dim HEFixValue As clsHiveEnum
    
    Set HEFixKey = New clsHiveEnum
    Set HEFixValue = New clsHiveEnum
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HEFixKey.Init HE_HIVE_ALL
    HEFixValue.Init HE_HIVE_ALL
    
    'key to check
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\explorer\Browser Helper Objects"
    
    'keys to fix + \{CLSID} placeholder
    HEFixKey.AddKey "HKCR\CLSID", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer\Extension Compatibility", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\Stats", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\Settings", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer", "\ApprovedExtensionsMigration{CLSID}"
    
    'values to fix (value == {CLSID})
    HEFixValue.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID"
    HEFixValue.AddKey "Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration"
    
    Do While HE.MoveNext
   
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
    
            i = 0
            Do
                sCLSID = String$(MAX_KEYNAME, vbNullChar)
                lpcName = Len(sCLSID)
                If RegEnumKeyExW(hKey, i, StrPtr(sCLSID), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
                
                sCLSID = Left$(sCLSID, lstrlen(StrPtr(sCLSID)))
                
                If Len(sCLSID) <> 0 And Not StrBeginWith(sCLSID, "MSHist") Then
                    
                    BHO_key = HE.KeyAndHive & "\" & sCLSID
                    
                    If InStr(sCLSID, "}}") > 0 Then
                        'the new searchwww.com trick - use a double
                        '}} in the IE toolbar registration, reg the toolbar
                        'with only one } - IE ignores the double }}, but
                        'HT didn't. It does now!
                        sCLSID = Left$(sCLSID, Len(sCLSID) - 1)
                    End If
                    
                    'get filename from HKCR\CLSID\sName + BHO name
                    
                    'get bho name from BHO regkey
                    sName = Reg.GetString(0&, BHO_key, vbNullString, HE.Redirected)
                    If HE.SharedKey And Len(sName) = 0 Then
                        sName = Reg.GetString(0&, BHO_key, vbNullString, Not HE.Redirected)
                    End If
                    
                    GetFileByCLSID sCLSID, sFile, , HE.Redirected, HE.SharedKey
                    
                    sProgId = Reg.GetString(HKEY_CLASSES_ROOT, "Clsid\" & sCLSID & "\ProgID", vbNullString, HE.Redirected)
                    If Len(sProgId) = 0 And HE.SharedKey Then
                        sProgId = Reg.GetString(HKEY_CLASSES_ROOT, "Clsid\" & sCLSID & "\ProgID", vbNullString, Not HE.Redirected)
                    End If
                    
                    If Len(sProgId) <> 0 Then
                        'safety check
                        sProgId_CLSID = Reg.GetString(HKEY_CLASSES_ROOT, sProgId & "\Clsid", vbNullString, HE.Redirected)
                        If Len(sProgId_CLSID) = 0 And HE.SharedKey Then
                            sProgId_CLSID = Reg.GetString(HKEY_CLASSES_ROOT, sProgId & "\Clsid", vbNullString, Not HE.Redirected)
                        End If
                        
                        If sProgId_CLSID <> sCLSID Then
                            sProgId = vbNullString
                        End If
                    End If
                    
                    sFile = FormatFileMissing(sFile)
                    
                    bSafe = False
                    If bHideMicrosoft Then
                        If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                            If IsMicrosoftFile(sFile) Then bSafe = True
                        Else
                            If WhiteListed(sFile, "ie_to_edge_bho.dll", True) Then bSafe = True
                            If WhiteListed(sFile, "ie_to_edge_bho_64.dll", True) Then bSafe = True
                        End If
                    End If
                    
                    If Not bSafe Then
                        'get bho name from CLSID
                        If Len(sName) = 0 Then GetTitleByCLSID sCLSID, sName, HE.Redirected, HE.SharedKey
                        
                        SignVerifyJack sFile, result.SignResult
                        
                        sHit = BitPrefix("O2", HE) & _
                            " - " & HE.HiveNameAndSID & "\..\BHO: " & sName & " - " & sCLSID & " - " & sFile & FormatSign(result.SignResult)
                        
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                        
                        If Not IsOnIgnoreList(sHit) Then
                            
                            With result
                                .Section = "O2"
                                .HitLineW = sHit
                                
                                AddRegToFix .Reg, REMOVE_KEY, 0, BHO_key, , , IIf(HE.SharedKey, REG_REDIRECTION_BOTH, HE.Redirected)

                                If 0 <> Len(sProgId) Then
                                    AddRegToFix .Reg, REMOVE_KEY, HKCR, sProgId, , , IIf(HE.SharedKey, REG_REDIRECTION_BOTH, HE.Redirected)
                                End If
                                
                                HEFixKey.Repeat
                                Do While HEFixKey.MoveNext
                                    AddRegToFix .Reg, REMOVE_KEY, HEFixKey.Hive, Replace$(HEFixKey.Key, "{CLSID}", sCLSID), , , HEFixKey.Redirected
                                Loop
                                
                                HEFixValue.Repeat
                                Do While HEFixValue.MoveNext
                                    AddRegToFix .Reg, REMOVE_VALUE, HEFixValue.Hive, HEFixValue.Key, sCLSID, , HEFixValue.Redirected
                                Loop
                                
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile 'Or UNREG_DLL
                                
                                .CureType = REGISTRY_BASED Or FILE_BASED
                            End With
                        
                            AddToScanResults result
                        End If
                    End If
                End If
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    Set HEFixKey = Nothing
    Set HEFixValue = Nothing
    
    AppendErrorLogCustom "CheckO2Item - End"
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO2Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RequestCloseIE()
    
    Dim bIE_Exist As Boolean
    Static bForceWarningDisplayed As Boolean
    
    bIE_Exist = ProcessExist("iexplore.exe", True)
    
    If Not bShownBHOWarning And bIE_Exist Then
        MsgBoxW Translate(310), vbExclamation
'        msgboxW "HiJackThis is about to remove a " & _
'               "BHO and the corresponding file from " & _
'               "your system. Close all Internet " & _
'               "Explorer windows AND all Windows " & _
'               "Explorer windows before continuing for " & _
'               "the best chance of success.", vbExclamation
        bShownBHOWarning = True
    End If
    
    If bIE_Exist And Not bForceWarningDisplayed Then
        If MsgBox(Translate(311), vbExclamation) = vbYes Then
            'Internet Explorer still run. Would you like HJT close IE forcibly?
            'WARNING: current browser session will be lost!
            Proc.ProcessClose ProcessName:="iexplore.exe", Async:=False, TimeoutMs:=1000, SendCloseMsg:=True
        End If
        bForceWarningDisplayed = True
    End If
    
End Sub

Public Sub FixO2Item(sItem$, result As SCAN_RESULT)
    'O2 - Enumeration of existing MSIE BHO's
    'O2 - BHO: AcroIEHlprObj Class - {00000...000} - C:\PROGRAM FILES\ADOBE\ACROBAT 5.0\ACROBAT\ACTIVEX\ACROIEHELPER.OCX
    'O2 - BHO: ... (no file)
    'O2 - BHO: ... c:\bla.dll (file missing)
    
    On Error GoTo ErrorHandler:
    
    RequestCloseIE
    
    '//TODO: Add:
    'HKLM\SOFTWARE\WOW6432NODE\MICROSOFT\INTERNET EXPLORER\LOW RIGHTS\ELEVATIONPOLICY\{CLSID}
    'HKLM\SOFTWARE\CLASSES\APPID\{Name}
    'HKLM\SOFTWARE\CLASSES\APPID\{GUID}
    'HKLM\SOFTWARE\WOW6432NODE\CLASSES\APPID\{Name}
    'HKLM\SOFTWARE\WOW6432NODE\CLASSES\APPID\{GUID}
    'HKLM\SOFTWARE\CLASSES\INTERFACE\{GUID}
    'HKLM\SOFTWARE\CLASSES\TYPELIB\{GUID}
    
    'file should go first bacause it can use reg. info for its dll unregistration.
    
    FixIt result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO2Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO3Item()
    'HKLM\Software\Microsoft\Internet Explorer\Toolbar
    'HKLM\Software\Microsoft\Internet Explorer\Explorer Bars
  
    '//TODO:
    'Add handling of:
    'Locked value: http://www.tweaklibrary.com/windows/Software_Applications/Internet-Explorer/27/Unlock-the-Internet-Explorer-toolbars/11245/index.htm
    'Explorer, ShellBrowser and subkeys with ITBarLayout (ITBar7Layout, ITBar7Layout64) values: https://support.microsoft.com/en-us/help/555460
    'BackBitmapIE5 value (need ???): https://msdn.microsoft.com/en-us/library/aa753592(v=vs.85).aspx
    '
    'Detailed description: http://www.winblog.ru/admin/1147761976-ippon_170506_02.html
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO3Item - Begin"
    
    Dim hKey&, i&, sCLSID$, sName$, result As SCAN_RESULT
    Dim sFile$, sHit$, SearchwwwTrick As Boolean, sProgId$, sProgId_CLSID$
    Dim bSafe As Boolean
    
    Dim HEFixKey As clsHiveEnum
    Dim HEFixValue As clsHiveEnum
    
    Set HEFixKey = New clsHiveEnum
    Set HEFixValue = New clsHiveEnum
    
    Dim aKeys(1) As String
    Dim aDescr(1) As String
    
    'keys to check
    aKeys(0) = "Software\Microsoft\Internet Explorer\Toolbar"
    aKeys(1) = "Software\Microsoft\Internet Explorer\Explorer Bars"
    
    aDescr(0) = "Toolbar"
    aDescr(1) = "Explorer Bars"
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HEFixKey.Init HE_HIVE_ALL
    HEFixValue.Init HE_HIVE_ALL
    
    HE.AddKeys aKeys
    
    'keys to fix + placeholder
    HEFixKey.AddKey "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\Stats", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\Settings", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Windows\CurrentVersion\Ext\PreApproved", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration", "\{CLSID}"
    HEFixKey.AddKey "Software\Microsoft\Internet Explorer", "\ApprovedExtensionsMigration{CLSID}"
    
    'values to fix (value == {CLSID})
    HEFixValue.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID"
    HEFixValue.AddKey "Software\Microsoft\Internet Explorer\ApprovedExtensionsMigration"
    
    Do While HE.MoveNext
        
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then

            i = 0
            Do
                sCLSID = String$(MAX_VALUENAME, 0)
                Dim uData() As String
                ReDim uData(MAX_VALUENAME) As String
    
                'enumerate MSIE toolbars / Explorer Bars
                If RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&) <> 0 Then Exit Do
                sCLSID = TrimNull(sCLSID)
    
                If InStr(sCLSID, "}}") > 0 Then
                    'the new searchwww.com trick - use a double
                    '}} in the IE toolbar registration, reg the toolbar
                    'with only one } - IE ignores the double }}, but
                    'HJT didn't. It does now!
        
                    sCLSID = Left$(sCLSID, Len(sCLSID) - 1)
                    SearchwwwTrick = True
                Else
                    SearchwwwTrick = False
                End If
    
                'found one? then check corresponding HKCR key
                GetFileByCLSID sCLSID, sFile, , HE.Redirected, HE.SharedKey
    
                '   sCLSID <> "BrandBitmap" And _
                '   sCLSID <> "SmBrandBitmap" And _
                '   sCLSID <> "BackBitmap" And _
                '   sCLSID <> "BackBitmapIE5" And _
                '   sCLSID <> "OLE (Part 1 of 5)" And _
                '   sCLSID <> "OLE (Part 2 of 5)" And _
                '   sCLSID <> "OLE (Part 3 of 5)" And _
                '   sCLSID <> "OLE (Part 4 of 5)" And _
                '   sCLSID <> "OLE (Part 5 of 5)" Then
    
                sProgId = Reg.GetString(HKEY_CLASSES_ROOT, "Clsid\" & sCLSID & "\ProgID", vbNullString, HE.Redirected)
                If Len(sProgId) = 0 And HE.SharedKey Then
                    sProgId = Reg.GetString(HKEY_CLASSES_ROOT, "Clsid\" & sCLSID & "\ProgID", vbNullString, Not HE.Redirected)
                End If
    
                If 0 <> Len(sProgId) Then
                    'safe check
                    sProgId_CLSID = Reg.GetString(HKEY_CLASSES_ROOT, sProgId & "\Clsid", vbNullString, HE.Redirected)
                    If Len(sProgId_CLSID) = 0 And HE.SharedKey Then
                        sProgId_CLSID = Reg.GetString(HKEY_CLASSES_ROOT, sProgId & "\Clsid", vbNullString, Not HE.Redirected)
                    End If
                    
                    If sProgId_CLSID <> sCLSID Then
                        sProgId = vbNullString
                    End If
                End If
                
                sFile = FormatFileMissing(sFile)
                
                bSafe = False
                
                If bHideMicrosoft Then
                    If OSver.MajorMinor = 5 Then 'Win2k
                        If WhiteListed(sFile, sWinDir & "\system32\msdxm.ocx") Then bSafe = True
                    End If
                End If
                
                If InStr(sCLSID, "{") <> 0 And Not bSafe Then
    
                    GetTitleByCLSID sCLSID, sName, HE.Redirected, HE.SharedKey
                    
                    SignVerifyJack sFile, result.SignResult
                    
                    sHit = BitPrefix("O3", HE) & _
                        " - " & HE.HiveNameAndSID & "\..\" & aDescr(HE.KeyIndex) & ": " & sName & " - " & sCLSID & " - " & sFile & FormatSign(result.SignResult)
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    
                    If Not IsOnIgnoreList(sHit) Then
                        
                        With result
                            .Section = "O3"
                            .HitLineW = sHit
                            
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\WebBrowser", sCLSID, , HE.Redirected
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\ShellBrowser", sCLSID, , HE.Redirected
                            
                            If 0 <> Len(sProgId) Then
                                AddRegToFix .Reg, REMOVE_VALUE, 0, "HKCR\" & sProgId, sCLSID, , IIf(HE.SharedKey, REG_REDIRECTION_BOTH, HE.Redirected)
                            End If
                            
                            HEFixKey.Repeat
                            Do While HEFixKey.MoveNext
                                AddRegToFix .Reg, REMOVE_KEY, HEFixKey.Hive, Replace$(HEFixKey.Key, "{CLSID}", sCLSID), , , HEFixKey.Redirected
                            Loop
                            
                            HEFixValue.Repeat
                            Do While HEFixValue.MoveNext
                                AddRegToFix .Reg, REMOVE_VALUE, HEFixValue.Hive, HEFixValue.Key, sCLSID, , HEFixValue.Redirected
                            Loop
                            
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile 'Or UNREG_DLL
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    Set HEFixKey = Nothing
    Set HEFixValue = Nothing
    
    AppendErrorLogCustom "CheckO3Item - End"
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO3Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO3Item(sItem$, result As SCAN_RESULT)
    FixIt result
End Sub


'returns array of SID strings, except of current user
Sub GetUserNamesAndSids(aSID() As String, aUser() As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetUserNamesAndSids - Begin"
    
    'get all users' SID and map it to the corresponding username
    'not all users visible in User Accounts screen have a SID in HKU hive though,
    'they get it when logged in

    Dim CurUserName$, i&, sUsername$, aTmpSID() As String, aTmpUser() As String

    CurUserName = OSver.UserName
    
    aTmpSID = SplitSafe(Reg.EnumSubKeys(HKEY_USERS, vbNullString), "|")
    ReDim aTmpUser(UBound(aTmpSID))
    For i = 0 To UBound(aTmpSID)
        If aTmpSID(i) Like "S-#-#-#*" And Not StrEndWith(aTmpSID(i), "_Classes") Then
            sUsername = MapSIDToUsername(aTmpSID(i))
            If 0 = Len(sUsername) Then sUsername = "?"
            If StrComp(sUsername, CurUserName, 1) <> 0 Then
                aTmpUser(i) = sUsername
            Else
                'filter current user key with HKCU
                aTmpSID(i) = vbNullString
                aTmpUser(i) = vbNullString
            End If
        Else
            aTmpSID(i) = vbNullString
            aTmpUser(i) = vbNullString
        End If
    Next i
    
    CompressArray aTmpSID
    CompressArray aTmpUser
    
    aSID = aTmpSID
    aUser = aTmpUser
    
    AppendErrorLogCustom "GetUserNamesAndSids - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.GetUserNamesAndSids"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_RegRuns()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4_RegRuns - Begin"
    
    Const BAT_LENGTH_LIMIT As Long = 300&
    
    Dim aRegRuns() As String, aDes() As String, result As SCAN_RESULT
    Dim i&, j&, sKey$, sData$, sHit$, sAlias$, sParam As String, sHash$, aValue() As String
    Dim bData() As Byte, isDisabledWin8 As Boolean, isDisabledWinXP As Boolean, flagDisabled As Long, sKeyDisable As String
    Dim sFile$, sArgs$, sUser$, bSafe As Boolean, aLines() As String
    Dim aData() As String, bDisabled As Boolean, bMicrosoft As Boolean, bMissing As Boolean
    Dim sOrigLine As String
    
    ReDim aRegRuns(1 To 9)
    ReDim aDes(UBound(aRegRuns))
    
    aRegRuns(1) = "Software\Microsoft\Windows\CurrentVersion\Run"
    aDes(1) = "Run"
    
    aRegRuns(2) = "Software\Microsoft\Windows\CurrentVersion\RunServices"
    aDes(2) = "RunServices"
    
    aRegRuns(3) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
    aDes(3) = "RunOnce"
    
    aRegRuns(4) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
    aDes(4) = "RunServicesOnce"
    
    aRegRuns(5) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
    aDes(5) = "Policies\Explorer\Run"
    
    aRegRuns(6) = "Software\Microsoft\Windows\CurrentVersion\Run-"
    aDes(6) = "Run-"

    aRegRuns(7) = "Software\Microsoft\Windows\CurrentVersion\RunServices-"
    aDes(7) = "RunServices-"

    aRegRuns(8) = "Software\Microsoft\Windows\CurrentVersion\RunOnce-"
    aDes(8) = "RunOnce-"

    aRegRuns(9) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce-"
    aDes(9) = "RunServicesOnce-"
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL
    HE.AddKeys aRegRuns
    
    Do While HE.MoveNext
        
        For i = 1 To Reg.NtEnumValuesToArray(HE.Hive, HE.Key, aValue(), HE.Redirected)
            
            isDisabledWin8 = False
            
            isDisabledWinXP = (Right$(HE.Key, 1) = "-")    ' Run- e.t.c.
            
            sData = Reg.NtGetData(HE.Hive, HE.Key, aValue(i), HE.Redirected)
            
            If OSver.IsWindows8OrGreater Then
                
                If StrComp(HE.Key, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 1) = 0 Then
                    
                    'Param. name is always "Run" on x32 bit. OS.
                    sKeyDisable = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And HE.Redirected, "Run32", "Run")
                    
                    If Reg.ValueExists(HE.Hive, sKeyDisable, aValue(i)) Then
                        
                        ReDim bData(0)
                        bData() = Reg.GetBinary(HE.Hive, sKeyDisable, aValue(i))
            
                        If UBoundSafe(bData) >= 11 Then
                            
                            GetMem4 ByVal VarPtr(bData(0)), flagDisabled
                            
                            If CBool(flagDisabled And 1) Then isDisabledWin8 = True
                        End If
                    End If
                End If
            End If
            
            If Len(sData) <> 0 And Not isDisabledWin8 Then
                
                'Example:
                '"O4 - HKLM\..\Run: "
                '"O4 - HKU\S-1-5-19\..\Run: "
                sAlias = BitPrefix("O4", HE) & _
                    " - " & IIf(isDisabledWinXP, "(disabled) ", vbNullString) & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": "
                
                sHit = sAlias & "[" & aValue(i) & "] = "
                
                If HE.IsSidUser Or HE.IsSidService Then
                    sUser = " (User '" & HE.UserName & "')"
                Else
                    sUser = vbNullString
                End If
                
                SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                
                sFile = FormatFileMissing(sFile, sArgs)
                
                bMicrosoft = False
                If InStr(1, sFile, "OneDrive", 1) <> 0 Then bMicrosoft = IsMicrosoftFile(sFile)
                
                bSafe = False
                
                If Not bIgnoreAllWhitelists And bHideMicrosoft Then
                    
                    '//TODO: narrow down to services' SID only: S-1-5-19 + S-1-5-20 + 'UpdatusUser' (NVIDIA)
                    
                    'Note: For services only
                    If StrComp(sFile, PF_64 & "\Windows Sidebar\Sidebar.exe", 1) = 0 And sArgs = "/autoRun" Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinDir & "\System32\mctadmin.exe", 1) = 0 And Len(sArgs) = 0 Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinSysDirWow64 & "\OneDriveSetup.exe", 1) = 0 And sArgs = "/thfirstsetup" Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinDir & "\system32\SecurityHealthSystray.exe", 1) = 0 And Len(sArgs) = 0 Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinDir & "\AzureArcSetup\Systray\AzureArcSysTray.exe", 1) = 0 And Len(sArgs) = 0 Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                    
'                    If OSver.IsWindows2000 Then
'                        If WhiteListed(sFile, sWinDir & "\system32\internat.exe") And Len(sArgs) = 0 Then bSafe = True
'                    End If
'                    If OSver.IsWindows2000 And OSver.IsServer Then
'                        If aDes(HE.KeyIndex) = "RunOnce" Then
'                            If WhiteListed(sFile, PF_32 & "\Internet Explorer\Connection Wizard\icwconn1.exe") And sArgs = "/desktop" Then bSafe = True
'                        End If
'                    End If
                    
                    If OSver.MajorMinor <= 6.1 Then 'Win2k-Win7
                        If WhiteListed(sFile, sWinDir & "\system32\CTFMON.EXE") And Len(sArgs) = 0 Then bSafe = True
                    End If

                    If OSver.IsWindowsVista Then 'Vista/2008
                        If WhiteListed(sFile, sWinDir & "\system32\rundll32.exe") And sArgs = "oobefldr.dll,ShowWelcomeCenter" Then
                            If IsMicrosoftFile(sWinDir & "\system32\oobefldr.dll") Then bSafe = True
                        End If
                        If WhiteListed(sFile, PF_64 & "\Windows Sidebar\sidebar.exe") And (sArgs = "/autoRun" Or sArgs = "/detectMem") Then bSafe = True
                        '\MSASCui.exe
                        If WhiteListed(sFile, PF_64 & "\" & STR_CONST.WINDOWS_DEFENDER & Caes_Decode("]PXH\NHx.xSB")) And sArgs = "-hide" Then bSafe = True
                    End If
                    
                    If OSver.IsWindows8OrGreater Then
                        If WhiteListed(sFile, PF_64 & "\" & STR_CONST.WINDOWS_DEFENDER & "\MSASCuiL.exe") And Len(sArgs) = 0 Then bSafe = True
                    End If

                End If
                
                SignVerifyJack sFile, result.SignResult
                
                sHit = sHit & ConcatFileArg(sFile, sArgs) & sUser & FormatSign(result.SignResult)
                
                If (Not bSafe) Or (Not bHideMicrosoft) Then
                
                    If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
                
                    If (Not IsOnIgnoreList(sHit)) Then
                        
                        With result
                            .Section = "O4"
                            .HitLineW = sHit
                            .Alias = sAlias
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i), , HE.Redirected
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                            
                            If Not isDisabledWinXP Then
                                AddProcessToFix .Process, FREEZE_OR_KILL_PROCESS, sFile
                            End If
                            .CureType = REGISTRY_BASED Or FILE_BASED Or PROCESS_BASED
                        End With
                        
                        AddToScanResults result
                    End If
                End If
            End If
        Next
    Loop
    
    'Certain param based checkings
    
    Dim aRegKey() As String
    Dim aRegParam() As String
    Dim aDefData() As String
    ReDim aRegKey(1 To 6) As String                   'key
    ReDim aRegParam(1 To UBound(aRegKey)) As String   'param
    ReDim aDefData(1 To UBound(aRegKey)) As String    'data
    
    aRegKey(1) = "Software\Microsoft\Command Processor" 'HKLM + HKU
    aRegParam(1) = "Autorun"
    aDefData(1) = vbNullString
    
    aRegKey(2) = "HKLM\SYSTEM\CurrentControlSet\Control\BootVerificationProgram"
    aRegParam(2) = "ImagePath"
    aDefData(2) = vbNullString
    
    aRegKey(3) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(3) = "BootExecute"
    
    aRegKey(4) = "HKLM\SYSTEM\CurrentControlSet\Control\SafeBoot"
    aRegParam(4) = "AlternateShell"
    aDefData(4) = "cmd.exe"
    
    aRegKey(5) = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager"
    aRegParam(5) = "SetupExecute"
    aDefData(5) = vbNullString
    '
    'see: https://guyrleech.wordpress.com/2014/07/16/reasons-for-reboots-part-2-2/
    
    If OSver.IsWindows8OrGreater Then
        aRegKey(6) = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager"
        aRegParam(6) = "BootShell"
        aDefData(6) = "%SystemRoot%\system32\bootim.exe"
    End If
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        
        sParam = aRegParam(HE.KeyIndex)
        
        sData = Reg.GetData(HE.Hive, HE.Key, sParam, HE.Redirected)
        
        aData = SplitSafe(sData, vbNullChar) 'if MULTI_SZ (BootExecute)
        
        ArrayRemoveEmptyItems aData
        
        For i = 0 To UBound(aData)
        
            bSafe = False
        
            sData = aData(i)
            sOrigLine = sData
        
            If StrComp(sParam, "BootExecute", 1) = 0 Then
                If i = 0 Then
                    If StrBeginWith(sData, "autocheck ") Then 'remove autocheck, because it is not a real filename
                        sData = mid$(sData, Len("autocheck ") + 1)
                    End If
                End If
                
                If bHideMicrosoft Then
                    If OSver.MajorMinor = 5 Then 'Win2k
                        If StrComp(sData, "autochk *", 1) = 0 Or StrComp(sData, "DfsInit", 1) = 0 Then bSafe = True
                    ElseIf OSver.MajorMinor >= 6.2 And OSver.IsServer Then '2012 Server, 2012 Server R2 (2016 too ?)
                        If StrComp(sData, "autochk /q /v *", 1) = 0 Then bSafe = True
                        If StrComp(sData, BuildPath(sWinSysDir, "autochk.exe") & " /q /v *", 1) = 0 Then bSafe = True
                    Else
                        If StrComp(sData, "autochk *", 1) = 0 Then bSafe = True
                    End If
                End If
            Else
                If sData = EnvironW(aDefData(HE.KeyIndex)) Then bSafe = True
            End If
            
            bDisabled = False
            If StrComp(sParam, "AlternateShell", 1) = 0 Then
                If 1 <> Reg.GetDword(HKEY_LOCAL_MACHINE, HE.Key & "\Options", "UseAlternateShell") Then
                    bDisabled = True
                End If
            End If
            
            If Not bSafe Or bIgnoreAllWhitelists Or Not bHideMicrosoft Then
                
                'HKLM\..\Command Processor: [Autorun] =
                sAlias = BitPrefix("O4", HE) & " - " & HE.HiveNameAndSID & "\..\" & GetFileName(HE.Key) & ": " & _
                    "[" & sParam & "] = "
                
                SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                
                sFile = FormatFileMissing(sFile)
                
                SignVerifyJack sFile, result.SignResult
                
                sHit = sAlias & ConcatFileArg(sFile, sArgs) & FormatSign(result.SignResult)
                
                If bDisabled Then sHit = sHit & " (disabled)"
                
                If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
                
                If Not IsOnIgnoreList(sHit) Then
                    
                    With result
                        .Section = "O4"
                        .HitLineW = sHit
                        .Alias = sAlias
                        If StrComp(sParam, "BootExecute", 1) = 0 Then
                            
                            AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE, _
                                HE.Hive, HE.Key, sParam, , HE.Redirected, REG_RESTORE_MULTI_SZ, _
                                sOrigLine, vbNullString, vbNullChar
                            
                            AddRegToFix .Reg, APPEND_VALUE_NO_DOUBLE, HE.Hive, HE.Key, sParam, _
                                "autocheck autochk *", HE.Redirected, REG_RESTORE_MULTI_SZ
                            
                            If OSver.MajorMinor = 5 Then
                                AddRegToFix .Reg, APPEND_VALUE_NO_DOUBLE, HE.Hive, HE.Key, sParam, _
                                    "DfsInit", HE.Redirected, REG_RESTORE_MULTI_SZ
                            End If
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                            
                        ElseIf StrComp(sParam, "SetupExecute", 1) = 0 Then
                            
                            AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sParam, vbNullString, HE.Redirected
                            
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                            AddJumpFiles .Jump, JUMP_FILE, ExtractFilesFromCommandLine(sData)

                            .CureType = REGISTRY_BASED Or FILE_BASED
                        Else
                            If Len(aDefData(HE.KeyIndex)) <> 0 Then
                                AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sParam, aDefData(HE.KeyIndex), HE.Redirected
                            Else
                                AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                            End If
                                
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs

                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End If
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    
  If bAdditional And Not bStartupScan Then
  
    ReDim aRegKey(1 To 2) As String                   'key
    ReDim aRegParam(1 To UBound(aRegKey)) As String   'param
    ReDim aDes(1 To UBound(aRegKey)) As String        'description
    
    'https://technet.microsoft.com/en-us/library/cc960241.aspx
    'PendingFileRenameOperations
    'Shared
    
    aRegKey(1) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(1) = "PendingFileRenameOperations"
    aDes(1) = "Session Manager"
    
    aRegKey(2) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(2) = "PendingFileRenameOperations2"
    aDes(2) = "Session Manager"
    
    HE.Init HE_HIVE_HKLM, , HE_REDIR_NO_WOW
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        
        sParam = aRegParam(HE.KeyIndex)
    
        sData = Reg.GetData(HE.Hive, HE.Key, sParam, HE.Redirected)
        
        If Len(sData) <> 0 Then
        
          'converting MULTI_SZ to [1] -> [2], [3] -> [4] ...
          aLines = SplitSafe(sData, vbNullChar)
        
          For j = 0 To UBound(aLines) Step 2
            bSafe = False
            sFile = PathNormalize(aLines(j))
            If j + 1 <= UBound(aLines) Then
                If Len(aLines(j + 1)) = 0 Then
                    sArgs = "-> DELETE"
                    If Not bIgnoreAllWhitelists Then bSafe = True
                Else
                    sArgs = "-> " & PathNormalize(aLines(j + 1))
                End If
            End If
            
            'HKLM\..\FileRenameOperations:
            sAlias = BitPrefix("O4", HE) & " - " & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": " & _
                "[" & sParam & "] = "
            
            sFile = FormatFileMissing(sFile, , bMissing)
            
            sHit = sAlias & ConcatFileArg(sFile, sArgs)
            
            If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
            
            If bIgnoreAllWhitelists Or ((Not bSafe) And (Not IsOnIgnoreList(sHit)) And Not bMissing) Then
                With result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                    
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults result
            End If
          Next
        End If
    Loop
    
  End If
    
    Dim aFiles() As String
    ReDim aFiles(5)
    
    'Win 9x
    aFiles(0) = BuildPath(sWinSysDir, "BatInit.bat")
    aFiles(1) = BuildPath(sWinDir, "WinStart.bat")
    aFiles(2) = BuildPath(sWinDir, "DosStart.bat")
    aFiles(3) = BuildPath(SysDisk, "AutoExec.bat")
    'Win NT
    aFiles(4) = BuildPath(sWinSysDir, "AutoExec.nt")
    aFiles(5) = BuildPath(sWinSysDir, "Config.nt")
    
    For i = 0 To UBound(aFiles)
        sFile = aFiles(i)
        If FileExists(sFile) Then
            
            sData = ReadFileContents(sFile, False)
            
            If Len(sData) <> 0 Then
                
                sData = Replace$(sData, vbCr, vbNullString)
                aData = Split(sData, vbLf)
                
                'exclude comments
                For j = 0 To UBound(aData)
                    If (StrBeginWith(aData(j), "REM") And Not bIgnoreAllWhitelists) Then
                        aData(j) = vbNullString
                    ElseIf (StrBeginWith(aData(j), "::") And Not bIgnoreAllWhitelists) Then
                        aData(j) = vbNullString
                    ElseIf bHideMicrosoft Then
                        'check whitelist
                        If StrEndWith(sFile, "AutoExec.nt") Then
                            If StrComp(aData(j), "@echo off", 1) = 0 Then
                                aData(j) = vbNullString
                            ElseIf StrComp(aData(j), "lh %SystemRoot%\system32\mscdexnt.exe", 1) = 0 Then
                                aData(j) = vbNullString
                            ElseIf StrComp(aData(j), "lh %SystemRoot%\system32\redir", 1) = 0 Then
                                aData(j) = vbNullString
                            ElseIf StrComp(aData(j), "lh %SystemRoot%\system32\dosx", 1) = 0 Then
                                aData(j) = vbNullString
                            ElseIf StrComp(aData(j), "SET BLASTER=A220 I5 D1 P330 T3", 1) = 0 Then
                                aData(j) = vbNullString
                            End If
                        ElseIf StrEndWith(sFile, "Config.nt") Then
                            If StrComp(aData(j), "dos=high, umb", 1) = 0 Then
                                aData(j) = vbNullString
                            ElseIf StrComp(aData(j), "device=%SystemRoot%\system32\himem.sys", 1) = 0 Then
                                aData(j) = vbNullString
                            ElseIf StrComp(aData(j), "Files=40", 1) = 0 Then
                                aData(j) = vbNullString
                            End If
                        End If
                    End If
                    
                    If 0 <> Len(aData(j)) Then
                        If Len(aData(j)) > BAT_LENGTH_LIMIT Then
                            aData(j) = Left$(aData(j), BAT_LENGTH_LIMIT) & " ... (" & Len(aData(j)) - BAT_LENGTH_LIMIT & " more characters)"
                        End If
                        If i < 4 Then
                            sAlias = "O4 - Win9x BAT: "
                        Else
                            sAlias = "O4 - WinNT BAT: "
                        End If
                        sHit = sAlias & sFile & " => " & EscapeSpecialChars(aData(j)) & IIf(Len(aData(j)) = 0, " (0 bytes)", vbNullString)
                        If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O4"
                                .HitLineW = sHit
                                .Alias = sAlias
                                AddFileToFix .File, REMOVE_FILE, sFile
                                .CureType = FILE_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    'RunOnceEx, RunServicesOnceEx
    'https://support.microsoft.com/en-us/kb/310593
    'http://www.oszone.net/2762
    '" DllFileName | FunctionName | CommandLineArguments "
    Dim aSubKey() As String

    ReDim aRegKey(1 To 2) As String                   'key
    ReDim aDes(1 To UBound(aRegKey)) As String        'description
    Dim pos As Long
    
    aRegKey(1) = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
    aDes(1) = "RunOnceEx"
    aRegKey(2) = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnceEx"
    aDes(2) = "RunServicesOnceEx"
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        If Reg.KeyHasSubKeys(HE.Hive, HE.Key, HE.Redirected) Then
            
            For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey(), HE.Redirected, , False)
                
                For j = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key & "\" & aSubKey(i), aValue(), HE.Redirected)
                    
                    sData = Reg.GetString(HE.Hive, HE.Key & "\" & aSubKey(i), aValue(j), HE.Redirected)
                    
                    'e.g. C:\PROGRA~2\COMMON~1\MICROS~1\Repostry\REPCDLG.OCX|DllRegisterServer
                    pos = InStr(sData, "|")
                    If pos <> 0 Then
                        sFile = Left$(sData, pos - 1)
                        sArgs = mid$(sData, pos)
                    Else
                        sFile = sData
                        sArgs = vbNullString
                    End If
                    
                    sFile = FormatFileMissing(sFile)
                    
                    SignVerifyJack sFile, result.SignResult
                    
                    sAlias = "O4 - " & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": "
                    sHit = sAlias & aSubKey(i) & " [" & aValue(j) & "] = " & ConcatFileArg(sFile, sArgs) & FormatSign(result.SignResult)
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O4"
                            .HitLineW = sHit
                            .Alias = sAlias
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\" & aSubKey(i), aValue(j), , HE.Redirected
                            AddRegToFix .Reg, REMOVE_KEY_IF_NO_VALUES, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                Next
            Next
        End If
    Loop
    
  If bAdditional Then
    
    'Autorun.inf
    'http://journeyintoir.blogspot.com/2011/01/autoplay-and-autorun-exploit-artifacts.html
    Dim aDrives() As String
    Dim sAutorun As String
    Dim aVerb() As String
    Dim bOnce As Boolean

    aVerb = Split("open|shellexecute|shell\open\command|shell\explore\command", "|")

    ' Mapping scheme for "inf. verb" -> to "registry" :
    '
    ' icon                  -> _Autorun\Defaulticon
    ' open                  -> shell\AutoRun\command
    ' shellexecute          -> shell\AutoRun\command
    ' shell\open\command    -> shell\open\command
    ' shell\explore\command -> shell\explore\command

    aDrives = GetDrives(DRIVE_BIT_FIXED Or DRIVE_BIT_REMOVABLE)

    For i = 1 To UBound(aDrives)
        sAutorun = BuildPath(aDrives(i), "autorun.inf")
        If FileExists(sAutorun) Then

            bOnce = False

            For j = 0 To UBound(aVerb)

                sFile = vbNullString
                sArgs = vbNullString
                sData = ReadIniA(sAutorun, "autorun", aVerb(j))

                If Len(sData) <> 0 Then
                    SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=False
                    sFile = FormatFileMissing(sFile)
                    SignVerifyJack sFile, result.SignResult

                    sHit = "O4 - Autorun.inf: " & sAutorun & " - " & aVerb(j) & " - " & ConcatFileArg(sFile, sArgs) & FormatSign(result.SignResult)

                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O4"
                            .HitLineW = sHit
                            AddFileToFix .File, REMOVE_FILE, sAutorun, sArgs
                            .CureType = FILE_BASED
                        End With
                        AddToScanResults result
                    End If

                    bOnce = True
                End If
            Next

            'if unknown data is inside autorun.inf
            If Not bOnce Then

                sHit = "O4 - Autorun.inf: " & sAutorun & " - " & "(unknown target)"

                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O4"
                        .HitLineW = sHit
                        AddFileToFix .File, REMOVE_FILE, sAutorun
                        .CureType = FILE_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        End If
    Next
    
  End If
  
  If bAdditional Then

    'MountPoints2
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"

    aVerb = Split("shell\AutoRun\command|shell\open\command|shell\explore\command", "|")

    Do While HE.MoveNext
        
        For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey, HE.Redirected)
            For j = 0 To UBound(aVerb)
                sKey = HE.Key & "\" & aSubKey(i) & "\" & aVerb(j)

                If Reg.KeyExists(HE.Hive, sKey, HE.Redirected) Then

                    sData = Reg.GetString(HE.Hive, sKey, vbNullString, HE.Redirected)

                    SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                    sFile = FormatFileMissing(sFile)
                    SignVerifyJack sFile, result.SignResult

                    sHit = BitPrefix("O4", HE) & " - MountPoints2: " & HE.HiveNameAndSID & "\..\" & aSubKey(i) & "\" & aVerb(j) & ": (default) = " & ConcatFileArg(sFile, sArgs) & FormatSign(result.SignResult)

                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O4"
                            .HitLineW = sHit
                            'remove MountPoints2\{CLSID}
                            'or
                            'remove MountPoints2\Letter
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                            AddJumpFile .Jump, JUMP_FILE, sFile
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        Next
    Loop
    
  End If
    
    'ScreenSaver
    
    HE.Init HE_HIVE_HKCU Or HE_HIVE_HKU, HE_SID_USER Or HE_SID_NO_VIRTUAL, HE_REDIR_NO_WOW
    HE.AddKey "Control Panel\Desktop"
    
    Do While HE.MoveNext
    
      sData = Reg.GetString(HE.Hive, HE.Key, "SCRNSAVE.EXE")
      If 0 <> Len(sData) Then
        bSafe = True
        
        SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
        sFile = FormatFileMissing(sFile)
        
        If FileMissing(sFile) Then
            '(���)
            If (sData <> STR_CONST.RU_NO And sData <> "(None)" Or bIgnoreAllWhitelists) Then
                bSafe = False
            End If
        Else
            If IsMicrosoftFile(sFile) Then
                If IsLoLBin(sFile) Then bSafe = False
            Else
                bSafe = False
            End If
        End If
        
        If Not bSafe Then
            SignVerifyJack sFile, result.SignResult
            
            sHit = "O4 - " & HE.HiveNameAndSID & "\Control Panel\Desktop: [SCRNSAVE.EXE] = " & sFile & FormatSign(result.SignResult)
        
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O4"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "SCRNSAVE.EXE"
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults result
            End If
        End If
      End If
    Loop
    
    AppendErrorLogCustom "CheckO4_RegRuns - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.CheckO4_RegRuns"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_MSConfig(aHives() As String, aUserOfHive() As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4_MSConfig - Begin"
    
    'HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupreg
    'HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupfolder
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32
    '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder -> checked in CheckO4_AutostartFolder()
    
    Dim sHive$, i&, j&, sAlias$, sHash$, result As SCAN_RESULT
    Dim aSubKey$(), sDay$, sMonth$, sYear$, sKey$, sFile$, sTime$, sHit$, SourceHive$, sArgs$, sUser$, sDate$, sAddition$
    Dim Values$(), bData() As Byte, flagDisabled As Long, dDate As Date, UseWow As Variant, Wow6432Redir As Boolean, sTarget$, sData$
    
    Const sDateEpoch As String = "1601/01/01"
    
    If OSver.MajorMinor >= 6.2 Then ' Win 8+
    
        For i = 0 To UBound(aHives) 'HKLM, HKCU, HKU\SID()

            sHive = aHives(i)
            
            For Each UseWow In Array(False, True)
    
                Wow6432Redir = UseWow
  
                If (bIsWin32 And Wow6432Redir) _
                  Or bIsWin64 And Wow6432Redir And (sHive = "HKCU" Or StrBeginWith(sHive, "HKU\")) Then
                    Exit For
                End If
            
                
                For j = 1 To Reg.EnumValuesToArray(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values())
            
                    flagDisabled = 2
                    ReDim bData(0)
                    
                    bData() = Reg.GetBinary(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values(j))
                    
                    'undoc. flag. Seen:
                    '0x02000000 - enabled
                    '0x06000000 - enabled
                    '0x03000000 - disabled
                    '0x07000000 - disabled
                    '---
                    'looks like, flag 1 - is disabled
                    If UBoundSafe(bData) >= 11 Then
                        GetMem4 ByVal VarPtr(bData(0)), flagDisabled
                    End If
                    
                    If AryItems(bData) And CBool(flagDisabled And 1) Then   'is Disabled ?
                    
                        dDate = ConvertFileTimeToLocalDate(VarPtr(bData(4)))
                        
                        If Reg.ValueExists(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), Wow6432Redir) Then
                        
                            sData = Reg.GetString(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), Wow6432Redir)
                        
                            'if you change it, change fix appropriate !!!
                            sAlias = "O4 - " & sHive & "\..\StartupApproved\" & IIf(bIsWin64 And Wow6432Redir, "Run32", "Run") & ": "
                            
                            sHit = sAlias & "[" & Values(j) & "] "
                            
                            sUser = vbNullString
                            If Len(aUserOfHive(i)) <> 0 And StrBeginWith(sHive, "HKU\") Then
                                If (sHive <> "HKU\S-1-5-18" And _
                                    sHive <> "HKU\S-1-5-19" And _
                                    sHive <> "HKU\S-1-5-20") Then sUser = " (User '" & aUserOfHive(i) & "')"
                            End If
                            
                            SplitIntoPathAndArgs sData, sFile, sArgs, True
                            
                            sFile = FormatFileMissing(sFile)
                            
                            sHit = sHit & "= " & ConcatFileArg(sFile, sArgs) & sUser
                            
                            sDate = Format$(dDate, "yyyy\/mm\/dd")
                            If sDate <> sDateEpoch Then sHit = sHit & " (" & sDate & ")"
                            
                            SignVerifyJack sFile, result.SignResult
                            sHit = sHit & FormatSign(result.SignResult)
                            
                            If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
                            
                            If Not IsOnIgnoreList(sHit) Then
                                With result
                                    .Section = "O4"
                                    .HitLineW = sHit
                                    .Alias = sAlias
                                    AddRegToFix .Reg, REMOVE_VALUE, 0, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values(j), , False
                                    AddRegToFix .Reg, REMOVE_VALUE, 0, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), , CLng(Wow6432Redir)
                                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                                    .CureType = REGISTRY_BASED Or FILE_BASED
                                End With
                                AddToScanResults result
                            End If
                        End If
                    End If
                Next
            Next
        Next
        
    Else
    
        sHive = "HKLM"
        sKey = sHive & "\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupreg"
        
        For i = 1 To Reg.EnumSubKeysToArray(0&, sKey, aSubKey())
        
            sData = Reg.GetData(0&, sKey & "\" & aSubKey(i), "command")
            
            sYear = Reg.GetData(0&, sKey & "\" & aSubKey(i), "YEAR")
            sMonth = Right$("0" & Reg.GetData(0&, sKey & "\" & aSubKey(i), "MONTH"), 2)
            sDay = Right$("0" & Reg.GetData(0&, sKey & "\" & aSubKey(i), "DAY"), 2)
            
            If Val(sYear) = 0 Or Val(sMonth) = 0 Or Val(sDay) = 0 Then
                sTime = Format$(Reg.GetKeyTime(0&, sKey & "\" & aSubKey(i)), "yyyy\/mm\/dd")
            Else
                sTime = sYear & "/" & sMonth & "/" & sDay
            End If
            
            SourceHive = Reg.GetData(0&, sKey & "\" & aSubKey(i), "hkey")
            If SourceHive <> "HKLM" And SourceHive <> "HKCU" Then SourceHive = vbNullString
            
            'O4 - MSConfig\startupreg: [RtHDVCpl] C:\Program Files\Realtek\Audio\HDA\RAVCpl64.exe -s (HKLM) (2016/10/13)
            sAlias = "O4 - MSConfig\startupreg: "
            
            sHit = sAlias & aSubKey(i) & " [command] = "
            
            SplitIntoPathAndArgs sData, sFile, sArgs, True
            
            sFile = FormatFileMissing(sFile)
            sAddition = sArgs
            
            If Len(SourceHive) <> 0 Then sAddition = sAddition & IIf(Len(sArgs) = 0, vbNullString, " ") & "(" & SourceHive & ")"
            
            SignVerifyJack sFile, result.SignResult
            
            sHit = sHit & ConcatFileArg(sFile, sAddition & " (" & sTime & ")") & FormatSign(result.SignResult)
            
            If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
           
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_KEY, 0, sKey & "\" & aSubKey(i), , , REG_NOTREDIRECTED
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        Next
        
        'Startup folder items
        
        sKey = "HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupfolder"
        
        
        For i = 1 To Reg.EnumSubKeysToArray(0&, sKey, aSubKey())
        
            sArgs = vbNullString
        
            sFile = Reg.GetData(0&, sKey & "\" & aSubKey(i), "backup")
            
            sTime = Format$(Reg.GetKeyTime(0&, sKey & "\" & aSubKey(i)), "yyyy\/mm\/dd")
        
            sAlias = "O4 - MSConfig\startupfolder: "    'if you change it, change fix appropriate !!!
            
            If UCase$(GetExtensionName(aSubKey(i))) = ".LNK" Then
                'expand LNK, like:
                'C:^ProgramData^Microsoft^Windows^Start Menu^Programs^Startup^GIGABYTE OC_GURU.lnk - C:\Windows\pss\GIGABYTE OC_GURU.lnk.CommonStartup
            
                If FileExists(sFile) Then
                    sTarget = GetFileFromShortcut(sFile, sArgs, True)
                End If
            End If
            
            If 0 <> Len(sTarget) Then
                SignVerifyJack sTarget, result.SignResult
                sHit = sAlias & aSubKey(i) & " [backup] => " & sTarget & IIf(Len(sArgs) <> 0, " " & sArgs, vbNullString) & " (" & sTime & ")" & IIf(Not FileExists(sTarget), " " & STR_FILE_MISSING, vbNullString) & FormatSign(result.SignResult)
            Else
                SignVerifyJack sFile, result.SignResult
                sHit = sAlias & aSubKey(i) & " [backup] = " & sFile & " (" & sTime & ")" & IIf(Len(sFile) = 0, " (no file)", IIf(Not FileExists(sFile), " " & STR_FILE_MISSING, vbNullString)) & FormatSign(result.SignResult)
            End If
            
            If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_KEY, 0&, sKey & "\" & aSubKey(i), , , REG_NOTREDIRECTED
                    AddFileToFix .File, REMOVE_FILE, sFile 'removing backup (.pss)
                    AddJumpFile .Jump, JUMP_FILE, sTarget, sArgs
                    .CureType = FILE_BASED Or REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        Next
    End If
    
    AppendErrorLogCustom "CheckO4_MSConfig - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.CheckO4_MSConfig"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_AutostartFolder(aSID() As String, aUserOfHive() As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4_AutostartFolder - Begin"
    
    Const dEpoch As Date = #1/1/1601#
    
    Dim aRegKeys() As String, aParams() As String, aDes() As String, aDesConst() As String, result As SCAN_RESULT
    Dim sAutostartFolder$(), sBaseFileName$, i&, k&, Wow6432Redir As Boolean, UseWow, sFolder$, sHit$
    Dim FldCnt&, sKey$, sSid$, sBasePath$, sBaseExt$, sTarget$, sFinalExecutable As String, bShortcut As Boolean, bPE_EXE As Boolean
    Dim bData() As Byte, isDisabled As Boolean, flagDisabled As Long, sKeyDisable As String, dDate As Date
    Dim StartupCU As String, aFiles() As String, sArguments As String, aUserNames() As String, aUserConst() As String, sUsername$
    Dim aFolders() As String, aHive() As String
    
    ReDim aRegKeys(1 To 8)
    ReDim aParams(1 To UBound(aRegKeys))
    ReDim aDesConst(1 To UBound(aRegKeys))
    ReDim aUserConst(1 To UBound(aRegKeys))

    ReDim sAutostartFolder(100) ' HKCU + HKLM + Wow64 + HKU
    ReDim aDes(100)
    ReDim aUserNames(100)
    ReDim aHive(100)
    
    'aRegKeys  - Key
    'aParams   - Value
    'aDesConst - Description for HJT Section
    
    'HKLM (HKLM hives should go first)
    aRegKeys(1) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(1) = "Common Startup"
    aDesConst(1) = "Startup Global"
    'aUserConst(1) = "All users" ' skip to make logs clear
    
    aRegKeys(2) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(2) = "Common AltStartup"
    aDesConst(2) = "StartupAlt Global"
    'aUserConst(2) = "All users"
    
    aRegKeys(3) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(3) = "Common Startup"
    aDesConst(3) = "Startup Global User"
    'aUserConst(3) = "All users"
    
    aRegKeys(4) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(4) = "Common AltStartup"
    aDesConst(4) = "StartupAlt Global User"
    'aUserConst(4) = "All users"
    
    'HKCU
    aRegKeys(5) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(5) = "Startup"
    aDesConst(5) = "Startup"
    'aUserConst(5) = envCurUser
    
    aRegKeys(6) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    aParams(6) = "AltStartup"
    aDesConst(6) = "StartupAlt"
    'aUserConst(6) = envCurUser
    
    aRegKeys(7) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(7) = "Startup"
    aDesConst(7) = "Startup User"
    'aUserConst(7) = envCurUser
    
    aRegKeys(8) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    aParams(8) = "AltStartup"
    aDesConst(8) = "StartupAlt User"
    'aUserConst(8) = envCurUser
    
    
    FldCnt = 0
    
    ' Get folder pathes
    For k = 1 To UBound(aRegKeys)
    
        For Each UseWow In Array(False, True)
            
            Wow6432Redir = UseWow
        
            'skip HKCU Wow64
            If (bIsWin32 And Wow6432Redir) _
              Or bIsWin64 And Wow6432Redir And StrBeginWith(aRegKeys(k), "HKCU") Then Exit For
    
            FldCnt = FldCnt + 1
            sAutostartFolder(FldCnt) = Reg.GetString(0&, aRegKeys(k), aParams(k), Wow6432Redir)
            aDes(FldCnt) = aDesConst(k)
            aUserNames(FldCnt) = aUserConst(k)
            aHive(FldCnt) = Reg.GetShortHiveName(aRegKeys(k))
            
            'save path of Startup for current user to substitute other user names
            If aParams(k) = "Startup" Then
                If Len(sAutostartFolder(FldCnt)) <> 0 Then
                    StartupCU = UnQuote(EnvironW(sAutostartFolder(FldCnt)))
                End If
            End If
        Next
    Next
    
    If InStr(1, StartupCU, UserProfile, 1) = 0 Then 'hijacked?
        If OSver.IsWindowsVistaOrGreater Then
            StartupCU = UserProfile & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        End If
    End If
    
    '+ HKU pathes
    For i = 0 To UBound(aSID)
        If Len(aSID(i)) <> 0 Then
            sSid = aSID(i)
            
            For k = 1 To UBound(aRegKeys)
            
                'only HKCU keys
                If StrBeginWith(aRegKeys(k), "HKCU") Then
                
                    ' Convert HKCU -> HKU
                    sKey = Replace$(aRegKeys(k), "HKCU\", "HKU\" & sSid & "\")
                
                    FldCnt = FldCnt + 1
                    If UBound(sAutostartFolder) < FldCnt Then
                        ReDim Preserve sAutostartFolder(UBound(sAutostartFolder) + 100)
                        ReDim Preserve aDes(UBound(aDes) + 100)
                        ReDim Preserve aUserNames(UBound(aUserNames) + 100)
                    End If
                    
                    sAutostartFolder(FldCnt) = Reg.GetString(0&, sKey, aParams(k), , bDoNotExpand:=True)
                    sAutostartFolder(FldCnt) = EnvironW(sAutostartFolder(FldCnt), , GetProfileDirBySID(sSid))
                    
                    aDes(FldCnt) = aDesConst(k) & " Other"
                    aHive(FldCnt) = Reg.GetShortHiveName(sKey, bIncludeSID:=True)
                    aUserNames(FldCnt) = MapSIDToUsername(sSid)
                End If
            Next
        End If
    Next
    
    ReDim Preserve sAutostartFolder(FldCnt)
    ReDim Preserve aDes(FldCnt)
    ReDim Preserve aUserNames(FldCnt)
    
    For k = 1 To UBound(sAutostartFolder)
        sAutostartFolder(k) = UnQuote(EnvironW(sAutostartFolder(k)))
    Next
    
    ' adding all similar folders in c:\users (in case user isn't logged - so HKU\SID willn't be exist for him, cos his hive is not mounted)
    
    For i = 1 To colProfiles.Count
        'not current user
        If StrComp(colProfiles(i), UserProfile, 1) <> 0 Then
            If Len(colProfiles(i)) <> 0 Then
                ReDim Preserve sAutostartFolder(UBound(sAutostartFolder) + 1)
                ReDim Preserve aDes(UBound(aDes) + 1)
                ReDim Preserve aUserNames(UBound(aUserNames) + 1)
                sAutostartFolder(UBound(sAutostartFolder)) = Replace$(StartupCU, UserProfile, colProfiles(i), 1, 1, 1)
                aDes(UBound(aDes)) = "Startup Other"
                aUserNames(UBound(aUserNames)) = colProfilesUser(i)
            End If
        End If
    Next
    
    DeleteDuplicatesInArray sAutostartFolder, vbTextCompare, DontCompress:=True
    
    For k = 1 To UBound(sAutostartFolder)
        
        sUsername = aUserNames(k)
        
        sFolder = sAutostartFolder(k)
        
        If 0 <> Len(sFolder) Then
          If FolderExists(sFolder) Then
            
            Erase aFolders
            aFolders = ListSubfolders(sFolder)
            
            For i = 0 To UBoundSafe(aFolders)
            
                sHit = "O4 - " & aDes(k) & ": " & aFolders(i) & " (folder)"
            
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O4"
                        .HitLineW = sHit
                        AddFileToFix .File, REMOVE_FOLDER, aFolders(i)
                        .CureType = FILE_BASED
                    End With
                    AddToScanResults result
                End If
            Next
            
            Erase aFiles
            aFiles = ListFiles(sFolder)
            
            For i = 0 To UBoundSafe(aFiles)
            
                sBasePath = aFiles(i)
                sBaseFileName = GetFileNameAndExt(sBasePath)
                sBaseExt = UCase$(GetExtensionName(sBaseFileName))
                
                If (LCase$(sBaseFileName) <> "desktop.ini" Or bIgnoreAllWhitelists) Then
                  
                  'wtf is this?
                  'If Not FolderExists(sFolder & "\" & sBaseFileName) Then
                  
                    isDisabled = False
              
                    If OSver.MajorMinor >= 6.2 Then  ' Win 8+

                        If InStr(aDes(k), "StartupAlt") = 0 And Len(aHive(k)) <> 0 Then

                            sKeyDisable = aHive(k) & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder"

                            If Reg.ValueExists(0&, sKeyDisable, sBaseFileName) Then

                                ReDim bData(0)
                                bData() = Reg.GetBinary(0&, sKeyDisable, sBaseFileName)
                                
                                If UBoundSafe(bData) >= 11 Then
                            
                                    GetMem4 ByVal VarPtr(bData(0)), flagDisabled

                                    If flagDisabled <> 2 Then
                        
                                        isDisabled = True
                                        dDate = ConvertFileTimeToLocalDate(VarPtr(bData(4)))
                                    End If
                                End If
                            End If
                        End If
                    End If
                  
                    bShortcut = False
                    bPE_EXE = False
                    
                    'Example:
                    '"O4 - Global User AltStartup: "
                    '"O4 - S-1-5-19 User AltStartup: "
                    If isDisabled Then
                        sHit = "O4 - " & aHive(k) & "\..\StartupApproved\StartupFolder: " 'if you change it, change fix also !!!
                    Else
                        sHit = "O4 - " & aDes(k) & ": "
                    End If
                    
                    If StrInParamArray(sBaseExt, ".LNK", ".URL", ".WEBSITE", ".PIF") Then bShortcut = True
                    
                    If Not bShortcut Or sBaseExt = ".PIF" Then  'not a Shortcut ?
                        bPE_EXE = isPE(sBasePath)       'PE EXE ?
                    End If
                    
                    sArguments = vbNullString
                    sTarget = vbNullString
                    sArguments = vbNullString
                    
                    If bShortcut Then
                        sTarget = GetFileFromShortcut(sBasePath, sArguments)
                        sTarget = FormatFileMissing(sTarget)
                        sFinalExecutable = sTarget
                        sHit = sHit & sBasePath & "    ->    " & sTarget & IIf(Len(sArguments) <> 0, " " & sArguments, vbNullString)
                    Else
                        sFinalExecutable = sBasePath
                        sHit = sHit & sBasePath & IIf(bPE_EXE, "    ->    (PE EXE)", vbNullString)
                    End If
                    
                    If Len(sUsername) <> 0 Then sHit = sHit & " (User '" & sUsername & "')"
                    
                    If isDisabled Then sHit = sHit & IIf(dDate <> dEpoch, " (" & Format$(dDate, "yyyy\/mm\/dd") & ")", vbNullString)
                    
                    If IsScriptExtension(sFinalExecutable) Then
                        sHit = sHit & "    =>    " & _
                            ReadFileContents(sFinalExecutable, FileGetTypeBOM(sFinalExecutable) = CP_UTF16LE)
                    End If
                    
                    If Not bShortcut Or bPE_EXE Then
                        SignVerifyJack sBasePath, result.SignResult
                        sHit = sHit & FormatSign(result.SignResult)
                        If g_bCheckSum Then
                            sHit = sHit & GetFileCheckSum(sBasePath)
                        End If
                    Else
                        If 0 <> Len(sTarget) Then
                            SignVerifyJack sTarget, result.SignResult
                            sHit = sHit & FormatSign(result.SignResult)
                            If g_bCheckSum Then
                                sHit = sHit & GetFileCheckSum(sTarget)
                            End If
                        End If
                    End If
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                          .Section = "O4"
                          .HitLineW = sHit
                          
                          If isDisabled Then
                            .Alias = aHive(k) & "\..\StartupApproved\StartupFolder:"
                            AddRegToFix .Reg, REMOVE_VALUE, 0&, sKeyDisable, sBaseFileName, , REG_NOTREDIRECTED
                            If bShortcut Then ' should go first (for VT)
                                AddProcessToFix .Process, FREEZE_OR_KILL_PROCESS, sTarget
                            Else
                                AddProcessToFix .Process, FREEZE_OR_KILL_PROCESS, sBasePath
                            End If
                            AddFileToFix .File, REMOVE_FILE, sBasePath
                            .CureType = FILE_BASED Or REGISTRY_BASED Or PROCESS_BASED
                          Else
                            .Alias = aDes(k)
                            AddProcessToFix .Process, FREEZE_OR_KILL_PROCESS, sTarget
                            If bShortcut Then ' should go first (for VT)
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sTarget
                            End If
                            AddFileToFix .File, REMOVE_FILE, sBasePath
                            AddJumpFile .Jump, JUMP_FILE, sTarget, sArguments
                            .CureType = FILE_BASED Or PROCESS_BASED
                          End If
                        End With
                        AddToScanResults result
                    End If
                  'End If
                End If
            Next
          End If
        End If
    Next
    
    AppendErrorLogCustom "CheckO4_AutostartFolder - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.CheckO4_AutostartFolder"
    If inIDE Then Stop: Resume Next
End Sub


Sub CheckO4_ActiveSetup() 'Thanks to Helge Klein for explanations
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4_ActiveSetup - Begin"
    
    Dim result As SCAN_RESULT
    Dim bDisabled As Boolean, bInstalled As Boolean, bSafe As Boolean
    Dim i As Long
    Dim sHit As String, sAlias As String, sData As String, sFile As String, sArgs As String, sHash As String, sDll As String
    Dim aSubKey() As String
    Dim dWhitelist As clsTrickHashTable
    Set dWhitelist = New clsTrickHashTable
    dWhitelist.CompareMode = vbTextCompare
    
    'key = file + argument directly as seen in log; value = additional file required to be checked by signature
    dWhitelist.Add BuildPath(PF_64, "Windows Mail\WinMail.exe OCInstallUserConfigOE"), ""
    dWhitelist.Add BuildPath(sWinSysDir, "ie4uinit.exe -DisableSSL3"), ""
    dWhitelist.Add BuildPath(sWinSysDir, "ie4uinit.exe -EnableTLS"), ""
    dWhitelist.Add BuildPath(sWinSysDir, "ie4uinit.exe -UserConfig"), ""
    dWhitelist.Add BuildPath(sWinSysDir, "regsvr32.exe /s /n /i:/UserInstall " & sWinSysDir & "\themeui.dll"), sWinSysDir & "\themeui.dll"
    dWhitelist.Add BuildPath(sWinSysDir, "regsvr32.exe /s /n /i:U shell32.dll"), ""
    dWhitelist.Add BuildPath(sWinSysDir, "Rundll32.exe " & sWinSysDir & "\mscories.dll,Install"), sWinSysDir & "\mscories.dll"
    dWhitelist.Add BuildPath(sWinSysDir, "unregmp2.exe /FirstLogon /Shortcuts /RegBrowsers /ResetMUI"), ""
    dWhitelist.Add BuildPath(sWinSysDir, "unregmp2.exe /ShowWMP"), ""
    dWhitelist.Add BuildPath(sWinSysDir, "unregmp2.exe /FirstLogon"), ""
    dWhitelist.Add BuildPath(sWinSysDir, "Rundll32.exe """ & sWinSysDir & "\iesetup.dll"",IEHardenAdmin"), sWinSysDir & "\iesetup.dll"
    dWhitelist.Add BuildPath(sWinSysDir, "Rundll32.exe """ & sWinSysDir & "\iesetup.dll"",IEHardenUser"), sWinSysDir & "\iesetup.dll"
    If OSver.IsWindowsVista Then
        dWhitelist.Add BuildPath(sWinSysDir, "ie4uinit.exe -BaseSettings"), ""
        dWhitelist.Add BuildPath(sWinSysDir, "ie4uinit.exe -UserIconConfig"), ""
        dWhitelist.Add BuildPath(sWinSysDir, "RunDLL32.exe IEDKCS32.DLL,BrandIE4 SIGNUP"), ""
    End If
    
    Dim HE As clsHiveEnum: Set HE = New clsHiveEnum
    HE.Init HE_HIVE_HKLM
    HE.AddKey "SOFTWARE\Microsoft\Active Setup\Installed Components"
    
    Do While HE.MoveNext
        
        For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey(), HE.Redirected)
        
            sData = Reg.GetString(HE.Hive, HE.Key & "\" & aSubKey(i), "StubPath")
            
            If Len(sData) <> 0 Then
                bDisabled = (0 = Reg.GetDword(HE.Hive, HE.Key & "\" & aSubKey(i), "IsInstalled"))
                If Reg.StatusCode <> ERROR_SUCCESS Then bDisabled = False
                
                'Disabled just in case
                'If (Not bDisabled) Or bIgnoreAllWhitelists Then
                If True Then
                
                    result.SignResult.FilePathVerified = vbNullString
                
                    bInstalled = Reg.KeyExists(HE.Hive, HE.Key & "\" & aSubKey(i), HE.Redirected)
                    
                    SplitIntoPathAndArgs sData, sFile, sArgs, True
                    
                    sFile = FormatFileMissing(sFile)
                    
                    sData = ConcatFileArg(sFile, sArgs)
                    
                    bSafe = dWhitelist.Exists(sData)
                    If bSafe Then
                        SignVerifyJack sFile, result.SignResult
                        bSafe = result.SignResult.isMicrosoftSign
                        If bSafe Then
                            sDll = dWhitelist(sData)
                            If Len(sDll) <> 0 Then
                                SignVerifyJack sDll, result.SignResult
                                bSafe = result.SignResult.isMicrosoftSign
                            End If
                        End If
                    Else 'dynamic database (e.g. version in path)
                        If StrBeginWith(sData, PF_32 & "\Microsoft\Edge\Application\") Then
                            If sArgs = "--configure-user-settings --verbose-logging --system-level --msedge --channel=stable" Then
                                SignVerifyJack sFile, result.SignResult
                                bSafe = result.SignResult.isMicrosoftSign
                            End If
                        End If
                    End If
                    
                    If (Not bSafe) Then
                        'garbage by MS :)
                        If sData = "U " & STR_FILE_MISSING Then
                            bSafe = True
                        ElseIf sData = "/UserInstall " & STR_FILE_MISSING Then
                            bSafe = True
                        End If
                    End If
                    
                    If (Not bSafe) Or (Not bHideMicrosoft) Then
                        
                        If Len(result.SignResult.FilePathVerified) = 0 Then SignVerifyJack sFile, result.SignResult
                        
                        sAlias = BitPrefix("O4", HE)
                        sHit = sAlias & " - ActiveSetup: HKLM\..\" & aSubKey(i) & ": [StubPath] = "
                        sHit = sHit & sData & FormatSign(result.SignResult) & _
                            IIf(bDisabled, " (disabled)", "") & IIf(bInstalled, "", " (fresh)")
                        
                        If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O4"
                                .HitLineW = sHit
                                .Alias = sAlias
                                AddRegToFix .Reg, REMOVE_KEY, HKLM, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                                'HKLM key is duplicated here when component is installed
                                AddRegToFix .Reg, REMOVE_KEY, HKCU, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                End If
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckO4_ActiveSetup - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain.CheckO4_ActiveSetup"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckO4Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO4Item - Begin"

    ' Keys affected by wow64 redirector:
    ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa384253(v=vs.85).aspx
    ' http://safezone.cc/threads/27567/
    
    CheckO4_RegRuns
    
    CheckO4_MSConfig gHives(), gUserOfHive()
    
    CheckO4_AutostartFolder gSIDs(), gUserOfHive()
    
    CheckO4_ActiveSetup
    
    AppendErrorLogCustom "CheckO4Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO4Item"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub FixO4Item(sItem$, result As SCAN_RESULT)
    'O4 - Enumeration of autoloading Regedit entries
    'O4 - HKLM\..\Run: [blah] program.exe
    'O4 - Startup: bla.lnk = c:\bla.exe
    'O4 - HKU\S-1-5-19\..\Run: [blah] program.exe (Username 'Joe')
    'O4 - Startup: bla.exe
    'O4 - MSConfig:
    'O4 - \..\StartupApproved\StartupFolder:
    '...
    
    On Error GoTo ErrorHandler:
    
    Dim sFile$

    FixProcessHandler result
    
    If InStr(sItem, "StartupApproved\StartupFolder") <> 0 Then
        
        sFile = result.File(0).path
        
        If FileExists(sFile) Then
            If DeleteFileForce(sFile) Then
                FixRegistryHandler result 'remove registry value if only file successfully deleted (!!!)
            End If
        Else
            FixRegistryHandler result
        End If
        
        Exit Sub
    End If
    
    FixFileHandler result
    FixRegistryHandler result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO4Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO5Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO5Item - Begin"
    
    Dim sControlIni$, sDummy$, sHit$, result As SCAN_RESULT
    Dim i&, aValues() As String, bSafe As Boolean, bFileExist As Boolean, sSnapIn As String, sDescr As String, sPath As String
    Dim aParams() As Variant
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    '// TODO: add also:
    
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => DisallowCpl = 1
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => RestrictCpl = 1
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "HKCU\Control Panel\don't load"
    HE.AddKey "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\don't load"
            
    Do While HE.MoveNext
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValues, HE.Redirected)
            
            bSafe = False
            sSnapIn = aValues(i)
            
            If bHideMicrosoft Then
                If HE.Hive = HKCU Or HE.Hive = HKU Then
                    If inArraySerialized(sSnapIn, sSafeO5Items_HKU, "|", , , 1) Then bSafe = True
                Else
                    If HE.Redirected Then
                        If inArraySerialized(sSnapIn, sSafeO5Items_HKLM_32, "|", , , 1) Then bSafe = True
                    Else
                        If inArraySerialized(sSnapIn, sSafeO5Items_HKLM, "|", , , 1) Then bSafe = True
                    End If
                End If
            End If
            
            If Not bSafe Then
                sPath = BuildPath(IIf(HE.Redirected, sWinSysDirWow64, sWinSysDir), sSnapIn)
                bFileExist = FileExists(sPath)
                sDescr = vbNullString
                If bFileExist Then
                    sDescr = GetFileProperty(sPath, "FileDescription")
                    SignVerifyJack sPath, result.SignResult
                End If
                
                sHit = BitPrefix("O5", HE) & " - " & HE.KeyAndHivePhysical & ": [" & sSnapIn & "]" & _
                    IIf(Len(sDescr) <> 0, " (" & sDescr & ")", vbNullString) & IIf(bFileExist, vbNullString, " " & STR_FILE_MISSING) & FormatSign(result.SignResult)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O5"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sSnapIn, , HE.Redirected
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    sControlIni = sWinDir & "\control.ini"
    If Not FileExists(sControlIni) Then GoTo SkipControlIni
    
    Dim cIni As clsIniFile
    Set cIni = New clsIniFile
    
    cIni.InitFile sControlIni, 1251
    
    If cIni.CountParams("don't load") > 0 Then
        aParams = cIni.GetParamNames("don't load")
        
        For i = 0 To UBound(aParams)
            sSnapIn = aParams(i)
            sDummy = Trim$(cIni.ReadParam("don't load", sSnapIn))
            
            If Len(sDummy) <> 0 Then
                sPath = BuildPath(IIf(HE.Redirected, sWinSysDirWow64, sWinSysDir), sSnapIn)
                bFileExist = FileExists(sPath)
                sDescr = vbNullString
                If bFileExist Then
                    sDescr = GetFileProperty(sPath, "FileDescription")
                    SignVerifyJack sPath, result.SignResult
                End If
                
                sHit = "O5 - control.ini: [don't load] " & sSnapIn & " = " & sDummy & _
                    IIf(Len(sDescr) <> 0, " (" & sDescr & ")", vbNullString) & IIf(bFileExist, vbNullString, " " & STR_FILE_MISSING) & FormatSign(result.SignResult)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O5"
                        .HitLineW = sHit
                        AddIniToFix .Reg, RESTORE_VALUE_INI, "control.ini", "don't load", sSnapIn, vbNullString
                        .CureType = INI_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    End If

    Set cIni = Nothing
    
SkipControlIni:
    
    Dim aFiles() As String
    Dim vFolder As Variant
    
    For Each vFolder In Array(sWinSysDir, sWinSysDirWow64)
    
        aFiles = ListFiles(CStr(vFolder), ".cpl")
        If AryItems(aFiles) Then
            For i = 0 To UBound(aFiles)
                sPath = aFiles(i)
                sPath = FormatFileMissing(sPath)
                
                bSafe = True
                
                SignVerifyJack sPath, result.SignResult
                If Not result.SignResult.isMicrosoftSign Then bSafe = False
                
                If Not bSafe Then
                    
                    sHit = "O5 - Applet: " & sPath & FormatSign(result.SignResult)
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sPath)
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O5"
                            .HitLineW = sHit
                            AddFileToFix .File, REMOVE_FILE Or RESTORE_FILE_SFC, sPath
                            .CureType = FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        End If
    Next
    
    '// todo:
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace
    '
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\Cpls
    '+ HKCU + wow
    
    AppendErrorLogCustom "CheckO5Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO5Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO5Item(sItem$, result As SCAN_RESULT)
    FixIt result
End Sub

Public Sub CheckO6Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO6Item - Begin"
    
    'If there are sub folders called
    '"restrictions" and/or "control panel", delete them
    
    Dim sHit$, Key$(2), result As SCAN_RESULT
    'keys 0,1,2 - are x6432 shared.
    
    Key(0) = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    Key(1) = "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions"
    Key(2) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys Key()
    
    Do While HE.MoveNext
        If Reg.KeyHasValues(HE.Hive, HE.Key, HE.Redirected) Then
            sHit = BitPrefix("O6", HE) & " - IE Policy: " & HE.HiveNameAndSID & "\" & HE.Key & " - present"
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O6"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key, , , HE.Redirected
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Loop
    
    'Restrict using IE settings by HKLM hive only ?
    If Reg.GetDword(HKLM, "SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings", "Security_HKLM_only") = 1 Then
        sHit = "O6 - IE Policy: HKLM\..\Internet Settings: [Security_HKLM_only] = 1"
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O6"
                .HitLineW = sHit
                AddRegToFix .Reg, REMOVE_VALUE, HKLM, "SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings", "Security_HKLM_only"
                .CureType = REGISTRY_BASED
            End With
            AddToScanResults result
        End If
    End If
    
    AppendErrorLogCustom "CheckO6Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO6Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO6Item(sItem$, result As SCAN_RESULT)
    FixIt result
End Sub

Public Sub CheckSystemProblems()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckSystemProblems - Begin"
    
    Call CheckKnownFolders
    Call CheckSystemProblemsEnvVars
    Call CheckSystemProblemsFreeSpace
    Call CheckSystemProblemsNetwork
    
    AppendErrorLogCustom "CheckSystemProblems - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblems"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckKnownFolders()

    'temporarily
    If Not OSver.IsWindowsVistaOrGreater Then Exit Sub

    CheckKnownFoldersHKLM
    
    If Not OSver.IsLocalSystemContext Then
        CheckKnownFoldersHKCU
    End If
End Sub

Public Sub CheckKnownFoldersHKLM()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "CheckKnownFoldersHKLM - Begin"
    
    Dim sHit As String, result As SCAN_RESULT
    
    Dim aKey(20) As String
    Dim aParam() As String
    Dim aValue() As String
    ReDim aParam(UBound(aKey)) As String
    ReDim aValue(UBound(aKey)) As String
    
    aKey(0) = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    
    aParam(0) = "Common Administrative Tools"
    aValue(0) = SysDisk & "\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools"
    
    aParam(1) = "Common AppData"
    aValue(1) = SysDisk & "\ProgramData"
    
    aParam(2) = "Common Desktop"
    aValue(2) = SysDisk & "\Users\Public\Desktop"
    
    aParam(3) = "Common Documents"
    aValue(3) = SysDisk & "\Users\Public\Documents"
    
    aParam(4) = "Common Programs"
    aValue(4) = SysDisk & "\ProgramData\Microsoft\Windows\Start Menu\Programs"
    
    aParam(5) = "Common Start Menu"
    aValue(5) = SysDisk & "\ProgramData\Microsoft\Windows\Start Menu"
    
    aParam(6) = "Common Startup"
    aValue(6) = SysDisk & "\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
    
    aParam(7) = "Common Templates"
    aValue(7) = SysDisk & "\ProgramData\Microsoft\Windows\Templates"
    
    aParam(8) = "CommonMusic"
    aValue(8) = SysDisk & "\Users\Public\Music"
    
    aParam(9) = "CommonPictures"
    aValue(9) = SysDisk & "\Users\Public\Pictures"
    
    aParam(10) = "CommonVideo"
    aValue(10) = SysDisk & "\Users\Public\Videos"
    
    '--------------
    
    aKey(11) = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    
    aParam(11) = "Common AppData"
    aValue(11) = "%ProgramData%"
    
    aParam(12) = "Common Desktop"
    aValue(12) = "%PUBLIC%\Desktop"
    
    aParam(13) = "Common Documents"
    aValue(13) = "%PUBLIC%\Documents"
    
    aParam(14) = "Common Programs"
    aValue(14) = "%ProgramData%\Microsoft\Windows\Start Menu\Programs"
    
    aParam(15) = "Common Start Menu"
    aValue(15) = "%ProgramData%\Microsoft\Windows\Start Menu"
    
    aParam(16) = "Common Startup"
    aValue(16) = "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Startup"
    
    aParam(17) = "Common Templates"
    aValue(17) = "%ProgramData%\Microsoft\Windows\Templates"
    
    aParam(18) = "CommonMusic"
    aValue(18) = "%PUBLIC%\Music"
    
    aParam(19) = "CommonPictures"
    aValue(19) = "%PUBLIC%\Pictures"
    
    aParam(20) = "CommonVideo"
    aValue(20) = "%PUBLIC%\Videos"
    
    Dim sKey As String, sValue As String, sDefValue As String, sValueExpanded As String
    Dim i As Long
    Dim dictChecked As clsTrickHashTable
    Set dictChecked = New clsTrickHashTable
    
    For i = 0 To UBound(aKey)
    
        If Len(aKey(i)) <> 0 Then sKey = aKey(i)
    
        If Len(aParam(i)) <> 0 Then
            
            sValue = Reg.GetString(0&, sKey, aParam(i), bDoNotExpand:=True)
            
            If (StrComp(sValue, aValue(i), vbTextCompare) <> 0) Or bIgnoreAllWhitelists Then
                
                sHit = "O7 - KnownFolder: " & sKey & ", " & aParam(i) & " = " & sValue
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, RESTORE_VALUE, 0, sKey, aParam(i), aValue(i), , REG_RESTORE_EXPAND_SZ
                        AddFileToFix .File, CREATE_FOLDER, EnvironW(aValue(i))
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
            
            'If the folder is redirected legally we don't need extra "folder missing" records
            'sDefValue = EnvironW(aValue(i))
            sValueExpanded = EnvironW(sValue)
            
            If Not dictChecked.Exists(sValueExpanded) Then
                dictChecked.Add sValueExpanded, 0
            
                If Not FolderExists(sValueExpanded) Then
                    
                    sHit = "O7 - KnownFolder: " & sValueExpanded & " " & STR_FOLDER_MISSING
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            AddFileToFix .File, CREATE_FOLDER, sValueExpanded
                            .CureType = FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
            
        End If
    Next
    
    AppendErrorLogCustom "CheckKnownFoldersHKLM - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckKnownFoldersHKLM"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckKnownFoldersHKCU()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "CheckKnownFoldersHKCU - Begin"
    
    Dim sHit As String, result As SCAN_RESULT
    
    Dim aKey(49) As String
    Dim aParam() As String
    Dim aValue() As String
    ReDim aParam(UBound(aKey)) As String
    ReDim aValue(UBound(aKey)) As String
    
    aKey(0) = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    
    If OSver.IsWindows7OrGreater Then
        aParam(0) = "{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}"
        aValue(0) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Libraries"
        
        aParam(1) = "{374DE290-123F-4565-9164-39C4925E467B}"
        aValue(1) = "%UserProfile%\Downloads"
        
        aParam(2) = "{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}"
        aValue(2) = "%UserProfile%\Saved Games"
        
        aParam(3) = "{56784854-C6CB-462B-8169-88E350ACB882}"
        aValue(3) = "%UserProfile%\Contacts"
        
        aParam(4) = "{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}"
        aValue(4) = "%UserProfile%\Searches"
        
        aParam(5) = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
        aValue(5) = "%UserProfile%\AppData\LocalLow"
        
        aParam(6) = "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}"
        aValue(6) = "%UserProfile%\Links"
    End If
        
    If OSver.IsWindowsVistaOrGreater Then
        aParam(7) = "Administrative Tools"
        aValue(7) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Administrative Tools"
    End If
    
    aParam(8) = "AppData"
    aValue(8) = "%UserProfile%\AppData\Roaming"
    
    aParam(9) = "Cache"
    If OSver.IsWindows8OrGreater Then
        aValue(9) = "%UserProfile%\AppData\Local\Microsoft\Windows\INetCache"
    Else
        aValue(9) = "%UserProfile%\AppData\Local\Microsoft\Windows\Temporary Internet Files"
    End If
    
    aParam(10) = "CD Burning"
    aValue(10) = "%UserProfile%\AppData\Local\Microsoft\Windows\Burn\Burn"
    
    aParam(11) = "Cookies"
    If OSver.IsWindows8OrGreater Then
        aValue(11) = "%UserProfile%\AppData\Local\Microsoft\Windows\INetCookies"
    Else
        aValue(11) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Cookies"
    End If
    
    aParam(12) = "Desktop"
    aValue(12) = "%UserProfile%\Desktop"
    
    aParam(13) = "Favorites"
    aValue(13) = "%UserProfile%\Favorites"
    
    aParam(14) = "Fonts"
    aValue(14) = sWinDir & "\Fonts"
    
    aParam(15) = "History"
    aValue(15) = "%UserProfile%\AppData\Local\Microsoft\Windows\History"
    
    aParam(16) = "Local AppData"
    aValue(16) = "%UserProfile%\AppData\Local"
    
    aParam(17) = "My Music"
    aValue(17) = "%UserProfile%\Music"
    
    aParam(18) = "My Pictures"
    aValue(18) = "%UserProfile%\Pictures"
    
    If OSver.IsWindowsVistaOrGreater Then
        aParam(19) = "My Video"
        aValue(19) = "%UserProfile%\Videos"
    End If
    
    aParam(20) = "NetHood"
    aValue(20) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Network Shortcuts"
    
    aParam(21) = "Personal"
    aValue(21) = "%UserProfile%\Documents"
    
    aParam(23) = "PrintHood"
    aValue(23) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Printer Shortcuts"
    
    aParam(24) = "Programs"
    aValue(24) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    
    aParam(25) = "Recent"
    aValue(25) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Recent"
    
    aParam(26) = "SendTo"
    aValue(26) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\SendTo"
    
    aParam(27) = "Start Menu"
    aValue(27) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Start Menu"
    
    aParam(28) = "Startup"
    aValue(28) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    
    aParam(29) = "Templates"
    aValue(29) = "%UserProfile%\AppData\Roaming\Microsoft\Windows\Templates"
    
    ' -----------------
    
    aKey(30) = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    
    If OSver.IsWindowsVistaOrGreater Then
        aParam(30) = "{374DE290-123F-4565-9164-39C4925E467B}"
        aValue(30) = "%USERPROFILE%\Downloads"
    End If
    
    aParam(31) = "AppData"
    aValue(31) = "%USERPROFILE%\AppData\Roaming"
    
    aParam(32) = "Cache"
    If OSver.IsWindows8OrGreater Then
        aValue(32) = "%USERPROFILE%\AppData\Local\Microsoft\Windows\INetCache"
    Else
        aValue(32) = "%USERPROFILE%\AppData\Local\Microsoft\Windows\Temporary Internet Files"
    End If
    
    aParam(33) = "Cookies"
    If OSver.IsWindows8OrGreater Then
        aValue(33) = "%USERPROFILE%\AppData\Local\Microsoft\Windows\INetCookies"
    Else
        aValue(33) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Cookies"
    End If
    
    aParam(34) = "Desktop"
    aValue(34) = "%USERPROFILE%\Desktop"
    
    aParam(35) = "Favorites"
    aValue(35) = "%USERPROFILE%\Favorites"
    
    aParam(36) = "History"
    aValue(36) = "%USERPROFILE%\AppData\Local\Microsoft\Windows\History"
    
    aParam(37) = "Local AppData"
    aValue(37) = "%USERPROFILE%\AppData\Local"
    
    aParam(38) = "My Pictures"
    aValue(38) = "%USERPROFILE%\Pictures"
    
    If OSver.IsWindowsVistaOrGreater Then
        aParam(39) = "My Music"
        aValue(39) = "%USERPROFILE%\Music"
    
        aParam(40) = "My Video"
        aValue(40) = "%USERPROFILE%\Videos"
    End If
    
    aParam(41) = "NetHood"
    aValue(41) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Network Shortcuts"
    
    aParam(42) = "Personal"
    aValue(42) = "%USERPROFILE%\Documents"
    
    aParam(43) = "PrintHood"
    aValue(43) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Printer Shortcuts"
    
    aParam(44) = "Programs"
    aValue(44) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    
    aParam(45) = "Recent"
    aValue(45) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Recent"
    
    aParam(46) = "SendTo"
    aValue(46) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\SendTo"
    
    aParam(47) = "Start Menu"
    aValue(47) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu"
    
    aParam(48) = "Startup"
    aValue(48) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    
    aParam(49) = "Templates"
    aValue(49) = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Templates"
    
    Dim sKey As String, sValue As String, sDefValue As String, sSid As String, sValueExpanded As String
    Dim i As Long, k As Long, pos As Long, sProfile As String
    Dim bSafe As Boolean
    Dim dictChecked As clsTrickHashTable
    Set dictChecked = New clsTrickHashTable
    
    For k = 0 To UBound(gHivesUser)

        For i = 0 To UBound(aKey)
        
            If Len(aKey(i)) <> 0 Then sKey = gHivesUser(k) & "\" & aKey(i)
        
            If Len(aParam(i)) <> 0 Then
                
                sValue = Reg.GetString(0&, sKey, aParam(i), bDoNotExpand:=True)
                
                bSafe = False
                
                If i >= 0 And i <= 29 Then
                    If gHivesUser(k) = "HKCU" Then
                        sDefValue = EnvironW(aValue(i))
                        sProfile = UserProfile
                    Else
                        pos = InStr(gHivesUser(k), "\")
                        If pos <> 0 Then
                            sSid = mid$(gHivesUser(k), pos + 1)
                            sProfile = GetProfileDirBySID(sSid)
                            sDefValue = EnvironW(aValue(i), , sProfile)
                        End If
                    End If
                Else
                    sDefValue = aValue(i)
                End If
                
                'sometimes 'Shell Folders' key is volatile!
                If Len(sValue) = 0 Then
                    If gHivesUser(k) <> "HKCU" And Not bIgnoreAllWhitelists Then bSafe = True
                End If
                
                If Not bSafe Then
                    If StrComp(sValue, sDefValue, vbTextCompare) <> 0 Then
                        
                        If Not bIgnoreAllWhitelists Then
                            'if Desktop is shared (moved) to OneDrive
                            If StrBeginWith(sValue, BuildPath(sProfile, "OneDrive")) Then bSafe = True
                        End If
                        
                        If Not bSafe Then
                            'sHit = "O7 - KnownFolder: " & sKey & ", " & aParam(i) & " = " & sValue & " - Should be: " & sDefValue
                            sHit = "O7 - KnownFolder: " & sKey & ", " & aParam(i) & " = " & sValue
                            
                            If Not IsOnIgnoreList(sHit) Then
                                With result
                                    .Section = "O7"
                                    .HitLineW = sHit
                                    AddRegToFix .Reg, RESTORE_VALUE, 0, sKey, aParam(i), sDefValue, , REG_RESTORE_EXPAND_SZ
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults result
                            End If
                        End If
                    End If
                End If
                
                'If the folder is redirected legally we don't need extra "folder missing" records
                'sDefValue = EnvironW(sDefValue)
                sValueExpanded = EnvironW(sValue, , sProfile)
                
                If Not dictChecked.Exists(sValueExpanded) Then
                    dictChecked.Add sValueExpanded, 0
                
                    If Not FolderExists(sValueExpanded) Then
                        
                        sHit = "O7 - KnownFolder: " & sValueExpanded & " " & STR_FOLDER_MISSING
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O7"
                                .HitLineW = sHit
                                AddFileToFix .File, CREATE_FOLDER, sValueExpanded
                                .CureType = FILE_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                End If
                
            End If
        Next
    Next
    
    AppendErrorLogCustom "CheckKnownFoldersHKCU - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckKnownFoldersHKCU"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckSystemProblemsEnvVars()
    CheckEnvVarPath
    CheckEnvVarPathExt
    CheckEnvVarTemp
    CheckEnvVarOther
End Sub

Public Sub CheckEnvVarPath()

    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckEnvVarPath - Begin"
    
    Dim sData As String
    Dim vParam, sKeyFull As String, sHit As String, result As SCAN_RESULT
    Dim aLine() As String, i As Long, vValue, bSafe As Boolean, sPsPath As String
    
    '// TODO:
    ' add checking %PATH% load order
    
    '// TODO:
    ' PATH len exceed the maximum allowed, see article:
    ' https://safezone.cc/threads/delo-o-zablokirovannoj-peremennoj-okruzhenija-path.31001/
    ' Check essential programs, e.g. scripting hosts, by search path.
    
    sKeyFull = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    
    sData = Reg.GetString(0, sKeyFull, "Path")
    aLine = Split(sData, ";")
    
    For i = 0 To UBoundSafe(aLine)
        aLine(i) = LTrim$(aLine(i))
        If StrEndWith(aLine(i), "\") Then aLine(i) = Left$(aLine(i), Len(aLine(i)) - 1) 'cut last \
    Next
    
    If OSver.IsWindows7OrGreater Then sPsPath = BuildPath(sWinSysDir, "WindowsPowerShell\v1.0")
    
    For Each vValue In Array(sWinDir, sWinSysDir, BuildPath(sWinSysDir, "Wbem"), sPsPath)
        
        bSafe = False
        
        If Len(vValue) = 0 Then
            bSafe = True
        Else
            If AryItems(aLine) Then
                If InArray(CStr(vValue), aLine, , , vbTextCompare) Then bSafe = True
            End If
        End If
        
        If Not bSafe Then
            sHit = "O7 - TroubleShooting (EV): %PATH% has missing system folder: " & vValue
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, APPEND_VALUE_NO_DOUBLE, HKEY_ANY, sKeyFull, "Path", EnvironUnexpand(CStr(vValue)) & ";", , REG_RESTORE_EXPAND_SZ, , , ";"
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    AppendErrorLogCustom "CheckEnvVarPath - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckEnvVarPath"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckEnvVarTemp()

    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckEnvVarTemp - Begin"
    
    'Check for %Temp% anomalies
    'Checking for present and correct type of parameters:
    'HKCU\Environment => temp, tmp
    'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment => temp, tmp ("%SystemRoot%\TEMP")
    
    Dim vParam, sKeyFull As String, sHit As String, result As SCAN_RESULT
    Dim bComply As Boolean, sData As String, sDataNonExpanded As String, sDefValue As String
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    If OSver.IsWindows7OrGreater Then
        '7+
        HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    Else
        'Vista-
        HE.Init HE_HIVE_ALL, (HE_SID_ALL And Not HE_SID_SERVICE) Or HE_SID_NO_VIRTUAL, HE_REDIR_NO_WOW
    End If
    
    Do While HE.MoveNext
        For Each vParam In Array("TEMP", "TMP")
            
            sHit = vbNullString
            
            If HE.Hive = HKLM Then
                sKeyFull = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
            Else
                sKeyFull = HE.HiveNameAndSID & "\Environment"
            End If
            
            If Not Reg.ValueExists(0, sKeyFull, CStr(vParam)) Then
                
                bComply = False
                
                If OSver.IsElevated Then
                    bComply = True
                Else
                    'these keys are access restricted for Limited users
                    If Not (HE.Hive = HKU And (HE.SID = "S-1-5-19" Or HE.SID = "S-1-5-20")) Then
                        bComply = True
                    End If
                End If
                
                If bComply Then
                    sHit = "O7 - TroubleShooting (EV): " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = (not exist)"
                End If
            Else
                sDataNonExpanded = Reg.GetString(0, sKeyFull, CStr(vParam), , True)
                sData = EnvironW(sDataNonExpanded, , GetProfileDirBySID(HE.SID))
                
                If InStr(sData, "%") <> 0 Then
                    sHit = "O7 - TroubleShooting (EV): " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = " & sData & " (wrong type of parameter)"
                    sData = EnvironW(sData)
                ElseIf Len(sData) = 0 Then
                    sHit = "O7 - TroubleShooting (EV): " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = (empty value)"
                End If
                
                If Len(sHit) = 0 And HE.SID <> "S-1-5-18" Then
                    If Not FolderExists(sData) Then
                        sHit = "O7 - TroubleShooting: (EV) " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = " & sData & " " & STR_FOLDER_MISSING
                    End If
                End If
            End If
            
            If Len(sHit) <> 0 Then
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        
                        If HE.Hive = HKLM Then
                            sDefValue = "%SystemRoot%\TEMP"
                        Else
                            If OSver.MajorMinor < 6 Then
                                sDefValue = "%USERPROFILE%\Local Settings\Temp"
                            Else
                                sDefValue = "%USERPROFILE%\AppData\Local\Temp"
                            End If
                        End If
                        
                        If StrEndWith(sHit, STR_FOLDER_MISSING) Then
                            AddFileToFix .File, CREATE_FOLDER, sData
                            .CureType = FILE_BASED
                        Else
                            AddRegToFix .Reg, RESTORE_VALUE, 0, sKeyFull, CStr(vParam), sDefValue, REG_NOTREDIRECTED, REG_RESTORE_EXPAND_SZ
                            .CureType = REGISTRY_BASED
                        End If
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckEnvVarTemp - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckEnvVarTemp"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckEnvVarOther()
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckEnvVarOther - Begin"
    
    Dim sKeyFull As String, sHit As String, result As SCAN_RESULT, bSafe As Boolean
    Dim i As Long, sData As String, sDataNonExpanded As String, aData() As String, sDefault As String
    
    'check additional HKLM env. vars
    
    ReDim aParam(1) As String
    Dim aDefValue() As String
    Dim aFileBased() As Boolean
    ReDim aDefValue(UBound(aParam)) As String
    ReDim aFileBased(UBound(aParam)) As Boolean
    
    Dim bMissing As Boolean
    Dim bMicrosoft As Boolean
    
    aParam(0) = "ComSpec"
    aDefValue(0) = "%SystemRoot%\system32\cmd.exe"
    aFileBased(0) = True

    aParam(1) = "windir"
    aDefValue(1) = "%SystemRoot%"
    
    sKeyFull = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    
    For i = 0 To UBound(aParam)

        sData = Reg.GetString(0, sKeyFull, aParam(i))
        sDataNonExpanded = Reg.GetString(0, sKeyFull, aParam(i), , True)

        bSafe = True
        bMissing = False
        bMicrosoft = False

        If StrComp(sDataNonExpanded, aDefValue(i), 1) <> 0 Then
            bSafe = False
        End If
        
        If aFileBased(i) Then
            If Not FileExists(sData) Then
                bMissing = True
                bSafe = False
            ElseIf SignVerifyJack(sData, result.SignResult) And result.SignResult.isMicrosoftSign Then
                bMicrosoft = True
            Else
                bSafe = False
            End If
        End If
        
        If Not bSafe Then
            
            sHit = "O7 - TroubleShooting (EV): HKLM\..\Environment: " & "[" & aParam(i) & "]" & " = " & sDataNonExpanded
            
            If aFileBased(i) Then
                sHit = sHit & IIf(bMissing, " " & STR_FILE_MISSING, vbNullString) & FormatSign(result.SignResult)
            End If
            
            If aFileBased(i) Then
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
            End If
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    
                    AddRegToFix .Reg, RESTORE_VALUE, HKEY_ANY, sKeyFull, aParam(i), aDefValue(i), REG_NOTREDIRECTED, REG_RESTORE_EXPAND_SZ, , , ";"
                    .CureType = REGISTRY_BASED
                    
                    If aFileBased(i) Then
                        AddFileToFix .File, JUMP_FILE, sData
                        .CureType = .CureType Or FILE_BASED
                    End If
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    ' %PSModulePath% - according to "Missing list"
    
    If OSver.IsWindows7OrGreater Then
    
        sKeyFull = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        
        sDataNonExpanded = Reg.GetString(0, sKeyFull, "PSModulePath", , True)
        aData = Split(sDataNonExpanded, ";")
        PathRemoveLastSlashInArray aData
        
        If OSver.IsWindows10OrGreater Then
            sDefault = "%ProgramFiles%\WindowsPowerShell\Modules;%SystemRoot%\system32\WindowsPowerShell\v1.0\Modules"
        Else
            sDefault = "%SystemRoot%\system32\WindowsPowerShell\v1.0\Modules"
        End If
        aDefValue = Split(sDefault, ";")
        
        For i = 0 To UBound(aDefValue)
            If Not InArray(aDefValue(i), aData, , , vbTextCompare) Then
                
                sHit = "O7 - TroubleShooting (EV): " & "HKLM\..\Environment: " & "[PSModulePath]" & " = " & sDataNonExpanded & _
                    " (Missing: " & aDefValue(i) & ")"
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, RESTORE_VALUE, 0, sKeyFull, "PSModulePath", sDefault & ";" & sDataNonExpanded, , REG_RESTORE_EXPAND_SZ
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
        
    End If
    
    AppendErrorLogCustom "CheckEnvVarOther - Begin"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckEnvVarOther"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckEnvVarPathExt()
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckEnvVarPathExt - Begin"
    
    Dim HE As clsHiveEnum
    Dim sData As String, sDefValue As String, bSafe As Boolean, i As Long
    Dim vParam, sKeyFull As String, sHit As String, result As SCAN_RESULT
    
    ' %PATHEXT%
    
    Set HE = New clsHiveEnum
    HE.Init HE_HIVE_ALL And Not HE_HIVE_HKLM, HE_SID_USER, HE_REDIR_NO_WOW
    
    'HKCU - no value;
    'for HKLM only:
    If OSver.IsWindowsVistaOrGreater Then
        sDefValue = ".COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC"
    Else
        sDefValue = ".COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH"
    End If
    
    Dim aData() As String
    Dim aDef() As String
    aDef = Split(sDefValue, ";")
    
    'If HKCU present, it has overwrite advantage, -> check is it have %PATHEXT% or all of its required components
    'If value empty, remove it allowing HKLM to take an advantage
    'If not empty, prepend %PATHEXT%; to it
    Do While HE.MoveNext

        sKeyFull = HE.HiveNameAndSID & "\Environment"
        
        If Reg.ValueExists(0, sKeyFull, "PATHEXT") Then
        
            sData = Reg.GetString(0, sKeyFull, "PATHEXT")
            aData = Split(sData, ";")
            
            For i = 0 To UBound(aDef)
                If Not InArray(aDef(i), aData, , , vbTextCompare) Then
                
                    sHit = "O7 - TroubleShooting (EV): " & HE.HiveNameAndSID & "\..\Environment: " & "[PATHEXT]" & " = " & sData & _
                        " (Missing: " & aDef(i) & ")"
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            If Len(sData) = 0 Then
                                AddRegToFix .Reg, REMOVE_VALUE, 0, sKeyFull, "PATHEXT"
                            Else
                                AddRegToFix .Reg, RESTORE_VALUE, 0, sKeyFull, "PATHEXT", "%PATHEXT%;" & sData, , REG_RESTORE_EXPAND_SZ
                            End If
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        End If
    Loop
    
    'HKLM - by "Missing list" (show item if only it has missing records)
    
    sKeyFull = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    sData = Reg.GetString(0, sKeyFull, "PATHEXT")
    
    aData = Split(sData, ";")
    
    For i = 0 To UBound(aDef)
        If Not InArray(aDef(i), aData, , , vbTextCompare) Then
        
            sHit = "O7 - TroubleShooting (EV): " & "HKLM\..\Environment: " & "[PATHEXT]" & " = " & sData & " (Missing: " & aDef(i) & ")"
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, 0, sKeyFull, "PATHEXT", sDefValue, , REG_RESTORE_SZ
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    AppendErrorLogCustom "CheckEnvVarPathExt - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckEnvVarPathExt"
    If inIDE Then Stop: Resume Next
End Sub
    
Public Sub CheckSystemProblemsFreeSpace()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckSystemProblemsFreeSpace - Begin"
    
    Dim cFreeSpace As Currency
    Dim sHit As String
    Dim result As SCAN_RESULT
    
    cFreeSpace = cDrives.GetFreeSpace(SysDisk, False)
    ' < 1 GB ?
    If (cFreeSpace < cMath.MBToInt64(1& * 1024)) And (cFreeSpace <> 0@) Then
        
        sHit = "O7 - TroubleShooting (Disk): Free disk space on " & SysDisk & " is too low = " & (cFreeSpace / 1024& / 1024& * 10000& \ 1) & " MB."
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O7"
                .HitLineW = sHit
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
    End If
    
    AppendErrorLogCustom "CheckSystemProblemsFreeSpace - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblemsFreeSpace"
    If inIDE Then Stop: Resume Next
End Sub
    
Public Sub CheckSystemProblemsNetwork()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckSystemProblemsNetwork - Begin"
    
    Dim sNetBiosName As String
    Dim sHit As String
    Dim result As SCAN_RESULT
    
    If Len(GetCompName(ComputerNamePhysicalDnsHostname)) = 0 Then
    
        sNetBiosName = GetCompName(ComputerNameNetBIOS)
        sHit = "O7 - TroubleShooting (Network): Computer name (hostname) is not set" & IIf(Len(sNetBiosName) <> 0, " (should be: " & sNetBiosName & ")", vbNullString)
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O7"
                .HitLineW = sHit
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
    End If
    
    AppendErrorLogCustom "CheckSystemProblemsNetwork - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblemsNetwork"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckCertificatesEDS()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckCertificatesEDS - Begin"
    
    'Checking for untrusted code signing root certificates
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SystemCertificates\Disallowed\Certificates
    
    'infections examples:
    'https://blog.malwarebytes.com/cybercrime/2015/11/vonteera-adware-uses-certificates-to-disable-anti-malware/
    'https://www.bleepingcomputer.com/news/security/certlock-trojan-blocks-security-programs-by-disallowing-their-certificates/
    'https://www.securitylab.ru/news/486648.php
    
    'reverse 'blob'
    'https://namecoin.org/2017/05/27/reverse-engineering-cryptoapi-cert-blobs.html
    'https://itsme.home.xs4all.nl/projects/xda/smartphone-certificates.html
    'https://msdn.microsoft.com/en-us/library/windows/desktop/aa376079%28v=vs.85%29.aspx
    'https://msdn.microsoft.com/en-us/library/windows/desktop/aa376573%28v=vs.85%29.aspx
    'https://msdn.microsoft.com/en-us/library/cc232282.aspx
    
    Dim i&, aSubKey$(), idx&, sTitle$, bSafe As Boolean, sHit$, result As SCAN_RESULT, resultAll As SCAN_RESULT
    Dim Blob() As Byte, CertHash As String, FriendlyName As String, IssuedTo As String, nItems As Long
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum

    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\SystemCertificates\Disallowed\Certificates"

    Do While HE.MoveNext
        For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey())

            bSafe = True
            sTitle = vbNullString

            Blob = Reg.GetBinary(HE.Hive, HE.Key & "\" & aSubKey(i), "Blob")

            If AryItems(Blob) Then
                ParseCertBlob Blob, CertHash, FriendlyName, IssuedTo

                If Len(CertHash) = 0 Then CertHash = aSubKey(i)

                idx = GetCollectionIndexByKey(CertHash, colDisallowedCert)

                If idx <> 0 Then
                    'it's safe
                    If Not bHideMicrosoft Or bIgnoreAllWhitelists Then
                        sTitle = colDisallowedCert(idx)
                        bSafe = False
                    End If
                Else
                    bSafe = False
                End If

                If Not bSafe Then
                    If Len(sTitle) = 0 Then sTitle = IssuedTo
                    If Len(sTitle) = 0 Then sTitle = "Unknown"
                    If Len(FriendlyName) <> 0 Then sTitle = sTitle & " (" & FriendlyName & ")"
                    If FriendlyName = "Fraudulent" Or FriendlyName = "Untrusted" Then sTitle = sTitle & " (HJT: possible, safe)"

                    'O7 - Policy: [Untrusted Certificate] Hash - 'Name, cert. issued to' (HJT rating, if possible)
                    sHit = "O7 - Policy: [Untrusted Certificate] " & HE.HiveNameAndSID & " - " & CertHash & " - " & sTitle

                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i)
                            AddRegToFix resultAll.Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                        nItems = nItems + 1
                    End If
                End If
            End If
        Next
    Loop

    If nItems > 10 Then
        With resultAll
            .Alias = "O7"
            .HitLineW = "O7 - Policy: [Untrusted Certificate] Fix all items from the log"
            .CureType = REGISTRY_BASED
        End With
        AddToScanResults resultAll
    End If
    
    'Check for new Microsoft Root certificates

    FindNewMicrosoftCodeSignCert

    AppendErrorLogCustom "CheckCertificatesEDS - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckCertificatesEDS"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub ParseCertBlob(Blob() As Byte, out_CertHash As String, out_FriendlyName As String, out_IssuedTo As String)
    On Error GoTo ErrorHandler:
    
    'Thanks to Willem Jan Hengeveld
    'https://itsme.home.xs4all.nl/projects/xda/smartphone-certificates.html
    
    Const SHA1_HASH As Long = 3
    Const FRIENDLY_NAME As Long = 11 'Fraudulent
    
    Dim pCertContext    As Long
    'Dim CertInfo        As CERT_INFO
    Dim Prop            As CERTIFICATE_BLOB_PROPERTY
    
    Dim cStream As clsStream
    Set cStream = New clsStream
    
    out_CertHash = vbNullString
    out_FriendlyName = vbNullString
    out_IssuedTo = vbNullString
    
    'registry blob is an array of CERTIFICATE_BLOB_PROPERTY structures.
    
    cStream.WriteData VarPtr(Blob(0)), UBound(Blob) + 1
    cStream.BufferPointer = 0
    
    Do While cStream.BufferPointer < cStream.Size
        cStream.ReadData VarPtr(Prop), 12
        If Prop.Length > 0 Then
            ReDim Prop.Data(Prop.Length - 1)
            cStream.ReadData VarPtr(Prop.Data(0)), Prop.Length
            
'            Debug.Print "PropID: " & prop.PropertyID
'            Debug.Print "Length: " & prop.length
'            Debug.Print "DataA:   " & Replace(StringFromPtrA(VarPtr(prop.Data(0))), vbNullChar, "-")
'            Debug.Print "DataW:   " & StringFromPtrW(VarPtr(prop.Data(0)))
'            Debug.Print "HexData: " & GetHexStringFromArray(prop.Data)
            
            'Notice: some prop. Ids supplied with a blob in unknown encoding form, not applicable for CertCreateCertificateContext
            'e.g. CERT_ENHKEY_USAGE_PROP_ID
            
            Select Case Prop.PropertyId
            Case SHA1_HASH
                out_CertHash = GetHexStringFromArray(Prop.Data)
            Case FRIENDLY_NAME
                out_FriendlyName = StringFromPtrW(VarPtr(Prop.Data(0)))
            Case 32
                pCertContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, VarPtr(Prop.Data(0)), UBound(Prop.Data) + 1)
            
                If pCertContext <> 0 Then
                    
                    'If GetCertInfoFromCertificate(pCertContext, CertInfo) Then
                    '    out_IssuedTo = GetSignerNameFromBLOB(CertInfo.Subject)
                    'End If
                    out_IssuedTo = ExtractStringFromCertificate(pCertContext, CERT_NAME_SIMPLE_DISPLAY_TYPE)
                    
                    CertFreeCertificateContext pCertContext
                Else
                    If inIDE Then Debug.Print "CertCreateCertificateContext failed with 0x" & Hex$(Err.LastDllError)
                End If
            End Select
            If Len(out_CertHash) <> 0 And Len(out_FriendlyName) <> 0 And Len(out_IssuedTo) <> 0 Then Exit Do
        End If
    Loop
    
    'If out_IssuedTo = vbnullstring Then Debug.Print "No SubjectName for cert: " & out_CertHash
    
    Set cStream = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ParseCertBlob"
    If inIDE Then Stop: Resume Next
End Sub

Private Function GetHexStringFromArray(a() As Byte) As String
    Dim sHex As String
    Dim i As Long
    For i = 0 To UBound(a)
        sHex = sHex & Right$("0" & Hex$(a(i)), 2)
    Next
    GetHexStringFromArray = sHex
End Function

Private Function HexStringToArray(sHexStr As String) As Byte()
    Dim i As Long
    Dim b() As Byte
    
    ReDim b(Len(sHexStr) \ 2 - 1)
    
    For i = 1 To Len(sHexStr) Step 2
        b((i - 1) \ 2) = CLng("&H" & mid$(sHexStr, i, 2))
    Next
    
    HexStringToArray = b
End Function

Public Sub CheckPolicyACL()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckPolicyACL - Begin"
    
    Dim result As SCAN_RESULT
    Dim i As Long
    Dim SDDL As String, sHit As String
    
    Dim aKey(3) As String
    aKey(0) = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies"
    aKey(1) = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft"
    aKey(2) = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\SystemCertificates"
    aKey(3) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SystemCertificates"
    
    For i = 0 To UBound(aKey)
        If Not CheckKeyAccess(0, aKey(i), KEY_READ) Then
            SDDL = GetKeyStringSD(0, aKey(i))
            
            sHit = "O7 - Policy: Permissions on key are restricted - " & aKey(i) & " - " & SDDL
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    '// TODO: improve it to use correct default permission (no write access for Admin. group / no propagate).
                    AddRegToFix .Reg, RESTORE_KEY_PERMISSIONS_RECURSE, 0, aKey(i)
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    AppendErrorLogCustom "CheckPolicyACL - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckPolicyACL"
    If inIDE Then Stop: Resume Next
End Sub

Sub CheckCredentials()

    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckCredentials - Begin"
    
    Dim sHit$, result As SCAN_RESULT
    Dim lData As Long, sValue As String
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    'Checking for plain login/password usage
    HE.AddKey "System\CurrentControlSet\Control\SecurityProviders\WDigest"
    sValue = "UseLogonCredential"

    Do While HE.MoveNext
        lData = Reg.GetDword(HE.Hive, HE.Key, sValue, HE.Redirected)
        If lData <> 0 Then
            sHit = BitPrefix("O7", HE) & " - Policy: " & HE.KeyAndHivePhysical & ": " & "[" & sValue & "] = " & lData
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sValue, , HE.Redirected
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Loop

    AppendErrorLogCustom "CheckCredentials - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckCredentials"
    If inIDE Then Stop: Resume Next
End Sub

Sub CheckPolicyScripts()
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckPolicyScripts - Begin"
    
    '
    'For quick overview:
    'gpresult.exe /v - console output of policy scrpits (these doesn't include "Local PC\User" scripts !!! )
    '
    'Keys for analysis:
    '
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Logon\*\*
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Logoff\*\*
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\<SID>\Scripts\Logon\*\*
    'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\State\<SID>\Scripts\Logon\*\*
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\*\*
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\*\*
    'C:\Windows\System32\GroupPolicy\User\Scripts\scripts.ini
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\scripts.ini
    'C:\Windows\System32\GroupPolicy\User\Scripts\psscripts.ini
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\psscripts.ini
    
    'Notice for HKCU:
    '
    'Always represented as:
    ' - Logon\X\Y
    ' - Logoff\X\Y
    '
    'where X defined as:
    '0 - HKCU policy for Local PC (user config)
    '1 - HKCU policy for Local PC\User (user config)*

    ' * How to setup user-specific group policies:
    'https://www.tenforums.com/tutorials/80043-apply-local-group-policy-specific-user-windows-10-a.html
    
    'Notice for HKLM:
    '
    'Always represented as:
    ' - Startup\X\Y
    ' - Shutdown\X\Y
    '
    'where X defined as:
    '0 - HKCU policy for Local PC (machine config)
    '
    'Y - is index number of script record, starting from 0. They are always consecutive*.
    '* non-consecutive indeces brake the chain!
    'When HJT fix the item, it is required to reconstruct the whole chain, so other items become valid.
    
    'Both cases include mirrors in C:\Windows\System32\GroupPolicy location ini-file.
    '* for "Local PC\User" the mirror is located under: C:\Windows\System32\GroupPolicyUsers\<SID>
    
    Dim sHit$, result As SCAN_RESULT
    Dim pos As Long, x As Long, y As Long, i As Long
    Dim vType As Variant, vFile As Variant
    Dim sFile$, sArgs$, sAlias$, sHash$, aKeyX$(), aKeyY$(), sKey$, aFiles$(), sIniPath$, sIniPathPS$, sFileSysPath$
    Dim oFiles As clsTrickHashTable
    Set oFiles = New clsTrickHashTable
    oFiles.CompareMode = 1
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    For Each vType In Array("Logon", "Logoff", "Startup", "Shutdown")
        HE.Init HE_HIVE_ALL
        HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\" & vType

        Do While HE.MoveNext

            For x = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aKeyX(), HE.Redirected, False, False, True)

                sFileSysPath = Reg.GetString(HE.Hive, HE.Key & "\" & aKeyX(x), "FileSysPath")
                
                sIniPath = EnvironW(sFileSysPath)
                sIniPath = sFileSysPath & "\Scripts\scripts.ini"
                sIniPathPS = sFileSysPath & "\Scripts\psscripts.ini"
                
                For y = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key & "\" & aKeyX(x), aKeyY(), HE.Redirected, False, False, True)

                    sKey = HE.Key & "\" & aKeyX(x) & "\" & aKeyY(y)

                    If Reg.ValueExists(HE.Hive, sKey, "Script", HE.Redirected) Then

                        sFile = Reg.GetString(HE.Hive, sKey, "Script", HE.Redirected)
                        sArgs = Reg.GetString(HE.Hive, sKey, "Parameters", HE.Redirected)

                        If InStr(sFile, ":") = 0 Then 'relative to script storage?
                            sFile = BuildPath(sFileSysPath, "Scripts", vType, sFile)
                        End If

                        sAlias = BitPrefix("O7", HE) & " - Policy Script: "

                        sFile = FormatFileMissing(sFile)

                        If Not oFiles.Exists(sFile) Then oFiles.Add sFile, 0&
                        
                        SignVerifyJack sFile, result.SignResult
                        
                        sHit = sAlias & HE.HiveNameAndSID & "\..\Group Policy\Scripts\" & vType & "\" & aKeyX(x) & "\" & aKeyY(y) & _
                            ": [" & "Script" & "] = " & ConcatFileArg(sFile, sArgs) & FormatSign(result.SignResult)
                        
                        If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash

                        If Not IsOnIgnoreList(sHit) Then

                            With result
                                .Section = "O7"
                                .HitLineW = sHit
                                .Alias = sAlias
                                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, sKey, , , HE.Redirected
                                
                                If 1 = Reg.GetDword(HE.Hive, sKey, "IsPowershell", HE.Redirected) Then
                                    
                                    AddIniToFix .Reg, REMOVE_VALUE_INI, sIniPathPS, vType, aKeyY(y) & "CmdLine"
                                    AddIniToFix .Reg, REMOVE_VALUE_INI, sIniPathPS, vType, aKeyY(y) & "Parameters"
                                Else
                                    AddIniToFix .Reg, REMOVE_VALUE_INI, sIniPath, vType, aKeyY(y) & "CmdLine"
                                    AddIniToFix .Reg, REMOVE_VALUE_INI, sIniPath, vType, aKeyY(y) & "Parameters"
                                End If
                                
                                'remove state data
                                If vType = "Startup" Or vType = "Shutdown" Then
                                    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Startup\0\0
                                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\" & vType & "\" & aKeyX(x) & "\" & aKeyY(y), , , HE.Redirected
                                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\" & vType & "\" & aKeyX(x) & "\" & aKeyY(y), , , HE.Redirected
                                
                                Else ' vType = "Logon" Or vType = "Logoff" Then
                                    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\S-1-5-21-4161311594-4244952198-1204953518-1000\Scripts\Logon\0\0
                                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\" & HE.SID & "\Scripts\" & vType & "\" & aKeyX(x) & "\" & aKeyY(y), , , HE.Redirected
                                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\State\" & HE.SID & "\Scripts\" & vType & "\" & aKeyX(x) & "\" & aKeyY(y), , , HE.Redirected
                                End If
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                                .CureType = FILE_BASED Or REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If

                Next
            Next
        Loop
    Next
    
    'Checking files inside script folders
    '
    'C:\Windows\System32\GroupPolicy\User\Scripts\Logon
    'C:\Windows\System32\GroupPolicy\User\Scripts\Logoff
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\Startup
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown
    'C:\Windows\System32\GroupPolicyUsers\<SID>\User\Scripts\Logoff
    'C:\Windows\System32\GroupPolicyUsers\<SID>\User\Scripts\Logon
    
    For Each vType In Array("User\Scripts\Logon", "User\Scripts\Logoff", "Machine\Scripts\Startup", "Machine\Scripts\Shutdown")
    
        aFiles = ListFiles(BuildPath(sWinSysDir, "GroupPolicy", vType))
        If AryItems(aFiles) Then
            For i = 0 To UBound(aFiles)
                sFile = aFiles(i)
                If Not oFiles.Exists(sFile) Then 'don't duplicate registry records scan results
                    
                    sAlias = "O7 - Policy Script: "
                    SignVerifyJack sFile, result.SignResult
                    sHit = sAlias & sFile & FormatSign(result.SignResult)
                    
                    If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            .Alias = sAlias
                            AddFileToFix .File, REMOVE_FILE, sFile
                            .CureType = FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        End If
    Next
    
    'C:\Windows\System32\GroupPolicy\User\Scripts\Scripts.ini
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\Scripts.ini
    'C:\Windows\System32\GroupPolicyUsers\<SID>\User\Scripts\Scripts.ini
    'C:\Windows\System32\GroupPolicy\User\Scripts\psscripts.ini
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\psscripts.ini
    'C:\Windows\System32\GroupPolicyUsers\<SID>\User\Scripts\psscripts.ini
    '
    '[Logon]
    '0CmdLine=C:\Users\Alex\Desktop\Alex1.ps1
    '0Parameters=-hi_there
    '1CmdLine=""C:\Users\Alex\Desktop\Alex2.ps1""
    '1Parameters=
    
    Dim cIni As clsIniFile
    Dim nSID As Long, nItem As Long
    Dim sIni As String, vIniName
    Dim aSections(), aSection, aNames(), aName
    
    For Each vIniName In Array("scripts.ini", "psscripts.ini")
    
        For Each vFile In Array( _
            BuildPath(sWinSysDir, "GroupPolicy\User\Scripts\"), _
            BuildPath(sWinSysDir, "GroupPolicy\Machine\Scripts\"), _
            "<SID>")
            
            nSID = LBound(gSID_All)
            Do
                If vFile = "<SID>" Then
                    sIni = BuildPath(sWinSysDir, "GroupPolicyUsers", gSID_All(nSID), "User\Scripts\", vIniName)
                    nSID = nSID + 1
                Else
                    sIni = vFile & vIniName
                End If
                
                If FileExists(sIni) Then
                
                    Set cIni = New clsIniFile
                    cIni.InitFile sIni, 0
                    
                    aSections = cIni.GetSections()
                    
                    For Each aSection In aSections()
                    
                        aNames = cIni.GetParamNames(aSection)
                        
                        For Each aName In aNames()
                        
                            pos = InStr(2, aName, "CmdLine", 1)
                            
                            If pos <> 0 Then
                            
                                If IsNumeric(Left$(aName, pos - 1)) Then
                                
                                    nItem = CLng(Left$(aName, pos - 1))
                                
                                    sFile = cIni.ReadParam(aSection, nItem & "CmdLine")
                                    sArgs = cIni.ReadParam(aSection, nItem & "Parameters")
                                    
                                    sFile = UnQuote(Replace$(sFile, """""", """"))
                                    
                                    If InStr(sFile, ":") = 0 Then 'relative to script storage?
                                        sFile = BuildPath(GetParentDir(sIni), aSection, sFile)
                                    End If
                                    
                                    sFile = FormatFileMissing(sFile)
                                    
                                    If Not oFiles.Exists(sFile) Then 'no such file in registry entries scan
                                        
                                        sAlias = "O7 - Policy Script: "
                                        SignVerifyJack sFile, result.SignResult
                                        
                                        sHit = sAlias & sIni & ": " & "[" & aSection & "] " & aName & " = " & ConcatFileArg(sFile, sArgs) & _
                                            FormatSign(result.SignResult)
                                        
                                        If g_bCheckSum Then sHash = GetFileCheckSum(sFile): sHit = sHit & sHash
                                        
                                        If Not IsOnIgnoreList(sHit) Then
                                            With result
                                                .Section = "O7"
                                                .HitLineW = sHit
                                                .Alias = sAlias
                                                AddIniToFix .Reg, REMOVE_VALUE_INI, sIni, aSection, aName
                                                AddIniToFix .Reg, REMOVE_VALUE_INI, sIni, aSection, nItem & "Parameters"
                                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                                                .CureType = FILE_BASED Or INI_BASED
                                            End With
                                            AddToScanResults result
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                    
                End If
                
            Loop Until vFile <> "<SID>" Or nSID > UBound(gSID_All)
        Next
    Next
    
    Set oFiles = Nothing
    
    AppendErrorLogCustom "CheckPolicyScripts - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckPolicyScripts"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub PolicyScripts_RebuildChain()

    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "PolicyScripts_RebuildChain - Begin"
    
    'rebuild registry chain
    
    Dim vType, vFile
    Dim sKey As String, sIniPath As String, sName As String
    Dim aFiles() As String, aKeyX() As String, aKeyY() As String
    Dim x As Long, y As Long, i As Long, idx As Long
    
    Dim oFiles As clsTrickHashTable
    Set oFiles = New clsTrickHashTable
    oFiles.CompareMode = 1
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    For Each vType In Array("Logon", "Logoff", "Startup", "Shutdown")
        HE.Init HE_HIVE_ALL
        HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\" & vType
        
        Do While HE.MoveNext
        
            For x = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aKeyX(), HE.Redirected, False, False, True)
                
                sIniPath = Reg.GetString(HE.Hive, HE.Key & "\" & aKeyX(x), "FileSysPath") & "\Scripts\scripts.ini"
                sIniPath = EnvironW(sIniPath)
                
                If Not oFiles.Exists(sIniPath) Then oFiles.Add sIniPath, 0&
                
                idx = 0
                
                For y = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key & "\" & aKeyX(x), aKeyY(), HE.Redirected, False, False, True)
                
                    If aKeyY(y) <> idx Then
                        
                        sKey = HE.Key & "\" & aKeyX(x) & "\" & aKeyY(y)
                        
                        Reg.RenameKey HE.Hive, sKey, CStr(idx), HE.Redirected
                    End If
                    
                    idx = idx + 1
                Next
            Next
        Loop
    Next
    
    'rebuild ini chain
    
    Dim cIni As clsIniFile
    
    'preparing list of ini files
    aFiles = ListFiles(BuildPath(sWinSysDir, "GroupPolicyUsers"), ".ini", True)
    
    If AryItems(aFiles) Then
        ReDim Preserve aFiles(UBound(aFiles) + 4)
    Else
        ReDim aFiles(3)
    End If
    aFiles(UBound(aFiles) - 3) = BuildPath(sWinSysDir, "GroupPolicy\User\Scripts\psscripts.ini")
    aFiles(UBound(aFiles) - 2) = BuildPath(sWinSysDir, "GroupPolicy\User\Scripts\scripts.ini")
    aFiles(UBound(aFiles) - 1) = BuildPath(sWinSysDir, "GroupPolicy\Machine\Scripts\psscripts.ini")
    aFiles(UBound(aFiles) - 0) = BuildPath(sWinSysDir, "GroupPolicy\Machine\Scripts\scripts.ini")
    For i = 0 To UBound(aFiles)
    
        sName = GetFileName(aFiles(i), True)
        
        If StrComp(sName, "scripts.ini", 1) = 0 _
            Or StrComp(sName, "psscripts.ini", 1) = 0 Then
        
            'append to files pointed by registry entries
            If Not oFiles.Exists(aFiles(i)) Then oFiles.Add aFiles(i), 0&
        End If
    Next
    
    Dim aParam() As Variant
    Dim iNum As Long
    
    For Each vFile In oFiles.Keys
        
        If FileExists(vFile) Then
        
            Set cIni = New clsIniFile
            cIni.InitFile CStr(vFile), 0
            
            'Debug.Print vFile
            
            For Each vType In Array("Logon", "Logoff", "Startup", "Shutdown")
                
                idx = -1
                
                aParam = cIni.GetParamNames(vType)
                
                For x = 0 To UBoundSafe(aParam)
                    
                    sName = mid$(aParam(x), 2)
                    
                    If StrComp(sName, "CmdLine", 1) = 0 Then
                        
                        If IsNumeric(Left$(aParam(x), 1)) Then
                            
                            idx = idx + 1
                            iNum = CLng(Left$(aParam(x), 1))
                            
                            If iNum <> idx Then
                                cIni.RenameParam vType, CStr(iNum) & "CmdLine", CStr(idx) & "CmdLine"
                                cIni.RenameParam vType, CStr(iNum) & "Parameters", CStr(idx) & "Parameters"
                            End If
                        End If
                    End If
                    
                    'Debug.Print aParam(X) & "  =  " & cIni.ReadParam(vType, aParam(X))
                Next
            Next
            
            Set cIni = Nothing
            
        End If
    Next
    
    Set oFiles = Nothing
    
    AppendErrorLogCustom "PolicyScripts_RebuildChain - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "PolicyScripts_RebuildChain"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckPolicies()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckPolicies - Begin"
    
    '//TODO:
    'HKEY_CURRENT_USER\Software\Policies\Microsoft
    
    Dim sDrv As String, aValue() As String, i&, lData&, bData() As Byte
    Dim sHit$, result As SCAN_RESULT
    Dim sKey As String, sValue As String
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\System" 'Shared
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\LocalUser\Software\Microsoft\Windows\CurrentVersion\Policies\System" 'Shared
    HE.AddKey "Software\Microsoft\Windows NT\CurrentVersion\Winlogon" 'WOW
    
    'DisableRegistryTools|DisableTaskMgr
    aValue = Split(Caes_Decode("ElxhkwravzDPSS\sVXW`|o\hX[gbSbvpTpC"), "|")
    
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i), HE.Redirected)
            If lData <> 0 Then
                sHit = BitPrefix("O7", HE) & " - Policy: " & HE.KeyAndHivePhysical & ": " & "[" & aValue(i) & "] = " & lData
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i), , HE.Redirected
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    'Taskbar policies
    '��. �������� �. �������� ������� Windows Vista. ����� � �������.
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" 'Shared
    
    aValue = Split("NoSetTaskbar|TaskbarLockAll|LockTaskbar|NoTrayItemsDisplay|NoChangeStartMenu|NoStartMenuMorePrograms|NoRun|" & _
        "NoSMConfigurePrograms|NoViewOnDrive|RestrictRun|DisallowRun|NoControlPanel|NoDispCpl|SettingsPageVisibility|NoViewContextMenu|" & _
        "NoSecurityTab", "|")
    
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i))
            If lData <> 0 Then
                sHit = "O7 - Taskbar policy: " & HE.HiveNameAndSID & "\..\Policies\Explorer: [" & aValue(i) & "] = " & lData
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    
    'hide drives in My PC window
    'https://support.microsoft.com/en-us/help/555438
    HE.Repeat
    Do While HE.MoveNext
        If Reg.ValueExists(HE.Hive, HE.Key, "NoDrives") Then
            bData = Reg.GetBinary(HE.Hive, HE.Key, "NoDrives")
            If AryItems(bData) Then
                GetMem4 bData(0), lData
                For i = 65 To 90
                    If lData And (2 ^ (i - 65)) Then
                        sDrv = sDrv & Chr$(i) & ", "
                    End If
                Next
                If Len(sDrv) <> 0 Then
                    sDrv = Left$(sDrv, Len(sDrv) - 2)
                    sHit = "O7 - Explorer Policy: " & HE.HiveNameAndSID & "\..\Policies\Explorer: [NoDrives] = 0x" & ByteArrayToHex(bData) & _
                        " (Disk: " & sDrv & ")"
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            .Reboot = True
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "NoDrives"
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
        End If
    Loop
    
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" 'WOW
    
    aValue = Split("Start_ShowRun", "|")
    
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i), HE.Redirected)
            If lData = 0 And Reg.StatusSuccess Then
                sHit = BitPrefix("O7", HE) & " - Taskbar policy: " & HE.HiveNameAndSID & "\..\Explorer\Advanced: [" & aValue(i) & "] = " & lData
            
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i), , HE.Redirected
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    'https://blog.malwarebytes.com/detections/pum-optional-disallowrun/
    Dim iEnabled As Long
    Dim sData As String
    Dim resultAll As SCAN_RESULT
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun" 'Shared
    
    Do While HE.MoveNext
    
        iEnabled = Reg.GetDword(0, HE.HiveNameAndSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowRun", HE.Redirected)
        
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue(), HE.Redirected)
   
            sData = Reg.GetData(HE.Hive, HE.Key, aValue(i), HE.Redirected)
            
            sHit = "O7 - Policy: " & HE.HiveNameAndSID & "\..\Policies\Explorer\DisallowRun: " & IIf(Len(aValue(i)) = 1, " ", vbNullString) & _
                "[" & aValue(i) & "] = " & sData & IIf(iEnabled = 0, " (disabled)", vbNullString)
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i), , HE.Redirected
                    .CureType = REGISTRY_BASED
                End With
                ConcatScanResults resultAll, result
                AddToScanResults result
            End If
        Next
        If iEnabled <> 0 Then 'for fix all
            AddRegToFix resultAll.Reg, RESTORE_VALUE, 0, HE.HiveNameAndSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowRun", 0, HE.Redirected, REG_RESTORE_SAME
        End If
    Loop
    If resultAll.CureType <> 0 Then
        resultAll.HitLineW = "O7 - Policy: *\..\Policies\Explorer\DisallowRun: Fix all"
        resultAll.FixAll = True
        resultAll.CureType = resultAll.CureType Or CUSTOM_BASED
        AddRegToFix resultAll.Reg, REMOVE_KEY, HKLM, "SOFTWARE\Policies\Microsoft\Windows\SrpV2"
        AddRegToFix resultAll.Reg, REMOVE_KEY, HKCU, "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy Objects"
        AddCustomToFix resultAll.Custom, CUSTOM_ACTION_APPLOCKER
        AddToScanResults resultAll
    End If
    
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "SOFTWARE\Policies\Microsoft\Windows\Explorer"
    
    Do While HE.MoveNext
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue(), HE.Redirected)
   
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i), HE.Redirected)
            If lData <> 0 Then
                sHit = BitPrefix("O7", HE) & " - Policy: " & HE.HiveNameAndSID & "\..\Windows\Explorer: [" & aValue(i) & "] = " & lData
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i), , HE.Redirected
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    
    'Check Windows Defender policies
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey Caes_Decode("TrkAFlEtm`DzQPVTM]GDX_Wdnl AdghsknCiavtG-mJPJ s]\cVVi`hi") ' "Software\Microsoft\Windows Defender\Real-Time Protection"
    aValue = Split(Caes_Decode("EsfKrDnqCxy|]JVFIUPyTR_i`f`JnolyvAtAv"), "|") '"DpaDisabled|DisableRealtimeMonitoring"
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i))
            If lData <> 0 Then
                sHit = "O7 - Policy: " & HE.KeyAndHivePhysical & ": " & "[" & aValue(i) & "] = " & lData
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                        FixWindowsDefender result
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey Caes_Decode("TrkAFlEtm`DzQPVTM]GDX_Wdnl AdghsknC") '"Software\Microsoft\Windows Defender"
    HE.AddKey Caes_Decode("TrkAFlEtmcJIHDLJZErVRcbhf_oYVjqivFD SvyzKCFU") '"Software\Policies\Microsoft\Windows Defender"
    aValue = Split(Caes_Decode("ElxhkwrPEMDjOZZFYN|kXdTWc^vksjYnyDD"), "|") '"DisableAntiSpyware|DisableAntiVirus"
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i))
            If lData <> 0 Then
                sHit = "O7 - Policy: " & HE.KeyAndHivePhysical & ": " & "[" & aValue(i) & "] = " & lData
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                        FixWindowsDefender result
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    If OSver.IsWindows10OrGreater And OSver.Build >= 18362 Then
        sHit = vbNullString
        sKey = Caes_Decode("INQTe^BuKPvODwjNJ[Z`^WgQNbianxv KnqrCuxMs_FDY\[P`") '"HKLM\Software\Microsoft\Windows Defender\Features"
        sValue = Caes_Decode("UdrwnC]GFMzzSJRS") '"TamperProtection"
        lData = Reg.GetDword(HKEY_ANY, sKey, sValue)
        If Reg.StatusCode <> ERROR_SUCCESS Then
            sHit = "O7 - Policy: " & sKey & ": [" & sValue & "] = " & Reg.StatusCodeDesc()
        ElseIf (lData And 1) = 0 Then
            sHit = "O7 - Policy: " & sKey & ": [" & sValue & "] = " & lData
        End If
        If Len(sHit) <> 0 Then
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HKEY_ANY, sKey, sValue, 5, , REG_RESTORE_DWORD
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    End If
    
    AppendErrorLogCustom "CheckPolicies - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckPolicies"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub FixWindowsDefender(result As SCAN_RESULT)
    With result
        AddRegToFix .Reg, CREATE_KEY, HKEY_LOCAL_MACHINE, "Software\Microsoft\AMSI\Providers\{2781761E-28E0-4109-99FE-B9D127C57AFE}"
        AddRegToFix .Reg, CREATE_KEY, HKEY_LOCAL_MACHINE, "Software\Microsoft\AMSI\Providers2\{2781761E-28E0-4109-99FE-B9D127C57AFE}"
        AddRegToFix .Reg, CREATE_KEY, HKEY_LOCAL_MACHINE, "Software\Microsoft\AMSI\UacProviders\{2781761E-28E2-4109-99FE-B9D127C57AFE}"
        'SOFTWARE\Microsoft\Windows Defender\Spynet (Cloud-delivered protection)
        AddRegToFix .Reg, RESTORE_VALUE, HKEY_LOCAL_MACHINE, Caes_Decode("TRK[`L_Tm`DzQPVTM]GDX_Wdnl AdghsknCibGRIBS"), "SpyNetReporting", 2
        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Policies\Microsoft\" & STR_CONST.WINDOWS_DEFENDER
        AddRegToFix .Reg, REMOVE_KEY, HKCU, "SOFTWARE\Policies\Microsoft\" & STR_CONST.WINDOWS_DEFENDER
        AddServiceToFix .Service, ENABLE_SERVICE Or START_SERVICE, "WinDefend"
        AddTaskToFix .Task, ENABLE_TASK, "\Microsoft\Windows\ExploitGuard\ExploitGuard MDM policy Refresh"
        AddCommandlineToFix .CommandLine, COMMANDLINE_POWERSHELL, , "Set-MpPreference -UILockdown 0", , False
        AddCommandlineToFix .CommandLine, COMMANDLINE_POWERSHELL, , "Set-MpPreference -DisableRealtimeMonitoring $false", , False
        AddCommandlineToFix .CommandLine, COMMANDLINE_RUN, BuildPath(PF_64, STR_CONST.WINDOWS_DEFENDER, "mpcmdrun.exe"), "-wdenable", SW_MINIMIZE, False
        .CureType = REGISTRY_BASED Or SERVICE_BASED Or TASK_BASED
        '// TODO: restore tasks
        .Reboot = True
    End With
End Sub

Public Sub CheckPolicyUAC()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckPolicyUAC - Begin"

    'http://www.oszone.net/11424
    
    If Not OSver.IsWindowsVistaOrGreater Then Exit Sub

    Dim lData&
    Dim sHit$, result As SCAN_RESULT
    Dim HE As clsHiveEnum:      Set HE = New clsHiveEnum
    Dim DC As clsDataChecker:   Set DC = New clsDataChecker
    
    HE.Init HE_HIVE_HKLM, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" 'key - x64 Shared
    
    DC.AddValueData "ConsentPromptBehaviorAdmin", Array(2, 5)
    DC.AddValueData "ConsentPromptBehaviorUser", IIf(OSver.IsWindowsVista, 1, 3)
    DC.AddValueData "EnableLUA", 1
    DC.AddValueData "PromptOnSecureDesktop", 1
    DC.AddValueData "EnableUIADesktopToggle", 0
    
    'Allow to be user defined:
    '
    'EnableInstallerDetection
    'EnableSecureUIAPaths
    'ValidateAdminCodeSignatures
    'EnableVirtualization
    'FilterAdministratorToken
    
    Do While HE.MoveNext
        Do While DC.MoveNext
            lData = Reg.GetDword(HE.Hive, HE.Key, DC.ValueName, HE.Redirected)
            
            If Not DC.ContainsData(lData) Then
                sHit = "O7 - Policy: (UAC) " & HE.KeyAndHivePhysical & ": " & "[" & DC.ValueName & "] = " & Reg.StatusCodeDescOnFail(lData)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, DC.ValueName, DC.DataLong, HE.Redirected, REG_RESTORE_DWORD
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Loop
    Loop
    
    AppendErrorLogCustom "CheckPolicyUAC - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckPolicyUAC"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckAppLocker()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckAppLocker - Begin"
    
    ' Details:
    ' secpol.msc => Application Control Policies
    ' http://www.oszone.net/11303/AppLocker
    ' https://oddvar.moe/2019/02/01/bypassing-applocker-as-an-admin/
    ' https://docs.microsoft.com/en-US/windows/security/threat-protection/windows-defender-application-control/applocker/delete-an-applocker-rule
    ' https://docs.microsoft.com/ru-ru/windows/security/threat-protection/windows-defender-application-control/configure-authorized-apps-deployed-with-a-managed-installer
    '
    ' Affected keys:
    ' HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\{GUID}Machine\Software\Policies\Microsoft\Windows\SrpV2
    ' HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\{GUID}User\Software\Policies\Microsoft\Windows\SrpV2
    ' HKLM\SOFTWARE\Policies\Microsoft\Windows\SrpV2
    ' HKLM\SYSTEM\CurrentControlSet\Control\Srp\Gp
    ' HKLM\SYSTEM\ControlSet*\Control\Srp\Gp
    '
    ' Affected files:
    ' c:\windows\system32\applocker\*
    ' C:\windows\system32\GroupPolicy\*
    ' c:\users\public\ntuser.pol
    ' - AppCache.dat
    ' - ???
    '
    ' Default keys state:
    ' - No key: HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects
    ' - No key: HKLM\SOFTWARE\Policies\Microsoft\Windows\SrpV2
    ' - No Subkeys in: HKLM\SYSTEM\CurrentControlSet\Control\Srp\Gp
    
    ' Enforcement param:
    ' Name: "EnforcementMode"
    ' Location:
    ' 1) HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\{GUID}Machine\Software\Policies\Microsoft\Windows\SrpV2\{type}
    ' 2) HKLM\SOFTWARE\Policies\Microsoft\Windows\SrpV2\{type}
    ' where {type} key name is:
    ' - "Exe"
    ' - "Msi"
    ' - "Script"
    ' - "Dll"
    ' - "Appx" (Windows 10)
    ' - "ManagedInstaller" (Windows 10)
    '
    ' Values (mode):
    ' 0 - audit only
    ' 1 - enforce (implicit deny)
    ' No value - same as "enforce" - by default
    
    ' Rule example:
    '
    ' HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\
    '    {GUID}Machine\Software\Policies\Microsoft\Windows\SrpV2\Exe\{GUID} (without brackets) => "Value" => xml data
    '
    ' Same record in:
    ' HKLM\SOFTWARE\Policies\Microsoft\Windows\SrpV2\Exe\{GUID} (without brackets) => "Value" => xml data
    '
    ' xml data example:
    ' <FilePathRule
    '   Id="921cc481-6e17-4653-8f75-050b80acca20"
    '   Name="(..text.. &quot;Program Files&quot;"
    '   Description="..text.. &quot;Program Files&quot;."
    '   UserOrGroupSid="S-1-1-0"
    '   Action="Allow"
    ' >
    '   <Conditions>
    '       <FilePathCondition
    '           Path="%PROGRAMFILES%\*"
    '       />
    '   </Conditions>
    ' </FilePathRule>
    '
    ' HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Srp\Gp\Exe\{GUID} (without brackets) =>
    ' param "ACE" (binary)
    ' param "Name" (rule name)
    ' param "SDDL" => D:(XA;;FX;;;S-1-1-0;(APPID://PATH Contains "%PROGRAMFILES%\*"))
    
    If Not OSver.IsWindows7OrGreater Then Exit Sub
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim aGpoKeys() As String
    Dim aRuleKeys() As String
    Dim vType As Variant
    Dim sKey As String
    Dim xmlData As String
    Dim bNewRule As Boolean
    Dim sHit As String
    Dim result As SCAN_RESULT
    Dim resultAll As SCAN_RESULT
    Dim aFiles() As String
    
    aFiles = ListFiles(BuildPath(sWinSysDir, "AppLocker"), vbNullString, True)
    
    Dim dRuleXml As clsTrickHashTable
    Dim dRuleKey As clsTrickHashTable
    
    Set dRuleXml = New clsTrickHashTable
    Set dRuleKey = New clsTrickHashTable
    
    dRuleXml.CompareMode = TextCompare
    dRuleKey.CompareMode = TextCompare
    
    'HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\{GUID}*\Software\Policies\Microsoft\Windows\SrpV2\{type}\{CLSID}
    
    For i = 1 To Reg.EnumSubKeysToArray(HKCU, "Software\Microsoft\Windows\CurrentVersion\Group Policy Objects", aGpoKeys(), bRecursively:=False, bEraseArray:=True)
        
        For Each vType In Array("Exe", "Msi", "Script", "Dll", "Appx", "ManagedInstaller")
            
            sKey = "Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\" & _
                aGpoKeys(i) & "\Software\Policies\Microsoft\Windows\SrpV2\" & vType
            
            For k = 1 To Reg.EnumSubKeysToArray(HKCU, sKey, aRuleKeys(), bRecursively:=False, bEraseArray:=True)
     
                xmlData = Reg.GetString(HKCU, sKey & "\" & aRuleKeys(k), "Value")

                If Len(xmlData) <> 0 Then
                    
                    If Not dRuleXml.Exists(aRuleKeys(k)) Then dRuleXml.Add aRuleKeys(k), xmlData
                    If Not dRuleKey.Exists(aRuleKeys(k)) Then dRuleKey.Add aRuleKeys(k), "HKCU\" & sKey & "\" & aRuleKeys(k)
                End If
            Next
        Next
    Next
    
    'HKLM\SOFTWARE\Policies\Microsoft\Windows\SrpV2\{type}\{CLSID}
    
    For Each vType In Array("Exe", "Msi", "Script", "Dll", "Appx", "ManagedInstaller")
        
        sKey = "SOFTWARE\Policies\Microsoft\Windows\SrpV2\" & vType
        
        For k = 1 To Reg.EnumSubKeysToArray(HKLM, sKey, aRuleKeys(), bRecursively:=False, bEraseArray:=True)
    
            xmlData = Reg.GetString(HKLM, sKey & "\" & aRuleKeys(k), "Value")

            If Len(xmlData) <> 0 Then
                bNewRule = False
                
                If dRuleXml.Exists(aRuleKeys(k)) Then
                    
                    If dRuleXml(aRuleKeys(k)) <> xmlData Then 'rule isn't identical to HKCU ?
                    
                        bNewRule = True
                    End If
                Else
                    bNewRule = True
                End If
                
                If bNewRule Then
                    If Not dRuleXml.Exists(aRuleKeys(k)) Then dRuleXml.Add aRuleKeys(k), xmlData
                    If Not dRuleKey.Exists(aRuleKeys(k)) Then dRuleKey.Add aRuleKeys(k), "HKLM\" & sKey & "\" & aRuleKeys(k)
                    
                    'Debug.Print xmlData
                End If
            End If
        Next
    Next
    
    Dim sRuleId         As String
    Dim sFilePath       As String
    Dim sPublisherName  As String
    Dim sHitTemplate    As String
    Dim bDeny           As Boolean
    Dim bException      As Boolean
    Dim numEntries      As Long
    Dim eRuleType       As APPLOCKER_RULE_TYPE
    Dim tHashRuleData() As APPLOCKER_HASH_RULE_DATA
    
    Dim xmlDoc          As XMLDocument
    Dim xmlElement      As CXmlElement
    Dim xmlSubNode      As CXmlElement
    Dim xmlAttribute    As CXmlAttribute
    
    Set xmlDoc = New XMLDocument
    
    For i = 0 To dRuleXml.Count - 1
        
        bDeny = False
        eRuleType = APPLOCKER_RULE_UNKNOWN
        sFilePath = vbNullString
        sPublisherName = vbNullString
        
        sRuleId = dRuleXml.Keys(i)
        xmlData = dRuleXml.Items(i)
        sKey = dRuleKey.Items(i)
        
        Call xmlDoc.LoadData(xmlData)
        
        'Debug.Print xmlDoc.Root.Name 'FilePathRule / FileHashRule / FilePublisherRule
        
        For k = 1 To xmlDoc.Root.AttributeCount
            
            Set xmlAttribute = xmlDoc.Root.ElementAttribute(k)
            
            'Header Keys:
            ' - Id
            ' - Name
            ' - Description
            ' - UserOrGroupSid
            ' - Action (Allow, Deny)
            'Debug.Print "key = " & xmlAttribute.KeyWord & " - " & xmlAttribute.Value
            
            If xmlAttribute.KeyWord = "Action" Then
                If xmlAttribute.Value = "Deny" Then bDeny = True
            End If
        Next

        Set xmlElement = xmlDoc.Root.NodeByName("Conditions")
        
        If Not (xmlElement Is Nothing) Then
        
            For j = 1 To xmlElement.NodeCount
                
                If StrComp(xmlElement.Node(j).Name, "FilePathCondition", 1) = 0 Then
                
                    eRuleType = APPLOCKER_RULE_FILE_PATH
                
                    For k = 1 To xmlElement.Node(j).AttributeCount
                
                        Set xmlAttribute = xmlElement.Node(j).ElementAttribute(k)
                        
                        'keys:
                        ' - Path
                        'Debug.Print "key = " & xmlAttribute.KeyWord & " - " & xmlAttribute.Value
                        
                        If xmlAttribute.KeyWord = "Path" Then sFilePath = xmlAttribute.Value
                    Next
                    
                ElseIf StrComp(xmlElement.Node(j).Name, "FileHashCondition", 1) = 0 Then
                    
                    eRuleType = APPLOCKER_RULE_FILE_HASH
                    
                    If xmlElement.Node(j).NodeCount <> 0 Then
                        ReDim tHashRuleData(xmlElement.Node(j).NodeCount - 1)
                    Else
                        ReDim tHashRuleData(0)
                    End If
                    
                    For m = 1 To xmlElement.Node(j).NodeCount
                        
                        Set xmlSubNode = xmlElement.Node(j).Node(m)
                        
                        If StrComp(xmlSubNode.Name, "FileHash", 1) = 0 Then
                        
                            For k = 1 To xmlSubNode.AttributeCount
                
                                Set xmlAttribute = xmlSubNode.ElementAttribute(k)
                                
                                'keys:
                                ' - Type
                                ' - Data
                                ' - SourceFileName
                                ' - SourceFileLength
                                'Debug.Print "key = " & xmlAttribute.KeyWord & " - " & xmlAttribute.Value
                                
                                If xmlAttribute.KeyWord = "Data" Then tHashRuleData(m - 1).Hash = xmlAttribute.Value
                                If xmlAttribute.KeyWord = "SourceFileName" Then tHashRuleData(m - 1).FileName = xmlAttribute.Value
                                If xmlAttribute.KeyWord = "SourceFileLength" Then tHashRuleData(m - 1).FileLength = xmlAttribute.Value
                                
                            Next
                            
                        End If
                    Next
                ElseIf StrComp(xmlElement.Node(j).Name, "FilePublisherCondition", 1) = 0 Then
                    
                    eRuleType = APPLOCKER_RULE_FILE_PUBLISHER
                
                    For k = 1 To xmlElement.Node(j).AttributeCount
                
                        Set xmlAttribute = xmlElement.Node(j).ElementAttribute(k)
                        
                        'keys:
                        ' - PublisherName
                        ' - ProductName
                        ' - BinaryName
                        'Debug.Print "key = " & xmlAttribute.KeyWord & " - " & xmlAttribute.Value
                        
                        '+ subnode "BinaryVersionRange" with keys:
                        ' - LowSection
                        ' - HighSection
                        
                        If xmlAttribute.KeyWord = "PublisherName" Then sPublisherName = xmlAttribute.Value
                    Next
                Else
                    If inIDE Then Debug.Print "------ No definition!"
                End If
            Next
        End If
        
        bException = Not (xmlDoc.Root.NodeByName("Exceptions") Is Nothing)
        
        sHit = "O7 - AppLocker: " & IIf(bDeny, "(Deny) ", "(Allow) ")
        
        Select Case LCase$(GetFileName(GetParentDir(sKey)))
            Case "exe":                 sHit = sHit & "[Executable] "
            Case "msi":                 sHit = sHit & "[Installer] "
            Case "script":              sHit = sHit & "[Script] "
            Case "dll":                 sHit = sHit & "[Library] "
            Case "appx":                sHit = sHit & "[AppX] "
            Case "managedinstaller":    sHit = sHit & "[ManagedInstaller] "
        End Select
        
        If eRuleType = APPLOCKER_RULE_FILE_HASH Then
            numEntries = UBound(tHashRuleData) + 1
        Else
            numEntries = 1
        End If
        
        sHitTemplate = sHit
        
        For m = 0 To numEntries - 1
        
            sHit = sHitTemplate
        
            Select Case eRuleType
                Case APPLOCKER_RULE_FILE_PATH:      sHit = sHit & "[Path] " & sFilePath
                Case APPLOCKER_RULE_FILE_HASH:      sHit = sHit & "[Hash] " & tHashRuleData(m).FileName & " (Size: " & tHashRuleData(m).FileLength & ") - " & tHashRuleData(m).Hash
                Case APPLOCKER_RULE_FILE_PUBLISHER: sHit = sHit & "[Publisher] " & sPublisherName
                Case APPLOCKER_RULE_UNKNOWN:        sHit = sHit & "[Unknown] " & sRuleId
            End Select
            
            If bException Then sHit = sHit & " (Exceptions present)"
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    
                    AddRegToFix .Reg, REMOVE_KEY, 0, sKey
                    
                    If StrBeginWith(sKey, "HKCU") Then 'remove the mirror
                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Policies\Microsoft\Windows\SrpV2\Exe\" & sRuleId
                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Policies\Microsoft\Windows\SrpV2\Msi\" & sRuleId
                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Policies\Microsoft\Windows\SrpV2\Script\" & sRuleId
                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Policies\Microsoft\Windows\SrpV2\Dll\" & sRuleId
                        AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Policies\Microsoft\Windows\SrpV2\Appx\" & sRuleId
                    End If
                    
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SYSTEM\CurrentControlSet\Control\Srp\Gp\Exe\" & sRuleId
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SYSTEM\CurrentControlSet\Control\Srp\Gp\Msi\" & sRuleId
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SYSTEM\CurrentControlSet\Control\Srp\Gp\Script\" & sRuleId
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SYSTEM\CurrentControlSet\Control\Srp\Gp\Dll\" & sRuleId
                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SYSTEM\CurrentControlSet\Control\Srp\Gp\Appx\" & sRuleId
                    
                    If AryPtr(aFiles) Then 'files in %SystemRoot%\System32\AppLocker\*
                        For k = 0 To UBound(aFiles)
                            AddFileToFix .File, REMOVE_FILE, aFiles(k)
                        Next
                    End If
                    
                    AddFileToFix .File, REMOVE_FILE, BuildPath(sWinSysDir, "GroupPolicy\Machine\Registry.pol")
                    AddFileToFix .File, REMOVE_FILE, BuildPath(AllUsersProfile, "ntuser.pol")
                    
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                ConcatScanResults resultAll, result
                AddToScanResults result
            End If
        Next
    Next
    
    If resultAll.CureType <> 0 Then
        resultAll.HitLineW = "O7 - AppLocker: Fix all (including policies)"
        resultAll.FixAll = True
        resultAll.CureType = resultAll.CureType Or CUSTOM_BASED
        AddRegToFix resultAll.Reg, REMOVE_KEY, HKLM, "SOFTWARE\Policies\Microsoft\Windows\SrpV2"
        AddRegToFix resultAll.Reg, REMOVE_KEY, HKCU, "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy Objects"
        AddCustomToFix resultAll.Custom, CUSTOM_ACTION_APPLOCKER
        AddToScanResults resultAll
    End If
    
    
    Exit Sub
    
    'SRP v1 // TODO
    
    Dim ccsID As Long
    Dim CSKey As String
    Dim sName As String
    Dim sSDDL As String
    
    ccsID = Reg.GetDword(HKLM, "SYSTEM\Select", "Current")
    
    'HKLM\SYSTEM\ControlSet*\Control\Srp\Gp\{type}\{CLSID}
    
    For j = 0 To 99    ' 0 - is CCS
    
        CSKey = IIf(j = 0, "System\CurrentControlSet", "System\ControlSet" & Format$(j, "000"))
        
        If j > 0 Then
            If Not Reg.KeyExists(HKEY_LOCAL_MACHINE, CSKey) Then Exit For
        End If
        
        If j = 0 Or j <> ccsID Then
    
            For Each vType In Array("Exe", "Msi", "Script", "Dll", "Appx") 'Appx ?
                
                sKey = CSKey & "\Control\Srp\Gp\" & vType
                
                For k = 1 To Reg.EnumSubKeysToArray(HKLM, sKey, aRuleKeys(), bRecursively:=False, bEraseArray:=True)
            
                    sName = Reg.GetString(HKLM, sKey & "\" & aRuleKeys(k), "Name")
                    sSDDL = Reg.GetString(HKLM, sKey & "\" & aRuleKeys(k), "SDDL")
                    
                    'Debug.Print sName & " - " & sSDDL
                    
                Next
            Next
        End If
    Next
    
    Set dRuleXml = Nothing
    Set dRuleKey = Nothing
    
    AppendErrorLogCustom "CheckAppLocker - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckAppLocker"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RestoreApplockerDefaults()
    On Error GoTo ErrorHandler:
    
    Dim hFile As Long
    Dim strPath As String: strPath = BuildPath(TempCU, "applocker_clear.xml")
    
    If OSver.IsWindows10OrGreater Then
        If OpenW(strPath, FOR_OVERWRITE_CREATE, hFile, g_FileBackupFlag) Then
            PrintLineW hFile, "<AppLockerPolicy Version=""1"">"
            PrintLineW hFile, "<RuleCollection Type=""Exe"" EnforcementMode=""NotConfigured"" />"
            PrintLineW hFile, "<RuleCollection Type=""Msi"" EnforcementMode=""NotConfigured"" />"
            PrintLineW hFile, "<RuleCollection Type=""Script"" EnforcementMode=""NotConfigured"" />"
            PrintLineW hFile, "<RuleCollection Type=""Dll"" EnforcementMode=""NotConfigured"" />"
            PrintLineW hFile, "<RuleCollection Type=""Appx"" EnforcementMode=""NotConfigured"" />"
            PrintLineW hFile, "<RuleCollection Type=""ManagedInstaller"" EnforcementMode=""NotConfigured"" />"
            PrintLineW hFile, "</AppLockerPolicy>"
            CloseW hFile
            Call Proc.RunPowershell("import-module AppLocker; Set-AppLockerPolicy -XMLPolicy '" & strPath & "'", True, 30000)
            DeleteFileW StrPtr(strPath)
        End If
        
        If Proc.ProcessRun(BuildPath(sWinSysDir, "appidtel.exe"), "stop -mionly", , vbHide) Then
            Proc.WaitForTerminate , , , 15000
        End If
    End If
    
    SetServiceStartMode "applockerfltr", SERVICE_MODE_MANUAL
    SetServiceStartMode "appidsvc", SERVICE_MODE_MANUAL
    SetServiceStartMode "appid", SERVICE_MODE_MANUAL
    StopService "applockerfltr"
    StopService "appidsvc"
    StopService "appid"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "RestoreApplockerDefaults"
    If inIDE Then Stop: Resume Next
End Sub

Public Function EnableApplocker() As Boolean
    Dim bSuccess As Boolean: bSuccess = True
    SetServiceStartMode "applockerfltr", SERVICE_MODE_AUTOMATIC
    SetServiceStartMode "appidsvc", SERVICE_MODE_AUTOMATIC
    SetServiceStartMode "appid", SERVICE_MODE_AUTOMATIC
    bSuccess = bSuccess And StartService("applockerfltr")
    bSuccess = bSuccess And StartService("appidsvc")
    bSuccess = bSuccess And StartService("appid")
End Function


Public Sub CheckO7Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO7Item - Begin"

    'Policies
    CheckPolicies
    
    'Policy - Logon scripts
    CheckPolicyScripts
    
    CheckCredentials
    
    CheckPolicyUAC
    
    'Untrusted certificates
    UpdateProgressBar "O7-Cert"
    Call CheckCertificatesEDS
    If Not bAutoLogSilent Then DoEvents
    
    'AppLocker software run restrictions
    CheckAppLocker
    
    ' System troubleshooting
    UpdateProgressBar "O7-Trouble"
    Call CheckSystemProblems '%temp%, %tmp%, disk free space < 1 GB.
    If Not bAutoLogSilent Then DoEvents
    
    'Check for DACL lock on Policy key
    UpdateProgressBar "O7-ACL"
    Call CheckPolicyACL
    If Not bAutoLogSilent Then DoEvents
    
    'IP Security
    UpdateProgressBar "O7-IPSec"
    Call CheckIPSec
    
    AppendErrorLogCustom "CheckO7Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO7Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO7Item_Bitcoin(sWalletAddr As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO7Item_Bitcoin - Begin"
    
    Dim sHit As String, sActualClip As String, result As SCAN_RESULT

    DoEvents
    sActualClip = ClipboardGetText()
    
    If (InStr(1, sActualClip, "0x", 1) <> 0) And (sWalletAddr <> sActualClip) Then

        sHit = "O7 - Policy: Bitcoin wallet address hijacker is present: blockchain.com/explorer/search?search=" & sActualClip & " " & STR_NO_FIX
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O7"
                .HitLineW = sHit
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
    End If
    
    AppendErrorLogCustom "CheckO7Item_Bitcoin - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO7Item_Bitcoin"
    If inIDE Then Stop: Resume Next
End Sub

Public Function GenWalletAddressETH() As String
    Dim i&, s$, Value&
    Randomize Time
    For i = 1 To 40
        Value = Int(Rnd * 16) '48-57 (0-9), 65-70 (A-F)
        Value = Value + 48
        If Value > 57 Then Value = Value + 7
        s = s & Chr$(Value)
    Next
    GenWalletAddressETH = "0x" & s
End Function

Public Sub CheckAutoLogon()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckAutoLogon - Begin"
    
    Dim sHit As String, sKey As String, sUser As String, sDomain As String, bEnabled As Boolean, result As SCAN_RESULT
    Dim sAccType As String
    
    sKey = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
    sUser = Reg.GetString(HKEY_LOCAL_MACHINE, sKey, "DefaultUserName")
    
    If Len(sUser) <> 0 Then
        
        bEnabled = 0 <> Reg.GetDword(HKEY_LOCAL_MACHINE, sKey, "AutoAdminLogon")
        
        If bEnabled Or bIgnoreAllWhitelists Then
        
            sDomain = Reg.GetString(HKEY_LOCAL_MACHINE, sKey, "DefaultDomainName")
            sAccType = GetUserAccountType(sUser)
            If Len(sAccType) <> 0 Then sAccType = " (type: " & sAccType & ")"
            
            sHit = "O27 - Account: (AutoLogon) HKLM\..\Winlogon: " & sDomain & "\" & sUser & sAccType & IIf(bEnabled, "", " (disabled)")
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O27"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HKLM, sKey, "AutoAdminLogon", "0", , REG_RESTORE_SZ
                    AddRegToFix .Reg, REMOVE_VALUE, HKLM, sKey, "DefaultDomainName"
                    AddRegToFix .Reg, REMOVE_VALUE, HKLM, sKey, "DefaultUserName"
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    End If
    
    'Logon screen text
    
    Dim sText As String, i As Long, iLength As Long
    ReDim aParam(1) As String
    aParam(0) = "LegalNoticeCaption"
    aParam(1) = "LegalNoticeText"
    
    For i = 0 To UBound(aParam)
        
        sText = Reg.GetString(HKEY_LOCAL_MACHINE, sKey, aParam(i))
        
        If Len(sText) <> 0 Then
            
            sText = Replace$(sText, vbCr, vbNullString)
            sText = Replace$(sText, vbLf, "\n")
            iLength = Len(sText)
            If iLength > 300 Then sText = Left$(sText, 300) & " ... (" & iLength & " characters)"
            
            sHit = "O7 - Policy: HKLM\..\Winlogon: [" & aParam(i) & "] = " & sText
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HKLM, sKey, aParam(i), "", , REG_RESTORE_SZ
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    AppendErrorLogCustom "CheckAutoLogon - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckAutoLogon"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckIPSec()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckIPSec - Begin"
    
    Dim sHit$, sHit1$, sHit2$, result As SCAN_RESULT
    Dim i As Long
    
    'IPSec policy
    
    'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\
    'secpol.msc
    
    'ipsecPolicy{GUID}                  'example: 5d57bbac-8464-48b2-a731-9dd7e6f65c9f
    
    '\ipsecName                         -> Name of policy
    '\whenChanged                       -> Date in Unix format ( ConvertUnixTimeToLocalDate )
    '\ipsecNFAReference [REG_MULTI_SZ]  -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNFA{GUID_1} '96372b24-f2bf-4f50-a036-5897aac92f2f
                                                   'SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNFA{GUID_2} '8c676c64-306c-47db-ab50-e0108a1621dd
    
    'Note: One of these ipsecNFA{GUID} key may not contain 'ipsecFilterReference' parameter
    
    '\ipsecISAKMPReference              -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecISAKMPPolicy{GUID} '738d84c5-d070-4c6c-9468-12b171cfd10e
    
    '--------------------------------
    'ipsecNFA{GUID}
    
    '\ipsecNegotiationPolicyReference     -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecNegotiationPolicy{GUID} '7c5a4ff0-ae4b-47aa-a2b6-9a72d2d6374c
    '\ipsecFilterReference [REG_MULTI_SZ] -> example: SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\ipsecFilter{GUID_1} 'c73baa5d-71a6-4533-bf7d-f640b1ff2eb8
    
    '--------------------------------
    'ipsecNegotiationPolicy{GUID}
    
    'Trojan example: https://www.trendmicro.com/vinfo/us/threat-encyclopedia/malware/troj_dloade.xn
    
    Dim KeyPolicy() As String, IPSecName$, KeyNFA() As String, KeyNegotiation() As String, dModify As Date, lModify As Long, IPSecID As String
    Dim KeyISAKMP As String, j As Long, KeyFilter() As String, k As Long, NegAction As String, NegType As String, bEnabled As Boolean, sActPolicy As String
    Dim bRegexpInit As Boolean, bFilterData() As Byte, IP(1) As String, RuleAction As String, bMirror As Boolean
    Dim Packet_Type(1) As String, m As Long, N As Long, PortNum(1) As Long, ProtocolType As String, IpFil As IPSEC_FILTER_RECORD
    Dim oMatches As IRegExpMatchCollection, IPTypeFlag(1) As Long, b() As Byte, bAtLeastOneFilter As Boolean, bNoFilter As Boolean
    Dim bSafe As Boolean
    Dim odHit As clsTrickHashTable
    Set odHit = New clsTrickHashTable
    
    
    For i = 1 To Reg.EnumSubKeysToArray(0&, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local", KeyPolicy())
    
      If StrBeginWith(KeyPolicy(i), "ipsecPolicy{") Then
        
        'what policy is currently active?
        sActPolicy = Reg.GetString(0&, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local", "ActivePolicy")
        
        bEnabled = (StrComp(sActPolicy, "SOFTWARE\Policies\Microsoft\Windows\IPSEC\Policy\Local\" & KeyPolicy(i), 1) = 0)
        
        'add prefix
        KeyPolicy(i) = "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyPolicy(i)
        
        bMirror = False
        RuleAction = vbNullString
        
        IPSecID = mid$(KeyPolicy(i), InStrRev(KeyPolicy(i), "{"))
        
        IPSecName = Reg.GetString(0&, KeyPolicy(i), "ipsecName")
        
        lModify = Reg.GetDword(0&, KeyPolicy(i), "whenChanged")
        
        dModify = ConvertUnixTimeToLocalDate(lModify)
        
        KeyISAKMP = Reg.GetString(0&, KeyPolicy(i), "ipsecISAKMPReference")
        KeyISAKMP = MidFromCharRev(KeyISAKMP, "\")
        KeyISAKMP = IIf(Len(KeyISAKMP) = 0, vbNullString, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyISAKMP)
        
        Erase KeyNFA
        Erase KeyFilter
        Erase KeyNegotiation
        Erase IP
        Erase Packet_Type: Packet_Type(0) = "Unknown": Packet_Type(1) = "Unknown"
        Erase PortNum
        RuleAction = vbNullString
        ProtocolType = vbNullString
        bMirror = False
        RuleAction = "Unknown"
        bNoFilter = False
        
        KeyNFA() = Reg.GetMultiSZ(0&, KeyPolicy(i), "ipsecNFAReference")
        '() -> ipsecNegotiationPolicy
        '() -> ipsecFilter (optional)
        
        If AryItems(KeyNFA) Then
            
          For j = 0 To UBound(KeyNFA)
            KeyNFA(j) = MidFromCharRev(KeyNFA(j), "\")
            KeyNFA(j) = IIf(Len(KeyNFA(j)) = 0, vbNullString, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyNFA(j))
          Next
          
          ReDim KeyNegotiation(UBound(KeyNFA))
          
          For j = 0 To UBound(KeyNFA)
            KeyNegotiation(j) = Reg.GetString(0&, KeyNFA(j), "ipsecNegotiationPolicyReference")
            KeyNegotiation(j) = MidFromCharRev(KeyNegotiation(j), "\")
            KeyNegotiation(j) = IIf(Len(KeyNegotiation(j)) = 0, vbNullString, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyNegotiation(j))
          Next
          
          For j = 0 To UBound(KeyNFA)
            
            NegType = Reg.GetString(0&, KeyNegotiation(j), "ipsecNegotiationPolicyType")
            NegAction = Reg.GetString(0&, KeyNegotiation(j), "ipsecNegotiationPolicyAction")
            
            'GUIDs: https://msdn.microsoft.com/en-us/library/cc232441.aspx
            
            If StrComp(NegType, "{62f49e10-6c37-11d1-864c-14a300000000}", 1) = 0 Then 'without last one "-" character (!)
                If StrComp(NegAction, "{8a171dd2-77e3-11d1-8659-a04f00000000}", 1) = 0 Then
                    RuleAction = "Allow"
                ElseIf StrComp(NegAction, "{3f91a819-7647-11d1-864d-d46a00000000}", 1) = 0 Then
                    RuleAction = "Block"
                ElseIf StrComp(NegAction, "{8a171dd3-77e3-11d1-8659-a04f00000000}", 1) = 0 Then
                    RuleAction = "Approve security"
                ElseIf StrComp(NegAction, "{3f91a81a-7647-11d1-864d-d46a00000000}", 1) = 0 Then
                    RuleAction = "Inbound pass-through"
                Else
                    RuleAction = "Unknown"
                End If
            ElseIf StrComp(NegType, "{62f49e13-6c37-11d1-864c-14a300000000}", 1) = 0 Then
                RuleAction = "Default response"
            Else
                RuleAction = "Unknown"
            End If
            
            Erase KeyFilter
            Erase IP
            Erase Packet_Type: Packet_Type(0) = "Unknown": Packet_Type(1) = "Unknown"
            Erase PortNum
            ProtocolType = vbNullString
            bMirror = False
            
            KeyFilter() = Reg.GetMultiSZ(0&, KeyNFA(j), "ipsecFilterReference")
                        
            If 0 = AryItems(KeyFilter) Then
            
                bAtLeastOneFilter = False
                
                For m = 0 To UBound(KeyNFA)
                    If Reg.ValueExists(0&, KeyNFA(m), "ipsecFilterReference") Then
                        bAtLeastOneFilter = True
                        Exit For
                    End If
                Next
                
                If Not bAtLeastOneFilter Then
                    bNoFilter = True
                    GoSub AddItem
                End If
            Else
                
                For k = 0 To UBound(KeyFilter)
                    KeyFilter(k) = MidFromCharRev(KeyFilter(k), "\")
                    KeyFilter(k) = IIf(Len(KeyFilter(k)) = 0, vbNullString, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyFilter(k))
                Next
                
                For k = 0 To UBound(KeyFilter)
                    
                    Erase IP
                    Erase Packet_Type: Packet_Type(0) = "Unknown": Packet_Type(1) = "Unknown"
                    Erase PortNum
                    ProtocolType = vbNullString
                    bMirror = False
                    
                    bFilterData() = Reg.GetBinary(0&, KeyFilter(k), "ipsecData")
                    
                    If AryItems(bFilterData) Then

                      AppendErrorLogCustom "CheckO7Item - Regexp - Begin"

                      If Not g_bRegexpInit Then
                        Set oRegexp = New cRegExp
                        g_bRegexpInit = True
                      End If

                      If Not bRegexpInit Then
                        bRegexpInit = True
                        oRegexp.IgnoreCase = True
                        oRegexp.Global = True
                        oRegexp.Pattern = "(00|01)(000000)(........)(00000000|FFFFFFFF)(........)(00000000|FFFFFFFF)(00000000)(((06|11)000000........)|((00|01|06|08|11|14|16|1B|42|FF|..)00000000000000))00(00|01|02|03|04|81|82|83|84)0000"
                      End If
                      
                      Set oMatches = oRegexp.Execute(SerializeByteArray(bFilterData))
                    
                      AppendErrorLogCustom "CheckO7Item - Regexp - End"
    
                      For N = 0 To oMatches.Count - 1
                      
                        b = DeSerializeToByteArray(oMatches(N))
                        
                        memcpy IpFil, b(0), Len(IpFil)

                        '00,00,00,00,00,00,00,00 -> any IP
                        'xx,xx,xx,xx,ff,ff,ff,ff -> specified IP / subnet
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 0 -> my IP
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 1 or 0x81 -> DNS-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 2 or 0x82 -> WINS-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 3 or 0x83 -> DHCP-servers
                        '00,00,00,00,ff,ff,ff,ff + [0x6F] == 4 or 0x84 -> Gateway
                        '
                        '[0x4E] == 1 -> mirrored
                        '
                        '[0x66] -> port type
                        '[0x6A] -> port number (source) (2 bytes)
                        '[0x6C] -> port number (destination) (2 bytes)

                        bMirror = (IpFil.Mirrored = 1)
                        PortNum(0) = cMath.ShortIntToUShortInt(IpFil.PortNum1)
                        PortNum(1) = cMath.ShortIntToUShortInt(IpFil.PortNum2)
                        
                        Select Case IpFil.ProtocolType
                            Case 0: ProtocolType = "Any"
                            Case 6: ProtocolType = "TCP"
                            Case 17: ProtocolType = "UDP"
                            Case 1: ProtocolType = "ICMP"
                            Case 27: ProtocolType = "RDP"
                            Case 8: ProtocolType = "EGP"
                            Case 20: ProtocolType = "HMP"
                            Case 255: ProtocolType = "RAW"
                            Case 66: ProtocolType = "RVD"
                            Case 22: ProtocolType = "XNS-IDP"
                            Case Else: ProtocolType = "type: " & CLng(bFilterData(&H66))
                        End Select
                        
                        IP(0) = IpFil.IP1(0) & "." & IpFil.IP1(1) & "." & IpFil.IP1(2) & "." & IpFil.IP1(3)
                        IP(1) = IpFil.IP2(0) & "." & IpFil.IP2(1) & "." & IpFil.IP2(2) & "." & IpFil.IP2(3)

                        IPTypeFlag(0) = IpFil.IPTypeFlag1
                        IPTypeFlag(1) = IpFil.IPTypeFlag2

                        For m = 0 To 1
                        
                            If IPTypeFlag(m) = 0 Then       '00,00,00,00,00,00,00,00
                                If IP(m) = "0.0.0.0" Then Packet_Type(m) = "Any IP"
                            
                            ElseIf IPTypeFlag(m) = -1 Then  '00,00,00,00,ff,ff,ff,ff
                                If IP(m) = "0.0.0.0" Then
                            
                                    Select Case IpFil.DynPacketType
                                        Case 0: Packet_Type(m) = "my IP"
                                        Case &H81, 1: Packet_Type(m) = "DNS-servers"
                                        Case &H82, 2: Packet_Type(m) = "WINS-servers"
                                        Case &H83, 3: Packet_Type(m) = "DHCP-servers"
                                        Case &H84, 4: Packet_Type(m) = "Gateway"
                                        Case Else: Packet_Type(m) = "Unknown"
                                        '1,2,3,4 - Source packets
                                        '81,82,83,84 - Destination packets
                                    End Select
                                Else                            'xx,xx,xx,xx,ff,ff,ff,ff
                                    Packet_Type(m) = "IP"
                                End If
                            Else
                                Packet_Type(m) = "Unknown"
                            End If
                        
                            If IP(m) = "0.0.0.0" Then IP(m) = vbNullString
                        Next
                        
                        GoSub AddItem
                        
                      Next
                      
                    Else
                        GoSub AddItem
                    End If
                Next

            End If
            
          Next
          
        Else
            GoSub AddItem
        End If
        
      End If
    Next
    
    Set odHit = Nothing
    
    If Not bAutoLogSilent Then DoEvents
    
    AppendErrorLogCustom "CheckIPSec - End"
    Exit Sub
AddItem:
    'keys:
    'KeyPolicy(i) - 1
    'KeyISAKMP - 1
    'KeyNFA(j) - 0 to ...
    'KeyNegotiation - 1
    'KeyFilter(k) - 0 to ...
    
    'flags:
    'bEnabled - policy enabled ?
    'bMirror - true, if rule also applies to reverse direction: from destination to source
    
    'Other:
    'IPSecName - name of policy
    'IPSecID - identifier in registry
    'dModify - date last modified
    'RuleAction - action for filter
    'PortNum()
    'ProtocolType
    
    'example:
    'O7 - IPSec: (Enabled) IP_Policy_Name [yyyy/mm/dd] - {5d57bbac-8464-48b2-a731-9dd7e6f65c9f} - Source: My IP - Destination: 8.8.8.8 (Port 80 TCP) - (mirrored) Action: Block
    
    'do not show in log disabled entries
    If Not bIgnoreAllWhitelists Then
        If Not bEnabled Then Return
    End If
    
    sHit1 = "O7 - IPSec: Name: " & IPSecName & " " & _
        "(" & Format$(dModify, "yyyy\/mm\/dd") & ")" & " - "
        
    sHit2 = IPSecID & " - " & _
        IIf(bNoFilter, "No rules ", _
        "Source: " & IIf(Packet_Type(0) = "IP", "IP: " & IP(0), Packet_Type(0)) & _
        IIf((ProtocolType = "TCP" Or ProtocolType = "UDP") And PortNum(0) <> 0, " (Port " & PortNum(0) & " " & ProtocolType & ")", vbNullString) & " - " & _
        "Destination: " & IIf(Packet_Type(1) = "IP", "IP: " & IP(1), Packet_Type(1)) & _
        IIf((ProtocolType = "TCP" Or ProtocolType = "UDP") And PortNum(1) <> 0, " (Port " & PortNum(1) & " " & ProtocolType & ")", vbNullString) & " " & _
        IIf(bMirror, "(mirrored) ", vbNullString)) & "- Action: " & RuleAction & IIf(bEnabled, vbNullString, " (disabled)")
    
    sHit = sHit1 & sHit2
    
    If odHit.Exists(sHit) Then 'skip several identical rules
        If Not bAutoLogSilent Then DoEvents
        Return
    Else
        odHit.Add sHit, 0&
    End If
    
    bSafe = False
    
    'Whitelists
    If bHideMicrosoft Then
      If (OSver.MajorMinor <= 5.2) Or (OSver.MajorMinor = 5.2 And OSver.IsWin64) Then  'Win2k / XP / XP x64
        If StrComp(sHit2, "{72385236-70fa-11d1-864c-14a300000000} - No rules - Action: Default response (disabled)", 1) = 0 Then
            bSafe = True
        ElseIf StrComp(sHit2, "{72385230-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Allow (disabled)", 1) = 0 Then
            bSafe = True
        ElseIf StrComp(sHit2, "{72385230-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Inbound pass-through (disabled)", 1) = 0 Then
            bSafe = True
        ElseIf StrComp(sHit2, "{7238523c-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Allow (disabled)", 1) = 0 Then
            bSafe = True
        ElseIf StrComp(sHit2, "{7238523c-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Inbound pass-through (disabled)", 1) = 0 Then
            bSafe = True
        End If
      End If
    End If
    
    If Not bSafe Then
      If Not IsOnIgnoreList(sHit) Then
        With result
            .Section = "O7"
            .HitLineW = sHit
            AddRegToFix .Reg, REMOVE_KEY, 0, KeyPolicy(i)
            If Len(KeyISAKMP) <> 0 Then AddRegToFix .Reg, REMOVE_KEY, 0, KeyISAKMP
            If AryItems(KeyNFA) Then
                For m = 0 To UBound(KeyNFA)
                    If Len(KeyNFA(m)) <> 0 Then
                        AddRegToFix .Reg, REMOVE_KEY, 0, KeyNFA(m)
                    End If
                    If Len(KeyNegotiation(m)) <> 0 Then
                        AddRegToFix .Reg, REMOVE_KEY, 0, KeyNegotiation(m)
                    End If
                Next
            End If
            If AryItems(KeyFilter) Then
                For m = 0 To UBound(KeyFilter)
                    If Len(KeyFilter(m)) <> 0 Then
                        AddRegToFix .Reg, REMOVE_KEY, 0, KeyFilter(m)
                    End If
                Next
            End If
            If bEnabled Then
                AddRegToFix .Reg, REMOVE_VALUE, 0, "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local", "ActivePolicy"
            End If
            .CureType = REGISTRY_BASED
        End With
        AddToScanResults result
      End If
    End If
    
    Return
ErrorHandler:
    ErrorMsg Err, "ModMain_CheckIPSec"
    If inIDE Then Stop: Resume Next
End Sub

'byte array -> to Hex String
Public Function SerializeByteArray(b() As Byte, Optional Delimiter As String = vbNullString) As String
    Dim i As Long
    Dim s As String
    SerializeByteArray = String$((UBound(b) + 1) * 2, "0")
       
    For i = 0 To UBound(b)
        s = Hex$(b(i))
        Mid$(SerializeByteArray, (i * 2) + 1 + IIf(Len(s) = 2, 0, 1)) = s
    Next
End Function

'Serialized Hex String of bytes -> byte array
Public Function DeSerializeToByteArray(s As String, Optional Delimiter As String = vbNullString) As Byte()
    On Error GoTo ErrorHandler:
    Dim i As Long
    Dim N As Long
    Dim b() As Byte
    Dim ArSize As Long
    If Len(s) = 0 Then Exit Function
    ArSize = (Len(s) + Len(Delimiter)) \ (2 + Len(Delimiter)) '2 chars on byte + add final delimiter
    ReDim b(ArSize - 1) As Byte
    For i = 1 To Len(s) Step 2 + Len(Delimiter)
        b(N) = CLng("&H" & mid$(s, i, 2))
        N = N + 1
    Next
    DeSerializeToByteArray = b
    Exit Function
ErrorHandler:
    If inIDE Then Debug.Print "Error in DeSerializeByteString"
End Function

Public Sub FixO7Item(sItem$, result As SCAN_RESULT)
    'O7 - Disabling of Policies
    On Error GoTo ErrorHandler:
    
    If InStr(1, result.HitLineW, "Policy Script:", 1) <> 0 Then bNeedRebuildPolicyChain = True
    
    If result.CureType = CUSTOM_BASED Then
    
        If InStr(1, result.HitLineW, "Free disk space", 1) <> 0 Then
            RunCleanMgr
            
        ElseIf InStr(1, result.HitLineW, "Computer name (hostname) is not set", 1) <> 0 Then
            Dim sNetBiosName As String
            sNetBiosName = GetCompName(ComputerNameNetBIOS)
            If Len(sNetBiosName) = 0 Then
                sNetBiosName = Environ$("USERDOMAIN")
                If Len(sNetBiosName) = 0 Then
                    sNetBiosName = "USER-PC"
                End If
                SetCompName ComputerNamePhysicalNetBIOS, sNetBiosName
            End If
            SetCompName ComputerNamePhysicalDnsHostname, sNetBiosName
            bRebootRequired = True
        End If
        
    Else
        FixIt result
        bUpdatePolicyNeeded = True
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO7Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub RunCleanMgr()

    'https://winaero.com/blog/cleanmgr-exe-command-line-arguments-in-windows-10/

    Dim sRootKey As String
    Dim cKeys As Collection
    Dim nid As Long
    Dim i As Long
    Dim sKey As String
    Dim sParam As String
    Dim lData As Long
    Dim sCleanMgr As String
    
    'all, except:
    'Recycle Bin
    'Windows update report
    'System crash dump
    'Remote Desktop Cache Files
    'Windows EDS files (needed to Refresh or Reset PC on Win 8/10)
    
    '//TODO: (�� sov44)
    '������ ������ ����� c:\Windows\Installer �� ���������� ���������� �����.
    '�� ������ ������� ��� �� 1 �� �����. ���������� ���� � ���� ������� ����� - �� ����� state �������. => look Sources\Cleaner
    '��������� ������ ����� c:\Windows\SoftwareDistribution\Download, �� � �� �������, �.�. ����� ��������
    '� ������ ��������� ���������� �� ������� "��������� � ��������".
    '���-�� ������ c:\Windows\winsxs\Backup, c:\Windows\winsxs\Temp,
    '�� ��� �� ������� �������� �������� � ������� ������� �� ������ �����.
    
    Set cKeys = New Collection
    
    nid = 777
    sRootKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
    cKeys.Add 2, "Active Setup Temp Folders"
    cKeys.Add 2, "BranchCache" '8/10
    cKeys.Add 2, "Compress old files" 'XP
    cKeys.Add 0, "Content Indexer Cleaner"
    cKeys.Add 2, "Downloaded Program Files"
    cKeys.Add 2, "GameUpdateFiles"
    cKeys.Add 2, "Internet Cache Files"
    cKeys.Add 2, "Memory Dump Files"
    cKeys.Add 2, "Offline Pages Files"
    cKeys.Add 2, "Old ChkDsk Files"
    cKeys.Add 2, "Previous Installations"
    cKeys.Add 0, "Recycle Bin"
    cKeys.Add 0, "Remote Desktop Cache Files" 'XP
    cKeys.Add 2, "RetailDemo Offline Content" '8/10
    cKeys.Add 2, "Service Pack Cleanup"
    cKeys.Add 0, "Setup Log Files"
    cKeys.Add 0, "System error memory dump files"
    cKeys.Add 0, "System error minidump files"
    cKeys.Add 2, "Temporary Files"
    cKeys.Add 2, "Temporary Setup Files"
    cKeys.Add 2, "Thumbnail Cache"
    cKeys.Add 2, "Update Cleanup"
    cKeys.Add 2, STR_CONST.WINDOWS_DEFENDER '8/10
    cKeys.Add 2, "User file versions" '8/10
    cKeys.Add 2, "Upgrade Discarded Files"
    cKeys.Add 2, "WebClient and WebPublisher Cache" 'XP
    cKeys.Add 2, "Windows Error Reporting Archive Files"
    cKeys.Add 2, "Windows Error Reporting Queue Files"
    cKeys.Add 2, "Windows Error Reporting System Archive Files"
    cKeys.Add 2, "Windows Error Reporting System Queue Files"
    cKeys.Add 2, "Windows Error Reporting Temp Files" '8/10
    cKeys.Add 0, "Windows ESD installation files"
    cKeys.Add 0, "Windows Upgrade Log Files"

    sParam = "StateFlags" & Right$("000" & nid, 4)
    'set preset
    For i = 1 To cKeys.Count
        lData = CLng(cKeys(i))
        sKey = sRootKey & "\" & GetCollectionKeyByIndex(i, cKeys)
        
        If Reg.KeyExists(0, sKey) Then
            Call Reg.SetDwordVal(0, sKey, sParam, lData)
        End If
    Next
    'run cleaner

    sCleanMgr = sSysNativeDir & "\CleanMgr.exe"
    
    If Proc.ProcessRun(sCleanMgr, "/SAGERUN:" & nid) Then
        Do While Proc.IsRunned
            'Please, wait until Microsoft disk cleanup manager finish its work and press OK.
            MsgBoxW TranslateNative(351), vbInformation
        Loop
    End If
    'remove preset
    For i = 1 To cKeys.Count
        sKey = sRootKey & "\" & GetCollectionKeyByIndex(i, cKeys)
        
        If Reg.KeyExists(0, sKey) Then
            Call Reg.DelVal(0, sKey, sParam)
        End If
    Next
    Set cKeys = Nothing
End Sub

Public Sub CheckO8Item()
    'O8 - Extra context menu items
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt
    'HKLM\Software\Microsoft\Internet Explorer\MenuExt
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO8Item - Begin"
    
    Dim hKey&, i&, sName$, lpcName&, sFile$, sHit$, result As SCAN_RESULT, pos&, bSafe As Boolean
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Internet Explorer\MenuExt"
    
    Do While HE.MoveNext
    
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, _
          KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
          
            i = 0
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            
            Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
                sName = RTrimNull(sName)
                sFile = Reg.GetString(HE.Hive, HE.Key & "\" & sName, vbNullString, HE.Redirected)
        
                If Len(sFile) = 0 Then
                    sFile = STR_NO_FILE
                Else
                    If InStr(1, sFile, "res://", vbTextCompare) = 1 Then
                        sFile = mid$(sFile, 7)
                    End If
            
                    If InStr(1, sFile, "file://", vbTextCompare) = 1 Then
                        sFile = mid$(sFile, 8)
                    End If
                    
                    pos = InStrRev(sFile, "/")
                    If pos <> 0 Then sFile = Left$(sFile, pos - 1)
                    
                    pos = InStrRev(sFile, "?")
                    If pos <> 0 Then sFile = Left$(sFile, pos - 1)
                    
                    sFile = FormatFileMissing(sFile)
                End If
                
                SignVerifyJack sFile, result.SignResult
                
                sHit = "O8 - Context menu item: " & HE.HiveNameAndSID & "\..\Internet Explorer\MenuExt\" & sName & ": (default) = " & sFile & _
                    FormatSign(result.SignResult)
                
                bSafe = False
                If WhiteListed(sFile, "EXCEL.EXE", True) Then bSafe = True 'MS Office
                If WhiteListed(sFile, "ONBttnIE.dll", True) Then bSafe = True 'MS Office
                
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                If Not IsOnIgnoreList(sHit) And (Not bSafe) Then
                    With result
                        .Section = "O8"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName
                        AddFileToFix .File, JUMP_FILE, sFile
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
                
                sName = String$(MAX_KEYNAME, 0&)
                lpcName = Len(sName)
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
        
    Loop
    
    AppendErrorLogCustom "CheckO8Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO8Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO8Item(sItem$, result As SCAN_RESULT)
    'O8 - Extra context menu items
    'O8 - Extra context menu item: [name] - html file
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt

    FixIt result
End Sub

Public Sub CheckO9Item()
    'HKLM\Software\Microsoft\Internet Explorer\Extensions
    'HKCU\..\etc
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO9Item - Begin"
    
    Dim hKey&, i&, sData$, sCLSID$, sCLSID2$, lpcName&, sFile$, sHit$, sBuf$, result As SCAN_RESULT
    Dim pos&, bSafe As Boolean
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Internet Explorer\Extensions"
    
    Do While HE.MoveNext
    
    'open root key
    If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
        i = 0
        sCLSID = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sCLSID)
        'start enum of root key subkeys (i.e., extensions)
        Do While RegEnumKeyExW(hKey, i, StrPtr(sCLSID), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
            sCLSID = TrimNull(sCLSID)
            If sCLSID = "CmdMapping" Then GoTo NextExt:
            
            'check for 'MenuText' or 'ButtonText'
            sData = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "ButtonText", HE.Redirected)
            
            'this clsid is mostly useless, always pointing to SHDOCVW.DLL
            'places to look for correct dll:
            '* Exec
            '* Script
            '* BandCLSID
            '* CLSIDExtension
            '* CLSIDExtension -> TreatAs CLSID
            '* CLSID
            '* ???
            '* actual CLSID of regkey (not used)
            sFile = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "Exec", HE.Redirected)
            If sFile = vbNullString Then
                sFile = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "Script", HE.Redirected)
                If sFile = vbNullString Then
                    sCLSID2 = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "BandCLSID", HE.Redirected)
                    sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, HE.Redirected)
                    If sFile = vbNullString Then
                        sCLSID2 = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "CLSIDExtension", HE.Redirected)
                        sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, HE.Redirected)
                        If sFile = vbNullString Then
                            sCLSID2 = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\TreatAs", vbNullString, HE.Redirected)
                            sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, HE.Redirected)
                            If sFile = vbNullString Then
                                sCLSID2 = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "CLSID", HE.Redirected)
                                sFile = Reg.GetString(HKEY_CLASSES_ROOT, "CLSID\" & sCLSID2 & "\InprocServer32", vbNullString, HE.Redirected)
                            End If
                        End If
                    End If
                End If
            End If
            
            If Len(sFile) = 0 Then
                sFile = STR_NO_FILE
            Else
                'expand %systemroot% var
                'sFile = replace$(sFile, "%systemroot%", sWinDir, , , vbTextCompare)
                sFile = UnQuote(EnvironW(sFile))
                
                'strip stuff from res://[dll]/page.htm to just [dll]
                If InStr(1, sFile, "res://", vbTextCompare) = 1 Then
                    'And (LCase$(Right$(sFile, 4)) = ".htm" Or LCase$(Right$(sFile, 4)) = "html") Then
                    sFile = mid$(sFile, 7)
                End If
                
                'remove other stupid prefixes
                If InStr(1, sFile, "file://", vbTextCompare) = 1 Then
                    sFile = mid$(sFile, 8)
                End If
                
                pos = InStrRev(sFile, "/")
                If pos <> 0 Then sFile = Left$(sFile, pos - 1)
                
                pos = InStrRev(sFile, "?")
                If pos <> 0 Then sFile = Left$(sFile, pos - 1)
                
                If InStr(1, sFile, "http:", 1) <> 1 And _
                  InStr(1, sFile, "https:", 1) <> 1 Then
                    '8.3 -> Full
                    If FileExists(sFile) Then
                        sFile = GetLongPath(EnvironW(sFile))
                    Else
                        sFile = GetLongPath(EnvironW(sFile)) & " " & STR_FILE_MISSING
                    End If
                End If
            End If
            
            bSafe = False
            If Not bIgnoreAllWhitelists And bHideMicrosoft Then
                If WhiteListed(sFile, PF_64 & "\Messenger\msmsgs.exe") Then bSafe = True
                
                If OSver.MajorMinor = 5 Then 'Win2k
                    If StrComp(sFile, sWinDir & "\web\related.htm", 1) = 0 Then bSafe = True
                End If
                If OSver.MajorMinor <= 5.2 Then 'win2k/xp/2003
                    If WhiteListed(sFile, sWinDir & "\Network Diagnostic\xpnetdiag.exe") Then bSafe = True
                End If
                If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                    If IsMicrosoftFile(sFile) Then bSafe = True
                End If
            End If
            
            If Not bSafe Then
            
              If sData = vbNullString Then sData = STR_NO_NAME
              If Left$(sData, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sData)
                If 0 <> Len(sBuf) Then sData = sBuf
              End If
              
              SignVerifyJack sFile, result.SignResult
              
              'O9 - Extra button:
              'O9-32 - Extra button:
              sHit = BitPrefix("O9", HE) & _
                " - Button: " & HE.HiveNameAndSID & "\..\" & sCLSID & ": " & sData & " - " & sFile & FormatSign(result.SignResult)
              
              If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
              
              If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O9"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sCLSID, , , HE.Redirected
                    AddRegToFix .Reg, REMOVE_VALUE, HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", sCLSID
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
              End If
            
              sData = Reg.GetString(HE.Hive, HE.Key & "\" & sCLSID, "MenuText", HE.Redirected)
            
              If Left$(sData, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sData)
                If 0 <> Len(sBuf) Then sData = sBuf
              End If

              bSafe = False
            
              If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                If OSver.MajorMinor = 5 Then 'Win2k
                  If StrComp(sFile, sWinDir & "\web\related.htm", 1) = 0 Then bSafe = True
                End If
                If OSver.MajorMinor <= 5.2 Then 'win2k/xp/2003
                    If WhiteListed(sFile, sWinDir & "\Network Diagnostic\xpnetdiag.exe") Then bSafe = True
                End If
                If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                    If IsMicrosoftFile(sFile) Then bSafe = True
                End If
              End If
            
              'don't show it again in case sdata=null
              If sData <> vbNullString And Not bSafe Then
                'O9 - Extra 'Tools' menuitem:
                'O9-32 - Extra 'Tools' menuitem:
                SignVerifyJack sFile, result.SignResult
                sHit = BitPrefix("O9", HE) & _
                  " - Tools menu item: " & HE.HiveNameAndSID & "\..\" & sCLSID & ": " & sData & " - " & sFile & FormatSign(result.SignResult)
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O9"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sCLSID, , , HE.Redirected
                        AddRegToFix .Reg, REMOVE_VALUE, HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", sCLSID
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
              End If
            End If
NextExt:
            sCLSID = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sCLSID)
            i = i + 1
        Loop
        RegCloseKey hKey
    End If
    Loop
    
    AppendErrorLogCustom "CheckO9Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO9Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO9Item(sItem$, result As SCAN_RESULT)
    'O9 - Extra buttons/Tools menu items
    'O9 - Extra button: [name] - [CLSID] - [file] [(HKCU)]
    
    FixIt result
End Sub

Public Sub CheckO10Item()
    CheckLSP
End Sub

Public Sub CheckO11Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO11Item - Begin"
    
    'HKLM\Software\Microsoft\Internet Explorer\AdvancedOptions
    Dim hKey&, i&, sSubKey$, sName$, lpcName&, sHit$, result As SCAN_RESULT
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Internet Explorer\AdvancedOptions"
    
    Do While HE.MoveNext
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
        
            sSubKey = String$(MAX_KEYNAME, 0)
            lpcName = Len(sSubKey)
            i = 0
            Do While RegEnumKeyExW(hKey, i, StrPtr(sSubKey), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
                sSubKey = TrimNull(sSubKey)
                
                If InStr("JAVA_VM.JAVA_SUN.BROWSE.ACCESSIBILITY.SEARCHING." & _
                  "HTTP1.1.MULTIMEDIA.Multimedia.CRYPTO.PRINT." & _
                  "TOEGANKELIJKHEID.TABS.INTERNATIONAL*.ACCELERATED_GRAPHICS", sSubKey) = 0 Or bIgnoreAllWhitelists Then
                  
                    sName = Reg.GetString(HE.Hive, HE.Key & "\" & sSubKey, "Text", HE.Redirected)
                  
                    If Len(sName) <> 0 Then
                        'O11 - Options group:
                        'O11-32 - Options group:
                        sHit = BitPrefix("O11", HE) & _
                          " - " & HE.HiveNameAndSID & "\..\Internet Explorer\AdvancedOptions\" & sSubKey & ": [Text] = " & sName
                
                        If bIgnoreAllWhitelists Or Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O11"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sSubKey, , , HE.Redirected
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                End If
                sSubKey = String$(MAX_KEYNAME, 0&)
                lpcName = Len(sSubKey)
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    AppendErrorLogCustom "CheckO11Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO11Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO11Item(sItem$, result As SCAN_RESULT)
    'O11 - Options group: [BLA] Blah"
    
    FixIt result
End Sub

Public Sub CheckO12Item()
    'HKLM\Software\Microsoft\Internet Explorer\Plugins\Extensions
    'HKLM\Software\Microsoft\Internet Explorer\Plugins\MIME
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO12Item - Begin"
    
    Dim hKey&, i&, sName$, sData$, sFile$, sArgs$, sHit$, lpcName&, result As SCAN_RESULT
    Dim aKey() As String, aDes() As String
    ReDim aKey(1), aDes(1)
    
    aKey(0) = "Software\Microsoft\Internet Explorer\Plugins\Extension"
    aDes(0) = "Internet Explorer\Plugins\Extension"
    
    aKey(1) = "Software\Microsoft\Internet Explorer\Plugins\MIME"
    aDes(1) = "Internet Explorer\Plugins\MIME"
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys aKey
    
    Do While HE.MoveNext
      
      If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
      
        sName = String$(MAX_KEYNAME, 0&)
        lpcName = Len(sName)
        i = 0
        
        Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
            sName = TrimNull(sName)
            sData = Reg.GetString(HE.Hive, HE.Key & "\" & sName, "Location", HE.Redirected)
            
            SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
            sFile = FormatFileMissing(sFile)
            SignVerifyJack sFile, result.SignResult
            
            'O12 - Plugin
            'O12-32 - Plugin
            sHit = BitPrefix("O12", HE) & " - " & _
              HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & "\" & sName & ": [Location] = " & ConcatFileArg(sFile, sArgs) & _
                FormatSign(result.SignResult)
            
            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O12"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName, , , HE.Redirected
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile, sArgs
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults result
            End If
            
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            i = i + 1
        Loop
        RegCloseKey hKey
      End If
    Loop
    
    AppendErrorLogCustom "CheckO12Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO12Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO12Item(sItem$, result As SCAN_RESULT)
    'O12 - Plugin for .ofb: C:\Win98\blah.dll
    'O12 - Plugin for text/blah: C:\Win98\blah.dll
    
    On Error GoTo ErrorHandler:
    
    If Not bShownToolbarWarning And ProcessExist("iexplore.exe", True) Then
        MsgBoxW Translate(330), vbExclamation
'        msgboxW "HiJackThis is about to remove a " & _
'               "plugin from " & _
'               "your system. Close all Internet " & _
'               "Explorer windows before continuing for " & _
'               "the best chance of success.", vbExclamation
        bShownToolbarWarning = True
    End If
    
    FixIt result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO12Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO13Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO13Item - Begin"
    
    Dim sDummy$, sHit$, result As SCAN_RESULT
    Dim aKey() As String, aVal() As String, aExa() As String, i As Long
    
    ReDim aKey(6)
    ReDim aVal(UBound(aKey))
    ReDim aExa(UBound(aKey))
    ReDim aDes(UBound(aKey))
    
    aKey(0) = "DefaultPrefix"
    aVal(0) = vbNullString
    aExa(0) = "http://"
    'aDes(0) = "DefaultPrefix"
    
    aKey(1) = "Prefixes"
    aVal(1) = "www"
    aExa(1) = "http://"
    'aDes(1) = "WWW Prefix"
    
    aKey(2) = "Prefixes"
    aVal(2) = "www."
    aExa(2) = vbNullString
    'aDes(2) = "WWW. Prefix"
    
    aKey(3) = "Prefixes"
    aVal(3) = "home"
    aExa(3) = "http://"
    'aDes(3) = "Home Prefix"
    
    aKey(4) = "Prefixes"
    aVal(4) = "mosaic"
    aExa(4) = "http://"
    'aDes(4) = "Mosaic Prefix"
    
    aKey(5) = "Prefixes"
    aVal(5) = "ftp"
    aExa(5) = "ftp://"
    'aDes(5) = "FTP Prefix"
    
    aKey(6) = "Prefixes"
    aVal(6) = "gopher"
    aExa(6) = "gopher://|"
    'aDes(6) = "Gopher Prefix"
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\URL"

    Do While HE.MoveNext
    
        For i = 0 To UBound(aKey)
        
            sDummy = Reg.GetString(HE.Hive, HE.Key & "\" & aKey(i), aVal(i), HE.Redirected)
            
            'exclude empty HKCU / HKU
            If Not (HE.Hive <> HKLM And Len(sDummy) = 0) Then
            
                If Not inArraySerialized(sDummy, aExa(i), "|", , , vbBinaryCompare) Or Not bHideMicrosoft Then
                    
                    sHit = BitPrefix("O13", HE) & " - " & HE.HiveNameAndSID & "\..\URL\" & aKey(i) & _
                        ": [" & aVal(i) & "] = " & sDummy
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O13"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key & "\" & aKey(i), aVal(i), SplitSafe(aExa(i), "|")(0), HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckO13Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO13Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO13Item(sItem$, result As SCAN_RESULT)
    'defaultprefix fix
    'O13 - DefaultPrefix: http://www.hijacker.com/redir.cgi?
    'O13 - [WWW/Home/Mosaic/FTP/Gopher] Prefix: ..

    FixIt result
End Sub

Public Sub CheckO14Item()
    'O14 - Reset Websettings check
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO14Item - Begin"
    
    Dim sLine$, sHit$, ff%
    Dim sStartPage$, sSearchPage$, sMsStartPage$
    Dim sSearchAssis$, sCustSearch$
    Dim sFile$, aLogStrings() As String, i&
    
    sFile = sWinDir & "\inf\iereset.inf"
    
    If Not FileExists(sFile) Then Exit Sub
    If FileLenW(sFile) = 0 Then Exit Sub
    
    aLogStrings = ReadFileToArray(sFile, FileGetTypeBOM(sFile) = CP_UTF16LE)
    
    For i = 0 To UBound(aLogStrings)
        sLine = aLogStrings(i)
        
            If InStr(sLine, "SearchAssistant") > 0 Then
                sSearchAssis = mid$(sLine, InStr(sLine, "http://"))
                sSearchAssis = Left$(sSearchAssis, Len(sSearchAssis) - 1)
            End If
            If InStr(sLine, "CustomizeSearch") > 0 Then
                sCustSearch = mid$(sLine, InStr(sLine, "http://"))
                sCustSearch = Left$(sCustSearch, Len(sCustSearch) - 1)
            End If
            If InStr(sLine, "START_PAGE_URL=") = 1 And _
               InStr(sLine, "MS_START_PAGE_URL") = 0 Then
                sStartPage = mid$(sLine, InStr(sLine, "=") + 1)
                sStartPage = UnQuote(sStartPage)
            End If
            If InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sSearchPage = mid$(sLine, InStr(sLine, "=") + 1)
                sSearchPage = UnQuote(sSearchPage)
            End If
            If InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sMsStartPage = mid$(sLine, InStr(sLine, "=") + 1)
                sMsStartPage = UnQuote(sMsStartPage)
            End If
    Next
    
    'SearchAssistant = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm
    If (sSearchAssis <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm" And _
      Len(sSearchAssis) <> 0) Or Not bHideMicrosoft Then
        sHit = "O14 - IERESET.INF: SearchAssistant = " & sSearchAssis
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'CustomizeSearch = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm
    If (sCustSearch <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm" And _
      Len(sCustSearch) <> 0) Or Not bHideMicrosoft Then
        sHit = "O14 - IERESET.INF: CustomizeSearch = " & sCustSearch
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'SEARCH_PAGE_URL = http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch
    If (sSearchPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch" And _
      sSearchPage <> "http://www.msn.com" And _
      sSearchPage <> "https://www.msn.com") Or Not bHideMicrosoft Then
        sHit = "O14 - IERESET.INF: [Strings] SEARCH_PAGE_URL = " & sSearchPage
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'START_PAGE_URL  = http://www.msn.com
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If (sStartPage <> "http://www.msn.com" And _
       sStartPage <> "https://www.msn.com" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome") Or Not bHideMicrosoft Then
        sHit = "O14 - IERESET.INF: [Strings] START_PAGE_URL = " & sStartPage
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'MS_START_PAGE_URL=http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '(=START_PAGE_URL) http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If sMsStartPage <> vbNullString Then
        If (sMsStartPage <> "http://www.msn.com" And _
           sMsStartPage <> "https://www.msn.com" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome") Or Not bHideMicrosoft Then
            sHit = "O14 - IERESET.INF: [Strings] MS_START_PAGE_URL = " & sMsStartPage
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
        End If
    End If
    
    AppendErrorLogCustom "CheckO14Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO14Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO14Item(sItem$, result As SCAN_RESULT)
    'resetwebsettings fix
    'O14 - IERESET.INF: [item]=[URL]
    
    On Error GoTo ErrorHandler:
    'sItem - not used
    Dim sLine$, sFixedIeResetInf$, ff%
    Dim i&, aLogStrings() As String, sFile$, isUnicode As Boolean
    
    sFile = sWinDir & "\INF\iereset.inf"
    
    If Not FileExists(sFile) Then Exit Sub
    
    BackupFile result, sFile
    
    isUnicode = (FileGetTypeBOM(sFile) = CP_UTF16LE)
    aLogStrings = ReadFileToArray(sFile, IIf(isUnicode, True, False))
    
    For i = 0 To UBound(aLogStrings)
        sLine = aLogStrings(i)

            If InStr(sLine, "SearchAssistant") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""SearchAssistant"",0,""" & _
                    vbNullString & """" & vbCrLf
            ElseIf InStr(sLine, "CustomizeSearch") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""CustomizeSearch"",0,""" & _
                    vbNullString & """" & vbCrLf
            ElseIf InStr(sLine, "START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "START_PAGE_URL=""" & "https://www.msn.com" & """" & vbCrLf
            ElseIf InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "SEARCH_PAGE_URL=""" & "https://www.msn.com" & """" & vbCrLf
            ElseIf InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "MS_START_PAGE_URL=""" & "https://www.msn.com" & """" & vbCrLf
            Else
                sFixedIeResetInf = sFixedIeResetInf & sLine & vbCrLf
            End If
        
    Next
    sFixedIeResetInf = Left$(sFixedIeResetInf, Len(sFixedIeResetInf) - 2)   '-CrLf
    
    DeleteFileForce sFile
    
    ff = FreeFile()
    
    If isUnicode Then
        Dim b() As Byte
        b() = ChrW$(-257) & sFixedIeResetInf
        Open sFile For Binary Access Write As #ff
        Put #ff, , b()
    Else
        Open sFile For Output As #ff
        Print #ff, sFixedIeResetInf
    End If
    
    Close #ff
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO14Item", "sItem=", sItem
    Close #ff
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO15Item()
    'the value * / http / https denotes the protocol for which the rule is valid. It's:
    '2 for Trusted Zone;
    '4 for Restricted Zone.
    
    'Checks:
    '* ZoneMap\Domains          - trusted domains
    '* ZoneMap\Ranges           - trusted IPs and IP ranges
    '* ZoneMap\ProtocolDefaults - what zone rules does a protocol obey
    '* ZoneMap\EscDomains       - trusted domains for Enhanced Security Configuration
    '* ZoneMap\EscRanges        - trusted IPs and IP ranges for ESC
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO15Item - Begin"
    
    Dim sDomains$(), sSubDomains$(), vProtocol, sProtPrefix$
    Dim i&, j&, sHit$, sAlias$, sIPRange$, bSafe As Boolean, result As SCAN_RESULT
    Dim dURL As clsTrickHashTable, aResult() As SCAN_RESULT, iRes As Long, iCur As Long, sURL As String
    Set dURL = New clsTrickHashTable
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains"
    
    'enum all domains
    'value = http - only http on subdomain is trusted
    'value = https - only https on subdomain is trusted
    'value = * - both
    
    Do While HE.MoveNext
        sDomains = Split(Reg.EnumSubKeys(HE.Hive, HE.Key, HE.Redirected), "|")
        If UBound(sDomains) > -1 Then
            If StrEndWith(HE.Key, "EscDomains") Then
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - " & HE.HiveNameAndSID & "\..\ESC Trusted Zone: "
            Else
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - " & HE.HiveNameAndSID & "\..\Trusted Zone: "
            End If
            For i = 0 To UBound(sDomains)
                bSafe = False
                If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                    bSafe = StrBeginWithArray(sDomains(i), aSafeRegDomains)
                End If
                If Not bSafe Then
                    sSubDomains = Split(Reg.EnumSubKeys(HE.Hive, HE.Key & "\" & sDomains(i), HE.Redirected), "|")
                    If UBound(sSubDomains) <> -1 Then
                        'list any trusted subdomains for main domain
                        For j = 0 To UBound(sSubDomains)
                            
                            For Each vProtocol In Array("*", "http", "https")
                                Select Case vProtocol
                                    Case "*": sProtPrefix = "*."
                                    Case "http": sProtPrefix = "http://"
                                    Case "https": sProtPrefix = "https://"
                                End Select
                                If Reg.GetDword(HE.Hive, HE.Key & "\" & sDomains(i) & "\" & sSubDomains(j), CStr(vProtocol), HE.Redirected) = 2 Then
                                
                                    bSafe = False
                                    If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                                        bSafe = StrBeginWithArray(sProtPrefix & sSubDomains(j) & "." & sDomains(i), aSafeRegDomains)
                                    End If
                                    
                                    If Not bSafe Then
                                        sURL = sSubDomains(j) & "." & sDomains(i)
                                        sAlias = BitPrefix("O15", HE) & " - Trusted Zone: "
                                        sHit = sAlias & sProtPrefix & sURL
                                        
                                        If Not IsOnIgnoreList(sHit) Then
                                        
                                            'concat several identical URLs to single log line
                                            If dURL.Exists(sURL) Then
                                                iCur = dURL(sURL)
                                            Else
                                                iRes = iRes + 1
                                                iCur = iRes
                                                ReDim Preserve aResult(iRes)
                                                dURL.Add sURL, iCur
                                            End If
                                            
                                            With aResult(iCur)
                                                .Section = "O15"
                                                .HitLineW = sHit
                                                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i) & "\" & sSubDomains(j), , , HE.Redirected
                                                .CureType = REGISTRY_BASED
                                            End With
                                            'AddToScanResults result
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    'list main domain as well if that's trusted too (*grumble*)
                    For Each vProtocol In Array("*", "http", "https")
                        Select Case vProtocol
                            Case "*": sProtPrefix = "*."
                            Case "http": sProtPrefix = "http://"
                            Case "https": sProtPrefix = "https://"
                        End Select
                        If Reg.GetDword(HE.Hive, HE.Key & "\" & sDomains(i), CStr(vProtocol), HE.Redirected) = 2 Then
                        
                            bSafe = False
                            If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                                If StrBeginWithArray(sProtPrefix & sDomains(i), aSafeRegDomains) Then bSafe = True
                            End If
                        
                            If Not bSafe Then
                                sURL = sDomains(i)
                                sAlias = BitPrefix("O15", HE) & " - Trusted Zone: "
                                sHit = sAlias & sProtPrefix & sURL
                                
                                If Not IsOnIgnoreList(sHit) Then
                                
                                    'concat several identical URLs to single log line
                                    If dURL.Exists(sURL) Then
                                        iCur = dURL(sURL)
                                    Else
                                        iRes = iRes + 1
                                        iCur = iRes
                                        ReDim Preserve aResult(iRes)
                                        dURL.Add sURL, iCur
                                    End If
                                
                                    With aResult(iCur)
                                        .Section = "O15"
                                        .HitLineW = sHit
                                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i), , , HE.Redirected
                                        .CureType = REGISTRY_BASED
                                    End With
                                    'AddToScanResults result
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Loop
    
    For i = 1 To iRes
        AddToScanResults aResult(i)
    Next
        
    Set dURL = Nothing
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscRanges"
    
    'enum all IP ranges
    Do While HE.MoveNext
        sDomains = Split(Reg.EnumSubKeys(HE.Hive, HE.Key, HE.Redirected), "|")
        If UBound(sDomains) > -1 Then
            If StrEndWith(HE.Key, "EscRanges") Then
                sAlias = BitPrefix("O15", HE) & " - " & HE.HiveNameAndSID & "\..\ESC Trusted IP range: "
            Else
                sAlias = BitPrefix("O15", HE) & " - " & HE.HiveNameAndSID & "\..\Trusted IP range: "
            End If
            For i = 0 To UBound(sDomains)
                sIPRange = Reg.GetString(HE.Hive, HE.Key & "\" & sDomains(i), ":Range", HE.Redirected)
                If Left$(sDomains(i), 5) = "Range" And sIPRange <> vbNullString Then
                    For Each vProtocol In Array("*", "http", "https")
                        Select Case vProtocol
                        Case "*": sProtPrefix = "*."
                        Case "http": sProtPrefix = "http://"
                        Case "https": sProtPrefix = "https://"
                        End Select
                        If Reg.GetDword(HE.Hive, HE.Key & "\" & sDomains(i), CStr(vProtocol), HE.Redirected) = 2 Then
                            sHit = sAlias & sProtPrefix & sIPRange
                            If Not IsOnIgnoreList(sHit) Then
                                With result
                                    .Section = "O15"
                                    .HitLineW = sHit
                                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i), , , HE.Redirected
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults result
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Loop
    
    'check all ProtocolDefaults values
    
    Dim sZoneNames$(), sProtVals$(), lProtZoneDefs&(11), lProtZones&(11), LastIndex&
    sZoneNames = Split("My Computer|Intranet|Trusted|Internet|Restricted|Unknown", "|")
    sProtVals = Split("@ivt|file|ftp|http|https|shell|ldap|news|nntp|oecmd|snews|knownfolder", "|")
    
    '0 = My Computer
    '1 = Intranet
    '2 = Trusted
    '3 = Internet
    '4 = Restricted
    '5 = Unknown
    
    'binding protocol to zone
    lProtZoneDefs(0) = 1 '@ivt '2k+
    lProtZoneDefs(1) = 3 'file '2k+
    lProtZoneDefs(2) = 3 'ftp '2k+
    lProtZoneDefs(3) = 3 'http '2k+
    lProtZoneDefs(4) = 3 'https '2k+
    lProtZoneDefs(5) = 0 'shell 'XP+
    lProtZoneDefs(6) = 4 'ldap '(HKLM only) 'Vista+
    lProtZoneDefs(7) = 4 'news '(HKLM only) 'Vista+
    lProtZoneDefs(8) = 4 'nntp '(HKLM only) 'Vista+
    lProtZoneDefs(9) = 4 'oecmd '(HKLM only) 'Vista+
    lProtZoneDefs(10) = 4 'snews '(HKLM only) 'Vista+
    lProtZoneDefs(11) = 0 'knownfolder '7+
    
    If OSver.MajorMinor = 5 Then 'Win2k
        HE.Init HE_HIVE_ALL, HE_SID_USER Or HE_SID_NO_VIRTUAL
    ElseIf OSver.MajorMinor = 5.2 And OSver.IsServer Then 'Win 2003, 2003 R2
        HE.Init HE_HIVE_ALL, HE_SID_USER Or HE_SID_NO_VIRTUAL
    Else 'Vista+
        HE.Init HE_HIVE_ALL, HE_SID_DEFAULT Or HE_SID_USER Or HE_SID_NO_VIRTUAL
    End If
    
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\ProtocolDefaults"
    
    If OSver.IsWindows7OrGreater Then
        LastIndex = 11
    ElseIf OSver.IsWindowsVistaOrGreater Then
        LastIndex = 10
    ElseIf OSver.IsWindowsXPOrGreater Then
        LastIndex = 5
    Else
        LastIndex = 4
    End If
    
    Do While HE.MoveNext
        For i = 0 To LastIndex
            bSafe = False
            lProtZones(i) = Reg.GetDword(HE.Hive, HE.Key, sProtVals(i), HE.Redirected)
            
            If lProtZones(i) = 0 Then
                If Not Reg.ValueExists(HE.Hive, HE.Key, sProtVals(i), HE.Redirected) Then
                    If i >= 6 And i <= 10 And HE.Hive <> HKLM Then
                        bSafe = True
                    Else
                        lProtZones(i) = 5 'Unknown
                    End If
                End If
                
                If lProtZones(i) = 5 Then
                    If sProtVals(i) = "knownfolder" And OSver.MajorMinor = 6.1 Then bSafe = True
                    If HE.UserName = "UpdatusUser" Then bSafe = True
                    If HE.UserName = "unknown" Then bSafe = True
                End If
            End If
            
            If Not bSafe Then
                If lProtZones(i) < 0 Or lProtZones(i) > 5 Then lProtZones(i) = 5 'Unknown
                If lProtZones(i) = 5 Then
                    If InStr(1, HE.UserName, "MSSQL", 1) <> 0 Then
                        bSafe = True
                    ElseIf InStr(1, HE.UserName, "MsDtsServer", 1) <> 0 Then
                        bSafe = True
                    ElseIf InStr(1, HE.UserName, "defaultuser0", 1) <> 0 Then
                        bSafe = True
                    ElseIf InStr(1, HE.UserName, "Acronis Agent User", 1) <> 0 Then
                        bSafe = True
                        '// TODO: improve it (logon as service)
                    ElseIf Reg.GetDword(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & HE.SID, "State") = 0 Then
                        bSafe = True
                    End If
                End If
                
                If Not bSafe And (lProtZones(i) <> lProtZoneDefs(i)) Then 'check for legit
                    
                    sHit = BitPrefix("O15", HE) & " - " & HE.HiveNameAndSID & "\..\ProtocolDefaults: " & _
                        " - [" & sProtVals(i) & "] protocol is in " & sZoneNames(lProtZones(i)) & " Zone, should be " & sZoneNames(lProtZoneDefs(i)) & " Zone" & _
                        IIf(HE.IsSidUser, " (User: '" & HE.UserName & "')", vbNullString)
                        
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O15"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sProtVals(i), lProtZoneDefs(i), HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckO15Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO15Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO15Item(sItem$, result As SCAN_RESULT)
'    'O15 - Trusted Zone: free.aol.com (HKLM)
'    'O15 - Trusted Zone: http://free.aol.com
'    'O15 - Trusted IP range: 66.66.66.66 (HKLM)
'    'O15 - Trusted IP range: http://66.66.66.*
'    'O15 - ESC Trusted Zone: free.aol.com (HKLM)
'    'O15 - ESC Trusted IP range: 66.66.66.66
'    'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)

    FixIt result
End Sub

Public Sub CheckO16Item()
    'O16 - Downloaded Program Files
    
    'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Internet Settings,ActiveXCache
    'is location of actual %WINDIR%\DPF\ folder
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO16Item - Begin"
    
    Dim sName$, sFriendlyName$, sCodebase$, i&, j&, hKey&, lpcName&, sHit$, result As SCAN_RESULT
    Dim sOSD$, sInf$, sInProcServer32$, aValue() As String
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Code Store Database\Distribution Units"
    
    Do While HE.MoveNext
    
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
    
            sName = String$(MAX_KEYNAME, 0)
            lpcName = Len(sName)
            i = 0
    
            Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
      
                sName = Left$(sName, InStr(sName, vbNullChar) - 1)
        
                sCodebase = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\DownloadInformation", "CODEBASE", HE.Redirected)
        
                If (InStr(sCodebase, "http://www.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://webresponse.one.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://rtc.webresponse.one.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://office.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://officeupdate.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://protect.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://dql.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://codecs.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://download.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://windowsupdate.microsoft.com") <> 1 And _
                  InStr(sCodebase, "http://v4.windowsupdate.microsoft.com") <> 1) _
                  Or Not bHideMicrosoft Then

                  'a DPF object can consist of:
                  '* DPF regkey           -> sDPFKey
                  '* CLSID regkey         -> CLSID\ & sName
                  '* OSD file             -> sOSD = Reg.GetString
                  '* INF file             -> sINF = Reg.GetString
                  '* InProcServer32 file  -> sIPS = Reg.GetString
                    
                    If Left$(sName, 1) = "{" And Right$(sName, 1) = "}" Then
                        Call GetFileByCLSID(sName, sInProcServer32, sFriendlyName, HE.Redirected, HE.SharedKey)
                    End If
                    
                    'not http ?
                    If mid$(sCodebase, 2, 1) = ":" Then
                        sCodebase = FormatFileMissing(PathNormalize(sCodebase))
                    End If
                    
                    ' "O16 - DPF: "
                    ' CODEBASE - is a URL
                    sHit = BitPrefix("O16", HE) & " - DPF: " & HE.HiveNameAndSID & "\..\" & _
                      sName & "\DownloadInformation: " & sFriendlyName & " [CODEBASE] = " & sCodebase
                    
                    'if file
                    If mid$(sCodebase, 2, 1) = ":" Then
                        SignVerifyJack sCodebase, result.SignResult
                        sHit = sHit & FormatSign(result.SignResult)
                        If g_bCheckSum Then
                             sHit = sHit & GetFileCheckSum(sCodebase)
                        End If
                    End If
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O16"
                            .HitLineW = sHit
                            
                            sOSD = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\DownloadInformation", "OSD", HE.Redirected)
                            sInf = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\DownloadInformation", "INF", HE.Redirected)

                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sInProcServer32 'Or UNREG_DLL
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sOSD
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sInf
                            
                            AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sName, , , REG_REDIRECTION_BOTH
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName, , , HE.Redirected
                            
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility\" & sName, , , HE.Redirected
                            AddRegToFix .Reg, REMOVE_KEY, HKCU, "Software\Microsoft\Windows\CurrentVersion\Ext\Stats\" & sName, , , HE.Redirected
                            
                            For j = 1 To Reg.EnumValuesToArray(HKLM, HE.Key & "\" & sName & "\Contains\Files", aValue, HE.Redirected)
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, aValue(j)
                            Next
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
                
                i = i + 1
                sName = String$(MAX_KEYNAME, 0)
                lpcName = Len(sName)
            
            Loop
            
            RegCloseKey hKey
        End If
    Loop
    
    AppendErrorLogCustom "CheckO16Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO16Item"
    RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO16Item(sItem$, result As SCAN_RESULT)
    'O16 - DPF: {0000000} (shit toolbar) - http://bla.com/bla.dll
    'O16 - DPF: Plugin - http://bla.com/bla.dll
    
    FixIt result
End Sub

Public Sub CheckO17Item()
    'check 'domain' and 'domainname' values in:
    'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters
    'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\*
    'HKLM\Software\Microsoft\Windows\CurrentVersion\Telephony
    'HKLM\System\CurrentControlSet\Services\VxD\MSTCP
    'and all values in other ControlSet's as well
    '
    'new one from UltimateSearch: value 'SearchList' in
    'HKLM\System\CurrentControlSet\Services\VxD\MSTCP
    '
    'just in case: NameServer as well, CoolWebSearch
    'maybe using this
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO17Item - Begin"
    
    Dim hKey&, i&, j&, sDomain$, sHit$, sParam$, vParam, CSKey$, N&, sData$, aNames() As String
    Dim UseWow, Wow6432Redir As Boolean, result As SCAN_RESULT, Data() As String, sTrimChar As String
    Dim TcpIpNameServers() As String: ReDim TcpIpNameServers(0)
    Dim aKeyDomain() As String
    ReDim aKeyDomain(0 To 1) As String
    Dim sProviderDNS As String
    Dim ccsID As Long
    
'    'get target of CurrentControlSet
'    If Reg.IsKeySymLink(HKLM, "SYSTEM\CurrentControlSet", sSymTarget) Then
'        If IsNumeric(Right$(sSymTarget, 3)) Then
'            ccsID = CLng(Right$(sSymTarget, 3))
'        End If
'    End If

    ccsID = Reg.GetDword(HKLM, "SYSTEM\Select", "Current")
    
    'these keys are x64 shared
    aKeyDomain(0) = "Services\Tcpip\Parameters"
    aKeyDomain(1) = "Services\VxD\MSTCP"
    
    For j = 0 To 99    ' 0 - is CCS
    
        CSKey = IIf(j = 0, "System\CurrentControlSet", "System\ControlSet" & Format$(j, "000"))
        
        If j > 0 Then
            If Not Reg.KeyExists(HKEY_LOCAL_MACHINE, CSKey) Then Exit For
            If j = ccsID Then GoTo Continue
        End If
    
        For Each vParam In Array("Domain", "DomainName", "SearchList", "NameServer")
            sParam = vParam
            
            For N = 0 To UBound(aKeyDomain)
                'HKLM\System\CCS\Services\Tcpip\Parameters,Domain
                'HKLM\System\CCS\Services\Tcpip\Parameters,DomainName
                'HKLM\System\CCS\Services\VxD\MSTCP,Domain
                'HKLM\System\CCS\Services\VxD\MSTCP,DomainName
                'new one from UltimateSearch!
                'HKLM\System\CCS\Services\VxD\MSTCP,SearchList
                'HKLM\System\CCS\Services\VxD\MSTCP,SearchList
                'HKLM\System\CCS\Services\Tcpip\Parameters,SearchList
                'HKLM\System\CCS\Services\Tcpip\Parameters,NameServer
                sData = Reg.GetString(HKEY_LOCAL_MACHINE, CSKey & "\" & aKeyDomain(N), sParam)
                
                If Len(sData) <> 0 Then
                    
                    ReDim Data(0)
                    Data(0) = sData
                    
                    If sParam = "NameServer" Then
                        Data = SplitByMultiDelims(Trim$(sData), True, sTrimChar, " ", ",")
                        ArrayRemoveEmptyItems Data
                    End If
                    
                    For i = 0 To UBound(Data)
                    
                        sData = Data(i)
                    
                        sHit = "O17 - HKLM\" & IIf(j = 0, "System\CCS", CSKey) & "\" & aKeyDomain(N) & ": [" & sParam & "] = " & sData
                    
                        If sParam = "NameServer" Then
                            sProviderDNS = GetCollectionItemByKey(sData, colSafeDNS)
                            If Len(sProviderDNS) <> 0 Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                        End If
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O17"
                                .HitLineW = sHit
                                'AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, CSKey & "\" & sKeyDomain(n), sParam
                                
                                AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                    HKEY_LOCAL_MACHINE, CSKey & "\" & aKeyDomain(N), sParam, _
                                    , , , CStr(Data(i)), vbNullString, sTrimChar
                                
                                AddCustomToFix .Custom, CUSTOM_ACTION_SPECIFIC, sData
                                
                                .CureType = REGISTRY_BASED Or CUSTOM_BASED
                            End With
                            AddToScanResults result
                        End If
                    
                    Next
                    
                End If
            Next
            
            'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\.. subkeys
            'HKLM\System\CS*\Services\Tcpip\Parameters\Interfaces\.. subkeys
            
            For N = 1 To Reg.EnumSubKeysToArray(HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces", aNames)
                
                sData = Reg.GetString(HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces\" & aNames(N), sParam)
                If sData <> vbNullString Then
                
                    ReDim Data(0)
                    Data(0) = sData
                    
                    If sParam = "NameServer" Then
                        
                        'Split lines like:
                        'O17 - HKLM\System\CCS\Services\Tcpip\..\{19B2C21E-CA09-48A1-9456-E4191BE91F00}: NameServer = 89.20.100.53 83.219.25.69
                        'O17 - HKLM\System\CCS\Services\Tcpip\..\{2A220B45-7A12-4A0B-92F0-00254794215A}: NameServer = 192.168.1.1,8.8.8.8
                        'into several separate
                        
                        Data = SplitByMultiDelims(Trim$(sData), True, sTrimChar, " ", ",")
                        ArrayRemoveEmptyItems Data
                        
                        For i = 0 To UBound(Data)
                            ReDim Preserve TcpIpNameServers(UBound(TcpIpNameServers) + 1)   'for using in filtering DNS DHCP later
                            TcpIpNameServers(UBound(TcpIpNameServers)) = Data(i)
                        Next
                    End If
                    
                    For i = 0 To UBound(Data)
                        
                        sHit = "O17 - HKLM\" & IIf(j = 0, "System\CCS", CSKey) & "\Services\Tcpip\..\" & aNames(N) & ": [" & sParam & "] = " & Data(i)
                        
                        If sParam = "NameServer" Then
                            sProviderDNS = GetCollectionItemByKey(CStr(Data(i)), colSafeDNS)
                            If Len(sProviderDNS) <> 0 Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                        End If
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O17"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                    HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces\" & aNames(N), sParam, _
                                    , , , CStr(Data(i)), vbNullString, sTrimChar
                                
                                AddCustomToFix .Custom, CUSTOM_ACTION_SPECIFIC, CStr(Data(i))
                                
                                .CureType = REGISTRY_BASED Or CUSTOM_BASED
                            End With
                            AddToScanResults result
                        End If
                    Next
                End If
            Next
        Next
Continue:
    Next
    
    Dim sTelephonyDomain$
    sTelephonyDomain = "Software\Microsoft\Windows\CurrentVersion\Telephony"
    
    For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If bIsWin32 And Wow6432Redir Then Exit For
    
        'HKLM\Software\MS\Windows\CurVer\Telephony,Domain
        'HKLM\Software\MS\Windows\CurVer\Telephony,DomainName
        For Each vParam In Array("Domain", "DomainName")
            sParam = vParam
            sDomain = Reg.GetString(HKEY_LOCAL_MACHINE, sTelephonyDomain, sParam, Wow6432Redir)
            If sDomain <> vbNullString Then
                'O17 - HKLM\Software\..\Telephony:
                sHit = IIf(bIsWin32, "O17", IIf(Wow6432Redir, "O17-32", "O17")) & " - HKLM\Software\..\Telephony: [" & sParam & "] = " & sDomain
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O17"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, sTelephonyDomain, sParam
                        AddCustomToFix .Custom, CUSTOM_ACTION_SPECIFIC, sParam
                        .CureType = REGISTRY_BASED Or CUSTOM_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Next
    
    Dim DNS() As String
    
    If GetDNS(DNS) Then
        For i = 0 To UBound(DNS)
            If Len(DNS(i)) <> 0 Then
                sHit = "O17 - DHCP DNS " & i + 1 & ": " & DNS(i)
                
                sProviderDNS = GetCollectionItemByKey(DNS(i), colSafeDNS)
                If Len(sProviderDNS) <> 0 Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                
                If Not IsOnIgnoreList(sHit) Then
                    
                    With result
                        .Section = "O17"
                        .HitLineW = sHit
                        AddCustomToFix .Custom, CUSTOM_ACTION_SPECIFIC, DNS(i)
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    End If
    
    AppendErrorLogCustom "CheckO17Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO17Item"
    If hKey <> 0 Then RegCloseKey hKey
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO17Item(sItem$, result As SCAN_RESULT)
    'O17 - Domain hijack
    'O17 - HKLM\System\CCS\Services\VxD\MSTCP: Domain[Name] = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\Parameters: Domain[Name] = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\..\{0000}: Domain[Name] = blah
    '                  CS1
    '                  CS2
    '                  ...
    'O17 - HKLM\Software\..\Telephony: SearchList = blah
    'O17 - HKLM\System\CCS\Services\VxD\MSTCP: SearchList = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\Parameters: SearchList = blah
    'O17 - HKLM\System\CCS\Services\Tcpip\..\{0000}: SearchList = blah
    '                  CS1
    '                  CS2
    '                  ...
    'ditto for NameServer
    
    If StrBeginWith(sItem, "O17 - DHCP DNS:") Then
        'Cure for this object is not provided: []
        'You need to manually set the DNS address on the router, which is issued to you by provider.
        MsgBoxW Replace$(TranslateNative(349), "[]", sItem), vbExclamation
        Exit Sub
    End If
    
    FixIt result
End Sub

Public Sub CheckO18Item()
    'enumerate everything in HKCR\Protocols\Handler
    'enumerate everything in HKCR\Protocols\Filters (section 2)
    'keys are x64 shared
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO18Item - Begin"
    
    ScanPrinterPorts
    
    Dim hKey&, i&, sName$, sCLSID$, sFile$, lpcName&, sHit$, result As SCAN_RESULT
    Dim bShared As Boolean, vkt As KEY_REDIRECTION_INFO, bSafe As Boolean, sFixKey As String
    Dim bBySubKey As Boolean, aSubKey() As String, j&, sDefCLSID As String, sDefCLSID_all As String, vDefCLSID As Variant, sDefFile As String
    Dim sHash As String
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Classes\Protocols\Handler"
    
    Do While HE.MoveNext
        vkt = Reg.GetKeyRedirectionType(HE.Hive, HE.Key)
        bShared = (vkt And KEY_REDIRECTION_SHARED) Or (vkt And KEY_REDIRECTION_REFLECTED)
        
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            i = 0
            Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
                sName = TrimNull(sName)
                sCLSID = UCase$(Reg.GetString(HE.Hive, HE.Key & "\" & sName, "CLSID", HE.Redirected))
                
                sFile = vbNullString
                If Len(sCLSID) <> 0 Then
                    Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, bShared)
                End If
                If Len(sCLSID) = 0 Then sCLSID = STR_NO_CLSID
                If Len(sFile) = 0 Then sFile = STR_NO_FILE
                
                bSafe = False
                
                If oDict.dSafeProtocols.Exists(sName) Then
                    sDefCLSID_all = oDict.dSafeProtocols(sName)
                    If bHideMicrosoft Then
                        If IsMicrosoftFile(sFile) Then
                            bSafe = True
                        Else
                            If StrComp(GetFileName(sFile, True), "MSITSS.DLL", 1) = 0 Then
                                'https://www.virustotal.com/gui/file/7941cd077397bc18e6dc46a478e196f25fd56ee0ad7ebcaede6b360c77d57de1/detection
                                'ms office 2003
                                If StrComp(GetFileSHA1(sFile, , True), "19255cb7154b30697431ec98e9c9698e39d80c7d", 1) = 0 Then bSafe = True
                            End If
                        End If
                    End If
                Else
                    sDefCLSID_all = vbNullString
                    If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                        If bHideMicrosoft Then
                            If IsMicrosoftFile(sFile) Then bSafe = True
                        End If
                    End If
                End If
                
                If Not bSafe And Len(sFile) <> 0 Then
                    If StrComp(GetFileNameAndExt(sFile), "MSDAIPP.DLL", 1) = 0 Then
                        sHash = GetFileSHA1(sFile, , True)
                        If sHash = "3F61C6698DEA48E0CA8C05019CF470E54B4782C6" Or _
                          sHash = "A5F1FE61B58F28A65AE189AF14749DD241A9830D" Then bSafe = True
                    End If
                End If
                
                'Repeat for subkey
                
                If Not bSafe Then
                    'Protocols key can contain several subkeys with similar contents
                    bBySubKey = False
                    If sCLSID = STR_NO_CLSID Then
                        If Reg.KeyHasSubKeys(HE.Hive, HE.Key & "\" & sName, HE.Redirected) Then
                            bBySubKey = True
                            For j = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key & "\" & sName, aSubKey, HE.Redirected)
                                sCLSID = UCase$(Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\" & aSubKey(j), "CLSID", HE.Redirected))
                                sFile = vbNullString
                                If Len(sCLSID) <> 0 Then
                                    Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, bShared)
                                End If
                                If Len(sCLSID) = 0 Then sCLSID = STR_NO_CLSID
                                If Len(sFile) = 0 Then sFile = STR_NO_FILE
                                
                                If oDict.dSafeProtocols.Exists(sName) Then
                                    sDefCLSID_all = oDict.dSafeProtocols(sName)
                                    If bHideMicrosoft Then
                                        If IsMicrosoftFile(sFile) Then bSafe = True
                                    End If
                                Else
                                    sDefCLSID_all = vbNullString
                                    If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                                        If bHideMicrosoft Then
                                            If IsMicrosoftFile(sFile) Then bSafe = True
                                        End If
                                    End If
                                End If
                                
                                If Not bSafe And Len(sFile) <> 0 Then
                                    If StrComp(GetFileNameAndExt(sFile), "MSDAIPP.DLL", 1) = 0 Then
                                        sHash = GetFileSHA1(sFile, , True)
                                        If sHash = "3F61C6698DEA48E0CA8C05019CF470E54B4782C6" Or _
                                          sHash = "A5F1FE61B58F28A65AE189AF14749DD241A9830D" Then bSafe = True
                                    End If
                                End If
                                
                                If Not bSafe Then
                                    sHit = BitPrefix("O18", HE) & " - " & HE.KeyAndHivePhysical & "\" & sName & "\" & aSubKey(j) & _
                                        ": [CLSID] = " & sCLSID & " - " & sFile
                                    sFixKey = HE.Key & "\" & sName & "\" & aSubKey(j)
                                    GoSub labelFix:
                                End If
                            Next
                        End If
                    End If
                    
                    If Not bBySubKey Then
                        'HKCU often has empty protocol keys, so skip them
                        
                        'If Not (sCLSID = "(no CLSID)" And (HE.Hive = HKCU Or HE.Hive = HKU)) Then
                
                            sHit = BitPrefix("O18", HE) & " - " & HE.KeyAndHivePhysical & "\" & sName & _
                                ": [CLSID] = " & sCLSID & " - " & sFile
                            sFixKey = HE.Key & "\" & sName
                            
                            GoSub labelFix:
                        'End If
                    End If
                End If
                
                sName = String$(MAX_KEYNAME, 0)
                lpcName = Len(sName)
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    '-------------------
    'Filters:
    
    HE.Clear
    HE.AddKey "Software\Classes\Protocols\Filter"
    
    hKey = 0
    sCLSID = vbNullString
    sFile = vbNullString
    
    Do While HE.MoveNext
    
        vkt = Reg.GetKeyRedirectionType(HE.Hive, HE.Key)
        bShared = (vkt And KEY_REDIRECTION_SHARED) Or (vkt And KEY_REDIRECTION_REFLECTED)
        
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            i = 0
            Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
                sName = TrimNull(sName)
                sCLSID = Reg.GetString(HE.Hive, HE.Key & "\" & sName, "CLSID", HE.Redirected)
                
                If Len(sCLSID) = 0 Then
                    sCLSID = STR_NO_CLSID
                    sFile = STR_NO_FILE
                Else
                    Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, bShared)
                End If
                
                bSafe = False
                If oDict.dSafeFilters.Exists(sName) Then
                    sDefCLSID_all = oDict.dSafeFilters(sName)
                    If bHideMicrosoft Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                Else
                    sDefCLSID_all = vbNullString
                    If InStr(1, sFile, "\Microsoft Shared\", 1) <> 0 Then
                        If bHideMicrosoft Then
                            If IsMicrosoftFile(sFile) Then bSafe = True
                        End If
                    End If
                End If
                
                'Repeat for subkey
                
                If Not bSafe Then
                    'Filters key, possibly, can also contain several subkeys with similar contents
                    bBySubKey = False
                    If sCLSID = STR_NO_CLSID Then
                        If Reg.KeyHasSubKeys(HE.Hive, HE.Key & "\" & sName, HE.Redirected) Then
                            bBySubKey = True
                            For j = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key & "\" & sName, aSubKey, HE.Redirected)
                                sCLSID = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\" & aSubKey(j), "CLSID", HE.Redirected)
                                
                                If Len(sCLSID) = 0 Then
                                    sCLSID = STR_NO_CLSID
                                    sFile = STR_NO_FILE
                                Else
                                    Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, bShared)
                                End If
                                
                                If oDict.dSafeFilters.Exists(sName) Then
                                    sDefCLSID_all = oDict.dSafeFilters(sName)
                                    If bHideMicrosoft Then
                                        If IsMicrosoftFile(sFile) Then bSafe = True
                                    End If
                                Else
                                    sDefCLSID_all = vbNullString
                                    If InStr(1, sFile, "\Microsoft Shared\", 1) <> 0 Then
                                        If bHideMicrosoft Then
                                            If IsMicrosoftFile(sFile) Then bSafe = True
                                        End If
                                    End If
                                End If
                                
                                If Not bSafe Then
                                    sHit = BitPrefix("O18", HE) & " - " & HE.KeyAndHivePhysical & "\" & sName & _
                                        ": [CLSID] = " & sCLSID & " - " & sFile
                                    sFixKey = HE.Key & "\" & sName & "\" & aSubKey(j)
                                    
                                    GoSub labelFix:
                                End If
                            Next
                        End If
                    End If
                    
                    If Not bBySubKey Then
                        'HKCU often has empty filter keys, so skip them
                        
                        'If Not (sCLSID = "(no CLSID)" And (HE.Hive = HKCU Or HE.Hive = HKU)) Then
                    
                            sHit = BitPrefix("O18", HE) & " - " & HE.KeyAndHivePhysical & "\" & sName & _
                                ": [CLSID] = " & sCLSID & " - " & sFile & FormatSign(result.SignResult)
                            sFixKey = HE.Key & "\" & sName
                            
                            GoSub labelFix:
                        'End If
                    End If
                End If
                
                sName = String$(MAX_KEYNAME, 0&)
                lpcName = Len(sName)
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    AppendErrorLogCustom "CheckO18Item - End"
    Exit Sub

labelFix:
    SignVerifyJack sFile, result.SignResult
    sHit = sHit & FormatSign(result.SignResult)
    
    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
    
    If Not IsOnIgnoreList(sHit) Then
        With result
            .Section = "O18"
            .HitLineW = sHit
            
            sDefCLSID = vbNullString
            'find suitable CLSID by database and check is it legit
            For Each vDefCLSID In SplitSafe(sDefCLSID_all, "|")
                If Len(vDefCLSID) <> 0 Then
                    Call GetFileByCLSID(CStr(vDefCLSID), sDefFile, , HE.Redirected, bShared)
                    If IsMicrosoftFile(sDefFile) Then
                        sDefCLSID = CStr(vDefCLSID)
                        Exit For
                    End If
                End If
            Next
            If Len(sDefCLSID) = 0 Then
                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, sFixKey, , , HE.Redirected
            Else
                AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, sFixKey, "CLSID", sDefCLSID, HE.Redirected
            End If

            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
            
            .CureType = REGISTRY_BASED Or FILE_BASED
        End With
        AddToScanResults result
    End If
    
    Return

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO18Item"
    If hKey <> 0 Then RegCloseKey hKey
End Sub

Public Sub ScanPrinterPorts() '// Thanks to Alex Ionescu
    On Error GoTo ErrorHandler:
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports

    Dim i&, aPorts$(), sFile$, sHit$, result As SCAN_RESULT
    
    For i = 1 To Reg.EnumValuesToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports", aPorts, False)
        
        sFile = EnvironW(aPorts(i))
        
        If InStr(sFile, ":\") <> 0 Or InStr(sFile, ":/") <> 0 Then 'look as file
            
            sHit = "O18 - Printer Port: " & sFile
            
            If Not IsOnIgnoreList(sHit) Then
                
                With result
                    .Section = "O18"
                    .HitLineW = sHit
                    AddCustomToFix .Custom, CUSTOM_ACTION_SPECIFIC, aPorts(i)
                    AddRegToFix .Reg, REMOVE_VALUE, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports", aPorts(i), , False
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ScanPrinterPorts"
End Sub

Public Function FileMissing(sFile$) As Boolean
    If Len(sFile) = 0 Then FileMissing = True: Exit Function
    If Right$(sFile, 1) = ")" Then
        If sFile = STR_NO_FILE Then FileMissing = True: Exit Function
        If StrEndWith(sFile, STR_FILE_MISSING) Then FileMissing = True: Exit Function
    End If
End Function

Public Sub RemoveFileMissingStr(sFile$)
    If Len(sFile) = 0 Then Exit Sub
    Dim i As Long
    i = InStr(sFile, " " & STR_FILE_MISSING)
    If i <> 0 Then sFile = Left$(sFile, Len(sFile) - Len(STR_FILE_MISSING) - 1)
End Sub

Public Sub FixO18Item(sItem$, result As SCAN_RESULT)
    'O18 - Protocol: cn
    'O18 - Filter: text/blah - {0} - c:\file.dll
    'O18 - Printer Port: c:\file.exe
    On Error GoTo ErrorHandler:
    
    If InStr(1, result.HitLineW, "Printer Port:", 1) <> 0 Then
        
        Dim sPort As String
        sPort = result.Custom(0).Name
        
        'get-printer / remove-printer are Win 8+ only?
        Call Proc.RunPowershell("$printer = get-printer * | where {$_.portname -eq '" & sPort & "'}; remove-printer -inputobject $printer", True)
        
    End If
    
    FixIt result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO18Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO19Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO19Item - Begin"
    
    'Software\Microsoft\Internet Explorer\Styles,Use My Stylesheet
    'Software\Microsoft\Internet Explorer\Styles,User Stylesheet
    
    Dim lUseMySS&, sUserSS$, sHit$, result As SCAN_RESULT
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_HKCU Or HE_HIVE_HKU
    HE.AddKey "Software\Microsoft\Internet Explorer\Styles"
    
    Do While HE.MoveNext
        lUseMySS = Reg.GetDword(HE.Hive, HE.Key, "Use My Stylesheet", HE.Redirected)
        
        If lUseMySS <> 0 Then
        
          sUserSS = Reg.GetString(HE.Hive, HE.Key, "User Stylesheet", HE.Redirected)
          sUserSS = FormatFileMissing(sUserSS)
        
          If Len(sUserSS) <> 0 Then
        
            'O19 - User stylesheet (HKCU,HKLM):
            'O19-32 - User stylesheet (HKCU,HKLM):
            sHit = BitPrefix("O19", HE) & " - " & HE.HiveNameAndSID & "\..\Internet Explorer\Styles: " & _
                "[User Stylesheet] = " & sUserSS
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O19"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, "Use My Stylesheet", 0&, HE.Redirected
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "User Stylesheet", , HE.Redirected
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
          End If
        End If
    Loop
    
    AppendErrorLogCustom "CheckO19Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO19Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO19Item(sItem$, result As SCAN_RESULT)
    'O19 - User stylesheet: c:\file.css (file missing)
    
    FixIt result
End Sub

Public Sub CheckO20Item()
    'AppInit_DLLs
    'https://support.microsoft.com/ru-ru/kb/197571
    'https://msdn.microsoft.com/en-us/library/windows/desktop/dd744762(v=vs.85).aspx
    
    'According to MSDN:
    ' - modules are delimited by spaces or commas
    ' - long file names are not permitted
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO20Item - Begin"
    
    'appinit_dlls + winlogon notify
    Dim sAppInit$, sFile$, sHit$, UseWow, Wow6432Redir As Boolean, result As SCAN_RESULT
    Dim bEnabled As Boolean, bRequireCodeSigned As Boolean, aFile() As String, bUnsigned As Boolean, i As Long
    Dim sTrimChar As String, sOrigLine As String
    
    For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If bIsWin32 And Wow6432Redir Then Exit For
    
        sAppInit = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
        
        If OSver.MajorMinor <= 5.2 Then 'XP/2003-
            bEnabled = True
        Else
            bEnabled = (1 = Reg.GetDword(HKEY_LOCAL_MACHINE, sAppInit, "LoadAppInit_DLLs", Wow6432Redir))
            bRequireCodeSigned = (1 = Reg.GetDword(HKEY_LOCAL_MACHINE, sAppInit, "RequireSignedAppInit_DLLs", Wow6432Redir))
        End If
        
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sAppInit, "AppInit_DLLs", Wow6432Redir)
        If Len(sFile) <> 0 Then
            
            aFile = SplitByMultiDelims(sFile, True, sTrimChar, ",", " ")
            ArrayRemoveEmptyItems aFile
            
            For i = 0 To UBound(aFile)
            
                sFile = aFile(i)
                sOrigLine = sFile
                
                If (InStr(1, "*" & sSafeAppInit & "*", "*" & sFile & "*", vbTextCompare) = 0) Or bIgnoreAllWhitelists Then
                    'item is not on whitelist
                    'O20 - AppInit_DLLs
                    'O20-32 - AppInit_DLLs
                    
                    sFile = FormatFileMissing(sFile)
                    
                    If bRequireCodeSigned Then
                        bUnsigned = False
                        If FileExists(sFile) Then
                            If Not IsLegitFileEDS(sFile) Then bUnsigned = True
                        End If
                    End If
                    
                    SignVerifyJack sFile, result.SignResult
                    
                    sHit = IIf(bIsWin32, "O20", IIf(Wow6432Redir, "O20-32", "O20")) & " - HKLM\..\Windows: [AppInit_DLLs] = " & sFile & _
                      FormatSign(result.SignResult) & _
                      IIf(bRequireCodeSigned And bUnsigned, " (disabled because not code signed)", vbNullString) & _
                      IIf(Not bEnabled, " (disabled by registry)", vbNullString) & IIf(OSver.SecureBoot, " (disabled by SecureBoot)", vbNullString)
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O20"
                            .HitLineW = sHit
                            
                            'to disable loading AppInit_DLLs
                            'AddRegToFix .Reg, RESTORE_VALUE, 0, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Windows", "LoadAppInit_DLLs", 0, CLng(Wow6432Redir), REG_RESTORE_DWORD
                            
                            AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE, _
                                HKLM, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", , CLng(Wow6432Redir), REG_RESTORE_SZ, _
                                sOrigLine, vbNullString, sTrimChar
                            
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        End If
        
        Dim sSubkeys$(), sWinLogon$
        sWinLogon = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify"
        sSubkeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sWinLogon, Wow6432Redir), "|")
        If UBound(sSubkeys) <> -1 Then
            For i = 0 To UBound(sSubkeys)
                sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sWinLogon & "\" & sSubkeys(i), "DllName", Wow6432Redir)
                sFile = FormatFileMissing(sFile)
                SignVerifyJack sFile, result.SignResult
                
                If (Not result.SignResult.isMicrosoftSign) Or (Not bHideMicrosoft) Or bIgnoreAllWhitelists Then
                    'O20 - Winlogon Notify:
                    sHit = IIf(bIsWin32, "O20", IIf(Wow6432Redir, "O20-32", "O20")) & " - HKLM\..\Winlogon\Notify\" & sSubkeys(i) & _
                        ": [DllName] = " & sFile & FormatSign(result.SignResult)
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O20"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify\" & sSubkeys(i), , , CLng(Wow6432Redir)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next i
        End If
    Next

    AppendErrorLogCustom "CheckO20Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO20Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO20Item(sItem$, result As SCAN_RESULT)
    'O20 - AppInit_DLLs: file.dll
    'O20 - Winlogon Notify: bladibla - c:\file.dll
    
    FixIt result
End Sub

Public Sub CheckO21Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO21Item - Begin"
    
    'Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad
    'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers
    'Software\Microsoft\Windows\CurrentVersion\explorer\ShellExecuteHooks
    
    '//TODO
    'Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved
    'HKCR\Folder\shellex\ColumnHandlers
    'HKCR\AllFilesystemObjects\shellex\ContextMenuHandlers
    'ShellServiceObjects
    '
    
    Dim sSSODL$, sHit$, sFile$
    Dim hKey&, i&, sName$, lNameLen&, sCLSID$, lDataLen&, sValueName$
    Dim result As SCAN_RESULT, bSafe As Boolean, bInList As Boolean
    
    sSSODL = "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
    
    'BE AWARE: SHELL32.dll - sometimes this file is patched
    '(e.g. seen after "Windown XP Update pack by Simplix" together with his certificate installed to trusted root storage)
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_HKLM
    HE.AddKey sSSODL

    Do While HE.MoveNext
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then

            Do
                lNameLen = MAX_VALUENAME
                sValueName = String$(lNameLen, 0&)
                lDataLen = MAX_VALUENAME
                sCLSID = String$(lDataLen, 0&)
                
                If RegEnumValueW(hKey, i, StrPtr(sValueName), lNameLen, 0&, REG_SZ, StrPtr(sCLSID), lDataLen) <> 0 Then Exit Do

                sValueName = Left$(sValueName, lNameLen)
                sCLSID = TrimNull(sCLSID)
                
                Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, HE.SharedKey)

                sFile = FormatFileMissing(sFile)
                
                bSafe = False
                If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                    
                    bInList = InArray(sCLSID, aSafeSSODL, , , vbTextCompare)

                    If bInList Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                    If Not bSafe Then
                        If WhiteListed(sFile, "GROOVEEX.DLL", True) Then bSafe = True
                    End If
                End If

                If Not bSafe Then
                    Call GetTitleByCLSID(sCLSID, sName, HE.Redirected, HE.SharedKey)
                    If sName = STR_NO_NAME Then sName = sValueName
                    
                    SignVerifyJack sFile, result.SignResult
                    
                    sHit = BitPrefix("O21", HE) & " - HKLM\..\ShellServiceObjectDelayLoad: " & _
                        IIf(sName <> sValueName, sName & " ", vbNullString) & "[" & sValueName & "] " & " = " & sCLSID & " - " & _
                        sFile & FormatSign(result.SignResult)

                    'some shit leftover by Microsoft ^)
                    If bHideMicrosoft And (sName = "WebCheck" And sCLSID = "{E6FB5E20-DE35-11CF-9C87-00AA005127ED}" And sFile = STR_NO_FILE) Then bSafe = True
                
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                    If Not IsOnIgnoreList(sHit) And Not bSafe Then
                        With result
                            .Section = "O21"
                            .HitLineW = sHit
                            If Len(sCLSID) <> 0 Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sName, , HE.Redirected
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop

    'ShellIconOverlayIdentifiers
    Dim aSubKey() As String
    Dim sSIOI As String
    Dim sPrevFile As String
    
    sSIOI = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers"
    
    HE.Init HE_HIVE_HKLM
    HE.AddKey sSIOI

    Do While HE.MoveNext
        
        If Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey, HE.Redirected) > 0 Then
        
            For i = 1 To UBound(aSubKey)

                sName = aSubKey(i)
                sCLSID = Reg.GetString(HE.Hive, HE.Key & "\" & aSubKey(i), vbNullString, HE.Redirected)
                
                Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, HE.SharedKey)

                sFile = FormatFileMissing(sFile)
                
                bSafe = False
                If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                    
                    bInList = InArray(sFile, aSafeSIOI, , , vbTextCompare)
                    
                    If StrComp(GetFileName(sFile, True), "GROOVEEX.DLL", 1) = 0 Then bInList = True
                    If StrComp(GetFileName(sFile, True), "FileSyncShell.dll", 1) = 0 Then bInList = True
                    If StrComp(GetFileName(sFile, True), "FileSyncShell64.dll", 1) = 0 Then bInList = True
                    
                    If bInList Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                End If

                If Not bSafe Then
                    Call GetTitleByCLSID(sCLSID, sName, HE.Redirected, HE.SharedKey)
                    
                    SignVerifyJack sFile, result.SignResult
                    
                    If sFile = STR_NO_FILE Then
                    
                        sHit = BitPrefix("O21", HE) & " - HKLM\..\ShellIconOverlayIdentifiers\" & _
                            aSubKey(i) & ": " & sName & " - " & sCLSID & " - " & sFile
                    Else
                    
                        sHit = BitPrefix("O21", HE) & " - HKLM\..\ShellIconOverlayIdentifiers\" & _
                            " - " & sFile & FormatSign(result.SignResult)
                    End If

                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)

                    If Not IsOnIgnoreList(sHit) And Not bSafe Then
                        
                        If result.CureType <> 0 Then
                            If StrComp(sFile, sPrevFile, 1) <> 0 Then
                                AddToScanResults result, DoNotDuplicate:=True
                            End If
                        End If

                        With result
                            .Section = "O21"
                            .HitLineW = sHit
                            
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile 'should go first!
                            
                            If Len(sCLSID) <> 0 Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        
                        If sFile = STR_NO_FILE Then
                            AddToScanResults result
                        Else
                            sPrevFile = sFile
                        End If
                    End If
                End If
            
            Next
        End If
    Loop
    If result.CureType <> 0 Then AddToScanResults result

    'ShellExecuteHooks
    'See: http://blog.zemana.com/2016/06/youndoocom-using-shellexecutehooks-to.html
    
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => EnableShellExecuteHooks
    
    Dim bDisabled As Boolean
    Dim aValue() As String
    
    HE.Init HE_HIVE_HKLM
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks"
    
    Do While HE.MoveNext
        
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue, HE.Redirected)
            sCLSID = aValue(i)
            
            Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, HE.SharedKey)
            
            sFile = FormatFileMissing(sFile)
            
            bSafe = False
            If bHideMicrosoft And Not bIgnoreAllWhitelists Then
            
                bInList = InArray(sFile, aSafeSEH, , , vbTextCompare)
                If StrComp(GetFileName(sFile, True), "GROOVEEX.DLL", 1) = 0 Then bInList = True
                
                If bInList Then
                    If IsMicrosoftFile(sFile) Then bSafe = True
                End If
            End If

            If Not bSafe Then
                Call GetTitleByCLSID(sCLSID, sName, HE.Redirected, HE.SharedKey)
                SignVerifyJack sFile, result.SignResult
                
                If OSver.MajorMinor >= 6 Then 'XP/2003 has no policy
                    bDisabled = Not (1 = Reg.GetDword(HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "EnableShellExecuteHooks"))
                End If
                
                sHit = BitPrefix("O21", HE) & " - HKLM\..\ShellExecuteHooks: [" & sCLSID & "] - " & sName & " - " & _
                    sFile & FormatSign(result.SignResult) & IIf(bDisabled, " (disabled)", vbNullString)
                
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O21"
                        .HitLineW = sHit
                        If Len(sCLSID) <> 0 Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                        AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    'ColumnHandlers
    'Note: Not available in Vista+
    'HKEY_CLASSES_ROOT\Folder\ShellEx\ColumnHandlers
    
    'see also: https://www.nirsoft.net/utils/shexview.html
    
'    HE.Init HE_HIVE_ALL
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\ColumnHandlers"
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\ContextMenuHandlers"
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\DragDropHandlers"
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\PropertySheetHandlers"
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\CopyHookHandlers"
    
    '����� ��������, ��� ��� ����������� � ����. ����������� ����� ���� ���������������� ��� ��� �����, ��� � ��� ��������� ���������� �����.
    '����������� �� ������? ����� ����� ��������� ���������� CLSID.
    
    'HKEY_CLASSES_ROOT\*\shellex
    'HKEY_CLASSES_ROOT\Folder\shellex - virtual combination of actual file folders, shell folders, drives, other special folders
    'HKEY_CLASSES_ROOT\Directory\shellex
    'and so ...
    
    'Explorer Shell extensions
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved
    'see:
    'https://msdn.microsoft.com/en-us/library/ms812054.aspx
    'https://forum.sysinternals.com/shell-extensions-approved_topic11891.html
    'So, this key can be needed only for heuristic cleaning of extension
    
    
    AppendErrorLogCustom "CheckO21Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO21Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO21Item(sItem$, result As SCAN_RESULT)
    'O21 - SSODL: webcheck - {000....000} - c:\file.dll (file missing)
    'actions to take:
    '* kill file
    '* kill regkey - ShellIconOverlayIdentifiers
    '* kill regparam - SSODL
    '* kill clsid regkey
    
    ShutdownExplorer
    FixIt result
    
End Sub

Public Sub CheckO22Item()
    'ScheduledTask
    'XP    - HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler
    'Vista - HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO22Item - Begin"
    
    EnumTasks
    EnumJobs
    'EnumBITS 'splitted in 2 stages for better optimiz.
    
    AppendErrorLogCustom "CheckO22Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO22Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO22Item(sItem$, result As SCAN_RESULT)
    'O22 - ScheduledTask: blah - {000...000} - file.dll
    
    FixIt result
End Sub

Public Sub CheckO23Item()
    'https://www.bleepingcomputer.com/tutorials/how-malware-hides-as-a-service/
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO23Item - Begin"
    
    'enum NT services
    Dim sServices$(), i&, j&, sName$, sDisplayName$, tmp$, result As SCAN_RESULT
    Dim lStart&, lType&, sFile$, sHit$, sBuf$, IsCompositeCmd As Boolean
    Dim bHideDisabled As Boolean, sServiceDll As String, sServiceDll_2 As String, bDllMissing As Boolean
    Dim ServState As SERVICE_STATE
    Dim argc As Long
    Dim argv() As String
    Dim isSafeMSCmdLine As Boolean
    Dim FoundFile As String
    Dim IsMSCert As Boolean
    Dim sImagePath As String
    Dim sArgument As String
    Dim pos As Long
    Dim bSuspicious As Boolean
    Dim sGroup As String
    Dim sServiceGroup As String
    Dim sFailureCommand As String
    
    Dim dLegitService As clsTrickHashTable
    Set dLegitService = New clsTrickHashTable
    dLegitService.CompareMode = vbTextCompare
    
    Dim dLegitGroups As clsTrickHashTable
    Set dLegitGroups = New clsTrickHashTable
    dLegitGroups.CompareMode = vbTextCompare
    
    If Not bIsWinNT Then Exit Sub
    
    If Not bIgnoreAllWhitelists Then
        bHideDisabled = True
    End If
    
    sServices = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    
    If UBound(sServices) = -1 Then Exit Sub
    
    For i = 0 To UBound(sServices)
        
        sName = sServices(i)
        Dbg sName
        
        'sFailureCommand = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "FailureCommand")
        'If Len(sFailureCommand) <> 0 Then
        '    Debug.Print sName; " -> " & sFailureCommand
        'End If
        
        lType = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Type")
        
        If (lType < 16 And lType <> -1) Then 'Driver
            If Not bAdditional Then 'if 'O23 - Driver' check is skipped
                If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 4&
                sGroup = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Group")
                If Len(sGroup) <> 0 Then
                    If Not dLegitGroups.Exists(sGroup) Then dLegitGroups.Add sGroup, 4&
                End If
            End If
            GoTo Continue
        End If
        
        lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start")
        
        If (lStart = 4 And bHideDisabled) Then
            If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 4&
            GoTo Continue
        End If
        
        UpdateProgressBar "O23", sName
        
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ImagePath")
        sServiceDll = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName & "\Parameters", "ServiceDll")
        
        bDllMissing = False
        
        result.SignResult.isMicrosoftSign = False
        
        'Checking Service Dll
        If Len(sServiceDll) <> 0 Then
            sServiceDll = EnvironW(UnQuote(sServiceDll))
            
            tmp = FindOnPath(sServiceDll)
            
            If Len(tmp) = 0 Then
                sServiceDll = sServiceDll & " " & STR_FILE_MISSING
                bDllMissing = True
            Else
                sServiceDll = tmp
            End If
        End If
        
        If bDllMissing Then
            
            sServiceDll_2 = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ServiceDll")
            
            If Len(sServiceDll_2) <> 0 Then
                
                sServiceDll_2 = EnvironW(UnQuote(sServiceDll_2))
                
                tmp = FindOnPath(sServiceDll_2)
                
                If Len(tmp) <> 0 Then sServiceDll = tmp: bDllMissing = False
            End If
        End If
        
        'cleanup filename
        sImagePath = sFile
        sFile = CleanServiceFileName(sFile, sName)
        
        '//// TODO: Check this !!!
        
        'https://technet.microsoft.com/en-us/library/cc959922.aspx
        
        'Performance key:
        'https://learn.microsoft.com/en-us/windows-hardware/drivers/install/hklm-system-currentcontrolset-services-registry-tree
        
        'Start
        '0 - Boot
        '1 - System
        '2 - Automatic
        '3 - Manual
        '4 - Disabled
        
        'Type
        '1 - Kernel device driver
        '2 - File System driver
        '4 - A set of arguments for an adapter
        '8 - File System driver service
        '16 - A Win32 program that runs in a process by itself. This type of Win32 service can be started by the service controller.
        '32 - A Win32 service that can share a process with other Win32 services
        '272 - A Win32 program that runs in a process by itself (like Type16) and that can interact with users.
        '288 - A Win32 program that shares a process and that can interact with users.
        '
        'According to Helge Klein (per-user services), see: https://helgeklein.com/blog/per-user-services-in-windows-info-and-configuration/
        '
        '80 (0x50) - (per-user service) individual Process template (SERVICE_USER_OWN_PROCESS)
        '96 (0x60) - (per-user service) shared process template (SERVICE_USER_SHARE_PROCESS)
        '208 (0xd0) - (per-user service) individual process (SERVICE_USER_OWN_PROCESS + SERVICE_USERSERVICE_INSTANCE + 0x10)
        '224 (0xe0) - (per-user service) shared process (SERVICE_USER_SHARE_PROCESS + SERVICE_USERSERVICE_INSTANCE + 0x10)
        
        ServState = GetServiceRunState(sName)
        
        sArgument = vbNullString
        bSuspicious = False
        
        '// TODO:
        'or lType = -1 (e.g. RDPNP)
        
        If (lType >= 16) Then
          If Not (lStart = 4 And bHideDisabled) Then
            
            IsCompositeCmd = False
            isSafeMSCmdLine = False
            
            '������� ��������� ������ - ��� ������ ��� ���� �� ���������� �� �����
            If Not FileExists(sImagePath) And Len(sImagePath) <> 0 Then
            
                ' ������ ���� ��������� �������� ��������� ������ � �������� ��������� ��� ������� ����� �� ���� �������
                ' ���� ����� ���� �� ������� �� �������� ��������, ������ ��������� ������������
            
                ParseCommandLine sImagePath, argc, argv
                
                If argc > 1 Then
                    pos = InStr(sImagePath, argv(1))
                    If pos <> 0 Then
                        sArgument = mid$(sImagePath, pos + Len(argv(1)))
                        If Left$(sArgument, 1) = """" Then sArgument = mid$(sArgument, 2)
                        sArgument = LTrim$(sArgument)
                    End If
                End If
                
                If Len(sArgument) <> 0 Then
                    If Len(sArgument) > 50 Then
                        bSuspicious = True
                    Else
                        sBuf = EnvironW(sArgument)
                        If Not bSuspicious Then If InStr(1, sBuf, "http:", 1) <> 0 Then bSuspicious = True
                        If Not bSuspicious Then If InStr(1, sBuf, "https:", 1) <> 0 Then bSuspicious = True
                    End If
                End If
                
                '// TODO: �������� � FindOnPath �����, � ������� ��������� �������� ����������� ������� ����
                
                '���� ���� � ������� ���������� ������, ��������: C:\WINDOWS\system32\svchost -k rpcss.exe
                
                If argc > 2 Then        ' 1 -> app exe self, 2 -> actual cmd, 3 -> arg
                
                  If Not FileExists(argv(1)) Then   ' ���� ����������� ���� �� ���������� -> ���� ���
                    FoundFile = FindOnPath(argv(1))
                    argv(1) = FoundFile
                  Else
                    FoundFile = argv(1)
                  End If
                
                  ' ���� ����������� ���� ���������� (�����, ��� ������ ��������� ��������� ���������)
                  If 0 <> Len(FoundFile) And Not StrBeginWith(sImagePath, sWinSysDir & "\svchost.exe -k") Then
                  
                    '���� � ���, ��� ������ ��������� ��������� ��������� ������, � ������� ��� ������� ������ (����������� ����) ����������
                    IsCompositeCmd = True
                
                    isSafeMSCmdLine = True
                 
                    For j = 1 To UBound(argv) ' argv[1] -> ����������� ���� � �������
                    
                        ' ��������� ��� ��������� ����������� ������� �� ��������� ��������� ������, ���� �� ��� ������ �� ��������� ����� Path
                        
                        FoundFile = FindOnPath(argv(j))
                        
                        If 0 <> Len(FoundFile) Then
                        
                            If IsWinServiceFileName(FoundFile) Then
                                SignVerifyJack FoundFile, result.SignResult
                                IsMSCert = result.SignResult.isMicrosoftSign
                            Else
                                IsMSCert = False
                            End If
                            
                            If Not IsMSCert Then isSafeMSCmdLine = False: Exit For
                        End If
                    Next
                  End If
                End If
            
            End If
            
            If 0 = Len(sFile) Then
                sFile = STR_NO_FILE
            Else
                If (Not FileExists(sFile)) And (Not IsCompositeCmd) Then
                    sFile = sFile & " " & STR_FILE_MISSING
                Else
'                    If IsCompositeCmd Then
'                        FoundFile = argv(1)
'                    Else
'                        FoundFile = sFile
'                    End If
                    
                    'sCompany = GetFilePropCompany(FoundFile)
                    'If Len(sCompany) = 0 Then sCompany = "Unknown owner"
                    
                End If
            End If
            
            result.SignResult.isMicrosoftSign = False
            
            If IsCompositeCmd Then
                If Not isSafeMSCmdLine Then bSuspicious = True
            Else
                If sFile <> STR_NO_FILE Then    '�����, ����� �������� ��� ��������� �����
                    If IsWinServiceFileName(sFile, sArgument) Then
                        SignVerifyJack sFile, result.SignResult
                    Else
                        'WipeSignResult SignResult
                    End If
                End If
            End If
            
            'override by checkind EDS of service dll if original file is Microsoft (usually, svchost)
            If bDllMissing Then
                result.SignResult.isMicrosoftSign = False
            Else
                If Len(sServiceDll) <> 0 Then
                    If IsWinServiceFileName(sServiceDll) Then
                        SignVerifyJack sServiceDll, result.SignResult
                    Else
                        result.SignResult.isMicrosoftSign = False
                    End If
                End If
            End If
            
            If True Then
                '��������� � ������ ���������� ����� ��� ����������� ������������� ��� �������� ������������
                If Not (bSuspicious Or bDllMissing Or Not (result.SignResult.isMicrosoftSign)) Then
                    If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 0&
                End If
                
                ' ���� �������� ���������� ������� ������� ����������� ���������� + � ������, ��� ���� �������� �� ����, �� ��������� ������ �� ����
                
                If bSuspicious Or bDllMissing Or Not (result.SignResult.isMicrosoftSign And bHideMicrosoft) Then
                    
                    sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DisplayName")

                    If Len(sDisplayName) = 0 Then
                        sDisplayName = sName
                    Else
                        If Left$(sDisplayName, 1) = "@" Then                    'extract string resource from file

                            sBuf = GetStringFromBinary(, , sDisplayName)

                            If 0 <> Len(sBuf) Then sDisplayName = sBuf
                        End If
                    End If
                    
                    sHit = "O23 - Service " & IIf(ServState <> SERVICE_STOPPED, "R", "S") & lStart & _
                        ": " & IIf(sDisplayName = sName, sName, sDisplayName & " - (" & sName & ")") & " - " & ConcatFileArg(sFile, sArgument)
                    
                    If Len(sServiceDll) <> 0 Then
                        sHit = sHit & "; ""ServiceDll"" = " & sServiceDll
                    End If
                    
                    SignVerifyJack IIf(Len(sServiceDll) = 0, sFile, sServiceDll), result.SignResult
                    sHit = sHit & FormatSign(result.SignResult)
                    
                    If IsSafemodeService(sName) Then sHit = sHit & " (+safe mode)"
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(IIf(Len(sServiceDll) = 0, sFile, sServiceDll))
                    
                    If Not IsOnIgnoreList(sHit) Then
                        
                        With result
                            .Section = "O23"
                            .HitLineW = sHit
                            .Name = sName 'used in "Disable" stuff
                            .State = IIf(lStart <> 4, ITEM_STATE_ENABLED, ITEM_STATE_DISABLED)
                            
                            AddServiceToFix .Service, DELETE_SERVICE Or USE_FEATURE_DISABLE, sName, , , , ServState, True
                        
                            If Len(sServiceDll) = 0 Then
                                AddFileToFix .File, BACKUP_FILE, sFile, sArgument
                            Else
                                AddFileToFix .File, BACKUP_FILE, sServiceDll
                            End If
                            
                            AddRegToFix .Reg, BACKUP_KEY, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName
                            .Reboot = True
                            .CureType = SERVICE_BASED Or FILE_BASED Or REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
          End If
        End If
Continue:
    Next i
    
    'checking drivers
    
    UpdateProgressBar "O23-D"
    
    If bAdditional Then
        CheckO23Item_Drivers sServices, dLegitService
    End If
    
    'Check dependency *(should go after 'O23 - Drivers' scan !!!)
    
    'Temporarily added to "Additional scan", until I figure out all cases with damaged EDS subsystem
    If bAdditional Then
        CheckO23Item_Dependency sServices, dLegitService, dLegitGroups
    End If
    
    Set dLegitService = Nothing
    
    AppendErrorLogCustom "CheckO23Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO23Item", "Service=", sDisplayName
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO23Item_Dependency(sServices() As String, dLegitService As clsTrickHashTable, dLegitGroups As clsTrickHashTable)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO23Item_Dependency - Begin"
    
    Dim i&, k&
    Dim aDepend()       As String
    Dim vName           As Variant
    Dim sName           As String
    Dim sGroup          As String
    Dim sHit            As String
    Dim bMissing        As Boolean
    Dim bSafe           As Boolean
    Dim result          As SCAN_RESULT
    Dim aSubKey()       As String
    
    '"DependOnService" parameter
    
    UpdateProgressBar "O23", "Dependency"
    
    'Appending list of legit services with services that have no "ImagePath" (XP/2003- only)
    If OSver.MajorMinor <= 5.2 Then
        For i = 1 To Reg.EnumSubKeysToArray(HKLM, "System\CurrentControlSet\Services", aSubKey)
            If Not Reg.ValueExists(HKLM, "System\CurrentControlSet\Services\" & aSubKey(i), "ImagePath") Then
                If Not dLegitService.Exists(aSubKey(i)) Then
                    dLegitService.Add aSubKey(i), 0&
                End If
            End If
        Next
    End If
    
    For Each vName In dLegitService.Keys
        
        If dLegitService(vName) <> 4 Then 'check only real legit services (4 - is unknown state)
        
            sName = vName
            
            aDepend = Reg.GetMultiSZ(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DependOnService")
    
            If AryItems(aDepend) Then
                For k = 0 To UBound(aDepend)
                    
                    bSafe = dLegitService.Exists(aDepend(k))
                    
                    If Not bSafe Then
                        
                        bMissing = Not Reg.KeyExists(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & aDepend(k))
                        
                        'Win10 "feature"
                        'See: https://answers.microsoft.com/en-us/windows/forum/windows_10-other_settings/dependonservice-refers-to-a-non-existent-service/f63265d1-70ee-4561-b473-e54085cdeaf2
                        If bMissing And OSver.IsWindows10OrGreater Then
                            If StrEndWith(aDepend(k), "x") Then 'UcmCx, UcmUcsiCx, GPIOClx
                                bSafe = dLegitService.Exists(aDepend(k) & "0101")
                            End If
                        End If
                        
                        If Not bSafe Then
                        
                            sHit = "O23 - Dependency: Microsoft Service '" & sName & "' depends on unknown service: '" & aDepend(k) & "'" & _
                                IIf(bMissing, " (service missing)", vbNullString)
                            
                            If Not IsOnIgnoreList(sHit) Then
                                
                                With result
                                    .Section = "O23"
                                    .HitLineW = sHit
                                    
                                    AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                        HKLM, "System\CurrentControlSet\Services\" & sName, "DependOnService", , , REG_RESTORE_MULTI_SZ, _
                                        aDepend(k), vbNullString, vbNullChar
                                    
                                    .Reboot = True
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults result
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    'Check dependency on groups
    
    '"DependOnGroup" parameter
    
    'Note: Serice group can be created by specifying "Group" registry parameter for some service.
    'There is no separate list.
    'Service groups loading order is stored in:
    ' - HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\GroupOrderList
    ' - HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\ServiceGroupOrder [List]
    ' - "Tag" reg. parameter - is an order of service loading in particular service group
    
    'Groups can be:
    '1. Legit (all services of this group belong to Microsoft)
    '2. Semi-legit (group contains both Microsoft and non-Microsoft services)
    '3. Non-legit (group contains non-Microsoft services only)
    
    'Firstly, we'll add legit and semi-legit groups to dLegitGroups dictionary.
    'And compare "Group" reg. parameter of each service with this dictionary.
    'Next, we'll list all semi-legit group to HJT log because wtf, that is wrong.
    
    'Gather groups
    
    For i = 0 To UBound(sServices)
        
        sName = sServices(i)
        sGroup = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Group")
        
        'idx description
        '1 - legit
        '2 - semi-legit
        '3 - non-legit
        '4 - state is unknown, because check is not performed. It can be due:
        ' > "Additional scan" is not performed (so, no Driver services check)
        ' > Service state is "disabled" (so such services is also not checked), unless "Ignore ALL Whitelist" is marked
        
        If Len(sGroup) <> 0 Then
            If dLegitGroups.Exists(sGroup) Then
            
                k = dLegitGroups(sGroup)
                
                If dLegitService.Exists(sName) Then
                    'Current is Legit: Make 3 => 2
                    If k = 3 Then dLegitGroups(sGroup) = 2
                Else
                    'Current is Non-legit: Make 1 => 2
                    If k = 1 Then dLegitGroups(sGroup) = 2
                End If
            Else
                If dLegitService.Exists(sName) Then
                    dLegitGroups.Add sName, 1 'make legit
                Else
                    dLegitGroups.Add sName, 3 'make non-legit
                End If
            End If
        End If
    Next
    'We should have only legit and semi-legit groups
    'So remove non-legit:
    For Each vName In dLegitGroups.Keys
        If dLegitGroups(vName) = 3 Then dLegitGroups.Remove vName
    Next
    
    'Ok, now we'll check all entries in "DependOnGroup" parameter of legit. services against "dLegitGroups" dictionary
    
    For Each vName In dLegitService.Keys
    
        If dLegitService(vName) <> 4 Then 'check only real legit services (4 - is unknown state)
        
            sName = vName
            
            aDepend = Reg.GetMultiSZ(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DependOnGroup")
            
            If AryItems(aDepend) Then
                For k = 0 To UBound(aDepend)
                    
                    If Not dLegitGroups.Exists(aDepend(k)) Then
                        
                        sHit = "O23 - Dependency: Microsoft Service '" & sName & "' depends on mixed group: '" & aDepend(k) & "'"
                        
                        If Not IsOnIgnoreList(sHit) Then
                            
                            With result
                                .Section = "O23"
                                .HitLineW = sHit
                                
                                AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                    HKLM, "System\CurrentControlSet\Services\" & sName, "DependOnGroup", , , REG_RESTORE_MULTI_SZ, _
                                    aDepend(k), vbNullString, vbNullChar
                                
                                .Reboot = True
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    'And the last: we'll list all semi-legit groups ("Additional scan" only) - can contain legit entries!
    
    If bAdditional Then
      For Each vName In dLegitGroups.Keys
        
        If dLegitGroups(vName) = 2 Then
            
            sGroup = vName
            
            'get services that belong to it
            For i = 0 To UBound(sServices)
                If StrComp(sGroup, Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServices(i), "Group"), 1) = 0 Then
                    If Not dLegitService.Exists(sServices(i)) Then
                    
                        sHit = "O23 - Dependency: Microsoft Service Group '" & sGroup & "' contains unknown service:  '" & sServices(i) & "'"
                        
                        If Not IsOnIgnoreList(sHit) Then
                            
                            With result
                                .Section = "O23"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REMOVE_VALUE, HKLM, "System\CurrentControlSet\Services\" & sServices(i), "Group"
                                .Reboot = True
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                End If
            Next
        End If
      Next
    End If
    
    Set dLegitGroups = Nothing
    
    AppendErrorLogCustom "CheckO23Item_Dependency - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO23Item_Dependency"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO23Item_Drivers(sServices() As String, dLegitService As clsTrickHashTable)
    'https://www.bleepingcomputer.com/tutorials/how-malware-hides-as-a-service/
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO23Item_Drivers - Begin"
    
    '+ ����� ������������, ����� ������� ����������� �������� ��������.
    '���� �� ������� ������� ����� ������, ����������� ����� NtQuerySystemInformation � ������ �������,
    '�� ���� ��������� ��������� �� ��������� ���������� ������ ��������, � � ���� ������ ��������� ��� �� �������.
    '�.�. ��� ���� ������� �������� ����� ����� ����� ��������.
    '
    ' Uninstall Devices
    '
    'https://docs.microsoft.com/en-us/windows-hardware/drivers/install/using-setupapi-to-uninstall-devices-and-driver-packages
    'https://stackoverflow.com/questions/12756712/windows-device-uninstall-using-c
    '
    ' Uninstall Drivers
    '
    'http://www.cyberforum.ru/drivers-programming/thread1300444.html#post6857698
    
    'Enum Drivers via NtQuerySystemInformation:
    
    Const DRIVER_INFORMATION            As Long = 11
    Const SYSTEM_MODULE_SIZE            As Long = 284
    Const STATUS_INFO_LENGTH_MISMATCH   As Long = &HC0000004
    
    Dim ret             As Long
    Dim buf()           As Byte
    Dim mdl             As SYSTEM_MODULE_INFORMATION
    Dim dDriver         As clsTrickHashTable 'Drivers loaded in memory atm
    Dim sFile           As String
    Dim i               As Long
    Dim sName           As String
    Dim lType           As Long
    Dim lStart          As Long
    Dim sGroup          As String
    Dim sDisplayName    As String
    Dim bHideDisabled   As Boolean
    Dim sBuf            As String
    Dim ServState       As SERVICE_STATE
    Dim sHit            As String
    Dim result          As SCAN_RESULT
    Dim bSafe           As Boolean
    Dim bSafeModeSvc    As Boolean
    Dim sFilename       As String
    Dim sHash           As String
    
    If Not bIgnoreAllWhitelists Then
        bHideDisabled = True
    End If
    
    Set dDriver = New clsTrickHashTable
    dDriver.CompareMode = TextCompare
    
    If NtQuerySystemInformation(DRIVER_INFORMATION, ByVal 0&, 0, ret) = STATUS_INFO_LENGTH_MISMATCH Then
        ReDim buf(ret - 1)
        If NT_SUCCESS(NtQuerySystemInformation(DRIVER_INFORMATION, buf(0), ret, ret)) Then
            mdl.ModulesCount = buf(0) Or (buf(1) * &H100&) Or (buf(2) * &H10000) Or (buf(3) * &H1000000)
            If mdl.ModulesCount Then
                ReDim mdl.Modules(mdl.ModulesCount - 1)
                For ret = 0 To mdl.ModulesCount - 1
                    memcpy mdl.Modules(ret), buf(ret * SYSTEM_MODULE_SIZE + 4), SYSTEM_MODULE_SIZE
                    sFile = TrimNull(mdl.Modules(ret).Name)
                    
                    sFile = CleanServiceFileName(sFile, vbNullString, bIsDriver:=True)
                    
                    sFile = FindOnPath(sFile, True, sWinSysDir & "\Drivers")
                    
                    UpdateProgressBar "O23-D", sFile

                    If Not IsMicrosoftDriverFileEx(sFile, result.SignResult) Or Not bHideMicrosoft Then
                        dDriver.Add sFile, 0&
                    End If

                Next
            End If
        End If
    End If
    
    'Enum Drivers via Registry
    
    For i = 0 To UBound(sServices)

        sName = sServices(i)

        lType = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Type")

        If lType >= 16 Then 'not a Driver
            GoTo Continue2
        End If

        lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start")

        If (lStart = 4 And bHideDisabled) Then
            If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 4&
            GoTo Continue2
        End If

        sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DisplayName")
        
        If Len(sDisplayName) = 0 Then
            sDisplayName = sName
        Else
            If Left$(sDisplayName, 1) = "@" Then                    'extract string resource from file

                sBuf = GetStringFromBinary(, , sDisplayName)

                If 0 <> Len(sBuf) Then sDisplayName = sBuf
            End If
        End If
        
        UpdateProgressBar "O23-D", sDisplayName

        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ImagePath")
        If Len(sFile) = 0 Then GoTo Continue2

        sFile = CleanServiceFileName(sFile, sName, bIsDriver:=True)

        ServState = GetServiceRunState(sName)
        
        If Not bAutoLogSilent Then DoEvents

        bSafe = IsMicrosoftDriverFileEx(sFile, result.SignResult)

        If bSafe Then
            If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 0&
        End If
        
        If Not bSafe Then
            If OSver.IsWindows8OrGreater Then
                sFilename = GetFileName(sFile, True)
                If Not bSafe Then If StrComp(sFilename, "BthA2dp.sys", 1) = 0 Then If GetFileSHA1(sFile, , True) = "8CE29225E3425898D862EB69D491091B693A1AE0" Then bSafe = True
                If Not bSafe Then If StrComp(sFilename, "BthHfEnum.sys", 1) = 0 Then If GetFileSHA1(sFile, , True) = "8EE57413F82B7ECF1BAC041484CD878B8409C090" Then bSafe = True
            End If
        End If
        
        If Not bSafe Or Not bHideMicrosoft Then
            
            If dDriver.Exists(sFile) Then dDriver.Remove sFile
            
            sFile = FormatFileMissing(sFile)
            
            sGroup = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Group")
            bSafeModeSvc = IsSafemodeDriver(sName, sGroup)
            
            sHit = "O23 - Driver " & IIf(ServState <> SERVICE_STOPPED, "R", "S") & lStart & _
                ": " & IIf(sDisplayName = sName, sName, sDisplayName & " - (" & sName & ")") & " - " & sFile & _
                IIf(bSafeModeSvc, " (+safe mode)", "") & _
                FormatSign(result.SignResult)
            
            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O23"
                    .HitLineW = sHit
                    AddServiceToFix .Service, DELETE_SERVICE Or USE_FEATURE_DISABLE, sName, , , , ServState, True
                    AddFileToFix .File, BACKUP_FILE, sFile
                    AddRegToFix .Reg, BACKUP_KEY, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName
                    .Reboot = True
                    .CureType = SERVICE_BASED Or FILE_BASED Or REGISTRY_BASED
                End With
                AddToScanResults result
            Else
                If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 0&
            End If
        End If

Continue2:
    Next
    
    'if currently there are more running drivers (e.g. loaded dynamically)
    If dDriver.Count > 0 Then
        For i = 0 To dDriver.Count - 1
        
            sFile = dDriver.Keys(i)
            
            bSafe = False
            'skip Microsoft drivers mapped to non-existent filename
            If Not FileExists(sFile) Then
                If Not bIgnoreAllWhitelists Then
                    If oDict.DriverMapped.Exists(sFile) Then
                        bSafe = True
                    End If
                End If
            End If
            
            If Not bSafe Or bIgnoreAllWhitelists Then
                
                sDisplayName = GetFileProperty(sFile, "FileDescription")
                If Len(sDisplayName) = 0 Then
                    sDisplayName = GetFileProperty(sFile, "ProductName")
                End If
                
                sFile = FormatFileMissing(sFile)
                SignVerifyJack sFile, result.SignResult
                
                sHit = "O23 - Driver R: " & IIf(Len(sDisplayName) = 0, STR_NO_NAME, sDisplayName) & " - " & sFile & FormatSign(result.SignResult)
                
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                If Not IsOnIgnoreList(sHit) Then
                    
                    With result
                        .Section = "O23"
                        .HitLineW = sHit
                        AddFileToFix .File, REMOVE_FILE, sFile
                        .CureType = FILE_BASED
                        .Reboot = True
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    End If
    
    AppendErrorLogCustom "CheckO23Item_Drivers - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO23Item_Drivers", "Service=", sDisplayName
    If inIDE Then Stop: Resume Next
End Sub
    

Public Function IsWinServiceFileName(sFilePath As String, Optional sArgument As String) As Boolean
    
    On Error GoTo ErrorHandler:
    
    Dim sFilename As String
    Dim sArgDB As String
    
    If oDict.dSafeSvcPath.Exists(sFilePath) Then
        If Len(sArgument) = 0 Then
            IsWinServiceFileName = True: Exit Function
        Else
            sArgDB = oDict.dSafeSvcPath(sFilePath)
            If inArraySerialized(sArgument, sArgDB, "|", , , vbTextCompare) Then
                IsWinServiceFileName = True
                Exit Function
            End If
        End If
    End If
    
    'by filename
    sFilename = GetFileNameAndExt(sFilePath)
    If oDict.dSafeSvcFilename.Exists(sFilename) Then IsWinServiceFileName = True: Exit Function
    
    If StrBeginWith(sFilePath, PF_64 & "\" & "Microsoft\Exchange Server") Then IsWinServiceFileName = True: Exit Function
    
    'if service file is not in list, check if it protected by SFC, excepting AV / Firewall services
    'also, separate blacklist nedeed to identify dangerous host-files like cmd.exe / powershell e.t.c.
    If IsFileSFC(sFilePath) Then
        If Not IsLoLBin(sFilePath) Then IsWinServiceFileName = True: Exit Function
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IsWinServiceFileName", "File: " & sFilePath
    If inIDE Then Stop: Resume Next
End Function


Public Sub FixO23Item(sItem$, result As SCAN_RESULT)
    'stop & disable & delete NT service
    'O23 - Service: <displayname> - <company> - <file>
    ' (file missing) or (filesize .., MD5 ..) can be appended
    If Not bIsWinNT Then Exit Sub
    
    FixIt result
End Sub

Public Sub CheckO24Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO24Item - Begin"
    
    'activex desktop components
    Dim sDCKey$, sComponents$(), i&
    Dim sSource$, sSubscr$, sName$, sHit$, Wow64key As Boolean, result As SCAN_RESULT
    
    Wow64key = False
    
    sDCKey = "Software\Microsoft\Internet Explorer\Desktop\Components"
    sComponents = Split(Reg.EnumSubKeys(HKEY_CURRENT_USER, sDCKey, Wow64key), "|")
    
    For i = 0 To UBound(sComponents)
        If Reg.KeyExists(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), Wow64key) Then
            sSource = Reg.GetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "Source", Wow64key)
            
            sSubscr = Reg.GetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "SubscribedURL", Wow64key)
            sSubscr = UnQuote(EnvironW(sSubscr))
            sSubscr = GetLongPath(sSubscr)  ' 8.3 -> Full
            
            sName = Reg.GetString(HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), "FriendlyName", Wow64key)
            If sName = vbNullString Then sName = STR_NO_NAME
            If (Not (LCase$(sSource) = "about:home" And LCase$(sSubscr) = "about:home") And _
               Not (UCase$(sSource) = "131A6951-7F78-11D0-A979-00C04FD705A2" And UCase$(sSubscr) = "131A6951-7F78-11D0-A979-00C04FD705A2")) _
               Or Not bHideMicrosoft Then
                
                'Example: <Windows folder>\screen.html
                sSource = Replace$(sSource, "<Windows folder>", sWinDir, , , 1)
                sSource = Replace$(sSource, "<System>", sWinSysDir, , , 1)
                If Left$(sSource, 8) = "file:///" Then sSource = mid$(sSource, 9)
                
                WipeSignResult result.SignResult
                If mid$(sSource, 2, 1) = ":" Then 'If file system object
                    sSource = FormatFileMissing(sSource)
                    SignVerifyJack sSource, result.SignResult
                End If
                
                sHit = "O24 - Desktop Component " & sComponents(i) & ": " & sName & " - " & _
                    IIf(Len(sSource) <> 0, "[Source] = " & sSource, IIf(Len(sSubscr) <> 0, "[SubscribedURL] = " & sSubscr, STR_NO_FILE)) & _
                    FormatSign(result.SignResult)
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Alias = "O24"
                        .Section = "O24"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), , , CLng(Wow64key)
                        If Len(sSource) <> 0 Then AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sSource
                        If Len(sSubscr) <> 0 Then AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sSubscr
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        End If
    Next i
    
    AppendErrorLogCustom "CheckO24Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO24Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO24Item(sItem$, result As SCAN_RESULT)
    'delete the entire registry key
    'O24 - Desktop Component 1: Internet Explorer Channel Bar - 131A6951-7F78-11D0-A979-00C04FD705A2
    'O24 - Desktop Component 2: Security - %windir%\index.html
    
    FixIt result

End Sub

Public Sub FixO24Item_Post()
    Const SPIF_UPDATEINIFILE As Long = 1&
    
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, 0&, SPIF_UPDATEINIFILE 'SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    Sleep 1000
    RestartExplorer
End Sub

Public Sub RestartExplorer()

    Dim iRestartShell As Long
    Dim bShouldRestore As Boolean
    
    'AutoRestartShell can be configured to prevent Explorer restarts automatically.
    'Also, in that case the attempt to start it manually will cause Explorer to run in 'Safe Mode'.
    'So, firstly we need to ensure the value is correct.
    
    iRestartShell = Reg.GetDword(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoRestartShell")
    
    If iRestartShell <> 1 Then
        Reg.SetDwordVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoRestartShell", 1
        bShouldRestore = True
    End If
    
    If ProcessExist("explorer.exe", True) Then
    
        'We could do it in official way, e.g. with RestartManager: https://jiangsheng.net/2013/01/22/how-to-restart-windows-explorer-programmatically-using-restart-manager/
        'but, consider we are dealing with malware, it is better to just kill process without notifying loaded modules about this action
    
        KillProcessByFile sWinDir & "\" & "explorer.exe", True, 0
        Sleep 1000
        
        If Not ProcessExist("explorer.exe", True) Then
            StartExplorer
        End If
    Else
        'When explorer does not exists, it could had been shut down with Exit Code 1,
        'which means the next time it will be started in 'Safe Mode',
        'so, we need to restart it twice in that case with 0 Exit Code.
    
        StartExplorer
        Sleep 3000
        
        KillProcessByFile sWinDir & "\" & "explorer.exe", True, 0
        Sleep 1000
        
        If Not ProcessExist("explorer.exe", True) Then
            StartExplorer
        End If
    End If
    
    If bShouldRestore Then
        Reg.SetDwordVal HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoRestartShell", iRestartShell
    End If
    
    Sleep 500
    
End Sub

Public Sub StartExplorer()
    ' Run unelevated (downgrade privileges)
    ' Same as CreateExplorerShellUnelevatedTask task, that uses /NOUACCHECK switch to override task policy
    ' I guess that switch used in task scheduler to prevent recurse call
    Proc.ProcessRunUnelevated2 sWinDir & "\" & "explorer.exe"
End Sub

Public Sub ShutdownExplorer()
    KillProcessByFile sWinDir & "\" & "explorer.exe", True, 1
End Sub
    
Public Function IsOnIgnoreList(ByRef sHit$, Optional UpdateList As Boolean, Optional EraseList As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "IsOnIgnoreList - Begin", "Line: " & sHit
    
    Static isInit As Boolean
    Static aIgnoreList() As String
    
    If EraseList Then
        ReDim aIgnoreList(0)
        Exit Function
    End If
    
    sHit = ScreenHitLine(LimitHitLineLength(sHit, LIMIT_CHARS_COUNT_FOR_LOGLINE))
    
    If isInit And Not UpdateList Then
        If InArray(sHit, aIgnoreList) Then IsOnIgnoreList = True
    Else
        Dim iIgnoreNum&, i&
        
        isInit = True
        ReDim aIgnoreList(0)
        
        iIgnoreNum = Val(RegReadHJT("IgnoreNum", "0"))
        If iIgnoreNum > 0 Then ReDim aIgnoreList(iIgnoreNum)
        
        For i = 1 To iIgnoreNum
            aIgnoreList(i) = DeCrypt(RegReadHJT("Ignore" & i, vbNullString))
        Next
    End If
    
    AppendErrorLogCustom "IsOnIgnoreList - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modMain_IsOnIgnoreList", sHit
    If inIDE Then Stop: Resume Next
End Function

Public Function DateTimeToStringUS(curTime As Date) As String
    DateTimeToStringUS = Year(curTime) & _
        "/" & Right$("0" & Month(curTime), 2) & _
        "/" & Right$("0" & Day(curTime), 2) & _
        " " & Right$("0" & Hour(curTime), 2) & _
        ":" & Right$("0" & Minute(curTime), 2) & _
        ":" & Right$("0" & Second(curTime), 2)
End Function

Public Function DateTimeToString(curTime As Date) As String
    DateTimeToString = Right$("0" & Day(curTime), 2) & _
        "." & Right$("0" & Month(curTime), 2) & _
        "." & Year(curTime) & _
        " " & Right$("0" & Hour(curTime), 2) & _
        ":" & Right$("0" & Minute(curTime), 2) & _
        ":" & Right$("0" & Second(curTime), 2)
End Function

Public Sub ErrorMsg(ErrObj As ErrObject, sProcedure$, ParamArray vCodeModule())
    Dim sMsg$, sParameters$, hResult$, HRESULT_LastDll$, sErrDesc$, iErrNum&, iErrLastDll&, i&
    Dim DateTime As String, curTime As Date, ErrText$
    Dim sErrHeader$
    
    If bSkipErrorMsg Then Exit Sub
    
    sErrDesc = ErrObj.Description
    iErrNum = ErrObj.Number
    iErrLastDll = ErrObj.LastDllError
    
    If iErrNum <> 33333 And iErrNum <> 0 Then    'error defined by HJT
        hResult = ErrMessageText(CLng(iErrNum))
    End If
    
    If iErrLastDll <> 0 Then
        HRESULT_LastDll = ErrMessageText(iErrLastDll)
    End If
    
    For i = 0 To UBound(vCodeModule)
        If Not IsMissing(vCodeModule(i)) Then
            sParameters = sParameters & vCodeModule(i) & " "
        End If
    Next
    
    If AryItems(TranslateNative) Then
        sErrHeader = TranslateNative(590)
    End If
    If 0 = Len(sErrHeader) Then
        If AryItems(Translate) Then
            sErrHeader = Translate(590)
        End If
    End If
    If 0 = Len(sErrHeader) Then
        ' Emergency mode (if translation module is not initialized yet)
        sErrHeader = "Please help us improve HiJackThis+ by reporting this error." & _
            vbCrLf & vbCrLf & "Error message has been copied to clipboard." & _
            vbCrLf & "Click 'Yes' to submit." & _
            vbCrLf & vbCrLf & "Error Details: " & _
            vbCrLf & vbCrLf & "An unexpected error has occurred at function: "
    End If
    
    Dim OSData As String
    
    If ObjPtr(OSver) <> 0 Then
        OSData = OSver.Bitness & " " & OSver.OSName & IIf(Len(OSver.Edition) <> 0, " (" & OSver.Edition & ")", vbNullString) & ", " & _
            OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & ", " & _
            "Service Pack: " & OSver.SPVer & IIf(OSver.IsSafeBoot, " (Safe Boot)", vbNullString)
    End If
    
    sMsg = sErrHeader & " " & _
        sProcedure & vbCrLf & _
        "Error # " & iErrNum & IIf(iErrNum <> 0, " - " & sErrDesc, vbNullString) & _
        vbCrLf & "HRESULT: " & hResult & _
        vbCrLf & "LastDllError # " & iErrLastDll & IIf(iErrLastDll <> 0, " (" & HRESULT_LastDll & ")", vbNullString) & _
        vbCrLf & "Trace info: " & sParameters & _
        vbCrLf & vbCrLf & "Windows version: " & OSData & _
        vbCrLf & AppVerPlusName & vbCrLf & _
        "--- EOF ---"
    
    If Not bAutoLogSilent Then
        ClipboardSetText sMsg
    End If
    
    curTime = Now()
    DateTime = DateTimeToString(curTime)
    
    ErrText = " - " & sProcedure & " - #" & iErrNum
    If iErrNum <> 0 Then ErrText = ErrText & " (" & sErrDesc & ")" & IIf(Len(hResult) <> 0, " (" & hResult & ")", vbNullString)
    ErrText = ErrText & " LastDllError = " & iErrLastDll
    If iErrLastDll <> 0 Then ErrText = ErrText & " (" & HRESULT_LastDll & ")"
    If Len(sParameters) <> 0 Then ErrText = ErrText & " " & sParameters
    
    If inIDE Then Debug.Print ErrText
    
    ErrReport = ErrReport & vbCrLf & _
        "- " & DateTime & ErrText
    
    AppendErrorLogCustom ">>> ERROR:" & vbCrLf & _
        "- " & DateTime & ErrText
    
    If Not bAutoLog Then
        frmError.Show vbModeless
        frmError.Label1.Caption = sMsg
        frmError.Hide
        frmError.Show vbModal
        
'        If vbYes = MsgBoxW(sMsg, vbCritical Or vbYesNo, Translate(591)) Then
'            Dim szParams As String
'            Dim szCrashUrl As String
'            szCrashUrl = "http://safezone.cc/threads/25222/" 'https://sourceforge.net/p/hjt/_list/tickets"
'
'            'szParams = "function=" & sProcedure
'            'szParams = szParams & "&params=" & sParameters
'            'szParams = szParams & "&errorno=" & iErrNum
'            'szParams = szParams & "&errorlastdll=" & iErrLastDll
'            'szParams = szParams & "&errortxt" & sErrDesc
'            'szParams = szParams & "&winver=" & sWinVersion
'            'szParams = szParams & "&hjtver=" & App.Major & "." & App.Minor & "." & App.Revision
'            'szCrashUrl = szCrashUrl & URLEncode(szParams)
'            If True = IsOnline Then
'                ShellExecute 0&, "open", szCrashUrl, vbNullString, vbNullString, vbNormalFocus
'            Else
'                'MsgBoxW "No Internet Connection Available"
'                MsgBoxW Translate(560)
'            End If
'        End If
    End If
    
    If inIDE Then Stop
End Sub

Public Function OpenClipboardEx(hWndOwner As Long) As Boolean 'thanks to wqweto
    Dim lr          As Long
    Dim lRetry      As Long
    
    Randomize Timer
    'ClipboardClose
    
    For lRetry = 1 To 5
        lr = OpenClipboard(hWndOwner)
        If lr <> 0 Then
            OpenClipboardEx = True
            Exit Function
        End If
        Call Sleep(Rnd() * 500)
    Next
End Function

Public Function ClipboardGetText() As String
    On Error GoTo ErrorHandler
        Dim hMem As Long
        Dim ptr  As Long
        Dim Size As Long
        Dim txt  As String
        If OpenClipboardEx(g_HwndMain) Then
            hMem = GetClipboardData(CF_UNICODETEXT)
            If hMem Then
                Size = GlobalSize(hMem)
                If Size Then
                    txt = Space$(Size \ 2 - 1)
                    ptr = GlobalLock(hMem)
                    lstrcpyn ByVal StrPtr(txt), ByVal ptr, Size
                    GlobalUnlock hMem
                    ClipboardGetText = txt
                End If
            End If
            CloseClipboard
        End If
    Exit Function
ErrorHandler:
    If inIDE Then Stop: Resume Next
End Function

Public Function ClipboardSetText(sText As String) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ClipboardSetText - Begin"
    
    Dim LangNonUnicodeCode As Long
    LangNonUnicodeCode = GetSystemDefaultLCID Mod &H10000
    
    Dim hMem As Long
    Dim ptr As Long
    If OpenClipboardEx(g_HwndMain) Then
        EmptyClipboard
        If Len(sText) <> 0 Then
            hMem = GlobalAlloc(GMEM_MOVEABLE, 4)
            If hMem <> 0 Then
                ptr = GlobalLock(hMem)
                If ptr <> 0 Then
                    GetMem4 LangNonUnicodeCode, ByVal ptr
                    GlobalUnlock hMem
                    If SetClipboardData(CF_LOCALE, hMem) = 0 Then
                        GlobalFree hMem
                    End If
                End If
            End If
            hMem = GlobalAlloc(GMEM_MOVEABLE, LenB(sText) + 2)
            If hMem <> 0 Then
                ptr = GlobalLock(hMem)
                If ptr <> 0 Then
                    lstrcpyn ByVal ptr, ByVal StrPtr(sText), LenB(sText)
                    GlobalUnlock hMem
                    ClipboardSetText = SetClipboardData(CF_UNICODETEXT, hMem)
                    If Not ClipboardSetText Then
                        GlobalFree hMem
                    End If
                End If
            End If
        End If
        CloseClipboard
    End If
    
    AppendErrorLogCustom "ClipboardSetText - End"
    Exit Function
ErrorHandler:
    'ErrorMsg Err, "ClipboardSetText" ' Out of stack space
    If inIDE Then Stop: Resume Next
End Function

Public Function ClipboardCopyFile(sFile As String) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "ClipboardCopyFile - Begin"
    
    Dim pidlFile As Long
    Dim pidlChild As Long
    Dim psfParent As IShellFolder
    Dim psData As IDataObject
    Dim Redirect As Boolean, bOldStatus As Boolean
    
    If Not FileExists(sFile) Then Exit Function
    
    Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    pidlFile = ILCreateFromPathW(StrPtr(sFile))
    
    If pidlFile <> 0 Then
    
        Call SHBindToParent(pidlFile, IID_IShellFolder, psfParent, pidlChild)
        If (psfParent Is Nothing) Or (pidlChild = 0) Then
            If inIDE Then Debug.Print "Failed to bind to parent of: " & sFile
        Else
            Dim aPidl() As Long
            ReDim aPidl(0) As Long
            aPidl(0) = pidlChild
            
            Call psfParent.GetUIObjectOf(0&, 1&, aPidl(0), IID_IDataObject, 0&, psData)
            If (psData Is Nothing) Then
                If inIDE Then Debug.Print "Failed to get IDataObject interface for: " & sFile
            Else
                Call OleSetClipboard(psData)
                Call OleFlushClipboard
            End If
        End If
        
        CoTaskMemFree pidlFile
    End If
    
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "ClipboardCopyFile - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ClipboardCopyFile"
    If inIDE Then Stop: Resume Next
End Function

Public Sub AppendErrorLogNoErr(ErrObj As ErrObject, sProcedure As String, ParamArray CodeModule())
    'to append error log without displaying error message to user
    
    On Error Resume Next
    
    Dim i           As Long
    Dim DateTime    As String
    Dim ErrText     As String
    Dim sErrDesc    As String
    Dim iErrNum     As Long
    Dim iErrLastDll As Long
    Dim hResult     As String
    Dim HRESULT_LastDll As String
    Dim sParameters As String

    DateTime = DateTimeToString(Now())
    
    sErrDesc = ErrObj.Description
    iErrNum = ErrObj.Number
    iErrLastDll = ErrObj.LastDllError
    
    If iErrNum <> 33333 And iErrNum <> 0 Then    'error defined by HJT
        hResult = ErrMessageText(CLng(iErrNum))
    End If
    
    If iErrLastDll <> 0 Then
        HRESULT_LastDll = ErrMessageText(iErrLastDll)
    End If
    
    For i = 0 To UBound(CodeModule)
        sParameters = sParameters & CodeModule(i) & " "
    Next

    ErrText = " - " & sProcedure & " - #" & iErrNum
    If iErrNum <> 0 Then ErrText = ErrText & " (" & sErrDesc & ")" & IIf(Len(hResult) <> 0, " (" & hResult & ")", vbNullString)
    ErrText = ErrText & " LastDllError = " & iErrLastDll
    If iErrLastDll <> 0 Then ErrText = ErrText & " (" & HRESULT_LastDll & ")"
    If Len(sParameters) <> 0 Then ErrText = ErrText & " " & sParameters
    
    ErrReport = ErrReport & vbCrLf & _
        "- " & DateTime & ErrText
    
    AppendErrorLogCustom ">>> ERROR:" & vbCrLf & _
        "- " & DateTime & ErrText
End Sub

Public Function ErrMessageText(lCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
    Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800&
    Const MSG_SIZE = 300&
    
    Dim sRtrnMsg   As String
    Dim lret       As Long
    Dim hLib       As Long
    
    sRtrnMsg = String$(MSG_SIZE, 0&)
    hLib = GetModuleHandle(StrPtr("wininet.dll"))
    If hLib = 0 Then
        hLib = LoadLibrary(StrPtr("wininet.dll"))
    End If
    
    lret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_FROM_HMODULE Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal hLib, lCode, g_CurrentLangID, StrPtr(sRtrnMsg), MSG_SIZE, 0&)
    
    If Err.LastDllError = 1815 Then 'lang id not found => fallback to english
        lret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_FROM_HMODULE Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal hLib, lCode, &H409&, StrPtr(sRtrnMsg), MSG_SIZE, 0&)
    End If
    If lret > 0 Then
        ErrMessageText = Left$(sRtrnMsg, lret)
        ErrMessageText = Replace$(ErrMessageText, vbCrLf, vbNullString)
    End If
End Function

'Public Sub CheckDateFormat()
'    Dim sBuffer$, uST As SYSTEMTIME
'    With uST
'        .wDay = 10
'        .wMonth = 11
'        .wYear = 2003
'    End With
'    sBuffer = String$(255, 0)
'    GetDateFormat 0&, 0&, uST, 0&, StrPtr(sBuffer), 255&
'    sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
'
'    'last try with GetLocaleInfo didn't work on Win2k/XP
'    If InStr(sBuffer, "10") < InStr(sBuffer, "11") Then
'        bIsUSADateFormat = False
'        'msgboxW "sBuffer = " & sBuffer & vbCrLf & "10 < 11, so bIsUSADateFormat False"
'    Else
'        bIsUSADateFormat = True
'        'msgboxW sBuffer & vbCrLf & "10 !< 11, so bIsUSADateFormat True"
'    End If
'
'    'Dim lLndID&, sDateFormat$
'    'lLndID = GetSystemDefaultLCID()
'    'sDateFormat = String$(255, 0)
'    'GetLocaleInfo lLndID, LOCALE_SSHORTDATE, sDateFormat, 255
'    'sDateFormat = left$(sDateFormat, InStr(sDateFormat, vbnullchar) - 1)
'    'If sDateFormat = vbNullString Then Exit Sub
'    ''sDateFormat = "dd-MM-yy" or "M/d/yy"
'    ''I hope this works - dunno what happens in
'    ''yyyy-mm-dd or yyyy-dd-mm format
'    'If InStr(1, sDateFormat, "d", vbTextCompare) < _
'    '   InStr(1, sDateFormat, "m", vbTextCompare) Then
'    '    bIsUSADateFormat = False
'    'Else
'    '    bIsUSADateFormat = True
'    'End If
'End Sub

Public Function UnEscape(ByVal StringToDecode As String) As String
    Dim i As Long
    Dim acode As Integer, lTmp As Long, HexChar As String

    On Error GoTo ErrorHandler

'    Set scr = CreateObject("MSScriptControl.ScriptControl")
'    scr.Language = "VBScript"
'    scr.Reset
'    Escape = scr.Eval("unescape(""" & s & """)")

    UnEscape = StringToDecode

    If InStr(UnEscape, "%") = 0 Then
         Exit Function
    End If
    For i = Len(UnEscape) To 1 Step -1
        acode = Asc(mid$(UnEscape, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars

            Case 37
                ' Decode % value
                HexChar = UCase$(mid$(UnEscape, i + 1, 2))
                If HexChar Like "[0123456789ABCDEF][0123456789ABCDEF]" Then
                    lTmp = CLng("&H" & HexChar)
                    UnEscape = Left$(UnEscape, i - 1) & Chr$(lTmp) & mid$(UnEscape, i + 3)
                End If
        End Select
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "UnEscape", "string:", StringToDecode
End Function

Public Function HasSpecialCharacters(sName$) As Boolean
    'function checks for special characters in string,
    'like Chinese or Japanese.
    'Used in CheckO3Item (IE Toolbar)
    HasSpecialCharacters = False
    
    'function disabled because of proper DBCS support
    Exit Function
    
    If Len(sName) <> lstrlen(StrPtr(sName)) Then
        HasSpecialCharacters = True
        Exit Function
    End If
    
    If Len(sName) <> LenB(StrConv(sName, vbFromUnicode)) Then
        HasSpecialCharacters = True
        Exit Function
    End If
End Function

Public Function CheckForReadOnlyMedia() As Boolean

    If Not CheckFileAccess(AppPath(), GENERIC_WRITE) Then
    
        bNoWriteAccess = True
        'It looks like you're running HiJackThis from
        'a read-only device like a CD-ROM.
        'If you want to make backups of items you fix,
        'you must copy HiJackThis.exe to your hard disk
        'first, and run it from there.
        MsgBoxW Translate(7), vbExclamation
    Else
        CheckForReadOnlyMedia = True
    End If
    
End Function

Public Sub SetAllFontCharset(frm As Form, Optional sFontName As String, Optional sFontSize As String, Optional bFontBold As Boolean)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "SetAllFontCharset - Begin"

    Dim Ctl         As Control
    Dim ctlBtn      As VBCCR17.CommandButtonW
    Dim ctlOptBtn   As VBCCR17.OptionButtonW
    Dim ctlCheckBox As VBCCR17.CheckBoxW
    Dim ctlTxtBox   As VBCCR17.TextBoxW
    Dim ctlLstBox   As VBCCR17.ListBoxW
    Dim CtlLbl      As VBCCR17.LabelW
    Dim CtlFrame    As VBCCR17.FrameW
    Dim CtlCombo    As VBCCR17.ComboBoxW
    Dim CtlTree     As VBCCR17.TreeView
    Dim CtlPict     As PictureBox
    Dim iOldTop     As Long
    Dim iOldSel     As Long
    
    If frm Is frmMain Then
        'save selected position on listbox
        If frmMain.lstResults.ListIndex <> -1 Then
            iOldSel = frmMain.lstResults.ListIndex
        End If
        
        iOldTop = frmMain.lstResults.TopIndex
    End If
    
    For Each Ctl In frm.Controls
        Select Case TypeName(Ctl)
            Case "CommandButtonW"
                Set ctlBtn = Ctl
                SetFontCharSet ctlBtn, sFontName, sFontSize, bFontBold
            Case "OptionButtonW"
                Set ctlOptBtn = Ctl
                SetFontCharSet ctlOptBtn, sFontName, sFontSize, bFontBold
            Case "TextBoxW"
                Set ctlTxtBox = Ctl
                SetFontCharSet ctlTxtBox, sFontName, sFontSize, bFontBold
            Case "ListBoxW"
                Set ctlLstBox = Ctl
                SetFontCharSet ctlLstBox, sFontName, sFontSize, bFontBold
            Case "LabelW"
                Set CtlLbl = Ctl
                SetFontCharSet CtlLbl, sFontName, sFontSize, bFontBold
            Case "CheckBoxW"
                Set ctlCheckBox = Ctl
                'If ctlCheckBox.Name <> "chkConfigTabs" Then
                    SetFontCharSet ctlCheckBox, sFontName, sFontSize, bFontBold
                'End If
            Case "FrameW"
                Set CtlFrame = Ctl
                SetFontCharSet CtlFrame, sFontName, sFontSize, bFontBold
            Case "ComboBoxW"
                Set CtlCombo = Ctl
                If Not ((Ctl Is frmMain.cmbFont) Or (Ctl Is frmMain.cmbFontSize) Or (Ctl Is frmMain.cmbDefaultFont) Or (Ctl Is frmMain.cmbDefaultFontSize)) Then
                    SetFontCharSet CtlCombo, sFontName, sFontSize, bFontBold
                End If
            Case "TreeView"
                Set CtlTree = Ctl
                SetFontCharSet CtlTree, sFontName, sFontSize, bFontBold
            Case "PictureBox"
                Set CtlPict = Ctl
                SetFontCharSet CtlPict, sFontName, sFontSize, bFontBold
        End Select
    Next Ctl
    
    're-applying subclassing is required, since the window destroys itself and re-creates as soon as font charset is changed
    If frm Is frmMain Then
        SubClassScroll_ScanResults True
        
        'restore listbox sel. position and scroll position
        If iOldSel > 0 Then frmMain.lstResults.ListIndex = iOldSel
        If iOldTop <> -1 Then frmMain.lstResults.TopIndex = iOldTop
    End If
    
    AppendErrorLogCustom "SetAllFontCharset - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_SetAllFontCharset"
    If inIDE Then Stop: Resume Next
End Sub

'reset font to initial defaults
Public Sub SetFontDefaults(Ctl As Control, Optional bRelease As Boolean)

    'Here we are saving default state of control and change the state to defaults before changing font,
    'because previous font can be such that has no some property (like it can be BOLD only).
    'In such case after changing font will be always BOLDed.
    
    Const DEFAULT_FONT_NAME As String = "Tahoma"
    
    If bRelease Then
        Set dFontDefault = Nothing
        Erase aFontDefProp
        Exit Sub
    End If
    
    If (dFontDefault Is Nothing) Then
        Set dFontDefault = New clsTrickHashTable
        ReDim aFontDefProp(0)
    End If
    
    Dim CtlPath As String
    Dim idx As Long
    CtlPath = Ctl.Parent.Name & "." & Ctl.Name
    
    If dFontDefault.Exists(CtlPath) Then
        idx = dFontDefault(CtlPath)
        With Ctl.Font
            .Bold = aFontDefProp(idx).Bold
            .Italic = aFontDefProp(idx).Italic
            .Underline = aFontDefProp(idx).Underline
            .Size = aFontDefProp(idx).Size
        End With
    Else
        idx = UBound(aFontDefProp) + 1
        dFontDefault.Add CtlPath, idx
        ReDim Preserve aFontDefProp(idx)
        With aFontDefProp(idx)
            .Bold = Ctl.Font.Bold
            .Italic = Ctl.Font.Italic
            .Underline = Ctl.Font.Underline
            .Size = Ctl.Font.Size
        End With
    End If
    With Ctl.Font
        .Name = DEFAULT_FONT_NAME
        .Weight = 400
        .Charset = DEFAULT_CHARSET
        .StrikeThrough = False
        'when font name is changed, all properties are resetted automatically => should re-apply
        idx = dFontDefault(CtlPath)
        With Ctl.Font
            .Bold = aFontDefProp(idx).Bold
            .Italic = aFontDefProp(idx).Italic
            .Underline = aFontDefProp(idx).Underline
            .Size = aFontDefProp(idx).Size
        End With
    End With
End Sub

'return BOOL, whether control is a list and require to apply separate font setting for it
Private Function IsControlRepresentList_ForFont(Ctl As Control) As Boolean
    Static CtlList() As String
    Dim CtlPath As String
    
    CtlPath = Ctl.Parent.Name & "." & Ctl.Name
    
    If 0 = AryPtr(CtlList) Then
        ReDim CtlList(16)
        CtlList(0) = "frmMain.lstResults"
        CtlList(1) = "frmMain.lstIgnore"
        CtlList(2) = "frmMain.lstBackups"
        CtlList(3) = "frmMain.lstHostsMan"
        CtlList(4) = "frmStartupList2.tvwMain"
        CtlList(5) = "frmADSspy.lstADSFound"
        CtlList(6) = "frmADSspy.txtADSContent"
        CtlList(7) = "frmADSspy.txtScanFolder"
        CtlList(8) = "frmCheckDigiSign.txtPaths"
        CtlList(9) = "frmCheckDigiSign.txtExtensions"
        CtlList(10) = "frmProcMan.lstProcessManager"
        CtlList(11) = "frmProcMan.lstProcManDLLs"
        CtlList(12) = "frmUninstMan.lstUninstMan"
        CtlList(13) = "frmUninstMan.txtName"
        CtlList(14) = "frmUnlockRegKey.txtKeys"
        CtlList(15) = "frmRegTypeChecker.txtKeys"
        CtlList(16) = "frmHostsMan.lstHostsMan"
    End If
    
    If InArray(CtlPath, CtlList, , , 1) Then
        IsControlRepresentList_ForFont = True
    End If
End Function

Public Sub SetFontCharSet(Ctl As Control, Optional ByVal sFontName As String, Optional ByVal sFontSize As String, Optional ByVal bFontBold As Boolean)
    On Error GoTo ErrorHandler:
    
    'A big thanks to 'Gun' and 'Adult', two Japanese users
    'who helped me greatly with this
    
    'https://msdn.microsoft.com/en-us/library/aa241713(v=vs.60).aspx
    
    Static isInit As Boolean
    Static lLCID As Long
    
    Dim bNonUsCharset As Boolean
    Dim ControlFont As Font
    Dim lFontSize As Long
    Dim bLists As Boolean
    
    '//TODO:
    'Set default Hewbrew 'Non-Unicode: Hebrew (0x40D)' to Arial Unicode MS (after testing)
    
    SetFontDefaults Ctl
    
    If IsControlRepresentList_ForFont(Ctl) Then
        bLists = True
    Else
        'use font defaults
        sFontName = g_DefaultFontName
        sFontSize = g_DefaultFontSize
        bFontBold = Ctl.Font.Bold
    End If
    
    Set ControlFont = Ctl.Font
    
    If Len(sFontName) <> 0 And sFontName <> "Automatic" Then 'if font specified explicitly by user
        ControlFont.Name = sFontName
        
        If sFontSize = "Auto" Or Len(sFontSize) = 0 Then
            lFontSize = 8
        Else
            lFontSize = CLng(sFontSize)
        End If
        ControlFont.Size = lFontSize
        
        'if Hebrew
        'https://msdn.microsoft.com/en-us/library/cc194829.aspx
        
        If OSver.LangDisplayCode = &H40D& Or OSver.LangNonUnicodeCode = &H40D& Then
            ControlFont.Charset = HEBREW_CHARSET
        End If
        ControlFont.Bold = bFontBold
        
        Exit Sub
    End If
    
    bNonUsCharset = True
    
    If Not isInit Then
        lLCID = OSver.LCID_UserDefault
        isInit = True
    End If
    
    'Hebrew default behaviour -> choose "Miriam", Size 10 (thanks to @limelect for tests)
    If OSver.LangDisplayCode = &H40D& Or OSver.LangNonUnicodeCode = &H40D& Then
        If FontExist("Miriam") Then
            ControlFont.Name = "Miriam"
            ControlFont.Size = 10
            ControlFont.Charset = HEBREW_CHARSET
            ControlFont.Bold = bFontBold
            Exit Sub
        End If
    End If
    
    Select Case lLCID
         Case &H404 ' Traditional Chinese
            ControlFont.Charset = CHINESEBIG5_CHARSET
            ControlFont.Name = ChrW$(&H65B0) & ChrW$(&H7D30) & ChrW$(&H660E) & ChrW$(&H9AD4)   'New Ming-Li
         Case &H411 ' Japan
            ControlFont.Charset = SHIFTJIS_CHARSET
            ControlFont.Name = ChrW$(&HFF2D) & ChrW$(&HFF33) & ChrW$(&H20) & ChrW$(&HFF30) & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
         Case &H412 ' Korea UserLCID
            ControlFont.Charset = HANGEUL_CHARSET
            ControlFont.Name = ChrW$(&HAD74) & ChrW$(&HB9BC)
         Case &H804 ' Simplified Chinese
            ControlFont.Charset = CHINESESIMPLIFIED_CHARSET
            ControlFont.Name = ChrW$(&H5B8B) & ChrW$(&H4F53)
         Case Else ' The other countries
            ControlFont.Charset = DEFAULT_CHARSET
            ControlFont.Name = "Tahoma"
            bNonUsCharset = False
    End Select
    
    If sFontSize = "Auto" Or Len(sFontSize) = 0 Then
        If bNonUsCharset And bLists Then
            lFontSize = 9
        Else
            lFontSize = 8
        End If
    Else
        lFontSize = CLng(sFontSize)
    End If
    
    ControlFont.Size = lFontSize
    ControlFont.Bold = bFontBold
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_SetFontCharSet"
    If inIDE Then Stop: Resume Next
End Sub

Public Function FontExist(sFontName As String) As Boolean
    Dim i As Long
    For i = 0 To Screen.FontCount - 1
        If StrComp(sFontName, Screen.Fonts(i), 1) = 0 Then
            FontExist = True
            Exit For
        End If
    Next i
End Function

Private Function TrimNull(s$) As String
    TrimNull = Left$(s, lstrlen(StrPtr(s)))
End Function

Public Function CheckForStartedFromTempDir() As Boolean
    'if user picks 'run from current location when downloading HiJackThis.exe,
    'or runs file directly from zip file, exe will be ran from temp folder,
    'meaning a reboot or cache clean could delete it, as well any backups
    'made. Also the user won't be able to find the exe anymore :P
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckForStartedFromTempDir - Begin"
    
    Dim v1          As String
    Dim v2          As String
    Dim cnt         As Long
    Dim sBuffer     As String
    Dim RunFromTemp As Boolean
    Dim sMsg        As String
    
'    sMsg = "HiJackThis appears to have been started from a temporary " & _
'               "folder. Since temp folders tend to be be emptied regularly, " & _
'               "it's wise to copy HiJackThis.exe to a folder of its own, " & _
'               "for instance C:\Program Files\HiJackThis." & vbCrLf & _
'               "This way, any backups that will be made of fixed items " & _
'               "won't be lost." & vbCrLf & vbCrLf & _
'               "May I unpack HJT to desktop for you ?"
'               '"Please quit HiJackThis and copy it to a separate folder " & _
'               '"first before fixing any items."

    'Just too many words
    'User can be shocked and he will close this program immediately and forewer :)
    'l'll try this simple (just only this time):
    
    'Launch from the archive is forbidden !" & vbCrLf & vbCrLf & "May I unzip to desktop for you ?"
    sMsg = TranslateNative(8)
    
    ' �������� �� ������ �� ������
    If Len(TempCU) <> 0& Then
    
        If StrBeginWith(AppPath(), TempCU) Then RunFromTemp = True
        If Not RunFromTemp Then

            'fix, ����� app.path ������������ � ����� 8.3
            sBuffer = String$(MAX_PATH, vbNullChar)
            cnt = GetLongPathName(StrPtr(AppPath()), StrPtr(sBuffer), Len(sBuffer))
            If cnt Then
                v1 = Left$(sBuffer, cnt)
            Else
                v1 = AppPath()
            End If

            sBuffer = String$(MAX_PATH, vbNullChar)
            cnt = GetLongPathName(StrPtr(TempCU), StrPtr(sBuffer), Len(sBuffer))
            If cnt Then
                v2 = Left$(sBuffer, cnt)
            Else
                v2 = TempCU
            End If
            
            If Len(v1) <> 0 And Len(v2) <> 0 And StrBeginWith(v1, v2) Then RunFromTemp = True
        End If
        
        If RunFromTemp And (Len(g_sCommandLine) = 0) Then
            'msgboxW "������ �� ������ �������� !" & vbCrLf & "����������� �� ������� ���� ��� ��� ?", vbExclamation, AppName
            If MsgBoxW(sMsg, vbExclamation Or vbYesNo, g_AppName) = vbYes Then
                Dim NewFile As String
                NewFile = Desktop & "\HiJackThis+\" & AppExeName(True)
                MkDirW NewFile, True
                If FileExists(NewFile) Then     ', Cache:=NO_CACHE
                    SetFileAttributes StrPtr(NewFile), GetFileAttributes(StrPtr(NewFile)) And Not FILE_ATTRIBUTE_READONLY
                    DeleteFileEx NewFile
                End If
                CopyFile StrPtr(AppPath(True)), StrPtr(NewFile), ByVal 0&
                If FileExists(NewFile) Then     ', Cache:=NO_CACHE
                    frmMain.ReleaseMutex
                    Proc.ProcessRun NewFile     ', "/twice"
                    CheckForStartedFromTempDir = True
                Else
                    'Could not unzip file to Desktop! Please, unzip it manually.
                    MsgBoxW Translate(1007), vbCritical
                    CheckForStartedFromTempDir = True
                    End
                End If
            Else
                CheckForStartedFromTempDir = True
            End If
        End If
    End If
    
    AppendErrorLogCustom "CheckForStartedFromTempDir - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CheckForStartedFromTempDir"
    If inIDE Then Stop: Resume Next
End Function

Public Function RestartSystem(Optional sExtraPrompt$, Optional bSilent As Boolean, Optional bForceRestartOnServer As Boolean) As Boolean
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "RestartSystem - Begin"
    
    Dim OpSysSet As Object
    Dim OpSys As Object
    Dim lret As Long
    
    If OSver.IsServer And Not bForceRestartOnServer Then
        'The server needs to be rebooted to complete required operations. Please, do it on your own.
        MsgBoxW TranslateNative(352), vbInformation
        Exit Function
    End If
    
    'HiJackThis needs to restart the system to apply the changes.
    'Please, save your work and press 'YES' if you agree to reboot now.
    If Not bSilent Then
        If MsgBoxW(IIf(Len(sExtraPrompt) <> 0, sExtraPrompt & vbCrLf & vbCrLf, vbNullString) & TranslateNative(350), vbYesNo Or vbQuestion) = vbNo Then
            Exit Function
        End If
    End If
    
    SetCurrentProcessPrivileges "SeShutdownPrivilege"
    
    If bIsWinNT Then
        If OSver.IsWindows7OrGreater Then
            lret = InitiateSystemShutdownExW(0, 0, 0, 1, 1, SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION Or SHTDN_REASON_FLAG_PLANNED)
        Else
            lret = ExitWindowsEx(EWX_REBOOT Or EWX_FORCEIFHUNG, SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION Or SHTDN_REASON_FLAG_PLANNED)
        End If
        If lret = 0 Then
            If RunWMI_Service(bWait:=True, bAskBeforeLaunch:=False, bSilent:=bSilent) Then
                'select * from Win32_OperatingSystem where Primary=true
                Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery(Caes_Decode("thqllE * yMLL tNU89LxaXgXmdkfTBxAnx LyxMB kUNTJ]f=eej\"))
                For Each OpSys In OpSysSet
                    RestartSystem = (0 = OpSys.Reboot())
                Next
            End If
        Else
            RestartSystem = True
        End If
    Else
        RestartSystem = SHRestartSystemMB(g_HwndMain, StrPtr(sExtraPrompt), EWX_FORCE)
    End If
    
    AppendErrorLogCustom "RestartSystem - End"
  Exit Function
ErrorHandler:
    ErrorMsg Err, "RestartSystem"
    If inIDE Then Stop: Resume Next
End Function

Public Function IsIPAddress(sIP$) As Boolean
    'IsIPAddress = IIf(inet_addr(sIP) <> -1, True, False)
    'can't really trust this API, sometimes it bails when the fourth
    'octet is >127
    Dim sOctets$()
    If InStr(sIP, ".") = 0 Then Exit Function
    sOctets = Split(sIP, ".")
    If UBound(sOctets) = 3 Then
        If IsNumeric(sOctets(0)) And _
           IsNumeric(sOctets(1)) And _
           IsNumeric(sOctets(2)) And _
           IsNumeric(sOctets(3)) Then
            If (sOctets(0) >= 0 And sOctets(0) <= 255) And _
               (sOctets(1) >= 0 And sOctets(1) <= 255) And _
               (sOctets(2) >= 0 And sOctets(2) <= 255) And _
               (sOctets(3) >= 0 And sOctets(3) <= 255) Then
                IsIPAddress = True
            End If
        End If
    End If
End Function

Public Function DomainHasDoubleTLD(sDomain$) As Boolean
    Dim sDoubleTLDs$(), i&
    sDoubleTLDs = Split(".co.uk|" & _
                        ".da.ru|" & _
                        ".h1.ru|" & _
                        ".me.uk|" & _
                        ".ss.ru|" & _
                        ".xu.pl", "|")
                        '".com.au|" & _
                        ".com.br|" & _
                        ".1gb.ru|" & _
                        ".biz.ua|" & _
                        ".jps.ru|" & _
                        ".psn.cn|" & _
                        ".spb.ru|" & _
                        'above stuff somehow isn't recognized by IE
                        'as a double TLD - it's not a bug, it's a feature!

    For i = 0 To UBound(sDoubleTLDs)
        If InStr(sDomain, sDoubleTLDs(i)) = Len(sDomain) - Len(sDoubleTLDs(i)) + 1 Then
            DomainHasDoubleTLD = True
            Exit Function
        End If
    Next i
End Function

Public Function MapSIDToUsername(ByVal sSid As String) As String
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "MapSIDToUsername - Begin", "SID: " & sSid
    
    '   PURPOSE: there are certain builtin accounts on Windows NT which do not have a mapped
    '   account name. LookupAccountSid will return the error ERROR_NONE_MAPPED.  This function
    '   generates SIDs for the following accounts that are not mapped:
    '    * ACCOUNT OPERATORS
    '    * SYSTEM OPERATORS
    '    * PRINTER OPERATORS
    '    * BACKUP OPERATORS
    '   the other SID it creates is a LOGON SID, it has a prefix of S-1-5-5.  a LOGON SID is a
    '   unique identifier for a user's logon session.
    
    If Len(sSid) = 0 Then Exit Function
    
    'Force ENG form, because by default LookupAccountSid() returns 'System' account name in localized form
    'SetThreadLocale is not applicable!
    
    If StrComp(sSid, ".DEFAULT", 1) = 0 Then
        MapSIDToUsername = "SYSTEM"
        Exit Function
    End If
    
    If StrComp(sSid, "S-1-5-18", 1) = 0 Then
        MapSIDToUsername = "SYSTEM"
        Exit Function
    End If
    
    'map predefined in HKU root for speed optimiz.
    
    If StrComp(sSid, "S-1-5-19", 1) = 0 Then
        MapSIDToUsername = "LOCAL SERVICE"
        Exit Function
    End If

    If StrComp(sSid, "S-1-5-20", 1) = 0 Then
        MapSIDToUsername = "NETWORK SERVICE"
        Exit Function
    End If
    
    If StrComp(sSid, SID_TEMPLATE, 1) = 0 Then
        MapSIDToUsername = "Profile Template"
        Exit Function
    End If
    
    Dim bufSid() As Byte
    Dim AccName As String
    Dim AccDomain As String
    Dim AccType As Long
    Dim ccAccName As Long
    Dim ccAccDomain As Long
    Dim vOtherName()
    Dim tpSidAuth As SID_IDENTIFIER_AUTHORITY
    Dim pSid(3) As Long
    Dim psidLogonSid As Long
    Dim psidCheck As Long
    Dim i As Long
    
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
    
    bufSid = CreateBufferedSID(sSid)
    
    If AryItems(bufSid) Then
    
        AccName = String$(MAX_NAME, 0)
        AccDomain = String$(MAX_NAME, 0)
        ccAccName = Len(AccName)
        ccAccDomain = Len(AccDomain)
        psidCheck = VarPtr(bufSid(0))
    
        If 0 <> LookupAccountSid(0&, psidCheck, StrPtr(AccName), ccAccName, StrPtr(AccDomain), ccAccDomain, AccType) Then
        
            MapSIDToUsername = Left$(AccName, ccAccName)
            
        Else
        
            If Err.LastDllError = ERROR_NONE_MAPPED Then
            
                vOtherName = Array("Account operators", "Server operators", "Printer operators", "Backup operators")
            
                ' Create account operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ACCOUNT_OPS, 0, 0, 0, 0, 0, 0, pSid(0))

                ' Create system operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_SYSTEM_OPS, 0, 0, 0, 0, 0, 0, pSid(1))
        
                ' Create printer operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_PRINT_OPS, 0, 0, 0, 0, 0, 0, pSid(2))
        
                ' Create backup operators.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_BACKUP_OPS, 0, 0, 0, 0, 0, 0, pSid(3))

                ' Create a logon SID.
                Call AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_LOGON_IDS_RID, 0, 0, 0, 0, 0, 0, 0, psidLogonSid)
                
                '*psnu =  SidTypeAlias;
                
                If EqualPrefixSid(psidCheck, psidLogonSid) Then
                    MapSIDToUsername = "LOGON SID"
                Else
                    For i = 0 To UBound(pSid)
                        If EqualSid(psidCheck, pSid(i)) Then
                            MapSIDToUsername = vOtherName(i)
                            Exit For
                        End If
                    Next
                End If
                
                For i = 0 To UBound(pSid)
                    FreeSid pSid(i)
                Next
                FreeSid psidLogonSid
            End If
        End If
    End If
    
    AppendErrorLogCustom "MapSIDToUsername - End"
  Exit Function
ErrorHandler:
    ErrorMsg Err, "MapSIDToUsername", "SID: ", sSid
    If inIDE Then Stop: Resume Next
End Function

Public Sub SilentDeleteOnReboot(sCmd$)
    Dim sDummy$, sFilename$
    'sCmd is all command-line parameters, like this
    '/param1 /deleteonreboot c:\progra~1\bla\bla.exe /param3
    '/param1 /deleteonreboot "c:\program files\bla\bla.exe" /param3
    
    '/deleteonreboot
    sDummy = mid$(sCmd, InStr(sCmd, "deleteonreboot") + Len("deleteonreboot") + 1)
    If InStr(sDummy, """") = 1 Then
        'enclosed in quotes, chop off at next quote
        sFilename = mid$(sDummy, 2)
        sFilename = Left$(sFilename, InStr(sFilename, """") - 1)
    Else
        'no quotes, chop off at next space if present
        If InStr(sDummy, " ") > 0 Then
            sFilename = Left$(sDummy, InStr(sDummy, " ") - 1)
        Else
            sFilename = sDummy
        End If
    End If
    DeleteFileOnReboot sFilename, True
End Sub

Public Function IsProcedureAvail(ByVal ProcedureName As String, ByVal DllFilename As String) As Boolean
    AppendErrorLogCustom "IsProcedureAvail - Begin", "Function: " & ProcedureName, "Dll: " & DllFilename
    Static bInit As Boolean
    Dim hModule As Long, procAddr As Long
    If Not bInit Then
        bInit = True
        Set oDictProcAvail = New clsTrickHashTable
    End If
    If oDictProcAvail.Exists(ProcedureName) Then
        IsProcedureAvail = oDictProcAvail(ProcedureName)
        Exit Function
    End If
    hModule = LoadLibrary(StrPtr(DllFilename))
    If hModule Then
        procAddr = GetProcAddress(hModule, StrPtr(StrConv(ProcedureName, vbFromUnicode)))
        FreeLibrary hModule
    End If
    IsProcedureAvail = (procAddr <> 0)
    oDictProcAvail.Add ProcedureName, IsProcedureAvail
    AppendErrorLogCustom "IsProcedureAvail - End"
End Function

Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String = " ") As VbMsgBoxResult
    Dim hActiveWnd As Long, hMyWnd As Long, frm As Form
    If inIDE Then
        MsgBoxW = VBA.MsgBox(Prompt, Buttons, Title) 'subclassing walkaround
    Else
        hActiveWnd = GetForegroundWindow()
        For Each frm In Forms
            If frm.hWnd = hActiveWnd Then hMyWnd = hActiveWnd: Exit For
        Next
        MsgBoxW = MessageBox(IIf(hMyWnd <> 0, hMyWnd, g_HwndMain), StrPtr(Prompt), StrPtr(Title), ByVal Buttons)
    End If
End Function

Public Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String = " ") As VbMsgBoxResult
    MsgBox = MsgBoxW(Prompt, Buttons, Title)
End Function

Public Function UnQuote(stri As String) As String   ' Trim quotes
    If Len(stri) = 0 Then Exit Function
    If Left$(stri, 1) = """" And Right$(stri, 1) = """" Then
        UnQuote = mid$(stri, 2, Len(stri) - 2)
    Else
        UnQuote = stri
    End If
End Function

Public Sub ReInitScanResults()  'Global results structure will be cleaned

    'ReDim Scan.Globals(0)
    ReDim Scan(0)

End Sub

Public Sub InitVariables()

    On Error GoTo ErrorHandler:
    
    'Const CSIDL_LOCAL_APPDATA       As Long = &H1C&
    'Const CSIDL_COMMON_PROGRAMS     As Long = &H17&
    'Const FOLDERID_ComputerFolderStr As String = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"

    'SysDisk
    'sWinDir
    'sWinSysDir
    'sSysNativeDir
    'sSysDir (the same as sWinSysDir)
    'sWinSysDirWow64
    'PF_32
    'PF_64
    'AppData
    'LocalAppData
    'Desktop
    'UserProfile
    'AllUsersProfile
    'TempCU
    'envCurUser
    'ProgramData
    'StartMenuPrograms
    
    'Special note:
    'Under Local System account some environment variables looks like this:
    '
    'APPDATA=C:\Windows\system32\config\systemprofile\AppData\Roaming
    'LOCALAPPDATA=C:\Windows\system32\config\systemprofile\AppData\Local
    'USERPROFILE=C:\Windows\system32\config\systemprofile
    '
    'also these internal variables will be modified:
    '
    'TempCU
    'UserProfile
    'Desktop
    'StartMenuPrograms
    '
    
    AppendErrorLogCustom "InitVariables - Begin"
    
    CRCinit
    
    Set oDictFileExist = New clsTrickHashTable  'file exists cache
    oDictFileExist.CompareMode = 1
    
    Dim lr As Long, i As Long, nChars As Long
    Dim path As String, dwBufSize As Long
    
    g_bIsReflectionSupported = IsProcedureAvail("RegQueryReflectionKey", "Advapi32.dll")
    
    sWinSysDir = Environ$("SystemRoot") & "\System32"
    
    If OSver.MajorMinor >= 5.1 And OSver.MajorMinor <= 5.2 Then bIsWinXP = True
    If OSver.MajorMinor = 5 Then bIsWin2k = True
    
    With OSver
        bIsWinVistaAndNewer = .Major >= 6
        bIsWin7AndNewer = .MajorMinor >= 6.1
        
        Select Case .PlatformID
            Case 0: bIsWin9x = True: bIsWinNT = False 'Win3x
            Case 1: bIsWin9x = True: bIsWinNT = False
            Case 2: bIsWinNT = True: bIsWin9x = False
        End Select
        
        If bIsWin9x Then
            If .Major = 4 Then
                If .Minor = 90 Then 'Windows Millennium Edition
                    bIsWinME = True
                End If
            End If
        End If
    End With
    
    SysDisk = String$(MAX_PATH, 0)
    lr = GetSystemWindowsDirectory(StrPtr(SysDisk), MAX_PATH)
    If lr Then
        sWinDir = Left$(SysDisk, lr)
        SysDisk = Left$(SysDisk, 2)
    Else
        sWinDir = EnvironW("%SystemRoot%")
        SysDisk = EnvironW("%SystemDrive%")
    End If
    sWinSysDir = sWinDir & "\" & IIf(bIsWinNT, "System32", "System")
    sSysDir = sWinSysDir
    sWinSysDirWow64 = sWinDir & "\SysWOW64"
    
    'enable redirector (just in case)
    If bIsWin64 Then ToggleWow64FSRedirection True
    
    If bIsWin64 And FolderExists(sWinDir & "\sysnative") And OSver.MajorMinor >= 6 Then
        sSysNativeDir = sWinDir & "\SysNative"
    Else
        sSysNativeDir = sWinDir & "\System32"
    End If
    
    If bIsWin64 Then
        If OSver.MajorMinor >= 6.1 Then     'Win 7 and later
            PF_64 = EnvironW("%ProgramW6432%")
        Else
            PF_64 = SysDisk & "\Program Files"
        End If
        PF_32 = EnvironW("%ProgramFiles%", True)
    Else
        PF_32 = EnvironW("%ProgramFiles%")
        PF_64 = PF_32
    End If
    
    PF_32_Common = PF_32 & "\Common Files"
    PF_64_Common = PF_64 & "\Common Files"
    
    UserProfile = GetSpecialFolderPath(CSIDL_PROFILE)
    If Len(UserProfile) = 0 Then UserProfile = EnvironW("%UserProfile%")
    
    If OSver.IsLocalSystemContext Then
        If OSver.IsWindowsVistaOrGreater Then
            path = SysDisk & "\Users"
        Else
            path = SysDisk & "\Documents and Settings"
        End If
    Else
        Call GetProfilesDirectory(StrPtr(path), dwBufSize)
        If dwBufSize > 0 Then
            path = String(dwBufSize, 0)
            dwBufSize = Len(path)
            
            If GetProfilesDirectory(StrPtr(path), dwBufSize) Then
                path = Left$(path, lstrlen(StrPtr(path)))
            Else
                path = vbNullString
            End If
        End If
    End If
    If Len(path) = 0 Then
        path = GetParentDir(UserProfile)
    End If
    
    ProfilesDir = path
    
    nChars = MAX_PATH
    AllUsersProfile = String$(nChars, 0)
    If GetAllUsersProfileDirectory(StrPtr(AllUsersProfile), nChars) Then
        AllUsersProfile = Left$(AllUsersProfile, nChars - 1)
    Else
        If Not OSver.IsWindowsVistaOrGreater Then
            If Len(ProfilesDir) = 0 Then
                ProfilesDir = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory")
            End If
            If Len(ProfilesDir) <> 0 Then
                AllUsersProfile = ProfilesDir & "\" & Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "AllUsersProfile")
            End If
        Else    'Win Vista +
            AllUsersProfile = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProgramData")
        End If
    End If
    If Len(AllUsersProfile) = 0 Then
        AllUsersProfile = EnvironW("%ALLUSERSPROFILE%")
    End If
    
    AppData = GetSpecialFolderPath(CSIDL_APPDATA)
    If Len(AppData) = 0 Then AppData = EnvironW("%AppData%")
    
    LocalAppData = GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)
    If Len(LocalAppData) = 0 Then
        If OSver.IsWindowsVistaOrGreater Then
            LocalAppData = EnvironW("%LocalAppData%")
        Else
            LocalAppData = UserProfile & "\Local Settings\Application Data"
        End If
    End If
    
    If OSver.MajorMinor < 6 Then
        AppDataLocalLow = AppData
    Else
        AppDataLocalLow = GetKnownFolderPath("{A520A1A4-1780-4FF6-BD18-167343C5AF16}")
    End If
    
    StartMenuPrograms = GetSpecialFolderPath(CSIDL_COMMON_PROGRAMS)
    
    Desktop = GetSpecialFolderPath(CSIDL_DESKTOP)
    
    'TempCU = Environ$("temp") ' will return path in format 8.3 on XP
    TempCU = Reg.GetData(HKEY_CURRENT_USER, "Environment", "Temp")
    ' if REG_EXPAND_SZ is missing
    If InStr(TempCU, "%") <> 0 Then
        TempCU = EnvironW(TempCU)
    End If
    If Len(TempCU) = 0 Or InStr(TempCU, "%") <> 0 Then ' if there TEMP is not defined
        If OSver.IsWindowsVistaOrGreater Then
            TempCU = UserProfile & "\Local\Temp"
        Else
            TempCU = UserProfile & "\Local Settings\Temp"
        End If
    End If
    
    envCurUser = OSver.UserName
    'envCurUser = EnvironW("%UserName%")
    
    ProgramData = EnvironW("%ProgramData%")
    
    'Override some special folders and substitute first found user if token = Local System
    
    If OSver.IsLocalSystemContext Then
    
        Dim ProfileListKey      As String
        Dim ProfileSubKey()     As String
        Dim sSid                As String
    
        ProfileListKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        'ProfilesDirectory = Reg.GetString(HKLM, ProfileListKey, "ProfilesDirectory")
    
        If Reg.EnumSubKeysToArray(HKLM, ProfileListKey, ProfileSubKey()) > 0 Then
            For i = LBound(ProfileSubKey) To UBound(ProfileSubKey)
            
                sSid = ProfileSubKey(i)
            
                If Not (sSid = "S-1-5-18" Or _
                        sSid = "S-1-5-19" Or _
                        sSid = "S-1-5-20") Then
                    
                    UserProfile = Reg.GetString(HKLM, ProfileListKey & "\" & sSid, "ProfileImagePath")
                    
                    'just in case
                    If Len(UserProfile) = 0 Then UserProfile = SysDisk & "\All Users"
                    
                    AppData = vbNullString
                    LocalAppData = vbNullString
                    Desktop = vbNullString
                    TempCU = vbNullString
                    
                    '���� ������� ��������
                    If Reg.KeyExists(HKEY_USERS, sSid) Then
                        
                        AppData = Reg.GetString(HKEY_USERS, sSid & _
                            "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AppData")
                    
                        LocalAppData = Reg.GetString(HKEY_USERS, sSid & _
                            "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Local AppData")
                    
                        Desktop = Reg.GetString(HKEY_USERS, sSid & _
                            "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Desktop")
                    
                        TempCU = Reg.GetString(HKEY_USERS, sSid & _
                            "\Environment", "TEMP")
                        
                        'HKU contains paths of REG_EXPAND_SZ type with %UserProfile% value, so they will be expanded automatically with wrong values,
                        'so we need to substitute correct values manually:
                        AppData = PathSubstituteProfile(AppData, UserProfile)
                        LocalAppData = PathSubstituteProfile(LocalAppData, UserProfile)
                        Desktop = PathSubstituteProfile(Desktop, UserProfile)
                        TempCU = PathSubstituteProfile(TempCU, UserProfile)
                        
                        If OSver.MajorMinor < 6 Then
                            AppDataLocalLow = AppData
                        Else
                            AppDataLocalLow = BuildPath(GetParentDir(AppData), "LocalLow")
                        End If
                    End If
    
                    '���� ������� �� ��������
                    
                    If OSver.IsWindowsVistaOrGreater Then
                        
                        If Len(AppData) = 0 Then AppData = UserProfile & "\AppData\Roaming"
                        If Len(AppDataLocalLow) = 0 Then AppDataLocalLow = BuildPath(GetParentDir(AppData), "LocalLow")
                        If Len(LocalAppData) = 0 Then LocalAppData = UserProfile & "\AppData\Local"
                        If Len(Desktop) = 0 Then Desktop = UserProfile & "\Desktop"
                        If Len(TempCU) = 0 Then TempCU = LocalAppData & "\Temp"
                        
                    Else
                        If Len(AppData) = 0 Then AppData = UserProfile & "\Application Data"
                        If Len(AppDataLocalLow) = 0 Then AppDataLocalLow = AppData
                        If Len(LocalAppData) = 0 Then LocalAppData = UserProfile & "\Local Settings"
                        If Len(TempCU) = 0 Then TempCU = LocalAppData & "\Temp"
                        
                        If Len(Desktop) = 0 Then
                            path = Reg.GetString(HKLM, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common Desktop")
                            
                            If Len(path) = 0 Then
                                If IsSlavianCultureCode(OSver.LangSystemCode) Then
                                    Desktop = UserProfile & "\" & LoadResString(606) '������� ����
                                Else
                                    Desktop = UserProfile & "\Desktop"
                                End If
                            Else
                                path = GetFileNameAndExt(path)
                                Desktop = UserProfile & "\" & path
                            End If
                        End If
                    End If
                    
                    Exit For
                
                End If
            Next
        End If
    End If
    
    ' Shortcut interfaces initialization
    'IURL_Init
    ISL_Init
    
    Set oDict.TaskWL_ID = New clsTrickHashTable
    oDict.TaskWL_ID.CompareMode = vbTextCompare
    
    Set colProfiles = New Collection
    Set colProfilesUser = New Collection
    GetProfiles
    
    FillUsers
    
    g_LocalGroupNames = GetLocalGroupNames()
    g_LocalUserNames = GetLocalUserNames()
    
    Set cMath = New clsMath
    'Set oRegexp = New cRegExp
    
    If OSver.MajorMinor >= 6.1 Then
        Set TaskBar = New TaskbarList
    End If
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    'Call CLSIDFromString(StrPtr(FOLDERID_ComputerFolderStr), FOLDERID_ComputerFolder)
    
    'Load ru phrases
    STR_CONST.RU_NO = LoadResString(601)
    STR_CONST.RU_MICROSOFT = LoadResString(604)
    STR_CONST.RU_PC = LoadResString(605)
    STR_CONST.SHA1_PCRE2 = LoadResString(700)
    STR_CONST.SHA1_ABR = LoadResString(701)
    STR_CONST.SHA1_OCX = LoadResString(702)
    STR_CONST.WINDOWS_DEFENDER = Caes_Decode("XlskxHF UxABMEHW") 'Windows Defender
    STR_CONST.VIRUSTOTAL = Caes_Decode("WlwBB_BIrE") 'VirusTotal
    STR_CONST.AUTORUNS = Caes_Decode("BxyvAFAH") 'Autoruns
    
    AppendErrorLogCustom "InitVariables - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "InitVariables"
    If inIDE Then Stop: Resume Next
End Sub

Public Function PathSubstituteProfile(path As String, Optional ByVal sUserProfileDir As String) As String
    'Substitute 'sUserProfileDir' to 'Path' if 'Path' goes through %UserProfile%.s
    'Note: sUserProfileDir can be not a profile at all. In such case substitution is not performed.
    
    Dim pos As Long
    Dim bComply As Boolean
    
    'expanded path contains current profile's dir?
    bComply = False
    
    If StrBeginWith(path, UserProfile & "\") Then bComply = True
    If Not bComply Then If StrComp(path, UserProfile, 1) = 0 Then bComply = True
    
    If bComply And StrComp(ProfilesDir, GetParentDir(sUserProfileDir), 1) = 0 Then
        'substitute
        PathSubstituteProfile = BuildPath(sUserProfileDir, mid$(path, Len(UserProfile) + 2))
        Exit Function
    End If
    
    PathSubstituteProfile = path
End Function

Public Function EnvironW(ByVal SrcEnv As String, Optional UseRedir As Boolean, Optional ByVal sUserProfileDir As String) As String
    Dim lr As Long
    Dim buf As String
    Static LastFile As String
    Static LastResult As String
    
    AppendErrorLogCustom "EnvironW - Begin", "SrcEnv: " & SrcEnv
    
    If Len(SrcEnv) = 0 Then Exit Function
    If InStr(SrcEnv, "%") = 0 Then
        EnvironW = SrcEnv
    Else
        If LastFile = SrcEnv And Len(sUserProfileDir) = 0 And UseRedir = False Then
            EnvironW = LastResult
            Exit Function
        End If
        'redirector correction
        If OSver.IsWow64 Then
            If Not UseRedir Then
                If InStr(1, SrcEnv, "%PROGRAMFILES%", 1) <> 0 Then
                    SrcEnv = Replace$(SrcEnv, "%PROGRAMFILES%", PF_64, 1, 1, 1)
                End If
                If InStr(1, SrcEnv, "%COMMONPROGRAMFILES%", 1) <> 0 Then
                    SrcEnv = Replace$(SrcEnv, "%COMMONPROGRAMFILES%", PF_64_Common, 1, 1, 1)
                End If
            End If
        End If
        buf = String$(MAX_PATH, vbNullChar)
        lr = ExpandEnvironmentStrings(StrPtr(SrcEnv), StrPtr(buf), MAX_PATH + 1)
        
        If lr > MAX_PATH Then
            buf = String$(lr, vbNullChar)
            lr = ExpandEnvironmentStrings(StrPtr(SrcEnv), StrPtr(buf), lr + 1)
        End If
        
        If lr Then
            EnvironW = Left$(buf, lr - 1)
        Else
            EnvironW = SrcEnv
        End If
        
        If InStr(EnvironW, "%") <> 0 Then
            If OSver.MajorMinor <= 6 Then
                If InStr(1, EnvironW, "%ProgramW6432%", 1) <> 0 Then
                    EnvironW = Replace$(EnvironW, "%ProgramW6432%", SysDisk & "\Program Files", 1, -1, 1)
                End If
            End If
        End If
        
        'if need expanding under certain user
        If Len(sUserProfileDir) <> 0 Then
            EnvironW = PathSubstituteProfile(EnvironW, sUserProfileDir)
        End If
    End If
    
    If Len(sUserProfileDir) = 0 And UseRedir = False Then
        LastFile = SrcEnv
        LastResult = EnvironW
    End If
    
    AppendErrorLogCustom "EnvironW - End", EnvironW
End Function

Public Function GetProfileDirBySID(sSid As String) As String
    
    GetProfileDirBySID = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & sSid, "ProfileImagePath")
    
    If Len(GetProfileDirBySID) = 0 Then
        
        If sSid = SID_TEMPLATE Then Exit Function
        
        Dim sUsername As String
        sUsername = MapSIDToUsername(sSid)
        
        If Len(sUsername) <> 0 Then
            GetProfileDirBySID = BuildPath(ProfilesDir, sUsername)
        End If
    End If

End Function

Public Function StrInParamArray(stri As String, ParamArray vEtalon()) As Boolean
    Dim i As Long
    For i = 0 To UBound(vEtalon)
        If StrComp(stri, vEtalon(i), 1) = 0 Then StrInParamArray = True: Exit For
    Next
End Function

' ���������� true, ���� ������� �������� ������� � ����� �� ��������� ������� (lB, uB ������������ ��������������� �������� ��������)
Public Function InArray( _
    stri As String, _
    MyArray() As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    If AryItems(MyArray) = 0 Then Exit Function
    If lB = -2147483647 Then lB = LBound(MyArray)   'some trick
    If uB = 2147483647 Then uB = UBound(MyArray)    'Thanks to ��������� :)
    Dim i As Long
    For i = lB To uB
        If StrComp(stri, MyArray(i), CompareMethod) = 0 Then InArray = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "inArray"
    If inIDE Then Stop: Resume Next
End Function

'Note: Serialized array - it is a string which stores all items of array delimited by some character (default delimiter in HJT is '|' and '*' chars)
'Example 1: "string1*string2*string3"
'Example 2: "string1|string2|string3" and so.

'this function returns true, if any of items in serialized array has exact match with 'Stri' variable
'you can restrict search with LBound and UBound items only.
Public Function inArraySerialized( _
    stri As String, _
    SerializedArray As String, _
    Delimiter As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    Dim MyArray() As String
    If 0 = Len(SerializedArray) Then
        If 0 = Len(stri) Then inArraySerialized = True
        Exit Function
    End If
    MyArray = Split(SerializedArray, Delimiter)
    If lB = -2147483647 Or lB < LBound(MyArray) Then lB = LBound(MyArray)  'some trick
    If uB = 2147483647 Or uB > UBound(MyArray) Then uB = UBound(MyArray)  'Thanks to ��������� :)
    
    Dim i As Long
    For i = lB To uB
        If StrComp(stri, MyArray(i), CompareMethod) = 0 Then inArraySerialized = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "inArraySerialized", "SerializedString: ", SerializedArray, "delim: ", Delimiter
    If inIDE Then Stop: Resume Next
End Function

'The same as Split(), except of proper error handling when source data is empty string and you assign result to variable defined as array.
'So, in case of empty string it return array with 1 empty item (0 to 0), unless bAllowReturnEmptyArray=true is specified.
'If that's the case, bounds(0 to -1) are returned.
'Also: return type is 'string()' instead of 'variant()'
'
'Warning note:
'To use this function in For each loop, specify bAllowReturnEmptyArray = true
Public Function SplitSafe(sComplexString As String, Optional Delimiter As String = " ", _
    Optional bAllowReturnEmptyArray As Boolean = False) As String()
    
    If 0 = Len(sComplexString) Then
        If bAllowReturnEmptyArray Then
            SplitSafe = EmptyArray()
        Else
            ReDim SplitSafe(0)
        End If
    Else
        SplitSafe = Split(sComplexString, Delimiter)
    End If
End Function

Public Sub ArrayRemoveEmptyItems(arr() As String)
    Dim i As Long
    Dim d As Long
    Dim bShift As Boolean
    
    If IsArrayEmpty(arr) Then Exit Sub
    
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) <> 0 Then
            If bShift Then arr(d) = arr(i): d = d + 1 'shifting items
        Else
            If Not bShift Then bShift = True: d = i
        End If
    Next
    
    If bShift Then
        If d > LBound(arr) Then
            ReDim Preserve arr(d - 1)
        ElseIf UBound(arr) > LBound(arr) Then
            ReDim Preserve arr(d)
        End If
    End If
End Sub

'get the first item of serilized array
Public Function SplitExGetFirst(sSerializedArray As String, Optional Delimiter As String = " ") As String
    SplitExGetFirst = SplitSafe(sSerializedArray, Delimiter)(0)
End Function

'get the last item of serialized array
Public Function SplitExGetLast(sSerializedArray As String, Optional Delimiter As String = " ") As String
    Dim ret() As String
    ret = SplitSafe(sSerializedArray, Delimiter)
    SplitExGetLast = ret(UBound(ret))
End Function

Public Function IsArrayEmpty(arr() As String) As Boolean
    IsArrayEmpty = (UBound(arr) < LBound(arr))
End Function

Private Sub DeleteDuplicatesInArray(arr() As String, CompareMethod As VbCompareMethod, Optional DontCompress As Boolean)
    On Error GoTo ErrorHandler:
    
    'DontCompress:
    'if true, do not move items:
    'function will return array with empty items in places where duplicate match were found
    'so, its structure will be similar to the source array
    
    'if false, returns new reconstructed array:
    'all subsequent array items are shifted to the item where duplicate was found.
    
    Dim i As Long
    If IsArrayEmpty(arr) Then Exit Sub
    
    If DontCompress Then
        For i = UBound(arr) To LBound(arr) + 1 Step -1
            If Len(arr(i)) <> 0 Then
                If InArray(arr(i), arr, LBound(arr), i - 1, CompareMethod) Then
                    arr(i) = vbNullString
                End If
            End If
        Next
    Else
        Dim TmpArr() As String
        ReDim TmpArr(LBound(arr) To UBound(arr))
        Dim cnt As Long
        cnt = LBound(arr)
        For i = LBound(arr) To UBound(arr)
            If Not InArray(arr(i), arr, i + 1, UBound(arr), CompareMethod) Then
                TmpArr(cnt) = arr(i)
                cnt = cnt + 1
            End If
        Next
        ReDim Preserve TmpArr(LBound(TmpArr) To cnt - 1)
        arr = TmpArr
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "DeleteDuplicatesInArray"
    If inIDE Then Stop: Resume Next
End Sub

'Remove empty strings from array
Public Sub CompressArray(arr() As String)
    On Error GoTo ErrorHandler:

    If 0 = AryPtr(arr) Then Exit Sub
    Dim i As Long
    Dim pIdx As Long
    pIdx = -1
    For i = 0 To UBound(arr)
        If Len(arr(i)) = 0 Then
            If pIdx = -1 Then
                pIdx = i
            End If
        Else
            If pIdx <> -1 Then
                arr(pIdx) = arr(i)
                pIdx = pIdx + 1
            End If
        End If
    Next
    If pIdx > 0 Then
        ReDim Preserve arr(pIdx - 1)
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CompressArray"
    If inIDE Then Stop: Resume Next
End Sub

Public Function StrBeginWith(Text As String, BeginPart As String) As Boolean
    StrBeginWith = (StrComp(Left$(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function

Public Function StrEndWith(Text As String, lastPart As String) As Boolean
    StrEndWith = (StrComp(Right$(Text, Len(lastPart)), lastPart, 1) = 0)
End Function

Public Function StrEndWithParamArray(Text As String, ParamArray vLastPart()) As Boolean
    Dim i As Long
    For i = 0 To UBound(vLastPart)
        If Len(vLastPart(i)) <> 0 Then
            If StrComp(Right$(Text, Len(vLastPart(i))), vLastPart(i), 1) = 0 Then
                StrEndWithParamArray = True
                Exit For
            End If
        End If
    Next
End Function

Public Function StrBeginWithArray(Text As String, BeginPart() As String) As Boolean
    Dim i As Long
    For i = 0 To UBound(BeginPart)
        If Len(BeginPart(i)) <> 0 Then
            If StrComp(Left$(Text, Len(BeginPart(i))), BeginPart(i), 1) = 0 Then
                StrBeginWithArray = True
                Exit For
            End If
        End If
    Next
End Function

Public Sub CenterForm(myForm As Form) ' ������������� ����� �� ������ � ������ ��������� �������
    On Error Resume Next
    Dim Left    As Long
    Dim Top     As Long
    Left = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN) / 2 - myForm.Width / 2
    Top = Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYFULLSCREEN) / 2 - myForm.Height / 2
    myForm.Move Left, Top
End Sub

Public Function LoadWindowPos(frm As Form, idSection As SETTINGS_SECTION) As Boolean
    
    If frm.WindowState = vbMinimized Or frm.WindowState = vbMaximized Then Exit Function
    
    LoadWindowPos = True
    
    If idSection <> SETTINGS_SECTION_MAIN Then
    
        Dim iHeight As Long, iWidth As Long
        iHeight = CLng(RegReadHJT("WinHeight", "-1", idSection))
        iWidth = CLng(RegReadHJT("WinWidth", "-1", idSection))
        
        If iHeight = -1 Or iWidth = -1 Then LoadWindowPos = False
        
        If iHeight > 0 And iWidth > 0 Then
            If iHeight > Screen.Height Then iHeight = Screen.Height
            If iWidth > Screen.Width Then iWidth = Screen.Width
            
            If iHeight < 500 Then iHeight = 500
            If iWidth < 1000 Then iWidth = 1000
            
            frm.Height = iHeight
            frm.Width = iWidth
        End If
    End If
    
    Dim iTop As Long, iLeft As Long
    iTop = CLng(RegReadHJT("WinTop", "-1", idSection))
    iLeft = CLng(RegReadHJT("WinLeft", "-1", idSection))
    
    If iTop = -1 Or iLeft = -1 Then
    
        LoadWindowPos = False
        CenterForm frm
    Else
        If iTop > (Screen.Height - 2500) Then iTop = Screen.Height - 2500
        If iLeft > (Screen.Width - 5000) Then iLeft = Screen.Width - 5000
        If iTop < 0 Then iTop = 0
        If iLeft < 0 Then iLeft = 0
        
        frm.Top = iTop
        frm.Left = iLeft
    End If
    
    If CLng(RegReadHJT("WinState", "0", idSection)) = vbMaximized Then frm.WindowState = vbMaximized
End Function

Public Sub SaveWindowPos(frm As Form, idSection As SETTINGS_SECTION)
    
    If g_UninstallState Then Exit Sub
    
    If frm.WindowState <> vbMinimized And frm.WindowState <> vbMaximized Then
        RegSaveHJT "WinTop", CStr(frm.Top), idSection
        RegSaveHJT "WinLeft", CStr(frm.Left), idSection
        RegSaveHJT "WinHeight", CStr(frm.Height), idSection
        RegSaveHJT "WinWidth", CStr(frm.Width), idSection
    End If
    RegSaveHJT "WinState", CStr(frm.WindowState), idSection
    
End Sub

Public Function ConvertVersionToNumber(sVersion As String) As Long  '"1.1.1.1" -> 1 number
    On Error GoTo ErrorHandler:
    Dim Ver() As String
    
    If 0 = Len(sVersion) Then Exit Function
    
    Ver = SplitSafe(sVersion, ".")
    If UBound(Ver) = 3 Then
        ConvertVersionToNumber = cMath.Shl(Val(Ver(0)), 24) + cMath.Shl(Val(Ver(1)), 16) + cMath.Shl(Val(Ver(2)), 8) + Val(Ver(3))
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ConvertVersionToNumber"
    If inIDE Then Stop: Resume Next
End Function

Public Sub UpdatePolicy(Optional noWait As Boolean)

    If OSver.IsWindows8OrGreater Then Exit Sub

    Dim GPUpdatePath$
    If bIsWin64 And FolderExists(sWinDir & "\sysnative") And OSver.MajorMinor >= 6 Then
        GPUpdatePath = sWinDir & "\sysnative\gpupdate.exe"
    Else
        GPUpdatePath = sWinDir & "\system32\gpupdate.exe"
    End If
    If Proc.ProcessRun(GPUpdatePath, "/force", , vbHide) Then
        If Not noWait Then
            Proc.WaitForTerminate , , , 15000
        End If
    End If
End Sub

Public Sub ConcatArrays(DestArray() As String, AddArray() As String)
    'Appends AddArray() to the end of DestArray.
    'DestArray() should be declared as dynamic
    
    'UnInitialized arrays are permitted
    'Warning: if both arrays is uninitialized - DestArray() will remain the same (with uninitialized state)
    
    On Error GoTo ErrorHandler
    
    Dim i&, idx&
    
    If 0 = AryItems(AddArray) Then Exit Sub
    If 0 = AryItems(DestArray) Then
        idx = -1
        ReDim DestArray(UBound(AddArray) - LBound(AddArray))
    Else
        idx = UBound(DestArray)
        ReDim Preserve DestArray(UBound(DestArray) + (UBound(AddArray) - LBound(AddArray)) + 1)
    End If
    
    For i = LBound(AddArray) To UBound(AddArray)
        idx = idx + 1
        DestArray(idx) = AddArray(i)
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Parser.ConcatArrays"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub QuickSort(j() As String, low As Long, high As Long, Optional CompareMethod As VbCompareMethod = vbTextCompare)
    On Error GoTo ErrorHandler
    Dim i As Long, L As Long, pM As Long, wsp As Long, pSA As Long, lcid As Long
    i = low: L = high: pM = StrPtr(j((i + L) \ 2))
    pSA = Deref(AryPtr(j) + 12) 'SAFEARRAY.pvData => BSTR
    lcid = OSver.LCID_UserDefault
    Do Until i > L
        Do While VarBstrCmp(StrPtr(j(i)), pM, lcid, CompareMethod) = VARCMP_LT: i = i + 1: Loop
        Do While VarBstrCmp(StrPtr(j(L)), pM, lcid, CompareMethod) = VARCMP_GT: L = L - 1: Loop
        If (i <= L) Then
            wsp = StrPtr(j(L))
            PutMem4 ByVal (pSA + 4 * L), StrPtr(j(i))
            PutMem4 ByVal (pSA + 4 * i), wsp
            i = i + 1: L = L - 1
        End If
    Loop
    If low < L Then QuickSort j, low, L, CompareMethod
    If i < high Then QuickSort j, i, high, CompareMethod
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "QuickSort"
    If inIDE Then Stop: Resume Next
End Sub

' For Variant/String() only!
'
Public Sub QuickSortV(j() As Variant, low As Long, high As Long, Optional CompareMethod As VbCompareMethod = vbTextCompare)
    On Error GoTo ErrorHandler
    Dim i As Long, L As Long, pM As Long, wsp As Long, pVA As Long, lcid As Long
    i = low: L = high: pM = StrPtr(j((i + L) \ 2))
    pVA = Deref(AryPtr(j) + 12) 'SAFEARRAY.pvData => VARIANT
    lcid = OSver.LCID_UserDefault
    Do Until i > L
        Do While VarBstrCmp(StrPtr(j(i)), pM, lcid, CompareMethod) = VARCMP_LT: i = i + 1: Loop
        Do While VarBstrCmp(StrPtr(j(L)), pM, lcid, CompareMethod) = VARCMP_GT: L = L - 1: Loop
        If (i <= L) Then
            wsp = StrPtr(j(L))
            PutMem4 ByVal (pVA + 16 * L + 8), StrPtr(j(i))
            PutMem4 ByVal (pVA + 16 * i + 8), wsp
            i = i + 1: L = L - 1
        End If
    Loop
    If low < L Then QuickSortV j, low, L, CompareMethod
    If i < high Then QuickSortV j, i, high, CompareMethod
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "QuickSortV"
    If inIDE Then Stop: Resume Next
End Sub

' exclude items from ArraySrc() that is not match 'Mask' and save to 'ArrayDest()'
' return value is a number of items in 'ArrayDest'
' if number of items is 0, ArrayDest() will have 1 empty item.
Public Function FilterArray(ArraySrc() As String, ArrayDest() As String, Mask As String) As Long
    On Error GoTo ErrorHandler:
    Dim i As Long, j As Long
    ReDim ArrayDest(LBound(ArraySrc) To UBound(ArraySrc))
    For i = LBound(ArraySrc) To UBound(ArraySrc)
        If ArraySrc(i) Like Mask Then
            j = j + 1
            ArrayDest(LBound(ArraySrc) + j - 1) = ArraySrc(i)
        End If
    Next
    If j = 0 Then
        ReDim ArrayDest(LBound(ArraySrc) To LBound(ArraySrc))
    Else
        ReDim Preserve ArrayDest(LBound(ArraySrc) To LBound(ArraySrc) + j - 1)
    End If
    FilterArray = j
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FilterArray"
    If inIDE Then Stop: Resume Next
End Function

'get a substring starting at the specified character (search begins with the end of the line)
Public Function MidFromCharRev(sText As String, Delimiter As String) As String
    On Error GoTo ErrorHandler:
    Dim iPos As Long
    If 0 <> Len(sText) Then
        iPos = InStrRev(sText, Delimiter)
        If iPos <> 0 Then
            MidFromCharRev = mid$(sText, iPos + 1)
        Else
            MidFromCharRev = vbNullString
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MidFromCharRev"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetCollectionKeyByIndex(ByVal Index As Long, col As Collection) As String ' Thanks to 'The Trick' (�. �������) for this code
    'Fixed by Dragokas
    On Error GoTo ErrorHandler:
    Dim lpSTR As Long, ptr As Long, Key As String
    If col Is Nothing Then Exit Function
    Select Case Index
    Case Is < 1, Is > col.Count: Exit Function
    Case Else
        ptr = ObjPtr(col)
        Do While Index
            GetMem4 ByVal ptr + 24, ptr
            Index = Index - 1
        Loop
    End Select
    GetMem4 ByVal VarPtr(Key), lpSTR
    GetMem4 ByVal ptr + 16, ByVal VarPtr(Key)
    GetCollectionKeyByIndex = Key
    GetMem4 lpSTR, ByVal VarPtr(Key)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetCollectionKeyByIndex"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetCollectionIndexByItem(sItem As String, col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(col.Item(i), sItem, CompareMode) = 0 Then
            GetCollectionIndexByItem = i
            Exit For
        End If
    Next
End Function

Public Function GetCollectionKeyByItem(sItem As String, col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As String
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(col.Item(i), sItem, CompareMode) = 0 Then
            GetCollectionKeyByItem = GetCollectionKeyByIndex(i, col)
            Exit For
        End If
    Next
End Function

Public Function isCollectionKeyExists(Key As String, col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(GetCollectionKeyByIndex(i, col), Key, CompareMode) = 0 Then isCollectionKeyExists = True: Exit For
    Next
End Function

Public Function isCollectionItemExists(vItem As Variant, col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
    isCollectionItemExists = (GetCollectionIndexByItem(CStr(vItem), col, CompareMode) <> 0)
End Function

Public Function GetCollectionItemByKey(Key As String, col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As String
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(GetCollectionKeyByIndex(i, col), Key, CompareMode) = 0 Then GetCollectionItemByKey = col.Item(i)
    Next
End Function

Public Function GetCollectionIndexByKey(Key As String, col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(GetCollectionKeyByIndex(i, col), Key, CompareMode) = 0 Then GetCollectionIndexByKey = i
    Next
End Function

Public Sub GetProfiles()    'result -> in global variable 'colProfiles' (collection)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetProfiles - Begin"
    
    'include all folders inside <c:\users>
    'without 'Public'
    
    Dim ProfileListKey      As String
    Dim ProfilesDirectory   As String
    Dim ProfileSubKey()     As String
    Dim ProfilePath         As String
    Dim SubFolders()        As String
    Dim i                   As Long
    Dim lr                  As Long
    Dim path                As String
    Dim objFolder           As Variant
    Dim sSid                As String
    
    ProfileListKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    ProfilesDirectory = Reg.GetData(0&, ProfileListKey, "ProfilesDirectory")

    
    If Reg.EnumSubKeysToArray(0&, ProfileListKey, ProfileSubKey()) > 0 Then
        For i = 1 To UBound(ProfileSubKey)
            sSid = ProfileSubKey(i)
            If Not (sSid = "S-1-5-18" Or _
                    sSid = "S-1-5-19" Or _
                    sSid = "S-1-5-20") Then
                
                ProfilePath = Reg.GetData(0&, ProfileListKey & "\" & sSid, "ProfileImagePath")
                
                If Len(ProfilePath) <> 0 Then
                    If FolderExists(ProfilePath) Then
                        If Not isCollectionKeyExists(ProfilePath, colProfiles) Then
                            On Error Resume Next
                            colProfiles.Add ProfilePath, ProfilePath
                            colProfilesUser.Add MapSIDToUsername(sSid)
                            On Error GoTo ErrorHandler:
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    '�������� �����, ������� ��������� � ����������� (�� 1 ������� ����) ������� �������� ������������
    
    If Len(UserProfile) <> 0 Then
        If FolderExists(UserProfile) Then
            path = UserProfile
            lr = PathRemoveFileSpec(StrPtr(path))   ' get Parent directory
            If lr Then path = Left$(path, lstrlen(StrPtr(path)))

            SubFolders() = ListSubfolders(path)

            If AryItems(SubFolders) Then
                For Each objFolder In SubFolders()
                    If Len(objFolder) <> 0 And Not (StrEndWith(CStr(objFolder), "\Public") And OSver.MajorMinor >= 6) Then
                        If FolderExists(CStr(objFolder)) Then
                            If Not isCollectionKeyExists(CStr(objFolder), colProfiles) Then
                                On Error Resume Next
                                colProfiles.Add CStr(objFolder), CStr(objFolder)
                                colProfilesUser.Add "Unknown"
                                On Error GoTo ErrorHandler:
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    AppendErrorLogCustom "GetProfiles - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetProfiles"
    If inIDE Then Stop: Resume Next
End Sub

Public Function UnpackCryptedFile(ResourceID As Long, DestinationPath As String) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "UnpackCryptedFile - Begin", "ID: " & ResourceID, "Destination: " & DestinationPath
    
    Dim b()     As Byte
    Dim hFile   As Long
    Dim lBytesWrote As Long
    
    b = LoadResData(ResourceID, "CUSTOM")
    Caes_DecodeBin b
    
    hFile = CreateFile(StrPtr(DestinationPath), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, ByVal 0&)
    
    If hFile <> INVALID_HANDLE_VALUE Then
        If WriteFile(hFile, VarPtr(b(0)), UBound(b) + 1, lBytesWrote, 0&) Then UnpackCryptedFile = True
        CloseHandle hFile
    End If
    AppendErrorLogCustom "UnpackCryptedFile - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "UnpackCryptedFile", "ID: " & ResourceID, "Destination path: " & DestinationPath
    UnpackCryptedFile = False
    If inIDE Then Stop: Resume Next
End Function

Public Sub Terminate_HJT()
    Unload frmMain
    End
End Sub

Public Sub AddHorizontalScrollBarToResults(lstControl As VBCCR17.ListBoxW)
    'Adds a horizontal scrollbar to the results display if it is needed (after the scan)
    Dim x As Long, s$
    Dim idx As Long
    
    With lstControl
        For idx = 0 To .ListCount - 1
            s = Replace$(.List(idx), vbTab, "12345678")
            If .Width < frmMain.TextWidth(s) + 300 And x < frmMain.TextWidth(s) + 300 Then
                x = frmMain.TextWidth(.List(idx)) + 300
            End If
        Next
        If x <> 0 Then
            x = x * 1.2
            If frmMain.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX + 50  ' if twips change to pixels (+50 to account for the width of the vertical scrollbar
        End If
        SendMessage .hWnd, LB_SETHORIZONTALEXTENT, x, ByVal 0&
    End With
End Sub

'Public Function IsArrDimmed(vArray As Variant) As Boolean
'    IsArrDimmed = (GetArrDims(vArray) > 0)
'End Function

Public Function AryItems(vArray As Variant) As Long
    Dim ppSA As Long
    Dim pSA As Long
    Dim VT As Long
    'Dim sa As SAFEARRAY
    Dim pvData As Long
    Const vbByRef As Integer = 16384

    If IsArray(vArray) Then
        GetMem4 ByVal VarPtr(vArray) + 8, ppSA      ' pV -> ppSA (pSA)
        If ppSA <> 0 Then
            GetMem2 vArray, VT
            If VT And vbByRef Then
                GetMem4 ByVal ppSA, pSA                 ' ppSA -> pSA
            Else
                pSA = ppSA
            End If
            If pSA <> 0 Then
                'memcpy sa, ByVal pSA, LenB(sa)
                'If sa.pvData <> 0 Then
                GetMem4 ByVal pSA + 12, pvData
                If pvData <> 0 Then
                    AryItems = UBound(vArray) - LBound(vArray) + 1
                End If
            End If
        End If
    End If
End Function

Public Function LBoundSafe(vArray As Variant) As Long
    If AryItems(vArray) Then
        LBoundSafe = LBound(vArray)
    Else
        LBoundSafe = 2147483647
    End If
End Function

Public Function UBoundSafe(vArray As Variant) As Long
    If AryItems(vArray) Then
        UBoundSafe = UBound(vArray)
    Else
        UBoundSafe = -1
    End If
End Function

' ������������� HTTP: -> HXXP:, HTTPS: -> HXXPS:, WWW -> VVV
Public Function doSafeURLPrefix(sURL As String) As String
    doSafeURLPrefix = Replace(Replace(Replace(sURL, "http:", "hxxp:", , , 1&), "www", "vvv", , , 1&), "https:", "hxxps:", , , 1&)
End Function

Public Sub Dbg(sMsg As String)
    If bDebugMode Then
        AppendErrorLogCustom sMsg
        'OutputDebugStringA sMsg ' -> because already is in AppendErrorLogCustom sub()
    End If
End Sub

Public Sub AppendErrorLogCustom(ParamArray CodeModule())    'trace info
    
    If Not (bDebugMode Or bDebugToFile) Then Exit Sub
    Static freq As Currency
    Static isInit As Boolean
    
    Dim Other       As String
    Dim i           As Long
    For i = 0 To UBound(CodeModule)
        Other = Other & CodeModule(i) & " | "
    Next
    
    Dim tim1 As Currency
    If Not isInit Then
        isInit = True
        QueryPerformanceFrequency freq
    End If
    QueryPerformanceCounter tim1
    
    If bDebugToFile Then
        If g_hDebugLog <> 0 Then
            Dim b() As Byte
            b = "- " & Time & " - " & Format$(tim1 / freq, "##0.000") & " - " & Other & vbCrLf
            PutW_NoLog g_hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
        End If
    End If
    
    If bDebugMode Then
    
        OutputDebugStringA Other

        If Not (ErrLogCustomText Is Nothing) Then
            ErrLogCustomText.Append (vbCrLf & "- " & Time & " - " & Format$(tim1 / freq, "##0.000") & " - " & Other)
        End If
    
        'If DebugHeavy Then AddtoLog vbCrLf & "- " & time & " - " & Other
    End If
End Sub

Public Sub OpenDebugLogHandle()
    If g_hDebugLog > 0 Then Exit Sub
    
    If Len(g_sDebugLogFile) = 0 Then
        g_sDebugLogFile = BuildPath(AppPath(), "HiJackThis_debug.log")
    End If
    
    If FileExists(g_sDebugLogFile) Then DeleteFileEx g_sDebugLogFile
    
    On Error Resume Next
    OpenW g_sDebugLogFile, FOR_OVERWRITE_CREATE, g_hDebugLog, g_FileBackupFlag
    
    If g_hDebugLog <= 0 Then
        g_sDebugLogFile = Left$(g_sDebugLogFile, Len(g_sDebugLogFile) - 4) & "_2.log"
                    
        Call OpenW(g_sDebugLogFile, FOR_OVERWRITE_CREATE, g_hDebugLog)
        
    End If
    
    Dim sCurTime$
    sCurTime = vbCrLf & vbCrLf & "Logging started at: " & Now() & vbCrLf & vbCrLf
    PutW_NoLog g_hDebugLog, 1&, StrPtr(sCurTime), LenB(sCurTime), doAppend:=True
End Sub

Public Sub OpenLogHandle()
    
    Dim ov As OVERLAPPED
    
    If Len(g_sLogFile) = 0 Then
        g_sLogFile = BuildPath(AppPath(), "HiJackThis_.log")
    End If
    
    If FileExists(g_sLogFile, , True) Then DeleteFileEx g_sLogFile
    
    On Error Resume Next
    OpenW g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog, g_FileBackupFlag
    
    If g_hLog > 0 Then
        ov.offset = 0
        ov.InternalHigh = 0
        ov.hEvent = 0
        
        Dim lret As Long
        
        lret = LockFileEx(g_hLog, LOCKFILE_EXCLUSIVE_LOCK Or LOCKFILE_FAIL_IMMEDIATELY, 0&, 1& * 1024 * 1024, 0&, VarPtr(ov))
        
        If lret Then
            g_LogLocked = True
        Else
            AppendErrorLogCustom "Can't lock file. Err = " & Err.LastDllError
        End If
    Else
        g_sLogFile = Left$(g_sLogFile, Len(g_sLogFile) - 4) & "_2.log"
        
        Call OpenW(g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog)
    End If
End Sub

Public Function StringFromPtrA(ByVal ptr As Long) As String
    If 0& <> ptr Then
        StringFromPtrA = SysAllocStringByteLen(ptr, lstrlenA(ptr))
    End If
End Function

Public Function StringFromPtrW(ByVal ptr As Long) As String
    Dim strSize As Long
    If 0 <> ptr Then
        strSize = lstrlen(ptr)
        If 0 <> strSize Then
            StringFromPtrW = String$(strSize, 0&)
            lstrcpyn StrPtr(StringFromPtrW), ptr, strSize + 1&
        End If
    End If
End Function

Public Sub DoCrash()
    memcpy 0, ByVal 0, 4
End Sub

Public Sub ParseKeysURL(ByVal sURL As String, aKey() As String, aVal() As String)
    On Error GoTo ErrorHandler:
    'Example:
    'http://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IE8SRC
    ' =>
    ' 1) Key(0) = q,    Val(0) = {searchTerms}
    ' 2) Key(1) = src,  Val(1) = IE-SearchBox
    ' 3) Key(2) = FORM, Val(2) = IE8SRC
    
    Dim pos As Long, aKeyPara() As String, aTmp() As String, i As Long
    
    Erase aKey
    Erase aVal
    
    pos = InStr(sURL, "?")
    If pos = 0 Or pos = Len(sURL) Then Exit Sub 'no '?' or last '?'
    sURL = mid$(sURL, pos + 1)
    
    aKeyPara = Split(sURL, "&")
    
    ReDim aKey(UBound(aKeyPara))
    ReDim aVal(UBound(aKeyPara))
    
    For i = 0 To UBound(aKeyPara)
        aTmp = Split(aKeyPara(i), "=", 2) 'split keypara
        If UBound(aTmp) >= 0 Then 'not empty key ?
            aKey(i) = aTmp(0)
            If StrBeginWith(aKey(i), "amp;") Then 'remove some strange amp; in key that happen sometimes
                If aKey(i) <> "amp;q" And aKey(i) <> "amp;query" Then 'restrict this rule for 'q' and 'query' important keys
                    aKey(i) = mid$(aKey(i), 5)
                End If
            End If
        End If
        If UBound(aTmp) > 0 Then 'not empty value ?
            aVal(i) = aTmp(1)
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ParseKeysURL", "sURL: " & sURL
    If inIDE Then Stop: Resume Next
End Sub

Public Sub LockMenu(bDoUnlock As Boolean)
    With frmMain
        .mnuFile.Enabled = bDoUnlock
        .mnuTools.Enabled = bDoUnlock
        .mnuHelp.Enabled = bDoUnlock
    End With
End Sub

Public Sub LockInterface(bAllowInfoButtons As Boolean, bDoUnlock As Boolean)
    'Lock controls when scanning
    On Error GoTo ErrorHandler:
    
    Dim mnu As Menu
    Dim Ctl As Control
    
    For Each Ctl In frmMain.Controls
        If TypeName(Ctl) = "Menu" Then
            Set mnu = Ctl
            If InStr(1, mnu.Name, "delim", 1) = 0 Then
                mnu.Enabled = bDoUnlock
            End If
        End If
    Next
    Set mnu = Nothing
    Set Ctl = Nothing
    
    With frmMain
        .cmdScan.Enabled = bDoUnlock
        .cmdFix.Enabled = bDoUnlock
        If Not bAllowInfoButtons Or bDoUnlock Then 'if allow pressing info..., analyze this
            .cmdInfo.Enabled = bDoUnlock
            .cmdAnalyze.Enabled = bDoUnlock
        End If
        .cmdMainMenu.Enabled = bDoUnlock
        .cmdHelp.Enabled = bDoUnlock
        .cmdConfig.Enabled = bDoUnlock
        .cmdSaveDef.Enabled = bDoUnlock
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "LockInterface", "bAllowInfoButtons: " & bAllowInfoButtons, "bDoUnlock: " & bDoUnlock
    If inIDE Then Stop: Resume Next
End Sub

Public Sub LockInterfaceMain(bDoUnlock As Boolean)
    With frmMain
        .cmdN00bLog.Enabled = bDoUnlock
        .cmdN00bScan.Enabled = bDoUnlock
        .cmdN00bBackups.Enabled = bDoUnlock
        .cmdFixing.Enabled = bDoUnlock
        .cmdN00bHJTQuickStart.Enabled = bDoUnlock
        .FraIncludeSections.Enabled = bDoUnlock
        .fraScanOpt.Enabled = bDoUnlock
        '.cmdN00bClose.Enabled = bDoUnlock
    End With
End Sub

Public Function TrimEx(ByVal sStr As String, sDelimiter As String) As String
    Dim iLenDelim As Long
    iLenDelim = Len(sDelimiter)
    If iLenDelim = 0 Then Exit Function
    Do While Left$(sStr, iLenDelim) = sDelimiter And Len(sStr) <> 0
        sStr = mid$(sStr, iLenDelim + 1)
    Loop
    Do While Right$(sStr, iLenDelim) = sDelimiter And Len(sStr) <> 0
        sStr = Left$(sStr, Len(sStr) - iLenDelim)
    Loop
    TrimEx = sStr
End Function

Private Sub GetSpecialFolders_XP_2(sLog As clsStringBuilder)
    On Error GoTo ErrorHandler:
    
    sLog.AppendLine "ADMINTOOLS = " & GetSpecialFolderPath(CSIDL_ADMINTOOLS)
    If OSver.MajorMinor >= 6 Then sLog.AppendLine "ALTSTARTUP = " & GetSpecialFolderPath(CSIDL_ALTSTARTUP) 'Vista+
    sLog.AppendLine "APPDATA = " & GetSpecialFolderPath(CSIDL_APPDATA)
    'sLog.AppendLine "CSIDL_BITBUCKET = " & "(virtual)"
    sLog.AppendLine "CDBURN_AREA = " & GetSpecialFolderPath(CSIDL_CDBURN_AREA)
    sLog.AppendLine "COMMON_ADMINTOOLS = " & GetSpecialFolderPath(CSIDL_COMMON_ADMINTOOLS)
    If OSver.MajorMinor >= 6 Then sLog.AppendLine "COMMON_ALTSTARTUP = " & GetSpecialFolderPath(CSIDL_COMMON_ALTSTARTUP) 'Vista+
    sLog.AppendLine "COMMON_APPDATA = " & GetSpecialFolderPath(CSIDL_COMMON_APPDATA)
    sLog.AppendLine "COMMON_DESKTOPDIRECTORY = " & GetSpecialFolderPath(CSIDL_COMMON_DESKTOPDIRECTORY)
    sLog.AppendLine "COMMON_DOCUMENTS = " & GetSpecialFolderPath(CSIDL_COMMON_DOCUMENTS)
    sLog.AppendLine "COMMON_FAVORITES = " & GetSpecialFolderPath(CSIDL_COMMON_FAVORITES)
    sLog.AppendLine "COMMON_MUSIC = " & GetSpecialFolderPath(CSIDL_COMMON_MUSIC)
    If OSver.MajorMinor >= 6 Then sLog.AppendLine "COMMON_OEM_LINKS = " & GetSpecialFolderPath(CSIDL_COMMON_OEM_LINKS) 'Vista+"
    sLog.AppendLine "COMMON_PICTURES = " & GetSpecialFolderPath(CSIDL_COMMON_PICTURES)
    sLog.AppendLine "COMMON_PROGRAMS = " & GetSpecialFolderPath(CSIDL_COMMON_PROGRAMS)
    sLog.AppendLine "COMMON_STARTMENU = " & GetSpecialFolderPath(CSIDL_COMMON_STARTMENU)
    sLog.AppendLine "COMMON_STARTUP = " & GetSpecialFolderPath(CSIDL_COMMON_STARTUP)
    sLog.AppendLine "COMMON_TEMPLATES = " & GetSpecialFolderPath(CSIDL_COMMON_TEMPLATES)
    sLog.AppendLine "COMMON_VIDEO = " & GetSpecialFolderPath(CSIDL_COMMON_VIDEO)
    'sLog.AppendLine "COMPUTERSNEARME = " & "(virtual)"
    'sLog.AppendLine "CONNECTIONS = " & "(virtual)"
    'sLog.AppendLine "CONTROLS = " & "(virtual)"
    sLog.AppendLine "COOKIES = " & GetSpecialFolderPath(CSIDL_COOKIES)
    sLog.AppendLine "DESKTOP = " & GetSpecialFolderPath(CSIDL_DESKTOP)
    sLog.AppendLine "DESKTOPDIRECTORY = " & GetSpecialFolderPath(CSIDL_DESKTOPDIRECTORY)
    'sLog.AppendLine "DRIVES = " & "(virtual)"
    sLog.AppendLine "FAVORITES = " & GetSpecialFolderPath(CSIDL_FAVORITES)
    sLog.AppendLine "FLAG_CREATE = " & GetSpecialFolderPath(CSIDL_FLAG_CREATE)
    sLog.AppendLine "FLAG_DONT_VERIFY = " & GetSpecialFolderPath(CSIDL_FLAG_DONT_VERIFY)
    sLog.AppendLine "FLAG_MASK = " & GetSpecialFolderPath(CSIDL_FLAG_MASK)
    sLog.AppendLine "FLAG_NO_ALIAS = " & GetSpecialFolderPath(CSIDL_FLAG_NO_ALIAS)
    sLog.AppendLine "FLAG_PER_USER_INIT = " & GetSpecialFolderPath(CSIDL_FLAG_PER_USER_INIT)
    sLog.AppendLine "FONTS = " & GetSpecialFolderPath(CSIDL_FONTS)
    sLog.AppendLine "HISTORY = " & GetSpecialFolderPath(CSIDL_HISTORY)
    'sLog.AppendLine "INTERNET = " & "(virtual)"
    sLog.AppendLine "INTERNET_CACHE = " & GetSpecialFolderPath(CSIDL_INTERNET_CACHE)
    sLog.AppendLine "LOCAL_APPDATA = " & GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)
    'sLog.AppendLine "MYDOCUMENTS = " & "(virtual)"
    sLog.AppendLine "MYMUSIC = " & GetSpecialFolderPath(CSIDL_MYMUSIC)
    sLog.AppendLine "MYPICTURES = " & GetSpecialFolderPath(CSIDL_MYPICTURES)
    sLog.AppendLine "MYVIDEO = " & GetSpecialFolderPath(CSIDL_MYVIDEO)
    sLog.AppendLine "NETHOOD = " & GetSpecialFolderPath(CSIDL_NETHOOD)
    'sLog.AppendLine "NETWORK = " & "(virtual)"
    sLog.AppendLine "PERSONAL = " & GetSpecialFolderPath(CSIDL_PERSONAL)
    'sLog.AppendLine "PRINTERS = " & "(virtual)"
    sLog.AppendLine "PRINTHOOD = " & GetSpecialFolderPath(CSIDL_PRINTHOOD)
    sLog.AppendLine "PROFILE = " & GetSpecialFolderPath(CSIDL_PROFILE)
    sLog.AppendLine "PROGRAM_FILES = " & GetSpecialFolderPath(CSIDL_PROGRAM_FILES)
    sLog.AppendLine "PROGRAM_FILES_COMMON = " & GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMON)
    If OSver.IsWin64 Then
        sLog.AppendLine "PROGRAM_FILES_COMMONX86 = " & GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMONX86) 'x64
        sLog.AppendLine "PROGRAM_FILESX86 = " & GetSpecialFolderPath(CSIDL_PROGRAM_FILESX86) 'x64
    End If
    sLog.AppendLine "PROGRAMS = " & GetSpecialFolderPath(CSIDL_PROGRAMS)
    sLog.AppendLine "RECENT = " & GetSpecialFolderPath(CSIDL_RECENT)
    sLog.AppendLine "RESOURCES = " & GetSpecialFolderPath(CSIDL_RESOURCES)
    sLog.AppendLine "RESOURCES_LOCALIZED = " & GetSpecialFolderPath(CSIDL_RESOURCES_LOCALIZED)
    sLog.AppendLine "SENDTO = " & GetSpecialFolderPath(CSIDL_SENDTO)
    sLog.AppendLine "STARTMENU = " & GetSpecialFolderPath(CSIDL_STARTMENU)
    sLog.AppendLine "STARTUP = " & GetSpecialFolderPath(CSIDL_STARTUP)
    sLog.AppendLine "SYSTEM = " & GetSpecialFolderPath(CSIDL_SYSTEM)
    sLog.AppendLine "SYSTEMX86 = " & GetSpecialFolderPath(CSIDL_SYSTEMX86)
    sLog.AppendLine "TEMPLATES = " & GetSpecialFolderPath(CSIDL_TEMPLATES)
    sLog.AppendLine "WINDOWS = " & GetSpecialFolderPath(CSIDL_WINDOWS)

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetSpecialFolders_XP_2"
    If inIDE Then Stop: Resume Next
End Sub
    

Private Sub GetSpecialFoldersRegistry(sLog As clsStringBuilder)
    On Error GoTo ErrorHandler:

    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum

    Dim i As Long
    Dim aValue() As String
    Dim sData As String
    
    HE.Init HE_HIVE_HKCU, , HE_REDIR_NO_WOW
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    
    Do While HE.MoveNext
        sLog.AppendLine
        sLog.AppendLine "[" & HE.KeyAndHivePhysical & "]"
        
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue, HE.Redirected)
            If InStr(1, aValue(i), "Do not use this registry key", 1) = 0 Then
                
                sData = Reg.GetString(HE.Hive, HE.Key, aValue(i), HE.Redirected, True)
                sLog.AppendLine aValue(i) & " = " & sData
            End If
        Next
    Loop
    
    ' Because of Hive enum order precede key enum. lol. Need an update... // TODO
    '
    HE.Init HE_HIVE_HKLM, , HE_REDIR_NO_WOW
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    
    Do While HE.MoveNext
        sLog.AppendLine
        sLog.AppendLine "[" & HE.KeyAndHivePhysical & "]"
        
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue, HE.Redirected)

            sData = Reg.GetString(HE.Hive, HE.Key, aValue(i), HE.Redirected, True)
            sLog.AppendLine aValue(i) & " = " & sData
        Next
    Loop
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetSpecialFoldersRegistry"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub GetSpecialFolders_Vista(sLog As clsStringBuilder)
    On Error GoTo ErrorHandler:
    Dim kfid As UUID
    Dim nCount As Long
    Dim pakfid As Long
    Dim i As Long
    Dim ptr As Long
    Dim Flags As Long
    Dim lpPath As Long
    Dim sPath$, sName$
    Dim aPath() As String
    
    Dim pKFM As KnownFolderManager
    Set pKFM = New KnownFolderManager
    
    Dim pKF As IKnownFolder
    Dim pKFD As KNOWNFOLDER_DEFINITION
    
    Dim pItem As IShellItem
    Dim penum1 As IEnumShellItems
    Dim pChild As IShellItem
    Dim pcl As Long
    
    Call pKFM.GetFolderIds(pakfid, nCount)
    
    If nCount > 0 Then
        
        ptr = pakfid
        ReDim aPath(nCount - 1)
        
        For i = 1 To nCount
            memcpy kfid, ByVal ptr, LenB(kfid)  'array[idx] -> UUID
            ptr = ptr + LenB(kfid)
            
            Call pKFM.GetFolder(kfid, pKF)  'UUID -> IKnownFolder
            
            If Not (pKF Is Nothing) Then
                
                sName = vbNullString: sPath = vbNullString
                
                Call pKF.GetFolderDefinition(pKFD)   'IKnownFolder -> KNOWNFOLDER_DEFINITION
                
                If pKFD.pszName <> 0 Then
                    sName = BStrFromLPWStr(pKFD.pszName)        'get name
                End If
                
                'FreeKnownFolderDefinitionFields pKFD
                
                If sName = "RecycleBinFolder" Then
                    On Error Resume Next
                    pKF.GetShellItem KF_FLAG_DEFAULT, IID_IShellItem, pItem
                    On Error GoTo ErrorHandler:

                    If Not (pItem Is Nothing) Then

                        'special method: retrieve path by listing first child item

                        pItem.BindToHandler ByVal 0&, BHID_EnumItems, IID_IEnumShellItems, penum1

                        If Not (penum1 Is Nothing) Then
                            If penum1.Next(1&, pChild, pcl) = S_OK Then

                                pChild.GetAttributes SFGAO_FILESYSTEM, Flags
                                If Flags And SFGAO_FILESYSTEM Then
                                    pChild.GetDisplayName SIGDN_FILESYSPATH, lpPath
                                    sPath = BStrFromLPWStr(lpPath, True)
                                    sPath = GetParentDir(sPath)
                                End If
                            End If
                        End If
                    End If
                Else
                    Flags = (KF_FLAG_SIMPLE_IDLIST Or KF_FLAG_DONT_VERIFY Or KF_FLAG_DEFAULT_PATH Or KF_FLAG_NOT_PARENT_RELATIVE)
                    On Error Resume Next
                    Call pKF.getPath(Flags, lpPath)      'IKnownFolder -> physical path
                    If lpPath <> 0 Then
                        sPath = BStrFromLPWStr(lpPath, True)
                    End If
                    On Error GoTo ErrorHandler:
                End If
                
                If Len(sPath) <> 0 Then
                    aPath(i - 1) = sName & " = " & sPath
                End If
            End If
        Next
        
        CoTaskMemFree pakfid
    End If
    
    Set pKFM = Nothing
    
    If AryPtr(aPath) Then
        CompressArray aPath
        QuickSort aPath, 0, UBound(aPath)
        For i = 0 To UBound(aPath)
            sLog.AppendLine aPath(i)
        Next
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetSpecialFolders_Vista"
    If inIDE Then Stop: Resume Next
End Sub

Public Function CreateLogFile() As String
    Dim sLog As clsStringBuilder
    Dim i&, j&, sProcessList$
    Dim lNumProcesses&
    Dim sProcessName$
    Dim col As New Collection, cnt&
    Dim sTmp$
    Dim bStadyMakeLog As Boolean
    Dim aPos() As Variant, aNames() As String
    Dim sScanMode As String
    Dim dModule As clsTrickHashTable
    Dim aModules() As String
    Dim sModule As String
    Dim sModuleList As String
    Dim bMicrosoft As Boolean
    
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.CreateLogFile - Begin"
    
    Set sLog = New clsStringBuilder 'speed optimization for huge logs
    
    If Not bLogProcesses Then GoTo MakeLog
    
    If AryPtr(gProcess) Then
        lNumProcesses = UBound(gProcess) + 1
    End If
    
    If (lNumProcesses > 0) Then
    
        For i = 0 To UBound(gProcess)
            
            sProcessName = gProcess(i).path
            
            If Len(gProcess(i).path) = 0 Or gProcess(i).Minimal Then
                If bIgnoreAllWhitelists Or Not IsMinimalProcess_ForLog(gProcess(i).pid, gProcess(i).Name) Then
                    sProcessName = gProcess(i).Name
                Else
                    sProcessName = vbNullString
                End If
            End If
            
            If Len(sProcessName) <> 0 Then
                If Not isCollectionKeyExists(sProcessName, col) Then
                    col.Add 1&, sProcessName          ' item - count, key - name of process
                Else
                    cnt = col.Item(sProcessName)      ' replacing item of collection
                    col.Remove (sProcessName)
                    col.Add cnt + 1&, sProcessName    ' increase count of identical processes
                End If
            End If
        Next
    End If
    
    'sProcessList = "Running processes:" & vbCrLf
    sProcessList = Translate(29) & ":" & vbCrLf
    
    'If bAdditional Then
    '    sProcessList = sProcessList & "  PID | " & Translate(1021) & vbCrLf
    'Else
        'Number | Path
        sProcessList = sProcessList & Translate(1020) & " | " & Translate(1021) & vbCrLf
    'End If
    
    If bLogModules Then
        'Additional mode => PID | Process Name
        
        ReDim aPos(UBound(gProcess)), aNames(UBound(gProcess))
        
        For i = 0 To UBound(gProcess)
            aPos(i) = i
            If Len(gProcess(i).path) = 0 Then
                gProcess(i).path = gProcess(i).Name
            End If
            aNames(i) = gProcess(i).path
        Next
        
        QuickSortSpecial aNames, aPos, 0, UBound(gProcess)
        
        For i = 0 To UBound(aPos)
            With gProcess(aPos(i))
                '// TODO: add 'is microsoft' check and mark
                sProcessList = sProcessList & Right$("     " & .pid & "  ", 8) & .path & vbCrLf
            End With
        Next
        
    Else
        'Normal mode => Number of processes | Process Name
    
        ' Sort using positions array method (Key - Process Path).
        Dim ProcLog() As MY_PROC_LOG
        ReDim ProcLog(col.Count) As MY_PROC_LOG
        ReDim aPos(col.Count) As Variant
        ReDim aNames(col.Count) As String
        
        'Dim SignResult  As SignResult_TYPE
        
        For i = 1& To col.Count
            With ProcLog(i)
                .ProcName = GetCollectionKeyByIndex(i, col)
                .Number = col(i)
                
    ' I temporarily disable EDS checking
    '            SignVerify .ProcName, 0&, SignResult
    '            If SignResult.isLegit Then
    '                .EDS_issued = SignResult.SubjectName
    '            End If
    '
    '            If Not bIgnoreAllWhitelists Then
    '                UpdateProgressBar "ProcList", .ProcName
    '                .IsMicrosoft = (IsMicrosoftCertHash(SignResult.HashRootCert) And SignResult.isLegit)  'validate EDS
    '            End If
                
                aNames(i) = IIf(.IsMicrosoft, "(sign: 'Microsoft') ", IIf(Len(.EDS_issued) <> 0, "(" & .EDS_issued & ") ", " " & STR_NOT_SIGNED)) & .ProcName
                aPos(i) = i
            End With
        Next
        
        QuickSortSpecial aNames, aPos, 0, col.Count
        
        For i = 1& To UBound(aPos)
            With ProcLog(aPos(i))
                sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & IIf(.IsMicrosoft, "(sign: 'Microsoft') ", vbNullString) & .ProcName & vbCrLf
            End With
        Next
    End If
    
    sProcessList = sProcessList & vbCrLf
    
    'show all PIDs in debug. mode
    If bDebug Or bDebugToFile Then
        If lNumProcesses Then
            sTmp = vbNullString
            For i = 0 To UBound(gProcess)
                sTmp = sTmp & gProcess(i).pid & " | " & IIf(Len(gProcess(i).path) <> 0, gProcess(i).path, gProcess(i).Name) & vbCrLf
            Next
            AppendErrorLogCustom sTmp
            sTmp = vbNullString
        End If
    End If
    
    If bLogModules Then
        'Loaded modules
        
        Set dModule = New clsTrickHashTable
        
        For i = 0 To UBound(gProcess)
        
            If gProcess(i).pid <> 0 And gProcess(i).pid <> 4 Then
        
                aModules = EnumModules64(gProcess(i).pid)
        
                If AryItems(aModules) Then
                    For j = 0 To UBound(aModules)
                        If Not dModule.Exists(aModules(j)) Then
                            dModule.Add aModules(j), "[" & CStr(gProcess(i).pid) & "]"
                        Else
                            dModule(aModules(j)) = dModule(aModules(j)) & " [" & CStr(gProcess(i).pid) & "]"
                        End If
                    Next
                End If
            End If
        Next
        
        If dModule.Count > 0 Then
            ReDim aPos(dModule.Count - 1)
            ReDim aNames(dModule.Count - 1)
        
            For i = 0 To dModule.Count - 1
                aPos(i) = i
                aNames(i) = dModule.Keys(i)
            Next
            
            QuickSortSpecial aNames, aPos, 0, UBound(aNames)
            
            '// TODO: add command line extraction via PEB
            
            'Loaded modules: Path | PID | Command line
            sModuleList = Translate(21) & ":" & vbCrLf & Translate(1021) & " | PID" & vbCrLf 'translate(1070)
            
            For i = 0 To UBound(aPos)
                If Not bAutoLogSilent Then
                    DoEvents
                    UpdateProgressBar "ModuleList", CStr(i + 1) & " / " & CStr(UBound(aPos) + 1)
                End If
                sModule = dModule.Keys(aPos(i))
                bMicrosoft = IsMicrosoftFile(sModule)
                
                If (Not bMicrosoft) Or Not bHideMicrosoft Then
                    sModuleList = sModuleList & sModule & vbTab & vbTab & vbTab & dModule.Items(aPos(i)) & vbCrLf
                End If
            Next
            
        End If
        
        Set dModule = Nothing
    End If
    
    '------------------------------
MakeLog:
    bStadyMakeLog = True
    
    UpdateProgressBar "Report"
    
    'UpdateProgressBar "Finish"
    'DoEvents
    
    sLog.Append ChrW$(-257) & "Logfile of " & AppVerPlusName & vbCrLf & vbCrLf ' + BOM UTF-16 LE
    
    sLog.Append MakeLogHeader()
    sLog.AppendLine ""
    
    Dim tmp$
    With GetBrowsersInfo() 'BROWSERS_VERSION_INFO
        tmp = .Opera.Version
        If Len(tmp) Then sLog.Append "Opera:   " & tmp & vbCrLf
        tmp = .Chrome.Version
        If Len(tmp) Then sLog.Append "Chrome:  " & tmp & vbCrLf
        tmp = .Firefox.Version
        If Len(tmp) Then sLog.Append "Firefox: " & tmp & vbCrLf
        tmp = .Edge.Version
        If Len(tmp) Then sLog.Append "Edge:    " & tmp & vbCrLf
        tmp = .IE.Version
        If Len(tmp) Then sLog.Append "Internet Explorer: " & tmp & vbCrLf
                         sLog.Append "Default: " & .Default & vbCrLf
    End With
    
    sLog.AppendLine ""
    
    Dim sBootMode As String
    sBootMode = "Boot mode: " & OSver.SafeBootMode
    If OSver.SecureBootSupported Then
        sBootMode = sBootMode & " (Secure Boot: " & IIf(OSver.SecureBoot, "On", "Off") & ")"
    End If
    If OSver.TestSigning Then
        sBootMode = sBootMode & " (Test Signing: On)"
    End If
    If OSver.DebugMode Then
        sBootMode = sBootMode & " (Debug Mode: On)"
    End If
    If OSver.CodeIntegrity Then
        sBootMode = sBootMode & " (Code Integrity: On)"
    End If
    
    sLog.AppendLine sBootMode
    
    If (Not bLogProcesses) Or (Not bAdditional) Or bLogModules Or bLogEnvVars Or (Not bHideMicrosoft) Or bIgnoreAllWhitelists Then
        If Not bAdditional Then
            sScanMode = "Skip Additional"
        End If
        If Not bLogProcesses Then
            sScanMode = sScanMode & "; Skip Processes"
        End If
        If bLogModules Then
            sScanMode = sScanMode & "; Loaded Modules"
        End If
        If bLogEnvVars Then
            sScanMode = sScanMode & "; Environment variables"
        End If
        If Not bHideMicrosoft Then
            sScanMode = sScanMode & "; Don't hide Microsoft"
        End If
        If bIgnoreAllWhitelists Then
            sScanMode = sScanMode & "; Ignore ALL Whitelists"
        End If
        If Left$(sScanMode, 2) = "; " Then sScanMode = mid$(sScanMode, 3)
        
        sLog.Append "Scan mode: " & sScanMode & vbCrLf
    End If
    
'    If OSver.IsSystemCaseSensitive Then
'        sLog.Append "Filenames: is in case sensitive mode" & vbCrLf
'    End If
    
    If bLogEnvVars Then
        
        Dim pEnv As Long
        Dim pEnvNext As Long
        Dim strEnv As String
        Dim sEnvName As String
        Dim sEnvValue As String
        Dim varCnt As Long
        Dim aValues() As String
        Dim pos As Long
        
        'Dim varDict As clsTrickHashTable
        'Set varDict = New clsTrickHashTable
        
        sLog.Append vbCrLf & "Environment variables:" & vbCrLf & vbCrLf
        
        'get System EV
        sLog.AppendLine "[System]"
        For i = 1 To Reg.EnumValuesToArray(HKLM, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", aValues, False)
            sEnvName = aValues(i)
            sEnvValue = Reg.GetString(HKLM, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", sEnvName, False)
            sLog.AppendLine sEnvName & " = " & sEnvValue
            'If Not varDict.Exists(sEnvName) Then varDict.Add sEnvName, sEnvValue
        Next
        
        'get User EV
        sLog.AppendLine
        sLog.AppendLine "[User]"
        varCnt = Reg.EnumValuesToArray(HKCU, "Environment", aValues, False)
        For i = 1 To varCnt
            sEnvName = aValues(i)
            sEnvValue = Reg.GetString(HKCU, "Environment", sEnvName, False)
            sLog.AppendLine sEnvName & " = " & sEnvValue
            'If Not varDict.Exists(sEnvName) Then varDict.Add sEnvName, sEnvValue
        Next
        If varCnt = 0 Then
            sLog.AppendLine "No variables."
        End If
        
        'get process EV
        sLog.AppendLine
        sLog.AppendLine "[Current process]"
        varCnt = 0
        
        pEnv = GetEnvironmentStrings()
        pEnvNext = pEnv
        
        If pEnv <> 0 Then
            Do
                strEnv = StringFromPtrW(pEnvNext)
                
                If Len(strEnv) <> 0 Then
                    pos = InStr(2, strEnv, "=")
                    If pos = 0 Then pos = InStr(1, strEnv, "=")
                    If pos <> 0 Then
                        sEnvName = Left$(strEnv, pos - 1)
                        sEnvValue = mid$(strEnv, pos + 1)
                        'bAddEV = True
                        'If varDict.Exists(sEnvName) Then
                        '    If varDict(sEnvName) = sEnvValue Then bAddEV = False
                        'End If
                        'If bAddEV Then
                            sLog.AppendLine sEnvName & " = " & sEnvValue
                            varCnt = varCnt + 1
                        'End If
                    End If
                    pEnvNext = pEnvNext + LenB(strEnv) + 2
                End If
            Loop Until Len(strEnv) = 0
            
            FreeEnvironmentStrings pEnv
        End If
        If varCnt = 0 Then
            'sLog.AppendLine "All variables are identical to System/User."
        End If
        
        'Set varDict = Nothing
        
        sLog.AppendLine
        sLog.AppendLine "Special folders:"
        sLog.AppendLine
        sLog.AppendLine "[CLSID]"

        If OSver.IsWindowsVistaOrGreater Then
            GetSpecialFolders_Vista sLog
        Else
            GetSpecialFolders_XP_2 sLog
        End If
        
        'because Microsoft doesn't care about its own "Best practice", mentioned in MSDN,
        'so Special Folders retrieved by CLSID become completely useless, e.g. for "Desktop" redirection in Win 10.
        GetSpecialFoldersRegistry sLog
    End If
    
    sLog.Append vbCrLf & sProcessList
    
    If bLogModules Then
        sLog.Append sModuleList & vbCrLf & vbCrLf
    End If
    
    ' -----> MAIN Sections <------
    
    ' in /silentautolog mode result screen is not fill due to speed optimization
    If (Not bAutoLogSilent) And frmMain.lstResults.ListCount = 0 Then
        If g_bGeneralScanned Then
            sLog.AppendLine Translate(1004) 'No suspicious items found!
            AppendErrorLogCustom "[lstResults] " & Translate(1004)
        End If
    Else
        If AryItems(HitSorted) Then
            For i = 0 To UBound(HitSorted)
                ' Adding empty lines beetween sections (cancelled)
                'sPrefix = rtrim$(Splitsafe(HitSorted(i), "-")(0))
                'If sPrefixLast <> vbnullstring And sPrefixLast <> sPrefix Then sLog = sLog & vbCrLf
                'sPrefixLast = sPrefix
                sLog.Append HitSorted(i) & vbCrLf
            Next
        End If
    End If
    
    ' ----------------------------
    
    Dim IgnoreCnt&
    IgnoreCnt = RegReadHJT("IgnoreNum", "0")
    If IgnoreCnt <> 0 Then
        If bSkipIgnoreList Or bLoadDefaults Then
            ' "Warning: Ignore list contains " & IgnoreCnt & " items, but they are displayed in log, because /default (/skipIgnoreList) switch is used." & vbCrLf
            sLog.Append vbCrLf & vbCrLf & Replace$(Replace$(Translate(1017), "[]", IgnoreCnt), "[*]", IIf(bLoadDefaults, "default", "skipIgnoreList")) & vbCrLf
        Else
            ' "Warning: Ignore list contains " & IgnoreCnt & " items." & vbCrLf
            sLog.Append vbCrLf & vbCrLf & Replace$(Translate(1011), "[]", IgnoreCnt) & vbCrLf
        End If
    End If
    If Not g_bGeneralScanned Then
        ' "Warning: General scanning was not performed." & vbCrLf
        sLog.Append vbCrLf & vbCrLf & Translate(1012) & vbCrLf
    End If
    Dim SignMsg As String
    If Not isEDS_Work(True, SignMsg) Then
        sLog.Append vbCrLf & vbCrLf & BuildPath(sWinDir, "system32\ntdll.dll") & " file doesn't pass digital signature verification. Error: " & SignMsg & vbCrLf
    End If
    If Not isEDS_CatExist() Then
        sLog.Append "Windows security catalogue is empty!" & vbCrLf
    End If
    
    If Len(EndReport) <> 0 Then
        sLog.AppendLine vbCrLf & EndReport
    End If
    EndReport = vbNullString
    
    'Append by Error Log
    If 0 <> Len(ErrReport) Then
        sLog.Append vbCrLf & vbCrLf & "Debug information:" & vbCrLf & ErrReport & vbCrLf
        '& vbCrLf & "CmdLine: " & AppPath(True) & " " & g_sCommandLine
    End If
    
    Dim b()     As Byte
    
    If bDebugToFile Then
        If g_hDebugLog <> 0 Then
            b() = vbCrLf & vbCrLf & "Contents of the main logfile:" & vbCrLf & vbCrLf & sLog.ToString & vbCrLf
            PutW_NoLog g_hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
        End If
    End If
    
    If 0 <> ErrLogCustomText.Length Then
        sLog.Append vbCrLf & vbCrLf & "Trace information:" & vbCrLf & ErrLogCustomText.ToString & vbCrLf
    End If
    
    If bAutoLog Then Perf.EndTime = GetTickCount()
    sLog.Append vbCrLf & "--" & vbCrLf & "End of file - " & "Time spent: " & ((Perf.EndTime - Perf.StartTime) \ 100) / 10 & " sec. - "
    
    If bDebugToFile Then
        If g_hDebugLog <> 0 Then
            b() = vbCrLf & "--" & vbCrLf & "End of file - " & "Time spent: " & ((Perf.EndTime - Perf.StartTime) \ 100) / 10 & " sec."
            PutW_NoLog g_hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
            CloseW g_hDebugLog, True: g_hDebugLog = 0
        End If
    End If
    
    Dim Size_1 As Long
    Dim Size_2 As Long
    Dim Size_3 As Long
    
    Size_1 = 2& * (sLog.Length + Len(" bytes, CRC32: FFFFFFFF. Sign:   "))   '���������� ������� ���� (� ������)
    Size_2 = Size_1 + 2& * Len(CStr(Size_1))                                 '� ������ ������ ����� "���-�� ����"
    Size_3 = Size_2 - 2& * Len(CStr(Size_1)) + 2& * Len(CStr(Size_2))        '��������, ���� ����� ���� ����������� �� 1 ������
    
    sLog.Append CStr(Size_3) & " bytes, CRC32: FFFFFFFF. Sign: "
    
    Dim ForwCRC As Long
    
    b() = sLog.ToString                                                 '������� CRC ����
    ForwCRC = CalcArrayCRCLong(b()) Xor -1
    
    Dim CorrBytes$: CorrBytes = RecoverCRC(ForwCRC, &HFFFFFFFF)         '������� ����� �������������
    
    ReDim Preserve b(UBound(b) + 4)                                     '��������� �� � ����� �������
    b(UBound(b) - 3) = Asc(mid$(CorrBytes, 1, 1))
    b(UBound(b) - 2) = Asc(mid$(CorrBytes, 2, 1))
    b(UBound(b) - 1) = Asc(mid$(CorrBytes, 3, 1))
    b(UBound(b) - 0) = Asc(mid$(CorrBytes, 4, 1))
    
    CreateLogFile = b()
    
    Set sLog = Nothing
    
    AppendErrorLogCustom "frmMain.CreateLogFile - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "frmMain_CreateLogFile"
    If inIDE Then Stop: Resume Next
    If Not bStadyMakeLog Then GoTo MakeLog
    Set sLog = Nothing
End Function

Private Function isEDS_CatExist() As Boolean
    'check is it have at least 10 cat. files (141 is in Win2k, 30 - is in XP SP2)
    'c:\Windows\System32\catroot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}
    
    Const NUM_REQUIRED = 10&
    
    Dim hFind As Long
    Dim cnt As Long
    Dim fd As WIN32_FIND_DATA
        
    hFind = FindFirstFile(StrPtr(BuildPath(sWinSysDir, "catroot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\*.cat")), fd)
    If hFind <> 0& Then
        Do Until FindNextFile(hFind, fd) = 0& Or cnt >= NUM_REQUIRED
            cnt = cnt + 1
        Loop
        FindClose hFind
    End If
    
    isEDS_CatExist = (cnt >= NUM_REQUIRED)
End Function

Public Function MakeLogHeader() As String

    If Len(g_sLogHeaderCache) <> 0 And bFirstScanAfterProgramStarted Then
        MakeLogHeader = g_sLogHeaderCache
        Exit Function
    End If

    Dim TimeCreated As String
    Dim bSPOld As Boolean
    Dim sUTC As String
    Dim sText As String
    
    'Service pack relevance checking
    
    Select Case OSver.MajorMinor
      
        Case 10
        
        Case 6.3
      
        Case 6.4 '10 Technical preview
            bSPOld = True
            
        Case 6.2 '8
            If Not OSver.IsServer Then bSPOld = True
            
        Case 6.1 '7 / Server 2008 R2
            If OSver.SPVer < 1 Then bSPOld = True
            
        Case 6 'Vista / Server 2008
            If OSver.SPVer < 2 Then bSPOld = True
            
        Case 5.2 'XP x64 / Server 2003 / Server 2003 R2
            If OSver.SPVer < 2 Then bSPOld = True
        
        Case 5.1 'XP
            If OSver.SPVer < 3 Then bSPOld = True
        
        Case 5 '2k / 2k Server
            If OSver.SPVer < 4 Then bSPOld = True
        
    End Select
    
    If GetTimeZone(sUTC) Then
        sUTC = "UTC" & sUTC
    Else
        sUTC = "UTC is unknown"
    End If
    
    TimeCreated = Right$("0" & Day(Now), 2) & "." & Right$("0" & Month(Now), 2) & "." & Year(Now) & " - " & _
        Right$("0" & Hour(Now), 2) & ":" & Right$("0" & Minute(Now), 2)
    
    sText = "Platform:  " & OSver.Bitness & " " & OSver.OSName & IIf(Len(OSver.Edition) <> 0, " (" & OSver.Edition & ")", vbNullString) & ", " & _
            OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & _
            IIf(OSver.ReleaseId <> 0, " (ReleaseId: " & OSver.ReleaseId & IIf(OSver.DisplayVersion <> "", ", " & OSver.DisplayVersion, "") & ")", vbNullString) & ", " & _
            "Service Pack: " & OSver.SPVer & IIf(bSPOld, " <=== Attention! (outdated SP)", vbNullString) & _
            IIf(OSver.MajorMinor <> OSver.MajorMinorNTDLL And OSver.MajorMinorNTDLL <> 0, " (ntdll.dll = " & OSver.NtDllVersion & ")", vbNullString) & _
            vbCrLf
    
    '," & vbTab & "Uptime: " & TrimSeconds(GetSystemUpTime()) & " h/m" & vbCrLf
    
    sText = sText & "Time:      " & TimeCreated & " (" & sUTC & ")" & vbCrLf
    sText = sText & "Language:  " & "OS: " & OSver.LangSystemNameFull & " (" & "0x" & Hex$(OSver.LangSystemCode) & "). " & _
            "Display: " & OSver.LangDisplayNameFull & " (" & "0x" & Hex$(OSver.LangDisplayCode) & "). " & _
            "Non-Unicode: " & OSver.LangNonUnicodeNameFull & " (" & "0x" & Hex$(OSver.LangNonUnicodeCode) & ")" & vbCrLf
    
    Dim iFreeSpace As Currency, dblFreeSpace As Double
    Dim iTotalSpace As Currency, dblTotalSpace As Double
    iFreeSpace = cDrives.GetFreeSpace(SysDisk, True, iTotalSpace)
    dblFreeSpace = iFreeSpace / 107374.1824
    dblTotalSpace = iTotalSpace / 107374.1824
    
    Dim diskTech As String
    Select Case cDrives.GetStorageTechnology(SysDisk)
        Case STORAGE_TECHNOLOGY_SSD: diskTech = "SSD"
        Case STORAGE_TECHNOLOGY_HDD: diskTech = "HDD"
        Case Else: diskTech = "Unknown tech"
    End Select
    
    Dim diskStyle As String
    Select Case cDrives.GetPartitionStyle(SysDisk)
        Case PARTITION_STYLE_MBR: diskStyle = "MBR"
        Case PARTITION_STYLE_GPT: diskStyle = "GPT"
        Case Else: diskTech = "Unknown style"
    End Select
    
    sText = sText & "Memory:    " & Format$(OSver.MemoryFree / 1024, "0.00") & " GiB Free / " & Round(OSver.MemoryTotal / 1024) & _
        ". Loading RAM (" & OSver.MemoryLoad & " %)"
    If OSver.IsWindowsVistaOrGreater Then
        sText = sText & ", CPU (" & IIf(g_iCpuUsage <> 0, g_iCpuUsage, CLng(OSver.CpuUsage)) & " %)" & vbCrLf
    Else
        sText = sText & vbCrLf
    End If
    
    sText = sText & "Disk " & SysDisk & "    " & Format$(dblFreeSpace, "0.00") & " GiB Free / " & Round(dblTotalSpace) & _
        " (" & diskTech & ", " & diskStyle & ")" & vbCrLf
    
    If OSver.MajorMinor >= 6 Then
        sText = sText & "Elevated:  " & IIf(OSver.IsElevated, "Yes", "No") & vbCrLf  '& vbTab & "IL: " & OSver.GetIntegrityLevel & vbCrLf
    End If
    
    Dim sAccType As String
    If OSver.IsWindows8OrGreater Then
        sAccType = GetCurrentUserAccountType()
        If Len(sAccType) <> 0 Then sAccType = "; type: " & sAccType
    End If
    
    sText = sText & "Ran by:    " & OSver.UserName & vbTab & "(group: " & OSver.UserType & sAccType & ") on " & OSver.ComputerName & _
        ", " & IIf(bDebugMode, "(SID: " & OSver.SID_CurrentProcess & ") ", vbNullString) & "FirstRun: " & IIf(bFirstRebootScan, "yes", "no") & _
        IIf(OSver.IsLocalSystemContext, " <=== Attention! ('Local System' account)", vbNullString) & vbCrLf
    
    MakeLogHeader = sText
    g_sLogHeaderCache = sText
End Function

' ���������� �� �����. �� ���� - ������ j(), �� ������ ������ k() � ��������� ������� j � ��������������� ������� + ��������������� ������.
' �������� ����� ������� �������� ��� ���������� User type arrays �� ������ �� �����.
Public Sub QuickSortSpecial(j() As String, k() As Variant, low As Long, high As Long, Optional CompareMethod As VbCompareMethod = vbTextCompare)
    On Error GoTo ErrorHandler
    Dim i As Long, L As Long, pM As Long, wsp As Long, pSA As Long, lcid As Long, VT As Variant
    i = low: L = high: pM = StrPtr(j((i + L) \ 2))
    pSA = Deref(AryPtr(j) + 12) 'SAFEARRAY.pvData => BSTR
    lcid = OSver.LCID_UserDefault
    Do Until i > L
        Do While VarBstrCmp(StrPtr(j(i)), pM, lcid, CompareMethod) = VARCMP_LT: i = i + 1: Loop
        Do While VarBstrCmp(StrPtr(j(L)), pM, lcid, CompareMethod) = VARCMP_GT: L = L - 1: Loop
        If (i <= L) Then
            wsp = StrPtr(j(L))
            PutMem4 ByVal (pSA + 4 * L), StrPtr(j(i))
            PutMem4 ByVal (pSA + 4 * i), wsp
            VT = k(i): k(i) = k(L): k(L) = VT
            i = i + 1: L = L - 1
        End If
    Loop
    If low < L Then QuickSortSpecial j, k, low, L, CompareMethod
    If i < high Then QuickSortSpecial j, k, i, high, CompareMethod
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "QuickSortSpecial"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub SortSectionsOfResultList()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.SortSectionsOfResultList - Begin"
    
    ' -----> Sorting of items in results window (ANSI)
    
    Dim Hit() As String
    Dim i As Long
    Dim sSelItem As String
    'HitSorted() -> is a global array
    
    'save selected position on listbox
    If frmMain.lstResults.ListIndex <> -1 Then
        sSelItem = frmMain.lstResults.List(frmMain.lstResults.ListIndex)
    End If
    
    Dim iOldTop&
    iOldTop = frmMain.lstResults.TopIndex
    
    Erase HitSorted
    
    If frmMain.lstResults.ListCount <> 0 Then
    
        ReDim Hit(frmMain.lstResults.ListCount - 1)
        
        AppendErrorLogCustom "SortSectionsOfResultList. Items before sorting:"
        
        For i = 0 To frmMain.lstResults.ListCount - 1
            Hit(i) = frmMain.lstResults.List(i)
            AppendErrorLogCustom "[lstResults] " & Hit(i)
        Next i
        
        SortSectionsOfResultList_Ex Hit, HitSorted
        
        AppendErrorLogCustom "SortSectionsOfResultList. Items after sorting:"
        
        ' Rearrange listbox data accorting to sorted list of sections
        frmMain.lstResults.Clear
        For i = 0 To UBound(HitSorted)
            frmMain.lstResults.AddItem HitSorted(i)
            AppendErrorLogCustom "[lstResults] " & HitSorted(i)
        Next
        
    End If
    
    ' -----> Sorting of items in global array (Unicode)
    '
    ' Number of items can be different beetween log and results window,
    ' e.g. O1 - Hosts limited to ~ 20 items for results windows, when in the same time all items are included in the logfile.
    
    If AryPtr(Scan) = 0 Then
    
        ReDim HitSorted(0)
    Else
        
        ReDim Hit(UBound(Scan))
        
        For i = 0 To UBound(Scan)
            Hit(i) = Scan(i).HitLineW
        Next i
        
        SortSectionsOfResultList_Ex Hit, HitSorted
    End If
    
    'restore selected position in listbox
    If Len(sSelItem) <> 0 Then
        For i = 0 To frmMain.lstResults.ListCount - 1
            If StrComp(frmMain.lstResults.List(i), sSelItem) = 0 Then
                frmMain.lstResults.ListIndex = i
                Exit For
            End If
        Next
    End If

    If iOldTop <> -1 Then frmMain.lstResults.TopIndex = iOldTop
    
    Perf.EndTime = GetTickCount()
    
    AppendErrorLogCustom "frmMain.SortSectionsOfResultList - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_SortSectionsOfResultList"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub SortSectionsOfResultList_Ex(aSrcArray() As String, aDstArray() As String)
    On Error GoTo ErrorHandler:

    ' Special sort procedure
    ' ---------------------------------

    Dim SectSorted() As String
    Dim SectName As String
    Dim nHit As Long
    Dim nSect As Long
    Dim nItemsSect As Long
    Dim nItemsHit As Long
    Dim pos As Long
    Dim i As Long
    Dim bComply As Boolean
    Dim cSectNames As Collection
    Set cSectNames = New Collection

    ReDim aDstArray(UBound(aSrcArray))
    
    cSectNames.Add "R0"
    cSectNames.Add "R1"
    cSectNames.Add "R2"
    cSectNames.Add "R3"
    cSectNames.Add "R4"
    cSectNames.Add "F0"
    cSectNames.Add "F1"
    cSectNames.Add "F2"
    cSectNames.Add "F3"
    cSectNames.Add "B"
    For i = 1 To 22
        cSectNames.Add "O" & i
    Next
    cSectNames.Add "O23 - Service"
    cSectNames.Add "O23 - Driver"
    cSectNames.Add "O23"
    For i = 24 To LAST_CHECK_OTHER_SECTION_NUMBER
        cSectNames.Add "O" & i
    Next
    
    '�������� ����������:
    '����������� ������� ������ �� ������ � ���� ��� � ������� ����
    '��� ������ ������ ������ ������, ��������� ��� � ��������� ���������� � �������������� ������
    
    nItemsHit = 0
    
    For nSect = 1 To cSectNames.Count
        nItemsSect = 0
        For nHit = 0 To UBound(aSrcArray)
            If 0 <> Len(aSrcArray(nHit)) Then
                pos = InStr(aSrcArray(nHit), "-")
                If pos = 0 Then
                    If Not bAutoLog Then
                        MsgBoxW "Warning! Wrong format of hit line. Must include dash after the name of the section!" & vbCrLf & "Line: " & aSrcArray(nHit)
                    End If
                Else
                    bComply = False
                    
                    '���� ������ ���������� �� �������� � ���������� �����
                    If InStr(cSectNames(nSect), "-") <> 0 Then
                        If StrBeginWith(aSrcArray(nHit), cSectNames(nSect)) Then bComply = True
                    Else
                        '���� ������ ���������� ������ �� ��������
                        SectName = RTrim$(Left$(aSrcArray(nHit), pos - 1))
                        If SectName = cSectNames(nSect) Then bComply = True
                    End If
                
                    '������ ���� ������������� �������� -> ��������� � ������
                    If bComply Then
                        ' ������ ��������� ������ ���� ������ ��� ����������
                        nItemsSect = nItemsSect + 1
                        ReDim Preserve SectSorted(nItemsSect - 1)
                        '�������� � SectSorted ������ �� aSrcArray
                        SectSorted(nItemsSect - 1) = aSrcArray(nHit)
                        '� � �������� ������� ������ ������
                        aSrcArray(nHit) = vbNullString
                    End If
                End If
            End If
        Next
        ' ������ ������ ���������.
        If 0 <> nItemsSect Then
            ' ������ ���������� ������
            ' O1 �� ��������� (hosts)
            If cSectNames(nSect) <> "O1" Then
                QuickSort SectSorted, 0, UBound(SectSorted)
            End If
            For i = 0 To UBound(SectSorted)
                If 0 <> Len(SectSorted(i)) Then
                    '��������� ��������������� ������ � ����� ������
                    aDstArray(nItemsHit) = SectSorted(i)
                    nItemsHit = nItemsHit + 1
                End If
            Next
        End If
    Next
    ' ���������, �� �������� �� ����������������� ���������
    ReDim SectSorted(0)
    nItemsSect = 0
    For nHit = 0 To UBound(aSrcArray)
        If 0 <> Len(aSrcArray(nHit)) Then
            '�������� ������� � ������ SectSorted
            nItemsSect = nItemsSect + 1
            ReDim Preserve SectSorted(nItemsSect - 1)
            SectSorted(nItemsSect - 1) = aSrcArray(nHit)
        End If
    Next
    If nItemsSect > 0 Then
        '��������� ���
        QuickSort SectSorted, 0, UBound(SectSorted)
        
        '� ���������� � ����� ��������������� �������
        For i = 0 To UBound(SectSorted)
            If Len(SectSorted(i)) <> 0 Then
                aDstArray(nItemsHit) = SectSorted(i)
                nItemsHit = nItemsHit + 1
            End If
        Next
    End If
    
    Set cSectNames = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_SortSectionsOfResultListEx"
    If inIDE Then Stop: Resume Next
End Sub

' Add key to jump list without actually doing cure on it
' Returns 'true' if key/value is exist and was add to jump list
Public Function AddJumpRegistry( _
    JumpArray() As JUMP_ENTRY, _
    ActionType As ENUM_REG_ACTION_BASED, _
    ByVal lHive As ENUM_REG_HIVE, _
    ByVal sKey As String, _
    Optional sParam As String = vbNullString, _
    Optional vDefaultData As Variant = vbNullString, _
    Optional eRedirected As ENUM_REG_REDIRECTION = REG_NOTREDIRECTED, _
    Optional ParamType As ENUM_REG_VALUE_TYPE_RESTORE = REG_RESTORE_SAME) As Boolean
    
    'speed hack
    If bAutoLogSilent Then Exit Function
    
    Dim KeyFix() As FIX_REG_KEY
    AddRegToFix KeyFix, ActionType, lHive, sKey, sParam, vDefaultData, eRedirected, ParamType
    
    If AryPtr(KeyFix) Then
        AddJumpRegistry = True
    
        If AryPtr(JumpArray) Then
            ReDim Preserve JumpArray(UBound(JumpArray) + 1)
        Else
            ReDim JumpArray(0)
        End If
        With JumpArray(UBound(JumpArray))
            .Type = JUMP_ENTRY_REGISTRY
            .Registry = KeyFix
        End With
    End If
End Function

' Append results array with new registry key
Public Sub AddRegToFix( _
    KeyArray() As FIX_REG_KEY, _
    ActionType As ENUM_REG_ACTION_BASED, _
    ByVal lHive As ENUM_REG_HIVE, _
    ByVal sKey As String, _
    Optional sParam As String = vbNullString, _
    Optional vDefaultData As Variant = vbNullString, _
    Optional eRedirected As ENUM_REG_REDIRECTION = REG_NOTREDIRECTED, _
    Optional ParamType As ENUM_REG_VALUE_TYPE_RESTORE = REG_RESTORE_SAME, _
    Optional ReplaceDataWhat As String = vbNullString, _
    Optional ReplaceDataInto As String = vbNullString, _
    Optional TrimDelimiter As String = vbNullString)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent And Not g_bFixing Then Exit Sub
    
    Dim vHiveFix As Variant, eHiveFix As ENUM_REG_HIVE_FIX
    Dim vUseWow As Variant, Wow6432Redir As Boolean
    Dim lActualHive As ENUM_REG_HIVE
    Dim bNoItem As Boolean
    Dim i As Long
    
    If Len(sKey) = 0 Then Exit Sub
    
    If lHive = 0 Then 'if hive handle defined by Key prefix -> ltrim prefix of sKey, and assign handle for lHive
        Call Reg.NormalizeKeyNameAndHiveHandle(lHive, sKey)
    End If
    
    If Not CBool(lHive And &H1000&) Then 'if not combined hive
        lHive = CombineHives(lHive)      'convert ENUM_REG_HIVE -> ENUM_REG_HIVE_FIX to be able to iterate
    End If
    
    For Each vHiveFix In Array(HKCR_FIX, HKCU_FIX, HKLM_FIX, HKU_FIX)
        
        eHiveFix = vHiveFix
        
        If lHive And eHiveFix Then
            
            For Each vUseWow In Array(False, True)
                
                Wow6432Redir = vUseWow
                
                If eRedirected = REG_REDIRECTION_BOTH _
                  Or ((eRedirected = REG_REDIRECTED) And (Wow6432Redir = True)) _
                  Or ((eRedirected = REG_NOTREDIRECTED) And (Wow6432Redir = False)) Then
    
                    lActualHive = ConvertHiveFixToHive(eHiveFix)
                    
                    bNoItem = False
                    
                    If (ActionType And BACKUP_KEY) Or (ActionType And REMOVE_KEY) Or (ActionType And REMOVE_KEY_IF_NO_VALUES) Or (ActionType And JUMP_KEY) Then
                        If Not Reg.HasSpecialChar(sKey) Then
                            If Not Reg.KeyExists(lActualHive, sKey, Wow6432Redir) Then bNoItem = True
                        End If
                        
                    ElseIf (ActionType And BACKUP_VALUE) Or (ActionType And REMOVE_VALUE) _
                      Or (ActionType And REMOVE_VALUE_IF_EMPTY) Or (ActionType And JUMP_VALUE) Then
                        If Not Reg.HasSpecialChar(sKey) And Not Reg.HasSpecialChar(sParam) Then
                            If Not Reg.ValueExists(lActualHive, sKey, sParam, Wow6432Redir) Then bNoItem = True
                        End If
                    End If
                    
                    If Not bNoItem Then

                        If AryPtr(KeyArray) Then
                        
                            'prevent duplicates
                            If Len(ReplaceDataWhat) = 0 And Len(ReplaceDataInto) = 0 And Len(TrimDelimiter) = 0 Then
                                For i = 0 To UBound(KeyArray)
                                    If sParam = KeyArray(i).Param Then
                                        If sKey = KeyArray(i).Key Then
                                            If lActualHive = KeyArray(i).Hive And ActionType = KeyArray(i).ActionType Then
                                                If Wow6432Redir = KeyArray(i).Redirected And ParamType = KeyArray(i).ParamType Then
                                                    If CStr(vDefaultData) = KeyArray(i).DefaultData Then
                                                        GoTo Continue
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        
                            ReDim Preserve KeyArray(UBound(KeyArray) + 1)
                        Else
                            ReDim KeyArray(0)
                        End If
                        
                        With KeyArray(UBound(KeyArray))
                            .ActionType = ActionType
                            .Hive = lActualHive
                            .Key = sKey
                            .Param = sParam
                            .DefaultData = CStr(vDefaultData)
                            .Redirected = Wow6432Redir
                            .ParamType = ParamType
                            .ReplaceDataWhat = ReplaceDataWhat
                            .ReplaceDataInto = ReplaceDataInto
                            .TrimDelimiter = TrimDelimiter
                        End With
                    End If
                End If
Continue:
            Next
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddRegToFix", ActionType, lHive, sKey, sParam, vDefaultData, eRedirected, ParamType, ReplaceDataWhat, ReplaceDataInto, TrimDelimiter
    If inIDE Then Stop: Resume Next
End Sub

Function CombineHives(ParamArray vHives() As Variant) As ENUM_REG_HIVE_FIX
    Dim vHive As Variant, lHive As ENUM_REG_HIVE_FIX
    
    For Each vHive In vHives
        If vHive = HKEY_CLASSES_ROOT Then lHive = lHive Or HKCR_FIX
        If vHive = HKEY_CURRENT_USER Then lHive = lHive Or HKCU_FIX
        If vHive = HKEY_LOCAL_MACHINE Then lHive = lHive Or HKLM_FIX
        If vHive = HKEY_USERS Then lHive = lHive Or HKU_FIX
    Next
    CombineHives = lHive Or &H1000&
End Function

Function ConvertHiveFixToHive(lHive As ENUM_REG_HIVE_FIX) As ENUM_REG_HIVE
    Select Case lHive
        Case HKCR_FIX: ConvertHiveFixToHive = HKEY_CLASSES_ROOT
        Case HKCU_FIX: ConvertHiveFixToHive = HKEY_CURRENT_USER
        Case HKLM_FIX: ConvertHiveFixToHive = HKEY_LOCAL_MACHINE
        Case HKU_FIX: ConvertHiveFixToHive = HKEY_USERS
    End Select
End Function

' Append results array with new ini record
Public Sub AddIniToFix( _
    KeyArray() As FIX_REG_KEY, _
    ActionType As ENUM_REG_ACTION_BASED, _
    ByVal sIniFile As String, _
    sSection As Variant, _
    Optional sParam As Variant = vbNullString, _
    Optional sDefaultData As Variant = vbNullString)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent And Not g_bFixing Then Exit Sub
    
    If Len(sIniFile) = 0 Then Exit Sub
    
    If ActionType And REMOVE_VALUE_INI Then
        If Not FileExists(sIniFile) Then Exit Sub
    End If
    
    If AryPtr(KeyArray) Then
        ReDim Preserve KeyArray(UBound(KeyArray) + 1)
    Else
        ReDim KeyArray(0)
    End If
    
    With KeyArray(UBound(KeyArray))
        .ActionType = ActionType
        .IniFile = sIniFile
        .Key = sSection
        .Param = sParam
        .DefaultData = sDefaultData
    End With

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddIniToFix", ActionType, sIniFile, sSection, sParam, sDefaultData
    If inIDE Then Stop: Resume Next
End Sub

' Add file to jump list without actually doing cure on it
' Returns 'true' if file exists and was add to jump list
Public Function AddJumpFile( _
    JumpArray() As JUMP_ENTRY, _
    ActionType As ENUM_FILE_ACTION_BASED, _
    sFilePath As String, _
    Optional sArguments As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent Then Exit Function
    
    If Len(sFilePath) = 0 Then Exit Function
    
    Dim FileFix() As FIX_FILE
    AddFileToFix FileFix, ActionType, sFilePath, sArguments

    If AryPtr(FileFix) Then
        AddJumpFile = True
    
        If AryPtr(JumpArray) Then
            ReDim Preserve JumpArray(UBound(JumpArray) + 1)
        Else
            ReDim JumpArray(0)
        End If
        With JumpArray(UBound(JumpArray))
            .Type = JUMP_ENTRY_FILE
            .File = FileFix
        End With
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "AddJumpFile", sFilePath
    If inIDE Then Stop: Resume Next
End Function

Public Function AddJumpFiles( _
    JumpArray() As JUMP_ENTRY, _
    ActionType As ENUM_FILE_ACTION_BASED, _
    aFilePath() As String) As Boolean
    
    'speed hack
    If bAutoLogSilent Then Exit Function
    
    Dim i&
    For i = 0 To UBoundSafe(aFilePath)
        AddJumpFiles = AddJumpFiles Or AddJumpFile(JumpArray, ActionType, aFilePath(i))
    Next
End Function

' Append results array with new ini record
Public Sub AddFileToFix( _
    FileArray() As FIX_FILE, _
    ActionType As ENUM_FILE_ACTION_BASED, _
    sFilePath As String, _
    Optional sArguments As String = vbNullString, _
    Optional sExpanded As String = vbNullString, _
    Optional sAutorun As String = vbNullString, _
    Optional sGoodFile As String = vbNullString)
    
    On Error GoTo ErrorHandler
    
    Dim bMissing As Boolean
    
    'speed hack
    If bAutoLogSilent And Not g_bFixing Then Exit Sub
    
    If Len(sFilePath) = 0 Then Exit Sub
    'If FileMissing(sFilePath) Then Exit Sub '!!! disabled because of 'RESTORE_FILE'
    bMissing = FileMissing(sFilePath)
    
    If Not bMissing Then
        bMissing = Not FileExists(sFilePath)
    Else
        RemoveFileMissingStr sFilePath
    End If
    
    Dim i As Long
    
    'prevent duplicates
    If AryPtr(FileArray) <> 0 Then
        If Len(sArguments) = 0 And Len(sExpanded) = 0 And Len(sAutorun) = 0 And Len(sGoodFile) = 0 Then
            For i = 0 To UBound(FileArray)
                If StrComp(sFilePath, FileArray(i).path, 1) = 0 Then
                    If ActionType = FileArray(i).ActionType Then Exit Sub
                End If
            Next
        End If
    End If
    
    'if restoring is not required
    If Not CBool(ActionType And RESTORE_FILE) And Not CBool(ActionType And RESTORE_FILE_SFC) Then
        
        'if no file to remove
        If ActionType And REMOVE_FILE Then
            If bMissing Then ActionType = ActionType And Not REMOVE_FILE
        End If
    
        'if nothing to backup
        If ActionType And BACKUP_FILE Then
            If bMissing Then ActionType = ActionType And Not BACKUP_FILE
        End If
        
        'if nothing to unreg.
        If ActionType And UNREG_DLL Then
            If bMissing Then ActionType = ActionType And Not UNREG_DLL
        End If
        
        'if no file to jump
        'If ActionType And JUMP_FILE Then
        '    If bMissing Then Exit Sub
        'End If
        
        'if no folder to remove
        If ActionType And REMOVE_FOLDER Then
            If Not FolderExists(sFilePath) Then ActionType = ActionType And Not REMOVE_FOLDER
        End If
        
        'if no folder to jump
        'If ActionType And JUMP_FOLDER Then
        '    If Not FolderExists(sFilePath) Then Exit Sub
        'End If
        
        'if folder is already exist
        If ActionType And CREATE_FOLDER Then
            If FolderExists(sFilePath) Then ActionType = ActionType And Not CREATE_FOLDER
        End If
        
        If ActionType = 0 Then ActionType = JUMP_FILE
        
    End If
    
    If AryPtr(FileArray) Then
        ReDim Preserve FileArray(UBound(FileArray) + 1)
    Else
        ReDim FileArray(0)
    End If
    
    With FileArray(UBound(FileArray))
        .ActionType = ActionType
        .path = sFilePath
        .Arguments = sArguments
        .GoodFile = sGoodFile
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddFileToFix", ActionType, sFilePath, sArguments, sExpanded, sAutorun, sGoodFile
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new process record
Public Sub AddProcessToFix( _
    ProcessArray() As FIX_PROCESS, _
    ActionType As ENUM_PROCESS_ACTION_BASED, _
    Optional PathOrName As String, _
    Optional pid As Long)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent And Not g_bFixing Then Exit Sub
    
    If Len(PathOrName) = 0 And pid = 0 Then Exit Sub
    
    If AryPtr(ProcessArray) Then
        ReDim Preserve ProcessArray(UBound(ProcessArray) + 1)
    Else
        ReDim ProcessArray(0)
    End If
    
    With ProcessArray(UBound(ProcessArray))
        .ActionType = ActionType
        .PathOrName = PathOrName
        .pid = pid
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddProcessToFix", ActionType, PathOrName, pid
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new custom record
Public Sub AddCustomToFix( _
    CustomArray() As FIX_CUSTOM, _
    ActionType As ENUM_CUSTOM_ACTION_BASED, _
    Optional sName As String, _
    Optional id As String, _
    Optional URL As String, _
    Optional sTargetOrUser As String, _
    Optional sCmd As String)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent And Not g_bFixing Then Exit Sub
    
    If AryPtr(CustomArray) Then
        ReDim Preserve CustomArray(UBound(CustomArray) + 1)
    Else
        ReDim CustomArray(0)
    End If
    
    With CustomArray(UBound(CustomArray))
        .ActionType = ActionType
        .Name = sName
        .id = id
        .URL = URL
        .TargetOrUser = sTargetOrUser
        .CommandLine = sCmd
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddCustomToFix", ActionType
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new command execution record
Public Sub AddCommandlineToFix( _
    CommandlineArray() As FIX_COMMANDLINE, _
    ActionType As ENUM_COMMANDLINE_ACTION_BASED, _
    Optional Executable As String, _
    Optional Arguments As String, _
    Optional Style As SHOWWINDOW_FLAGS, _
    Optional bWait As Boolean = True, _
    Optional TimeoutMs As Long = 30000)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent And Not g_bFixing Then Exit Sub
    
    If AryPtr(CommandlineArray) Then
        ReDim Preserve CommandlineArray(UBound(CommandlineArray) + 1)
    Else
        ReDim CommandlineArray(0)
    End If
    
    With CommandlineArray(UBound(CommandlineArray))
        .ActionType = ActionType
        .Executable = Executable
        .Arguments = Arguments
        'just in case
        If .ActionType = COMMANDLINE_POWERSHELL And Len(Arguments) = 0 And Len(Executable) <> 0 Then
            .Arguments = Executable
        End If
        .Style = Style
        .Wait = bWait
        .TimeoutMs = TimeoutMs
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddCommandlineToFix", ActionType
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new process record
Public Sub AddServiceToFix( _
    ServiceArray() As FIX_SERVICE, _
    ActionType As ENUM_SERVICE_ACTION_BASED, _
    sServiceName As String, _
    Optional sServiceDisplay As String = vbNullString, _
    Optional sImagePath As String = vbNullString, _
    Optional sDllPath As String = vbNullString, _
    Optional RunState As SERVICE_STATE, _
    Optional ForceMicrosoft As Boolean)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent And Not g_bFixing Then Exit Sub
    
    If Len(sServiceName) = 0 Then Exit Sub
    
    If AryPtr(ServiceArray) Then
        ReDim Preserve ServiceArray(UBound(ServiceArray) + 1)
    Else
        ReDim ServiceArray(0)
    End If
    
    With ServiceArray(UBound(ServiceArray))
        .ActionType = ActionType
        .ImagePath = sImagePath
        .DllPath = sDllPath
        .serviceName = sServiceName
        .ServiceDisplay = sServiceDisplay
        .RunState = RunState
        .ForceMicrosoft = ForceMicrosoft
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddServiceToFix", ActionType, sServiceName, sServiceDisplay, sImagePath, sDllPath
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new process record
Public Sub AddTaskToFix( _
    TaskArray() As FIX_TASK, _
    ActionType As ENUM_TASK_ACTION_BASED, _
    sTaskPath As String)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent And Not g_bFixing Then Exit Sub
    
    If Len(sTaskPath) = 0 Then Exit Sub
    
    If AryPtr(TaskArray) Then
        ReDim Preserve TaskArray(UBound(TaskArray) + 1)
    Else
        ReDim TaskArray(0)
    End If
    
    With TaskArray(UBound(TaskArray))
        .ActionType = ActionType
        .TaskPath = sTaskPath
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddTaskToFix", ActionType, sTaskPath
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixIt(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FixIt - Begin"
    
    If result.CureType And COMMANDLINE_BASED Then FixCommandlineHandler result
    If result.CureType And SERVICE_BASED Then FixServiceHandler result
    If result.CureType And PROCESS_BASED Then FixProcessHandler result
    If result.CureType And FILE_BASED Then FixFileHandler result
    If result.CureType And (REGISTRY_BASED Or INI_BASED) Then FixRegistryHandler result
    If result.CureType And TASK_BASED Then FixTaskHandler result
    If result.CureType And CUSTOM_BASED Then FixCustomHandler result
    
    AppendErrorLogCustom "FixIt - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixIt", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub


Public Sub FixCustomHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FixCustomHandler - Begin"
    
    Dim i As Long
    
    If result.CureType And CUSTOM_BASED Then
        If AryPtr(result.Custom) Then
            For i = 0 To UBound(result.Custom)
                With result.Custom(i)
                    Select Case .ActionType
                    
                    Case CUSTOM_ACTION_O25
                        RemoveSubscriptionWMI result.O25
                    
                    Case CUSTOM_ACTION_BITS
                        RemoveBitsJob result.Custom(i).id, bRemoveAll:=result.FixAll
                    
                    Case CUSTOM_ACTION_APPLOCKER
                        RestoreApplockerDefaults
                    
                    Case CUSTOM_ACTION_REMOVE_GROUP_MEMBERSHIP
                        RemoveUserGroupMembership result.Custom(i).TargetOrUser, result.Custom(i).Name
                    
                    Case CUSTOM_ACTION_FIREWALL_RULE
                        FW_RuleSetState result.Custom(i).Name, False
                    
                    End Select
                End With
            Next
        End If
    End If
    
    AppendErrorLogCustom "FixCustomHandler - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixCustomHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixCommandlineHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FixCommandlineHandler - Begin"
    
    Dim i As Long
    
    If result.CureType And COMMANDLINE_BASED Then
        If AryPtr(result.CommandLine) Then
            For i = 0 To UBound(result.CommandLine)
                With result.CommandLine(i)
                    Select Case .ActionType
                    
                    Case COMMANDLINE_RUN
                        If Proc.ProcessRun(.Executable, .Arguments, , .Style) Then
                            If .Wait Then
                                Proc.WaitForTerminate , , , .TimeoutMs
                            End If
                        End If
                    
                    Case COMMANDLINE_POWERSHELL
                        Proc.RunPowershell .Arguments, .Wait, .TimeoutMs, .Style
                    
                    End Select
                End With
            Next
        End If
    End If
    
    AppendErrorLogCustom "FixCommandlineHandler - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixCustomHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixProcessHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FixProcessHandler - Begin"
    
    Dim i As Long, k As Long
    Dim ActionType As ENUM_PROCESS_ACTION_BASED
    
    If result.CureType And PROCESS_BASED Then
        If AryPtr(result.Process) Then
            For i = 0 To UBound(result.Process)
                
                ActionType = result.Process(i).ActionType
                
                Dim lNumProcesses As Long
                Dim Process() As MY_PROC_ENTRY
                Dim bByPID As Boolean: bByPID = result.Process(i).pid <> 0
                Dim bByName As Boolean: bByName = InStr(1, result.Process(i).PathOrName, "\") = 0
                
                If bByPID Then
                    lNumProcesses = 1
                    ReDim Process(0)
                    Process(0).pid = result.Process(i).pid
                    
                ElseIf bByName Then
                    lNumProcesses = GetProcessesByName(Process, result.Process(i).PathOrName)
                    
                Else
                    lNumProcesses = 1
                    ReDim Process(0)
                    Process(0).path = result.Process(i).PathOrName
                End If
                
                For k = 0 To lNumProcesses - 1
                  With Process(k)
                
                    'my parent and not explorer ?
                    'Dim bParentProtected As Boolean
                    'bParentProtected = StrComp(.Path, MyParentProc.Path, 1) = 0 And Not StrEndWith(.Path, "explorer.exe")
                    
                    If Not IsSystemCriticalProcessPath(.path) Then 'And Not bParentProtected Then
                        
                        If bByPID Or bByName Then .path = "" 'should operate by PID
                        
                        If (ActionType And USE_FEATURE_DISABLE) And g_bDelmodeDisabling Then
                            Exit Sub
                        End If
                        
                        If ActionType And FREEZE_PROCESS Then
                            
                            PauseProcessByFileOrPID .path, .pid
                            
                        End If
                    
                        If ActionType And KILL_PROCESS Then
                            
                            KillProcessByFileOrPID .path, .pid, bForceMicrosoft:=True
                            
                        End If
                        
                        If ActionType And FREEZE_OR_KILL_PROCESS Then
                        
                            If Not PauseProcessByFileOrPID(.path, .pid) Then
                                KillProcessByFileOrPID .path, .pid, bForceMicrosoft:=True
                            End If
                        End If
                        
                        If ActionType And CLOSE_PROCESS Then
                        
                            ProcessCloseWindowByFileOrPID .path, .pid, bForce:=False, bWait:=True, TimeoutMs:=5000
                        
                        End If
                        
                        If ActionType And CLOSE_OR_KILL_PROCESS Then
                        
                            If Not ProcessCloseWindowByFileOrPID(.path, .pid, bForce:=False, bWait:=True, TimeoutMs:=5000) Then
                                KillProcessByFileOrPID .path, .pid, bForceMicrosoft:=True
                            End If
                        
                        End If
                        
                    End If
                  End With
                Next
            Next
        End If
    End If
    
    AppendErrorLogCustom "FixProcessHandler - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixProcessHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixRegistryHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FixRegistryHandler - Begin"
    
    Dim sData As String, i As Long
    Dim lType As REG_VALUE_TYPE
    Dim aData() As String
    Dim bDouble As Boolean
    Dim sDelim As String
    
    'Note: REG_RESTORE_SAME - is a default type if it was not specified in the argument of 'AddRegToFix'
    
    'If Result.CureType And REGISTRY_BASED Then
        If AryPtr(result.Reg) Then
            For i = 0 To UBound(result.Reg)
                With result.Reg(i)
                    
                    If (.ActionType And USE_FEATURE_DISABLE) And g_bDelmodeDisabling Then
                        Exit Sub
                    End If
                    
                    'if need to leave the same type of value
                    If .ParamType = REG_RESTORE_SAME Then
                        If Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then
                            .ParamType = MapRegValueTypeToRegRestoreType(Reg.GetValueType(.Hive, .Key, .Param, .Redirected))
                        Else
                            If Len(.DefaultData) <> 0 Then
                                If IsNumeric(.DefaultData) Then
                                    .ParamType = REG_RESTORE_DWORD
                                Else
                                    .ParamType = REG_RESTORE_EXPAND_SZ
                                End If
                            Else
                                .ParamType = REG_RESTORE_EXPAND_SZ
                            End If
                        End If
                    End If
                    
                    If .ActionType And RESTORE_VALUE Then
                        
                        Reg.DelVal .Hive, .Key, .Param, .Redirected
                        
                        Select Case .ParamType
                        
                        Case REG_RESTORE_SZ
                            Reg.SetStringVal .Hive, .Key, .Param, CStr(.DefaultData), .Redirected
                        
                        Case REG_RESTORE_EXPAND_SZ
                            Reg.SetExpandStringVal .Hive, .Key, .Param, CStr(.DefaultData), .Redirected
                        
                        Case REG_RESTORE_BINARY
                            Reg.SetBinaryVal .Hive, .Key, .Param, HexStringToArray(CStr(.DefaultData)), .Redirected
                        
                        Case REG_RESTORE_DWORD
                            Reg.SetDwordVal .Hive, .Key, .Param, CLng(.DefaultData), .Redirected
                        
                        Case REG_RESTORE_QWORD
                            Reg.SetQwordVal .Hive, .Key, .Param, CLng(.DefaultData), .Redirected
                        
                        'Case REG_RESTORE_LINK
                        
                        Case REG_RESTORE_MULTI_SZ
                            aData = SplitSafe(CStr(.DefaultData), vbNullChar)
                            Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                        
                        End Select
                    End If
                    
                    If .ActionType And CREATE_KEY Then
                    
                        Reg.CreateKey .Hive, .Key, .Redirected
                    End If
                    
                    If .ActionType And REMOVE_VALUE Then
                    
                        Reg.DelVal .Hive, .Key, .Param, .Redirected
                    End If
                    
                    If .ActionType And REMOVE_KEY Then
                    
                        Reg.DelKey .Hive, .Key, .Redirected
                    End If
                    
                    If .ActionType And REMOVE_KEY_IF_NO_VALUES Then
                    
                        If Not Reg.KeyHasValues(.Hive, .Key, .Redirected) Then
                            Reg.DelKey .Hive, .Key, .Redirected
                        End If
                    End If
                    
                    If .ActionType And RESTORE_VALUE_INI Then
                    
                        IniSetString .IniFile, .Key, .Param, CStr(.DefaultData)
                    End If
                        
                    If .ActionType And REMOVE_VALUE_INI Then
                        
                        IniRemoveString .IniFile, .Key, .Param
                    End If
                    
                    If .ActionType And APPEND_VALUE_NO_DOUBLE Then
                    
                        'check if value already contains data planned to be written
                        bDouble = False
                        sDelim = vbNullString
                        If Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then
                        
                            sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected, True)) 'true - do not expand
                            
                            If .ParamType = REG_RESTORE_MULTI_SZ Then
                                sDelim = vbNullChar
                            ElseIf Len(.TrimDelimiter) <> 0 Then
                                sDelim = .TrimDelimiter
                            End If
                            If Len(sDelim) <> 0 Then
                                bDouble = inArraySerialized(CStr(.DefaultData), sData, sDelim, , , vbTextCompare)
                            End If
                        End If
                        
                        If Not bDouble Then
                            'adding data to the beginning of the value
                            If Len(sData) = 0 Or Len(TrimEx(sData, sDelim)) = 0 Then
                                sData = .DefaultData
                            Else
                                sData = .DefaultData & sDelim & sData
                            End If
                            
                            Select Case .ParamType
                            
                            Case REG_RESTORE_MULTI_SZ
                                aData = SplitSafe(sData, vbNullChar)
                                Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                                
                            Case REG_RESTORE_SZ
                                Reg.SetStringVal .Hive, .Key, .Param, sData, .Redirected
                                
                            Case REG_RESTORE_EXPAND_SZ
                                Reg.SetExpandStringVal .Hive, .Key, .Param, sData, .Redirected
                                
                            End Select
                        End If
                    End If
                    
                    If .ActionType And REPLACE_VALUE Then
                    
                        If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                    
                        sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                        
                        If .ActionType And TRIM_VALUE Then 'if further 'trim' is planned to be in action
                            'replace by exact value
                            sData = Replace$(.TrimDelimiter & sData & .TrimDelimiter, _
                                .TrimDelimiter & .ReplaceDataWhat & .TrimDelimiter, _
                                .TrimDelimiter & .ReplaceDataInto & .TrimDelimiter, 1, 1, vbBinaryCompare) 'restrict to maximum 1 replacing
                        Else
                            'replace with part of value
                            sData = Replace$(sData, .ReplaceDataWhat, .ReplaceDataInto, 1, 1, vbTextCompare) 'restrict to maximum 1 replacing
                        End If
                        
                        lType = Reg.GetValueType(.Hive, .Key, .Param, .Redirected)
                        
                        Select Case lType
                        
                        Case REG_SZ
                            Reg.SetStringVal .Hive, .Key, .Param, sData, .Redirected
                                
                        Case REG_EXPAND_SZ
                            Reg.SetExpandStringVal .Hive, .Key, .Param, sData, .Redirected
                        
                        Case REG_MULTI_SZ
                            aData = SplitSafe(sData, vbNullChar)
                            Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                            
                        End Select
                    End If
                    
                    If .ActionType And TRIM_VALUE Then
                        
                        If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                        
                        sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                        
                        sData = TrimEx(sData, .TrimDelimiter)
                        
                        sData = Replace$(sData, .TrimDelimiter & .TrimDelimiter, .TrimDelimiter) '2 delims -> 1 delim
                        
                        lType = Reg.GetValueType(.Hive, .Key, .Param, .Redirected)
                        
                        Select Case lType
                        
                        Case REG_SZ
                            Reg.SetStringVal .Hive, .Key, .Param, sData, .Redirected
                                
                        Case REG_EXPAND_SZ
                            Reg.SetExpandStringVal .Hive, .Key, .Param, sData, .Redirected
                        
                        Case REG_MULTI_SZ
                            aData = SplitSafe(sData, vbNullChar)
                            Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                        
                        End Select
                    End If
                    
                    If .ActionType And REMOVE_VALUE_IF_EMPTY Then
                        
                        If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                        
                        sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                        If Len(sData) = 0 Then
                            Reg.DelVal .Hive, .Key, .Param, .Redirected
                        End If
                    End If
                    
                    If .ActionType And RESTORE_KEY_PERMISSIONS Then
                        RegKeyResetDACL .Hive, .Key, .Redirected, False
                    End If
                    
                    If .ActionType And RESTORE_KEY_PERMISSIONS_RECURSE Then
                        RegKeyResetDACL .Hive, .Key, .Redirected, True
                    End If

                End With
            Next
        End If
    'End If
    
    AppendErrorLogCustom "FixRegistryHandler - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixRegistryHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Function MapRegValueTypeToRegRestoreType(ordType As REG_VALUE_TYPE, Optional vData As Variant) As ENUM_REG_VALUE_TYPE_RESTORE
    On Error GoTo ErrorHandler
    
    Dim bRequiredDefault As Boolean
    Dim rType As ENUM_REG_VALUE_TYPE_RESTORE
    
    Select Case ordType
    
    Case REG_NONE
        bRequiredDefault = True
        
    Case REG_SZ
        rType = REG_RESTORE_SZ
    
    Case REG_EXPAND_SZ
        rType = REG_RESTORE_EXPAND_SZ
    
    Case REG_BINARY
        bRequiredDefault = True
        
    Case REG_DWORD
        rType = REG_RESTORE_DWORD
    
    Case REG_DWORDLittleEndian
        rType = REG_RESTORE_DWORD
    
    Case REG_DWORDBigEndian
        rType = REG_RESTORE_DWORD
    
    Case REG_LINK
        bRequiredDefault = True
        
    Case REG_MULTI_SZ
        rType = REG_RESTORE_MULTI_SZ
        
    Case REG_ResourceList
        bRequiredDefault = True
        
    Case REG_FullResourceDescriptor
        bRequiredDefault = True
        
    Case REG_ResourceRequirementsList
        bRequiredDefault = True
        
    Case REG_QWORD
        rType = REG_RESTORE_QWORD
        
    Case REG_QWORD_LITTLE_ENDIAN
        rType = REG_RESTORE_QWORD
    
    Case Else
        bRequiredDefault = True
    End Select
    
    If bRequiredDefault Then
        If IsMissing(vData) Then
            rType = REG_RESTORE_EXPAND_SZ 'default restore value type
        Else
            If IsNumeric(vData) Then
                rType = REG_RESTORE_DWORD
            Else
                rType = REG_RESTORE_EXPAND_SZ
            End If
        End If
    End If
    
    MapRegValueTypeToRegRestoreType = rType

    Exit Function
ErrorHandler:
    ErrorMsg Err, "MapRegValueTypeToRegRestoreType", ordType, vData
    If inIDE Then Stop: Resume Next
End Function

Public Sub FixFileHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FixFileHandler - Begin"
    
    Dim i As Long
    
    If result.CureType And FILE_BASED Then
        If AryPtr(result.File) Then
            For i = 0 To UBound(result.File)
                With result.File(i)
                    
                    If (.ActionType And USE_FEATURE_DISABLE) And g_bDelmodeDisabling Then
                        Exit Sub
                    End If
                
                    If .ActionType And UNREG_DLL Then
                        If Not IsMicrosoftFile(.path, True) Or result.ForceMicrosoft Then
                            Reg.UnRegisterDll .path
                        End If
                    End If
                    
                    If .ActionType And REMOVE_FILE Then
                        If FileExists(.path) Then
                            DeleteFileEx .path, result.ForceMicrosoft
                        End If
                    End If
                    
                    If .ActionType And REMOVE_FOLDER Then
                        If FolderExists(.path) Then
                            DeleteFolderForce .path, result.ForceMicrosoft
                        End If
                    End If
                    
                    If .ActionType And RESTORE_FILE Then
                        If FileExists(.GoodFile) Then
                            '// TODO: PendingFileOperation with replacing
                            If DeleteFileEx(.path, True, True) Then
                                FileCopyW .GoodFile, .path, True
                            End If
                        End If
                    End If
                    
                    If .ActionType And RESTORE_FILE_SFC Then
                        SFC_RestoreFile .path
                    End If
                    
                    If .ActionType And CREATE_FOLDER Then
                        MkDirW .path
                    End If
                End With
            Next
        End If
    End If
    
    AppendErrorLogCustom "FixFileHandler - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixFileHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Function SFC_RestoreFile(sHijacker As String, Optional bAsync As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "SFC_RestoreFile - Begin", "File: ", sHijacker
    Dim SFC As String
    Dim sHashOld As String
    Dim sHashNew As String
    Dim bNoOldFile As Boolean
    If OSver.IsWin64 And FolderExists(sWinDir & "\sysnative") Then 'Vista+
        'sfc.exe
        SFC = EnvironW("%SystemRoot%") & "\sysnative\" & Caes_Decode("tih.nIr")
    Else
        SFC = EnvironW("%SystemRoot%") & "\System32\" & Caes_Decode("tih.nIr")
    End If
    If FileExists(SFC) Then
        bNoOldFile = Not FileExists(sHijacker)
        If Not bNoOldFile Then
            TryUnlock sHijacker
            sHashOld = GetFileSHA1(sHijacker, , True)
        End If
        If Proc.ProcessRun(SFC, "/SCANFILE=" & """" & sHijacker & """", , 0) Then
            If Not bAsync Then
                If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 15000) Then
                    Proc.ProcessClose , , True
                End If
                If FileExists(sHijacker) Then
                    sHashNew = GetFileSHA1(sHijacker, , True)
                    If (sHashOld <> sHashNew) Or bNoOldFile Then
                        SFC_RestoreFile = True
                    End If
                End If
            Else
                SFC_RestoreFile = True
            End If
        End If
    End If
    AppendErrorLogCustom "SFC_RestoreFile - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SFC_RestoreFile", sHijacker
    If inIDE Then Stop: Resume Next
End Function

Public Sub FixServiceHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FixServiceHandler - Begin"
    
    Dim i As Long, j As Long, k As Long
    Dim aService() As String
    Dim aDepend() As String
    Dim bFixReg As Boolean
    
    If result.CureType And SERVICE_BASED Then
        If AryPtr(result.Service) Then
            For i = 0 To UBound(result.Service)
                With result.Service(i)
                
                    If (.ActionType And USE_FEATURE_DISABLE) And g_bDelmodeDisabling And (.ActionType And DELETE_SERVICE) Then
                        .ActionType = DISABLE_SERVICE
                    End If
                
                    If .ActionType And DELETE_SERVICE Then
                    
                        DeleteNTService .serviceName, , .ForceMicrosoft
                        
                        'Remove dependency
                        For j = 1 To Reg.EnumSubKeysToArray(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services", aService())
                            
                            aDepend = Reg.GetMultiSZ(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & aService(j), "DependOnService")
                            
                            If AryItems(aDepend) Then
                                For k = 0 To UBound(aDepend)
                                    If StrComp(aDepend(k), .serviceName, 1) = 0 Then
                                        BackupKey result, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & aService(j), "DependOnService"
                                        
                                        AddRegToFix result.Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                            HKLM, "System\CurrentControlSet\Services\" & aService(j), "DependOnService", , , REG_RESTORE_MULTI_SZ, _
                                            .serviceName, vbNullString, vbNullChar
                                        
                                        result.CureType = result.CureType Or REGISTRY_BASED
                                        bFixReg = True
                                    End If
                                Next
                            End If
                        Next
                        If bFixReg Then
                            FixRegistryHandler result
                        End If
                    End If
                    
                    If .ActionType And DISABLE_SERVICE Then
                    
                        SetServiceStartMode .serviceName, SERVICE_MODE_DISABLED
                    End If
                    
                    If .ActionType And RESTORE_SERVICE Then
                        '// TODO
                    End If
                    
                    If .ActionType And ENABLE_SERVICE Then
                        
                        SetServiceStartMode .serviceName, SERVICE_MODE_AUTOMATIC
                    End If
                    
                    If .ActionType And MANUAL_SERVICE Then
                        
                        SetServiceStartMode .serviceName, SERVICE_MODE_MANUAL
                    End If
                    
                    If .ActionType And STOP_SERVICE Then
                        
                        StopService .serviceName
                    End If
                    
                    If .ActionType And START_SERVICE Then
                        
                        StartService .serviceName
                    End If

                End With
            Next
        End If
    End If
    
    AppendErrorLogCustom "FixServiceHandler - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixServiceHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixTaskHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FixTaskHandler - Begin"
    
    Dim i As Long, j As Long, k As Long
    Dim aTask() As String
    
    If result.CureType And TASK_BASED Then
        If AryPtr(result.Task) Then
            For i = 0 To UBound(result.Task)
                With result.Task(i)
                
                    If .ActionType And ENABLE_TASK Then
                        EnableTask .TaskPath
                    End If
                    
                    If .ActionType And DISABLE_TASK Then
                        DisableTask .TaskPath
                    End If
                    
                End With
            Next
        End If
    End If
    
    AppendErrorLogCustom "FixTaskHandler - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixServiceHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Function CheckIntegrityHJT() As Boolean
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "CheckIntegrityHJT - Begin"
    'Checking consistency of HiJackThis.exe
    Dim SignResult As SignResult_TYPE
    Dim dModif As Date
    CheckIntegrityHJT = True
    If Not inIDE Then
        If OSver.IsWindows7OrGreater Then 'because my signature hash is SHA2
            If Not (OSver.SPVer < 1 And (OSver.MajorMinor = 6.1)) Then
                'ensure EDS subsystem is working correctly
                If isEDS_Work() Then
                    SignVerify AppPath(True), SV_PreferInternalSign Or SV_AllowExpired, SignResult
                    If Not IsDragokasSign(SignResult) Then
                        'not a developer machine ?
                        dModif = GetFileDate(AppPath(True), DATE_MODIFIED)
                        If (GetDateAtMidnight(dModif) <> GetDateAtMidnight(Now()) And InStr(AppPath(), Caes_Decode("`D[a")) = 0) Then '_AVZ
                            If (DateDiff("n", Now(), dModif) > 10) Or (DateDiff("n", Now(), dModif) < 0) Then
                                CheckIntegrityHJT = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    AppendErrorLogCustom "CheckIntegrityHJT - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CheckIntegrityHJT"
    If inIDE Then Stop: Resume Next
End Function

Public Function SetTaskBarProgressValue(frm As Form, ByVal Value As Single) As Boolean
    If Value < 0 Or Value > 1 Then Exit Function
    If Not (TaskBar Is Nothing) Then
        If Value = 0 Then
            TaskBar.SetProgressState g_HwndMain, TBPF_NOPROGRESS
        Else
            TaskBar.SetProgressValue frm.hWnd, CCur(Value * 10000), CCur(10000)
        End If
    End If
End Function

'// compare self version with installed one and update it
Public Function CheckInstalledVersionHJT() As Boolean
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "InstallUpdatedHJT - Begin"
    
    Dim sInstVer As String
    Dim HJT_LocationExe As String
    HJT_LocationExe = GetFullPathForInstallationHJT()
    
    'if self
    If StrComp(AppPath(True), HJT_LocationExe, 1) = 0 Then Exit Function
    
    'Installed version is present?
    If IsInstalledHJT() Then
        'compare versions
        sInstVer = GetFilePropVersion(HJT_LocationExe)
        
        Dbg "Installed version: " & sInstVer
        Dbg "Running version: " & AppVerString
        
        If ConvertVersionToNumber(sInstVer) < ConvertVersionToNumber(AppVerString) Then
            'The version of HiJackThis you launched is newer than installed one. Update it?
            If MsgBoxW(Translate(1402), vbQuestion Or vbYesNo) = vbYes Then
                'Replace 'Program Files' version with me
                CheckInstalledVersionHJT = InstallHJT(False)
            End If
        End If
    End If
    
    AppendErrorLogCustom "InstallUpdatedHJT - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "InstallUpdatedHJT"
    If inIDE Then Stop: Resume Next
End Function

Public Function InstallHJT( _
    Optional bAskToCreateDesktopShortcut As Boolean, _
    Optional bSilentCreateDesktopShortcut As Boolean) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim HJT_LocationExe As String
    Dim HJT_LocationDir As String
    Dim sScanToolsDir As String
    Dim sScanToolsDirDest As String
    Dim sHelperAppsDir As String
    Dim sHelperAppsDirDest As String
    Dim bInstInPlace As Boolean
    Dim hFile As Long
    Dim aEXE() As String
    Dim i As Long
    
    AppendErrorLogCustom "InstallHJT - Begin"
    
    HJT_LocationDir = GetDirForInstallationHJT()
    HJT_LocationExe = GetFullPathForInstallationHJT()
    
    sScanToolsDir = BuildPath(AppPath(), "tools\Scan")
    sScanToolsDirDest = BuildPath(HJT_LocationDir, "tools\Scan")
    
    sHelperAppsDir = BuildPath(AppPath(), "apps")
    sHelperAppsDirDest = BuildPath(HJT_LocationDir, "apps")
    
    If StrComp(HJT_LocationDir, AppPath(), 1) = 0 Then
        bInstInPlace = True
    End If
    
    If Not bInstInPlace Then
        If Not MkDirW(HJT_LocationDir) Then
            MsgBoxW "Installation failed. Cannot create folder: " & HJT_LocationDir, vbCritical
            InstallHJT = False
            Exit Function
        End If
        'Copy exe to Program Files dir
        If Not FileCopyW(AppPath(True), HJT_LocationExe, True) Then
            'MsgBoxW "Error while installing HiJackThis to program files folder. Cannot copy. Error = " & Err.LastDllError, vbCritical
            MsgBoxW Translate(593) & " " & Err.LastDllError, vbCritical
            InstallHJT = False
            Exit Function
        End If
        
        If FolderExists(sScanToolsDir) Then
            MkDirW sScanToolsDirDest
            FileCopyW BuildPath(sScanToolsDir, "auto.exe"), BuildPath(sScanToolsDirDest, "auto.exe")
            FileCopyW BuildPath(sScanToolsDir, "auto64.exe"), BuildPath(sScanToolsDirDest, "auto64.exe")
            FileCopyW BuildPath(sScanToolsDir, "executed.exe"), BuildPath(sScanToolsDirDest, "executed.exe")
            FileCopyW BuildPath(sScanToolsDir, "lastactivity.exe"), BuildPath(sScanToolsDirDest, "lastactivity.exe")
            FileCopyW BuildPath(sScanToolsDir, "serwin.exe"), BuildPath(sScanToolsDirDest, "serwin.exe")
            FileCopyW BuildPath(sScanToolsDir, "sheduler.exe"), BuildPath(sScanToolsDirDest, "sheduler.exe")
        End If
        
        If FolderExists(sHelperAppsDir) Then
            MkDirW sHelperAppsDirDest
            CopyFolderContents sHelperAppsDir, sHelperAppsDirDest
        End If
    End If
    
    If (FileExists(BuildPath(AppPath(), "whitelists.txt"))) Then
    
        If Not bInstInPlace Then
            FileCopyW BuildPath(AppPath(), "whitelists.txt"), BuildPath(HJT_LocationDir, "whitelists.txt")
        End If
        
        'Add HJT and supporting tools to exclude
        If OpenW(BuildPath(HJT_LocationDir, "whitelists.txt"), FOR_READ_WRITE, hFile) Then
            
            aEXE = ListFiles(BuildPath(HJT_LocationDir, "tools"), ".exe", True)
            If (AryPtr(aEXE)) Then
                For i = 0 To UBound(aEXE)
                    PrintLineW hFile, aEXE(i)
                Next
            End If
            PrintLineW hFile, AppPath(True)
            CloseW hFile
        End If
    End If
    
    'create Control panel -> 'Uninstall programs' entry
    CreateUninstallKey True, HJT_LocationExe
    
    'Shortcuts in Start Menu
    InstallHJT = CreateHJTShortcuts(HJT_LocationExe)
    
    If Not HasCommandLineKey("noShortcuts") Then
        If bAskToCreateDesktopShortcut Then
            If bSilentCreateDesktopShortcut Then
                CreateHJTShortcutDesktop HJT_LocationExe
            Else
                'Installation is completed. Do you want to create shortcut in Desktop?
                If MsgBoxW(Translate(69), vbYesNo, g_AppName) = vbYes Then
                    CreateHJTShortcutDesktop HJT_LocationExe
                End If
            End If
        End If
    End If
    
    InstallHJT = True
    
    AppendErrorLogCustom "InstallHJT - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "InstallHJT"
    If inIDE Then Stop: Resume Next
End Function

Public Function InstallAutorunHJT( _
    Optional bSilent As Boolean, _
    Optional lDelay As Long = 60, _
    Optional bForceCreateDesktopShortcuts As Boolean) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim HJT_Location As String
    Dim HJT_Command As String
    Dim Delay As String
    Dim sArguments As String
    Dim pos As Long
    
    Delay = CStr(lDelay)
    
    AppendErrorLogCustom "InstallAutorunHJT - Begin"
    
    HJT_Location = BuildPath(GetDirForInstallationHJT(), AppExeName(True))
    
'    If MsgBox("This will install HJT to 'Program Files' folder and set Windows for automatically run HJT scan at system startup." & _
'        vbCrLf & vbCrLf & "Continue?" & vbCrLf & vbCrLf & "Note: it is recommended that you add all safe items to ignore list, so " & _
'        "the results window will appear at system startup if only new item will be found.", vbYesNo Or vbQuestion) = vbNo Then
    If Not bSilent Then
        If MsgBoxW(Translate(66), vbYesNo Or vbQuestion) = vbNo Then
            gNotUserClick = True
            frmMain.chkConfigStartupScan.Value = 0
            gNotUserClick = False
            Exit Function
        End If
        'To increase system loading speed it is recommended to set a delay
        'before launching HiJackThis on user logon. Specify the delay (in seconds):
        Delay = InputBox(Translate(1403), g_AppName, "60")
        If Not IsNumeric(Delay) Then Delay = "60"
        If CLng(Delay) < 0 Then Delay = "60"
    End If
    
    If OSver.IsWindowsVistaOrGreater Then
        'check if 'Schedule' service is launched
        If Not RunScheduler_Service(True, Not bSilent, bSilent) Then
            Exit Function
        End If
    End If
    
    pos = InStr(1, g_sCommandLine, "/!")
    
    If pos = 0 Then
        sArguments = "/startupscan"
    Else
        sArguments = mid$(g_sCommandLine, pos + 3)
    End If
    
    If InstallHJT(bForceCreateDesktopShortcuts, HasCommandLineKey("noGUI") Or bForceCreateDesktopShortcuts) Then
    
        If OSver.IsWindowsVistaOrGreater Then
        
'            'delay after system startup for 1 min.
'            JobCommand = "/create /tn ""HiJackThis Autostart Scan"" /SC ONSTART /DELAY 0001:00 /F /RL HIGHEST " & _
'                "/tr ""\""" & HJT_Location & "\"" /startupscan"""
'
'            If Proc.ProcessRun("schtasks.exe", JobCommand, , 0) Then
'                iExitCode = Proc.WaitForTerminate(, , , 15000)     'if ExitCode = 0, 15 sec for timeout
'                If ERROR_SUCCESS <> iExitCode Then
'                    Proc.ProcessClose , , True
'                    'MsgBoxW "Error while creating task. Error = " & iExitCode, vbCritical
'                    MsgBoxW Translate(594) & " " & iExitCode, vbCritical
'                Else
'                    InstallAutorunHJT = True
'                End If
'            End If
            
            If CreateTask("HiJackThis Autostart Scan", HJT_Location, sArguments, _
              "Automatically scan the system with HiJackThis at user logon", CLng(Delay)) Then
                InstallAutorunHJT = True
            Else
                'MsgBoxW "Error while creating task", vbCritical
                MsgBoxW Translate(594), vbCritical
            End If
        Else
            'XP-
            'to add to 'Run' registry key
            HJT_Command = """" & HJT_Location & """" & " " & sArguments
            InstallAutorunHJT = Reg.SetExpandStringVal(HKLM, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis Autostart Scan", HJT_Command)
        End If
    End If
    
    AppendErrorLogCustom "InstallAutorunHJT - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "InstallAutorunHJT"
    If inIDE Then Stop: Resume Next
End Function

Public Function RemoveAutorunHJT() As Boolean
    If OSver.IsWindowsVistaOrGreater Then
        RemoveAutorunHJT = KillTask2("\HiJackThis Autostart Scan")
    Else
        RemoveAutorunHJT = Reg.DelVal(HKLM, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis Autostart Scan")
    End If
End Function

Public Sub OpenFolder(sFolder As String)
    Shell sWinDir & "\explorer.exe " & """" & sFolder & """", vbNormalFocus
End Sub

Public Sub OpenAndSelectFile(sFile As String)
    On Error GoTo ErrorHandler:

    Dim hRet As Long
    Dim pidl As Long
    Dim sFileIDL As String
    
    If OSver.MajorMinor >= 5.1 Then '(XP+)
    
        sFileIDL = sFile
        
        'commented because such pIDL doesn't work on SHOpenFolderAndSelectItems
'        If OSver.IsWin64 Then
'            If StrBeginWith(sFileIDL, sWinSysDir) Then
'                sFileIDL = Replace$(sFileIDL, sWinSysDir, sWinDir & "\sysnative", 1, 1, 1)
'            End If
'        End If
    
        pidl = ILCreateFromPath(StrPtr(sFileIDL))

        If pidl <> 0 Then
            hRet = SHOpenFolderAndSelectItems(pidl, 0, 0, 0)
            
            CoTaskMemFree pidl
        End If
    End If
    
    If pidl = 0 Or hRet <> S_OK Then
        'alternate
        If OSver.IsWin64 Then 'fix for 64 bit
            Shell sWinDir & "\explorer.exe /select," & """" & sFile & """", vbNormalFocus
        Else
            Shell "explorer.exe /select," & """" & sFile & """", vbNormalFocus
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "OpenAndSelectFile", "File:", sFile
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetDateAtMidnight(dDate As Date) As Date
    GetDateAtMidnight = DateAdd("s", -Second(dDate), DateAdd("n", -Minute(dDate), DateAdd("h", -Hour(dDate), dDate)))
End Function

Public Sub HJT_SaveReport(Optional nTry As Long)
    On Error GoTo ErrorHandler:
    Dim idx&
    
    AppendErrorLogCustom "HJT_SaveReport - Begin"

    idx = 7
    
    If bAutoLog Then
        If Len(g_sLogFile) = 0 Then
            g_sLogFile = BuildPath(AppPath(), "HiJackThis.log")
        End If
    Else
        bGlobalDontFocusListBox = True
        'sLogFile = SaveFileDialog("Save logfile...", "Log files (*.log)|*.log|All files (*.*)|*.*", "HiJackThis.log")
        g_sLogFile = SaveFileDialog(Translate(1001), AppPath(), "HiJackThis.log", Translate(1002) & " (*.log)|*.log|" & Translate(1003) & " (*.*)|*.*")
        bGlobalDontFocusListBox = False
    End If
    
    idx = 8
    
    If 0 <> Len(g_sLogFile) Then
        
        idx = 11
        
        Dim b() As Byte
        
        b = CreateLogFile() '<<<<<< ------- preparing all text for log file
        
        idx = 13
        
        'in /silentautolog mode log handle is already opened
        
        If g_hLog <= 0 Then
            If Not OpenW(g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog, g_FileBackupFlag) Then

                If Not bAutoLogSilent Then 'not via AutoLogger
                    'try another name

                    g_sLogFile = Left$(g_sLogFile, Len(g_sLogFile) - 4) & "_2.log"

                    Call OpenW(g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog)
                End If
            End If
        End If
        
        If g_hLog <= 0 Then
            If bAutoLogSilent Then 'via AutoLogger
                Exit Sub
            Else
            
                If bAutoLog Then ' if user clicked 1-st button (and HJT on ReadOnly media) => try another folder
                
                  Do
                    bGlobalDontFocusListBox = True
                    'sLogFile = SaveFileDialog("Save logfile...", "Log files (*.log)|*.log|All files (*.*)|*.*", "HiJackThis.log")
                    g_sLogFile = SaveFileDialog(Translate(1001), AppPath(), "HiJackThis.log", Translate(1002) & " (*.log)|*.log|" & Translate(1003) & " (*.*)|*.*")
                    bGlobalDontFocusListBox = False
                    
                    If Len(g_sLogFile) = 0 Then Exit Sub
                    
                    If Not OpenW(g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog) Then    '2-nd try
                        MsgBoxW Translate(26), vbExclamation
                    End If
                  
                  Loop While (g_hLog <= 0)
                    
                Else 'if user already clicked button "Save report"
                
'                   msgboxW "Write access was denied to the " & _
'                       "location you specified. Try a " & _
'                       "different location please.", vbExclamation
                    MsgBoxW Translate(26), vbExclamation
                    Exit Sub
                End If
            End If
        End If
        
        PutW g_hLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
        
        Dim lret As Long
        Dim ov As OVERLAPPED
        ov.offset = 0
        ov.InternalHigh = 0
        ov.hEvent = 0
        
        If g_LogLocked Then
            lret = UnlockFileEx(g_hLog, 0&, 1& * 1024 * 1024, 0&, VarPtr(ov))
            
            If lret Then
                g_LogLocked = False
            Else
                AppendErrorLogCustom "UnlockFileEx is failed with err = " & Err.LastDllError
            End If
        End If
        
        FlushFileBuffers g_hLog
        CloseW g_hLog: g_hLog = 0
        
        'Check the size of the log
        If 0 = FileLenW(g_sLogFile) Then
            If nTry <> 2 Then
                SleepNoLock 100
                DeleteFileEx g_sLogFile
                SleepNoLock 400
                HJT_SaveReport 2
                Exit Sub
            End If
        Else 'success
            'We doing it to be able detect reliably by external applications if logfile created successfully,
            'without requirement to check for zero size.
            'So, the usual logfile is appearing at the very last moment instantly with non-zero size.
            If StrEndWith(g_sLogFile, "HiJackThis_.log") Then
                Dim sNewFile As String
                sNewFile = BuildPath(GetParentDir(g_sLogFile), "HiJackThis.log")
                If RenameFile(g_sLogFile, sNewFile, bOverwrite:=True) Then
                    g_sLogFile = sNewFile
                End If
            End If
        End If
        
        idx = 14
        
        If (Not bAutoLogSilent) Or inIDE Then OpenLogFile g_sLogFile
    End If
    
    AppendErrorLogCustom "HJT_SaveReport - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "HJT_SaveReport", "Stady: ", idx
    If inIDE Then Stop: Resume Next
End Sub

' Opens log file in default editor / or notepad if editor is not assigned to the extension
' / or in explorer window with selection if all other methods failed
'
Public Sub OpenLogFile(ByVal sLogFile As String)
    If Not FileExists(sLogFile, , True) Then Exit Sub
    
    Dim bAssoc As Boolean
    Dim bFailed As Boolean
    Dim sClassID As String
    Dim sOpenCmd As String
    Dim sOpenProg As String
    
    sLogFile = PathX64(sLogFile)
    
    sClassID = Reg.GetString(HKEY_CLASSES_ROOT, GetExtensionName(sLogFile), vbNullString)
    
    If Len(sClassID) <> 0 Then
        sOpenCmd = EnvironW(Reg.GetString(HKEY_CLASSES_ROOT, sClassID & "\shell\open\command", vbNullString))
        
        SplitIntoPathAndArgs sOpenCmd, sOpenProg, , True
        
        If FileExists(sOpenProg) Then
            bAssoc = True
        End If
    End If

    If bAssoc Then
        If OSver.IsWindowsXPOrGreater Then
            If Proc.ProcessRunUnelevated2(BuildPath(sWinDir, "explorer.exe"), sLogFile) Then Exit Sub
        End If
    End If
    
    If bAssoc Then
        bFailed = ShellExecute(g_HwndMain, StrPtr("open"), StrPtr(sLogFile), 0&, 0&, 1) <= 32
    End If
    
    If Not bAssoc Or bFailed Then
        'system doesn't know what .log is
        If FileExists(sWinDir & "\notepad.exe") Then
            ShellExecute g_HwndMain, StrPtr("open"), StrPtr(sWinDir & "\notepad.exe"), StrPtr(sLogFile), 0&, 1
        Else
            If FileExists(sWinDir & IIf(bIsWinNT, "\system32", "\system") & "\notepad.exe") Then
                ShellExecute g_HwndMain, StrPtr("open"), StrPtr(sWinDir & IIf(bIsWinNT, "\sytem32", "\system") & "\notepad.exe"), StrPtr(sLogFile), 0&, 1
            Else
                'MsgBoxW Replace$(Translate(27), "[]", sLogFile), vbInformation
'                        msgboxW "The logfile has been saved to " & sLogFile & "." & vbCrLf & _
'                               "You can open it in a text editor like Notepad.", vbInformation
            
                OpenAndSelectFile sLogFile
            End If
        End If
    End If
End Sub

Public Sub HJT_Shutdown()   ' emergency exits the program due to exceeding the timeout limit
    
    '!!! HiJackThis was shut down due to exceeding the maximum allowed timeout: [] sec. !!! Report file will be incomplete!
    'Please, restart the program manually (not via Autologger).
    ErrReport = ErrReport & vbCrLf & Replace$(Translate(1027), "[]", Perf.MAX_TimeOut)
    
    Dim s$
    If g_hDebugLog <> 0 Then
        s = vbCrLf & vbCrLf & String$(39, "=") & vbCrLf & "!!! WARNING !!! Timeout is detected !!!" & vbCrLf & String$(39, "=") & vbCrLf & vbCrLf
        PutW_NoLog g_hDebugLog, 1, StrPtr(s), LenB(s), True
    End If
    
    'CloseW hLog, True
    'DeleteFileW BuildPath(AppPath(), "HiJackThis.log")
    
    'SortSectionsOfResultList
    HJT_SaveReport
    
    If inIDE Then
        Unload frmMain
        If inIDE Then Debug.Print "HJT_Shutdown is raised!"
        End
    Else
        ExitProcess 1001&
    End If
End Sub

Public Function WhiteListed(sFile As String, sWhiteListedPath As String, Optional bCheckFileNamePartOnly As Boolean) As Boolean
    'to check matching the file with the specified name and verify it by EDS

    If bHideMicrosoft And Not bIgnoreAllWhitelists Then
        If bCheckFileNamePartOnly Then
            If StrComp(GetFileName(sFile, True), sWhiteListedPath, 1) = 0 Then
                If IsMicrosoftFile(sFile) Then WhiteListed = True
            End If
        Else
            If StrComp(sFile, sWhiteListedPath, 1) = 0 Then
                If IsMicrosoftFile(sFile) Then WhiteListed = True
            End If
        End If
    End If
End Function

Public Function SplitByMultiDelims(ByVal sLine As String, bFirstMatchOnly As Boolean, s_out_UsedDelim As String, ParamArray delims()) As String()
    On Error GoTo ErrorHandler
    Dim i As Long
    If Not bFirstMatchOnly Then
        'replace all delimiters by first one
        For i = 1 To UBound(delims)
            sLine = Replace$(sLine, delims(i), delims(0))
        Next
        s_out_UsedDelim = delims(0)
        SplitByMultiDelims = SplitSafe(sLine, CStr(delims(0)))
    Else
        For i = 0 To UBound(delims)
            'substitute each delimiter
            If InStr(sLine, delims(i)) <> 0 Then
                SplitByMultiDelims = SplitSafe(sLine, CStr(delims(i)))
                s_out_UsedDelim = delims(i)
                Exit Function
            End If
        Next
        'if no delimiters found, set initial string
        Dim arr(0) As String
        arr(0) = sLine
        SplitByMultiDelims = arr
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SplitByMultiDelims", sLine
    ReDim SplitByMultiDelims(0)
    If inIDE Then Stop: Resume Next
End Function

' adds some warning to the end of the log (before debugging info)
Public Sub AddWarning(sMsg As String)
    If bSkipErrorMsg Then Exit Sub
    EndReport = EndReport & vbCrLf & "Warning: " & sMsg
End Sub

'"Welcome to HJT" / or "Below are the results..." depending on "txtNothing" label vision state, availability of lstResults records and progressbar.
Public Sub pvSetVisionForLabelResults()
    If isRanHJT_Scan Then
        frmMain.lblInfo(0).Visible = False
        If Not frmMain.shpBackground.Visible Then
            ResumeProgressbar
        End If
    Else
        If frmMain.txtNothing.Visible Or frmMain.lstResults.ListCount <> 0 Then
            frmMain.lblInfo(1).Visible = True
            frmMain.lblInfo(0).Visible = False
        Else
            frmMain.lblInfo(1).Visible = False
            frmMain.lblInfo(0).Visible = True
        End If
    End If
End Sub

Public Function LenSafe(var As Variant) As Long
    If IsMissing(var) Then
        LenSafe = 0
    Else
        LenSafe = Len(CStr(var))
    End If
End Function

Public Function LoadResString(idfrom As Long, Optional idTo As Long) As String
    If idTo = 0 Then
        LoadResString = LoadResData(idfrom, 6)
    Else
        Dim i&, s$
        For i = idfrom To idTo
            s = s & LoadResData(i, 6)
        Next
        LoadResString = s
    End If
End Function

Public Function ConvertDateToUSFormat(d As Date) As String 'DD.MM.YYYY HH:MM:SS -> YYYY/MM/DD HH:MM:SS (for sorting purposes)
    ConvertDateToUSFormat = Format$(d, "yyyy\/mm\/dd hh:nn:ss", vbMonday)
End Function

'@ sCmdLine - in. full command line
'@ sBaseKey - in. key to search subkeys for
'@ aKey - out. array of subkeys
'@ aValue - out. array of values corresponding to subkeys
'ret - number of items in "aKey" and "aValue" arrays
'
'SubKey example: /autostart d:600
'
Public Function ParseSubCmdLine(sCmdLine As String, sBaseKey As String, aKey() As String, aValue() As String) As Long
    On Error GoTo ErrorHandler
    
    Dim pos As Long, pd As Long, cnt As Long
    Dim ch As String

    pos = InStr(1, sCmdLine, sBaseKey, 1)
    If pos <> 0 Then
        pos = pos + Len(sBaseKey) + 1
        Do
            ReDim Preserve aKey(cnt)
            ReDim Preserve aValue(cnt)
            ch = mid$(sCmdLine, pos, 1)
            If (ch = "-" Or ch = "/") Then Exit Do
            pd = InStr(pos, sCmdLine, ":")
            If pd = 0 Then Exit Do
            aKey(cnt) = LTrim$(mid$(sCmdLine, pos, pd - pos))
            If (mid$(sCmdLine, pd + 1, 1) = """") Then
                pos = InStr(pd + 2, sCmdLine, """")
            Else
                pos = InStr(pd + 1, sCmdLine, " ")
            End If
            If (pos = 0) Then
                aValue(cnt) = mid$(sCmdLine, pd + 1)
            Else
                aValue(cnt) = mid$(sCmdLine, pd + 1, pos - pd - 1)
                pos = pos + 1
                Do While mid$(sCmdLine, pos, 1) = " "
                    pos = pos + 1
                Loop
            End If
            cnt = cnt + 1
        Loop While pos
    End If
    If (cnt > 0) Then
        ReDim Preserve aKey(cnt - 1)
        ReDim Preserve aValue(cnt - 1)
    Else
        Erase aKey
        Erase aValue
    End If
    ParseSubCmdLine = cnt
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ParseSubCmdLine", sCmdLine, "BaseKey", sBaseKey
    If inIDE Then Stop: Resume Next
End Function

'@ sCmdLine - in. full command line
'@ sKey - in. key to get value of
'@ sValue - out. value of the specified key (if found)
'ret - true, if specified key was found
'
'Key example: /instDir:"c:\temp"
'
Public Function ParseCmdLineKey(ByVal sCmdLine As String, sKey As String, sValue As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim pos As Long
    Dim ch As String

    pos = InStr(1, sCmdLine, sKey, 1)
    If pos <> 0 Then
        sCmdLine = mid$(sCmdLine, pos + Len(sKey) + 1)
        ch = Left$(sCmdLine, 1)
        If ch = """" Then
            pos = InStr(2, sCmdLine, """")
            If pos <> 0 Then
                sValue = mid$(sCmdLine, 2, pos - 2)
            End If
        Else
            pos = InStr(1, sCmdLine, " ")
            If pos <> 0 Then
                sValue = Left$(sCmdLine, pos - 1)
            Else
                sValue = sCmdLine
            End If
        End If
        ParseCmdLineKey = True
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ParseCmdLineKey", sCmdLine, "BaseKey", sKey
    If inIDE Then Stop: Resume Next
End Function

Public Function GetDirForInstallationHJT() As String
    Dim sValue As String
    Dim sInstDir As String
    If ParseCmdLineKey(g_sCommandLine, "instDir", sValue) Then
        sInstDir = sValue
        sInstDir = GetLongPath(sInstDir)
        sInstDir = GetFullPath(sInstDir)
    Else
        sInstDir = Reg.GetString(HKLM, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HiJackThis Fork", "InstallLocation")
        If Len(sInstDir) = 0 Or Not FolderExists(sInstDir) Then
            sInstDir = BuildPath(PF_32, "HiJackThis Fork")
        End If
    End If
    GetDirForInstallationHJT = sInstDir
End Function

Public Function GetFullPathForInstallationHJT() As String
    Dim sInstDir As String
    Dim sIconStr As String
    Dim sExeName As String
    
    sInstDir = GetDirForInstallationHJT()
    sIconStr = Reg.GetString(HKLM, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HiJackThis Fork", "DisplayIcon")
    If Len(sIconStr) <> 0 And StrBeginWith(sIconStr, sInstDir) Then
        sExeName = GetFileNameAndExt(GetPathFromIconString(sIconStr))
    End If
    If Len(sExeName) = 0 Then sExeName = AppExeName(True)
    
    GetFullPathForInstallationHJT = BuildPath(sInstDir, sExeName)
End Function

Public Function GetFullPathOfInstalledHJT() As String
    Dim sInstDir As String
    Dim sIconStr As String
    Dim sExeName As String
    
    sInstDir = Reg.GetString(HKLM, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HiJackThis Fork", "InstallLocation")
    If Len(sInstDir) <> 0 Then
    
        sIconStr = Reg.GetString(HKLM, "Software\Microsoft\Windows\CurrentVersion\Uninstall\HiJackThis Fork", "DisplayIcon")
        If Len(sIconStr) <> 0 And StrBeginWith(sIconStr, sInstDir) Then
        
            sExeName = GetFileNameAndExt(GetPathFromIconString(sIconStr))
            
            GetFullPathOfInstalledHJT = BuildPath(sInstDir, sExeName)
        End If
    End If
End Function

Public Function IsInstalledHJT() As Boolean

    IsInstalledHJT = FileExists(GetFullPathOfInstalledHJT())

End Function

Public Sub NotifyChangeFrame(NewFrame As FRAME_ALIAS)

    g_CurFrame = NewFrame
    
    If g_CurFrame <> FRAME_ALIAS_SCAN Then
    
        If IsFormInit(frmSearch) Then
    
            frmSearch.Hide
        End If
    End If
End Sub

Public Function BitPrefix(sPrefix As String, HE As clsHiveEnum) As String
    If HE.Redirected Then
        BitPrefix = sPrefix & "-32"
    Else
        BitPrefix = sPrefix
    End If
End Function

