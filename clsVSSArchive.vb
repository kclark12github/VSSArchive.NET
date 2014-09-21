'clsVSSArchive.vb
'   VSSArchive.NET Component...
'   Copyright © 2005, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   09/03/05    None        Ken Clark       Created;
'=================================================================================================================================
'Notes to Self...
'=================================================================================================================================
Option Strict Off
Option Explicit On 
Public Class clsVSSArchive
    Inherits clsBPEBase
    Public Sub New(ByVal objSupport As clsSupport, ByVal args() As String)
        MyBase.New(objSupport, "clsVSSArchive", trcOption.trcApplication)
        AppName = MyBase.ApplicationName
        GetCommandLine(args)
    End Sub

#Region "Events"
    'None at this time...
#End Region
#Region "Properties"
    Private AppName As String = vbNullString
    Private mAdminID As String = vbNullString
    Private mBackupDir As String = vbNullString
    Private mCL As clsCommandLine = Nothing
    Private mDatabaseName As String = vbNullString
    Private mDatabasePath As String = vbNullString
    Private mINIpath As String = vbNullString
    Private mLogFileName As String = vbNullString
    Private mPassword As String = vbNullString
    Private mProject As String = vbNullString
    Public ReadOnly Property AdminID() As String
        Get
            Return mAdminID
        End Get
    End Property
    Public ReadOnly Property BackupDir() As String
        Get
            Return mBackupDir
        End Get
    End Property
    Public ReadOnly Property DatabaseName() As String
        Get
            Return mDatabaseName
        End Get
    End Property
    Public ReadOnly Property DatabasePath() As String
        Get
            Return mDatabasePath
        End Get
    End Property
    Public ReadOnly Property INIpath() As String
        Get
            Return mINIpath
        End Get
    End Property
    Public ReadOnly Property LogFileName() As String
        Get
            Return mLogFileName
        End Get
    End Property
    Public ReadOnly Property Password() As String
        Get
            Return mPassword
        End Get
    End Property
    Public ReadOnly Property Project() As String
        Get
            Return mProject
        End Get
    End Property
#End Region
#Region "Methods"
#Region "Private Methods"
    Private Sub GetCommandLine(ByVal args() As String)
        Const EntryName As String = "GetCommandLine"
        Dim bTraceModeOnEntry As String

        'Note: The CommandLine Property should be executed prior to InitializeApplication if that
        '      routine is to be governed by the Trace information collected here...

        Try
            Dim strArgs As String = vbNullString
            For i As Integer = 0 To args.Length - 1
                strArgs &= args(i) & ";"
            Next
            RecordEntry(EntryName, "args:=""" & strArgs & """", trcOption.trcApplication, True)
            bTraceModeOnEntry = mSupport.Trace.TraceMode.ToString

            'CommandLine = Trim(Command())   'Command() returns data only if the command-line was issued for the component itself (i.e. only the Parent Application)...
            'CommandLine = "/U=SUNGARD /P=Z000000 /DSN=""V:\FiRRe\Config\WildFiRRe - WSRV08.DSN"" /TRACEFILE=""V:\FiRRe\SIASFiRRe.log"" /TraceOptions=trcSQL+trcApplicationDetail+trcADO /TrustedConnection /ConfirmCredentials=True /DBSecurity=False"
            'CommandLine = "/TRACEFILE=""V:\SIASFiRRe.log"" /TraceOptions=trcSQL+trcApplicationDetail+trcADO"

            'Step 1 Setup our SIASCL.CommandLine object...
            If mCL Is Nothing Then mCL = New clsCommandLine(mSupport)

            'Step 2 Process all the command-line arguments applicable to all FiRRe utilities
            '       Note that additional argument evaluation will be deferred until the specific
            '       function is executed...

            'Note that these command-line arguments will act to override defaults found in VSSArchive.ini...
            mCL.CommandLineArgs.Add("INIFILE", "/", "INIFile,INIFILE,inifile,IF,if", "File", False, 0)
            mCL.CommandLineArgs.Add("LOGFILE", "/", "LogFile,LOGFILE,logfile,LF,lf", "File", False, 0)
            mCL.CommandLineArgs.Add("TRACEFILE", "/", "TraceFile,TRACEFILE,tracefile,TF,tf", "File", False, 0)
            mCL.CommandLineArgs.Add("TRACEOPTIONS", "/", "TraceOptions,TRACEOPTIONS,traceoptions,TO,to", "String", False, 0)
            'Application-Specific parameters...
            mCL.CommandLineArgs.Add("ADMIN", "/", "Admin,ADMIN,admin,A,a", "String", False, 0)
            mCL.CommandLineArgs.Add("BACKUPDIR", "/", "BackupDir,BACKUPDIR,backupdir,BD,bd", "Directory", False, 0)
            mCL.CommandLineArgs.Add("DATABASE", "/", "Database,DATABASE,database,DB,db", "String", False, 0)
            mCL.CommandLineArgs.Add("PASSWORD", "/", "Password,PASSWORD,password,P,p", "String", False, 0)
            mCL.CommandLineArgs.Add("PROJECT", "/", "Project,PROJECT,project", "String", False, 0)

            'Step 3 Use .ParseCommandLine to evaluate text string into the .CommandLineArgs structure...
            Dim CommandLine As String = vbNullString
            For i As Integer = 1 To args.Length - 1
                CommandLine &= args(i) & " "
            Next
            If CommandLine <> vbNullString Then mCL.ParseCommandLine(CommandLine)

            'Step 4 Apply the command-line argument information to our properties as appropriate...
            'Dim AppPath As String = FindMainDir()
            With mCL.CommandLineArgs
                mINIpath = GetRegistrySetting(RootKeyConstants.HKEY_LOCAL_MACHINE, "SOFTWARE\SunGard\" & AppName, "INIfile", vbNullString)
                'Override INIpath if there's one on the CommandLine...
                If .Item("INIFILE").Present Then mINIpath = .Item("INIFILE").Value
                If mINIpath = vbNullString Then mINIpath = String.Format("{0}\{1}.ini", MyBase.ApplicationPath, AppName)
                If mINIpath = vbNullString Then Throw New ArgumentException("""/INI=<INI-File>"" must be specified.")
                If Dir(mINIpath, FileAttribute.Normal) = vbNullString Then Throw New ArgumentException(String.Format("""{0}"" (INIpath) not found.", mINIpath))
                If mSupport.FileSystem.ParsePath(mINIpath, ParseParts.FileNameBase) <> AppName Then AppName = mSupport.FileSystem.ParsePath(mINIpath, ParseParts.FileNameBase)

                'Trace arguments...
                If mSupport.Trace.TraceFile = vbNullString Then
                    mSupport.Trace.TraceFile = IIf(.Item("TRACEFILE").Present, .Item("TRACEFILE").Value, GetINIKey(mINIpath, AppName, "TraceFile", GetRegistrySetting(clsRegistry.RootKeyConstants.HKEY_LOCAL_MACHINE, "SOFTWARE\SunGard\FiRReServer", "TraceFile", mSupport.Trace.DefaultTraceFile)))
                End If
                If mSupport.Trace.TraceOptions = trcOption.trcNone Then
                    mSupport.Trace.TraceOptions = IIf(.Item("TRACEOPTIONS").Present, mSupport.Trace.ParseTraceOptions(.Item("TRACEOPTIONS").Value), mSupport.Trace.ParseTraceOptions(GetINIKey(mINIpath, AppName, "TraceOptions", "0")))
                End If
                If Not mSupport.Trace.TraceMode Then mSupport.Trace.TraceMode = CBool(mSupport.Trace.TraceOptions <> 0)
                mLogFileName = IIf(.Item("LOGFILE").Present, .Item("LOGFILE").Value, GetINIKey(mINIpath, AppName, "LogFile", vbNullString))

                mAdminID = IIf(.Item("ADMIN").Present, .Item("ADMIN").Value, GetINIKey(mINIpath, AppName, "AdminID", vbNullString))
                If mAdminID = vbNullString Then Throw New ArgumentException("/Admin=<AdministratorID> must be specified.")

                mBackupDir = IIf(.Item("BACKUPDIR").Present, .Item("BACKUPDIR").Value, GetINIKey(mINIpath, AppName, "BackupDir", vbNullString))
                If mBackupDir = vbNullString Then Throw New ArgumentException("/BackupDir=<Directory> must be specified.")

                mDatabaseName = IIf(.Item("DATABASE").Present, .Item("DATABASE").Value, GetINIKey(mINIpath, AppName, "DatabaseName", vbNullString))
                If mDatabaseName = vbNullString Then Throw New ArgumentException("/Database=<SourceSafeDatabaseName> must be specified.")

                mPassword = IIf(.Item("PASSWORD").Present, .Item("PASSWORD").Value, mSupport.Security.DecryptPassword(GetINIKey(mINIpath, AppName, "Password", vbNullString)))
                If mPassword = vbNullString Then Throw New ArgumentException("/Password=<AdministratorPassword> must be specified.")

                mProject = IIf(.Item("PROJECT").Present, .Item("PROJECT").Value, GetINIKey(mINIpath, AppName, "Project", vbNullString))
                'If mProject = vbNullString Then Throw New ArgumentException("/Project=<SourceSafeProjectName> must be specified.")
            End With
        Catch ex As Exception
            RaiseError()
        End Try

ExitSub:
        RecordExit(EntryName)
        Exit Sub
    End Sub
#Region "Win32 API Declarations"
    Private Const SYNCHRONIZE As Integer = &H100000
    Private Const INFINITE As Integer = &HFFFFFFFF '  Infinite timeout
    Private Const DEBUG_PROCESS As Short = &H1S
    Private Const DEBUG_ONLY_THIS_PROCESS As Short = &H2S

    Private Const CREATE_SUSPENDED As Short = &H4S

    Private Const DETACHED_PROCESS As Short = &H8S

    Private Const CREATE_NEW_CONSOLE As Short = &H10S

    Private Const NORMAL_PRIORITY_CLASS As Short = &H20S
    Private Const IDLE_PRIORITY_CLASS As Short = &H40S
    Private Const HIGH_PRIORITY_CLASS As Short = &H80S
    Private Const REALTIME_PRIORITY_CLASS As Short = &H100S

    Private Const CREATE_NEW_PROCESS_GROUP As Short = &H200S

    Private Const CREATE_NO_WINDOW As Integer = &H8000000

    Private Const WAIT_FAILED As Short = -1
    Private Const WAIT_OBJECT_0 As Short = 0
    Private Const WAIT_ABANDONED As Integer = &H80
    Private Const WAIT_ABANDONED_0 As Integer = &H80

    Private Const WAIT_TIMEOUT As Integer = &H102

    Private Const SW_SHOW As Short = 5

    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
    Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Integer, ByVal dwMilliseconds As Integer) As Integer
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    Declare Function AllocConsole Lib "kernel32" () As Integer
    Declare Function FreeConsole Lib "kernel32" () As Integer
    Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Integer) As Integer

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Integer, ByVal bAlertable As Integer) As Integer

    Public Const STD_OUTPUT_HANDLE As Short = -11
    Dim hConsole As Integer
#End Region
    Sub WaitForProcess(ByRef pid As Integer)
        Const lMilliseconds As Integer = 1000
        Dim phnd As Integer
        While True
            phnd = OpenProcess(SYNCHRONIZE, 0, pid)
            If phnd = 0 Then Exit Sub
            'WaitForSingleObject(phnd, INFINITE)
            WaitForSingleObject(phnd, lMilliseconds)
            Application.DoEvents()
            CloseHandle(phnd)
        End While
    End Sub
#End Region
    Public Sub Archive()
        Const EntryName As String = "Archive"

        Try
            RecordEntry(EntryName, Nothing, trcOption.trcApplication)
            Dim SCCServerPath As String = GetRegistrySetting(RootKeyConstants.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SourceSafe", "SCCServerPath", vbNullString)
            Dim SSARCPath As String = mSupport.FileSystem.ParsePath(SCCServerPath, ParseParts.DrvDir) & "SSARC.exe"
            Dim TimeStamp As String = VB6.Format(Now, "yyyymmdd.hhmmss")
            Dim VSSini As String = GetRegistrySetting(RootKeyConstants.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SourceSafe\Databases", mDatabaseName, vbNullString)
            Dim BaseName As String = mDatabaseName & IIf(mProject <> vbNullString, "." & mProject, vbNullString)
            Dim LogFile As String = String.Format("{0}\{1}.{2}.log", mBackupDir, BaseName, TimeStamp)
            Dim ArcFile As String = String.Format("{0}\{1}.{2}.ssa", mBackupDir, BaseName, TimeStamp)
            Dim strCommandLine() As String = {SSARCPath, LogFile, mAdminID, mPassword, ArcFile, mProject}
            Dim CommandLine As String = String.Format("{0} -d- -s..\ ""-o{1}"" -i- -y{2},{3} ""{4}"" $/{5}", strCommandLine)

            ChDrive(mSupport.FileSystem.ParsePath(SSARCPath, clsFileSystem.ParseParts.DrvOnly))
            ChDir(mSupport.FileSystem.ParsePath(SSARCPath, ParseParts.DrvDirNoSlash))

            LogMessage(mLogFileName, "[VSSarchive]", 0, False)
            LogMessage(mLogFileName, CommandLine, 0, False)

            Dim ProcessId As Integer = Shell(CommandLine, AppWinStyle.Hide)
            If ProcessId <> 0 Then WaitForProcess(ProcessId)

            'If we successfully backed-up our database, purge files older than a month (i.e. 28 days)
            Dim ArcFileInfo As New FileInfo(ArcFile)
            If ArcFileInfo.Exists Then
                Dim BackupDirInfo As New DirectoryInfo(BackupDir)
                Dim BackupFileList() As FileInfo = BackupDirInfo.GetFiles(String.Format("{0}.*.*", BaseName))
                For Each iFileInfo As FileInfo In BackupFileList
                    If DateDiff(DateInterval.DayOfYear, iFileInfo.LastWriteTime, Now) > 28 Then iFileInfo.Delete()
                Next
            End If
        Catch ex As Exception
            RaiseError()
        End Try

ExitSub:
        RecordExit(EntryName)
        Exit Sub
    End Sub
#End Region
#Region "Event Handlers"
    'None at this time...
#End Region
    Shared Sub Main()
        Dim objSupport As New clsSupport(Nothing)
        'Dim objSupport As New clsSupport(Nothing, "V:\FiRReSentinel.trace", trcOption.trcAll, True)
        'Dim objSupport As New clsSupport(Nothing, "V:\FiRReSentinel.trace", trcOption.trcServer Or trcOption.trcFileWatcher Or trcOption.trcMSMQWatcher Or trcOption.trcTaskWatcher Or trcOption.trcTCPIPWatcher, True)
        Dim mMyModuleName As String = "clsVSSArchive"
        Dim mMyTraceID As String = String.Format("{0}.{1}.{2}", objSupport.ApplicationName, objSupport.EntryComponent, mMyModuleName)
        Dim objVSSArchive As clsVSSArchive

        Application.EnableVisualStyles()
        Application.DoEvents()

        Try
            objSupport.CallStack.RecordEntry(objSupport.ApplicationName, objSupport.EntryComponent, mMyModuleName, "Main")

            objVSSArchive = New clsVSSArchive(objSupport, Environment.GetCommandLineArgs())
            With objVSSArchive
                .Archive()
                System.Environment.ExitCode = 0
            End With
        Catch ex As Exception
            Dim Message As String = String.Format("{0}: {1}{2}", ex.GetType.ToString, ex.Message, vbCrLf)
            Message &= vbCrLf
            Message &= String.Format("StackTrace: {0}", vbCrLf)
            Message &= String.Format("{0}{1}", ex.StackTrace.ToString, vbCrLf)
            objSupport.UI.ShowMsgBox(Message, MsgBoxStyle.Critical)
        Finally
        End Try

ExitSub:
        If Not IsNothing(objVSSArchive) Then objVSSArchive.Dispose() : objVSSArchive = Nothing
        objSupport.CallStack.RecordExit(objSupport.ApplicationName, objSupport.EntryComponent, mMyModuleName, "Main")
        If Not IsNothing(objSupport) Then objSupport.Dispose() : objSupport = Nothing
        Application.Exit()
        Exit Sub
    End Sub
End Class
