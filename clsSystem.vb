#Region " System "

Public Class clsSystem
    Inherits Collection(Of clsWindows)

    Public BuildYear As Integer = 2020
    Public User As String
    Public DomainName As String
    Public Is64bit As Boolean
    Public AdminRole As Boolean
    Public LgeCzech As Boolean
    Public Path As New clsPath
    Public ComputerName As String

#Region " Sub New "

    Sub New()
        sFramework()
        Is64bit = Environment.Is64BitOperatingSystem
        LgeCzech = GetLocalLanguage()
        ComputerName = Environment.MachineName
        User = System.Environment.UserName
        DomainName = System.Environment.UserDomainName
        My.User.InitializeWithWindowsUser()
        AdminRole = My.User.IsInRole(Microsoft.VisualBasic.ApplicationServices.BuiltInRole.Administrator)
    End Sub

#End Region

#Region " Paths "

    Public Class clsPath
        Private Company As String = "pyramidak"
        Public Roaming As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\" & Company
        Public Documents As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & Company
    End Class
#End Region

#Region " Framework "

    Public Function Framework() As Integer
        Return System.Environment.Version.Major
    End Function

    Public Function sFramework() As String
        Dim sFW As String = System.Environment.Version.Major.ToString + "." + System.Environment.Version.Minor.ToString
        If sFW = "2.0" Then sFW += " (3.5)"
        Return sFW
    End Function

#End Region

#Region " Local Language "

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function GetLocaleInfo(ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
    End Function

    Private Function GetLocalLanguage() As Boolean
        Dim LOCALE_USER_DEFAULT As Integer = &H400
        Dim LOCALE_SENGLANGUAGE As Integer = &H1001
        Dim Buffer As String, Ret As Integer
        Buffer = New String(Chr(0), 256)
        Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGLANGUAGE, Buffer, Len(Buffer))
        Dim SysTrue As Boolean = If(Buffer.Substring(0, Ret - 1) = "Czech", True, False)
        Dim RegTruePath As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + Application.ProductName
        Dim RegTrue As String = myRegister.GetValue(HKEY.LOCALE_MACHINE, RegTruePath, "DefaultLanguage", "")
        If RegTrue = "" Then RegTrue = myRegister.GetValue(HKEY.CURRENT_USER, RegTruePath, "DefaultLanguage", "")
        Return If(RegTrue = "", SysTrue, If(RegTrue = "Czech", True, False))
    End Function

    Public Sub LoadLanguageDictionary(Czech As Boolean, Folder As String)
        Dim LgeDict As New ResourceDictionary()
        LgeDict.Source = New Uri("/" + Application.ExeName + ";component/" & Folder & "/" + If(Czech, "CZ", "EN") + "-String.xaml", UriKind.Relative)
        For Each Resource As ResourceDictionary In Application.Current.Resources.MergedDictionaries
            If Resource.Source.ToString.EndsWith("-String.xaml") Then 'remove current language
                Application.Current.Resources.MergedDictionaries.Remove(Resource)
                Exit For
            End If
        Next
        Application.Current.Resources.MergedDictionaries.Add(LgeDict) 'add selected language
    End Sub

#End Region

#Region " Shutdown "

    <DllImport("user32.dll")>
    Private Shared Function ExitWindowsEx(ByVal uFlags As Integer, ByVal dwReason As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function LockWorkStation() As Boolean
    End Function

    <DllImport("powrprof.dll", SetLastError:=True)>
    Public Shared Function SetSuspendState(<[In](), MarshalAs(UnmanagedType.I1)> ByVal Hibernate As Boolean, <[In](), MarshalAs(UnmanagedType.I1)> ByVal ForceCritical As Boolean, <[In](), MarshalAs(UnmanagedType.I1)> ByVal DisableWakeEvent As Boolean) As <MarshalAs(UnmanagedType.I1)> Boolean
    End Function

    <DllImport("advapi32.dll")>
    Private Shared Function InitiateSystemShutdownEx(<MarshalAs(UnmanagedType.LPStr)> ByVal lpMachinename As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpMessage As String, ByVal dwTimeout As Int32, ByVal bForceAppsClosed As Boolean, ByVal bRebootAfterShutdown As Boolean, ByVal dwReason As Int32) As Boolean
    End Function

    Public Sub StandBy()
        SetSuspendState(False, False, False)
    End Sub

    Public Sub Lock()
        LockWorkStation()
    End Sub

    Public Sub PowerOff()
        System.Diagnostics.Process.Start("Shutdown", "-s -t 0")
    End Sub

    Public Sub Restart()
        System.Diagnostics.Process.Start("Shutdown", "-r -t 0")
    End Sub

    Public Sub LogOff()
        'ExitWindowsEx(0, 0) Log Off
        ExitWindowsEx(4, 0) 'forced Log Off
    End Sub
#End Region

#Region " Process "

    Public Function isAppRunning(sProcessName As String, Optional sUser As String = "") As Boolean
        Return GetProcessOwner(sProcessName, sUser)
        If UBound(Diagnostics.Process.GetProcessesByName(sProcessName)) > 0 Then
            If sUser = "" Then
                Return True
            Else
                Return GetProcessOwner(sProcessName, sUser)
            End If
        Else
            Return False
        End If
    End Function

    Private Function GetProcessOwner(ProcessName As String, UserName As String) As Boolean
        Try
            Dim CountInstance As Integer
            Dim selectQuery As SelectQuery = New SelectQuery("Select * from Win32_Process Where Name = '" + ProcessName + ".exe' ")
            Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher(selectQuery)
            Dim y As System.Management.ManagementObjectCollection
            y = searcher.Get
            For Each proc As ManagementObject In y
                Dim sOwner(1) As String
                proc.InvokeMethod("GetOwner", CType(sOwner, Object()))
                If proc("Name").ToString = ProcessName & ".exe" Then
                    If sOwner(0) = UserName Then
                        CountInstance = CountInstance + 1
                        If CountInstance > 1 Then Return True
                    End If
                End If
            Next
        Catch ex As Exception
        End Try
        Return False
    End Function

    Public Shared Function isProcess(ByVal ProcessID As Integer) As Boolean
        If ProcessID = 0 Then Return False
        Try
            Process.GetProcessById(ProcessID, ".")
            Return True
        Catch
            Return False
        End Try
    End Function

    Public Shared Function isProcess(ByVal ProcessName As String) As Boolean
        If ProcessName = "" Then Return False
        Try
            If Process.GetProcessesByName(ProcessName, ".").Length > 0 Then Return True
        Catch
        End Try
        Return False
    End Function

    Public Shared Function GetProcess(ByVal ProcessID As Integer) As Process
        If ProcessID = 0 Then Return Nothing
        Try
            Return Process.GetProcessById(ProcessID, ".")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Function GetProcess(ByVal ProcessName As String) As Process()
        If ProcessName = "" Then Return Nothing
        Try
            Return Process.GetProcessesByName(ProcessName, ".")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Function GetProcess(ByVal ProcessID As Integer, ByVal WindowTitle As String) As Process
        If WindowTitle = "" Then Return Nothing
        Dim allProcesses(), thisProcess As Process
        allProcesses = System.Diagnostics.Process.GetProcesses
        For Each thisProcess In allProcesses
            If thisProcess.MainWindowTitle = WindowTitle Then
                If ProcessID = 0 Then
                    Return thisProcess
                Else
                    If thisProcess.Id = ProcessID Then Return thisProcess
                End If
            End If
        Next
        Return Nothing
    End Function

#End Region

#Region " Physical Harddisks "

    Public DiskLetter As String = Environment.SystemDirectory.Substring(0, 1)
    Private Harddisky As New Collection(Of clsHarddisk)

    Public Property HardDisks() As Collection(Of clsHarddisk)
        Get
            If Harddisky.Count = 0 Then LoadHarddisks()
            Return Harddisky
        End Get
        Set(ByVal value As Collection(Of clsHarddisk))
            Harddisky = value
        End Set
    End Property

    Class clsHarddisk
        Public Property DeviceID As String
        Public Property Model As String
        Public Property SerialNumber As String
        Public Property Letter As String
        Public Property Type As DiskTypes
    End Class

    Public Sub LoadHarddisks()
        If Harddisky.Count = 0 Then
            Try
                For Each drive As ManagementObject In New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive").[Get]()
                    Dim disk As New clsHarddisk
                    For Each partition As ManagementObject In drive.GetRelated("Win32_DiskPartition")
                        For Each logical As ManagementObject In partition.GetRelated("Win32_LogicalDisk")
                            disk.Letter = logical("Name").ToString
                            If disk.Letter.Length > 1 Then disk.Letter = disk.Letter.Substring(0, 1)
                            'Exit For
                        Next
                    Next
                    disk.DeviceID = drive("DeviceId").ToString
                    disk.Type = If(drive("Model").ToString.ToLower.Contains("usb"), DiskTypes.Flashdisk_8, DiskTypes.Harddisk_7)
                    disk.Model = disk.Letter & ":   " & drive("Model").ToString
                    disk.SerialNumber = drive("SerialNumber").ToString.Trim.Replace("-", "")
                    If disk.SerialNumber.Length >= 8 AndAlso disk.SerialNumber.Substring(0, 5) <> "00000" Then
                        If Harddisky.FirstOrDefault(Function(x) x.Model = disk.Model) Is Nothing Then Harddisky.Add(disk)
                    End If
                Next
            Catch ex As Exception
            End Try
        End If
    End Sub

#End Region

#Region " Taskbar Location "

    Public Enum TaskbarLocation
        Top
        Bottom
        Left
        Right
    End Enum

    Public Function GetTaskbarLocation() As TaskbarLocation
        Dim bounds As New Rect(New Size(System.Windows.SystemParameters.PrimaryScreenWidth, System.Windows.SystemParameters.PrimaryScreenHeight))
        Dim working As Rect = System.Windows.SystemParameters.WorkArea
        If working.Height < bounds.Height And working.Y > 0 Then
            Return TaskbarLocation.Top
        ElseIf working.Height < bounds.Height And working.Y = 0 Then
            Return TaskbarLocation.Bottom
        ElseIf working.Width < bounds.Width And working.X > 0 Then
            Return TaskbarLocation.Left
        ElseIf working.Width < bounds.Width And working.X = 0 Then
            Return TaskbarLocation.Right
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region " Change Screen "
    Enum DisplayMode
        Internal
        External
        Extend
        Duplicate
    End Enum
    Public Sub SetDisplayMode(ByVal Mode As DisplayMode)
        Dim proc = New Process()
        proc.StartInfo.FileName = "DisplaySwitch.exe"

        Select Case Mode
            Case DisplayMode.External
                proc.StartInfo.Arguments = "/external"
            Case DisplayMode.Internal
                proc.StartInfo.Arguments = "/internal"
            Case DisplayMode.Extend
                proc.StartInfo.Arguments = "/extend"
            Case DisplayMode.Duplicate
                proc.StartInfo.Arguments = "/clone"
        End Select

        proc.Start()
    End Sub

#End Region

#Region " Mute sound "

    <DllImport("user32.dll")>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    Private Const WM_APPCOMMAND As Integer = &H319
    Private Const APPCOMMAND_VOLUME_MUTE As Integer = &H80000
    Private Const APPCOMMAND_VOLUME_DOWN As Integer = &H90000
    Private Const APPCOMMAND_VOLUME_UP As Integer = &HA0000

    Public Sub MuteSound(Wnd As Window)
        Dim wMainPtr As IntPtr = New System.Windows.Interop.WindowInteropHelper(Wnd).Handle
        SendMessage(New Interop.WindowInteropHelper(Wnd).Handle, WM_APPCOMMAND, IntPtr.Zero, New IntPtr(APPCOMMAND_VOLUME_MUTE))
    End Sub

#End Region

End Class

#End Region