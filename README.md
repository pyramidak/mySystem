# mySystem
Class clsSystem has lots of usefull stuff. Just walk through:

    Public User As String
    Public DomainName As String
    Public Is64bit As Boolean
    Public AdminRole As Boolean
    Public LgeCzech As Boolean
    Public Path As New clsPath
    Public ComputerName As String
    Public Roaming As String
    Public Documents As String
    Public Function Framework() As Integer 
    Private Function GetLocalLanguage() As Boolean "get language setting for your app
    Public Sub LoadLanguageDictionary(Czech As Boolean, Folder As String) "change wpf app language file
    #Region " All about Shutdown and other commands "
    Public Function isAppRunning(sProcessName As String, Optional sUser As String = "") As Boolean
    Private Function GetProcessOwner(ProcessName As String, UserName As String) As Boolean
    Public Sub LoadHarddisks() "information about your disks with serial numbers
    Public Function GetTaskbarLocation() As TaskbarLocation
    Public Sub SetDisplayMode(ByVal Mode As DisplayMode) "First screen, Duplicate, Extend, Only second
    Public Sub MuteSound(Wnd As Window)
    
I would like to pass on my experience in VB.Net to others and thus support this language for future generations. There are difficult beginnings in any programming language, when you have to learn to use new language libraries so that you can take even the smallest step. I want to help you with that now. You're welcome, signed Zdeněk Jantač.
