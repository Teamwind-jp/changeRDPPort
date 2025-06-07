'-------------------------------------------------
'rdpポート　引数に変更して再起動　bye

Public Class Form1


    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function GetCurrentProcess() As IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr,
    ByVal DesiredAccess As Integer,
    ByRef TokenHandle As IntPtr) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function CloseHandle(ByVal hHandle As IntPtr) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True,
    CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function LookupPrivilegeValue(ByVal lpSystemName As String,
    ByVal lpName As String,
    ByRef lpLuid As Long) As Boolean
    End Function

    <System.Runtime.InteropServices.StructLayout(
    System.Runtime.InteropServices.LayoutKind.Sequential, Pack:=1)>
    Private Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As Integer
        Public Luid As Long
        Public Attributes As Integer
    End Structure

    <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function AdjustTokenPrivileges(ByVal TokenHandle As IntPtr,
    ByVal DisableAllPrivileges As Boolean,
    ByRef NewState As TOKEN_PRIVILEGES,
    ByVal BufferLength As Integer,
    ByVal PreviousState As IntPtr,
    ByVal ReturnLength As IntPtr) As Boolean
    End Function

    'シャットダウンするためのセキュリティ特権を有効にする
    Public Shared Sub AdjustToken()
        Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
        Const TOKEN_QUERY As Integer = &H8
        Const SE_PRIVILEGE_ENABLED As Integer = &H2
        Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"

        If Environment.OSVersion.Platform <> PlatformID.Win32NT Then
            Return
        End If

        Dim procHandle As IntPtr = GetCurrentProcess()

        'トークンを取得する
        Dim tokenHandle As IntPtr
        OpenProcessToken(procHandle,
        TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, tokenHandle)
        'LUIDを取得する
        Dim tp As New TOKEN_PRIVILEGES()
        tp.Attributes = SE_PRIVILEGE_ENABLED
        tp.PrivilegeCount = 1
        LookupPrivilegeValue(Nothing, SE_SHUTDOWN_NAME, tp.Luid)
        '特権を有効にする
        AdjustTokenPrivileges(tokenHandle, False, tp, 0, IntPtr.Zero, IntPtr.Zero)

        '閉じる
        CloseHandle(tokenHandle)
    End Sub

    Public Enum ExitWindows
        EWX_LOGOFF = &H0
        EWX_SHUTDOWN = &H1
        EWX_REBOOT = &H2
        EWX_POWEROFF = &H8
        EWX_RESTARTAPPS = &H40
        EWX_FORCE = &H4
        EWX_FORCEIFHUNG = &H10
    End Enum

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ExitWindowsEx(ByVal uFlags As ExitWindows,
    ByVal dwReason As Integer) As Boolean
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'rdpポート変更

        My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp", "PortNumber", 50514)

        AdjustToken()
        ExitWindowsEx(ExitWindows.EWX_REBOOT, 0)


        End


    End Sub



End Class
