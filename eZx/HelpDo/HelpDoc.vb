Imports System.IO
Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon

Partial Public Class Ribbon_eZx

    Private frmHelpLocation As HelpLocation
    Private settings1 As New HelpLocationSettings

    Private Sub Group_Help_DialogLauncherClick(sender As Object, e As RibbonControlEventArgs) Handles Group_Help.DialogLauncherClick
        If frmHelpLocation Is Nothing Then
            frmHelpLocation = New HelpLocation
        End If
        frmHelpLocation.ShowDialog()
    End Sub

#Region "加载文档"

    ' 打开帮助文档所在文件夹
    Private Sub btn_OfficeHelp_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_OfficeHelp.Click
        settings1.Reload()
        Dim DirePath As String = settings1.OfficeHelp
        '
        If Directory.Exists(DirePath) Then

            Process.Start(DirePath)
        Else
            MessageBox.Show("指定的帮助文档不存在，请重新设置帮助文档路径。", "出错", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    ' 打开帮助文档
    Private Sub btn_ExcelHelp_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_ExcelHelp.Click

        settings1.Reload()
        Dim filePath As String = settings1.ExcelHelp
        '
        If File.Exists(filePath) Then
            Try
                Process.Start(filePath)
            Catch ex As Exception
                MessageBox.Show("指定的帮助文档无法打开。", "出错", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("指定的帮助文档不存在，请重新设置帮助文档路径。", "出错", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub


#End Region



End Class
