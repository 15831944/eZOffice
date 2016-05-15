''' <summary>
''' 帮助文档的位置
''' </summary>
Public Class HelpLocation

    Private settings1 As New HelpLocationSettings

    Private Sub HelpLocation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Me
            .TextBox_OfficeHelp.Text = settings1.OfficeHelp
            .TextBox_ExcelHelp.Text = settings1.ExcelHelp
        End With
    End Sub

    Private Sub Form1_FormClosing_1(ByVal sender As Object, ByVal e As _
            FormClosingEventArgs) Handles MyBase.FormClosing
        ' Save settings manually.
        settings1.Save()
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        With Me
            settings1.OfficeHelp = .TextBox_OfficeHelp.Text
            settings1.ExcelHelp = .TextBox_ExcelHelp.Text
        End With
        Me.Close()
    End Sub
End Class