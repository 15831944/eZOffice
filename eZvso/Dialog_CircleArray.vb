Imports System.Windows.Forms
Public Class Dialog_CircleArray
    Private Num As UShort = 4
    Private Angle As Double = 360
    Private blnCenter As Boolean = True
    Private blnPreserveDirection As Boolean = False


    Public Shadows Function ShowDialog(Optional ByRef Num As UShort = 4, _
                                     Optional ByRef Angle As Double = 360, _
                                     Optional ByRef blnCenter As Boolean = True, _
                                     Optional ByRef blnPreserveDirection As Boolean = False) As DialogResult
        Dim res As DialogResult = MyBase.ShowDialog()
        If res = Windows.Forms.DialogResult.OK Then
            With Me
                Num = .Num
                Angle = .Angle
                blnCenter = .blnCenter
                blnPreserveDirection = .blnPreserveDirection
            End With
        End If
        Return res

    End Function

#Region "   ---   事件处理"
    Private Sub txtAngle_TextChanged(sender As Object, e As EventArgs) Handles txtAngle.TextChanged
        Dim txt As TextBox = DirectCast(sender, TextBox)
        Double.TryParse(txt.Text, Angle)
    End Sub
    Private Sub txtNum_TextChanged(sender As Object, e As EventArgs) Handles txtNum.TextChanged
        Dim txt As TextBox = DirectCast(sender, TextBox)
        UShort.TryParse(txt.Text, Num)
    End Sub

    Private Sub CheckBox_preserveDirection_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_preserveDirection.CheckedChanged
        Dim box As CheckBox = DirectCast(sender, CheckBox)
        If box.Checked Then
            Me.blnPreserveDirection = True
        Else
            Me.blnPreserveDirection = False
        End If
    End Sub
    Private Sub RadioButton_Center_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Center.CheckedChanged, RadioButton_Border.CheckedChanged
        If Me.RadioButton_Center.Checked Then
            Me.blnCenter = True
        Else
            Me.blnCenter = False
        End If

    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

#End Region

End Class