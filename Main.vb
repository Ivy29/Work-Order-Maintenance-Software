Public Class Main

    Private Sub btnHardware_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHardware.Click
        hardware.Show()

    End Sub

    Private Sub btnSoftware_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSoftware.Click
        software.Show()

    End Sub

    Private Sub btnNetwork_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNetwork.Click
        network.Show()

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        If MsgBox("Are you sure you want to logout?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            login.Show()
            Me.Close()
        End If
    End Sub

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If USERLEVEL = "Guest" Then
            user.Text = USERLEVEL
            labelRemove.Visible = False
        ElseIf USERLEVEL = "Technician" Then
            labelRemove.Visible = False
            user.Text = USERNAME
        Else
            user.Text = USERNAME
        End If
    End Sub

    Private Sub labelRemove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles labelRemove.LinkClicked
        deleteUser.Show()
        Me.Close()
    End Sub
End Class