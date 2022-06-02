Imports MySql.Data.MySqlClient
Imports System.Windows.Forms
Public Class deleteUser
    Dim mydata As New DataTable
    Dim myquery As New MyClassLibrary.MyHelper
    Private Sub deleteUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connect(myquery)
        form_reload()
    End Sub
    Public Sub form_reload()
        mydata = myquery.runQuery("SELECT * FROM login")
        dgvUser.Rows.Clear()
        For i As Integer = 0 To mydata.Rows.Count - 1
            With dgvUser
                .Rows.Add()
                With .Rows(i)
                    dgvUser.Rows(i).Cells(0).Value = mydata.Rows(i).Item("user_id").ToString
                    dgvUser.Rows(i).Cells(1).Value = mydata.Rows(i).Item("Name").ToString
                    dgvUser.Rows(i).Cells(2).Value = mydata.Rows(i).Item("Username").ToString
                    dgvUser.Rows(i).Cells(3).Value = mydata.Rows(i).Item("Role").ToString
                End With
            End With
        Next
    End Sub

    Private Sub DeleteUserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteUserToolStripMenuItem.Click
        If dgvUser.Rows.Count > 0 Then
            If MsgBox("Are you sure you want to delete this user account?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                With myquery
                    connect(myquery)
                    mydata = myquery.runQuery("DELETE FROM login WHERE user_id = '" & dgvUser.SelectedRows(0).Cells(0).Value & "'")
                    form_reload()
                End With
            End If
        Else
            MsgBox("There is no data to be deleted! ", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Main.Show()
        Me.Close()
    End Sub
End Class