Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Public Class Register
    Dim mydata As New DataTable
    Dim myquery As New MyClassLibrary.MyHelper

    Private Sub register_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connect(myquery)
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtName.Text = "" Or txtUsername.Text = "" Or txtPass.Text = "" Then
            MsgBox("Please leave no blank in each item", MsgBoxStyle.Critical)
        Else
            With myquery
                .setTable("login")

                If txtPass1.Text = txtPass.Text Then
                    .addValue(txtName.Text, "name")
                    .addValue(txtUsername.Text, "username")
                    .addValue(txtUser_level.Text, "role")
                    '.addValue(txtPass.Text, "password")
                    .addValue("md5('" + AddSlashes(txtPass.Text) + "')", "Password")

                    .runInsert()
                    MsgBox("Successfully Saved", MsgBoxStyle.OkOnly)

                    login.Show()
                    Me.Close()
                Else
                    MsgBox("Password didn't matched!", MsgBoxStyle.Critical)
                End If
            End With
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        login.Show()
        Me.Close()
    End Sub

    Private Sub labelRemove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        deleteUser.Show()
        Me.Close()
    End Sub

    Private Sub txtPass1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPass1.KeyPress
        If e.KeyChar = Chr(13) Then
            btnSave_Click(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub txtName_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtUsername.Focus()
        End If
    End Sub

    Private Sub txtUsername_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUsername.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtUser_level.Focus()
        End If
    End Sub

    Private Sub txtUser_level_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUser_level.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtPass.Focus()
        End If
    End Sub

    Private Sub txtPass_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtPass1.Focus()
        End If
    End Sub
End Class