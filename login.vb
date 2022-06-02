Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Public Class login
    Dim mydata As New DataTable
    Dim myquery As New MyClassLibrary.MyHelper

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        login()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Register.Show()
        Me.Close()
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Public Sub login()
        connect(myquery)
        mydata = myquery.runQuery("SELECT user_id, role, username,name FROM login WHERE username ='" + txtUser.Text + "'and Password = MD5('" + AddSlashes(txtPass.Text) + "')")

        If mydata.Rows.Count > 0 Then
            USERID = mydata.Rows(0).Item("user_id")
            USERNAME = mydata.Rows(0).Item("name")
            Name = mydata.Rows(0).Item("username")
            USERLEVEL = mydata.Rows(0).Item("role")
        End If

        '  MsgBox(USERNAME)
        If Name.ToString = txtUser.Text.ToString And Name.ToString <> "" Then
            txtUser.Focus()

            txtUser.Text = ""
            txtPass.Text = ""
            Main.Show()
            Me.Close()
        Else
            MsgBox("Invalid username or password!", MsgBoxStyle.Exclamation)
            txtPass.Focus()

            txtPass.SelectAll()
            'txtUser.Text = ""
            'txtPass.Text = ""
        End If
        'mydata = myquery.runQuery("SELECT * FROM login WHERE username= '" + txtUser.Text.ToString + "'and Password= MD5('" + AddSlashes(txtPass.Text) + "')")

        'If mydata.Rows.Count <= 0 Then
        'MessageBox.Show("Username or Password Invalid! ")
        'Else
        'MessageBox.Show("Login Successful!")
        'Main.Show()
        'Me.Close()
        'End If
    End Sub

    Private Sub labelGuest_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles labelGuest.LinkClicked
        USERLEVEL = "Guest"
        Main.Show()
        Me.Close()
    End Sub

    Private Sub txtPass_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPass.KeyPress
        If e.KeyChar = Chr(13) Then 
            btnLogin_Click(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub txtUser_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUser.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtPass.Focus()
        End If
    End Sub
End Class