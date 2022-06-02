Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Public Class AddNet
    Dim mydata As New DataTable
    Dim myquery As New MyClassLibrary.MyHelper
    Public ID As String = ""


    Private Sub AddNet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connect(myquery)

        dateAccomplished.Format = DateTimePickerFormat.Custom
        dateAccomplished.CustomFormat = "yyyy-MM-dd hh:mm:ss"
        dateAccomplished.Value = Now()

        dateRequested.Format = DateTimePickerFormat.Custom
        dateRequested.CustomFormat = "yyyy-MM-dd hh:mm:ss"
        dateRequested.Value = Now()

        generate()

        mydata = myquery.runQuery("SELECT DISTINCT Description FROM tbl_network")
        txtDescription.DataSource = mydata
        txtDescription.DisplayMember = "tbl_network"
        txtDescription.ValueMember = "Description"

        If net_stat.Text = "Add" Then
            reset_text()
            txtStatus.Text = "Pending"
            txtStatus.Enabled = False
            load_template_technician()
        ElseIf net_stat.Text = "Edit" Then
            form_load()
        End If
    End Sub
    Public Sub load_template_technician()
        connect(myquery)

        Dim mydata As DataTable = myquery.runQuery("SELECT * FROM login WHERE role='Technician'")
        txtTechnician.Items.Clear()
        For i As Integer = 0 To mydata.Rows.Count - 1
            txtTechnician.Items.Add(mydata.Rows(i).Item("name"))
        Next
        mydata = Nothing

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If net_stat.Text = "Add" Then
            save()
        ElseIf net_stat.Text = "Edit" Then
            UpdateNetwork()
        End If
    End Sub

    Public Sub form_load()
        connect(myquery)
        With myquery

            mydata = .runQuery("SELECT * FROM tbl_network WHERE Request_No='" + ID.ToString + "'")
            If net_stat.Text = "Add" Then
                reset_text()
                'do nothing

            ElseIf net_stat.Text = "Edit" Then
                txtRequest.Text = mydata.Rows(0).Item("Request_No")
                txtStatus.Text = mydata.Rows(0).Item("Status")
                txtRemarks.Text = mydata.Rows(0).Item("Remarks")
                txtTechnician.Text = mydata.Rows(0).Item("Technician")
                txtOff.Text = mydata.Rows(0).Item("Requesting_office")
                txtRequested.Text = mydata.Rows(0).Item("Requested_by")
                txtContact.Text = mydata.Rows(0).Item("Contact_no")
                txtDescription.Text = mydata.Rows(0).Item("Description")
                txtUnit.Text = mydata.Rows(0).Item("Unit")
                dateRequested.Text = mydata.Rows(0).Item("Date_requested")
                dateAccomplished.Text = mydata.Rows(0).Item("Date_accomplished")
                txtInv.Text = mydata.Rows(0).Item("Inventory_tag")
            End If
        End With
    End Sub

    Public Sub save()
        If txtRequest.Text = "" Then
            MsgBox("Please leave no blank in each item", MsgBoxStyle.Critical)

        Else
            With myquery
                .setTable("tbl_network")
                .addValue(txtRequest.Text, "Request_No")
                .addValue(txtStatus.Text, "Status")
                .addValue(txtRemarks.Text, "Remarks")
                .addValue(txtTechnician.Text, "Technician")
                .addValue(txtOff.Text, "Requesting_office")
                .addValue(txtRequested.Text, "Requested_by")
                .addValue(txtContact.Text, "Contact_no")
                .addValue(txtDescription.Text, "Description")
                .addValue(txtUnit.Text, "Unit")
                .addValue(dateRequested.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_requested")
                .addValue(dateAccomplished.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_accomplished")
                .addValue(txtInv.Text, "Inventory_tag")

                .runInsert()
                MsgBox("Successfully added!", MsgBoxStyle.OkOnly)
                reset_text()

                network.form_reload()
                Me.Close()

            End With
        End If
    End Sub

    Public Sub UpdateNetwork()
        With myquery
            .setTable("tbl_network")
            .addValue(txtRequest.Text, "Request_No")
            .addValue(txtStatus.Text, "Status")
            .addValue(txtRemarks.Text, "Remarks")
            .addValue(txtTechnician.Text, "Technician")
            .addValue(txtOff.Text, "Requesting_office")
            .addValue(txtRequested.Text, "Requested_by")
            .addValue(txtContact.Text, "Contact_no")
            .addValue(txtDescription.Text, "Description")
            .addValue(txtUnit.Text, "Unit")
            .addValue(dateRequested.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_requested")
            .addValue(dateAccomplished.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_accomplished")
            .addValue(txtInv.Text, "Inventory_tag")

            .runUpdate("Request_No= '" & txtRequest.Text.ToString & "'")

            MsgBox("Successfully updated!", MsgBoxStyle.OkOnly)
            reset_text()

            network.form_reload()
            Me.Close()
        End With
    End Sub

    Public Sub reset_text()
        txtStatus.ResetText()
        txtRemarks.ResetText()
        txtTechnician.ResetText()
        txtOff.ResetText()
        txtRequested.ResetText()
        txtContact.ResetText()
        txtDescription.ResetText()
        txtUnit.ResetText()
        dateRequested.ResetText()
        dateAccomplished.ResetText()
        txtInv.ResetText()

    End Sub

    Public Sub generate()
        connect(myquery)
        Dim i As Integer

        mydata = myquery.runQuery("SELECT MAX(Request_No) FROM tbl_network")

        If mydata.Rows.Count > 0 And Not IsDBNull(mydata.Rows(0).Item(0)) Then
            i = CInt(mydata.Rows(0).Item(0) + 1)
        Else
            i = 1
        End If

        txtRequest.Text = i

    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        network.Show()
        Me.Close()
    End Sub

    Private Sub txtTechnician_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTechnician.TextChanged
        txtStatus.Text = ""
        txtStatus.Enabled = True
    End Sub
End Class