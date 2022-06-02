Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Public Class AddHard
    Dim mydata As New DataTable
    Dim myquery As New MyClassLibrary.MyHelper
    Public ID As String = ""

    Private Sub AddHard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connect(myquery)

        dateDelivered.Format = DateTimePickerFormat.Custom
        dateDelivered.CustomFormat = "yyyy-MM-dd hh:mm:ss"
        dateDelivered.Value = Now()

        dateAccomplished.Format = DateTimePickerFormat.Custom
        dateAccomplished.CustomFormat = "yyyy-MM-dd hh:mm:ss"
        dateAccomplished.Value = Now()

        dateRetrieved.Format = DateTimePickerFormat.Custom
        dateRetrieved.CustomFormat = "yyyy-MM-dd hh:mm:ss"
        dateRetrieved.Value = Now()

        generate()

        comboLoadDesc()
        comboLoadItem()

           
        If hard_stat.Text = "Add" Then
            reset_text()
            load_template_technician()
            'txtTechnician.Enabled = False
            txtStatus.Text = "Pending"
            txtStatus.Enabled = False

        ElseIf hard_stat.Text = "Edit" Then
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


    Private Sub btnSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If hard_stat.Text = "Add" Then
            save()
        ElseIf hard_stat.Text = "Edit" Then
            UpdateHardware()
        End If
    End Sub

    Public Sub form_load()
        connect(myquery)
        With myquery

            mydata = .runQuery("SELECT * FROM tbl_hardware WHERE Request_No='" + ID.ToString + "'")
            If hard_stat.Text = "Add" Then
                reset_text()
                'do nothing

            ElseIf hard_stat.Text = "Edit" Then
                txtRequest.Text = mydata.Rows(0).Item("Request_No")
                txtStatus.Text = mydata.Rows(0).Item("Status")
                txtRemarks.Text = mydata.Rows(0).Item("Remarks")
                txtTechnician.Text = mydata.Rows(0).Item("Technician")
                txtInventory.Text = mydata.Rows(0).Item("Inventory_Tag_No")
                txtDesc.Text = mydata.Rows(0).Item("Description")
                txtItem.Text = mydata.Rows(0).Item("Item")
                txtOff.Text = mydata.Rows(0).Item("Office")
                txtAccOff.Text = mydata.Rows(0).Item("Accountable_Officer")
                txtDelivered.Text = mydata.Rows(0).Item("Delivered_by")
                txtContact.Text = mydata.Rows(0).Item("Contact_no")
                txtID.Text = mydata.Rows(0).Item("ID_no")
                dateDelivered.Value = mydata.Rows(0).Item("Date_Delivered")
                dateAccomplished.Value = mydata.Rows(0).Item("Date_Accomplished")
                txtFindings.Text = mydata.Rows(0).Item("Findings")
                txtAction.Text = mydata.Rows(0).Item("Action_Taken")
                txtRetrieved.Text = mydata.Rows(0).Item("Retrieved_by")
                dateRetrieved.Value = mydata.Rows(0).Item("Date_Retrieved")
            End If
        End With
    End Sub

    Public Sub save()

        If txtRequest.Text = "" Then
            MsgBox("Please leave no blank in each item", MsgBoxStyle.Critical)
        Else
            With myquery
                .setTable("tbl_hardware")
                .addValue(txtRequest.Text, "Request_No")
                .addValue(txtStatus.Text, "Status")
                .addValue(txtRemarks.Text, "Remarks")
                .addValue(txtTechnician.Text, "Technician")
                .addValue(txtInventory.Text, "Inventory_Tag_No")
                .addValue(txtDesc.Text, "Description")
                .addValue(txtItem.Text, "Item")
                .addValue(txtOff.Text, "Office")
                .addValue(dateDelivered.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_Delivered")
                .addValue(txtDelivered.Text, "Delivered_by")
                .addValue(txtContact.Text, "Contact_no")
                .addValue(txtID.Text, "ID_no")
                .addValue(txtAccOff.Text, "Accountable_Officer")
                .addValue(dateAccomplished.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_Accomplished")
                .addValue(txtFindings.Text, "Findings")
                .addValue(txtAction.Text, "Action_taken")
                .addValue(txtRetrieved.Text, "Retrieved_by")
                .addValue(dateRetrieved.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_Retrieved")

                .runInsert()
                MsgBox("Successfully added!", MsgBoxStyle.OkOnly)

                reset_text()
                hardware.form_reload()
                Me.Close()
            End With
        End If
    End Sub

    Public Sub UpdateHardware()
        comboLoadDesc()
        comboLoadItem()

        With myquery
            .setTable("tbl_hardware")
            .addValue(txtRequest.Text, "Request_No")
            .addValue(txtStatus.Text, "Status")
            .addValue(txtRemarks.Text, "Remarks")
            .addValue(txtTechnician.Text, "Technician")
            .addValue(txtInventory.Text, "Inventory_Tag_No")
            .addValue(txtDesc.Text, "Description")
            .addValue(txtItem.Text, "Item")
            .addValue(txtOff.Text, "Office")
            .addValue(dateDelivered.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_Delivered")
            .addValue(txtDelivered.Text, "Delivered_by")
            .addValue(txtContact.Text, "Contact_no")
            .addValue(txtID.Text, "ID_no")
            .addValue(txtAccOff.Text, "Accountable_Officer")
            .addValue(dateAccomplished.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_Accomplished")
            .addValue(txtFindings.Text, "Findings")
            .addValue(txtAction.Text, "Action_taken")
            .addValue(txtRetrieved.Text, "Retrieved_by")
            .addValue(dateRetrieved.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_Retrieved")

            .runUpdate("Request_No= '" & txtRequest.Text.ToString & "'")

            MsgBox("Successfully updated!", MsgBoxStyle.OkOnly)
            reset_text()

            hardware.form_reload()
            Me.Close()
        End With
    End Sub

    Public Sub reset_text()
        txtStatus.ResetText()
        txtRemarks.ResetText()
        txtTechnician.ResetText()
        txtInventory.ResetText()
        txtDesc.ResetText()
        txtItem.ResetText()
        txtOff.ResetText()
        txtAccOff.ResetText()
        txtDelivered.ResetText()
        txtContact.ResetText()
        txtID.ResetText()
        dateDelivered.ResetText()
        dateAccomplished.ResetText()
        txtFindings.ResetText()
        txtAction.ResetText()
        txtRetrieved.ResetText()
        dateRetrieved.ResetText()
    End Sub
    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        hardware.Show()
        Me.Close()
    End Sub
    Public Sub generate()
        connect(myquery)
        Dim i As Integer

        mydata = myquery.runQuery("SELECT MAX(Request_No) FROM tbl_hardware")

        If mydata.Rows.Count > 0 And Not IsDBNull(mydata.Rows(0).Item(0)) Then
            i = CInt(mydata.Rows(0).Item(0) + 1)
        Else
            i = 1
        End If

        txtRequest.Text = i

    End Sub

    Public Sub comboLoadItem()
        mydata = myquery.runQuery("SELECT DISTINCT Item FROM tbl_hardware")
        txtItem.DataSource = mydata
        txtItem.DisplayMember = "tbl_hardware"
        txtItem.ValueMember = "Item"
    End Sub

    Public Sub comboLoadDesc()
        mydata = myquery.runQuery("SELECT DISTINCT Description FROM tbl_hardware")
        txtDesc.DataSource = mydata
        txtDesc.DisplayMember = "tbl_hardware"
        txtDesc.ValueMember = "Description"
    End Sub

    Private Sub txtTechnician_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTechnician.TextChanged
        txtStatus.Text = ""
        txtStatus.Enabled = True
    End Sub
End Class