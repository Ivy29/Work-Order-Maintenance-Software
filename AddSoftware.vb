Imports MySql.Data.MySqlClient
Imports System.Windows.Forms


Public Class AddSoftware
    Dim mydata As New DataTable
    Dim myquery As New MyClassLibrary.MyHelper
    Public ID As String = ""

    Private Sub AddSoftware_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connect(myquery)
        dateAccomplished.Format = DateTimePickerFormat.Custom
        dateAccomplished.CustomFormat = "yyyy-MM-dd hh:mm:ss"
        dateAccomplished.Value = Now()

        dateRequested.Format = DateTimePickerFormat.Custom
        dateRequested.CustomFormat = "yyyy-MM-dd hh:mm:ss"
        dateRequested.Value = Now()

        generate()
        load_template_software()
        load_template_system()
        load_template_description()

        If soft_stat.Text = "Add" Then
            reset_text()
            'txtTechnician.Enabled = False
            txtStatus.Text = "Pending"
            txtStatus.Enabled = False
            load_template_technician()

        ElseIf soft_stat.Text = "Edit" Then
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

    Public Sub generate()
        connect(myquery)
        Dim i As Integer

        mydata = myquery.runQuery("SELECT MAX(Request_No) FROM tbl_software")

        If mydata.Rows.Count > 0 And Not IsDBNull(mydata.Rows(0).Item(0)) Then
            i = CInt(mydata.Rows(0).Item(0) + 1)
        Else
            i = 1
        End If
        txtRequest.Text = i
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If soft_stat.Text = "Add" Then
            save()
        ElseIf soft_stat.Text = "Edit" Then
            UpdateSoftware()
        End If
    End Sub
    Public Sub form_load()
        connect(myquery)
        With myquery
            mydata = .runQuery("SELECT * FROM tbl_software WHERE Request_No='" + ID.ToString + "'")

            If soft_stat.Text = "Add" Then
                reset_text()
                'do nothing

            ElseIf soft_stat.Text = "Edit" Then
                txtRequest.Text = mydata.Rows(0).Item("Request_No")
                txtStatus.Text = mydata.Rows(0).Item("Status")
                txtRemarks.Text = mydata.Rows(0).Item("Remarks")
                txtTechnician.Text = mydata.Rows(0).Item("Technician")
                txtOff.Text = mydata.Rows(0).Item("Requesting_office")
                txtRequested.Text = mydata.Rows(0).Item("Requested_by")
                txtContact.Text = mydata.Rows(0).Item("Contact_No")
                dateRequested.Text = mydata.Rows(0).Item("Date_requested")
                txtDesc.Text = mydata.Rows(0).Item("Description")
                txtSoftware.Text = mydata.Rows(0).Item("Software")
                txtSystem.Text = mydata.Rows(0).Item("System")
                dateAccomplished.Text = mydata.Rows(0).Item("Date_accomplished")
                txtInventory.Text = mydata.Rows(0).Item("Inventory_tag_no")
            End If
        End With
    End Sub

    Public Sub UpdateSoftware()
        With myquery
            .setTable("tbl_software")
            .addValue(txtRequest.Text, "Request_No")
            .addValue(txtStatus.Text, "Status")
            .addValue(txtRemarks.Text, "Remarks")
            .addValue(txtTechnician.Text, "Technician")
            .addValue(txtOff.Text, "Requesting_office")
            .addValue(txtRequested.Text, "Requested_by")
            .addValue(txtContact.Text, "Contact_no")
            .addValue(dateRequested.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_Requested")
            .addValue(txtDesc.Text, "Description")
            .addValue(txtSoftware.Text, "Software")
            .addValue(txtSystem.Text, "System")
            .addValue(dateAccomplished.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_accomplished")
            .addValue(txtInventory.Text, "Inventory_tag_no")

            .runUpdate("Request_No= '" & txtRequest.Text.ToString & "'")

            MsgBox("Successfully updated!", MsgBoxStyle.OkOnly)
            reset_text()

            software.form_reload()
            Me.Close()
        End With
    End Sub

    Public Sub save()
        If txtRequest.Text = "" Then
            MsgBox("Please leave no blank in each item", MsgBoxStyle.Critical)
        Else
            With myquery
                .setTable("tbl_software")
                .addValue(txtRequest.Text, "Request_No")
                .addValue(txtStatus.Text, "Status")
                .addValue(txtRemarks.Text, "Remarks")
                .addValue(txtTechnician.Text, "Technician")
                .addValue(txtOff.Text, "Requesting_office")
                .addValue(txtRequested.Text, "Requested_by")
                .addValue(txtContact.Text, "Contact_no")
                .addValue(dateRequested.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_requested")
                .addValue(txtSoftware.Text, "Software")
                .addValue(txtSystem.Text, "System")
                .addValue(txtDesc.Text, "Description")
                .addValue(dateAccomplished.Value.ToString("yyyy-MM-dd hh:mm:ss"), "Date_accomplished")
                .addValue(txtInventory.Text, "Inventory_tag_no")

                .runInsert()
                MsgBox("Successfully added!", MsgBoxStyle.OkOnly)
                reset_text()

                software.form_reload()
                Me.Close()
            End With
        End If
    End Sub
    Public Sub reset_text()
        txtStatus.ResetText()
        txtRemarks.ResetText()
        txtTechnician.ResetText()
        txtOff.ResetText()
        txtRequested.ResetText()
        txtContact.ResetText()
        dateRequested.ResetText()
        txtSoftware.ResetText()
        txtSystem.ResetText()
        txtDesc.ResetText()
        dateAccomplished.ResetText()
        txtInventory.ResetText()
    End Sub
    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        software.Show()
        Me.Close()
    End Sub

    Private Sub txtTechnician_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTechnician.TextChanged
        txtStatus.Text = ""
        txtStatus.Enabled = True
    End Sub

    Public Sub load_template_software()
        mydata = myquery.runQuery("SELECT DISTINCT Software FROM tbl_software")
        txtSoftware.DataSource = mydata
        txtSoftware.DisplayMember = "tbl_software"
        txtSoftware.ValueMember = "Software"
    End Sub

    Public Sub load_template_system()
        mydata = myquery.runQuery("SELECT DISTINCT System FROM tbl_software")
        txtSystem.DataSource = mydata
        txtSystem.DisplayMember = "tbl_software"
        txtSystem.ValueMember = "System"
    End Sub
    Public Sub load_template_description()
        mydata = myquery.runQuery("SELECT DISTINCT Description FROM tbl_software")
        txtDesc.DataSource = mydata
        txtDesc.DisplayMember = "tbl_software"
        txtDesc.ValueMember = "Description"
    End Sub
End Class