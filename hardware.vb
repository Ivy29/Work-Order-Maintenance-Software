Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Public Class hardware
    Dim mydata As New DataTable
    Dim myquery As New MyClassLibrary.MyHelper
    Private Sub hardware_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connect(myquery)
        form_reload()
        generate()

        txtOff.Enabled = False
        txtTechnician.Enabled = False
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

        AddHard.txtRequest.Text = i
    End Sub

    Public Sub form_reload()
        If USERLEVEL = "Administrator" Then
            If chkPending.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Pending'")
            ElseIf chkFinished.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Finished'")
            ElseIf chkUnserviceable.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Unserviceable'")
            Else
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Date_Delivered BETWEEN '" & periodFrom.Value.ToString("yyyy-MM-01") & "' AND '" & periodTo.Value.ToString("yyyy-MM-31") & "'")
            End If

        ElseIf USERLEVEL = "Guest" Then
            btnAdd.Enabled = False
            btnDelete.Enabled = False
            btnEdit.Enabled = False
            EditToolStripMenuItem.Enabled = False
            AddNewRecordToolStripMenuItem.Enabled = False
            DeleteAllToolStripMenuItem.Enabled = False
            DeleteToolStripMenuItem.Enabled = False

            If chkPending.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Pending'")
            ElseIf chkFinished.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Finished'")
            ElseIf chkUnserviceable.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Unserviceable'")
            Else
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Date_Delivered BETWEEN '" & periodFrom.Value.ToString("yyyy-MM-01") & "' AND '" & periodTo.Value.ToString("yyyy-MM-31") & "'")
            End If

        Else
            DeleteAllToolStripMenuItem.Enabled = False
            If chkPending.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Pending' AND Technician='" + USERNAME + "'")
            ElseIf chkFinished.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Finished' AND Technician='" + USERNAME + "'")
            ElseIf chkUnserviceable.Checked = True Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Status= 'Unserviceable' AND Technician='" + USERNAME + "'")
            Else
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Technician='" + USERNAME + "' AND Date_Delivered BETWEEN '" & periodFrom.Value.ToString("yyyy-MM-01") & "' AND '" & periodTo.Value.ToString("yyyy-MM-31") & "'")
            End If
        End If

        dgvHardware.Rows.Clear()
        For i As Integer = 0 To mydata.Rows.Count - 1
            With dgvHardware
                .Rows.Add()
                With .Rows(i)
                    dgvHardware.Rows(i).Cells(0).Value = mydata.Rows(i).Item("Request_No").ToString
                    dgvHardware.Rows(i).Cells(1).Value = mydata.Rows(i).Item("Status").ToString
                    dgvHardware.Rows(i).Cells(2).Value = mydata.Rows(i).Item("Remarks").ToString
                    dgvHardware.Rows(i).Cells(3).Value = mydata.Rows(i).Item("Technician").ToString
                    dgvHardware.Rows(i).Cells(4).Value = mydata.Rows(i).Item("Inventory_Tag_No").ToString
                    dgvHardware.Rows(i).Cells(5).Value = mydata.Rows(i).Item("Description").ToString
                    dgvHardware.Rows(i).Cells(6).Value = mydata.Rows(i).Item("Item").ToString
                    dgvHardware.Rows(i).Cells(7).Value = mydata.Rows(i).Item("Office").ToString
                    dgvHardware.Rows(i).Cells(8).Value = mydata.Rows(i).Item("Date_delivered").ToString
                    dgvHardware.Rows(i).Cells(9).Value = mydata.Rows(i).Item("Delivered_by").ToString
                    dgvHardware.Rows(i).Cells(10).Value = mydata.Rows(i).Item("Contact_no").ToString
                    dgvHardware.Rows(i).Cells(11).Value = mydata.Rows(i).Item("ID_no").ToString
                    dgvHardware.Rows(i).Cells(12).Value = mydata.Rows(i).Item("Accountable_officer").ToString
                    dgvHardware.Rows(i).Cells(13).Value = mydata.Rows(i).Item("Date_accomplished").ToString
                    dgvHardware.Rows(i).Cells(14).Value = mydata.Rows(i).Item("Findings").ToString
                    dgvHardware.Rows(i).Cells(15).Value = mydata.Rows(i).Item("Action_taken").ToString
                    dgvHardware.Rows(i).Cells(16).Value = mydata.Rows(i).Item("Retrieved_by").ToString
                    dgvHardware.Rows(i).Cells(17).Value = mydata.Rows(i).Item("Date_Retrieved").ToString
                End With
            End With
        Next
    End Sub

    Private Sub dgvHardware_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvHardware.CellFormatting
        If e.ColumnIndex = 1 And e.Value IsNot Nothing Then
            If e.Value = "PENDING" Then
                e.CellStyle.BackColor = Color.Red
            ElseIf e.Value = "FINISHED" Then
                e.CellStyle.BackColor = Color.Green
            ElseIf e.Value = "UNSERVICEABLE" Then
                e.CellStyle.BackColor = Color.Yellow
            Else
                e.CellStyle.BackColor = Color.LightCyan
            End If
        End If
    End Sub
    Private Sub AddNewRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddNewRecordToolStripMenuItem.Click
        addRecord()
    End Sub
    Private Sub EditToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        editRecord()
    End Sub
    Public Sub addRecord()
        AddHard.hard_stat.Text = "Add"
        AddHard.ShowDialog()
    End Sub

    Public Sub deleteRecord()
        If dgvHardware.Rows.Count > 0 Then
            If MsgBox("Are you sure you want to delete?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                With myquery
                    connect(myquery)
                    mydata = myquery.runQuery("DELETE FROM tbl_hardware WHERE Request_No = '" & dgvHardware.SelectedRows(0).Cells(0).Value & "'")
                    form_reload()
                End With
            End If
        Else
            MsgBox("There is no data to be deleted! ", MsgBoxStyle.Critical)
        End If
    End Sub

    Public Sub editRecord()
        AddHard.hard_stat.Text = "Edit"
        If dgvHardware.Rows.Count > 0 Then
            AddHard.ID = dgvHardware.CurrentRow.Cells(0).Value.ToString
            AddHard.ShowDialog()
        Else
            MsgBox("There is no data to edit! ", MsgBoxStyle.Critical)
        End If
    End Sub
    Private Sub SelectedRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectedRowToolStripMenuItem.Click
        deleteRecord()
    End Sub

    Private Sub DeleteAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteAllToolStripMenuItem.Click
        If dgvHardware.Rows.Count > 0 Then
            If MsgBox("Are you sure you want to delete all data?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                With myquery
                    connect(myquery)
                    mydata = myquery.runQuery("DELETE FROM tbl_hardware")
                    form_reload()
                End With
            End If
        Else
            MsgBox("There is no data to be deleted! ", MsgBoxStyle.Critical)
        End If
    End Sub
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        connect(myquery)
        Dim mydata As New DataTable

        'TRY
        If USERLEVEL = "Technician" Then
            If CmbSearch.Text <> "" Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE " & CmbSearch.Text & " LIKE '%" & txtSearch.Text.ToString & "%'")
            Else
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Technician= '" + USERNAME + "'")
            End If
        Else
            If CmbSearch.Text <> "" Then
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE " & CmbSearch.Text & " LIKE '%" & txtSearch.Text.ToString & "%'")
            Else
                mydata = myquery.runQuery("SELECT * FROM tbl_hardware")
            End If
        End If

        dgvHardware.Rows.Clear()
        For i As Integer = 0 To mydata.Rows.Count - 1
            With dgvHardware
                .Rows.Add()
                With .Rows(i)
                    dgvHardware.Rows(i).Cells(0).Value = mydata.Rows(i).Item("Request_No").ToString
                    dgvHardware.Rows(i).Cells(1).Value = mydata.Rows(i).Item("Status").ToString
                    dgvHardware.Rows(i).Cells(2).Value = mydata.Rows(i).Item("Remarks").ToString
                    dgvHardware.Rows(i).Cells(3).Value = mydata.Rows(i).Item("Technician").ToString
                    dgvHardware.Rows(i).Cells(4).Value = mydata.Rows(i).Item("Inventory_Tag_No").ToString
                    dgvHardware.Rows(i).Cells(5).Value = mydata.Rows(i).Item("Description").ToString
                    dgvHardware.Rows(i).Cells(6).Value = mydata.Rows(i).Item("Item").ToString
                    dgvHardware.Rows(i).Cells(7).Value = mydata.Rows(i).Item("Office").ToString
                    dgvHardware.Rows(i).Cells(8).Value = mydata.Rows(i).Item("Date_delivered").ToString
                    dgvHardware.Rows(i).Cells(9).Value = mydata.Rows(i).Item("Delivered_by").ToString
                    dgvHardware.Rows(i).Cells(10).Value = mydata.Rows(i).Item("Contact_no").ToString
                    dgvHardware.Rows(i).Cells(11).Value = mydata.Rows(i).Item("ID_no").ToString
                    dgvHardware.Rows(i).Cells(12).Value = mydata.Rows(i).Item("Accountable_officer").ToString
                    dgvHardware.Rows(i).Cells(13).Value = mydata.Rows(i).Item("Date_accomplished").ToString
                    dgvHardware.Rows(i).Cells(14).Value = mydata.Rows(i).Item("Findings").ToString
                    dgvHardware.Rows(i).Cells(15).Value = mydata.Rows(i).Item("Action_taken").ToString
                    dgvHardware.Rows(i).Cells(16).Value = mydata.Rows(i).Item("Retrieved_by").ToString
                    dgvHardware.Rows(i).Cells(17).Value = mydata.Rows(i).Item("Date_Retrieved").ToString
                End With
            End With
        Next

        CmbSearch.Text = ""
        txtSearch.Text = ""
    End Sub

    Private Sub chkPending_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPending.CheckedChanged
        form_reload()
    End Sub

    Private Sub chkFinished_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFinished.CheckedChanged
        form_reload()
    End Sub

    Private Sub chkUnserviceable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUnserviceable.CheckedChanged
        form_reload()
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        connect(myquery)

        mydata = myquery.runQuery("SELECT * FROM tbl_hardware WHERE Date_Delivered BETWEEN '" & periodFrom.Value.ToString("yyyy-MM-01") & "' AND '" & periodTo.Value.ToString("yyyy-MM-31") & "'")

        dgvHardware.Rows.Clear()
        For i As Integer = 0 To mydata.Rows.Count - 1
            With dgvHardware
                .Rows.Add()
                With .Rows(i)
                    dgvHardware.Rows(i).Cells(0).Value = mydata.Rows(i).Item("Request_No").ToString
                    dgvHardware.Rows(i).Cells(1).Value = mydata.Rows(i).Item("Status").ToString
                    dgvHardware.Rows(i).Cells(2).Value = mydata.Rows(i).Item("Remarks").ToString
                    dgvHardware.Rows(i).Cells(3).Value = mydata.Rows(i).Item("Technician").ToString
                    dgvHardware.Rows(i).Cells(4).Value = mydata.Rows(i).Item("Inventory_Tag_No").ToString
                    dgvHardware.Rows(i).Cells(5).Value = mydata.Rows(i).Item("Description").ToString
                    dgvHardware.Rows(i).Cells(6).Value = mydata.Rows(i).Item("Item").ToString
                    dgvHardware.Rows(i).Cells(7).Value = mydata.Rows(i).Item("Office").ToString
                    dgvHardware.Rows(i).Cells(8).Value = mydata.Rows(i).Item("Date_delivered").ToString
                    dgvHardware.Rows(i).Cells(9).Value = mydata.Rows(i).Item("Delivered_by").ToString
                    dgvHardware.Rows(i).Cells(10).Value = mydata.Rows(i).Item("Contact_no").ToString
                    dgvHardware.Rows(i).Cells(11).Value = mydata.Rows(i).Item("ID_no").ToString
                    dgvHardware.Rows(i).Cells(12).Value = mydata.Rows(i).Item("Accountable_officer").ToString
                    dgvHardware.Rows(i).Cells(13).Value = mydata.Rows(i).Item("Date_accomplished").ToString
                    dgvHardware.Rows(i).Cells(14).Value = mydata.Rows(i).Item("Findings").ToString
                    dgvHardware.Rows(i).Cells(15).Value = mydata.Rows(i).Item("Action_taken").ToString
                    dgvHardware.Rows(i).Cells(16).Value = mydata.Rows(i).Item("Retrieved_by").ToString
                    dgvHardware.Rows(i).Cells(17).Value = mydata.Rows(i).Item("Date_Retrieved").ToString
                End With
            End With
        Next

    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        periodFrom.ResetText()
        periodTo.ResetText()
        form_reload()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        addRecord()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        editRecord()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        deleteRecord()
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Main.Show()
        Me.Close()
    End Sub

    Private Sub rbOffice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOffice.CheckedChanged
        txtTechnician.Enabled = False
        txtOff.Enabled = True
        txtTechnician.ResetText()
    End Sub

    Private Sub rbTechnician_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTechnician.CheckedChanged
        txtTechnician.Enabled = True
        txtOff.Enabled = False
        txtOff.ResetText()
    End Sub

    Private Sub rbOverall_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOverall.CheckedChanged
        txtTechnician.Enabled = False
        txtOff.Enabled = False
        txtOff.ResetText()
        txtTechnician.ResetText()
    End Sub


    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim dtpo As New DataTable
        Dim crFile As String = Nothing
        Dim rpTitle As String = Nothing
        Dim rptSource As String = Nothing
        Dim myDs As New DataSet
        Dim GeneratorReport As New rpt_Viewer 'vb viewer
        rpTitle = "Summary of Hardware"

        connect(myquery)

        Dim rptDoc As CrystalDecisions.CrystalReports.Engine.ReportDocument

        If rbOverall.Checked = True Then
            With myquery
                mydata = myquery.runQuery("Select Office, Technician, Inventory_Tag_No,Item, Description, Remarks, Findings, Action_Taken, Delivered_by, Date_Delivered, Date_Accomplished, Retrieved_by, Date_Retrieved from tbl_hardware where status = 'FINISHED'")
                mydata.TableName = "tbl_hardware"
                myDs.Tables.Add(mydata)
            End With
            rptDoc = New rptHardwareOverall
            rptDoc.SetDataSource(myDs)
        ElseIf rbTechnician.Checked = True Then
            With myquery
                mydata = myquery.runQuery("Select Office, Technician, Inventory_Tag_No,Item, Description, Remarks, Findings, Action_Taken, Delivered_by, Date_Delivered, Date_Accomplished, Retrieved_by, Date_Retrieved from tbl_hardware where status = 'FINISHED' AND Technician= '" & txtTechnician.Text.ToString & "'")
                mydata.TableName = "tbl_hardware"
                myDs.Tables.Add(mydata)
            End With
            rptDoc = New rptHardwareTech
            rptDoc.SetDataSource(myDs)
        Else
            With myquery
                mydata = myquery.runQuery("Select Office, Technician, Inventory_Tag_No,Item, Description, Remarks, Findings, Action_Taken, Delivered_by, Date_Delivered, Date_Accomplished, Retrieved_by, Date_Retrieved from tbl_hardware where status = 'FINISHED' AND Office= '" & txtOff.Text.ToString & "'")
                mydata.TableName = "tbl_hardware"
                myDs.Tables.Add(mydata)
            End With
            rptDoc = New rptHardwareOffice
            rptDoc.SetDataSource(myDs)
        End If

        'para ni sa txt box same parameter name

        With GeneratorReport
            .blnDataSource = True
            .ReportSource = rptSource
            .ReportPath = crFile
            .ReportTitle = rpTitle
            .crptViewer.ReportSource = rptDoc
            .WindowState = FormWindowState.Maximized
            .crptViewer.Zoom(75)
            .ShowDialog()
            .crptViewer.Refresh()
            .Dispose()
        End With
    End Sub

    Private Sub txtOff_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOff.KeyPress
        If e.KeyChar = Chr(13) Then
            btnPrint_Click(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub txtTechnician_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTechnician.KeyPress
        If e.KeyChar = Chr(13) Then
            btnPrint_Click(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub txtSearch_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If e.KeyChar = Chr(13) Then
            btnSearch_Click(Me, EventArgs.Empty)
        End If
    End Sub
End Class