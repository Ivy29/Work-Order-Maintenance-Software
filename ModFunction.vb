Module ModFunction
    Dim myquery As New MyClassLibrary.MyHelper
    Public Function isSelected(ByVal dgv As System.Windows.Forms.DataGridView) As Boolean
        If dgv.SelectedRows.Count <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub connect(ByRef myquery As MyClassLibrary.MyHelper)
        myquery.setConnection(LOCALHOST, PORT, DBUSER, DBPASSWORD, DATABASE)
    End Sub

    Public Sub makeUpper(ByRef txtcontrols As System.Windows.Forms.Control.ControlCollection)
        Dim txttemp As New System.Windows.Forms.TextBox
        For i As Integer = 0 To txtcontrols.Count - 1
            If txtcontrols.Item(i).ToString.Contains("Forms.TextBox") Then
                Try
                    txttemp = txtcontrols.Item(i)
                    txttemp.Text = txttemp.Text.ToUpper.Trim
                Catch ex As Exception
                End Try
            End If
        Next
    End Sub

    Public Enum State
        Add = 0
        Edit = 1
    End Enum

    'Public Sub initializeDataTables()
    '    connect(myquery)
    '    accounts = myquery.runQuery("SELECT   `ra_ID`   , `Account`   , `Acc_No`, `rCategory` FROM   `ctois`.`ref_account` WHERE (`isDeleted` =0);")
    '    account_category = myquery.runQuery("SELECT   `rac_ID`   , `Category` FROM   `ctois`.`ref_account_category` WHERE (`isDeleted` =0);")
    'End Sub
End Module
