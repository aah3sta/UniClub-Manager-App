Public Class ucAttendance
    Dim members As New List(Of String)
    Dim presentMembers As New List(Of String)


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAttendance.CellContentClick

    End Sub

    Private Sub btnMarkPresent_Click(sender As Object, e As EventArgs)

    End Sub



    Private Sub dgvAttendance_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAttendance.CellClick


        If e.ColumnIndex = 2 And e.RowIndex >= 0 Then
            dgvAttendance.Rows(e.RowIndex).Cells(1).Value = "Present"
        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        MessageBox.Show("Attendance saved successfully!")
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        lstAttendance.Items.Clear()
        lstAttendance.Visible = True

        Dim searchText As String = txtSearchName.Text.ToLower()

        For Each member In members
            If member.ToLower().Contains(searchText) Then
                lstAttendance.Items.Add(member)
            End If
        Next

        If lstAttendance.Items.Count = 0 Then
            MessageBox.Show("No member found")
            lstAttendance.Visible = False
        End If

    End Sub

    Private Sub lstAttendance_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstAttendance.SelectedIndexChanged

        Visible = False
    End Sub

    Private Sub btnClearAttendance_Click(sender As Object, e As EventArgs) Handles btnClearAttendance.Click
        dgvAttendance.Rows.Clear()
    End Sub

    Private Sub txtSearchName_TextChanged(sender As Object, e As EventArgs) Handles txtSearchName.TextChanged

    End Sub
End Class


