Imports System.Data.OleDb

Public Class ucAttendance
    ' In-memory member cache: (FullName, School)
    Private members As New List(Of (FullName As String, School As String))()
    Private presentMembers As New List(Of String)()

    Private Sub ucAttendance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lstAttendance.Visible = False
        LoadMembers()
    End Sub

    Private Sub LoadMembers()
        members.Clear()
        If String.IsNullOrWhiteSpace(DbConfig.ConnString) Then Return

        Try
            Using conn As New OleDbConnection(DbConfig.ConnString)
                conn.Open()
                Using cmd As New OleDbCommand("SELECT FullName, Department FROM Members ORDER BY FullName", conn)
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim name = If(rdr.IsDBNull(0), String.Empty, rdr.GetString(0))
                            Dim school = If(rdr.FieldCount > 1 AndAlso Not rdr.IsDBNull(1), rdr.GetString(1), String.Empty)
                            members.Add((name, school))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.Print("LoadMembers error: " & ex.Message)
        End Try
    End Sub

    Private Sub dgvAttendance_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAttendance.CellClick
        ' If user clicks the ATTENDANCE column toggle/mark Present
        If e.ColumnIndex = 1 And e.RowIndex >= 0 Then
            Dim row = dgvAttendance.Rows(e.RowIndex)
            Dim name = If(row.Cells(0).Value, String.Empty).ToString()
            row.Cells(1).Value = "Present"
            If Not String.IsNullOrEmpty(name) AndAlso Not presentMembers.Contains(name) Then
                presentMembers.Add(name)
            End If
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        MessageBox.Show("Attendance saved successfully!")
        ' No DB persistence as requested
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        DoSearch(txtSearchName.Text)
    End Sub

    Private Sub txtSearchName_TextChanged(sender As Object, e As EventArgs) Handles txtSearchName.TextChanged
        ' Live prefix search as user types
        DoSearch(txtSearchName.Text)
    End Sub

    Private Sub DoSearch(query As String)
        lstAttendance.Items.Clear()
        Dim searchText As String = If(query, String.Empty).Trim()
        If searchText = String.Empty Then
            lstAttendance.Visible = False
            Return
        End If

        For Each m In members
            If m.FullName.StartsWith(searchText, StringComparison.OrdinalIgnoreCase) Then
                lstAttendance.Items.Add(m.FullName)
            End If
        Next

        If lstAttendance.Items.Count = 0 Then
            lstAttendance.Visible = False
        Else
            lstAttendance.Visible = True
            lstAttendance.BringToFront()
            lstAttendance.Focus()
        End If
    End Sub

    Private Sub lstAttendance_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstAttendance.SelectedIndexChanged
        If lstAttendance.SelectedIndex < 0 Then Return

        Dim selectedName As String = lstAttendance.SelectedItem?.ToString()
        If String.IsNullOrEmpty(selectedName) Then Return

        ' Find member record (first match)
        Dim memberIndex As Integer = -1
        For i As Integer = 0 To members.Count - 1
            If String.Equals(members(i).FullName, selectedName, StringComparison.OrdinalIgnoreCase) Then
                memberIndex = i
                Exit For
            End If
        Next

        Dim memberSchool As String = String.Empty
        If memberIndex >= 0 Then
            memberSchool = members(memberIndex).School
        End If

        ' See if the name already exists in the grid
        Dim foundRow As DataGridViewRow = Nothing
        For Each r As DataGridViewRow In dgvAttendance.Rows
            Dim cellName = If(r.Cells(0).Value, String.Empty).ToString()
            If String.Equals(cellName, selectedName, StringComparison.OrdinalIgnoreCase) Then
                foundRow = r
                Exit For
            End If
        Next

        If foundRow IsNot Nothing Then
            foundRow.Cells(1).Value = "Present"
        Else
            dgvAttendance.Rows.Add(selectedName, "Present", memberSchool)
        End If

        If Not presentMembers.Contains(selectedName) Then
            presentMembers.Add(selectedName)
        End If

        ' Clear search UI
        lstAttendance.Visible = False
        txtSearchName.Clear()
    End Sub

    Private Sub btnClearAttendance_Click(sender As Object, e As EventArgs) Handles btnClearAttendance.Click
        dgvAttendance.Rows.Clear()
        presentMembers.Clear()
    End Sub
End Class


