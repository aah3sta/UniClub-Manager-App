Imports System.Data.OleDb
'Importing the data-handling package

Public Class ucDashboard
    'Connecting the database such that the content on it is accessible
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\UniClubDB.accdb"
    Dim conn As New OleDbConnection(connString)
    Private Sub pnlMain_Paint(sender As Object, e As PaintEventArgs) Handles pnlMain.Paint
        'Get the last 3 entries from the database and display them as the most recent entries
        Try
            If conn.State = ConnectionState.Open Then conn.Close()
            conn.Open()

            Dim sql As String = "SELECT TOP 3 FullName, Department, PhoneNumber, Course FROM Members ORDER BY JoinDate DESC"
            Dim adapter As New OleDbDataAdapter(sql, conn)
            Dim dt As New DataTable()

            adapter.Fill(dt)
            dgvRecentMembers.DataSource = dt
            conn.Close()

            'Error - handling if the recent activity fails to load
        Catch ex As Exception
            MessageBox.Show("Could not load recent activity: " & ex.Message)
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try

        'Get member count and display on screen
        Try
            conn.Open()
            Dim cmd As New OleDbCommand("SELECT COUNT(*) FROM Members", conn)
            Dim count As Integer = CInt(cmd.ExecuteScalar())
            lblMemberCount.Text = count.ToString()
            conn.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading stats: " & ex.Message)
            conn.Close()
        End Try

        'Get event count and display on the screen
        Try
            conn.Open()
            Dim cmd As New OleDbCommand("SELECT COUNT(*) FROM Events", conn)
            Dim count As Integer = CInt(cmd.ExecuteScalar())
            lblEventCount.Text = count.ToString()
            conn.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading stats: " & ex.Message)
            conn.Close()
        End Try
    End Sub

End Class
''''''''''''''''''''''''''''''
