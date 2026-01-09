Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class ucEvents
    Private EventsList As New List(Of String)()
    Private ReadOnly connString As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=|DataDirectory|\UniClubDB BD.accdb"

    Private Sub btnAddEvent_Click(sender As Object, e As EventArgs) Handles btnAddEvent.Click
        ' Validate required fields
        If txtEventName.Text = "" Or txtVenue.Text = "" Or
            txtLocation.Text = "" Or txtCapacity.Text = "" Then
            MessageBox.Show("Please fill all fields",
                        "Missing Information",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
            Exit Sub
        End If

        'Connect to database and insert new event
        Try
            Dim sql As String = "INSERT INTO Events ([EventName], [Venue], [Location], [EventDate], [Capacity], [EventDescription]) " &
                           "VALUES (@EventName, @Venue, @Location, @EventDate, @Capacity, @EventDescription)"
            Using conn As New OleDbConnection(connString)
                Using cmd As New OleDbCommand(sql, conn)
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Add("@EventName", OleDbType.VarWChar).Value = txtEventName.Text.Trim()
                    cmd.Parameters.Add("@Venue", OleDbType.VarWChar).Value = txtVenue.Text.Trim()
                    cmd.Parameters.Add("@Location", OleDbType.VarWChar).Value = txtLocation.Text.Trim()
                    cmd.Parameters.Add("@EventDate", OleDbType.Date).Value = DateTimePicker1.Value.Date
                    cmd.Parameters.Add("@Capacity", OleDbType.Integer).Value = Convert.ToInt32(txtCapacity.Text.Trim())
                    cmd.Parameters.Add("@EventDescription", OleDbType.VarWChar).Value = "No Description Provided"

                    conn.Open()
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                    conn.Close()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Event registered successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Clear form
                        txtEventName.Clear()
                        txtVenue.Clear()
                        txtLocation.Clear()
                        txtCapacity.Clear()
                        txtEventID.Clear()
                        DateTimePicker1.Value = Date.Now
                    Else
                        MessageBox.Show("Event logging failed: no rows were inserted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error saving event: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)
        ' Clear form
        txtEventName.Clear()
        txtVenue.Clear()
        txtLocation.Clear()
        txtCapacity.Clear()
        txtEventID.Clear()
        DateTimePicker1.Value = Date.Now
    End Sub
End Class
