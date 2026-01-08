Imports System.Data.OleDb

Public Class ucEvents

    Private EventsList As New List(Of String)
    Private lstEvents As Object

    Private Sub ucEvents_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private ReadOnly connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\UniClubDB BD.accdb"

        Private Sub btnAddEvent_Click_1(sender As Object, e As EventArgs) Handles btnAddEvent.Click
        ' Validate required fields
        If txtEventName.Text.Trim() = "" OrElse
               txtVenue.Text.Trim() = "" OrElse
               txtLocation.Text.Trim() = "" OrElse
               TextBox1.Text.Trim() = "" Then

            MessageBox.Show("Please fill in all required fields.",
                                "Missing Information",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Validate capacity
        Dim capacity As Integer
        If Not Integer.TryParse(TextBox1.Text.Trim(), capacity) OrElse capacity < 0 Then
            MessageBox.Show("Please enter a valid non-negative integer for Capacity.",
                                "Invalid Capacity",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Prepare parameters for saved Access query named "AddEvent".
        ' Expected parameter order in the saved query (positional): EventID, EventName, Venue, Location, EventDate, Capacity
        ' If EventID is AutoNumber in the DB leave as DBNull to let Access auto-generate it.
        Dim eventIdValue As Object = DBNull.Value
        Dim parsedEventId As Integer
        If Integer.TryParse(TextBox2.Text.Trim(), parsedEventId) Then
            eventIdValue = parsedEventId
        End If

        Try
            Using conn As New OleDbConnection(connString)
                Using cmd As New OleDbCommand("AddEvent", conn)
                    cmd.CommandType = CommandType.StoredProcedure

                    ' Add parameters positionally. OleDb ignores parameter names for Access; order matters.
                    cmd.Parameters.AddWithValue("@EventName", txtEventName.Text.Trim())
                    cmd.Parameters.AddWithValue("@Venue", txtVenue.Text.Trim())
                    cmd.Parameters.AddWithValue("@Location", txtLocation.Text.Trim())
                    cmd.Parameters.AddWithValue("@EventDate", DateTimePicker1.Value)
                    cmd.Parameters.AddWithValue("@Capacity", capacity)

                    conn.Open()
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                    conn.Close()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Event saved successfully!",
                                            "Success",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Information)

                        ' Optionally update in-memory/UI list
                        Dim eventDetails = String.Format("{0} | {1:d} | {2}", txtEventName.Text.Trim(), DateTimePicker1.Value, txtVenue.Text.Trim())
                        EventsList.Add(eventDetails)
                        Try
                            If lstEvents IsNot Nothing Then lstEvents.Items.Add(eventDetails)
                        Catch
                            ' ignore if lstEvents is not a control here
                        End Try

                        ' Clear form
                        TextBox2.Clear()
                        txtEventName.Clear()
                        txtVenue.Clear()
                        txtLocation.Clear()
                        TextBox1.Clear()
                        DateTimePicker1.Value = Date.Now
                    Else
                        MessageBox.Show("No rows were inserted. Verify the AddEvent query and table schema.",
                                            "Insert Failed",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error saving event: " & ex.Message,
                                "Database Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClear_Click_1(sender As Object, e As EventArgs) Handles btnClear.Click
        txtEventName.Clear()
        txtVenue.Clear()
        txtLocation.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        DateTimePicker1.Value = Date.Now
    End Sub

    ' Legacy/non-handled methods retained (no Handles clause) — safe to remove if unused
    Private Sub btnAddEvent_Click(sender As Object, e As EventArgs)
        ' kept for compatibility; main handler is btnAddEvent_Click_1
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)
        ' kept for compatibility; main handler is btnClear_Click_1
    End Sub
End Class
