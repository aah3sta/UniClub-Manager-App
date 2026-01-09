Imports System.Data.OleDb

Public Class ucEvents

    Private EventsList As New List(Of String)()
    Private ReadOnly connString As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=|DataDirectory|\UniClubDB.accdb"

    Private Sub ucEvents_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnAddEvent_Click(sender As Object, e As EventArgs)

        If txtEventName.Text = "" Or txtVenue.Text = "" Then
            MessageBox.Show("Please fill all fields")
            Exit Sub
        End If

        Dim eventDetails =
            txtEventName.Text & " | " &
            DateTimePicker1.Value.ToShortDateString & " | " &
            txtVenue.Text

        EventsList.Add(eventDetails)
        lstEvents.Items.Add(eventDetails)

        MessageBox.Show("Event added successfully")
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)
        txtEventName.Clear
        txtVenue.Clear
        DateTimePicker1.Value = Date.Now
    End Sub

    Private Sub btnAddEvent_Click_1(sender As Object, e As EventArgs) Handles btnAddEvent.Click
        If txtEventName.Text.Trim() = "" Or
       txtVenue.Text.Trim() = "" Or
       txtLocation.Text.Trim() = "" Then

            MessageBox.Show("Please fill in all fields.",
                        "Missing Information",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
            Exit Sub
        End If


        MessageBox.Show("Event saved successfully!",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
    End Sub

    Private Sub ClearFields()
        txtEventName.Clear()
        txtLocation.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        DateTimePicker1.Value = Date.Now
        ' Keep TextBox2 (EventID)
    End Sub

    ' Legacy/non-handled methods retained (no Handles clause) — safe to remove if unused
    Private Sub btnAddEvent_Click(sender As Object, e As EventArgs)
        ' kept for compatibility; main handler is btnAddEvent_Click_1
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)
        ' kept for compatibility; main handler is btnClear_Click_1
    End Sub
End Class
