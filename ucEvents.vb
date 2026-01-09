Imports System.Data.OleDb

Public Class ucEvents

    Private EventsList As New List(Of String)()
    Private ReadOnly connString As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=|DataDirectory|\UniClubDB.accdb"

    Private Sub ucEvents_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadEvents()
    End Sub

    Private Sub btnAddEvent_Click(sender As Object, e As EventArgs) Handles btnAddEvent.Click
        If txtEventName.Text.Trim() = "" OrElse
           txtLocation.Text.Trim() = "" Then

            MessageBox.Show("Please fill in all required fields (Event Name and Location).",
                        "Missing Information",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
            Exit Sub
        End If

        Try
            Using conn As New OleDbConnection(connString)
                conn.Open()

                ' Determine available columns in Events table
                Dim schema As DataTable = conn.GetSchema("Columns")
                Dim rows() As DataRow = schema.Select("TABLE_NAME = 'Events' OR TABLE_NAME = 'events'")
                Dim cols As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each r As DataRow In rows
                    cols.Add(Convert.ToString(r("COLUMN_NAME")))
                Next

                ' Map UI fields to DB column names (EventName, EventLocation, EventDate)
                Dim insertCols As New List(Of String)()
                Dim paramValues As New List(Of OleDbParameter)()

                If cols.Contains("EventName") Then
                    insertCols.Add("EventName")
                    paramValues.Add(New OleDbParameter("?", OleDbType.VarWChar) With {.Value = txtEventName.Text.Trim()})
                End If

                If cols.Contains("EventLocation") Then
                    insertCols.Add("EventLocation")
                    paramValues.Add(New OleDbParameter("?", OleDbType.VarWChar) With {.Value = txtLocation.Text.Trim()})
                ElseIf cols.Contains("Location") Then
                    insertCols.Add("Location")
                    paramValues.Add(New OleDbParameter("?", OleDbType.VarWChar) With {.Value = txtLocation.Text.Trim()})
                End If

                If cols.Contains("EventDate") Then
                    insertCols.Add("EventDate")
                    paramValues.Add(New OleDbParameter("?", OleDbType.Date) With {.Value = DateTimePicker1.Value})
                ElseIf cols.Contains("Date") Then
                    insertCols.Add("Date")
                    paramValues.Add(New OleDbParameter("?", OleDbType.Date) With {.Value = DateTimePicker1.Value})
                End If

                If insertCols.Count = 0 Then
                    MessageBox.Show("No matching columns were found in the Events table. Check database schema.", "Schema Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim sbCols As New System.Text.StringBuilder()
                Dim sbParams As New System.Text.StringBuilder()
                For i As Integer = 0 To insertCols.Count - 1
                    If i > 0 Then
                        sbCols.Append(", ")
                        sbParams.Append(", ")
                    End If
                    sbCols.Append(insertCols(i))
                    sbParams.Append("?")
                Next

                Dim sql As String = $"INSERT INTO Events ({sbCols}) VALUES ({sbParams})"

                Using cmd As New OleDbCommand(sql, conn)
                    For Each p As OleDbParameter In paramValues
                        cmd.Parameters.Add(p)
                    Next

                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        ' If EventID is AutoNumber, obtain the last identity
                        If cols.Contains("EventID") Then
                            Using idCmd As New OleDbCommand("SELECT @@IDENTITY", conn)
                                Dim idObj = idCmd.ExecuteScalar()
                                If idObj IsNot Nothing AndAlso Not IsDBNull(idObj) Then
                                    TextBox2.Text = Convert.ToInt32(idObj).ToString()
                                End If
                            End Using
                        End If

                        MessageBox.Show("Event saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ClearFields()
                        LoadEvents()
                    Else
                        MessageBox.Show("No rows were inserted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error saving event: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadEvents()
        EventsList.Clear()
        If lstEvents IsNot Nothing Then
            lstEvents.Items.Clear()
        End If

        Try
            Using conn As New OleDbConnection(connString)
                conn.Open()

                ' Determine which columns exist and build SELECT accordingly
                Dim schema As DataTable = conn.GetSchema("Columns")
                Dim rows() As DataRow = schema.Select("TABLE_NAME = 'Events' OR TABLE_NAME = 'events'")
                Dim cols As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each r As DataRow In rows
                    cols.Add(Convert.ToString(r("COLUMN_NAME")))
                Next

                Dim selectCols As New List(Of String)()
                If cols.Contains("EventID") Then selectCols.Add("EventID")
                If cols.Contains("EventName") Then selectCols.Add("EventName")
                If cols.Contains("EventLocation") Then selectCols.Add("EventLocation")
                If cols.Contains("Location") And Not selectCols.Contains("EventLocation") Then selectCols.Add("Location")
                If cols.Contains("EventDate") Then selectCols.Add("EventDate")
                If cols.Contains("Date") And Not selectCols.Contains("EventDate") Then selectCols.Add("Date")

                If selectCols.Count = 0 Then
                    Return
                End If

                Dim sb As New System.Text.StringBuilder()
                sb.Append("SELECT ")
                sb.Append(String.Join(", ", selectCols))
                sb.Append(" FROM Events")

                Using cmd As New OleDbCommand(sb.ToString(), conn)
                    Using rdr As OleDbDataReader = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim parts As New List(Of String)()
                            For i As Integer = 0 To rdr.FieldCount - 1
                                If rdr.IsDBNull(i) Then
                                    parts.Add("")
                                Else
                                    parts.Add(rdr.GetValue(i).ToString())
                                End If
                            Next
                            Dim line = String.Join(" | ", parts)
                            EventsList.Add(line)
                            If lstEvents IsNot Nothing Then lstEvents.Items.Add(line)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.Print("LoadEvents error: " & ex.Message)
        End Try
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ClearFields()
    End Sub

    Private Sub ClearFields()
        txtEventName.Clear()
        txtLocation.Clear()
        DateTimePicker1.Value = Date.Now
        ' Keep TextBox2 (EventID)
    End Sub
End Class
