Imports System.Data.OleDb
Imports System.Data

Public Class ucRegister

    Private ReadOnly connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\UniClubDB.accdb"

    Private Sub ucRegister_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnSubmitRegistration_Click(sender As Object, e As EventArgs) Handles btnSubmitRegistration.Click
        ' Validate required fields
        If txtFullName.Text.Trim() = "" OrElse
         txtEmail.Text.Trim() = "" OrElse
         txtCourse.Text.Trim() = "" OrElse
         txtDepartment.Text.Trim() = "" Then

            MessageBox.Show("Please fill in all required fields", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Basic email check
        If Not txtEmail.Text.Contains("@") OrElse Not txtEmail.Text.Contains(".") Then
            MessageBox.Show("Please enter a valid email address", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Try
            ' Use brackets to avoid reserved word conflicts; explicit parameter types to avoid type inference issues
            ' Exclude JoinDate so the database default (auto timestamp) is used
            Dim sql As String = "INSERT INTO Members ([FullName], [Email], [Course], [Department], [PhoneNumber]) VALUES (@FullName, @Email, @Course, @Department, @PhoneNumber)"
            Using conn As New OleDbConnection(connString)
                Using cmd As New OleDbCommand(sql, conn)
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Add(New OleDbParameter("@FullName", OleDbType.VarWChar)).Value = txtFullName.Text.Trim()
                    cmd.Parameters.Add(New OleDbParameter("@Email", OleDbType.VarWChar)).Value = txtEmail.Text.Trim()
                    cmd.Parameters.Add(New OleDbParameter("@Course", OleDbType.VarWChar)).Value = txtCourse.Text.Trim()
                    cmd.Parameters.Add(New OleDbParameter("@Department", OleDbType.VarWChar)).Value = txtDepartment.Text.Trim()

                    ' PhoneNumber is Short Text in Access; pass DBNull if empty
                    Dim phoneParam As New OleDbParameter("@PhoneNumber", OleDbType.VarWChar)
                    If String.IsNullOrWhiteSpace(TextBox1.Text) Then
                        phoneParam.Value = DBNull.Value
                    Else
                        phoneParam.Value = TextBox1.Text.Trim()
                    End If
                    cmd.Parameters.Add(phoneParam)

                    conn.Open()
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                    conn.Close()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Member registered successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ' Clear form
                        txtFullName.Clear()
                        txtEmail.Clear()
                        txtCourse.Clear()
                        txtDepartment.Clear()
                        TextBox1.Clear()
                        TextBox2.Clear()
                    Else
                        MessageBox.Show("Registration failed: no rows were inserted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error saving member: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCancelRegistration_Click(sender As Object, e As EventArgs) Handles btnCancelRegistration.Click
        txtFullName.Clear()
        txtEmail.Clear()
        txtCourse.Clear()
        txtDepartment.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

End Class