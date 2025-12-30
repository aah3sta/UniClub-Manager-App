Imports System.Data.OleDb
Public Class Form1
    Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\UniClubDB BD.accdb"
    Dim conn As New OleDbConnection(connString)

    Public Sub LoadRecentActivity()
        Try
            If conn.State = ConnectionState.Open Then conn.Close()
            conn.Open()

            Dim sql As String = "SELECT TOP 3 FullName, Department, PhoneNumber, Course FROM Members ORDER BY JoinDate DESC"
            Dim adapter As New OleDbDataAdapter(sql, conn)
            Dim dt As New DataTable()

            adapter.Fill(dt)
            dgvRecentMembers.DataSource = dt
            conn.Close()

        Catch ex As Exception
            MessageBox.Show("Could not load recent activity: " & ex.Message)
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadRecentActivity()

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

    Private Sub picClose_Click(sender As Object, e As EventArgs) Handles picClose.Click
        Close()
    End Sub

    Private Sub picMinimize_Click(sender As Object, e As EventArgs) Handles picMinimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub picMaximize_Click(sender As Object, e As EventArgs) Handles picMaximize.Click
        If Me.WindowState = FormWindowState.Normal Then
            Me.WindowState = FormWindowState.Maximized
        ElseIf Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
        End If
    End Sub

    Public Sub SwitchScreen(ByVal newScreen As UserControl)
        pnlMain.Controls.Clear()
        newScreen.Dock = DockStyle.Fill
        pnlMain.Controls.Add(newScreen)
    End Sub

    Private Sub btnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
        Dim regScreen As New ucRegister
        SwitchScreen(regScreen)
    End Sub

    Private Sub BtnAttendance_Click(sender As Object, e As EventArgs) Handles btnAttendance.Click
        Dim attendancePage As New ucAttendance()
        SwitchScreen(attendancePage)
    End Sub

    Private Sub btnEvents_Click(sender As Object, e As EventArgs) Handles btnEvents.Click
        Dim eventsPage As New ucEvents()
        SwitchScreen(eventsPage)
    End Sub

    Private Sub btnDashboard_Click(sender As Object, e As EventArgs) Handles btnDashboard.Click
        Dim dashboardPage As New ucDashboard()
        SwitchScreen(dashboardPage)
    End Sub
End Class
