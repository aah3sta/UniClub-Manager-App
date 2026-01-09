Imports System.Data.OleDb
Public Class Form1
    ' Remove hardcoded provider; use DbConfig.ConnString
    Dim conn As OleDbConnection

    Public Sub LoadRecentActivity()
        Try
            If conn Is Nothing Then
                conn = New OleDbConnection(DbConfig.ConnString)
            End If
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
            If conn IsNot Nothing AndAlso conn.State = ConnectionState.Open Then conn.Close()
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Ensure |DataDirectory| points to the running exe folder so the DB resolves reliably
        AppDomain.CurrentDomain.SetData("DataDirectory", Application.StartupPath)

        ' Build a runtime connection string that matches installed provider
        DbConfig.ConnString = ResolveConnectionString()

        ' Initialize connection once resolved
        conn = New OleDbConnection(DbConfig.ConnString)

        LoadRecentActivity()

        Try
            conn.Open()
            Dim cmd As New OleDbCommand("SELECT COUNT(*) FROM Members", conn)
            Dim count As Integer = CInt(cmd.ExecuteScalar())
            lblMemberCount.Text = count.ToString()
            conn.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading stats: " & ex.Message)
            If conn IsNot Nothing AndAlso conn.State = ConnectionState.Open Then conn.Close()
        End Try

        Try
            conn.Open()
            Dim cmd As New OleDbCommand("SELECT COUNT(*) FROM Events", conn)
            Dim count As Integer = CInt(cmd.ExecuteScalar())
            lblEventCount.Text = count.ToString()
            conn.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading stats: " & ex.Message)
            If conn IsNot Nothing AndAlso conn.State = ConnectionState.Open Then conn.Close()
        End Try
    End Sub

    Private Function ResolveConnectionString() As String
        ' Try providers in order; adjust to installed bits
        Dim dbPath As String = System.IO.Path.Combine(Application.StartupPath, "UniClubDB.accdb")
        Dim providers As String() = {"Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0"}
        For Each prov In providers
            Dim cs As String = $"Provider={prov};Data Source={dbPath}"
            Try
                Using c As New OleDbConnection(cs)
                    c.Open()
                    c.Close()
                End Using
                Return cs
            Catch
                ' try next
            End Try
        Next

        ' Fallback to ODBC driver
        Dim odbc As String = $"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={dbPath};"
        Try
            Using c As New OleDbConnection(odbc)
                c.Open()
                c.Close()
            End Using
            Return odbc
        Catch ex As Exception
            ' As last resort return a prov-12.0 string so the app fails with clear message
            Return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}"
        End Try
    End Function

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
