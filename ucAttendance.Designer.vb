<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ucAttendance
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        dgvAttendance = New DataGridView()
        Column1 = New DataGridViewTextBoxColumn()
        Column2 = New DataGridViewTextBoxColumn()
        Column3 = New DataGridViewTextBoxColumn()
        dtpMeetingDate = New DateTimePicker()
        btnClearAttendance = New Button()
        btnSave = New Button()
        txtSearchName = New TextBox()
        btnSearch = New Button()
        lstAttendance = New ListBox()
        CType(dgvAttendance, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' dgvAttendance
        ' 
        dgvAttendance.AllowUserToOrderColumns = True
        dgvAttendance.BackgroundColor = SystemColors.Highlight
        dgvAttendance.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvAttendance.Columns.AddRange(New DataGridViewColumn() {Column1, Column2, Column3})
        dgvAttendance.Location = New Point(68, 142)
        dgvAttendance.Margin = New Padding(2)
        dgvAttendance.Name = "dgvAttendance"
        dgvAttendance.RowHeadersWidth = 62
        dgvAttendance.Size = New Size(632, 302)
        dgvAttendance.TabIndex = 0
        ' 
        ' Column1
        ' 
        Column1.HeaderText = "NAME"
        Column1.MinimumWidth = 8
        Column1.Name = "Column1"
        Column1.Width = 200
        ' 
        ' Column2
        ' 
        Column2.HeaderText = "ATTENDANCE"
        Column2.MinimumWidth = 8
        Column2.Name = "Column2"
        Column2.Width = 200
        ' 
        ' Column3
        ' 
        Column3.HeaderText = "SCHOOL"
        Column3.MinimumWidth = 8
        Column3.Name = "Column3"
        Column3.Width = 150
        ' 
        ' dtpMeetingDate
        ' 
        dtpMeetingDate.Location = New Point(399, 34)
        dtpMeetingDate.Margin = New Padding(2)
        dtpMeetingDate.Name = "dtpMeetingDate"
        dtpMeetingDate.Size = New Size(300, 31)
        dtpMeetingDate.TabIndex = 1
        ' 
        ' btnClearAttendance
        ' 
        btnClearAttendance.FlatStyle = FlatStyle.Popup
        btnClearAttendance.Location = New Point(416, 451)
        btnClearAttendance.Margin = New Padding(2)
        btnClearAttendance.Name = "btnClearAttendance"
        btnClearAttendance.Size = New Size(128, 45)
        btnClearAttendance.TabIndex = 3
        btnClearAttendance.Text = "CLEAR"
        btnClearAttendance.UseVisualStyleBackColor = True
        ' 
        ' btnSave
        ' 
        btnSave.BackColor = Color.SkyBlue
        btnSave.FlatStyle = FlatStyle.Popup
        btnSave.Font = New Font("Segoe UI Semibold", 10.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnSave.Location = New Point(224, 451)
        btnSave.Margin = New Padding(2)
        btnSave.Name = "btnSave"
        btnSave.Size = New Size(130, 45)
        btnSave.TabIndex = 4
        btnSave.Text = "SAVE"
        btnSave.UseVisualStyleBackColor = False
        ' 
        ' txtSearchName
        ' 
        txtSearchName.Location = New Point(68, 36)
        txtSearchName.Margin = New Padding(2)
        txtSearchName.Name = "txtSearchName"
        txtSearchName.Size = New Size(206, 31)
        txtSearchName.TabIndex = 6
        ' 
        ' btnSearch
        ' 
        btnSearch.Location = New Point(68, 85)
        btnSearch.Margin = New Padding(2)
        btnSearch.Name = "btnSearch"
        btnSearch.Size = New Size(206, 31)
        btnSearch.TabIndex = 7
        btnSearch.Text = "Search"
        btnSearch.TextAlign = ContentAlignment.MiddleLeft
        btnSearch.UseVisualStyleBackColor = True
        ' 
        ' lstAttendance
        ' 
        lstAttendance.FormattingEnabled = True
        lstAttendance.ItemHeight = 25
        lstAttendance.Location = New Point(68, 75)
        lstAttendance.Margin = New Padding(2)
        lstAttendance.Name = "lstAttendance"
        lstAttendance.Size = New Size(206, 4)
        lstAttendance.TabIndex = 8
        ' 
        ' ucAttendance
        ' 
        AutoScaleDimensions = New SizeF(8.0F, 20.0F)
        AutoScaleMode = AutoScaleMode.Font
        Controls.Add(lstAttendance)
        Controls.Add(btnSearch)
        Controls.Add(txtSearchName)
        Controls.Add(btnSave)
        Controls.Add(btnClearAttendance)
        Controls.Add(dtpMeetingDate)
        Controls.Add(dgvAttendance)
        Margin = New Padding(2)
        Name = "ucAttendance"
        Size = New Size(771, 522)
        CType(dgvAttendance, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents dgvAttendance As DataGridView
    Friend WithEvents dtpMeetingDate As DateTimePicker
    Friend WithEvents btnClearAttendance As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents txtSearchName As TextBox
    Friend WithEvents btnSearch As Button
    Friend WithEvents lstAttendance As ListBox
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn

End Class
