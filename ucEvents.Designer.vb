<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucEvents
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        FlowLayoutPanel1 = New FlowLayoutPanel()
        Label3 = New Label()
        txtEventID = New TextBox()
        lblFullName = New Label()
        txtEventName = New TextBox()
        lblEmail = New Label()
        txtVenue = New TextBox()
        lblSpeaker = New Label()
        txtLocation = New TextBox()
        Label1 = New Label()
        DateTimePicker1 = New DateTimePicker()
        Label2 = New Label()
        txtCapacity = New TextBox()
        btnAddEvent = New Button()
        btnClear = New Button()
        lblRegisterMember = New Label()
        PictureBox1 = New PictureBox()
        FlowLayoutPanel1.SuspendLayout()
        CType(PictureBox1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' FlowLayoutPanel1
        ' 
        FlowLayoutPanel1.BackColor = SystemColors.MenuHighlight
        FlowLayoutPanel1.Controls.Add(Label3)
        FlowLayoutPanel1.Controls.Add(txtEventID)
        FlowLayoutPanel1.Controls.Add(lblFullName)
        FlowLayoutPanel1.Controls.Add(txtEventName)
        FlowLayoutPanel1.Controls.Add(lblEmail)
        FlowLayoutPanel1.Controls.Add(txtVenue)
        FlowLayoutPanel1.Controls.Add(lblSpeaker)
        FlowLayoutPanel1.Controls.Add(txtLocation)
        FlowLayoutPanel1.Controls.Add(Label1)
        FlowLayoutPanel1.Controls.Add(DateTimePicker1)
        FlowLayoutPanel1.Controls.Add(Label2)
        FlowLayoutPanel1.Controls.Add(txtCapacity)
        FlowLayoutPanel1.Location = New Point(147, 49)
        FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        FlowLayoutPanel1.Size = New Size(332, 312)
        FlowLayoutPanel1.TabIndex = 22
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Font = New Font("Segoe UI Semibold", 10.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label3.ForeColor = SystemColors.Control
        Label3.Location = New Point(3, 0)
        Label3.Name = "Label3"
        Label3.Size = New Size(75, 23)
        Label3.TabIndex = 28
        Label3.Text = "Event ID"
        ' 
        ' txtEventID
        ' 
        txtEventID.Location = New Point(84, 3)
        txtEventID.Name = "txtEventID"
        txtEventID.Size = New Size(238, 27)
        txtEventID.TabIndex = 29
        ' 
        ' lblFullName
        ' 
        lblFullName.AutoSize = True
        lblFullName.Font = New Font("Segoe UI Semibold", 10.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblFullName.ForeColor = SystemColors.Control
        lblFullName.Location = New Point(3, 33)
        lblFullName.Name = "lblFullName"
        lblFullName.Size = New Size(104, 23)
        lblFullName.TabIndex = 9
        lblFullName.Text = "Event Name"
        ' 
        ' txtEventName
        ' 
        txtEventName.Location = New Point(3, 59)
        txtEventName.Name = "txtEventName"
        txtEventName.Size = New Size(319, 27)
        txtEventName.TabIndex = 17
        ' 
        ' lblEmail
        ' 
        lblEmail.AutoSize = True
        lblEmail.Font = New Font("Segoe UI Semibold", 10.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblEmail.ForeColor = SystemColors.Control
        lblEmail.Location = New Point(3, 89)
        lblEmail.Name = "lblEmail"
        lblEmail.Size = New Size(58, 23)
        lblEmail.TabIndex = 18
        lblEmail.Text = "Venue"
        ' 
        ' txtVenue
        ' 
        txtVenue.Location = New Point(3, 115)
        txtVenue.Name = "txtVenue"
        txtVenue.Size = New Size(319, 27)
        txtVenue.TabIndex = 16
        ' 
        ' lblSpeaker
        ' 
        lblSpeaker.AutoSize = True
        lblSpeaker.Font = New Font("Segoe UI Semibold", 10.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblSpeaker.ForeColor = SystemColors.Control
        lblSpeaker.Location = New Point(3, 145)
        lblSpeaker.Name = "lblSpeaker"
        lblSpeaker.Size = New Size(75, 23)
        lblSpeaker.TabIndex = 19
        lblSpeaker.Text = "Location"
        ' 
        ' txtLocation
        ' 
        txtLocation.Location = New Point(3, 171)
        txtLocation.Name = "txtLocation"
        txtLocation.Size = New Size(319, 27)
        txtLocation.TabIndex = 20
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI Semibold", 10.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label1.ForeColor = SystemColors.Control
        Label1.Location = New Point(3, 201)
        Label1.Name = "Label1"
        Label1.Size = New Size(46, 23)
        Label1.TabIndex = 21
        Label1.Text = "Date"
        ' 
        ' DateTimePicker1
        ' 
        DateTimePicker1.Location = New Point(2, 226)
        DateTimePicker1.Margin = New Padding(2)
        DateTimePicker1.Name = "DateTimePicker1"
        DateTimePicker1.Size = New Size(320, 27)
        DateTimePicker1.TabIndex = 20
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI Semibold", 10.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label2.ForeColor = SystemColors.Control
        Label2.Location = New Point(3, 255)
        Label2.Name = "Label2"
        Label2.Size = New Size(76, 23)
        Label2.TabIndex = 26
        Label2.Text = "Capacity"
        ' 
        ' txtCapacity
        ' 
        txtCapacity.Location = New Point(3, 281)
        txtCapacity.Name = "txtCapacity"
        txtCapacity.Size = New Size(319, 27)
        txtCapacity.TabIndex = 27
        ' 
        ' btnAddEvent
        ' 
        btnAddEvent.BackColor = Color.SkyBlue
        btnAddEvent.FlatStyle = FlatStyle.Popup
        btnAddEvent.Font = New Font("Segoe UI Semibold", 10.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnAddEvent.Location = New Point(192, 366)
        btnAddEvent.Margin = New Padding(2)
        btnAddEvent.Name = "btnAddEvent"
        btnAddEvent.Size = New Size(117, 35)
        btnAddEvent.TabIndex = 21
        btnAddEvent.Text = "ADD"
        btnAddEvent.UseVisualStyleBackColor = False
        ' 
        ' btnClear
        ' 
        btnClear.Font = New Font("Segoe UI", 10.2F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        btnClear.Location = New Point(313, 366)
        btnClear.Margin = New Padding(2)
        btnClear.Name = "btnClear"
        btnClear.Size = New Size(112, 35)
        btnClear.TabIndex = 22
        btnClear.Text = "CLEAR"
        btnClear.UseVisualStyleBackColor = True
        ' 
        ' lblRegisterMember
        ' 
        lblRegisterMember.AutoSize = True
        lblRegisterMember.Font = New Font("Imprint MT Shadow", 16.2F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblRegisterMember.ForeColor = SystemColors.Highlight
        lblRegisterMember.Location = New Point(232, 13)
        lblRegisterMember.Name = "lblRegisterMember"
        lblRegisterMember.Size = New Size(193, 33)
        lblRegisterMember.TabIndex = 23
        lblRegisterMember.Text = "Event Details"
        lblRegisterMember.TextAlign = ContentAlignment.TopCenter
        ' 
        ' PictureBox1
        ' 
        PictureBox1.Image = My.Resources.Resources.group
        PictureBox1.Location = New Point(203, 11)
        PictureBox1.Name = "PictureBox1"
        PictureBox1.Size = New Size(32, 32)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.TabIndex = 24
        PictureBox1.TabStop = False
        ' 
        ' ucEvents
        ' 
        AutoScaleDimensions = New SizeF(8F, 20F)
        AutoScaleMode = AutoScaleMode.Font
        Controls.Add(PictureBox1)
        Controls.Add(lblRegisterMember)
        Controls.Add(FlowLayoutPanel1)
        Controls.Add(btnClear)
        Controls.Add(btnAddEvent)
        Margin = New Padding(2)
        Name = "ucEvents"
        Size = New Size(617, 418)
        FlowLayoutPanel1.ResumeLayout(False)
        FlowLayoutPanel1.PerformLayout()
        CType(PictureBox1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents lblFullName As Label
    Friend WithEvents txtEventName As TextBox
    Friend WithEvents lblEmail As Label
    Friend WithEvents txtVenue As TextBox
    Friend WithEvents lblSpeaker As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents btnAddEvent As Button
    Friend WithEvents btnClear As Button
    Friend WithEvents txtLocation As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblRegisterMember As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtCapacity As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtEventID As TextBox
    Friend WithEvents PictureBox1 As PictureBox

End Class
