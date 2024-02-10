Imports MySqlConnector
Imports System.Timers
Imports RJCP.IO.Ports
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Drawing.Printing
Imports System.IO.Compression
Imports OfficeOpenXml
Imports OfficeOpenXml.LicenseContext
Imports ExcelDataReader
Imports System.Text
Imports System.Data
Imports SharpCompress.Archives
Imports SharpCompress.Common


Public Class MainPage
    'server=localhost; user=yout_database_user; password=your_database_password; database=your_database_name
    'Dim Connection As New MySqlConnection("server=localhost; user=root; password=; database=project")'
    Dim Connection As New MySqlConnection("Data Source=localhost;Initial Catalog=monitorsystem;User ID=root;Password=")
    Dim connectionString As String = "Data Source=localhost;Initial Catalog=monitorsystem;User ID=root;Password="

    Dim MySQLCMD As New MySqlCommand
    Dim MySQLDA As New MySqlDataAdapter
    Dim DT As New DataTable
    Dim Table_Name As String = "student_list"
    Dim Table_Name2 As String = "attendance"
    Dim Table_Name3 As String = "image"
    Dim Table_Name4 As String = "sms_messages"
    Dim IMG_FileNameInput As String
    Dim LoadImagesStr As Boolean = False
    Private serialPort As SerialPortStream
    Dim Data As Integer
    Dim IDRam As String
    Dim StatusInput As String = "Save"
    Dim SqlCmdSearchstr As String

    Public Shared StrSerialIn As String
    Dim GetID As Boolean = False
    Dim ViewUserData As Boolean = False

    Public Property Cancel As Boolean

    Private rowIndex As Integer = 0

    Private Sub MainPage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            PanelDashboard.Visible = True
            PanelSearch.Visible = False
            PanelConnection.Visible = False
            PanelUserDetails.Visible = False
            PanelRegistration.Visible = False
            PanelStudentReport.Visible = False
            PanelMessageReport.Visible = False
            PanelAttendanceReport.Visible = False
            PanelExport1.Visible = False
            PanelExport2.Visible = False
            PanelExport3.Visible = False
            PanelRetrieve.Visible = False
            PanelDropdown.Height = 0

            AddHandler DataGridView1.DataBindingComplete, AddressOf DataGridView1_DataBindingComplete

            ComboBoxBaudRate.SelectedIndex = 3
            ComboBoxYearlevelinRegistration.SelectedIndex = 0
            ComboBoxYearlevelinAttendanceReport.SelectedIndex = 0
            ComboBoxStatusinAttendanceReport.SelectedIndex = 0
            ComboBoxDate.SelectedIndex = 0
            ComboBoxDateinAttendance.SelectedIndex = 0
            ComboBoxDepartmentinRetrieve.SelectedIndex = 0
            ComboBoxYearLevel.SelectedIndex = 0

            ComboBoxDateinSearchPanel.SelectedIndex = 0
            ComboBoxStatusinSearchPanel.SelectedIndex = 0
            ComboBoxDepartmentinSearchPanel.SelectedIndex = 0
            Dim dateFilter2 As String = ComboBoxDateinSearchPanel.SelectedItem.ToString()
            Dim departmentFilter2 As String = ComboBoxDepartmentinSearchPanel.SelectedItem.ToString()
            Dim statusFilter2 As String = ComboBoxStatusinSearchPanel.SelectedItem.ToString()
            Dim searchKeyword2 As String = SearchTextBoxinSearchPanel.Text
            Dim startDate2 As Date = DateTimePickerSearchPanelStart.Value.Date
            Dim endDate2 As Date = DateTimePickerSearchPanelEnd.Value.Date

            Dim newData2 As DataTable = ShowDataInSearch(dateFilter2, departmentFilter2, statusFilter2, searchKeyword2, startDate2, endDate2)
            DataGridView2.DataSource = newData2

            ComboBoxYearlevelinRegistration.SelectedIndex = 0
            Dim yearLevel3 As String = ComboBoxYearlevelinRegistration.SelectedItem.ToString()
            Dim newData3 As DataTable = ShowDataInRegistration(yearLevel3, "")
            DataGridView3.DataSource = newData3

            ComboBoxYearLevel.SelectedIndex = 0
            Dim yearLevel As String = ComboBoxYearLevel.SelectedItem.ToString()
            Dim newData4 As DataTable = ShowDataByYearLevel(yearLevel, False, False, "")
            DataGridView4.DataSource = newData4

            ComboBoxInMessageReport.SelectedIndex = 0
            Dim successOrFail As String = ComboBoxInMessageReport.SelectedItem.ToString()
            Dim dateFilter As String = ComboBoxDate.SelectedItem.ToString()
            Dim newData As DataTable = ShowDataInMessageReport(successOrFail, "", dateFilter)
            DataGridView5.DataSource = newData

            Dim dateFilter6 As String = ComboBoxDateinAttendance.SelectedItem.ToString()
            Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
            Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
            Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
            Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

            Dim newData6 As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
            DataGridView6.DataSource = newData6

            'Call Count Function
            CountSuccessMessages()
            CountFailMessages()
            CountAllStudents()
            CountCollegeStudents()
            CountSeniorHighStudents()
            CountAllActiveStudentToday()
            CountAllINStudentToday()
            CountAllOUTStudentToday()
            CountAllSMS()

            'Timers
            RefreshCountActive().Interval = 5000 ' 5 seconds
            RefreshCountActive().Start()
            RefreshCountIn().Interval = 5000 ' 5 seconds
            RefreshCountIn().Start()
            RefreshCountOut().Interval = 5000 ' 5 seconds
            RefreshCountOut().Start()
            RefreshSuccessSMS.Interval = 5000 ' 5 seconds
            RefreshSuccessSMS.Start()
            RefreshFailSMS.Interval = 5000 ' 5 seconds
            RefreshFailSMS.Start()
            RefreshSMStotal.Interval = 5000 ' 5 seconds
            RefreshSMStotal.Start()
            TimerDateandTime.Interval = 1000 ' 1 second
            TimerDateandTime.Start()
            Timer1.Interval = 1000 ' 1 second
            Timer1.Enabled = True

            DataGridView7.DefaultCellStyle.ForeColor = Color.Black
            DataGridView7.ClearSelection()
            DataGridView7.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView7.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
        Catch ex As Exception
            MessageBox.Show("An error occurred in the MainPage_Load event: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Load all datagridview
    Private Sub LoadDataforDatagridview()
        Try
            ' Show the loading panel
            PanelLoading.Visible = True

            ' Retrieve the data from the database
            Dim dataTable As DataTable = ShowDatainDashboard()
            DataGridView1.DataSource = dataTable

            ' Retrieve the data for DataGridView4
            ComboBoxYearLevel_SelectedIndexChanged(ComboBoxYearLevel, EventArgs.Empty)

            ' Retrieve the data for DataGridView5
            ComboBoxInMessageReport_SelectedIndexChanged(ComboBoxInMessageReport, EventArgs.Empty)

            ' Retrieve the data for DataGridView5
            ComboBoxYearlevelinRegistration_SelectedIndexChanged(ComboBoxYearlevelinRegistration, EventArgs.Empty)
        Catch ex As Exception
            MessageBox.Show("An error occurred in the LoadDataforDatagridview method: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Hide the loading panel
            PanelLoading.Visible = False
        End Try
    End Sub

    '*************************************************
    '
    '
    '   All the Component of Dashboard Panel
    '
    '
    '*************************************************

    'Show data in DatagridView1 in Dashboard Panel'
    Private Function ShowDatainDashboard() As DataTable
        Dim dataTable As New DataTable()
        Try
            Dim philippineTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Manila")
            Dim philippinesCurrentTime As DateTimeOffset = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, philippineTimeZone)
            Dim philippinesDate As String = philippinesCurrentTime.ToString("yyyy-MM-dd")

            Dim query As String = "SELECT a.RFID_UID, a.Student_ID, a.Student_Name, a.Department, a.Course, a.Year, a.Status AS Attendance_Status, a.Time, a.Date, s.Status AS Sms_Status
                                    FROM attendance a
                                    INNER JOIN sms_messages s ON a.ID = s.Attendance_ID
                                    WHERE a.Date = '" & philippinesDate & "'
                                    ORDER BY a.Time DESC;
                                    "

            Using connection As New MySqlConnection(connectionString)
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred in the ShowDatainDashboard function: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dataTable
    End Function

    'Count all the Active Student in current day
    Public Sub CountAllActiveStudentToday()
        Dim philippineTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Manila")
        Dim philippinesCurrentTime As DateTimeOffset = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, philippineTimeZone)
        Dim philippinesDate As String = philippinesCurrentTime.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT COUNT(DISTINCT Student_ID) FROM attendance WHERE DATE(Date) = '" & philippinesDate & "';"

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim successCount As Integer = Convert.ToInt32(command.ExecuteScalar())

                    ' Assuming you have a Label control named LabelTotalAttendance on your form
                    LabelTotalAttendance.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting all active students today: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshCountActive_Tick(sender As Object, e As EventArgs) Handles RefreshCountActive.Tick
        CountAllActiveStudentToday()
    End Sub

    'Count all the Student that inside of school
    Public Sub CountAllINStudentToday()
        Dim philippineTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Manila")
        Dim philippinesCurrentTime As DateTimeOffset = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, philippineTimeZone)
        Dim philippinesDate As String = philippinesCurrentTime.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT SUM(Out_Status_Count) AS Total_Status_Count
                            FROM (
                                SELECT Student_ID, Status, 
                                       SUM(CASE WHEN Status = 'IN' THEN 1 ELSE 0 END) AS Out_Status_Count
                                FROM (
                                    SELECT Student_ID, Status,
                                           ROW_NUMBER() OVER (PARTITION BY Student_ID ORDER BY Date DESC, Time DESC) AS rn
                                    FROM attendance
                                    WHERE DATE(Date) = '" & philippinesDate & "'
                                ) subquery
                                WHERE rn = 1
                                GROUP BY Student_ID, Status
                            ) subquery2;
                             "
        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim result As Object = command.ExecuteScalar()
                    Dim successCount As Integer = 0

                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        successCount = Convert.ToInt32(result)
                    End If

                    ' Assuming you have a Label control named LabelcountIN on your form
                    LabelcountIN.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting all students inside school today: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshCountIn_Tick(sender As Object, e As EventArgs) Handles RefreshCountIn.Tick
        CountAllINStudentToday()
    End Sub

    'Count all the student outside of school
    Public Sub CountAllOUTStudentToday()
        Dim philippineTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Manila")
        Dim philippinesCurrentTime As DateTimeOffset = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, philippineTimeZone)
        Dim philippinesDate As String = philippinesCurrentTime.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT SUM(Out_Status_Count) AS Total_Status_Count
                            FROM (
                                SELECT Student_ID, Status, 
                                       SUM(CASE WHEN Status = 'OUT' THEN 1 ELSE 0 END) AS Out_Status_Count
                                FROM (
                                    SELECT Student_ID, Status,
                                           ROW_NUMBER() OVER (PARTITION BY Student_ID ORDER BY Date DESC, Time DESC) AS rn
                                    FROM attendance
                                    WHERE DATE(Date) = '" & philippinesDate & "'
                                ) subquery
                                WHERE rn = 1
                                GROUP BY Student_ID, Status
                            ) subquery2;
                            "
        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim result As Object = command.ExecuteScalar()
                    Dim successCount As Integer = 0

                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        successCount = Convert.ToInt32(result)
                    End If

                    ' Assuming you have a Label control named LabelcountOUT on your form
                    LabelcountOUT.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting all students outside school today: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshCountOut_Tick(sender As Object, e As EventArgs) Handles RefreshCountOut.Tick
        CountAllOUTStudentToday()
    End Sub

    ' Hide the loading panel
    Private Sub DataGridView1_DataBindingComplete(sender As Object, e As EventArgs)
        PanelLoading.Visible = False
    End Sub

    'Refresh DataGridView1 in every 1 second in Dashboard Panel'
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim newData As DataTable = ShowDatainDashboard()
            DataGridView1.DataSource = newData

            DataGridView1.DefaultCellStyle.ForeColor = Color.Black
            DataGridView1.ClearSelection()
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
        Catch ex As MySqlException
            MessageBox.Show("A MySQL error occurred in the Timer1_Tick event handler: " & ex.Message, "MySQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("An error occurred in the Timer1_Tick event handler: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    '*************************************************
    '
    '
    '   All the Component of Search Panel
    '
    '
    '*************************************************

    'Display Attendance Data in Datagridview2
    Private Function ShowDataInSearch(dateFilter2 As String, departmentFilter2 As String, statusFilter2 As String, searchKeyword2 As String, startDate2 As Date, endDate2 As Date) As DataTable
        Dim dataTable As New DataTable()
        Try
            Dim query As String = "SELECT a.RFID_UID, a.Student_ID, a.Student_Name, a.Department, a.Course, a.Year, a.Status AS Attendance_Status, a.Time, a.Date, s.Status AS Sms_Status
                        FROM attendance a
                        INNER JOIN sms_messages s ON a.ID = s.Attendance_ID
                        WHERE 1 = 1"

            If dateFilter2 = "Today" Then
                query &= " AND a.Date = CURDATE()"
            ElseIf dateFilter2 = "This Week" Then
                query &= " AND YEARWEEK(a.Date, 1) = YEARWEEK(CURDATE(), 1)"
            ElseIf dateFilter2 = "This Month" Then
                query &= " AND MONTH(a.Date) = MONTH(CURDATE()) AND YEAR(a.Date) = YEAR(CURDATE())"
            ElseIf dateFilter2 = "Date Range" Then
                query &= " AND a.Date BETWEEN @StartDate AND @EndDate"
            End If

            If departmentFilter2 <> "All" Then
                query &= " AND a.Department = @departmentFilter"
            End If

            If statusFilter2 <> "All" Then
                query &= " AND a.Status = @StatusFilter"
            End If

            If searchKeyword2 <> "" Then
                If CheckBoxInAttendanceReportFindByStudent_ID.Checked Then
                    query &= " AND a.Student_ID LIKE @Keyword"
                ElseIf CheckBoxInAttendanceReportFindByName.Checked Then
                    query &= " AND a.Student_Name LIKE @Keyword"
                End If
            End If

            query &= " ORDER BY a.Time DESC;"

            Using connection As New MySqlConnection(connectionString)
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    If departmentFilter2 <> "All" Then
                        command.Parameters.AddWithValue("@departmentFilter", departmentFilter2)
                    End If

                    If statusFilter2 <> "All" Then
                        command.Parameters.AddWithValue("@StatusFilter", statusFilter2)
                    End If

                    If searchKeyword2 <> "" Then
                        command.Parameters.AddWithValue("@Keyword", "%" & searchKeyword2 & "%")
                    End If

                    If dateFilter2 = "Date Range" Then
                        command.Parameters.AddWithValue("@StartDate", startDate2)
                        command.Parameters.AddWithValue("@EndDate", endDate2)
                    End If

                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End Using
            End Using

            DataGridView2.DefaultCellStyle.ForeColor = Color.Black
            DataGridView2.ClearSelection()
            DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView2.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
            DataGridView2.DataSource = dataTable ' Update DataGridView6
        Catch ex As Exception
            MessageBox.Show("An error occurred in the ShowDataInAttendanceReport function: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dataTable
    End Function


    'Search by Name
    Private Sub CheckBoxName_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxName.CheckedChanged
        If CheckBoxName.Checked = True Then
            CheckBoxStudentID.Checked = False
        End If
        If CheckBoxName.Checked = False Then
            CheckBoxStudentID.Checked = True
        End If
    End Sub

    'Search by Student_ID
    Private Sub CheckBoxStudentID_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxStudentID.CheckedChanged
        If CheckBoxStudentID.Checked = True Then
            CheckBoxName.Checked = False
        End If
        If CheckBoxStudentID.Checked = False Then
            CheckBoxName.Checked = True
        End If
    End Sub

    Private Sub DateTimePickerSearchPanelStart_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerSearchPanelStart.ValueChanged
        If ComboBoxDepartmentinSearchPanel.SelectedItem IsNot Nothing Then
            Dim dateFilter2 As String = ComboBoxDateinSearchPanel.SelectedItem.ToString()
            Dim departmentFilter2 As String = ComboBoxDepartmentinSearchPanel.SelectedItem.ToString()
            Dim statusFilter2 As String = ComboBoxStatusinSearchPanel.SelectedItem.ToString()
            Dim searchKeyword2 As String = SearchTextBoxinSearchPanel.Text
            Dim startDate2 As Date = DateTimePickerSearchPanelStart.Value.Date
            Dim endDate2 As Date = DateTimePickerSearchPanelEnd.Value.Date

            Dim newData2 As DataTable = ShowDataInSearch(dateFilter2, departmentFilter2, statusFilter2, searchKeyword2, startDate2, endDate2)
            DataGridView2.DataSource = newData2
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem2.Click
        Dim dateFilter2 As String = ComboBoxDateinSearchPanel.SelectedItem.ToString()
        Dim departmentFilter2 As String = ComboBoxDepartmentinSearchPanel.SelectedItem.ToString()
        Dim statusFilter2 As String = ComboBoxStatusinSearchPanel.SelectedItem.ToString()
        Dim searchKeyword2 As String = SearchTextBoxinSearchPanel.Text
        Dim startDate2 As Date = DateTimePickerSearchPanelStart.Value.Date
        Dim endDate2 As Date = DateTimePickerSearchPanelEnd.Value.Date

        Dim newData2 As DataTable = ShowDataInSearch(dateFilter2, departmentFilter2, statusFilter2, searchKeyword2, startDate2, endDate2)
        DataGridView2.DataSource = newData2
    End Sub

    Private Sub DateTimePickerSearchPanelEnd_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerSearchPanelEnd.ValueChanged
        If ComboBoxDepartmentinSearchPanel.SelectedItem IsNot Nothing Then
            Dim dateFilter2 As String = ComboBoxDateinSearchPanel.SelectedItem.ToString()
            Dim departmentFilter2 As String = ComboBoxDepartmentinSearchPanel.SelectedItem.ToString()
            Dim statusFilter2 As String = ComboBoxStatusinSearchPanel.SelectedItem.ToString()
            Dim searchKeyword2 As String = SearchTextBoxinSearchPanel.Text
            Dim startDate2 As Date = DateTimePickerSearchPanelStart.Value.Date
            Dim endDate2 As Date = DateTimePickerSearchPanelEnd.Value.Date

            Dim newData2 As DataTable = ShowDataInSearch(dateFilter2, departmentFilter2, statusFilter2, searchKeyword2, startDate2, endDate2)
            DataGridView2.DataSource = newData2
        End If
    End Sub

    Private Sub SearchTextBox_TextChanged(sender As Object, e As EventArgs) Handles SearchTextBoxinSearchPanel.TextChanged
        If ComboBoxDepartmentinSearchPanel.SelectedItem IsNot Nothing Then
            Dim dateFilter2 As String = ComboBoxDateinSearchPanel.SelectedItem.ToString()
            Dim departmentFilter2 As String = ComboBoxDepartmentinSearchPanel.SelectedItem.ToString()
            Dim statusFilter2 As String = ComboBoxStatusinSearchPanel.SelectedItem.ToString()
            Dim searchKeyword2 As String = SearchTextBoxinSearchPanel.Text
            Dim startDate2 As Date = DateTimePickerSearchPanelStart.Value.Date
            Dim endDate2 As Date = DateTimePickerSearchPanelEnd.Value.Date

            Dim newData2 As DataTable = ShowDataInSearch(dateFilter2, departmentFilter2, statusFilter2, searchKeyword2, startDate2, endDate2)
            DataGridView2.DataSource = newData2
        End If
    End Sub
    Private Sub ComboBoxDateinSearchPanel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxDateinSearchPanel.SelectedIndexChanged
        If ComboBoxDepartmentinSearchPanel.SelectedItem IsNot Nothing Then
            Dim dateFilter2 As String = ComboBoxDateinSearchPanel.SelectedItem.ToString()
            Dim departmentFilter2 As String = ComboBoxDepartmentinSearchPanel.SelectedItem.ToString()
            Dim statusFilter2 As String = ComboBoxStatusinSearchPanel.SelectedItem.ToString()
            Dim searchKeyword2 As String = SearchTextBoxinSearchPanel.Text
            Dim startDate2 As Date = DateTimePickerSearchPanelStart.Value.Date
            Dim endDate2 As Date = DateTimePickerSearchPanelEnd.Value.Date

            Dim newData2 As DataTable = ShowDataInSearch(dateFilter2, departmentFilter2, statusFilter2, searchKeyword2, startDate2, endDate2)
            DataGridView2.DataSource = newData2
        End If
    End Sub

    Private Sub ComboBoxDepartmentinSearchPanel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxDepartmentinSearchPanel.SelectedIndexChanged
        If ComboBoxDepartmentinSearchPanel.SelectedItem IsNot Nothing Then
            Dim dateFilter2 As String = ComboBoxDateinSearchPanel.SelectedItem.ToString()
            Dim departmentFilter2 As String = ComboBoxDepartmentinSearchPanel.SelectedItem.ToString()
            Dim statusFilter2 As String = ComboBoxStatusinSearchPanel.SelectedItem.ToString()
            Dim searchKeyword2 As String = SearchTextBoxinSearchPanel.Text
            Dim startDate2 As Date = DateTimePickerSearchPanelStart.Value.Date
            Dim endDate2 As Date = DateTimePickerSearchPanelEnd.Value.Date

            Dim newData2 As DataTable = ShowDataInSearch(dateFilter2, departmentFilter2, statusFilter2, searchKeyword2, startDate2, endDate2)
            DataGridView2.DataSource = newData2
        End If
    End Sub

    Private Sub ComboBoxStatusinSearchPanel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxStatusinSearchPanel.SelectedIndexChanged
        If ComboBoxDepartmentinSearchPanel.SelectedItem IsNot Nothing Then
            Dim dateFilter2 As String = ComboBoxDateinSearchPanel.SelectedItem.ToString()
            Dim departmentFilter2 As String = ComboBoxDepartmentinSearchPanel.SelectedItem.ToString()
            Dim statusFilter2 As String = ComboBoxStatusinSearchPanel.SelectedItem.ToString()
            Dim searchKeyword2 As String = SearchTextBoxinSearchPanel.Text
            Dim startDate2 As Date = DateTimePickerSearchPanelStart.Value.Date
            Dim endDate2 As Date = DateTimePickerSearchPanelEnd.Value.Date

            Dim newData2 As DataTable = ShowDataInSearch(dateFilter2, departmentFilter2, statusFilter2, searchKeyword2, startDate2, endDate2)
            DataGridView2.DataSource = newData2
        End If
    End Sub

    Private Sub ButtonClearinSearchPanel_Click(sender As Object, e As EventArgs) Handles ButtonClearinSearchPanel.Click
        ComboBoxDateinSearchPanel.SelectedIndex = 0
        ComboBoxDepartmentinSearchPanel.SelectedIndex = 0
        ComboBoxStatusinSearchPanel.SelectedIndex = 0
        SearchTextBoxinSearchPanel.Text = ""
        DateTimePickerSearchPanelStart.Value = DateTime.Today
        DateTimePickerSearchPanelEnd.Value = DateTime.Today
    End Sub


    '*************************************************
    '
    '
    '   All the Component of Connection Panel
    '
    '
    '*************************************************

    'Scan the Port
    Private Sub ButtonScanPort_Click(sender As Object, e As EventArgs) Handles ButtonScanPort.Click
        Try
            ComboBoxPort.Items.Clear()
            Dim myPortNames As String() = SerialPortStream.GetPortNames()
            ComboBoxPort.Items.AddRange(myPortNames)

            If ComboBoxPort.Items.Count > 0 Then
                ComboBoxPort.SelectedIndex = 0
            Else
                MessageBox.Show("No COM ports detected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ComboBoxPort.Text = ""
            End If

            ComboBoxPort.DroppedDown = True
        Catch ex As UnauthorizedAccessException
            MessageBox.Show("Access to COM ports is denied. Make sure you have the necessary permissions.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("An error occurred while scanning COM ports: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Connect the Arduino to Application
    Private Sub ButtonConnect_Click(sender As Object, e As EventArgs) Handles ButtonConnect.Click
        If ButtonConnect.Text = "Connect" Then
            If ComboBoxBaudRate.SelectedItem IsNot Nothing AndAlso ComboBoxPort.SelectedItem IsNot Nothing Then
                serialPort = New SerialPortStream() ' Create a new instance of SerialPortStream
                serialPort.BaudRate = CInt(ComboBoxBaudRate.SelectedItem)
                serialPort.PortName = ComboBoxPort.SelectedItem.ToString()

                Try
                    serialPort.Open()
                    TimerSerialIn.Start()
                    ButtonConnect.Text = "Disconnect"
                    PictureBoxStatusConnect.Image = My.Resources.Connected
                    LabelConnectionStatus.Text = "Connection Status: Connected"
                Catch ex As UnauthorizedAccessException
                    MsgBox("Failed to connect to the serial port. Access denied.", MsgBoxStyle.Critical, "Error Message")
                    PictureBoxStatusConnect.Image = My.Resources.Disconnect
                    LabelConnectionStatus.Text = "Connection Status: Disconnected"
                Catch ex As IOException
                    MsgBox("Failed to connect to the serial port. The port is already in use.", MsgBoxStyle.Critical, "Error Message")
                    PictureBoxStatusConnect.Image = My.Resources.Disconnect
                    LabelConnectionStatus.Text = "Connection Status: Disconnected"
                Catch ex As ArgumentException
                    MsgBox("Failed to connect to the serial port. Invalid port name or settings.", MsgBoxStyle.Critical, "Error Message")
                    PictureBoxStatusConnect.Image = My.Resources.Disconnect
                    LabelConnectionStatus.Text = "Connection Status: Disconnected"
                Catch ex As Exception
                    MsgBox("Failed to connect to the serial port." & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Message")
                    PictureBoxStatusConnect.Image = My.Resources.Disconnect
                    LabelConnectionStatus.Text = "Connection Status: Disconnected"
                End Try
            Else
                MsgBox("Please select a baud rate and a port.", MsgBoxStyle.Information, "Error Message")
            End If
        ElseIf ButtonConnect.Text = "Disconnect" Then
            PictureBoxStatusConnect.Image = My.Resources.Disconnect
            ButtonConnect.Text = "Connect"
            LabelConnectionStatus.Text = "Connection Status: Disconnected"
            TimerSerialIn.Stop()

            Try
                serialPort.Close()
            Catch ex As Exception
                MsgBox("Failed to disconnect from the serial port." & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Message")
            End Try
        End If
    End Sub

    '*************************************************
    '
    '
    '   All the Component of UserDetails Panel
    '
    '
    '*************************************************


    'Show data in UserDetails Panel'
    'You can check the UserDetails via RFID Card and display the user information'
    Private Sub ShowDataUser()
        Dim errorMessage As String
        Try
            Using connection As New MySqlConnection(connectionString)
                Try
                    connection.Open()
                Catch ex As Exception
                    errorMessage = "Connection failed!" & vbCrLf & "Please check that the server is ready!"
                    MessageBox.Show(errorMessage, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try

                Try
                    Dim query As String = "SELECT * FROM " & Table_Name & " sl INNER JOIN " & Table_Name3 & " im ON sl.Student_ID = im.Student_ID WHERE sl.RFID_UID = @RFID_UID"

                    Using command As New MySqlCommand(query, connection)
                        command.Parameters.AddWithValue("@RFID_UID", LabelID.Text.Replace(vbCrLf, ""))

                        Dim adapter As New MySqlDataAdapter(command)
                        Dim dataTable As New DataTable()
                        Dim dataCount As Integer = adapter.Fill(dataTable)

                        If dataCount > 0 Then
                            Dim imgArray() As Byte = DirectCast(dataTable.Rows(0).Item("Images"), Byte())
                            Using imgStream As New System.IO.MemoryStream(imgArray)
                                PictureBoxUserImage.Image = Image.FromStream(imgStream)
                            End Using

                            LabelID.Text = dataTable.Rows(0).Item("RFID_UID").ToString()
                            LabelStudentID.Text = dataTable.Rows(0).Item("Student_ID").ToString()
                            LabelFirstname.Text = dataTable.Rows(0).Item("Firstname").ToString()
                            LabelLastname.Text = dataTable.Rows(0).Item("Lastname").ToString()
                            LabelMiddlename.Text = dataTable.Rows(0).Item("Middlename").ToString()
                            LabelAge.Text = dataTable.Rows(0).Item("Age").ToString()
                            LabelDepartment.Text = dataTable.Rows(0).Item("Department").ToString()
                            LabelCourse.Text = dataTable.Rows(0).Item("Course").ToString()
                            LabelYear.Text = dataTable.Rows(0).Item("Year").ToString()
                            LabelParent_Number.Text = dataTable.Rows(0).Item("Parent_Number").ToString()
                        Else
                            errorMessage = "RFID_UID not found!" & vbCr & "Please register your RFID_UID."
                            MessageBox.Show(errorMessage, "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End Using
                Catch ex As Exception
                    errorMessage = "Failed to load data from the database!" & vbCr & ex.Message
                    MessageBox.Show(errorMessage, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try
            End Using
        Catch ex As Exception
            errorMessage = "An error occurred!" & vbCr & ex.Message
            MessageBox.Show(errorMessage, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
    End Sub

    'To Read the RFID card'
    'Display the connected status'
    Private Sub TimerSerialIn_Tick(sender As Object, e As EventArgs) Handles TimerSerialIn.Tick
        Try
            StrSerialIn = serialPort.ReadExisting
            LabelConnectionStatus.Text = "Connection Status: Connected"
            If StrSerialIn <> "" Then
                If GetID = True Then
                    LabelGetID.Text = StrSerialIn
                    serialPort.Write("A")
                    GetID = False
                    If LabelGetID.Text <> "____________" Then
                        PanelReadingTagProcess.Visible = False
                        IDCheck()
                    End If
                End If
                If ViewUserData = True Then
                    ViewData()
                End If
            End If
        Catch ex As TimeoutException
            TimerSerialIn.Stop()
            serialPort.Close()
            LabelConnectionStatus.Text = "Connection Status: Disconnected"
            PictureBoxStatusConnect.Image = My.Resources.Disconnect
            MsgBox("Failed to read from the serial port. Timeout occurred.", MsgBoxStyle.Critical, "Error Message")
        Catch ex As IOException
            TimerSerialIn.Stop()
            serialPort.Close()
            LabelConnectionStatus.Text = "Connection Status: Disconnected"
            PictureBoxStatusConnect.Image = My.Resources.Disconnect
            MsgBox("Failed to read from the serial port. The port is closed or disconnected.", MsgBoxStyle.Critical, "Error Message")
        Catch ex As Exception
            TimerSerialIn.Stop()
            serialPort.Close()
            LabelConnectionStatus.Text = "Connection Status: Disconnected"
            PictureBoxStatusConnect.Image = My.Resources.Disconnect
            MsgBox("Failed to read from the serial port." & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Message")
        End Try

        If PictureBoxStatusConnect.Visible = True Then
            PictureBoxStatusConnect.Visible = False
        ElseIf PictureBoxStatusConnect.Visible = False Then
            PictureBoxStatusConnect.Visible = True
        End If
    End Sub

    'Check the RFID card if already exist in database'
    Private Sub IDCheck()
        Using connection As New MySqlConnection(connectionString)
            Try
                connection.Open()
            Catch ex As Exception
                MessageBox.Show("Connection failed!" & vbCrLf & "Please check that the server is ready.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            Try
                MySQLCMD.Connection = connection
                MySQLCMD.CommandType = CommandType.Text
                MySQLCMD.CommandText = "SELECT * FROM " & Table_Name & " WHERE RFID_UID ='" & LabelGetID.Text.Replace(vbLf, "").Replace(vbCr, "") & "'"
                MySQLDA = New MySqlDataAdapter(MySQLCMD.CommandText, connection)
                DT = New DataTable
                Data = MySQLDA.Fill(DT)
                If Data > 0 Then
                    If MsgBox("RFID_UID registered!" & vbCrLf & "Do you want to edit the data?", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "Confirmation") = MsgBoxResult.Cancel Then
                        DT = Nothing
                        connection.Close()
                        ButtonScanID.Enabled = True
                        GetID = False
                        LabelGetID.Text = "________"
                        Return
                    Else
                        TextBoxStudent_ID.Text = DT.Rows(0).Item("Student_ID")
                        TextBoxFirstname.Text = DT.Rows(0).Item("Firstname")
                        TextBoxMiddlename.Text = DT.Rows(0).Item("Middlename")
                        TextBoxLastname.Text = DT.Rows(0).Item("Lastname")
                        TextBoxAge.Text = DT.Rows(0).Item("Age")
                        ComboBoxDepartment.SelectedIndex = ComboBoxDepartment.FindStringExact(DT.Rows(0).Item("Department").ToString())
                        ComboBoxCourse.SelectedIndex = ComboBoxCourse.FindStringExact(DT.Rows(0).Item("Course").ToString())
                        ComboBoxYear.SelectedIndex = ComboBoxYear.FindStringExact(DT.Rows(0).Item("Year").ToString())

                        TextBoxParent_Number.Text = DT.Rows(0).Item("Parent_Number")
                        TextBoxStudentID.Text = DT.Rows(0).Item("ID")
                        StatusInput = "Update"

                        ' Retrieve the student image and ID from the separate table
                        Dim imageQuery As String = "SELECT ID, images FROM " & Table_Name3 & " WHERE Student_ID = '" & DT.Rows(0).Item("Student_ID") & "'"
                        Dim imageCMD As New MySqlCommand(imageQuery, connection)
                        Dim reader As MySqlDataReader = imageCMD.ExecuteReader()

                        ' Check if there is a result
                        If reader.Read() Then
                            ' Get the image data and ID
                            Dim imageData As Byte() = DirectCast(reader("images"), Byte())
                            Dim studentID As Integer = Convert.ToInt32(reader("ID"))

                            ' Display the student image in a PictureBox control
                            If imageData IsNot Nothing Then
                                Using ms As New MemoryStream(imageData)
                                    PictureBoxImageInput.Image = Image.FromStream(ms)
                                End Using
                            Else
                                ' No image found for the student
                                PictureBoxImageInput.Image = Nothing
                            End If

                            ' Display the ID in a TextBox control
                            TextBoxID.Text = studentID.ToString()
                        Else
                            ' No record found
                            PictureBoxImageInput.Image = Nothing
                            TextBoxID.Text = ""
                        End If

                        reader.Close()
                    End If
                Else
                    MessageBox.Show("RFID_UID not found!" & vbCrLf & "Please register your RFID_UID.", "Information Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As MySqlException
                MessageBox.Show("Database error!" & vbCrLf & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch ex As Exception
                MessageBox.Show("An error occurred!" & vbCrLf & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            DT = Nothing
            connection.Close()

            ButtonScanID.Enabled = True
            GetID = False
        End Using
    End Sub

    'To display the RFID_UID in UserDetails'
    Private Sub ViewData()
        Try
            LabelID.Text = StrSerialIn
            serialPort.Write("A")
            If LabelID.Text = "" Then
                ViewData()
            Else
                ShowDataUser()
            End If
        Catch ex As TimeoutException
            MsgBox("Failed to read from the serial port. Timeout occurred.", MsgBoxStyle.Critical, "Error Message")
        Catch ex As IOException
            MsgBox("Failed to read from the serial port. The port is closed or disconnected.", MsgBoxStyle.Critical, "Error Message")
        Catch ex As Exception
            MsgBox("Failed to read from the serial port." & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Message")
        End Try
    End Sub

    'Clear the UserDetails
    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        LabelID.Text = "______________"
        LabelStudentID.Text = "Waiting..."
        LabelFirstname.Text = "Waiting..."
        LabelMiddlename.Text = "Waiting..."
        LabelLastname.Text = "Waiting..."
        LabelAge.Text = "Waiting..."
        LabelDepartment.Text = "Waiting..."
        LabelCourse.Text = "Waiting..."
        LabelYear.Text = "Waiting..."
        LabelParent_Number.Text = "Waiting..."
        PictureBoxUserImage.Image = My.Resources.icons8_image_96

    End Sub


    '*************************************************
    '
    '
    '   All the Component of Registration Panel
    '
    '
    '*************************************************


    'Display Student Data in Datagridview3
    Private Function ShowDataInRegistration(yearLevel As String, searchKeyword As String) As DataTable
        Dim dataTable As New DataTable()
        Try
            Dim query As String = "SELECT `ID`, `Student_ID`, `Firstname`, `Middlename`, `Lastname`, `Age`, `Department`, `Course`, `Year`, `Parent_Number`, `RFID_UID` FROM " & Table_Name

            If yearLevel <> "All" Then
                query &= " WHERE `Department` = @YearLevel"
            End If

            If searchKeyword <> "" Then
                Dim searchColumn As String
                If CheckBoxByID.Checked Then
                    searchColumn = "Student_ID"
                ElseIf CheckBoxByName.Checked Then
                    searchColumn = "`Firstname` OR `Lastname`"
                Else
                    searchColumn = "`Firstname` OR `Lastname` OR `Student_ID`"
                End If

                If yearLevel = "All" Then
                    query &= " WHERE"
                Else
                    query &= " AND"
                End If

                query &= " (" & searchColumn & " LIKE @Keyword)"
            End If

            query &= " ORDER BY ID"

            Using connection As New MySqlConnection(connectionString)
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    If yearLevel <> "All" Then
                        command.Parameters.AddWithValue("@YearLevel", yearLevel)
                    End If

                    If searchKeyword <> "" Then
                        command.Parameters.AddWithValue("@Keyword", "%" & searchKeyword & "%")
                    End If

                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End Using
            End Using

            DataGridView3.DefaultCellStyle.ForeColor = Color.Black
            DataGridView3.ClearSelection()
            DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView3.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
        Catch ex As MySqlException
            MessageBox.Show("A MySQL error occurred in the ShowDataInRegistration function: " & ex.Message, "MySQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("An error occurred in the ShowDataInRegistration function: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dataTable
    End Function

    'Display the Student Image to Registration Panel
    Private Sub DataGridView3_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView3.SelectionChanged
        Try
            If DataGridView3.SelectedRows.Count > 0 Then
                Dim studentID As String = DataGridView3.SelectedRows(0).Cells("Student_ID").Value.ToString()
                Dim query As String = "SELECT `images` FROM " & Table_Name3 & " WHERE `Student_ID` = '" & studentID & "'"

                Using connection As New MySqlConnection(connectionString)
                    connection.Open()

                    Using command As New MySqlCommand(query, connection)
                        Using reader As MySqlDataReader = command.ExecuteReader()
                            If reader.Read() Then
                                Dim imageData As Byte() = DirectCast(reader("images"), Byte())
                                Using ms As New MemoryStream(imageData)
                                    PictureBoxImagePreview.Image = Image.FromStream(ms)
                                End Using
                            End If
                        End Using
                    End Using
                End Using
            End If
        Catch ex As MySqlException
            MessageBox.Show("A MySQL error occurred in the DataGridView3_SelectionChanged event handler: " & ex.Message, "MySQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("An error occurred in the DataGridView3_SelectionChanged event handler: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Button for scanning the RFID CARD
    Private Sub ButtonScanID_Click(sender As Object, e As EventArgs) Handles ButtonScanID.Click
        Try
            If TimerSerialIn.Enabled = True Then
                PanelReadingTagProcess.Visible = True
                GetID = True
                ButtonScanID.Enabled = False
            Else
                MsgBox("Failed to open the scanner of RFID Card !!!" & vbCrLf & "Click the Connection menu then click the Connect button.", MsgBoxStyle.Critical, "Error Message")
            End If
        Catch ex As Exception
            MsgBox("An error occurred while scanning the RFID card." & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Message")
        End Try
    End Sub

    'CLear Registration Form
    Private Sub ButtonClear2_Click(sender As Object, e As EventArgs) Handles ButtonClear2.Click
        TextBoxStudent_ID.Text = ""
        TextBoxFirstname.Text = ""
        TextBoxMiddlename.Text = ""
        TextBoxLastname.Text = ""
        LabelGetID.Text = "____________"
        TextBoxAge.Text = ""
        ComboBoxDepartment.Text = ""
        ComboBoxCourse.Text = ""
        ComboBoxYear.Text = ""
        TextBoxParent_Number.Text = "+63"
        PictureBoxImageInput.Image = My.Resources.icons8_upload_image_96
    End Sub

    'Scanning the RFID Card
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PanelReadingTagProcess.Visible = False
        ButtonScanID.Enabled = True
    End Sub

    'Input image
    Private Sub PictureBoxImageInput_Click(sender As Object, e As EventArgs) Handles PictureBoxImageInput.Click
        Try
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "JPEG (*.jpeg;*.jpg)|*.jpeg;*.jpg"

            If OpenFileDialog1.ShowDialog(Me) = DialogResult.OK Then
                IMG_FileNameInput = OpenFileDialog1.FileName
                PictureBoxImageInput.ImageLocation = IMG_FileNameInput
            End If
        Catch ex As Exception
            MsgBox("An error occurred while selecting the image!" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Message")
        End Try
    End Sub


    'Save data of Student
    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
        Try
            Dim mstream As New System.IO.MemoryStream()
            Dim arrImage() As Byte

            If TextBoxStudent_ID.Text = "" Then
                MessageBox.Show("Student ID cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If TextBoxFirstname.Text = "" Then
                MessageBox.Show("Firstname cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If TextBoxMiddlename.Text = "" Then
                MessageBox.Show("Middlename cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If TextBoxLastname.Text = "" Then
                MessageBox.Show("Lastname cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If TextBoxAge.Text = "" Then
                MessageBox.Show("Age cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If ComboBoxDepartment.Text = "" Then
                MessageBox.Show("Department cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If ComboBoxCourse.Text = "" Then
                MessageBox.Show("Course cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If ComboBoxYear.Text = "" Then
                MessageBox.Show("Year cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If TextBoxParent_Number.Text = "" Then
                MessageBox.Show("Parent Phone Number cannot be empty!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If StatusInput = "Save" Then
                ' Check if the student ID already exists in the database
                Dim isStudentRegistered As Boolean = CheckIfStudentRegistered(TextBoxStudent_ID.Text)

                If isStudentRegistered Then
                    MessageBox.Show("Student ID already exists in the database!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                If IMG_FileNameInput <> "" Then
                    PictureBoxImageInput.Image.Save(mstream, System.Drawing.Imaging.ImageFormat.Jpeg)
                    arrImage = mstream.GetBuffer()
                Else
                    MessageBox.Show("The image has not been selected!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Try
                    Connection.Open()
                    MySQLCMD = New MySqlCommand
                    With MySQLCMD
                        .CommandText = "INSERT INTO " & Table_Name & " (`Student_ID`, `Firstname`, `Middlename`, `Lastname`, `Age`, `Department`, `Course`, `Year`, `Parent_Number`, `RFID_UID`) VALUES (@Student_ID,@Firstname,@Middlename,@Lastname,@Age,@Department,@Course,@Year,@Parent_Number,@RFID_UID)"
                        .Connection = Connection
                        .Parameters.AddWithValue("@RFID_UID", LabelGetID.Text.Replace(vbCrLf, ""))
                        .Parameters.AddWithValue("@Student_ID", TextBoxStudent_ID.Text)
                        .Parameters.AddWithValue("@Firstname", TextBoxFirstname.Text)
                        .Parameters.AddWithValue("@Middlename", TextBoxMiddlename.Text)
                        .Parameters.AddWithValue("@Lastname", TextBoxLastname.Text)
                        .Parameters.AddWithValue("@Age", TextBoxAge.Text)
                        .Parameters.AddWithValue("@Department", ComboBoxDepartment.Text)
                        .Parameters.AddWithValue("@Course", ComboBoxCourse.Text)
                        .Parameters.AddWithValue("@Year", ComboBoxYear.Text)
                        .Parameters.AddWithValue("@Parent_Number", TextBoxParent_Number.Text)
                        .ExecuteNonQuery()
                    End With

                    MySQLCMD = New MySqlCommand
                    With MySQLCMD
                        .CommandText = "INSERT INTO " & Table_Name3 & " (`Student_ID`, `images`) VALUES (@Student_ID, @images)"
                        .Connection = Connection
                        .Parameters.AddWithValue("@Student_ID", TextBoxStudent_ID.Text)
                        .Parameters.AddWithValue("@images", arrImage)
                        .ExecuteNonQuery()
                    End With

                    MsgBox("Student Registered Successfully", MsgBoxStyle.Information, "Information")
                    IMG_FileNameInput = ""
                Catch ex As Exception
                    MsgBox("Data failed to save!" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Message")
                    Connection.Close()
                    Return
                End Try
                Connection.Close()
            End If

            PictureBoxImagePreview.Image = Nothing
            Dim yearLevel As String = ComboBoxYearlevelinRegistration.SelectedItem.ToString()
            Dim newData As DataTable = ShowDataInRegistration(yearLevel, "")
            DataGridView3.DataSource = newData
        Catch ex As Exception
            MsgBox("An error occurred while saving the data!" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Message")
        End Try
    End Sub


    'Check if the student already registered
    Private Function CheckIfStudentRegistered(studentID As String) As Boolean
        ' Modify the SQL query to check if the student ID exists in the database
        Dim query As String = "SELECT COUNT(*) FROM " & Table_Name & " WHERE `Student_ID` = @Student_ID"
        Try
            Connection.Open()
            MySQLCMD = New MySqlCommand(query, Connection)
            MySQLCMD.Parameters.AddWithValue("@Student_ID", studentID)
            Dim count As Integer = CInt(MySQLCMD.ExecuteScalar())
            Return count > 0
        Catch ex As Exception
            MessageBox.Show("Error checking student registration: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            Connection.Close()
        End Try
    End Function



    'Delete Student in Registration Panel
    Private Sub ButtonDelete_Click(sender As Object, e As EventArgs) Handles ButtonDelete.Click
        Dim selectedData As String = DataGridView3.SelectedCells(0).Value.ToString()

        ' Display a confirmation prompt
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this data?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' If the user clicks Yes, proceed with deletion
        If result = DialogResult.Yes Then
            Try
                Using connection As New MySqlConnection(connectionString)
                    Dim query As String = "DELETE FROM student_list WHERE ID = @selectedData"
                    Dim command As New MySqlCommand(query, connection)
                    command.Parameters.AddWithValue("@selectedData", selectedData)

                    connection.Open()
                    command.ExecuteNonQuery()
                    MessageBox.Show("Data deleted successfully.")
                End Using
            Catch ex As Exception
                MessageBox.Show("Error deleting data: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                ' Retrieve new data for DataGridView3 after the update
                Dim yearLevel As String = ComboBoxYearlevelinRegistration.SelectedItem.ToString()
                Dim searchKeyword As String = TextBoxSearch.Text

                Dim newData As DataTable = ShowDataInRegistration(yearLevel, searchKeyword)
                DataGridView3.DataSource = newData
            End Try
        End If
    End Sub


    'Select Student_ID in Registration Panel
    Private Sub DataGridView3_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown
        Try
            If AllCellsSelected(DataGridView1) = False Then
                If e.Button = MouseButtons.Left Then
                    DataGridView1.CurrentCell = DataGridView1(e.ColumnIndex, e.RowIndex)
                    Dim i As Integer
                    With DataGridView1
                        If e.RowIndex >= 0 Then
                            i = .CurrentRow.Index
                            LoadImagesStr = True
                            IDRam = .Rows(i).Cells("Student_ID").Value.ToString
                            ' Retrieve new data for DataGridView3 after the update
                            Dim yearLevel As String = ComboBoxYearlevelinRegistration.SelectedItem.ToString()
                            Dim searchKeyword As String = TextBoxSearch.Text

                            Dim newData As DataTable = ShowDataInRegistration(yearLevel, searchKeyword)
                            DataGridView3.DataSource = newData
                        End If
                    End With
                End If
            End If
        Catch ex As Exception
            ' Handle the exception or display an error message
            MessageBox.Show("Error: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Perform any additional error handling as needed
        End Try
    End Sub


    'All select in datagridview3
    Private Function AllCellsSelected(dgv As DataGridView) As Boolean
        AllCellsSelected = (DataGridView3.SelectedCells.Count = (DataGridView3.RowCount * DataGridView3.Columns.GetColumnCount(DataGridViewElementStates.Visible)))
    End Function

    'Delete in datagridview3 using ContextMenuStrip
    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView3.RowCount = 0 Then
            MsgBox("Cannot delete, table data is empty", MsgBoxStyle.Critical, "Error Message")
            Return
        End If

        If DataGridView3.SelectedRows.Count = 0 Then
            MsgBox("Cannot delete, select the table data to be deleted", MsgBoxStyle.Critical, "Error Message")
            Return
        End If

        If MsgBox("Delete record?", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "Confirmation") = MsgBoxResult.Cancel Then
            Return
        End If

        Try
            Connection.Open()
        Catch ex As Exception
            MessageBox.Show("Connection failed !!!" & vbCrLf & "Please check that the server is ready !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        Try
            For Each row As DataGridViewRow In DataGridView3.SelectedRows
                If row.Selected = True Then
                    Dim recordID As String = row.Cells("ID").Value.ToString()
                    Dim studentID As String = row.Cells("Student_ID").Value.ToString()

                    ' Delete record from Table_Name
                    Dim deleteQuery As String = "DELETE FROM " & Table_Name & " WHERE ID = @ID"
                    Dim deleteCommand As New MySqlCommand(deleteQuery, Connection)
                    deleteCommand.Parameters.AddWithValue("@ID", recordID)
                    deleteCommand.ExecuteNonQuery()

                    ' Delete image from Table_Name3
                    Dim deleteImageQuery As String = "DELETE FROM " & Table_Name3 & " WHERE Student_ID = @Student_ID"
                    Dim deleteImageCommand As New MySqlCommand(deleteImageQuery, Connection)
                    deleteImageCommand.Parameters.AddWithValue("@Student_ID", studentID)
                    deleteImageCommand.ExecuteNonQuery()
                End If
            Next

            ' Display successful deletion message
            MessageBox.Show("Deletion successful.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Failed to delete: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Connection.Close()

            ' Retrieve new data for DataGridView3 after the update
            Dim yearLevel As String = ComboBoxYearlevelinRegistration.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearch.Text

            Dim newData As DataTable = ShowDataInRegistration(yearLevel, searchKeyword)
            DataGridView3.DataSource = newData
        End Try
    End Sub

    'Select all in datagridview3
    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        DataGridView3.SelectAll()
    End Sub

    'Clear selection in datagridvew3
    Private Sub ClearSelectionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearSelectionToolStripMenuItem.Click
        DataGridView3.ClearSelection()
    End Sub

    'Refresh in datagridview3
    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        ' Retrieve new data for DataGridView3 after the update
        Dim yearLevel As String = ComboBoxYearlevelinRegistration.SelectedItem.ToString()
        Dim searchKeyword As String = TextBoxSearch.Text

        Dim newData As DataTable = ShowDataInRegistration(yearLevel, searchKeyword)
        DataGridView3.DataSource = newData
    End Sub

    'Display the Student data
    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        ' Get the selected row in the DataGridView
        Dim selectedRow As DataGridViewRow = DataGridView3.CurrentRow

        ' Check if a row is selected
        If selectedRow IsNot Nothing Then
            Try
                ' Get the value from the desired column in the selected row
                Dim ID As String = selectedRow.Cells("ID").Value.ToString()
                Dim Student_ID As String = selectedRow.Cells("Student_ID").Value.ToString()
                Dim Firstname As String = selectedRow.Cells("Firstname").Value.ToString()
                Dim Middlename As String = selectedRow.Cells("Middlename").Value.ToString()
                Dim Lastname As String = selectedRow.Cells("Lastname").Value.ToString()
                Dim Age As String = selectedRow.Cells("Age").Value.ToString()
                Dim Department As String = selectedRow.Cells("Department").Value.ToString()
                Dim Course As String = selectedRow.Cells("Course").Value.ToString()
                Dim Year As String = selectedRow.Cells("Year").Value.ToString()
                Dim Parent_Number As String = selectedRow.Cells("Parent_Number").Value.ToString()
                Dim RFID_UID As String = selectedRow.Cells("RFID_UID").Value.ToString()

                Dim query As String = "SELECT * FROM " & Table_Name3 & " WHERE Student_ID = @StudentID"

                Using connection As New MySqlConnection(connectionString)
                    connection.Open()

                    Using command As New MySqlCommand(query, connection)
                        command.Parameters.AddWithValue("@StudentID", Student_ID)

                        Using reader As MySqlDataReader = command.ExecuteReader()
                            ' Check if the reader has rows
                            If reader.HasRows Then
                                ' Read through each row in the result set
                                While reader.Read()
                                    ' Access the columns in the current row by name or index
                                    Dim columnValue As Integer = reader.GetInt32(reader.GetOrdinal("id")) ' Replace "id" with the appropriate column name
                                    ' Retrieve the image data as a byte array
                                    Dim imageData As Byte() = DirectCast(reader("images"), Byte())

                                    ' Create a MemoryStream to store the image data
                                    Using ms As New MemoryStream(imageData)
                                        ' Convert the byte array to an Image object
                                        Dim image As Image = Image.FromStream(ms)

                                        TextBoxID.Text = columnValue
                                        ' Display the image in the PictureBox control
                                        PictureBoxImageInput.Image = image
                                    End Using
                                End While
                            Else
                                MessageBox.Show("No data found.")
                            End If
                        End Using
                    End Using
                End Using

                ' Display the student data in the TextBox
                TextBoxStudent_ID.Text = Student_ID
                TextBoxFirstname.Text = Firstname
                TextBoxMiddlename.Text = Middlename
                TextBoxLastname.Text = Lastname
                TextBoxStudentID.Text = ID
                LabelGetID.Text = RFID_UID
                TextBoxAge.Text = Age
                ComboBoxDepartment.Text = Department
                ComboBoxCourse.Text = Course
                ComboBoxYear.Text = Year
                TextBoxParent_Number.Text = Parent_Number
            Catch ex As Exception
                MessageBox.Show("Error retrieving data: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub


    Private Sub ComboBoxDepartment_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxDepartment.SelectedIndexChanged
        Select Case ComboBoxDepartment.SelectedItem.ToString()
            Case "College"
                ComboBoxCourse.Items.Clear()
                ComboBoxYear.Items.Clear()
                ComboBoxCourse.Items.Add("BSA")
                ComboBoxCourse.Items.Add("BSMA")
                ComboBoxCourse.Items.Add("BSAIS")
                ComboBoxCourse.Items.Add("BSENT")
                ComboBoxCourse.Items.Add("BSIS")
                ComboBoxCourse.Items.Add("BSDC")
                ComboBoxCourse.Items.Add("BSIT")
                ComboBoxCourse.Items.Add("BSHM")
                ComboBoxYear.Items.Add("1st Year")
                ComboBoxYear.Items.Add("2nd Year")
                ComboBoxYear.Items.Add("3rd Year")
                ComboBoxYear.Items.Add("4th Year")
            Case "Senior High"
                ComboBoxCourse.Items.Clear()
                ComboBoxYear.Items.Clear()
                ComboBoxCourse.Items.Add("ABM")
                ComboBoxCourse.Items.Add("HUMSS")
                ComboBoxCourse.Items.Add("STEM")
                ComboBoxYear.Items.Add("Grade 11")
                ComboBoxYear.Items.Add("Grade 12")
        End Select
    End Sub

    Private Sub ClearInputUpdateData()
        TextBoxStudent_ID.Text = ""
        TextBoxFirstname.Text = ""
        TextBoxMiddlename.Text = ""
        TextBoxLastname.Text = ""
        LabelGetID.Text = "____________"
        TextBoxAge.Text = ""
        ComboBoxDepartment.Text = ""
        ComboBoxCourse.Text = ""
        ComboBoxYear.Text = ""
        TextBoxParent_Number.Text = "+63"
        PictureBoxImageInput.Image = My.Resources.icons8_upload_image_96
    End Sub

    Private Sub CheckBoxByName_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxByName.CheckedChanged
        ComboBoxYearlevelinRegistration_SelectedIndexChanged(ComboBoxYearlevelinRegistration, EventArgs.Empty)
        If CheckBoxByName.Checked = True Then
            CheckBoxByID.Checked = False
        End If
    End Sub

    Private Sub CheckBoxByID_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxByID.CheckedChanged
        ComboBoxYearlevelinRegistration_SelectedIndexChanged(ComboBoxYearlevelinRegistration, EventArgs.Empty)
        If CheckBoxByID.Checked = True Then
            CheckBoxByName.Checked = False
        End If
    End Sub

    Private Sub TextBoxSearch_TextChanged(sender As Object, e As EventArgs) Handles TextBoxSearch.TextChanged
        ComboBoxYearlevelinRegistration_SelectedIndexChanged(ComboBoxYearlevelinRegistration, EventArgs.Empty)
    End Sub

    'Update Student
    Private Sub ButtonUpdate_Click(sender As Object, e As EventArgs) Handles ButtonUpdate.Click
        If TextBoxStudent_ID.Text = "" Then
            MessageBox.Show("Student ID cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If TextBoxFirstname.Text = "" Then
            MessageBox.Show("Firstname cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If TextBoxMiddlename.Text = "" Then
            MessageBox.Show("Middlename cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If TextBoxLastname.Text = "" Then
            MessageBox.Show("Lastname cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If TextBoxAge.Text = "" Then
            MessageBox.Show("Age cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If ComboBoxDepartment.Text = "" Then
            MessageBox.Show("Department cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If ComboBoxCourse.Text = "" Then
            MessageBox.Show("Course cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If ComboBoxYear.Text = "" Then
            MessageBox.Show("Year cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If TextBoxParent_Number.Text = "" Then
            MessageBox.Show("Parent Phone Number cannot be empty !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim imageUpdated As Boolean = (IMG_FileNameInput <> "")

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Dim query As String = "UPDATE " & Table_Name & " SET Student_ID=@Student_ID, Firstname=@Firstname, Middlename=@Middlename, Lastname=@Lastname, Age=@Age, Department=@Department, Course=@Course, Year=@Year, Parent_Number=@Parent_Number, RFID_UID=@RFID_UID WHERE ID=@ID"

                Using command As New MySqlCommand(query, connection)
                    command.Parameters.AddWithValue("@Student_ID", TextBoxStudent_ID.Text)
                    command.Parameters.AddWithValue("@Firstname", TextBoxFirstname.Text)
                    command.Parameters.AddWithValue("@Middlename", TextBoxMiddlename.Text)
                    command.Parameters.AddWithValue("@Lastname", TextBoxLastname.Text)
                    command.Parameters.AddWithValue("@Age", TextBoxAge.Text)
                    command.Parameters.AddWithValue("@Department", ComboBoxDepartment.Text)
                    command.Parameters.AddWithValue("@Course", ComboBoxCourse.Text)
                    command.Parameters.AddWithValue("@Year", ComboBoxYear.Text)
                    command.Parameters.AddWithValue("@Parent_Number", TextBoxParent_Number.Text)
                    command.Parameters.AddWithValue("@ID", TextBoxStudentID.Text)
                    command.Parameters.AddWithValue("@RFID_UID", LabelGetID.Text.Replace(vbCrLf, ""))
                    command.ExecuteNonQuery()

                    If imageUpdated Then
                        UpdateImage(command)
                    End If

                    ' Create a new MySqlCommand for the second update query
                    Dim command2 As New MySqlCommand()
                    command2.Connection = connection
                    command2.CommandText = "UPDATE " & Table_Name3 & " SET Student_ID=@Student_ID WHERE id=@ID_Image"
                    command2.Parameters.AddWithValue("@Student_ID", TextBoxStudent_ID.Text)
                    command2.Parameters.AddWithValue("@ID_Image", TextBoxID.Text)
                    command2.ExecuteNonQuery()

                    MsgBox("Data updated successfully", MsgBoxStyle.Information, "Information")
                    ClearInputUpdateData()
                End Using

                PictureBoxImagePreview.Image = Nothing
            End Using
        Catch ex As Exception
            MsgBox("Data failed to update!" & vbCr & ex.Message, MsgBoxStyle.Critical, "Error Message")
        End Try

        ' Retrieve new data for DataGridView3 after the update
        Dim yearLevel As String = ComboBoxYearlevelinRegistration.SelectedItem.ToString()
        Dim newData As DataTable = ShowDataInRegistration(yearLevel, "")
        DataGridView3.DataSource = newData
    End Sub


    'Update Image
    Private Sub UpdateImage(command As MySqlCommand)
        Using mstream As New System.IO.MemoryStream()
            Dim arrImage() As Byte

            If PictureBoxImageInput.Image IsNot Nothing Then
                Try
                    PictureBoxImageInput.Image.Save(mstream, System.Drawing.Imaging.ImageFormat.Jpeg)
                    arrImage = mstream.GetBuffer()

                    command.CommandText = "UPDATE " & Table_Name3 & " SET images=@images WHERE id=@ID_Image"
                    command.Parameters.AddWithValue("@images", arrImage)
                    command.Parameters.AddWithValue("@ID_Image", TextBoxID.Text)
                    command.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("Failed to update image: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Else
                MessageBox.Show("No image available.")
            End If
        End Using
    End Sub

    Private Sub ComboBoxYearlevelinRegistration_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxYearlevelinRegistration.SelectedIndexChanged
        If ComboBoxYearlevelinRegistration.SelectedItem IsNot Nothing Then
            Dim yearLevel As String = ComboBoxYearlevelinRegistration.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearch.Text

            Dim newData As DataTable = ShowDataInRegistration(yearLevel, searchKeyword)
            DataGridView3.DataSource = newData
        End If
    End Sub

    '*************************************************
    '
    '
    '   All the Component of Retrieve Panel
    '
    '
    '*************************************************


    Private Sub ButtonChoosefile_Click(sender As Object, e As EventArgs) Handles ButtonChoosefile.Click
        Try
            ' Register the encoding provider
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)

            ' Create OpenFileDialog
            Dim openFileDialog As New OpenFileDialog()

            ' Set dialog properties
            openFileDialog.Title = "Select an Archive File"
            openFileDialog.Filter = "Archive Files (*.zip, *.rar)|*.zip;*.rar|All Files (*.*)|*.*"

            ' Show the dialog and get the selected archive file
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                ' Get the selected archive file path
                Dim archiveFilePath As String = openFileDialog.FileName

                ' Extract the archive to a temporary directory
                Dim tempDir As String = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString())
                Directory.CreateDirectory(tempDir)

                Using archive As IArchive = ArchiveFactory.Open(archiveFilePath)
                    For Each entry As IArchiveEntry In archive.Entries
                        If Not entry.IsDirectory Then
                            ' Extract only Excel files
                            If Path.GetExtension(entry.Key).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) Then
                                entry.WriteToFile(Path.Combine(tempDir, entry.Key), New ExtractionOptions() With {
                            .ExtractFullPath = True,
                            .Overwrite = True
                        })
                            End If
                        End If
                    Next
                End Using

                ' Create OpenFileDialog for the extracted files
                Dim fileOpenDialog As New OpenFileDialog()
                fileOpenDialog.Title = "Select an Excel File"
                fileOpenDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
                fileOpenDialog.InitialDirectory = tempDir ' Set the initial directory to the extracted files directory

                ' Show the dialog and get the selected file
                If fileOpenDialog.ShowDialog() = DialogResult.OK Then
                    ' Get the selected file path
                    Dim excelFilePath As String = fileOpenDialog.FileName

                    ' Specify the encoding for the ExcelDataReader
                    Dim encoding As Encoding = Encoding.GetEncoding("utf-8")

                    ' Read Excel file using ExcelDataReader
                    Using stream As FileStream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read)
                        Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream, New ExcelReaderConfiguration() With {
                    .FallbackEncoding = encoding
                })
                            ' Create a new DataTable to hold the data
                            Dim dataTable As New DataTable()

                            ' Read the header row
                            reader.Read()

                            ' Add columns to the DataTable based on the header row
                            For i As Integer = 0 To reader.FieldCount - 1
                                dataTable.Columns.Add(reader.GetString(i))
                            Next

                            ' Read the remaining rows and populate the DataTable
                            While reader.Read()
                                Dim row As DataRow = dataTable.NewRow()

                                For i As Integer = 0 To reader.FieldCount - 1
                                    row(i) = reader.GetValue(i)
                                Next

                                dataTable.Rows.Add(row)
                            End While

                            ' Display the data in the DataGridView
                            DataGridView7.DataSource = dataTable


                            CheckStatusColumn(dataTable)

                            ' Set the first date and last date in the DateTimePicker controls
                            If dataTable.Rows.Count > 0 Then
                                Dim oldestDate As Date = DateTime.MaxValue
                                Dim latestDate As Date = DateTime.MinValue

                                For Each row As DataRow In dataTable.Rows
                                    Dim currentDate As Date = CDate(row("Date"))
                                    If currentDate < oldestDate Then
                                        oldestDate = currentDate
                                    End If
                                    If currentDate > latestDate Then
                                        latestDate = currentDate
                                    End If
                                Next

                                DateTimePickerStartinRetrieve.Value = oldestDate
                                DateTimePickerEndinRetrieve.Value = latestDate
                            End If
                        End Using
                    End Using
                End If

                ' Delete the temporary directory and its contents
                Directory.Delete(tempDir, True)
            End If
        Catch ex As Exception
            ' Handle the specific error here
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    Private Sub CheckStatusColumn(dataTable As DataTable)
        Dim hasStatusColumn As Boolean = dataTable.Columns.Contains("Status")
        Dim hasAttendanceStatusColumn As Boolean = dataTable.Columns.Contains("Attendance_Status")

        If hasStatusColumn Then
            ComboBoxStatusinRetrieve.Items.Clear()
            ComboBoxStatusinRetrieve.Items.AddRange({"All", "Success", "Failed"})
            ComboBoxStatusinRetrieve.SelectedItem = "All"
        End If

        If hasAttendanceStatusColumn Then
            ComboBoxStatusinRetrieve.Items.Clear()
            ComboBoxStatusinRetrieve.Items.AddRange({"All", "IN", "OUT"})
            ComboBoxStatusinRetrieve.SelectedItem = "All"
        End If
    End Sub

    Private Sub ModifyFilterExpression(selectedStatus As String, ByRef filterExpression As String, hasStatusColumn As Boolean, hasAttendanceStatusColumn As Boolean)
        If selectedStatus <> "" AndAlso selectedStatus <> "All" Then
            If filterExpression <> "" Then
                filterExpression &= " AND "
            End If

            ' Replace the existing filter expression with the new one
            If hasStatusColumn Then
                filterExpression &= $"Status = '{selectedStatus}'"
            ElseIf hasAttendanceStatusColumn Then
                filterExpression &= $"Attendance_Status = '{selectedStatus}'"
            End If
        End If
    End Sub


    Private Sub ApplyFilters()
        Try
            Dim selectedDepartment As String = If(ComboBoxDepartmentinRetrieve.SelectedItem IsNot Nothing, ComboBoxDepartmentinRetrieve.SelectedItem.ToString(), "")
            Dim selectedStatus As String = If(ComboBoxStatusinRetrieve.SelectedItem IsNot Nothing, ComboBoxStatusinRetrieve.SelectedItem.ToString(), "")
            Dim searchValue As String = TextBoxinRetrievePanel.Text.Trim()
            Dim startDate As Date = DateTimePickerStartinRetrieve.Value
            Dim endDate As Date = DateTimePickerEndinRetrieve.Value

            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = DataGridView7.DataSource

            Dim filterExpression As String = ""

            If selectedDepartment <> "All" Then
                filterExpression = $"Department = '{selectedDepartment}'"
            End If

            Dim hasStatusColumn As Boolean = DataGridView7.Columns.Contains("Status")
            Dim hasAttendanceStatusColumn As Boolean = DataGridView7.Columns.Contains("Attendance_Status")
            ModifyFilterExpression(selectedStatus, filterExpression, hasStatusColumn, hasAttendanceStatusColumn)

            If searchValue <> "" Then
                Dim searchExpression As String = ""
                If CheckBoxNameinRetrieve.Checked Then
                    searchExpression &= $"Student_Name LIKE '%{searchValue}%'"
                End If

                If CheckBoxStudent_IDinRetrieve.Checked Then
                    If searchExpression <> "" Then
                        searchExpression &= " OR "
                    End If
                    searchExpression &= $"Student_ID LIKE '%{searchValue}%'"
                End If

                If filterExpression <> "" AndAlso searchExpression <> "" Then
                    filterExpression &= " AND "
                End If
                filterExpression &= searchExpression
            End If

            ' Filter by date range
            If startDate <= endDate Then
                If filterExpression <> "" Then
                    filterExpression &= " AND "
                End If
                filterExpression &= $"Date >= #{startDate.ToString("MM/dd/yyyy")}# AND Date <= #{endDate.ToString("MM/dd/yyyy")}#"
            End If

            bindingSource.Filter = filterExpression
            DataGridView7.DataSource = bindingSource
        Catch ex As Exception
            ' Handle the specific error here
            MessageBox.Show("An error occurred while applying filters: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub ComboBoxStatusinRetrieve_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxStatusinRetrieve.SelectedIndexChanged
        ApplyFilters()
    End Sub

    Private Sub ComboBoxDepartmentinRetrieve_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxDepartmentinRetrieve.SelectedIndexChanged
        ApplyFilters()
    End Sub

    Private Sub TextBoxinRetrievePanel_TextChanged(sender As Object, e As EventArgs) Handles TextBoxinRetrievePanel.TextChanged
        ApplyFilters()
    End Sub

    Private Sub ButtonClearinRetrieve_Click(sender As Object, e As EventArgs) Handles ButtonClearinRetrieve.Click
        TextBoxinRetrievePanel.Text = ""
        CheckBoxNameinRetrieve.Checked = True
        ComboBoxDepartmentinRetrieve.SelectedIndex = 0
        DateTimePickerStartinRetrieve.Value = DateTime.Today
        DateTimePickerEndinRetrieve.Value = DateTime.Today
    End Sub

    Private Sub DateTimePickerEndinRetrieve_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerEndinRetrieve.ValueChanged
        ApplyFilters()
    End Sub

    Private Sub DateTimePickerStartinRetrieve_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerStartinRetrieve.ValueChanged
        ApplyFilters()
    End Sub

    Private Sub CheckBoxNameinRetrieve_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxNameinRetrieve.CheckedChanged
        If CheckBoxNameinRetrieve.Checked = True Then
            CheckBoxStudent_IDinRetrieve.Checked = False
            ApplyFilters()
        End If
        If CheckBoxNameinRetrieve.Checked = False Then
            CheckBoxStudent_IDinRetrieve.Checked = True
        End If
    End Sub

    Private Sub CheckBoxStudent_IDinRetrieve_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxStudent_IDinRetrieve.CheckedChanged
        If CheckBoxStudent_IDinRetrieve.Checked = True Then
            CheckBoxNameinRetrieve.Checked = False
            ApplyFilters()
        End If
        If CheckBoxStudent_IDinRetrieve.Checked = False Then
            CheckBoxNameinRetrieve.Checked = True
        End If
    End Sub

    '*************************************************
    '
    '
    '   All the Component of Student Report Panel
    '
    '
    '*************************************************


    'Display Student data based on Year level and also Search function in Datagridview4
    Private Function ShowDataByYearLevel(yearLevel As String, searchByName As Boolean, searchByID As Boolean, searchKeyword As String) As DataTable
        Dim dataTable As New DataTable()
        Try
            Dim query As String

            If yearLevel = "All" Then
                query = "SELECT `Student_ID`, `Firstname`, `Middlename`, `Lastname`, `Age`, `Department`, `Course`, `Year`, `Parent_Number`, `RFID_UID` FROM " & Table_Name & " WHERE 1 = 1"
            Else
                query = "SELECT `Student_ID`, `Firstname`, `Middlename`, `Lastname`, `Age`, `Department`, `Course`, `Year`, `Parent_Number`, `RFID_UID` FROM " & Table_Name & " WHERE `Department` = @YearLevel"
            End If

            If searchByName Then
                query &= " AND (`Firstname` LIKE @Keyword OR `Middlename` LIKE @Keyword OR `Lastname` LIKE @Keyword)"
            End If

            If searchByID Then
                query &= " AND `Student_ID` LIKE @Keyword"
            End If

            query &= " ORDER BY ID"

            Using connection As New MySqlConnection(connectionString)
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    command.Parameters.AddWithValue("@YearLevel", yearLevel)

                    If searchByName Or searchByID Then
                        command.Parameters.AddWithValue("@Keyword", "%" & searchKeyword & "%")
                    End If

                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End Using
            End Using

            DataGridView4.DefaultCellStyle.ForeColor = Color.Black
            DataGridView4.ClearSelection()
            DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView4.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
        Catch ex As MySqlException
            MessageBox.Show("A MySQL error occurred in the ShowDataByYearLevel function: " & ex.Message, "MySQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("An error occurred in the ShowDataByYearLevel function: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dataTable
    End Function

    Private Sub SearchTextboxStudentReport_TextChanged(sender As Object, e As EventArgs) Handles SearchTextboxStudentReport.TextChanged
        ComboBoxYearLevel_SelectedIndexChanged(ComboBoxYearLevel, EventArgs.Empty)
    End Sub

    Private Sub CheckBoxSearchNameStudentReport_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSearchNameStudentReport.CheckedChanged
        ComboBoxYearLevel_SelectedIndexChanged(ComboBoxYearLevel, EventArgs.Empty)
        If CheckBoxSearchNameStudentReport.Checked = True Then
            CheckBoxSearchByIDStudentReport.Checked = False
        End If
        If CheckBoxSearchNameStudentReport.Checked = False Then
            CheckBoxSearchByIDStudentReport.Checked = True
        End If
    End Sub

    Private Sub CheckBoxSearchByIDStudentReport_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSearchByIDStudentReport.CheckedChanged
        ComboBoxYearLevel_SelectedIndexChanged(ComboBoxYearLevel, EventArgs.Empty)
        If CheckBoxSearchByIDStudentReport.Checked = True Then
            CheckBoxSearchNameStudentReport.Checked = False
        End If
        If CheckBoxSearchByIDStudentReport.Checked = False Then
            CheckBoxSearchNameStudentReport.Checked = True
        End If
    End Sub

    Private Sub ComboBoxYearLevel_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBoxYearLevel.SelectedIndexChanged
        If ComboBoxYearLevel.SelectedItem IsNot Nothing Then
            Dim yearLevel As String = ComboBoxYearLevel.SelectedItem.ToString()
            Dim searchByName As Boolean = CheckBoxSearchNameStudentReport.Checked
            Dim searchByID As Boolean = CheckBoxSearchByIDStudentReport.Checked
            Dim searchKeyword As String = SearchTextboxStudentReport.Text

            Dim newData As DataTable = ShowDataByYearLevel(yearLevel, searchByName, searchByID, searchKeyword)
            DataGridView4.DataSource = newData
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem1.Click
        Dim yearLevel As String = ComboBoxYearLevel.SelectedItem.ToString()
        Dim searchByName As Boolean = CheckBoxSearchNameStudentReport.Checked
        Dim searchByID As Boolean = CheckBoxSearchByIDStudentReport.Checked
        Dim searchKeyword As String = SearchTextboxStudentReport.Text

        Dim newData As DataTable = ShowDataByYearLevel(yearLevel, searchByName, searchByID, searchKeyword)
        DataGridView4.DataSource = newData
    End Sub

    Private Sub ComboBoxYearLevel_SelectedIndexChanged(sender As Object, e As EventArgs)
        If ComboBoxYearLevel.SelectedItem IsNot Nothing Then
            Dim yearLevel As String = ComboBoxYearLevel.SelectedItem.ToString()
            Dim searchByName As Boolean = CheckBoxSearchNameStudentReport.Checked
            Dim searchByID As Boolean = CheckBoxSearchByIDStudentReport.Checked
            Dim searchKeyword As String = SearchTextboxStudentReport.Text

            Dim newData As DataTable = ShowDataByYearLevel(yearLevel, searchByName, searchByID, searchKeyword)
            DataGridView4.DataSource = newData
        End If
    End Sub

    'Count all the Senior High
    Public Sub CountSeniorHighStudents()
        Dim query As String = "SELECT COUNT(*) FROM " & Table_Name & " WHERE Department = 'Senior High'"

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim successCount As Integer = Convert.ToInt32(command.ExecuteScalar())

                    ' Assuming you have a Label control named lblSeniorHighCount on your form
                    lblSeniorHighCount.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting Senior High students: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Count all the College
    Public Sub CountCollegeStudents()
        Dim query As String = "SELECT COUNT(*) FROM " & Table_Name & " WHERE Department = 'College'"

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim successCount As Integer = Convert.ToInt32(command.ExecuteScalar())

                    ' Assuming you have a Label control named lblCollegeCount on your form
                    lblCollegeCount.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting College students: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Count all Students
    Public Sub CountAllStudents()
        Dim query As String = "SELECT COUNT(*) FROM " & Table_Name

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim successCount As Integer = Convert.ToInt32(command.ExecuteScalar())

                    ' Assuming you have a Label control named lblTotalStudents on your form
                    lblTotalStudents.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting all students: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Print function in Student Report
    Private Sub ButtonPrint1_Click(sender As Object, e As EventArgs) Handles ButtonPrint1.Click
        rowIndex = 0 ' Reset the rowIndex variable

        Dim printDialog As New PrintDialog()
        PrintDocument1.DefaultPageSettings.Landscape = True ' Set the page orientation to landscape if needed

        If printDialog.ShowDialog() = DialogResult.OK Then
            Try
                PrintDocument1.Print()
                MessageBox.Show("Print Successful!")
            Catch ex As Exception
                MessageBox.Show("Print Error: " & ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim dataGridView As DataGridView = DataGridView4 ' Replace DataGridView4 with your actual DataGridView control name
        Dim startX As Integer = 10
        Dim startY As Integer = 10
        Dim cellHeight As Integer = 40
        Dim cellWidth As Integer = 110

        Try
            ' Print column headers
            For i As Integer = 0 To dataGridView.ColumnCount - 1
                e.Graphics.FillRectangle(Brushes.LightGray, New Rectangle(startX, startY, cellWidth, cellHeight))
                e.Graphics.DrawRectangle(Pens.Black, New Rectangle(startX, startY, cellWidth, cellHeight))
                e.Graphics.DrawString(dataGridView.Columns(i).HeaderText, dataGridView.Font, Brushes.Black, New RectangleF(startX, startY, cellWidth, cellHeight))
                startX += cellWidth
            Next

            startY += cellHeight

            ' Print rows
            While rowIndex < dataGridView.Rows.Count AndAlso rowIndex < 10
                startX = 10

                For columnIndex As Integer = 0 To dataGridView.Columns.Count - 1
                    e.Graphics.DrawRectangle(Pens.Black, New Rectangle(startX, startY, cellWidth, cellHeight))
                    e.Graphics.DrawString(dataGridView.Rows(rowIndex).Cells(columnIndex).FormattedValue.ToString(), dataGridView.Font, Brushes.Black, New RectangleF(startX, startY, cellWidth, cellHeight))
                    startX += cellWidth
                Next

                startY += cellHeight
                rowIndex += 1
            End While

            ' If there are more entities to print, set the HasMorePages property to True
            If rowIndex < dataGridView.Rows.Count Then
                e.HasMorePages = True
            Else
                e.HasMorePages = False
            End If
        Catch ex As Exception
            MessageBox.Show("Print Error: " & ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ButtonArchive1_Click(sender As Object, e As EventArgs) Handles ButtonArchive1.Click
        PanelExport1.Visible = True
    End Sub

    Private Sub ButtonCancel1_Click(sender As Object, e As EventArgs) Handles ButtonCancel1.Click
        PanelExport1.Visible = False
    End Sub

    Private Sub ButtonExport1_Click(sender As Object, e As EventArgs) Handles ButtonExport1.Click
        ExportToFolder1("Student Report")
    End Sub

    Private Sub ExportToFolder1(folderName As String)
        Dim archiveName As String = txtArchiveName1.Text.Trim()

        If String.IsNullOrWhiteSpace(archiveName) Then
            MessageBox.Show("Archive name cannot be empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim archiveFolderPath As String = Path.Combine(GetUserFolderPath(), folderName)

            Try
                ArchiveDataGridViewData(DataGridView4, archiveFolderPath, archiveName)
                PanelExport1.Visible = False
                txtArchiveName1.Text = ""
            Catch ex As Exception
                MessageBox.Show("An error occurred while creating the archive: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    '*************************************************
    '
    '
    '   All the Component of Message Report Panel
    '
    '
    '*************************************************


    'Display Sms_Message Data in Datagridview5
    Private Function ShowDataInMessageReport(successOrFail As String, searchKeyword As String, dateFilter As String) As DataTable
        Dim dataTable As New DataTable()
        Try
            Dim query As String

            If successOrFail = "All" Then
                query = "SELECT `Student_ID`, `Student_Name`, `Message`, `Department`, `Phone_number`, `Status`, `Time`, `Date` FROM " & Table_Name4 & " WHERE 1 = 1"
            Else
                query = "SELECT `Student_ID`, `Student_Name`, `Message`, `Department`,`Phone_number`, `Status`, `Time`, `Date` FROM " & Table_Name4 & " WHERE `Status` = @SuccessOrFail"
            End If

            If searchKeyword <> "" Then
                If CheckBoxInMessageReportFindByStudent_ID.Checked Then
                    query &= " AND `Student_ID` LIKE @Keyword"
                ElseIf CheckBoxSearchNameStudentReport.Checked Then
                    query &= " AND `Student_Name` LIKE @Keyword"
                Else
                    query &= " AND (`Student_ID` LIKE @Keyword OR `Student_Name` LIKE @Keyword OR `Message` LIKE @Keyword OR `Phone_number` LIKE @Keyword)"
                End If
            End If

            If dateFilter = "Today" Then
                query &= " AND `Date` = CURDATE()"
            ElseIf dateFilter = "This Week" Then
                query &= " AND YEARWEEK(`Date`, 1) = YEARWEEK(CURDATE(), 1)"
            ElseIf dateFilter = "This Month" Then
                query &= " AND MONTH(`Date`) = MONTH(CURDATE()) AND YEAR(`Date`) = YEAR(CURDATE())"
            End If

            query &= " ORDER BY ID"

            Using connection As New MySqlConnection(connectionString)
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    command.Parameters.AddWithValue("@SuccessOrFail", successOrFail)

                    If searchKeyword <> "" Then
                        command.Parameters.AddWithValue("@Keyword", "%" & searchKeyword & "%")
                    End If

                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End Using
            End Using

            DataGridView5.DefaultCellStyle.ForeColor = Color.Black
            DataGridView5.ClearSelection()
            DataGridView5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView5.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
        Catch ex As MySqlException
            MessageBox.Show("A MySQL error occurred in the ShowDataInMessageReport function: " & ex.Message, "MySQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("An error occurred in the ShowDataInMessageReport function: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dataTable
    End Function

    Private Sub SearchBoxInMessageReport_TextChanged(sender As Object, e As EventArgs) Handles SearchBoxInMessageReport.TextChanged
        ComboBoxInMessageReport_SelectedIndexChanged(ComboBoxInMessageReport, EventArgs.Empty)
    End Sub

    Private Sub CheckBoxInMessageReportFindByName_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxInMessageReportFindByName.CheckedChanged
        If CheckBoxInMessageReportFindByName.Checked = True Then
            CheckBoxInMessageReportFindByStudent_ID.Checked = False
        End If
        If CheckBoxInMessageReportFindByName.Checked = False Then
            CheckBoxInMessageReportFindByStudent_ID.Checked = True
        End If
    End Sub

    Private Sub FilterMessageData()
        If ComboBoxInMessageReport.SelectedItem IsNot Nothing AndAlso SearchBoxInMessageReport IsNot Nothing Then
            Dim successOrFail As String = ComboBoxInMessageReport.SelectedItem.ToString()
            Dim searchKeyword As String = SearchBoxInMessageReport.Text.Trim()

            Dim newData As DataTable = ShowDataInMessageReport(successOrFail, searchKeyword, ComboBoxDate.SelectedItem.ToString())
            DataGridView5.DataSource = newData
        End If
    End Sub

    'ComboBox for Message Success or Fail
    Private Sub ComboBoxInMessageReport_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxInMessageReport.SelectedIndexChanged
        If ComboBoxInMessageReport.SelectedItem IsNot Nothing Then
            Dim successOrFail As String = ComboBoxInMessageReport.SelectedItem.ToString()
            Dim searchKeyword As String = SearchBoxInMessageReport.Text

            Dim newData As DataTable = ShowDataInMessageReport(successOrFail, searchKeyword, ComboBoxDate.SelectedItem.ToString())
            DataGridView5.DataSource = newData
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem4.Click
        Dim successOrFail As String = ComboBoxInMessageReport.SelectedItem.ToString()
        Dim searchKeyword As String = SearchBoxInMessageReport.Text

        Dim newData As DataTable = ShowDataInMessageReport(successOrFail, searchKeyword, ComboBoxDate.SelectedItem.ToString())
        DataGridView5.DataSource = newData
    End Sub

    Private Sub CheckBoxInMessageReportFindByStudent_ID_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxInMessageReportFindByStudent_ID.CheckedChanged
        If CheckBoxInMessageReportFindByStudent_ID.Checked = True Then
            CheckBoxInMessageReportFindByName.Checked = False
            FilterMessageData()
        End If
        If CheckBoxInMessageReportFindByStudent_ID.Checked = False Then
            CheckBoxInMessageReportFindByName.Checked = True
        End If
    End Sub

    Private Sub ComboBoxDate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxDate.SelectedIndexChanged
        If ComboBoxInMessageReport.SelectedItem IsNot Nothing Then
            Dim successOrFail As String = ComboBoxInMessageReport.SelectedItem.ToString()
            Dim searchKeyword As String = SearchBoxInMessageReport.Text
            Dim dateFilter As String = ComboBoxDate.SelectedItem.ToString()

            Dim newData As DataTable = ShowDataInMessageReport(successOrFail, searchKeyword, dateFilter)
            DataGridView5.DataSource = newData
        End If
    End Sub

    'Count all the Success Messages
    Public Sub CountSuccessMessages()
        Dim philippineTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Manila")
        Dim philippinesCurrentTime As DateTimeOffset = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, philippineTimeZone)
        Dim philippinesDate As String = philippinesCurrentTime.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT COUNT(*) FROM " & Table_Name4 & " WHERE Status = 'Success' AND Date = '" & philippinesDate & "';"

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim successCount As Integer = Convert.ToInt32(command.ExecuteScalar())

                    ' Assuming you have a Label control named LabelSuccessCount on your form
                    LabelSuccessCount.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting success messages: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshSuccessSMS_Tick(sender As Object, e As EventArgs) Handles RefreshSuccessSMS.Tick
        CountSuccessMessages()
    End Sub

    Public Sub CountFailMessages()
        Dim philippineTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Manila")
        Dim philippinesCurrentTime As DateTimeOffset = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, philippineTimeZone)
        Dim philippinesDate As String = philippinesCurrentTime.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT COUNT(*) FROM " & Table_Name4 & " WHERE Status = 'Failed' AND Date = '" & philippinesDate & "';"

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim successCount As Integer = Convert.ToInt32(command.ExecuteScalar())

                    ' Assuming you have a Label control named LabelFailCount on your form
                    LabelFailCount.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting fail messages: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshFailSMS_Tick(sender As Object, e As EventArgs) Handles RefreshFailSMS.Tick
        CountFailMessages()
    End Sub

    'Count all the Messages
    Public Sub CountAllSMS()
        Dim philippineTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Manila")
        Dim philippinesCurrentTime As DateTimeOffset = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, philippineTimeZone)
        Dim philippinesDate As String = philippinesCurrentTime.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT COUNT(*) FROM " & Table_Name4 & " WHERE Date = '" & philippinesDate & "';"

        Try
            Using connection As New MySqlConnection(connectionString)
                connection.Open()

                Using command As New MySqlCommand(query, connection)
                    Dim successCount As Integer = Convert.ToInt32(command.ExecuteScalar())

                    ' Assuming you have a Label control named LabelTotalofSMS on your form
                    LabelTotalofSMS.Text = successCount.ToString()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error counting all messages: " & ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshSMStotal_Tick(sender As Object, e As EventArgs) Handles RefreshSMStotal.Tick
        CountAllSMS()
    End Sub

    'Print function in Message Report
    Private Sub ButtonPrint2_Click(sender As Object, e As EventArgs) Handles ButtonPrint2.Click
        rowIndex = 0 ' Reset the rowIndex variable

        Dim printDialog As New PrintDialog()
        PrintDocument2.DefaultPageSettings.Landscape = True ' Set the page orientation to landscape if needed

        If printDialog.ShowDialog() = DialogResult.OK Then
            Try
                PrintDocument2.Print()
                MessageBox.Show("Print Successful!")
            Catch ex As Exception
                MessageBox.Show("Print Error: " & ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub PrintDocument2_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument2.PrintPage
        Dim dataGridView As DataGridView = DataGridView5 ' Replace DataGridView5 with your actual DataGridView control name
        Dim startX As Integer = 10
        Dim startY As Integer = 10
        Dim cellHeight As Integer = 120
        Dim cellWidth As Integer = 140

        Try
            ' Print column headers
            For i As Integer = 0 To dataGridView.ColumnCount - 1
                e.Graphics.FillRectangle(Brushes.LightGray, New Rectangle(startX, startY, cellWidth, cellHeight))
                e.Graphics.DrawRectangle(Pens.Black, New Rectangle(startX, startY, cellWidth, cellHeight))
                e.Graphics.DrawString(dataGridView.Columns(i).HeaderText, dataGridView.Font, Brushes.Black, New RectangleF(startX, startY, cellWidth, cellHeight))
                startX += cellWidth
            Next

            startY += cellHeight

            ' Print rows
            While rowIndex < dataGridView.Rows.Count AndAlso rowIndex < 10
                startX = 10

                For columnIndex As Integer = 0 To dataGridView.Columns.Count - 1
                    e.Graphics.DrawRectangle(Pens.Black, New Rectangle(startX, startY, cellWidth, cellHeight))
                    e.Graphics.DrawString(dataGridView.Rows(rowIndex).Cells(columnIndex).FormattedValue.ToString(), dataGridView.Font, Brushes.Black, New RectangleF(startX, startY, cellWidth, cellHeight))
                    startX += cellWidth
                Next

                startY += cellHeight
                rowIndex += 1
            End While

            ' If there are more entities to print, set the HasMorePages property to True
            If rowIndex < dataGridView.Rows.Count Then
                e.HasMorePages = True
            Else
                e.HasMorePages = False
            End If
        Catch ex As Exception
            MessageBox.Show("Print Error: " & ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ButtonArchive2_Click(sender As Object, e As EventArgs) Handles ButtonArchive2.Click
        PanelExport2.Visible = True
    End Sub

    Private Sub ButtonCancel2_Click(sender As Object, e As EventArgs) Handles ButtonCancel2.Click
        PanelExport2.Visible = False
    End Sub

    Private Sub ButtonExport2_Click(sender As Object, e As EventArgs) Handles ButtonExport2.Click
        ExportToFolder2("Message Report")
    End Sub

    Private Sub ExportToFolder2(folderName As String)
        Dim archiveName As String = txtArchiveName2.Text.Trim()

        If String.IsNullOrWhiteSpace(archiveName) Then
            MessageBox.Show("Archive name cannot be empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim archiveFolderPath As String = Path.Combine(GetUserFolderPath(), folderName)

            Try
                ArchiveDataGridViewData(DataGridView5, archiveFolderPath, archiveName)
                PanelExport2.Visible = False
                txtArchiveName2.Text = ""
            Catch ex As Exception
                MessageBox.Show("An error occurred while creating the archive: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    '*************************************************
    '
    '
    '   All the Component of Attendance Report Panel
    '
    '
    '*************************************************


    'Show data in Datagridview6 in Attendance Report Panel
    Private Function ShowDataInAttendanceReport(dateFilter As String, departmentFilter As String, statusFilter As String, searchKeyword As String, startDate As Date, endDate As Date) As DataTable
        Dim dataTable As New DataTable()
        Try
            Dim query As String = "SELECT a.RFID_UID, a.Student_ID, a.Student_Name, a.Department, a.Course, a.Year, a.Status AS Attendance_Status, a.Time, a.Date, s.Status AS Sms_Status
                        FROM attendance a
                        INNER JOIN sms_messages s ON a.ID = s.Attendance_ID
                        WHERE 1 = 1"

            If dateFilter = "Today" Then
                query &= " AND a.Date = CURDATE()"
            ElseIf dateFilter = "This Week" Then
                query &= " AND YEARWEEK(a.Date, 1) = YEARWEEK(CURDATE(), 1)"
            ElseIf dateFilter = "This Month" Then
                query &= " AND MONTH(a.Date) = MONTH(CURDATE()) AND YEAR(a.Date) = YEAR(CURDATE())"
            ElseIf dateFilter = "Date Range" Then
                query &= " AND a.Date BETWEEN @StartDate AND @EndDate"
            End If

            If departmentFilter <> "All" Then
                query &= " AND a.Department = @departmentFilter"
            End If

            If statusFilter <> "All" Then
                query &= " AND a.Status = @StatusFilter"
            End If

            If searchKeyword <> "" Then
                If CheckBoxInAttendanceReportFindByStudent_ID.Checked Then
                    query &= " AND a.Student_ID LIKE @Keyword"
                ElseIf CheckBoxInAttendanceReportFindByName.Checked Then
                    query &= " AND a.Student_Name LIKE @Keyword"
                End If
            End If

            query &= " ORDER BY a.Time DESC;"

            Using connection As New MySqlConnection(connectionString)
                connection.Open()
                Using command As New MySqlCommand(query, connection)
                    If departmentFilter <> "All" Then
                        command.Parameters.AddWithValue("@departmentFilter", departmentFilter)
                    End If

                    If statusFilter <> "All" Then
                        command.Parameters.AddWithValue("@StatusFilter", statusFilter)
                    End If

                    If searchKeyword <> "" Then
                        command.Parameters.AddWithValue("@Keyword", "%" & searchKeyword & "%")
                    End If

                    If dateFilter = "Date Range" Then
                        command.Parameters.AddWithValue("@StartDate", startDate)
                        command.Parameters.AddWithValue("@EndDate", endDate)
                    End If

                    Using adapter As New MySqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End Using
            End Using

            DataGridView6.DefaultCellStyle.ForeColor = Color.Black
            DataGridView6.ClearSelection()
            DataGridView6.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView6.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
            DataGridView6.DataSource = dataTable ' Update DataGridView6

        Catch ex As Exception
            MessageBox.Show("An error occurred in the ShowDataInAttendanceReport function: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dataTable
    End Function

    Private Sub DateTimePickerInAttendanceReportStart_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerInAttendanceReportStart.ValueChanged
        If ComboBoxDateinAttendance.SelectedItem IsNot Nothing Then
            Dim dateFilter As String = ComboBoxDateinAttendance.SelectedItem.ToString()
            Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
            Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
            Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
            Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

            Dim newData As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
            DataGridView6.DataSource = newData
        End If
    End Sub

    Private Sub DateTimePickerInAttendanceReportEnd_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerInAttendanceReportEnd.ValueChanged
        If ComboBoxDateinAttendance.SelectedItem IsNot Nothing Then
            Dim dateFilter As String = ComboBoxDateinAttendance.SelectedItem.ToString()
            Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
            Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
            Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
            Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

            Dim newData As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
            DataGridView6.DataSource = newData
        End If
    End Sub


    Private Sub TextBoxSearchinAttendanceReport_TextChanged(sender As Object, e As EventArgs) Handles TextBoxSearchinAttendanceReport.TextChanged
        If ComboBoxDateinAttendance.SelectedItem IsNot Nothing Then
            Dim dateFilter As String = ComboBoxDateinAttendance.SelectedItem.ToString()
            Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
            Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
            Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
            Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

            Dim newData As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
            DataGridView6.DataSource = newData
        End If
    End Sub

    Private Sub ComboBoxDateinAttendance_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxDateinAttendance.SelectedIndexChanged
        If ComboBoxDateinAttendance.SelectedItem IsNot Nothing Then
            Dim dateFilter As String = ComboBoxDateinAttendance.SelectedItem.ToString()
            Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
            Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
            Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
            Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

            Dim newData As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
            DataGridView6.DataSource = newData
        End If
    End Sub

    Private Sub ComboBoxYearlevelinAttendanceReport_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxYearlevelinAttendanceReport.SelectedIndexChanged
        If ComboBoxDateinAttendance.SelectedItem IsNot Nothing Then
            Dim dateFilter As String = ComboBoxDateinAttendance.SelectedItem.ToString()
            Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
            Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
            Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
            Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

            Dim newData As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
            DataGridView6.DataSource = newData
        End If
    End Sub

    Private Sub ComboBoxStatusinAttendanceReport_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxStatusinAttendanceReport.SelectedIndexChanged
        If ComboBoxDateinAttendance.SelectedItem IsNot Nothing Then
            Dim dateFilter As String = ComboBoxDateinAttendance.SelectedItem.ToString()
            Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
            Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
            Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
            Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
            Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

            Dim newData As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
            DataGridView6.DataSource = newData
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem3.Click
        Dim dateFilter As String = ComboBoxDateinAttendance.SelectedItem.ToString()
        Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
        Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
        Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
        Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
        Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

        Dim newData As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
        DataGridView6.DataSource = newData
    End Sub

    Private Sub CheckBoxInAttendanceReportFindByName_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxInAttendanceReportFindByName.CheckedChanged
        If CheckBoxInAttendanceReportFindByName.Checked = True Then
            CheckBoxInAttendanceReportFindByStudent_ID.Checked = False
        End If
        If CheckBoxInAttendanceReportFindByName.Checked = False Then
            CheckBoxInAttendanceReportFindByStudent_ID.Checked = True
        End If
    End Sub

    Private Sub CheckBoxInAttendanceReportFindByStudent_ID_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxInAttendanceReportFindByStudent_ID.CheckedChanged
        If CheckBoxInAttendanceReportFindByStudent_ID.Checked = True Then
            CheckBoxInAttendanceReportFindByName.Checked = False
        End If
        If CheckBoxInAttendanceReportFindByStudent_ID.Checked = False Then
            CheckBoxInAttendanceReportFindByName.Checked = True
        End If
    End Sub

    Private Sub ButtonClearinAttendanceReport_Click(sender As Object, e As EventArgs) Handles ButtonClearinAttendanceReport.Click
        ComboBoxDateinAttendance.SelectedIndex = 0
        ComboBoxYearlevelinAttendanceReport.SelectedIndex = 0
        ComboBoxStatusinAttendanceReport.SelectedIndex = 0
        TextBoxSearchinAttendanceReport.Text = ""
        DateTimePickerInAttendanceReportStart.Value = DateTime.Today
        DateTimePickerInAttendanceReportEnd.Value = DateTime.Today
        CheckBoxInAttendanceReportFindByName.Checked = True
    End Sub

    'Print function in Attendance Report
    Private Sub ButtonPrint3_Click(sender As Object, e As EventArgs) Handles ButtonPrint3.Click
        rowIndex = 0 ' Reset the rowIndex variable

        Dim printDialog As New PrintDialog()
        PrintDocument3.DefaultPageSettings.Landscape = True ' Set the page orientation to landscape if needed

        If printDialog.ShowDialog() = DialogResult.OK Then
            Try
                PrintDocument3.Print()
                MessageBox.Show("Print Successful!")
            Catch ex As Exception
                MessageBox.Show("Print Error: " & ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub PrintDocument3_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument3.PrintPage
        Dim dataGridView As DataGridView = DataGridView6 ' Replace DataGridView5 with your actual DataGridView control name
        Dim startX As Integer = 10
        Dim startY As Integer = 10
        Dim cellHeight As Integer = 40
        Dim cellWidth As Integer = 110

        Try
            ' Print column headers
            For i As Integer = 0 To dataGridView.ColumnCount - 1
                e.Graphics.FillRectangle(Brushes.LightGray, New Rectangle(startX, startY, cellWidth, cellHeight))
                e.Graphics.DrawRectangle(Pens.Black, New Rectangle(startX, startY, cellWidth, cellHeight))
                e.Graphics.DrawString(dataGridView.Columns(i).HeaderText, dataGridView.Font, Brushes.Black, New RectangleF(startX, startY, cellWidth, cellHeight))
                startX += cellWidth
            Next

            startY += cellHeight

            ' Print rows
            While rowIndex < dataGridView.Rows.Count AndAlso rowIndex < 10
                startX = 10

                For columnIndex As Integer = 0 To dataGridView.Columns.Count - 1
                    e.Graphics.DrawRectangle(Pens.Black, New Rectangle(startX, startY, cellWidth, cellHeight))
                    e.Graphics.DrawString(dataGridView.Rows(rowIndex).Cells(columnIndex).FormattedValue.ToString(), dataGridView.Font, Brushes.Black, New RectangleF(startX, startY, cellWidth, cellHeight))
                    startX += cellWidth
                Next

                startY += cellHeight
                rowIndex += 1
            End While

            ' If there are more entities to print, set the HasMorePages property to True
            If rowIndex < dataGridView.Rows.Count Then
                e.HasMorePages = True
            Else
                e.HasMorePages = False
            End If
        Catch ex As Exception
            MessageBox.Show("Print Error: " & ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ButtonArchive3_Click(sender As Object, e As EventArgs) Handles ButtonArchive3.Click
        PanelExport3.Visible = True
    End Sub

    Private Sub ButtonExportCancel_Click(sender As Object, e As EventArgs) Handles ButtonExportCancel3.Click
        PanelExport3.Visible = False
    End Sub

    Private Sub ButtonExport3_Click(sender As Object, e As EventArgs) Handles ButtonExport3.Click
        ExportToFolder3("Attendance Report")
    End Sub

    Private Sub ExportToFolder3(folderName As String)
        Dim archiveName As String = txtArchiveName3.Text.Trim()

        If String.IsNullOrWhiteSpace(archiveName) Then
            MessageBox.Show("Archive name cannot be empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim archiveFolderPath As String = Path.Combine(GetUserFolderPath(), folderName)

            Try
                ArchiveDataGridViewData(DataGridView6, archiveFolderPath, archiveName)
                PanelExport3.Visible = False
                txtArchiveName3.Text = ""
            Catch ex As Exception
                MessageBox.Show("An error occurred while creating the archive: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub



    '*************************************************
    '
    '
    '   Other Components
    '
    '
    '*************************************************

    'Show DashboardPanel
    Private Sub ButtonDashboard_Click(sender As Object, e As EventArgs) Handles ButtonDashboard.Click
        PictureBoxSelect.Top = ButtonDashboard.Top
        PanelDashboard.Visible = True
        PanelSearch.Visible = False
        PanelConnection.Visible = False
        PanelUserDetails.Visible = False
        PanelRegistration.Visible = False
        PanelStudentReport.Visible = False
        PanelMessageReport.Visible = False
        PanelRetrieve.Visible = False
        PanelAttendanceReport.Visible = False
    End Sub

    'Show SearchPanel
    Private Sub ButtonSearch_Click(sender As Object, e As EventArgs) Handles ButtonSearch.Click
        PictureBoxSelect.Top = ButtonSearch.Top
        PanelDashboard.Visible = False
        PanelSearch.Visible = True
        PanelConnection.Visible = False
        PanelUserDetails.Visible = False
        PanelRegistration.Visible = False
        PanelStudentReport.Visible = False
        PanelMessageReport.Visible = False
        PanelRetrieve.Visible = False
        PanelAttendanceReport.Visible = False
    End Sub

    'Show ConnectionPanel
    Private Sub ButtonConnection_Click(sender As Object, e As EventArgs) Handles ButtonConnection.Click
        PictureBoxSelect.Top = ButtonConnection.Top
        PanelDashboard.Visible = False
        PanelSearch.Visible = False
        PanelConnection.Visible = True
        PanelUserDetails.Visible = False
        PanelRegistration.Visible = False
        PanelStudentReport.Visible = False
        PanelMessageReport.Visible = False
        PanelRetrieve.Visible = False
        PanelAttendanceReport.Visible = False
    End Sub

    'Show Registration Panel
    Private Sub ButtonRegister_Click(sender As Object, e As EventArgs) Handles ButtonRegister.Click
        PictureBoxSelect.Top = ButtonRegister.Top
        PanelDashboard.Visible = False
        PanelSearch.Visible = False
        PanelConnection.Visible = False
        PanelUserDetails.Visible = False
        PanelRegistration.Visible = True
        PanelStudentReport.Visible = False
        PanelRetrieve.Visible = False
        PanelMessageReport.Visible = False
        PanelAttendanceReport.Visible = False
        PanelReadingTagProcess.Visible = False
        TextBoxID.Visible = False
        TextBoxStudentID.Visible = False
    End Sub
    Private Sub ButtonRetrieve_Click(sender As Object, e As EventArgs) Handles ButtonRetrieve.Click
        PictureBoxSelect.Top = ButtonRetrieve.Top
        PanelDashboard.Visible = False
        PanelSearch.Visible = False
        PanelConnection.Visible = False
        PanelUserDetails.Visible = False
        PanelRegistration.Visible = False
        PanelStudentReport.Visible = False
        PanelRetrieve.Visible = True
        PanelMessageReport.Visible = False
        PanelAttendanceReport.Visible = False
    End Sub

    'Show the Report Panel
    Private Sub ButtonReport_Click(sender As Object, e As EventArgs) Handles ButtonReport.Click
        Try
            PictureBoxSelect.Top = ButtonReport.Top
            Timer2.Start()

            If Not PanelDropdown.Height = 0 Then
                PanelDropdown.Height = 0
                ButtonLogout.Location = New Point(50, 520)
                Timer2.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred while processing the report button click event: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        PanelDropdown.Height = 157
        ButtonLogout.Location = New Point(50, 650)
        Timer2.Stop()
    End Sub

    'Show the User Details Panel
    Private Sub ButtonUserDetails_Click(sender As Object, e As EventArgs) Handles ButtonUserDetails.Click
        If TimerSerialIn.Enabled = False Then
            MsgBox("Failed to open User Data !!!" & vbCr & "Click the Connection menu then click the Connect button.", MsgBoxStyle.Information, "Information")
            Return
        Else
            StrSerialIn = ""
            ViewUserData = True
            PictureBoxSelect.Top = ButtonUserDetails.Top
            PanelDashboard.Visible = False
            PanelSearch.Visible = False
            PanelConnection.Visible = False
            PanelRetrieve.Visible = False
            PanelUserDetails.Visible = True
            PanelRegistration.Visible = False
            PanelStudentReport.Visible = False
            PanelMessageReport.Visible = False
            PanelAttendanceReport.Visible = False
        End If
    End Sub

    'Show Student Report Panel
    Private Sub ButtonStudent_Click(sender As Object, e As EventArgs) Handles ButtonStudent.Click
        PanelDashboard.Visible = False
        PanelSearch.Visible = False
        PanelConnection.Visible = False
        PanelUserDetails.Visible = False
        PanelRegistration.Visible = False
        PanelStudentReport.Visible = True
        PanelRetrieve.Visible = False
        PanelMessageReport.Visible = False
        PanelAttendanceReport.Visible = False

        ComboBoxYearLevel.SelectedIndex = 0
        Dim yearLevel As String = ComboBoxYearLevel.SelectedItem.ToString()
        Dim newData4 As DataTable = ShowDataByYearLevel(yearLevel, False, False, "")
        DataGridView4.DataSource = newData4
    End Sub

    'Show Message Report Panel
    Private Sub ButtonMessage_Click(sender As Object, e As EventArgs) Handles ButtonMessage.Click
        PanelDashboard.Visible = False
        PanelSearch.Visible = False
        PanelConnection.Visible = False
        PanelUserDetails.Visible = False
        PanelRegistration.Visible = False
        PanelStudentReport.Visible = False
        PanelRetrieve.Visible = False
        PanelMessageReport.Visible = True
        PanelAttendanceReport.Visible = False

        ComboBoxInMessageReport.SelectedIndex = 0
        Dim successOrFail = ComboBoxInMessageReport.SelectedItem.ToString
        Dim dateFilter = ComboBoxDate.SelectedItem.ToString
        Dim newData = ShowDataInMessageReport(successOrFail, "", dateFilter)
        DataGridView5.DataSource = newData
    End Sub

    'Show Message Report Panel
    Private Sub ButtonAttendance_Click(sender As Object, e As EventArgs) Handles ButtonAttendance.Click
        PanelDashboard.Visible = False
        PanelSearch.Visible = False
        PanelConnection.Visible = False
        PanelUserDetails.Visible = False
        PanelRegistration.Visible = False
        PanelStudentReport.Visible = False
        PanelRetrieve.Visible = False
        PanelMessageReport.Visible = False
        PanelAttendanceReport.Visible = True

        Dim dateFilter As String = ComboBoxDate.SelectedItem.ToString()
        Dim dateFilter6 As String = ComboBoxDateinAttendance.SelectedItem.ToString()
        Dim departmentFilter As String = ComboBoxYearlevelinAttendanceReport.SelectedItem.ToString()
        Dim statusFilter As String = ComboBoxStatusinAttendanceReport.SelectedItem.ToString()
        Dim searchKeyword As String = TextBoxSearchinAttendanceReport.Text
        Dim startDate As Date = DateTimePickerInAttendanceReportStart.Value.Date
        Dim endDate As Date = DateTimePickerInAttendanceReportEnd.Value.Date

        Dim newData6 As DataTable = ShowDataInAttendanceReport(dateFilter, departmentFilter, statusFilter, searchKeyword, startDate, endDate)
        DataGridView6.DataSource = newData6
    End Sub

    'Show the data and time
    Private Sub TimerDateAndTime_Tick(sender As Object, e As EventArgs) Handles TimerDateandTime.Tick
        LabelDateandTime.Text = "Time: " & DateTime.Now.ToString("HH:mm:ss") & "  Date: " & DateTime.Now.ToString("dd MMM, yyyy")
    End Sub

    'Logout Function
    Private Sub ButtonLogout_Click(sender As Object, e As EventArgs) Handles ButtonLogout.Click
        Dim dialog As DialogResult

        dialog = MessageBox.Show("Do you really want to Exit the app?", "Exit", MessageBoxButtons.YesNo)
        If dialog = DialogResult.No Then
            Me.Cancel = True
        Else
            Me.Hide() 'hide the current form
            Dim LoginPage As New LoginPage() 'create an instance of the login form
            LoginPage.Show() 'show the login form
        End If
    End Sub

    Private Sub ArchiveDataGridViewData(dataGridView As DataGridView, archiveFolderPath As String, archiveName As String)
        ' Set the LicenseContext
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        ' Convert DataGridView to DataTable
        Dim data As DataTable = ConvertDataGridViewToDataTable(dataGridView)

        ' Create a new folder
        Directory.CreateDirectory(archiveFolderPath)

        ' Check if the archive already exists
        Dim archiveFilePath As String = Path.Combine(archiveFolderPath, $"{archiveName}.zip")
        If File.Exists(archiveFilePath) Then
            MessageBox.Show("Archive already exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ' Exit the method
        End If

        ' Convert DataTable to Excel file
        Dim excelFilePath As String = Path.Combine(archiveFolderPath, $"{archiveName}.xlsx")
        Try
            ConvertDataTableToExcel(data, excelFilePath, archiveName)

            ' Create a new ZIP archive
            Using archive As ZipArchive = ZipFile.Open(archiveFilePath, ZipArchiveMode.Create)
                ' Add the Excel file to the archive
                archive.CreateEntryFromFile(excelFilePath, $"{archiveName}.xlsx")
            End Using

            ' Delete the temporary Excel file
            File.Delete(excelFilePath)

            MessageBox.Show("Archive created successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As FileNotFoundException
            ' Handle the specific file not found error
            MessageBox.Show($"File not found: {ex.FileName}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            ' Handle any other exceptions
            MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function ConvertDataGridViewToDataTable(dataGridView As DataGridView) As DataTable
        Dim data As New DataTable()

        Try
            ' Create columns in DataTable
            For Each column As DataGridViewColumn In dataGridView.Columns
                data.Columns.Add(column.HeaderText, GetType(Object))
            Next

            ' Add rows to DataTable
            For Each row As DataGridViewRow In dataGridView.Rows
                If Not row.IsNewRow Then
                    Dim dataRow As DataRow = data.Rows.Add()
                    For Each cell As DataGridViewCell In row.Cells
                        dataRow(cell.ColumnIndex) = cell.Value
                    Next
                End If
            Next
        Catch ex As Exception
            ' Handle the specific error here
            MessageBox.Show("An error occurred while converting DataGridView to DataTable: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return data
    End Function


    Private Function GetUserFolderPath() As String
        Dim userFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Return Path.Combine(userFolder, "Monitor_Archive_Folder")
    End Function

    Private Sub ConvertDataTableToExcel(data As DataTable, filePath As String, archiveName As String)
        Try
            Using package As New ExcelPackage()
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(archiveName)

                ' Write header row
                Dim columnIndex As Integer = 1
                For Each column As DataColumn In data.Columns
                    worksheet.Cells(1, columnIndex).Value = column.ColumnName

                    ' Apply bold font style to the header cells
                    worksheet.Cells(1, columnIndex).Style.Font.Bold = True
                    columnIndex += 1
                Next

                ' Write data rows
                Dim rowIndex As Integer = 2
                For Each row As DataRow In data.Rows
                    columnIndex = 1
                    For Each item As Object In row.ItemArray
                        Dim cellValue As Object = item
                        If TypeOf item Is DateTime Then
                            ' Format DateTime value to display only the date part
                            cellValue = DirectCast(item, DateTime).ToString("yyyy-MM-dd")
                        ElseIf TypeOf item Is TimeSpan Then
                            ' Format TimeSpan value to display as hh:mm:ss
                            cellValue = DirectCast(item, TimeSpan).ToString("hh\:mm\:ss")
                        End If
                        worksheet.Cells(rowIndex, columnIndex).Value = cellValue
                        columnIndex += 1
                    Next
                    rowIndex += 1
                Next

                ' Auto-fit columns
                worksheet.Cells.AutoFitColumns()

                ' Save Excel file
                Dim fileInfo As New FileInfo(filePath)
                package.SaveAs(fileInfo)
            End Using
        Catch ex As Exception
            ' Handle the specific error here
            MessageBox.Show("An error occurred while converting DataTable to Excel: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub MainPage_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        ' Check if the user wants to close the form
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to close the Application?", "Form Closing", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' If the user clicks No, cancel the form closing event
        If result = DialogResult.No Then
            e.Cancel = True
        End If
    End Sub


End Class