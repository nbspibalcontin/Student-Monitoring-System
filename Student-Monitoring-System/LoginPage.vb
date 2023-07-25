Imports MySqlConnector

Public Class LoginPage
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            TextBoxPassword.UseSystemPasswordChar = False
        Else
            TextBoxPassword.UseSystemPasswordChar = True
        End If
    End Sub

    Private Sub LoginPage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        LoadingScreen.Hide()
        loading.Visible = False
    End Sub

    Private Sub LoginPage_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Check if the user wants to close the form
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to close the Application?", "Form Closing", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' If the user clicks No, cancel the form closing event
        If result = DialogResult.No Then
            e.Cancel = True
        End If
    End Sub

    Private Sub ButtonLogin_Click(sender As Object, e As EventArgs) Handles ButtonLogin.Click
        ' Show the loading PictureBox
        loading.Visible = True
        Application.DoEvents()

        ' Connect to the database
        Dim connectionString As String = "Data Source=localhost;Initial Catalog=monitorsystem;User ID=root;Password="

        Try
            Using Connection As New MySqlConnection(connectionString)
                Connection.Open()

                ' Execute the SQL query
                Dim query As String = "SELECT * FROM admin WHERE username=@username AND password=@password"
                Using command As New MySqlCommand(query, Connection)
                    command.Parameters.AddWithValue("@username", TextBoxUsername.Text)
                    command.Parameters.AddWithValue("@password", TextBoxPassword.Text)

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If reader.Read() Then
                            ' Hide the loading PictureBox and show the main form
                            loading.Visible = False
                            MainPage.Show()
                            Me.Hide()
                        Else
                            ' Hide the loading PictureBox and show an error message
                            loading.Visible = False
                            MessageBox.Show("Wrong Username or Password", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Hide the loading PictureBox and show an error message
            loading.Visible = False
            MessageBox.Show("Connection failed !!!" & vbCrLf & "Please check that the server is ready !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
    End Sub


End Class