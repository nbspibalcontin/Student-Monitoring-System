<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoginPage
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Panel1 = New Panel()
        loading = New PictureBox()
        ButtonLogin = New Button()
        CheckBox1 = New CheckBox()
        Label2 = New Label()
        Label1 = New Label()
        TextBoxPassword = New TextBox()
        TextBoxUsername = New TextBox()
        PictureBox1 = New PictureBox()
        Panel1.SuspendLayout()
        CType(loading, ComponentModel.ISupportInitialize).BeginInit()
        CType(PictureBox1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' Panel1
        ' 
        Panel1.BackColor = SystemColors.AppWorkspace
        Panel1.Controls.Add(loading)
        Panel1.Controls.Add(ButtonLogin)
        Panel1.Controls.Add(CheckBox1)
        Panel1.Controls.Add(Label2)
        Panel1.Controls.Add(Label1)
        Panel1.Controls.Add(TextBoxPassword)
        Panel1.Controls.Add(TextBoxUsername)
        Panel1.Controls.Add(PictureBox1)
        Panel1.Location = New Point(0, 0)
        Panel1.Name = "Panel1"
        Panel1.Size = New Size(801, 451)
        Panel1.TabIndex = 0
        ' 
        ' loading
        ' 
        loading.Image = My.Resources.Resources.output_onlinegiftools
        loading.Location = New Point(473, 348)
        loading.Name = "loading"
        loading.Size = New Size(50, 38)
        loading.SizeMode = PictureBoxSizeMode.StretchImage
        loading.TabIndex = 8
        loading.TabStop = False
        ' 
        ' ButtonLogin
        ' 
        ButtonLogin.Location = New Point(360, 348)
        ButtonLogin.Name = "ButtonLogin"
        ButtonLogin.Size = New Size(107, 38)
        ButtonLogin.TabIndex = 7
        ButtonLogin.Text = "Login"
        ButtonLogin.UseVisualStyleBackColor = True
        ' 
        ' CheckBox1
        ' 
        CheckBox1.AutoSize = True
        CheckBox1.Location = New Point(359, 314)
        CheckBox1.Name = "CheckBox1"
        CheckBox1.Size = New Size(108, 19)
        CheckBox1.TabIndex = 6
        CheckBox1.Text = "Show Password"
        CheckBox1.UseVisualStyleBackColor = True
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point)
        Label2.Location = New Point(271, 284)
        Label2.Name = "Label2"
        Label2.Size = New Size(77, 20)
        Label2.TabIndex = 5
        Label2.Text = "Password :"
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point)
        Label1.Location = New Point(271, 245)
        Label1.Name = "Label1"
        Label1.Size = New Size(82, 20)
        Label1.TabIndex = 4
        Label1.Text = "Username :"
        ' 
        ' TextBoxPassword
        ' 
        TextBoxPassword.Location = New Point(359, 285)
        TextBoxPassword.Name = "TextBoxPassword"
        TextBoxPassword.Size = New Size(154, 23)
        TextBoxPassword.TabIndex = 3
        TextBoxPassword.UseSystemPasswordChar = True
        ' 
        ' TextBoxUsername
        ' 
        TextBoxUsername.Location = New Point(359, 246)
        TextBoxUsername.Name = "TextBoxUsername"
        TextBoxUsername.Size = New Size(154, 23)
        TextBoxUsername.TabIndex = 2
        ' 
        ' PictureBox1
        ' 
        PictureBox1.Image = My.Resources.Resources.NBSPI_Text_Logo__Transparent_
        PictureBox1.Location = New Point(176, 30)
        PictureBox1.Name = "PictureBox1"
        PictureBox1.Size = New Size(420, 129)
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        PictureBox1.TabIndex = 1
        PictureBox1.TabStop = False
        ' 
        ' LoginPage
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(Panel1)
        Name = "LoginPage"
        Text = "Student Monitoring System"
        Panel1.ResumeLayout(False)
        Panel1.PerformLayout()
        CType(loading, ComponentModel.ISupportInitialize).EndInit()
        CType(PictureBox1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents ButtonLogin As Button
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBoxPassword As TextBox
    Friend WithEvents TextBoxUsername As TextBox
    Friend WithEvents loading As PictureBox
End Class
