Imports System.IO
Imports System.Configuration
Imports Microsoft.Win32
Public Class Authentication
    Private configFilePath As String = "user.config"
    Public Sub New()
        InitializeComponent()

        If Not File.Exists(configFilePath) Then
            File.Create(configFilePath).Dispose()
        End If

        Dim fileInfo As New FileInfo(configFilePath)
        If fileInfo.Length <> 0 Then
            Debug.WriteLine("File is not empty")
            Register.IsEnabled = False
            Register.Background = Brushes.Gray
        End If

        File.SetAttributes(configFilePath, FileAttributes.Hidden)
    End Sub

    Private Sub Login_Click(sender As Object, e As RoutedEventArgs) Handles Login.Click
        ' Validate that the username and password fields are not empty
        If User_name.Text <> "" And Password.Password <> "" Then
            ' Perform the actual credential validation
            If ValidateCredentials(User_name.Text, Password.Password) Then
                ' Clear the input fields


                ' Open the main form
                Dim mainForm As New MainWindow(User_name.Text)
                mainForm.Show()
                User_name.Text = ""
                Password.Password = ""
                ' Close the authentication form
                Me.Close()
                Return
            Else
                MessageBox.Show("Invalid credentials, please try again.")
            End If
        End If

        Debug.WriteLine("Login button clicked")
    End Sub

    ' Method to check if the username exists
    Private Function UserExists(username As String) As Boolean
        ' Read all lines from the config file
        Dim lines As String() = File.ReadAllLines(configFilePath)
        For Each line In lines
            Dim parts As String() = line.Split(","c)
            If parts.Length > 0 AndAlso parts(0) = username Then
                Return True
            End If
        Next
        Return False
    End Function

    ' Method to register a new user
    Private Sub RegisterUser(username As String, password As String)
        ' Append the new user credentials to the config file
        Using writer As StreamWriter = New StreamWriter(configFilePath, True)
            writer.WriteLine($"{username},{password}")
        End Using
    End Sub

    ' Method to validate the credentials
    Private Function ValidateCredentials(username As String, password As String) As Boolean
        ' Read all lines from the config file
        Dim lines As String() = File.ReadAllLines(configFilePath)
        For Each line In lines
            Dim parts As String() = line.Split(","c)
            If parts.Length = 2 Then
                If parts(0) = username AndAlso parts(1) = password Then
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Sub Register_Click(sender As Object, e As RoutedEventArgs) Handles Register.Click
        Dim fileInfo As New FileInfo(configFilePath)

        If User_name.Text <> "" And Password.Password <> "" Then
            If Not UserExists(User_name.Text) Then
                RegisterUser(User_name.Text, Password.Password)
                MessageBox.Show("Registration successful. You can now log in.")
            Else
                MessageBox.Show("Username already exists. Please choose a different username.")
            End If
        Else
            MessageBox.Show("Please enter a username and password to register.")
        End If

        Debug.WriteLine("Register button clicked")
    End Sub
End Class
