Public Class Authentication
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Login_Click(sender As Object, e As RoutedEventArgs) Handles Login.Click
        ' Validate that the username and password fields are not empty
        If User_name.Text <> "" And Password.Password <> "" Then
            ' Perform the actual credential validation
            If ValidateCredentials(User_name.Text, Password.Password) Then
                ' Clear the input fields
                User_name.Text = ""
                Password.Password = ""

                ' Open the main form
                Dim mainForm As New MainWindow()
                mainForm.Show()

                ' Close the authentication form
                Me.Close()
                Return
            Else
                MessageBox.Show("Invalid credentials, please try again.")
            End If
        End If

        Debug.WriteLine("Login button clicked")
    End Sub

    ' Method to validate the credentials
    Private Function ValidateCredentials(username As String, password As String) As Boolean
        ' Example hard-coded validation, replace with actual logic
        Return username = "admin" AndAlso password = "password"
    End Function

End Class
