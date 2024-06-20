
Public Class Authentication
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Delete_user_Click(sender As Object, e As RoutedEventArgs) Handles Delete_user.Click
        MessageBox.Show("Are you sure you want to delete the user? Everything will be lost.", "Delete user", MessageBoxButton.YesNoCancel)
        If MessageBoxButton.YesNoCancel = MessageBoxResult.Yes Then
            If User_name.Text <> "" And Password.Password <> "" Then
                User_name.Text = ""
                Password.Password = "" ' Assuming you also want to clear the password field
            End If

            MessageBox.Show("User deleted", "Delete user", MessageBoxButton.OK)
        End If
    End Sub

    Private Sub Login_Click(sender As Object, e As RoutedEventArgs) Handles Login.Click
        If User_name.Text <> "" And Password.Password <> "" Then
            User_name.Text = ""
            Password.Password = ""
            Me.DialogResult = True
        End If

        Me.Close()
        Debug.WriteLine("Login button clicked")
    End Sub

End Class
