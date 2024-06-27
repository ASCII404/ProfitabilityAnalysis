Public Class Authentication
    Public Sub New()
        InitializeComponent()
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
