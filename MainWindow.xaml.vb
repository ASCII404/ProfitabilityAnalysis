Class MainWindow
    Private dbHelper As Database

    Public Sub New()
        InitializeComponent()
        dbHelper = New Database("dbtest.db")
    End Sub

    ' Event handlers to switch tabs
    Private Sub DashboardButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = DashboardTab
    End Sub

    Private Sub PreviewDatabaseButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = PreviewDatabaseTab
    End Sub

    Private Sub AnalysisButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = AnalysisTab
    End Sub

    Private Sub InputDataButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = InputDataTab
    End Sub

    Private Sub ScenarioAnalysisButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = ScenarioAnalysisTab
    End Sub

    ' Event handler for adding a user
    Private Sub AddUserButton_Click(sender As Object, e As RoutedEventArgs)
        Dim newUser As New UserTest With {
            .Username = "new_user",
            .Email = "new_user@example.com",
            .PasswordHash = "hashed_password"
        }
        dbHelper.AddUser(newUser)
        MessageBox.Show("User added successfully!")
    End Sub

    ' Event handler for previewing users
    Private Sub PreviewUsersButton_Click(sender As Object, e As RoutedEventArgs)
        Dim users As List(Of UserTest) = dbHelper.GetAllUsers()
        UsersListBox.Items.Clear()
        For Each user In users
            UsersListBox.Items.Add($"{user.UserID}: {user.Username} - {user.Email}")
        Next
    End Sub
End Class