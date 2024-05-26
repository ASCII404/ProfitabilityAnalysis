Imports System.Data.SQLite
Imports System.IO
Public Class Database
    Private connectionString As String

    Public Sub New(dbPath As String)
        connectionString = $"Data Source={dbPath};Version=3;"
        If Not File.Exists(dbPath) Then
            CreateDatabase()
        End If
    End Sub

    Private Sub CreateDatabase()
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            ' Create Users table
            Dim createUserTable As String = "
                CREATE TABLE IF NOT EXISTS Users (
                    UserID INTEGER PRIMARY KEY AUTOINCREMENT,
                    Username TEXT NOT NULL UNIQUE,
                    Email TEXT NOT NULL UNIQUE,
                    PasswordHash TEXT NOT NULL
                );"
            Dim cmd As New SQLiteCommand(createUserTable, connection)
            cmd.ExecuteNonQuery()

            ' Create Posts table
            Dim createPostTable As String = "
                CREATE TABLE IF NOT EXISTS Posts (
                    PostID INTEGER PRIMARY KEY AUTOINCREMENT,
                    Title TEXT NOT NULL,
                    Content TEXT NOT NULL,
                    CreatedAt DATETIME DEFAULT CURRENT_TIMESTAMP,
                    AuthorID INTEGER,
                    FOREIGN KEY (AuthorID) REFERENCES Users(UserID)
                );"
            cmd.CommandText = createPostTable
            cmd.ExecuteNonQuery()

            ' Create Comments table
            Dim createCommentTable As String = "
                CREATE TABLE IF NOT EXISTS Comments (
                    CommentID INTEGER PRIMARY KEY AUTOINCREMENT,
                    PostID INTEGER,
                    AuthorID INTEGER,
                    Content TEXT NOT NULL,
                    CreatedAt DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (PostID) REFERENCES Posts(PostID),
                    FOREIGN KEY (AuthorID) REFERENCES Users(UserID)
                );"
            cmd.CommandText = createCommentTable
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' Updated AddUser method to use UserTest class
    Public Sub AddUser(user As UserTest)
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Dim cmd As New SQLiteCommand("
                INSERT INTO Users (Username, Email, PasswordHash) 
                VALUES (@Username, @Email, @PasswordHash)", connection)
            cmd.Parameters.AddWithValue("@Username", user.Username)
            cmd.Parameters.AddWithValue("@Email", user.Email)
            cmd.Parameters.AddWithValue("@PasswordHash", user.PasswordHash)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' Method to get all users from the database
    Public Function GetAllUsers() As List(Of UserTest)
        Dim users As New List(Of UserTest)
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Dim cmd As New SQLiteCommand("SELECT UserID, Username, Email, PasswordHash FROM Users", connection)
            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read()
                    Dim user As New UserTest() With {
                        .UserID = reader("UserID"),
                        .Username = reader("Username"),
                        .Email = reader("Email"),
                        .PasswordHash = reader("PasswordHash")
                    }
                    users.Add(user)
                End While
            End Using
        End Using
        Return users
    End Function

    ' Similar methods for Posts and Comments can be added here
End Class