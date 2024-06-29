Imports System.Data.SQLite
Imports System.IO

Public Class Database
    Private connectionString As String

    ' Constructor: Creates a new SQLite database file if it doesn't exist
    ' Always done when the app starts
    Public Sub New(dbPath As String)
        connectionString = $"Data Source={dbPath};Version=3;"
        If Not File.Exists(dbPath) Then
            CreateDatabase()
        End If
    End Sub

    'It creates the tables in the database
    Private Sub CreateDatabase()
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            ' Create DateDimension table
            Dim createDateDimensionTable As String = "
                CREATE TABLE IF NOT EXISTS DateDimension (
                    DateID INTEGER PRIMARY KEY AUTOINCREMENT,
                    Date DATE NOT NULL,
                    Year INTEGER,
                    Quarter INTEGER,
                    Month INTEGER,
                    Day INTEGER
                );"
            Dim cmd As New SQLiteCommand(createDateDimensionTable, connection)
            cmd.ExecuteNonQuery()

            ' Create FinancialData table
            Dim createFinancialDataTable As String = "
                CREATE TABLE IF NOT EXISTS FinancialData (
                    FinancialDataID INTEGER PRIMARY KEY AUTOINCREMENT,
                    DateID INTEGER,
                    Revenue DECIMAL,
                    CostOfGoodsSold DECIMAL,
                    OperatingExpenses DECIMAL,
                    NetIncome DECIMAL,
                    TotalAssets DECIMAL,
                    TotalEquity DECIMAL,
                    EBITDA DECIMAL,
                    CurrentAssets DECIMAL,
                    CurrentLiabilities DECIMAL,
                    TotalLiabilities DECIMAL,
                    InterestExpense DECIMAL,
                    VariableCosts DECIMAL,
                    FixedCosts DECIMAL,
                    SalesRevenuePerUnit DECIMAL,
                    VariableCostPerUnit DECIMAL,
                    FOREIGN KEY (DateID) REFERENCES DateDimension(DateID)
                );"
            cmd.CommandText = createFinancialDataTable
            cmd.ExecuteNonQuery()

            ' Create Ratios table (optional, for storing calculated ratios)
            Dim createRatiosTable As String = "
                CREATE TABLE IF NOT EXISTS Ratios (
                    RatioID INTEGER PRIMARY KEY AUTOINCREMENT,
                    FinancialDataID INTEGER,
                    ROA DECIMAL,
                    ROE DECIMAL,
                    GrossProfitMargin DECIMAL,
                    OperatingProfitMargin DECIMAL,
                    NetProfitMargin DECIMAL,
                    EBITDAValue DECIMAL,
                    CurrentRatio DECIMAL,
                    DebtToEquityRatio DECIMAL,
                    InterestCoverageRatio DECIMAL,
                    ContributionMargin DECIMAL,
                    BreakEvenPoint DECIMAL,
                    FOREIGN KEY (FinancialDataID) REFERENCES FinancialData(FinancialDataID)
                );"
            cmd.CommandText = createRatiosTable
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    'It adds a new financial data record to the database
    Public Sub AddFinancialData(dateValue As Date, revenue As Decimal, costOfGoodsSold As Decimal, operatingExpenses As Decimal, netIncome As Decimal, totalAssets As Decimal, totalEquity As Decimal, ebitda As Decimal, currentAssets As Decimal, currentLiabilities As Decimal, totalLiabilities As Decimal, interestExpense As Decimal, variableCosts As Decimal, fixedCosts As Decimal, salesRevenuePerUnit As Decimal, variableCostPerUnit As Decimal)
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            Dim dateId As Integer = GetOrCreateDateID(dateValue)

            Dim cmd As New SQLiteCommand("
                INSERT INTO FinancialData (DateID, Revenue, CostOfGoodsSold, OperatingExpenses, NetIncome, TotalAssets, TotalEquity, EBITDA, CurrentAssets, CurrentLiabilities, TotalLiabilities, InterestExpense, VariableCosts, FixedCosts, SalesRevenuePerUnit, VariableCostPerUnit)
                VALUES (@DateID, @Revenue, @CostOfGoodsSold, @OperatingExpenses, @NetIncome, @TotalAssets, @TotalEquity, @EBITDA, @CurrentAssets, @CurrentLiabilities, @TotalLiabilities, @InterestExpense, @VariableCosts, @FixedCosts, @SalesRevenuePerUnit, @VariableCostPerUnit)", connection)

            cmd.Parameters.AddWithValue("@DateID", dateId)
            cmd.Parameters.AddWithValue("@Revenue", revenue)
            cmd.Parameters.AddWithValue("@CostOfGoodsSold", costOfGoodsSold)
            cmd.Parameters.AddWithValue("@OperatingExpenses", operatingExpenses)
            cmd.Parameters.AddWithValue("@NetIncome", netIncome)
            cmd.Parameters.AddWithValue("@TotalAssets", totalAssets)
            cmd.Parameters.AddWithValue("@TotalEquity", totalEquity)
            cmd.Parameters.AddWithValue("@EBITDA", ebitda)
            cmd.Parameters.AddWithValue("@CurrentAssets", currentAssets)
            cmd.Parameters.AddWithValue("@CurrentLiabilities", currentLiabilities)
            cmd.Parameters.AddWithValue("@TotalLiabilities", totalLiabilities)
            cmd.Parameters.AddWithValue("@InterestExpense", interestExpense)
            cmd.Parameters.AddWithValue("@VariableCosts", variableCosts)
            cmd.Parameters.AddWithValue("@FixedCosts", fixedCosts)
            cmd.Parameters.AddWithValue("@SalesRevenuePerUnit", salesRevenuePerUnit)
            cmd.Parameters.AddWithValue("@VariableCostPerUnit", variableCostPerUnit)

            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub UpdateFinancialData(financialDataId As Integer, dateValue As Date, revenue As Decimal, costOfGoodsSold As Decimal, operatingExpenses As Decimal, netIncome As Decimal, totalAssets As Decimal, totalEquity As Decimal, ebitda As Decimal, currentAssets As Decimal, currentLiabilities As Decimal, totalLiabilities As Decimal, interestExpense As Decimal, variableCosts As Decimal, fixedCosts As Decimal, salesRevenuePerUnit As Decimal, variableCostPerUnit As Decimal)
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            Dim dateId As Integer = GetOrCreateDateID(dateValue)

            Dim cmd As New SQLiteCommand("
                UPDATE FinancialData
                SET DateID = @DateID, Revenue = @Revenue, CostOfGoodsSold = @CostOfGoodsSold, OperatingExpenses = @OperatingExpenses, NetIncome = @NetIncome, TotalAssets = @TotalAssets, TotalEquity = @TotalEquity, EBITDA = @EBITDA, CurrentAssets = @CurrentAssets, CurrentLiabilities = @CurrentLiabilities, TotalLiabilities = @TotalLiabilities, InterestExpense = @InterestExpense, VariableCosts = @VariableCosts, FixedCosts = @FixedCosts, SalesRevenuePerUnit = @SalesRevenuePerUnit, VariableCostPerUnit = @VariableCostPerUnit
                WHERE FinancialDataID = @FinancialDataID", connection)

            cmd.Parameters.AddWithValue("@DateID", dateId)
            cmd.Parameters.AddWithValue("@Revenue", revenue)
            cmd.Parameters.AddWithValue("@CostOfGoodsSold", costOfGoodsSold)
            cmd.Parameters.AddWithValue("@OperatingExpenses", operatingExpenses)
            cmd.Parameters.AddWithValue("@NetIncome", netIncome)
            cmd.Parameters.AddWithValue("@TotalAssets", totalAssets)
            cmd.Parameters.AddWithValue("@TotalEquity", totalEquity)
            cmd.Parameters.AddWithValue("@EBITDA", ebitda)
            cmd.Parameters.AddWithValue("@CurrentAssets", currentAssets)
            cmd.Parameters.AddWithValue("@CurrentLiabilities", currentLiabilities)
            cmd.Parameters.AddWithValue("@TotalLiabilities", totalLiabilities)
            cmd.Parameters.AddWithValue("@InterestExpense", interestExpense)
            cmd.Parameters.AddWithValue("@VariableCosts", variableCosts)
            cmd.Parameters.AddWithValue("@FixedCosts", fixedCosts)
            cmd.Parameters.AddWithValue("@SalesRevenuePerUnit", salesRevenuePerUnit)
            cmd.Parameters.AddWithValue("@VariableCostPerUnit", variableCostPerUnit)
            cmd.Parameters.AddWithValue("@FinancialDataID", financialDataId)

            cmd.ExecuteNonQuery()
        End Using
    End Sub

    'It deletes a financial data record from the database
    Public Async Function DeleteFinancialDataAsync(financialDataId As Integer) As Task
        Using connection As New SQLiteConnection(connectionString)
            Await connection.OpenAsync()

            Dim cmd As New SQLiteCommand("
                DELETE FROM FinancialData WHERE FinancialDataID = @FinancialDataID", connection)

            cmd.Parameters.AddWithValue("@FinancialDataID", financialDataId)
            Await cmd.ExecuteNonQueryAsync()
        End Using
    End Function

    Public Function GetOrCreateDateID(dateValue As Date) As Integer
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            ' Check if the date already exists in the DateDimension table
            Dim cmd As New SQLiteCommand("
                SELECT DateID FROM DateDimension WHERE Date = @Date", connection)
            cmd.Parameters.AddWithValue("@Date", dateValue)
            Dim result As Object = cmd.ExecuteScalar()

            ' If the date exists, return the DateID
            If result IsNot Nothing Then
                Return Convert.ToInt32(result)
            Else
                ' If the date does not exist, insert a new date and return the new DateID
                Dim year As Integer = dateValue.Year
                Dim quarter As Integer = ((dateValue.Month - 1) \ 3) + 1
                Dim month As Integer = dateValue.Month
                Dim day As Integer = dateValue.Day

                Dim insertCmd As New SQLiteCommand("
                    INSERT INTO DateDimension (Date, Year, Quarter, Month, Day)
                    VALUES (@Date, @Year, @Quarter, @Month, @Day);
                    SELECT last_insert_rowid();", connection)
                insertCmd.Parameters.AddWithValue("@Date", dateValue)
                insertCmd.Parameters.AddWithValue("@Year", year)
                insertCmd.Parameters.AddWithValue("@Quarter", quarter)
                insertCmd.Parameters.AddWithValue("@Month", month)
                insertCmd.Parameters.AddWithValue("@Day", day)

                Return Convert.ToInt32(insertCmd.ExecuteScalar())
            End If
        End Using
    End Function

    'It fetches all financial data records from the database
    Public Async Function GetFinancialDataAsync() As Task(Of List(Of FinancialData))
        Dim financialDataList As New List(Of FinancialData)

        Using connection As New SQLiteConnection(connectionString)
            Await connection.OpenAsync()

            Dim cmd As New SQLiteCommand("
                SELECT FinancialDataID, DateID, Revenue, CostOfGoodsSold, OperatingExpenses, NetIncome, TotalAssets, TotalEquity, EBITDA, CurrentAssets, CurrentLiabilities, TotalLiabilities, InterestExpense, VariableCosts, FixedCosts, SalesRevenuePerUnit, VariableCostPerUnit 
                FROM FinancialData", connection)

            Using reader As SQLiteDataReader = Await cmd.ExecuteReaderAsync()
                While Await reader.ReadAsync()
                    Dim dateId As Integer = Convert.ToInt32(reader("DateID"))

                    ' Fetch date details from DateDimension
                    Dim dateCmd As New SQLiteCommand("
                        SELECT Date FROM DateDimension WHERE DateID = @DateID", connection)
                    dateCmd.Parameters.AddWithValue("@DateID", dateId)
                    Dim dateValue As Date = Convert.ToDateTime(Await dateCmd.ExecuteScalarAsync())

                    Dim data As New FinancialData With {
                        .FinancialDataID = Convert.ToInt32(reader("FinancialDataID")),
                        .DateValue = dateValue,
                        .Revenue = Convert.ToDecimal(reader("Revenue")),
                        .CostOfGoodsSold = Convert.ToDecimal(reader("CostOfGoodsSold")),
                        .OperatingExpenses = Convert.ToDecimal(reader("OperatingExpenses")),
                        .NetIncome = Convert.ToDecimal(reader("NetIncome")),
                        .TotalAssets = Convert.ToDecimal(reader("TotalAssets")),
                        .TotalEquity = Convert.ToDecimal(reader("TotalEquity")),
                        .EBITDA = Convert.ToDecimal(reader("EBITDA")),
                        .CurrentAssets = Convert.ToDecimal(reader("CurrentAssets")),
                        .CurrentLiabilities = Convert.ToDecimal(reader("CurrentLiabilities")),
                        .TotalLiabilities = Convert.ToDecimal(reader("TotalLiabilities")),
                        .InterestExpense = Convert.ToDecimal(reader("InterestExpense")),
                        .VariableCosts = Convert.ToDecimal(reader("VariableCosts")),
                        .FixedCosts = Convert.ToDecimal(reader("FixedCosts")),
                        .SalesRevenuePerUnit = Convert.ToDecimal(reader("SalesRevenuePerUnit")),
                        .VariableCostPerUnit = Convert.ToDecimal(reader("VariableCostPerUnit"))
                    }
                    financialDataList.Add(data)
                End While
            End Using
        End Using

        Return financialDataList
    End Function


    'TEMP FUNCTION, JUST TESTING A FUNCTIONALITY
    Public Sub PrintDateComponents(dateId As Integer)
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            ' Query to get the date components for the given DateID
            Dim cmd As New SQLiteCommand("
            SELECT Date, Year, Quarter, Month, Day 
            FROM DateDimension 
            WHERE DateID = @DateID", connection)
            cmd.Parameters.AddWithValue("@DateID", dateId)

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read() Then
                    Dim dateValue As Date = Convert.ToDateTime(reader("Date"))
                    Dim year As Integer = Convert.ToInt32(reader("Year"))
                    Dim quarter As Integer = Convert.ToInt32(reader("Quarter"))
                    Dim month As Integer = Convert.ToInt32(reader("Month"))
                    Dim day As Integer = Convert.ToInt32(reader("Day"))

                    ' Print the components
                    Debug.WriteLine("DateID: " & dateId)
                    Debug.WriteLine("Date: " & dateValue.ToShortDateString())
                    Debug.WriteLine("Year: " & year)
                    Debug.WriteLine("Quarter: " & quarter)
                    Debug.WriteLine("Month: " & month)
                    Debug.WriteLine("Day: " & day)
                Else
                    Debug.WriteLine("DateID not found.")
                End If
            End Using
        End Using
    End Sub

End Class