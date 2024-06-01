﻿Imports System.Data
Imports System.IO
Imports System.Windows.Threading
Imports Microsoft.Win32
Imports OfficeOpenXml
Imports System.Net
Imports Newtonsoft.Json.Linq
Imports System.Net.Http
'Delete this when done with the project
Imports System.Runtime.InteropServices
Class MainWindow

    'Delete this when done with the project
    <DllImport("kernel32.dll")>
    Public Shared Function AllocConsole() As Boolean
    End Function

    Private dbHelper As Database
    Private apiKey As String
    Private FinancialData As FinancialData
    Private financialDataList As List(Of FinancialData)
    Public Sub New()
        'Delete this when done with the project
        AllocConsole()

        apiKey = "your Alpha Vantage API Key"
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        InitializeComponent()
        dbHelper = New Database("dbtest.db")
        FinancialData = New FinancialData()
    End Sub
    ' Event handlers to switch tabs
    Private Sub DashboardButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = DashboardTab
    End Sub

    Private Async Sub PreviewDatabaseButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = PreviewDatabaseTab
        Await LoadFinancialDataAsync()
    End Sub

    Private Async Sub DeleteDataButton_Click(sender As Object, e As RoutedEventArgs)
        ' Validate input
        Dim financialDataId As Integer
        If Not Integer.TryParse(FinancialDataIDInput.Text, financialDataId) Then
            MessageBox.Show("Invalid input for FinancialDataID. Please enter a valid integer.")
            Return
        End If

        ' Perform the delete operation asynchronously
        Try
            Debug.WriteLine("Starting delete operation...")
            Await dbHelper.DeleteFinancialDataAsync(financialDataId)
            Debug.WriteLine("Delete operation completed.")
            MessageBox.Show("Financial data deleted successfully!")

            ' Refresh the DataGrid
            Await LoadFinancialDataAsync()
        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}")
            Debug.WriteLine($"An error occurred: {ex.Message}")
        End Try

        ' Clear the input field
        FinancialDataIDInput.Text = String.Empty
    End Sub


    Private Async Sub AnalysisButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = AnalysisTab
        Await LoadFinancialDataAsync()
        Console.WriteLine(FinancialData.TotalAssets)
    End Sub

    Private Sub InputDataButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = InputDataTab
    End Sub

    Private Sub AddInputData_Click(sender As Object, e As RoutedEventArgs)
        ' Validate date input
        If Not DateInput.SelectedDate.HasValue Then
            MessageBox.Show("Please select a date.")
            Return
        End If

        Dim dateValue As Date = DateInput.SelectedDate.Value
        Dim revenue As Decimal
        Dim costOfGoodsSold As Decimal
        Dim operatingExpenses As Decimal
        Dim netIncome As Decimal
        Dim totalAssets As Decimal
        Dim totalEquity As Decimal
        Dim ebitda As Decimal
        Dim currentAssets As Decimal
        Dim currentLiabilities As Decimal
        Dim totalLiabilities As Decimal
        Dim interestExpense As Decimal
        Dim variableCosts As Decimal
        Dim fixedCosts As Decimal
        Dim salesRevenuePerUnit As Decimal
        Dim variableCostPerUnit As Decimal

        ' Validate and parse input values
        If Not Decimal.TryParse(RevenueInput.Text, revenue) Then
            MessageBox.Show("Invalid input for Revenue.")
            Return
        End If
        If Not Decimal.TryParse(CostOfGoodsSoldInput.Text, costOfGoodsSold) Then
            MessageBox.Show("Invalid input for Cost of Goods Sold.")
            Return
        End If
        If Not Decimal.TryParse(OperatingExpensesInput.Text, operatingExpenses) Then
            MessageBox.Show("Invalid input for Operating Expenses.")
            Return
        End If
        If Not Decimal.TryParse(NetIncomeInput.Text, netIncome) Then
            MessageBox.Show("Invalid input for Net Income.")
            Return
        End If
        If Not Decimal.TryParse(TotalAssetsInput.Text, totalAssets) Then
            MessageBox.Show("Invalid input for Total Assets.")
            Return
        End If
        If Not Decimal.TryParse(TotalEquityInput.Text, totalEquity) Then
            MessageBox.Show("Invalid input for Shareholders' Equity.")
            Return
        End If
        If Not Decimal.TryParse(EBITDAInput.Text, ebitda) Then
            MessageBox.Show("Invalid input for EBITDA.")
            Return
        End If
        If Not Decimal.TryParse(CurrentAssetsInput.Text, currentAssets) Then
            MessageBox.Show("Invalid input for Current Assets.")
            Return
        End If
        If Not Decimal.TryParse(CurrentLiabilitiesInput.Text, currentLiabilities) Then
            MessageBox.Show("Invalid input for Current Liabilities.")
            Return
        End If
        If Not Decimal.TryParse(TotalLiabilitiesInput.Text, totalLiabilities) Then
            MessageBox.Show("Invalid input for Total Liabilities.")
            Return
        End If
        If Not Decimal.TryParse(InterestExpenseInput.Text, interestExpense) Then
            MessageBox.Show("Invalid input for Interest Expense.")
            Return
        End If
        If Not Decimal.TryParse(VariableCostsInput.Text, variableCosts) Then
            MessageBox.Show("Invalid input for Variable Costs.")
            Return
        End If
        If Not Decimal.TryParse(FixedCostsInput.Text, fixedCosts) Then
            MessageBox.Show("Invalid input for Fixed Costs.")
            Return
        End If
        If Not Decimal.TryParse(SalesRevenuePerUnitInput.Text, salesRevenuePerUnit) Then
            MessageBox.Show("Invalid input for Sales Revenue Per Unit.")
            Return
        End If
        If Not Decimal.TryParse(VariableCostPerUnitInput.Text, variableCostPerUnit) Then
            MessageBox.Show("Invalid input for Variable Cost Per Unit.")
            Return
        End If

        ' Add the financial data to the database
        dbHelper.AddFinancialData(dateValue, revenue, costOfGoodsSold, operatingExpenses, netIncome, totalAssets, totalEquity, ebitda, currentAssets, currentLiabilities, totalLiabilities, interestExpense, variableCosts, fixedCosts, salesRevenuePerUnit, variableCostPerUnit)

        MessageBox.Show("Financial data added successfully!")
        ClearInputFields()
    End Sub

    Private Sub ClearInputButton_Click(sender As Object, e As RoutedEventArgs)
        ClearInputFields()
    End Sub
    Private Sub ClearInputFields()
        DateInput.SelectedDate = Nothing
        RevenueInput.Text = String.Empty
        CostOfGoodsSoldInput.Text = String.Empty
        OperatingExpensesInput.Text = String.Empty
        NetIncomeInput.Text = String.Empty
        TotalAssetsInput.Text = String.Empty
        TotalEquityInput.Text = String.Empty
        EBITDAInput.Text = String.Empty
        CurrentAssetsInput.Text = String.Empty
        CurrentLiabilitiesInput.Text = String.Empty
        TotalLiabilitiesInput.Text = String.Empty
        InterestExpenseInput.Text = String.Empty
        VariableCostsInput.Text = String.Empty
        FixedCostsInput.Text = String.Empty
        SalesRevenuePerUnitInput.Text = String.Empty
        VariableCostPerUnitInput.Text = String.Empty
    End Sub

    Private Sub ScenarioAnalysisButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = ScenarioAnalysisTab
    End Sub

    Private Async Function LoadFinancialDataAsync() As Task
        Try
            Debug.WriteLine("Loading financial data...")
            Dim financialDataList As List(Of FinancialData) = Await dbHelper.GetFinancialDataAsync()
            Dispatcher.Invoke(Sub()
                                  FinancialDataGrid.ItemsSource = financialDataList
                                  Debug.WriteLine("Financial data loaded.")
                              End Sub)
        Catch ex As Exception
            MessageBox.Show($"An error occurred while loading data: {ex.Message}")
            Debug.WriteLine($"An error occurred while loading data: {ex.Message}")
        End Try
    End Function

    Private Sub ImportExcelButton_Click(sender As Object, e As RoutedEventArgs)
        ' Open a file dialog to select an Excel file
        Dim openFileDialog As New OpenFileDialog() With {
            .Filter = "Excel Files|*.xls;*.xlsx"
        }

        If openFileDialog.ShowDialog() = True Then
            ImportExcelData(openFileDialog.FileName)
        End If
    End Sub

    Private Sub ImportExcelData(filePath As String)
        Try
            Using package As New ExcelPackage(New FileInfo(filePath))
                Dim worksheet = package.Workbook.Worksheets.FirstOrDefault()
                If worksheet Is Nothing Then
                    MessageBox.Show("No worksheet found in the Excel file.")
                    Return
                End If

                ' Create a DataTable to hold the data
                Dim dataTable As New DataTable()

                ' Add columns to the DataTable
                For Each firstRowCell In worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
                    dataTable.Columns.Add(firstRowCell.Text)
                Next

                ' Add rows to the DataTable
                For rowNum = 2 To worksheet.Dimension.End.Row
                    Dim wsRow = worksheet.Cells(rowNum, 1, rowNum, worksheet.Dimension.End.Column)
                    Dim row = dataTable.NewRow()
                    For Each cell In wsRow
                        row(cell.Start.Column - 1) = cell.Text
                    Next
                    dataTable.Rows.Add(row)
                Next

                ' Bind the DataTable to the DataGrid
                ImportedDataGrid.ItemsSource = dataTable.DefaultView
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error reading Excel file: {ex.Message}")
        End Try
    End Sub

    Private Sub AddImportedDataButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dataTable As DataTable = CType(ImportedDataGrid.ItemsSource, DataView).Table

        For rowIndex As Integer = 1 To dataTable.Rows.Count - 1 ' Start from the second row (index 1)
            Try
                Dim row As DataRow = dataTable.Rows(rowIndex)
                Dim dateValue As Date = Date.Parse(row(0).ToString())
                Dim revenue As Decimal = Decimal.Parse(row(1).ToString())
                Dim costOfGoodsSold As Decimal = Decimal.Parse(row(2).ToString())
                Dim operatingExpenses As Decimal = Decimal.Parse(row(3).ToString())
                Dim netIncome As Decimal = Decimal.Parse(row(4).ToString())
                Dim totalAssets As Decimal = Decimal.Parse(row(5).ToString())
                Dim totalEquity As Decimal = Decimal.Parse(row(6).ToString())
                Dim ebitda As Decimal = Decimal.Parse(row(7).ToString())
                Dim currentAssets As Decimal = Decimal.Parse(row(8).ToString())
                Dim currentLiabilities As Decimal = Decimal.Parse(row(9).ToString())
                Dim totalLiabilities As Decimal = Decimal.Parse(row(10).ToString())
                Dim interestExpense As Decimal = Decimal.Parse(row(11).ToString())
                Dim variableCosts As Decimal = Decimal.Parse(row(12).ToString())
                Dim fixedCosts As Decimal = Decimal.Parse(row(13).ToString())
                Dim salesRevenuePerUnit As Decimal = Decimal.Parse(row(14).ToString())
                Dim variableCostPerUnit As Decimal = Decimal.Parse(row(15).ToString())

                ' Create a summary of the data
                Dim dataSummary As String = $"Date: {dateValue}" & Environment.NewLine &
                                            $"Revenue: {revenue}" & Environment.NewLine &
                                            $"Cost of Goods Sold: {costOfGoodsSold}" & Environment.NewLine &
                                            $"Operating Expenses: {operatingExpenses}" & Environment.NewLine &
                                            $"Net Income: {netIncome}" & Environment.NewLine &
                                            $"Total Assets: {totalAssets}" & Environment.NewLine &
                                            $"Shareholders' Equity: {totalEquity}" & Environment.NewLine &
                                            $"EBITDA: {ebitda}" & Environment.NewLine &
                                            $"Current Assets: {currentAssets}" & Environment.NewLine &
                                            $"Current Liabilities: {currentLiabilities}" & Environment.NewLine &
                                            $"Total Liabilities: {totalLiabilities}" & Environment.NewLine &
                                            $"Interest Expense: {interestExpense}" & Environment.NewLine &
                                            $"Variable Costs: {variableCosts}" & Environment.NewLine &
                                            $"Fixed Costs: {fixedCosts}" & Environment.NewLine &
                                            $"Sales Revenue Per Unit: {salesRevenuePerUnit}" & Environment.NewLine &
                                            $"Variable Cost Per Unit: {variableCostPerUnit}"

                ' Show the confirmation message box
                Dim result As MessageBoxResult = MessageBox.Show($"Please confirm the following data:{Environment.NewLine}{Environment.NewLine}{dataSummary}", "Confirm Data Entry", MessageBoxButton.OKCancel, MessageBoxImage.Information)

                ' If the user clicks OK, add the data to the database
                If result = MessageBoxResult.OK Then
                    dbHelper.AddFinancialData(dateValue, revenue, costOfGoodsSold, operatingExpenses, netIncome, totalAssets, totalEquity, ebitda, currentAssets, currentLiabilities, totalLiabilities, interestExpense, variableCosts, fixedCosts, salesRevenuePerUnit, variableCostPerUnit)
                    MessageBox.Show("Financial data added successfully!")
                    ClearInputFields()
                End If
            Catch ex As Exception
                MessageBox.Show($"Error parsing data at row {rowIndex + 1}: {ex.Message}")
            End Try
        Next
    End Sub

    Private Sub MainTabControl_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles MainTabControl.SelectionChanged

    End Sub

    Private Async Sub GetAPIdata_Click(sender As Object, e As RoutedEventArgs)
        Dim comboBoxItem1 As ComboBoxItem = CType(SymbolOptions.SelectedItem, ComboBoxItem)
        Dim selectedSymbol As String = comboBoxItem1.Content.ToString()
        Dim selectedFiscalYearIndex As Integer = PeriodOptions.SelectedIndex

        Console.WriteLine("This is the selected symbol: " & selectedSymbol)
        Console.WriteLine("This is the selected fiscal year index: " & selectedFiscalYearIndex)

        Dim financialData As New FinancialData()
        Await financialData.LoadFinancialData(selectedSymbol, selectedFiscalYearIndex)
        financialData.PrintFinancialData()
        ' Additional code to display ratios or handle results...
    End Sub

    'TO DO: CHECK WHY IT IS NOT WORKING (Values obtained are 0)
    Private Async Sub CalculateRatiosButton_Click(sender As Object, e As RoutedEventArgs)
        If financialDataList Is Nothing OrElse financialDataList.Count = 0 Then
            Console.WriteLine(financialDataList.Count)
            Console.WriteLine("Data empty, why?")
            Await LoadFinancialDataAsync()
        End If
        Try
            If CK_ROA.IsChecked = True Then
                Dim roa As Decimal = FinancialData.ReturnOnAssets()
                Console.WriteLine($"Return on Assets: {roa}")
            End If
        Catch ex As Exception
            Console.WriteLine(FinancialData.TotalAssets)
            Console.WriteLine($"An error occurred while calculating ratios: {ex.Message}")
        End Try

    End Sub

End Class