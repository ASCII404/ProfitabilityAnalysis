Imports System.Data
Imports System.IO
Imports System.Reflection.Emit
Imports System.Windows.Threading
Imports LiveCharts
Imports LiveCharts.Wpf
Imports Microsoft.Win32
Imports OfficeOpenXml
Imports System.Collections.Generic
Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports LiveCharts.Defaults



Class MainWindow

    Private dbHelper As Database
    Private apiKey As String
    Private FinancialData As FinancialData
    Private financialDataList As List(Of FinancialData)
    Private previous_selected_tab As String
    Private helpButton_content As String
    Private authenatication_win As Authentication

    Public Property Chart1Values As ChartValues(Of ObservablePoint)
    Public Property Chart2Values As ChartValues(Of ObservablePoint)
    Public Property Chart3Values As ChartValues(Of ObservablePoint)
    Public Property Chart4Values As ChartValues(Of ObservableValue)
    Public Sub New()

        apiKey = "your Alpha Vantage API Key"
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        InitializeComponent()
        dbHelper = New Database("dbtest.db")
        FinancialData = New FinancialData()



        InitializeChartData1()
        InitializeChartData2()
        InitializeChartData3()
        InitializeChartData4()

    End Sub

    Private Sub InitializeChartData1()
        ' Example: Populate data for Chart 1 (Net Income)
        Chart1Values = New ChartValues(Of ObservablePoint) From {
    New ObservablePoint(1, 1000),
    New ObservablePoint(2, 2200),
    New ObservablePoint(3, 1500),
    New ObservablePoint(4, 3000),
    New ObservablePoint(5, 2800),
    New ObservablePoint(6, 3500),
    New ObservablePoint(7, 2500),
    New ObservablePoint(8, 4000),
    New ObservablePoint(9, 3200),
    New ObservablePoint(10, 4500)}
        ' Add more points as needed

        ' Bind data to Chart 1
        DataContext = Me

        ' Example: Placeholder text update for Chart 1
        placeholder1.Visibility = Visibility.Collapsed
        chart1.Visibility = Visibility.Visible
    End Sub

    Private Sub InitializeChartData2()
        ' Example: Populate data for Chart 2 (Total Assets)
        Chart2Values = New ChartValues(Of ObservablePoint) From {
    New ObservablePoint(1, 1200),
    New ObservablePoint(2, 1800),
    New ObservablePoint(3, 2500),
    New ObservablePoint(4, 2200),
    New ObservablePoint(5, 3800),
    New ObservablePoint(6, 3000),
    New ObservablePoint(7, 2800),
    New ObservablePoint(8, 3500),
    New ObservablePoint(9, 4000),
    New ObservablePoint(10, 5000)}
        ' Bind data to Chart 2
        chart2.Series = New SeriesCollection From {
    New ColumnSeries With {
        .Values = Chart2Values,
        .Title = "Total Assets",
        .Fill = Brushes.CadetBlue,
        .Stroke = Brushes.CadetBlue
    }
}
        DataContext = Me

        ' Placeholder text update for Chart 2
        placeholder2.Visibility = Visibility.Collapsed
        chart2.Visibility = Visibility.Visible
    End Sub

    Private Sub InitializeChartData3()
        ' Example: Populate data for Chart 3 (COGS)
        Chart3Values = New ChartValues(Of ObservablePoint) From {
        New ObservablePoint(1, 1400),
        New ObservablePoint(2, 2500),
        New ObservablePoint(3, 1800),
        New ObservablePoint(4, 2800),
        New ObservablePoint(5, 3500),
        New ObservablePoint(6, 2200),
        New ObservablePoint(7, 4000),
        New ObservablePoint(8, 3800),
        New ObservablePoint(9, 3000),
        New ObservablePoint(10, 4500)
    }

        ' Bind data to Chart 3 as a Row Chart
        chart3.Series = New SeriesCollection From {
        New RowSeries With {
            .Values = Chart3Values,
            .Title = "COGS",
            .Fill = Brushes.CadetBlue,
            .Stroke = Brushes.CadetBlue
        }
    }

        ' Set X and Y axes
        chart3.AxisX.Clear()
        chart3.AxisX.Add(New Axis With {
        .Title = "Amount",
        .Foreground = Brushes.Black
    })

        chart3.AxisY.Clear()
        chart3.AxisY.Add(New Axis With {
        .Title = "Categories",
        .Foreground = Brushes.Black
    })

        DataContext = Me

        ' Placeholder text update for Chart 3
        placeholder3.Visibility = Visibility.Collapsed
        chart3.Visibility = Visibility.Visible
    End Sub

    Public Sub InitializeChartData4()
        ' Ensure chart4 is initialized and not null here
        If chart4 Is Nothing Then
            Throw New NullReferenceException("chart4 is not initialized.")
        End If

        ' Sample data for the chart
        Dim revenueValues As New ChartValues(Of Double) From {2000, 2200, 2400, 2600, 2800, 3000, 3200, 3400, 3600, 3800}
        Dim expenseValues As New ChartValues(Of Double) From {1800, 1900, 2100, 2300, 2500, 2700, 2900, 3100, 3300, 3500}
        Dim dateLabels As New List(Of String) From {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"}

        ' Bind data to the chart
        chart4.Series = New SeriesCollection From {
        New LineSeries With {
            .Title = "Revenue",
            .Values = revenueValues,
            .LineSmoothness = 0.5,
            .PointGeometry = Nothing,
            .StrokeThickness = 2,
            .Fill = Brushes.Transparent,
            .Stroke = Brushes.CadetBlue
        },
        New LineSeries With {
            .Title = "Expenses",
            .Values = expenseValues,
            .LineSmoothness = 0.5,
            .PointGeometry = Nothing,
            .StrokeThickness = 2,
            .Fill = Brushes.Transparent,
            .Stroke = Brushes.Red
        }
    }

        ' Bind the date labels to the X-axis
        chart4.AxisX.Clear()
        chart4.AxisX.Add(New Axis With {
        .Title = "Date",
        .Labels = dateLabels,
        .Foreground = Brushes.Black
    })

        ' Set the Y-axis title
        chart4.AxisY.Clear()
        chart4.AxisY.Add(New Axis With {
        .Title = "Amount",
        .Foreground = Brushes.Black
    })

        ' Update visibility if needed
        If placeholder4 IsNot Nothing Then
            placeholder4.Visibility = Visibility.Collapsed
        End If

        chart4.Visibility = Visibility.Visible
    End Sub


    Private Sub DashboardButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = DashboardTab
    End Sub

    Private Async Sub PreviewDatabaseButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = PreviewDatabaseTab
        Await LoadFinancialDataAsync()
        dbHelper.PrintDateComponents(300)
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

    Private Sub ScenarioAnalysisButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = ScenarioAnalysisTab
    End Sub

    Private Sub ImportExcelButton_Click(sender As Object, e As RoutedEventArgs)
        ' Open a file dialog to select an Excel file
        Dim openFileDialog As New OpenFileDialog() With {
            .Filter = "Excel Files|*.xls;*.xlsx"
        }

        If openFileDialog.ShowDialog() = True Then
            ImportExcelData(openFileDialog.FileName)
        End If
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

    Private Sub AddImportedDataButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dataTable As DataTable = CType(ImportedDataGrid.ItemsSource, DataView).Table
        Debug.WriteLine("this is the no. of rows of import:" & dataTable.Rows.Count)
        Dim rowIndex As Integer = 1 ' Start from the second row (index 1)

        Do While rowIndex < dataTable.Rows.Count
            Try
                Dim row As DataRow = dataTable.Rows(rowIndex)

                ' Check if the row is empty by examining the first column
                If String.IsNullOrWhiteSpace(row(0).ToString()) Then
                    Exit Do
                End If

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

                '' Show the confirmation message box
                'Dim result As MessageBoxResult = MessageBox.Show($"Please confirm the following data:{Environment.NewLine}{Environment.NewLine}{dataSummary}", "Confirm Data Entry", MessageBoxButton.OKCancel, MessageBoxImage.Information)

                '' If the user clicks OK, add the data to the database
                'If result = MessageBoxResult.OK Then
                dbHelper.AddFinancialData(dateValue, revenue, costOfGoodsSold, operatingExpenses, netIncome, totalAssets, totalEquity, ebitda, currentAssets, currentLiabilities, totalLiabilities, interestExpense, variableCosts, fixedCosts, salesRevenuePerUnit, variableCostPerUnit)
                '    ClearInputFields()
                'End If
            Catch ex As Exception
                MessageBox.Show($"Error parsing data at row {rowIndex + 1}: {ex.Message}")
            End Try

            rowIndex += 1
        Loop

        MessageBox.Show("Financial data added successfully!")
    End Sub

    Private Async Sub GetAPIdata_Click(sender As Object, e As RoutedEventArgs)
        Dim comboBoxItem1 As ComboBoxItem = CType(SymbolOptions.SelectedItem, ComboBoxItem)
        Dim selectedSymbol As String = comboBoxItem1.Content.ToString()
        Dim selectedFiscalYearIndex As Integer = PeriodOptions.SelectedIndex

        Debug.WriteLine("This is the selected symbol: " & selectedSymbol)
        Debug.WriteLine("This is the selected fiscal year index: " & selectedFiscalYearIndex)

        Dim financialData As New FinancialData()
        Await financialData.LoadFinancialData(selectedSymbol, selectedFiscalYearIndex)
        Debug.WriteLine(financialData.Revenue)
        incomeS_Rev.Text = financialData.Revenue.ToString("C0")
        incomeS_COGS.Text = financialData.CostOfGoodsSold.ToString("C0")
        incomeS_OP_Expenses.Text = financialData.OperatingExpenses.ToString("C0")
        incomeS_NetIncome.Text = financialData.NetIncome.ToString("C0")
        incomeS_EBITDA.Text = financialData.EBITDA.ToString("C0")
        incomeS_IExpenses.Text = financialData.InterestExpense.ToString("C0")
    End Sub

    Private Async Sub GetAPIdata2_Click(sender As Object, e As RoutedEventArgs)
        Dim comboBoxItem2 As ComboBoxItem = CType(SymbolOptions.SelectedItem, ComboBoxItem)
        Dim selectedSymbol As String = comboBoxItem2.Content.ToString()
        Dim selectedFiscalYearIndex As Integer = PeriodOptions.SelectedIndex

        Debug.WriteLine("This is the selected symbol: " & selectedSymbol)
        Debug.WriteLine("This is the selected fiscal year index: " & selectedFiscalYearIndex)

        Dim financialData As New FinancialData()
        Await financialData.LoadFinancialData2(selectedSymbol, selectedFiscalYearIndex)
        Debug.WriteLine(financialData.TotalAssets)
        balanceS_TA.Text = financialData.TotalAssets.ToString("C0")
        balanceS_TE.Text = financialData.TotalEquity.ToString("C0")
        balanceS_CA.Text = financialData.CurrentAssets.ToString("C0")
        balanceS_CL.Text = financialData.CurrentLiabilities.ToString("C0")
        balanceS_TL.Text = financialData.TotalLiabilities.ToString("C0")

    End Sub
    'Help button content
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        MessageBox.Show(helpButton_content)
    End Sub

    'Function to import excel data from a file
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

    'It is used to check when a tab is selected and change the color of the buttons in the nav. bar and help button content
    Private Sub MainTabControl_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles MainTabControl.SelectionChanged
        If TypeOf e.Source Is TabControl Then
            Dim selectedTab As TabItem = CType(MainTabControl.SelectedItem, TabItem)
            Dim selectedTabName As String = selectedTab.Name

            'HELP BUTTON CONTENT
            ' Only show the message if the selected tab has changed
            If selectedTabName <> previous_selected_tab Then
                If selectedTabName = "DashboardTab" Then
                    helpButton_content = "Welcome to the Financial Data Visualization Tool!" & vbCrLf &
                                         "This is your main dashboard for visualizing financial data from our database." & vbCrLf &
                                         "If you see no data, means the database is empty." & vbCrLf &
                                         "Simply switch to the 'Input Data' tab to enter or update your financial information." & vbCrLf &
                                         "Make sure to input valid data for accurate visualization." & vbCrLf &
                                         "Once the data is entered, return to the dashboard tab to see your visualizations." & vbCrLf
                ElseIf selectedTabName = "PreviewDatabaseTab" Then
                    helpButton_content = "In this tab you can see your financial data that you entered." & vbCrLf &
                                         "You can also delete any data by entering the number of a line and by clicking the 'Delete' button."
                ElseIf selectedTabName = "AnalysisTab" Then
                    helpButton_content = "This is the content from the AnalysisTab"
                ElseIf selectedTabName = "InputDataTab" Then
                    helpButton_content = "This is the content from the InputDataTab"
                ElseIf selectedTabName = "ScenarioAnalysisTab" Then
                    helpButton_content = "This is the content from the ScenarioAnalysisTab"
                End If
                Debug.WriteLine("Current Tab: " & selectedTabName)
                previous_selected_tab = selectedTabName ' Update the previous selected tab name
            End If

            'Reset the background color of all buttons from navigation bar
            dashboard_button.Background = Brushes.MintCream
            database_button.Background = Brushes.MintCream
            analysis_button.Background = Brushes.MintCream
            inputData_button.Background = Brushes.MintCream
            results_button.Background = Brushes.MintCream

            ' Set the background color of the selected tab's button
            Select Case selectedTabName
                Case "DashboardTab"
                    dashboard_button.Background = Brushes.Brown
                Case "PreviewDatabaseTab"
                    database_button.Background = Brushes.Brown
                Case "AnalysisTab"
                    analysis_button.Background = Brushes.Brown
                Case "InputDataTab"
                    inputData_button.Background = Brushes.Brown
                Case "ScenarioAnalysisTab"
                    results_button.Background = Brushes.Brown
            End Select
        End If
    End Sub

    Private Async Sub CalculateRatiosButton_Click(sender As Object, e As RoutedEventArgs)
        Dim period As String = CType(internalPeriod.SelectedItem, ComboBoxItem).Content.ToString()
        Dim periodValue As Integer

        If Not Integer.TryParse(periodInput.Text, periodValue) Then
            MessageBox.Show("Please enter a valid number for the period value.")
            Return
        End If

        Try
            ' Retrieve financial data asynchronously based on selected period and period value
            Dim financialDataList As List(Of FinancialData) = Await dbHelper.GetFinancialDataAsync()

            If financialDataList Is Nothing OrElse financialDataList.Count = 0 Then
                MessageBox.Show("No financial data found for the selected period.")
                Return
            End If

            Dim results As New Dictionary(Of String, List(Of Double))
            Dim helperMethods As New FinancialData()

            Dim totalAssets As Double = 0
            Dim totalNetIncome As Double = 0
            Dim totalEquity As Double = 0
            Dim totalRevenue As Double = 0
            Dim TotalOperatingExpenses As Double = 0
            Dim totalCostOfGoodsSold As Double = 0
            Dim totalInterestExpense As Double = 0
            Dim totalVariableCosts As Double = 0
            Dim totalFixedCosts As Double = 0
            Dim totalSalesRevenuePerUnit As Double = 0
            Dim totalVariableCostPerUnit As Double = 0
            Dim totalLiabilities As Double = 0
            Dim totalCurrentLiabilities As Double = 0
            Dim totalCurrentAssets As Double = 0
            Dim totalEbitda As Double = 0

            For Each data As FinancialData In financialDataList
                If period = "Month" AndAlso data.DateValue.Month = periodValue Then
                    Debug.WriteLine(data.DateValue.ToString("yyyy-MM-dd"))
                    totalAssets += data.TotalAssets
                    Debug.WriteLine("TotalAssets: " & totalAssets)

                    totalNetIncome += data.NetIncome
                    Debug.WriteLine("TotalNetIncome: " & totalNetIncome)

                    totalEquity += data.TotalEquity
                    Debug.WriteLine("TotalEquity: " & totalEquity)

                    totalRevenue += data.Revenue
                    Debug.WriteLine("TotalRevenue: " & totalRevenue)

                    TotalOperatingExpenses += data.OperatingExpenses
                    Debug.WriteLine("TotalOperatingExpenses: " & TotalOperatingExpenses)

                    totalCostOfGoodsSold += data.CostOfGoodsSold
                    Debug.WriteLine("TotalCostOfGoodsSold: " & totalCostOfGoodsSold)

                    totalInterestExpense += data.InterestExpense
                    Debug.WriteLine("TotalInterestExpense: " & totalInterestExpense)

                    totalVariableCosts += data.VariableCosts
                    Debug.WriteLine("TotalVariableCosts: " & totalVariableCosts)

                    totalFixedCosts += data.FixedCosts
                    Debug.WriteLine("TotalFixedCosts: " & totalFixedCosts)

                    totalSalesRevenuePerUnit += data.SalesRevenuePerUnit
                    Debug.WriteLine("TotalSalesRevenuePerUnit: " & totalSalesRevenuePerUnit)

                    totalVariableCostPerUnit += data.VariableCostPerUnit
                    Debug.WriteLine("TotalVariableCostPerUnit: " & totalVariableCostPerUnit)

                    totalLiabilities += data.TotalLiabilities
                    Debug.WriteLine("TotalLiabilities: " & totalLiabilities)

                    totalCurrentLiabilities += data.CurrentLiabilities
                    Debug.WriteLine("TotalCurrentLiabilities: " & totalCurrentLiabilities)

                    totalCurrentAssets += data.CurrentAssets
                    Debug.WriteLine("TotalCurrentAssets: " & totalCurrentAssets)

                    totalEbitda += data.EBITDA
                    Debug.WriteLine("TotalEbitda: " & totalEbitda)
                ElseIf period = "Quarter" AndAlso data.DateValue.Month >= (periodValue - 2) AndAlso data.DateValue.Month <= periodValue Then
                    Debug.WriteLine(data.DateValue.ToString("yyyy-MM-dd"))
                    totalAssets += data.TotalAssets
                    Debug.WriteLine("TotalAssets: " & totalAssets)

                    totalNetIncome += data.NetIncome
                    Debug.WriteLine("TotalNetIncome: " & totalNetIncome)

                    totalEquity += data.TotalEquity
                    Debug.WriteLine("TotalEquity: " & totalEquity)

                    totalRevenue += data.Revenue
                    Debug.WriteLine("TotalRevenue: " & totalRevenue)

                    TotalOperatingExpenses += data.OperatingExpenses
                    Debug.WriteLine("TotalOperatingExpenses: " & TotalOperatingExpenses)

                    totalCostOfGoodsSold += data.CostOfGoodsSold
                    Debug.WriteLine("TotalCostOfGoodsSold: " & totalCostOfGoodsSold)

                    totalInterestExpense += data.InterestExpense
                    Debug.WriteLine("TotalInterestExpense: " & totalInterestExpense)

                    totalVariableCosts += data.VariableCosts
                    Debug.WriteLine("TotalVariableCosts: " & totalVariableCosts)

                    totalFixedCosts += data.FixedCosts
                    Debug.WriteLine("TotalFixedCosts: " & totalFixedCosts)

                    totalSalesRevenuePerUnit += data.SalesRevenuePerUnit
                    Debug.WriteLine("TotalSalesRevenuePerUnit: " & totalSalesRevenuePerUnit)

                    totalVariableCostPerUnit += data.VariableCostPerUnit
                    Debug.WriteLine("TotalVariableCostPerUnit: " & totalVariableCostPerUnit)

                    totalLiabilities += data.TotalLiabilities
                    Debug.WriteLine("TotalLiabilities: " & totalLiabilities)

                    totalCurrentLiabilities += data.CurrentLiabilities
                    Debug.WriteLine("TotalCurrentLiabilities: " & totalCurrentLiabilities)

                    totalCurrentAssets += data.CurrentAssets
                    Debug.WriteLine("TotalCurrentAssets: " & totalCurrentAssets)

                    totalEbitda += data.EBITDA
                    Debug.WriteLine("TotalEbitda: " & totalEbitda)

                ElseIf period = "Year" AndAlso data.DateValue.Year = periodValue Then
                    Debug.WriteLine(data.DateValue.ToString("yyyy-MM-dd"))
                    totalAssets += data.TotalAssets
                    Debug.WriteLine("TotalAssets: " & totalAssets)

                    totalNetIncome += data.NetIncome
                    Debug.WriteLine("TotalNetIncome: " & totalNetIncome)

                    totalEquity += data.TotalEquity
                    Debug.WriteLine("TotalEquity: " & totalEquity)

                    totalRevenue += data.Revenue
                    Debug.WriteLine("TotalRevenue: " & totalRevenue)

                    TotalOperatingExpenses += data.OperatingExpenses
                    Debug.WriteLine("TotalOperatingExpenses: " & TotalOperatingExpenses)

                    totalCostOfGoodsSold += data.CostOfGoodsSold
                    Debug.WriteLine("TotalCostOfGoodsSold: " & totalCostOfGoodsSold)

                    totalInterestExpense += data.InterestExpense
                    Debug.WriteLine("TotalInterestExpense: " & totalInterestExpense)

                    totalVariableCosts += data.VariableCosts
                    Debug.WriteLine("TotalVariableCosts: " & totalVariableCosts)

                    totalFixedCosts += data.FixedCosts
                    Debug.WriteLine("TotalFixedCosts: " & totalFixedCosts)

                    totalSalesRevenuePerUnit += data.SalesRevenuePerUnit
                    Debug.WriteLine("TotalSalesRevenuePerUnit: " & totalSalesRevenuePerUnit)

                    totalVariableCostPerUnit += data.VariableCostPerUnit
                    Debug.WriteLine("TotalVariableCostPerUnit: " & totalVariableCostPerUnit)

                    totalLiabilities += data.TotalLiabilities
                    Debug.WriteLine("TotalLiabilities: " & totalLiabilities)

                    totalCurrentLiabilities += data.CurrentLiabilities
                    Debug.WriteLine("TotalCurrentLiabilities: " & totalCurrentLiabilities)

                    totalCurrentAssets += data.CurrentAssets
                    Debug.WriteLine("TotalCurrentAssets: " & totalCurrentAssets)

                    totalEbitda += data.EBITDA
                    Debug.WriteLine("TotalEbitda: " & totalEbitda)
                End If
            Next

            'The FinancialData.ReturnOnAssets() is used from the constructor initialization of FinancialData. 
            If CK_ROA.IsChecked Then
                If totalAssets > 0 Then
                    Dim roa As Double = FinancialData.ReturnOnAssets(totalNetIncome, totalAssets)
                    Debug.WriteLine($"TotalAssets: {totalAssets}, TotalNetIncome: {totalNetIncome}, ROA: {roa}")
                Else
                    Debug.WriteLine("TotalAssets is zero or less, cannot calculate ROA")
                End If
            End If

            If CK_ROE.IsChecked Then
                If totalEquity > 0 Then
                    Dim roe As Double = totalNetIncome / totalEquity
                    Debug.WriteLine($"TotalEquity: {totalEquity}, TotalNetIncome: {totalNetIncome}, ROE: {roe}")
                Else
                    Debug.WriteLine("TotalEquity is zero or less, cannot calculate ROE")
                End If
            End If

            For Each result In results
                Debug.WriteLine($"{result.Key}: {result.Value}")
            Next

            If CK_OperatingMargin.IsChecked Then
                Dim operatingMargin As Double = helperMethods.OperatingProfitMargin(totalRevenue, totalCostOfGoodsSold, TotalOperatingExpenses)
                Debug.WriteLine($"Operating Margin: {operatingMargin}")
            End If

            If CK_NetProfitMargin.IsChecked Then
                If totalRevenue > 0 Then
                    Dim netProfitMargin As Double = helperMethods.NetProfitMargin(totalRevenue, totalCostOfGoodsSold, TotalOperatingExpenses, totalNetIncome)
                    Debug.WriteLine($"TotalRevenue: {totalRevenue}, TotalCostOfGoodsSold: {totalCostOfGoodsSold}, TotalOperatingExpenses: {TotalOperatingExpenses}, TotalNetIncome: {totalNetIncome}, NPM: {netProfitMargin}")
                Else
                    Debug.WriteLine("Total Revenue is zero or less, cannot calculate NPM")
                End If
            End If

            If CK_GrossProfitMargin.IsChecked Then
                If totalRevenue > 0 Then
                    Dim grossProfitMargin As Double = helperMethods.GrossProfitMargin(totalRevenue, totalCostOfGoodsSold)
                    Debug.WriteLine($"TotalRevenue: {totalRevenue}, TotalCostOfGoodsSold: {totalCostOfGoodsSold}, GPM: {grossProfitMargin}")
                Else
                    Debug.WriteLine("Total Revenue is zero or less, cannot calculate GPM")
                End If
            End If

            If CK_CurrentRatios.IsChecked Then
                If totalCurrentLiabilities > 0 Then
                    Dim currentRatio As Double = helperMethods.CurrentRatio(totalCurrentAssets, totalCurrentLiabilities)
                    Debug.WriteLine($"TotalCurrentAssets: {totalCurrentAssets}, TotalCurrentLiabilities: {totalCurrentLiabilities}, CR: {currentRatio}")
                Else
                    Debug.WriteLine("Total Current Liabilities is zero or less, cannot calculate CR")
                End If
            End If

            If CK_DebtToEquity.IsChecked Then
                If totalEquity > 0 Then
                    Dim debtToEquityRatio As Double = helperMethods.DebtToEquityRatio(totalLiabilities, totalEquity)
                    Debug.WriteLine($"TotalLiabilities: {totalLiabilities}, TotalEquity: {totalEquity}, D/E: {debtToEquityRatio}")
                Else
                    Debug.WriteLine("Total Equity is zero or less, cannot calculate D/E")
                End If
            End If

            If CK_InterestCoverage.IsChecked Then
                If totalInterestExpense > 0 Then
                    Dim interestCoverageRatio As Double = helperMethods.InterestCoverageRatio(totalEbitda, totalInterestExpense)
                    Debug.WriteLine($"TotalNetIncome: {totalEbitda}, TotalInterestExpense: {totalInterestExpense}, ICR: {interestCoverageRatio}")
                Else
                    Debug.WriteLine("Total Interest Expense is zero or less, cannot calculate ICR")
                End If
            End If

            If CK_ContributionMargin.IsChecked Then
                If totalSalesRevenuePerUnit > 0 Then
                    Dim contributionMargin As Double = helperMethods.ContributionMarginRatio(totalSalesRevenuePerUnit, totalVariableCostPerUnit)
                    Debug.WriteLine($"TotalSalesRevenuePerUnit: {totalSalesRevenuePerUnit}, TotalVariableCostPerUnit: {totalVariableCostPerUnit}, CM: {contributionMargin}")
                Else
                    Debug.WriteLine("Total Sales Revenue Per Unit is zero or less, cannot calculate CM")
                End If
            End If

            If CK_BreakEvenPoint.IsChecked Then
                If totalFixedCosts > 0 Then
                    Dim breakEvenPoint As Double = helperMethods.BreakEvenPoint(totalFixedCosts, totalSalesRevenuePerUnit, totalVariableCostPerUnit)
                    Debug.WriteLine($"TotalFixedCosts: {totalFixedCosts}, TotalSalesRevenuePerUnit: {totalSalesRevenuePerUnit}, TotalVariableCostPerUnit: {totalVariableCostPerUnit}, BEP: {breakEvenPoint}")
                Else
                    Debug.WriteLine("Total Fixed Costs is zero or less, cannot calculate BEP")
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error fetching or calculating financial ratios: {ex.Message}")
        End Try
    End Sub


    Private Sub ClickableText_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
        Debug.WriteLine("Clickable text was clicked!")
        authenatication_win = New Authentication()
        Dim result As Nullable(Of Boolean) = authenatication_win.ShowDialog()

        If result.HasValue AndAlso result.Value Then
            Debug.WriteLine("User authenticated successfully!")
            LogIn.Text = "Welcome, User"
            LogIn.IsEnabled = False
        Else
            Debug.WriteLine("User authentication failed.")
        End If
    End Sub

    'It clears all the input fields from the InputDataTab
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

    'It is used to export files
    Private Sub ExportPDFButton_Click(sender As Object, e As RoutedEventArgs)
        Dim saveFileDialog As New SaveFileDialog() With {
            .Filter = "PDF Files|*.pdf",
            .Title = "Save PDF File"
        }

        If saveFileDialog.ShowDialog() = True Then
            Dim filePath As String = saveFileDialog.FileName

            ' Create the PDF file
            Try
                ' Create a new PDF document
                Dim document As New PdfDocument()
                document.Info.Title = "Created with PDFsharp"

                ' Create an empty page
                Dim page As PdfPage = document.AddPage()

                ' Get an XGraphics object for drawing
                Dim gfx As XGraphics = XGraphics.FromPdfPage(page)

                ' Create a font
                Dim font As XFont = New XFont("Verdana", 20)
                ' Draw the text
                gfx.DrawString("Hello, World!", font, XBrushes.Black, New XRect(0, 0, page.Width, page.Height), XStringFormats.Center)
                gfx.DrawString("This is a sample PDF file created using PDFsharp.", font, XBrushes.Black, New XRect(0, 40, page.Width, page.Height), XStringFormats.Center)

                ' Save the document
                document.Save(filePath)
                document.Close()

                MessageBox.Show("PDF file exported successfully!")
            Catch ex As IOException
                MessageBox.Show($"An IO exception occurred while exporting the PDF: {ex.Message}")
            Catch ex As UnauthorizedAccessException
                MessageBox.Show($"An access exception occurred while exporting the PDF: {ex.Message}")
            Catch ex As Exception
                MessageBox.Show($"An unexpected error occurred while exporting the PDF: {ex.Message}")
            End Try

            Debug.WriteLine($"Exporting PDF file to: {filePath}")
        End If
    End Sub



End Class