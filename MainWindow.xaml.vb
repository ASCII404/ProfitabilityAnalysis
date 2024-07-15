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
Imports System.Runtime.Serialization
Imports System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder
Imports System.Diagnostics.Eventing.Reader
Imports System.Text



Class MainWindow

    Private dbHelper As Database
    Private apiKey As String
    Private FinancialData As FinancialData
    Private financialDataList As List(Of FinancialData)
    Private previous_selected_tab As String
    Private helpButton_content As String
    Private authenatication_win As Authentication

    Private financialDataDict As New Dictionary(Of String, Double)
    Private analysis_results As New Dictionary(Of String, Double)
    Private _userName As String
    Public Property Chart1Values As ChartValues(Of ObservablePoint)
    Public Property Chart2Values As ChartValues(Of ObservablePoint)
    Public Property Chart3Values As ChartValues(Of ObservablePoint)
    Public Property Chart4Values As ChartValues(Of ObservablePoint)
    Public Sub New(userName As String)

        apiKey = "your Alpha Vantage API Key"
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        InitializeComponent()
        dbHelper = New Database("dbtest.db")
        FinancialData = New FinancialData()

        _userName = userName
        LogIn.Text = "Welcome, " + _userName
        InitializeChartData1()
        InitializeChartData2()
        InitializeChartData3()
        InitializeChartData4()

    End Sub

    Private Async Sub InitializeChartData1()
        Dim financialDataList As List(Of FinancialData) = Nothing

        Try
            financialDataList = Await dbHelper.GetFinancialDataAsync()
        Catch ex As Exception
            MessageBox.Show("Data could not have been loaded")
        End Try

        Await Dispatcher.InvokeAsync(Sub()
                                         If financialDataList Is Nothing OrElse financialDataList.Count = 0 Then
                                             ' Show placeholder and hide chart if data is not available
                                             placeholder1.Visibility = Visibility.Visible
                                             chart1.Visibility = Visibility.Collapsed
                                         Else
                                             ' Initialize Chart1Values if not already done
                                             If Chart1Values Is Nothing Then
                                                 Chart1Values = New ChartValues(Of ObservablePoint)()
                                             End If

                                             ' Clear existing data
                                             Chart1Values.Clear()

                                             ' Define a dictionary to store cumulative sum and count for each month
                                             Dim monthTotals As New Dictionary(Of Integer, Tuple(Of Integer, Integer))

                                             ' Iterate through financialDataList to calculate cumulative sum and count for each month
                                             For Each data As FinancialData In financialDataList
                                                 Dim month As Integer = data.DateValue.Month
                                                 Dim cost As Integer = Convert.ToInt32(data.NetIncome)
                                                 Debug.WriteLine("This is data" & data.NetIncome)

                                                 If Not monthTotals.ContainsKey(month) Then
                                                     ' Initialize cumulative sum and count for the month
                                                     monthTotals(month) = Tuple.Create(cost, 1)
                                                 Else
                                                     ' Accumulate cumulative sum and count for the month
                                                     Dim currentTotal = monthTotals(month)
                                                     monthTotals(month) = Tuple.Create(currentTotal.Item1 + cost, currentTotal.Item2 + 1)
                                                 End If
                                             Next

                                             ' Populate Chart1Values with the average for each month
                                             For Each monthTotal In monthTotals
                                                 Dim month As Integer = monthTotal.Key
                                                 Dim sum As Integer = monthTotal.Value.Item1
                                                 Dim count As Integer = monthTotal.Value.Item2
                                                 Dim average As Integer = sum \ count ' Use integer division

                                                 ' Add the average to Chart1Values
                                                 Chart1Values.Add(New ObservablePoint(month, average))
                                             Next

                                             ' Bind data to Chart1
                                             chart1.Series = New SeriesCollection From {
                                         New LineSeries With {
                                             .Values = Chart1Values,
                                             .Title = "Net income:",
                                             .Fill = Brushes.CadetBlue,
                                             .Stroke = Brushes.CadetBlue
                                         }
                                     }

                                             ' Set the axis ranges
                                             chart1.AxisX.Clear()
                                             chart1.AxisX.Add(New Axis With {
                                         .Title = "Month",
                                         .MinValue = 1,
                                         .MaxValue = 12,
                                         .Foreground = Brushes.Black
                                     })

                                             chart1.AxisY.Clear()
                                             chart1.AxisY.Add(New Axis With {
                                         .Title = "Amount",
                                         .MinValue = 0,
                                         .MaxValue = 30000,
                                         .Foreground = Brushes.Black
                                     })

                                             ' Update UI elements visibility
                                             placeholder1.Visibility = Visibility.Collapsed
                                             chart1.Visibility = Visibility.Visible
                                         End If
                                     End Sub)
    End Sub

    Private Async Sub InitializeChartData2()
        Dim financialDataList As List(Of FinancialData) = Nothing

        Try
            financialDataList = Await dbHelper.GetFinancialDataAsync()
        Catch ex As Exception
            MessageBox.Show("Data could not have been loaded")
        End Try

        Await Dispatcher.InvokeAsync(Sub()
                                         If financialDataList Is Nothing OrElse financialDataList.Count = 0 Then
                                             ' Show placeholder and hide chart if data is not available
                                             placeholder2.Visibility = Visibility.Visible
                                             chart2.Visibility = Visibility.Collapsed
                                         Else
                                             ' Initialize Chart2Values if not already done
                                             If Chart2Values Is Nothing Then
                                                 Chart2Values = New ChartValues(Of ObservablePoint)()
                                             End If

                                             ' Clear existing data
                                             Chart2Values.Clear()

                                             ' Define a dictionary to store cumulative sum and count for each month
                                             Dim monthTotals As New Dictionary(Of Integer, Tuple(Of Integer, Integer))

                                             ' Iterate through financialDataList to calculate cumulative sum and count for each month
                                             For Each data As FinancialData In financialDataList
                                                 Dim month As Integer = data.DateValue.Month
                                                 Dim cost As Integer = Convert.ToInt32(data.TotalAssets)
                                                 Debug.WriteLine("This is data" & data.TotalAssets)

                                                 If Not monthTotals.ContainsKey(month) Then
                                                     ' Initialize cumulative sum and count for the month
                                                     monthTotals(month) = Tuple.Create(cost, 1)
                                                 Else
                                                     ' Accumulate cumulative sum and count for the month
                                                     Dim currentTotal = monthTotals(month)
                                                     monthTotals(month) = Tuple.Create(currentTotal.Item1 + cost, currentTotal.Item2 + 1)
                                                 End If
                                             Next

                                             ' Populate Chart2Values with the average for each month
                                             For Each monthTotal In monthTotals
                                                 Dim month As Integer = monthTotal.Key
                                                 Dim sum As Integer = monthTotal.Value.Item1
                                                 Dim count As Integer = monthTotal.Value.Item2
                                                 Dim average As Integer = sum \ count ' Use integer division

                                                 ' Add the average to Chart2Values
                                                 Chart2Values.Add(New ObservablePoint(month, average))
                                             Next

                                             ' Bind data to Chart 2
                                             chart2.Series = New SeriesCollection From {
                                         New ColumnSeries With {
                                             .Values = Chart2Values,
                                             .Title = "Assets value:",
                                             .Fill = Brushes.CadetBlue,
                                             .Stroke = Brushes.CadetBlue
                                         }
                                     }

                                             ' Set the axis ranges
                                             chart2.AxisX.Clear()
                                             chart2.AxisX.Add(New Axis With {
                                         .Title = "Month",
                                         .MinValue = 1,
                                         .MaxValue = 13,
                                         .Foreground = Brushes.Black
                                     })

                                             chart2.AxisY.Clear()
                                             chart2.AxisY.Add(New Axis With {
                                         .Title = "Amount",
                                         .MinValue = 0,
                                         .MaxValue = 30000,
                                         .Foreground = Brushes.Black
                                     })

                                             ' Update UI elements visibility
                                             placeholder2.Visibility = Visibility.Collapsed
                                             chart2.Visibility = Visibility.Visible
                                         End If
                                     End Sub)
    End Sub

    Private Async Sub InitializeChartData3()
        Dim financialDataList As List(Of FinancialData) = Nothing

        Try
            financialDataList = Await dbHelper.GetFinancialDataAsync()
        Catch ex As Exception
            MessageBox.Show("Data could not have been loaded")
        End Try

        Await Dispatcher.InvokeAsync(Sub()
                                         If financialDataList Is Nothing OrElse financialDataList.Count = 0 Then
                                             ' Show placeholder and hide chart if data is not available
                                             placeholder3.Visibility = Visibility.Visible
                                             chart3.Visibility = Visibility.Collapsed
                                         Else
                                             ' Initialize Chart2Values if not already done
                                             If Chart3Values Is Nothing Then
                                                 Chart3Values = New ChartValues(Of ObservablePoint)()
                                             End If

                                             ' Clear existing data
                                             Chart3Values.Clear()

                                             ' Define a dictionary to store cumulative sum and count for each month
                                             Dim monthTotals As New Dictionary(Of Integer, Tuple(Of Integer, Integer))

                                             ' Iterate through financialDataList to calculate cumulative sum and count for each month
                                             For Each data As FinancialData In financialDataList
                                                 Dim month As Integer = data.DateValue.Month
                                                 Dim cost As Integer = Convert.ToInt32(data.TotalAssets)
                                                 Debug.WriteLine("This is data" & data.TotalAssets)

                                                 If Not monthTotals.ContainsKey(month) Then
                                                     ' Initialize cumulative sum and count for the month
                                                     monthTotals(month) = Tuple.Create(cost, 1)
                                                 Else
                                                     ' Accumulate cumulative sum and count for the month
                                                     Dim currentTotal = monthTotals(month)
                                                     monthTotals(month) = Tuple.Create(currentTotal.Item1 + cost, currentTotal.Item2 + 1)
                                                 End If
                                             Next

                                             ' Populate Chart2Values with the average for each month
                                             For Each monthTotal In monthTotals
                                                 Dim month As Integer = monthTotal.Key
                                                 Dim sum As Integer = monthTotal.Value.Item1
                                                 Dim count As Integer = monthTotal.Value.Item2
                                                 Dim average As Integer = sum \ count ' Use integer division

                                                 ' Add the average to Chart2Values
                                                 Chart3Values.Add(New ObservablePoint(month, average))
                                             Next

                                             ' Bind data to Chart 2
                                             chart3.Series = New SeriesCollection From {
                                         New ColumnSeries With {
                                             .Values = Chart3Values,
                                             .Title = "Costs of Goods Sold:",
                                             .Fill = Brushes.CadetBlue,
                                             .Stroke = Brushes.CadetBlue
                                         }
                                     }

                                             ' Set the axis ranges
                                             chart3.AxisX.Clear()
                                             chart3.AxisX.Add(New Axis With {
                                         .Title = "Month",
                                         .MinValue = 1,
                                         .MaxValue = 13,
                                         .Foreground = Brushes.Black
                                     })

                                             chart3.AxisY.Clear()
                                             chart3.AxisY.Add(New Axis With {
                                         .Title = "Amount",
                                         .MinValue = 0,
                                         .MaxValue = 30000,
                                         .Foreground = Brushes.Black
                                     })

                                             ' Update UI elements visibility
                                             placeholder3.Visibility = Visibility.Collapsed
                                             chart3.Visibility = Visibility.Visible
                                         End If
                                     End Sub)
    End Sub

    Private Async Sub InitializeChartData4()
        Dim financialDataList As List(Of FinancialData) = Nothing

        Try
            financialDataList = Await dbHelper.GetFinancialDataAsync()
        Catch ex As Exception
            MessageBox.Show("Data could not have been loaded")
        End Try

        Await Dispatcher.InvokeAsync(Sub()
                                         If financialDataList Is Nothing OrElse financialDataList.Count = 0 Then
                                             ' Show placeholder and hide chart if data is not available
                                             placeholder4.Visibility = Visibility.Visible
                                             chart4.Visibility = Visibility.Collapsed
                                         Else
                                             ' Initialize Chart1Values if not already done
                                             If Chart4Values Is Nothing Then
                                                 Chart4Values = New ChartValues(Of ObservablePoint)()
                                             End If

                                             ' Clear existing data
                                             Chart4Values.Clear()

                                             ' Define a dictionary to store cumulative sum and count for each month
                                             Dim monthTotals As New Dictionary(Of Integer, Tuple(Of Integer, Integer))

                                             ' Iterate through financialDataList to calculate cumulative sum and count for each month
                                             For Each data As FinancialData In financialDataList
                                                 Dim month As Integer = data.DateValue.Month
                                                 Dim cost As Integer = Convert.ToInt32(data.Revenue)
                                                 Debug.WriteLine("This is data" & data.Revenue)

                                                 If Not monthTotals.ContainsKey(month) Then
                                                     ' Initialize cumulative sum and count for the month
                                                     monthTotals(month) = Tuple.Create(cost, 1)
                                                 Else
                                                     ' Accumulate cumulative sum and count for the month
                                                     Dim currentTotal = monthTotals(month)
                                                     monthTotals(month) = Tuple.Create(currentTotal.Item1 + cost, currentTotal.Item2 + 1)
                                                 End If
                                             Next

                                             ' Populate Chart1Values with the average for each month
                                             For Each monthTotal In monthTotals
                                                 Dim month As Integer = monthTotal.Key
                                                 Dim sum As Integer = monthTotal.Value.Item1
                                                 Dim count As Integer = monthTotal.Value.Item2
                                                 Dim average As Integer = sum \ count ' Use integer division

                                                 ' Add the average to Chart1Values
                                                 Chart4Values.Add(New ObservablePoint(month, average))
                                             Next

                                             ' Bind data to Chart1
                                             chart4.Series = New SeriesCollection From {
                                         New LineSeries With {
                                             .Values = Chart4Values,
                                             .Title = "Revenue:",
                                             .Fill = Brushes.CadetBlue,
                                             .Stroke = Brushes.CadetBlue
                                         }
                                     }

                                             ' Set the axis ranges
                                             chart4.AxisX.Clear()
                                             chart4.AxisX.Add(New Axis With {
                                         .Title = "Month",
                                         .MinValue = 1,
                                         .MaxValue = 12,
                                         .Foreground = Brushes.Black
                                     })

                                             chart4.AxisY.Clear()
                                             chart4.AxisY.Add(New Axis With {
                                         .Title = "Amount",
                                         .MinValue = 0,
                                         .MaxValue = 30000,
                                         .Foreground = Brushes.Black
                                     })

                                             ' Update UI elements visibility
                                             placeholder4.Visibility = Visibility.Collapsed
                                             chart4.Visibility = Visibility.Visible
                                         End If
                                     End Sub)
    End Sub



    Private Sub DashboardButton_Click(sender As Object, e As RoutedEventArgs)
        MainTabControl.SelectedItem = DashboardTab
        InitializeChartData1()
        InitializeChartData2()
        InitializeChartData3()
        InitializeChartData4()
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

        financialDataDict("Revenue") = financialData.Revenue
        financialDataDict("CostOfGoodsSold") = financialData.CostOfGoodsSold
        financialDataDict("OperatingExpenses") = financialData.OperatingExpenses
        financialDataDict("NetIncome") = financialData.NetIncome
        financialDataDict("EBITDA") = financialData.EBITDA
        financialDataDict("InterestExpense") = financialData.InterestExpense

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


        financialDataDict("TotalAssets") = financialData.TotalAssets
        financialDataDict("TotalEquity") = financialData.TotalEquity
        financialDataDict("CurrentAssets") = financialData.CurrentAssets
        financialDataDict("CurrentLiabilities") = financialData.CurrentLiabilities
        financialDataDict("TotalLiabilities") = financialData.TotalLiabilities

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
                    helpButton_content = "In order to do the analysis, pick a specific period like a month or a quarter and enter the period value. " & vbCrLf &
                                         "Then click the 'Calculate Ratios' button." & vbCrLf &
                                         "The tool will calculate the ratios for the selected period and period value." & vbCrLf &
                                         "For example, if you picked Quarter and entered 1, the tool will calculate the ratios for the first quarter." & vbCrLf &
                                         "Also, you can calculate specific metrics for available businesses by selecting a symbol and a fiscal year." & vbCrLf &
                                         "Then click the 'Get API Data' button to get the data from the API." & vbCrLf &
                                         "After data is retrieved, you can check boxes like ROA, ROE, etc. and click the 'Calculate Ratios' button."
                ElseIf selectedTabName = "InputDataTab" Then
                    helpButton_content = "This tab offers 2 possibilities: Enter specific values for only 1 day or import your financial data directly through an excel file" & vbCrLf &
                                         "For the import functionality to work, you need to have the same column names and the same order as the ones in the database." & vbCrLf &
                                         "The first row of the excel file should contain the column names." & vbCrLf &
                                         "After importing the data, you can click the 'Add to database' button to add the data to the database."
                ElseIf selectedTabName = "ScenarioAnalysisTab" Then
                    helpButton_content = "Here you can see your results from the analysis procedure for both internal and external perspectives" & vbCrLf &
                                         "You can also export the data to a PDF file by clicking the 'Export to PDF' button."
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

    Private Sub reset_values()
        ROA_res.Text = "0%"
        ROE_res.Text = "0%"
        OPM_res.Text = "0%"
        GPM_res.Text = "0%"
        NPM_res.Text = "0%"
        CRR_res.Text = "0"
        DTE_res.Text = "0"
        IC_res.Text = "0"
        CM_res.Text = "0%"
        BEP_res.Text = "0"

        ROA_api.Text = "0%"
        ROE_api.Text = "0%"
        OPM_api.Text = "0%"
        GPM_api.Text = "0%"
        NPM_api.Text = "0%"
        CRR_api.Text = "0"
        DTE_api.Text = "0"
        IC_api.Text = "0"
        CM_api.Text = "0%"
        BEP_api.Text = "0"

    End Sub
    Private Async Sub CalculateRatiosButton_Click(sender As Object, e As RoutedEventArgs)
        reset_values()
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
                If data.DateValue.Year = DateTime.Now.Year Then
                    If period = "Month" AndAlso data.DateValue.Month = periodValue Then
                        Debug.WriteLine(data.DateValue.ToString("yyyy-MM-dd"))
                        totalAssets = data.TotalAssets
                        totalNetIncome += data.NetIncome
                        totalEquity = data.TotalEquity
                        totalRevenue += data.Revenue
                        TotalOperatingExpenses += data.OperatingExpenses
                        totalCostOfGoodsSold += data.CostOfGoodsSold
                        totalInterestExpense += data.InterestExpense
                        totalVariableCosts += data.VariableCosts
                        totalFixedCosts += data.FixedCosts
                        totalSalesRevenuePerUnit += data.SalesRevenuePerUnit
                        totalVariableCostPerUnit += data.VariableCostPerUnit
                        totalLiabilities += data.TotalLiabilities
                        totalCurrentLiabilities += data.CurrentLiabilities
                        totalCurrentAssets += data.CurrentAssets
                        totalEbitda += data.EBITDA
                    ElseIf period = "Quarter" AndAlso data.DateValue.Month >= (periodValue * 4) - 3 AndAlso data.DateValue.Month <= periodValue * 4 Then
                        Debug.WriteLine(data.DateValue.ToString("yyyy-MM-dd"))
                        totalAssets = data.TotalAssets
                        totalNetIncome += data.NetIncome
                        totalEquity = data.TotalEquity
                        totalRevenue += data.Revenue
                        TotalOperatingExpenses += data.OperatingExpenses
                        totalCostOfGoodsSold += data.CostOfGoodsSold
                        totalInterestExpense += data.InterestExpense
                        totalVariableCosts += data.VariableCosts
                        totalFixedCosts += data.FixedCosts
                        totalSalesRevenuePerUnit += data.SalesRevenuePerUnit
                        totalVariableCostPerUnit += data.VariableCostPerUnit
                        totalLiabilities += data.TotalLiabilities
                        totalCurrentLiabilities += data.CurrentLiabilities
                        totalCurrentAssets += data.CurrentAssets
                        totalEbitda += data.EBITDA
                    End If
                End If

                If period = "Year" AndAlso data.DateValue.Year = periodValue Then
                    Debug.WriteLine(data.DateValue.ToString("yyyy-MM-dd"))
                    totalAssets = data.TotalAssets
                    totalNetIncome += data.NetIncome
                    totalEquity = data.TotalEquity
                    totalRevenue += data.Revenue
                    TotalOperatingExpenses += data.OperatingExpenses
                    totalCostOfGoodsSold += data.CostOfGoodsSold
                    totalInterestExpense += data.InterestExpense
                    totalVariableCosts += data.VariableCosts
                    totalFixedCosts += data.FixedCosts
                    totalSalesRevenuePerUnit += data.SalesRevenuePerUnit
                    totalVariableCostPerUnit += data.VariableCostPerUnit
                    totalLiabilities += data.TotalLiabilities
                    totalCurrentLiabilities += data.CurrentLiabilities
                    totalCurrentAssets += data.CurrentAssets
                    totalEbitda += data.EBITDA
                End If
            Next

            If period = "Quarter" Then
                totalNetIncome = totalNetIncome / 120
                totalRevenue = totalRevenue / 120
                TotalOperatingExpenses = TotalOperatingExpenses / 120
                totalCostOfGoodsSold = totalCostOfGoodsSold / 120
                totalInterestExpense = totalInterestExpense / 120
                totalVariableCosts = totalVariableCosts / 120
                totalFixedCosts = totalFixedCosts / 120
                totalSalesRevenuePerUnit = totalSalesRevenuePerUnit / 120
                totalVariableCostPerUnit = totalVariableCostPerUnit / 120
                totalLiabilities = totalLiabilities / 120
                totalCurrentLiabilities = totalCurrentLiabilities / 120
                totalCurrentAssets = totalCurrentAssets / 120
                totalEbitda = totalEbitda / 120
            End If
            If period = "Year" Then
                totalNetIncome = totalNetIncome / 360
                totalRevenue = totalRevenue / 360
                TotalOperatingExpenses = TotalOperatingExpenses / 360
                totalCostOfGoodsSold = totalCostOfGoodsSold / 360
                totalInterestExpense = totalInterestExpense / 360
                totalVariableCosts = totalVariableCosts / 360
                totalFixedCosts = totalFixedCosts / 360
                totalSalesRevenuePerUnit = totalSalesRevenuePerUnit / 360
                totalVariableCostPerUnit = totalVariableCostPerUnit / 360
                totalLiabilities = totalLiabilities / 360
                totalCurrentLiabilities = totalCurrentLiabilities / 360
                totalCurrentAssets = totalCurrentAssets / 360
                totalEbitda = totalEbitda / 360
            End If
            'The FinancialData.ReturnOnAssets() is used from the constructor initialization of FinancialData. 
            If CK_ROA.IsChecked Then
                If totalAssets > 0 Then
                    Dim roa As Double = FinancialData.ReturnOnAssets(totalNetIncome, totalAssets)
                    Debug.WriteLine($"TotalAssets: {totalAssets}, TotalNetIncome: {totalNetIncome}, ROA: {roa}")
                    roa = Math.Floor(roa * 100) / 100
                    ROA_res.Text = roa.ToString() & "%"
                    analysis_results("ROA") = roa
                Else
                    Debug.WriteLine("TotalAssets is zero or less, cannot calculate ROA")
                End If
            End If

            If CK_ROE.IsChecked Then
                If totalEquity > 0 Then
                    Dim roe As Double = totalNetIncome / totalEquity
                    roe = Math.Floor(roe * 100) / 100
                    ROE_res.Text = roe.ToString() & "%"
                    Debug.WriteLine($"TotalEquity: {totalEquity}, TotalNetIncome: {totalNetIncome}, ROE: {roe}")
                    analysis_results("ROE") = roe
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
                operatingMargin = Math.Floor(operatingMargin * 100) / 100
                OPM_res.Text = operatingMargin.ToString() & "%"
                analysis_results("OPM") = operatingMargin
            End If

            If CK_NetProfitMargin.IsChecked Then
                If totalRevenue > 0 Then
                    Dim netProfitMargin As Double = helperMethods.NetProfitMargin(totalRevenue, totalCostOfGoodsSold, TotalOperatingExpenses, totalNetIncome)
                    Debug.WriteLine($"TotalRevenue: {totalRevenue}, TotalCostOfGoodsSold: {totalCostOfGoodsSold}, TotalOperatingExpenses: {TotalOperatingExpenses}, TotalNetIncome: {totalNetIncome}, NPM: {netProfitMargin}")
                    netProfitMargin = Math.Floor(netProfitMargin * 100) / 100
                    NPM_res.Text = netProfitMargin.ToString() & "%"
                    analysis_results("NPM") = netProfitMargin
                Else
                    Debug.WriteLine("Total Revenue is zero or less, cannot calculate NPM")
                End If
            End If

            If CK_GrossProfitMargin.IsChecked Then
                If totalRevenue > 0 Then
                    Dim grossProfitMargin As Double = helperMethods.GrossProfitMargin(totalRevenue, totalCostOfGoodsSold)
                    Debug.WriteLine($"TotalRevenue: {totalRevenue}, TotalCostOfGoodsSold: {totalCostOfGoodsSold}, GPM: {grossProfitMargin}")
                    grossProfitMargin = Math.Floor(grossProfitMargin * 100) / 100
                    GPM_res.Text = grossProfitMargin.ToString() & "%"
                    analysis_results("GPM") = grossProfitMargin
                Else
                    Debug.WriteLine("Total Revenue is zero or less, cannot calculate GPM")
                End If
            End If

            If CK_CurrentRatios.IsChecked Then
                If totalCurrentLiabilities > 0 Then
                    Dim currentRatio As Double = helperMethods.CurrentRatio(totalCurrentAssets, totalCurrentLiabilities)
                    Debug.WriteLine($"TotalCurrentAssets: {totalCurrentAssets}, TotalCurrentLiabilities: {totalCurrentLiabilities}, CR: {currentRatio}")
                    currentRatio = Math.Floor(currentRatio * 100) / 100
                    CRR_res.Text = currentRatio.ToString()
                    analysis_results("CR") = currentRatio
                Else
                    Debug.WriteLine("Total Current Liabilities is zero or less, cannot calculate CR")
                End If
            End If

            If CK_DebtToEquity.IsChecked Then
                If totalEquity > 0 Then
                    Dim debtToEquityRatio As Double = helperMethods.DebtToEquityRatio(totalLiabilities, totalEquity)
                    Debug.WriteLine($"TotalLiabilities: {totalLiabilities}, TotalEquity: {totalEquity}, D/E: {debtToEquityRatio}")
                    debtToEquityRatio = Math.Floor(debtToEquityRatio * 100) / 100
                    DTE_res.Text = debtToEquityRatio.ToString()
                    analysis_results("DTE") = debtToEquityRatio
                Else
                    Debug.WriteLine("Total Equity is zero or less, cannot calculate D/E")
                End If
            End If

            If CK_InterestCoverage.IsChecked Then
                If totalInterestExpense > 0 Then
                    Dim interestCoverageRatio As Double = helperMethods.InterestCoverageRatio(totalEbitda, totalInterestExpense)
                    Debug.WriteLine($"TotalNetIncome: {totalEbitda}, TotalInterestExpense: {totalInterestExpense}, ICR: {interestCoverageRatio}")
                    interestCoverageRatio = Math.Floor(interestCoverageRatio * 100) / 100
                    IC_res.Text = interestCoverageRatio.ToString()
                    analysis_results("ICR") = interestCoverageRatio
                Else
                    Debug.WriteLine("Total Interest Expense is zero or less, cannot calculate ICR")
                End If
            End If

            If CK_ContributionMargin.IsChecked Then
                If totalSalesRevenuePerUnit > 0 Then
                    Dim contributionMargin As Double = helperMethods.ContributionMarginRatio(totalSalesRevenuePerUnit, totalVariableCostPerUnit)
                    Debug.WriteLine($"TotalSalesRevenuePerUnit: {totalSalesRevenuePerUnit}, TotalVariableCostPerUnit: {totalVariableCostPerUnit}, CM: {contributionMargin}")
                    contributionMargin = Math.Floor(contributionMargin * 100) / 100
                    CM_res.Text = contributionMargin.ToString() & "%"
                    analysis_results("CM") = contributionMargin
                Else
                    Debug.WriteLine("Total Sales Revenue Per Unit is zero or less, cannot calculate CM")
                End If
            End If

            If CK_BreakEvenPoint.IsChecked Then
                If totalFixedCosts > 0 Then
                    Dim breakEvenPoint As Double = helperMethods.BreakEvenPoint(totalFixedCosts, totalSalesRevenuePerUnit, totalVariableCostPerUnit)
                    Debug.WriteLine($"TotalFixedCosts: {totalFixedCosts}, TotalSalesRevenuePerUnit: {totalSalesRevenuePerUnit}, TotalVariableCostPerUnit: {totalVariableCostPerUnit}, BEP: {breakEvenPoint}")
                    breakEvenPoint = Math.Floor(breakEvenPoint * 100) / 100
                    BEP_res.Text = breakEvenPoint.ToString("C0")
                    analysis_results("BEP") = breakEvenPoint
                Else
                    Debug.WriteLine("Total Fixed Costs is zero or less, cannot calculate BEP")
                End If
            End If

            'THIS IS FOR EXTERNAL API DATA
            If CK_ROA_API.IsChecked Then
                ' Retrieve TotalNetIncome from dictionary
                Dim api_totalNetIncome As Double
                If Not financialDataDict.TryGetValue("NetIncome", api_totalNetIncome) Then
                    Debug.WriteLine("Net Income not found in dictionary.")
                    Exit Sub ' Exit the subroutine if Net Income is not found
                End If

                ' Retrieve TotalAssets from dictionary
                Dim api_totalAssets As Double
                If Not financialDataDict.TryGetValue("TotalAssets", api_totalAssets) Then
                    Debug.WriteLine("Total Assets not found in dictionary.")
                    Exit Sub ' Exit the subroutine if Total Assets is not found
                End If

                If api_totalAssets > 0 Then
                    Dim roa As Double = FinancialData.ReturnOnAssets(api_totalNetIncome, api_totalAssets)
                    roa = Math.Floor(roa * 100) / 100
                    ROA_api.Text = roa.ToString() & "%"
                    analysis_results("ROA_API") = roa
                Else
                    ' If TotalAssets is zero or less, cannot calculate ROA
                    Debug.WriteLine("TotalAssets is zero or less, cannot calculate ROA")
                End If
            End If

            If CK_ROE_API.IsChecked Then
                ' Retrieve Net Income from dictionary
                Dim roeNetIncome As Decimal
                If Not financialDataDict.TryGetValue("NetIncome", roeNetIncome) Then
                    Debug.WriteLine("Net Income not found in dictionary.")
                Else
                    ' Retrieve Total Equity from dictionary
                    Dim roeTotalEquity As Decimal
                    If Not financialDataDict.TryGetValue("TotalEquity", roeTotalEquity) Then
                        Debug.WriteLine("Total Equity not found in dictionary.")
                    Else
                        ' Calculate ROE if TotalEquity is greater than 0
                        If roeTotalEquity > 0 Then
                            Dim roe As Double = roeNetIncome / roeTotalEquity
                            roe = Math.Floor(roe * 100) / 100
                            ROE_api.Text = roe.ToString() & "%"
                            analysis_results("ROE_API") = roe
                        Else
                            Debug.WriteLine("TotalEquity is zero or less, cannot calculate ROE")
                        End If
                    End If
                End If
            End If

            ' Check if CK_OperatingMargin_API is checked
            If CK_OperatingMargin_API.IsChecked Then
                ' Retrieve relevant values from dictionary
                Dim opMarginTotalRevenue As Decimal
                If Not financialDataDict.TryGetValue("Revenue", opMarginTotalRevenue) Then
                    Debug.WriteLine("Total Revenue not found in dictionary.")
                Else
                    Dim opMarginTotalCostOfGoodsSold As Decimal
                    If Not financialDataDict.TryGetValue("CostOfGoodsSold", opMarginTotalCostOfGoodsSold) Then
                        Debug.WriteLine("Total Cost of Goods Sold not found in dictionary.")
                    Else
                        Dim opMarginTotalOperatingExpenses As Decimal
                        If Not financialDataDict.TryGetValue("OperatingExpenses", opMarginTotalOperatingExpenses) Then
                            Debug.WriteLine("Total Operating Expenses not found in dictionary.")
                        Else
                            ' Calculate Operating Profit Margin
                            Dim operatingMargin As Double = helperMethods.OperatingProfitMargin(opMarginTotalRevenue, opMarginTotalCostOfGoodsSold, opMarginTotalOperatingExpenses)
                            operatingMargin = Math.Floor(operatingMargin * 100) / 100
                            OPM_api.Text = operatingMargin.ToString() & "%"
                            analysis_results("OPM_API") = operatingMargin
                        End If
                    End If
                End If
            End If

            If CK_NetProfitMargin_API.IsChecked Then
                Dim npmTotalRevenue As Decimal
                If Not financialDataDict.TryGetValue("Revenue", npmTotalRevenue) Then
                    Debug.WriteLine("Total Revenue not found in dictionary.")
                Else
                    Dim npmTotalCostOfGoodsSold As Decimal
                    If Not financialDataDict.TryGetValue("CostOfGoodsSold", npmTotalCostOfGoodsSold) Then
                        Debug.WriteLine("Total Cost of Goods Sold not found in dictionary.")
                    Else
                        Dim npmTotalOperatingExpenses As Decimal
                        If Not financialDataDict.TryGetValue("OperatingExpenses", npmTotalOperatingExpenses) Then
                            Debug.WriteLine("Total Operating Expenses not found in dictionary.")
                        Else
                            Dim npmTotalNetIncome As Decimal
                            If Not financialDataDict.TryGetValue("NetIncome", npmTotalNetIncome) Then
                                Debug.WriteLine("Net Income not found in dictionary.")
                            Else
                                ' Calculate Net Profit Margin
                                If npmTotalRevenue > 0 Then
                                    Dim netProfitMargin As Double = helperMethods.NetProfitMargin(npmTotalRevenue, npmTotalCostOfGoodsSold, npmTotalOperatingExpenses, npmTotalNetIncome)
                                    netProfitMargin = Math.Floor(netProfitMargin * 100) / 100
                                    NPM_api.Text = netProfitMargin.ToString() & "%"
                                    analysis_results("NPM_API") = netProfitMargin
                                Else
                                    Debug.WriteLine("Total Revenue is zero or less, cannot calculate NPM")
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            ' Check if CK_GrossProfitMargin_API is checked
            If CK_GrossProfitMargin_API.IsChecked Then
                ' Retrieve relevant values from dictionary
                Dim gpmTotalRevenue As Decimal
                If Not financialDataDict.TryGetValue("Revenue", gpmTotalRevenue) Then
                    Debug.WriteLine("Revenue not found in dictionary.")
                Else
                    Dim gpmTotalCostOfGoodsSold As Decimal
                    If Not financialDataDict.TryGetValue("CostOfGoodsSold", gpmTotalCostOfGoodsSold) Then
                    Else
                        ' Calculate Gross Profit Margin if Total Revenue is greater than 0
                        If gpmTotalRevenue > 0 Then
                            Dim grossProfitMargin As Double = helperMethods.GrossProfitMargin(gpmTotalRevenue, gpmTotalCostOfGoodsSold)
                            grossProfitMargin = Math.Floor(grossProfitMargin * 100) / 100
                            GPM_api.Text = grossProfitMargin.ToString() & "%"
                            analysis_results("GPM_API") = grossProfitMargin
                        Else
                            Debug.WriteLine("Total Revenue is zero or less, cannot calculate GPM")
                        End If
                    End If
                End If
            End If

            ' Check if CK_CurrentRatios_API is checked
            If CK_CurrentRatios_API.IsChecked Then
                ' Retrieve relevant values from dictionary
                Dim crTotalCurrentAssets As Decimal
                If Not financialDataDict.TryGetValue("CurrentAssets", crTotalCurrentAssets) Then
                Else
                    Dim crTotalCurrentLiabilities As Decimal
                    If Not financialDataDict.TryGetValue("CurrentLiabilities", crTotalCurrentLiabilities) Then
                    Else
                        ' Calculate Current Ratio if Total Current Liabilities is greater than 0
                        If crTotalCurrentLiabilities > 0 Then
                            Dim currentRatio As Double = helperMethods.CurrentRatio(crTotalCurrentAssets, crTotalCurrentLiabilities)
                            currentRatio = Math.Floor(currentRatio * 100) / 100
                            CRR_api.Text = currentRatio.ToString()
                            analysis_results("CR_API") = currentRatio
                        Else
                            Debug.WriteLine("Total Current Liabilities is zero or less, cannot calculate CR")
                        End If
                    End If
                End If
            End If

            ' Check if CK_DebtToEquity_API is checked
            If CK_DebtToEquity_API.IsChecked Then
                ' Retrieve relevant values from dictionary
                Dim dteTotalLiabilities As Decimal
                If Not financialDataDict.TryGetValue("TotalLiabilities", dteTotalLiabilities) Then
                    Debug.WriteLine("Total Liabilities not found in dictionary.")
                Else
                    Dim dteTotalEquity As Decimal
                    If Not financialDataDict.TryGetValue("TotalEquity", dteTotalEquity) Then
                        Debug.WriteLine("Total Equity not found in dictionary.")
                    Else
                        ' Calculate Debt to Equity Ratio if TotalEquity is greater than 0
                        If dteTotalEquity > 0 Then
                            Dim debtToEquityRatio As Double = helperMethods.DebtToEquityRatio(dteTotalLiabilities, dteTotalEquity)
                            Debug.WriteLine($"TotalLiabilities: {dteTotalLiabilities}, TotalEquity: {dteTotalEquity}, D/E: {debtToEquityRatio}")
                            debtToEquityRatio = Math.Floor(debtToEquityRatio * 100) / 100
                            DTE_api.Text = debtToEquityRatio.ToString()
                            analysis_results("DTE_API") = debtToEquityRatio
                        Else
                            Debug.WriteLine("Total Equity is zero or less, cannot calculate D/E")
                        End If
                    End If
                End If
            End If

            ' Check if CK_InterestCoverage_API is checked
            If CK_InterestCoverage_API.IsChecked Then
                ' Retrieve relevant values from dictionary
                Dim icrTotalEbitda As Decimal
                If Not financialDataDict.TryGetValue("EBITDA", icrTotalEbitda) Then
                    Debug.WriteLine("Total EBITDA not found in dictionary.")
                Else
                    Dim icrTotalInterestExpense As Decimal
                    If Not financialDataDict.TryGetValue("InterestExpense", icrTotalInterestExpense) Then
                        Debug.WriteLine("Total Interest Expense not found in dictionary.")
                    Else
                        ' Calculate Interest Coverage Ratio if TotalInterestExpense is greater than 0
                        If icrTotalInterestExpense > 0 Then
                            Dim interestCoverageRatio As Double = helperMethods.InterestCoverageRatio(icrTotalEbitda, icrTotalInterestExpense)
                            Debug.WriteLine($"TotalEbitda: {icrTotalEbitda}, TotalInterestExpense: {icrTotalInterestExpense}, ICR: {interestCoverageRatio}")
                            interestCoverageRatio = Math.Floor(interestCoverageRatio * 100) / 100
                            IC_api.Text = interestCoverageRatio.ToString()
                            analysis_results("ICR_API") = interestCoverageRatio
                        Else
                            Debug.WriteLine("Total Interest Expense is zero or less, cannot calculate ICR")
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"Error fetching or calculating financial ratios: {ex.Message}")
        End Try
        check_result_dict()
    End Sub

    'TEMP DELETE AFTER TEST
    Private Sub check_result_dict()
        For Each result In analysis_results
            Debug.WriteLine($"{result.Key}: {result.Value}")
        Next
    End Sub

    Private Sub ClickableText_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
        Debug.WriteLine("Clickable text was clicked!")
        authenatication_win = New Authentication()
        Dim result As Nullable(Of Boolean) = authenatication_win.ShowDialog()

        If result.HasValue AndAlso result.Value Then
            Debug.WriteLine("User authenticated successfully!")
            LogIn.Text = $"Welcome, {_userName}!"
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
                document.Info.Title = "Analysis Results"

                ' Create a page
                Dim page As PdfPage = document.AddPage()

                ' Get an XGraphics object for drawing
                Dim gfx As XGraphics = XGraphics.FromPdfPage(page)

                ' Define fonts
                Dim fontTitle As New XFont("Arial", 24)
                Dim fontText As New XFont("Arial", 12)

                ' Draw title
                gfx.DrawString("Analysis Results", fontTitle, XBrushes.Black, New XRect(0, 20, page.Width, 0), XStringFormats.TopCenter)

                ' Draw professional text
                Dim professionalText As String = "Here are the results of our analysis:"
                gfx.DrawString(professionalText, fontText, XBrushes.Black, New XRect(40, 60, page.Width - 80, 0), XStringFormats.TopLeft)

                ' Draw results from analysis_results dictionary
                Dim startY As Double = 100
                For Each kvp In analysis_results
                    Dim line = $"{kvp.Key}: {kvp.Value}"
                    gfx.DrawString(line, fontText, XBrushes.Black, New XRect(40, startY, page.Width - 80, 0), XStringFormats.TopLeft)
                    startY += 20
                Next

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


    Private Sub CopyMetricsToClipboard_Click(sender As Object, e As RoutedEventArgs)
        Dim metricsText As New StringBuilder()

        ' Iterate through ListView items (StackPanels in this case)
        For Each item As StackPanel In MetricsResults_List.Items
            ' Extract metric name and value from each StackPanel's children (TextBlocks)
            Dim metricName As String = TryCast(item.Children(0), TextBlock)?.Text
            Dim metricValue As String = TryCast((TryCast(item.Children(1), TextBlock))?.Text, String)

            ' Append metric name and value to StringBuilder
            metricsText.AppendLine($"{metricName}{metricValue}")
        Next

        ' Copy metricsText to clipboard
        Clipboard.SetText(metricsText.ToString())

        ' Optionally, show a message or perform other actions after copying
        MessageBox.Show("Metrics copied to clipboard!")
    End Sub


    Private Sub CopyMetricsToClipboardAPI_Click(sender As Object, e As RoutedEventArgs)
        Dim metricsText As New StringBuilder()

        ' Iterate through ListView items (StackPanels in this case)
        For Each item As StackPanel In MetricsResults_ListAPI.Items
            ' Extract metric name and value from each StackPanel's children (TextBlocks)
            Dim metricName As String = TryCast(item.Children(0), TextBlock)?.Text
            Dim metricValue As String = TryCast((TryCast(item.Children(1), TextBlock))?.Text, String)

            ' Append metric name and value to StringBuilder
            metricsText.AppendLine($"{metricName}{metricValue}")
        Next

        ' Copy metricsText to clipboard
        Clipboard.SetText(metricsText.ToString())

        ' Optionally, show a message or perform other actions after copying
        MessageBox.Show("Metrics copied to clipboard!")
    End Sub




End Class