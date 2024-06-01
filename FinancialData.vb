Imports Newtonsoft.Json.Linq

Public Class FinancialData
    Public Property FinancialDataID As Integer
    Public Property DateValue As Date
    Public Property Revenue As Decimal
    Public Property CostOfGoodsSold As Decimal
    Public Property OperatingExpenses As Decimal
    Public Property NetIncome As Decimal
    Public Property TotalAssets As Decimal
    Public Property TotalEquity As Decimal
    Public Property EBITDA As Decimal
    Public Property CurrentAssets As Decimal
    Public Property CurrentLiabilities As Decimal
    Public Property TotalLiabilities As Decimal
    Public Property InterestExpense As Decimal
    Public Property VariableCosts As Decimal
    Public Property FixedCosts As Decimal
    Public Property SalesRevenuePerUnit As Decimal
    Public Property VariableCostPerUnit As Decimal

    Private AV_API As API

    Public Sub New()
        AV_API = New API()
    End Sub

    Public Sub New(financialDataID As Integer, dateValue As Date, revenue As Decimal, costOfGoodsSold As Decimal, operatingExpenses As Decimal, netIncome As Decimal, totalAssets As Decimal, totalEquity As Decimal, ebitda As Decimal, currentAssets As Decimal, currentLiabilities As Decimal, totalLiabilities As Decimal, interestExpense As Decimal, variableCosts As Decimal, fixedCosts As Decimal, salesRevenuePerUnit As Decimal, variableCostPerUnit As Decimal)
        Me.FinancialDataID = financialDataID
        Me.DateValue = dateValue
        Me.Revenue = revenue
        Me.CostOfGoodsSold = costOfGoodsSold
        Me.OperatingExpenses = operatingExpenses
        Me.NetIncome = netIncome
        Me.TotalAssets = totalAssets
        Me.TotalEquity = totalEquity
        Me.EBITDA = ebitda
        Me.CurrentAssets = currentAssets
        Me.CurrentLiabilities = currentLiabilities
        Me.TotalLiabilities = totalLiabilities
        Me.InterestExpense = interestExpense
        Me.VariableCosts = variableCosts
        Me.FixedCosts = fixedCosts
        Me.SalesRevenuePerUnit = salesRevenuePerUnit
        Me.VariableCostPerUnit = variableCostPerUnit
    End Sub

    Public Async Function LoadFinancialData(symbol As String, selectedFiscalYearIndex As Integer) As Task
        AV_API.Symbol = symbol
        Await AV_API.LoadIncomeStatement_API(selectedFiscalYearIndex)

        ' Populate FinancialData properties from the API data
        Dim report As JObject = AV_API.IncomeReport

        If report IsNot Nothing Then
            Revenue = If(report("totalRevenue") IsNot Nothing, CDec(report("totalRevenue")), 0D)
            CostOfGoodsSold = If(report("costOfRevenue") IsNot Nothing, CDec(report("costOfRevenue")), 0D)
            OperatingExpenses = If(report("operatingExpenses") IsNot Nothing, CDec(report("operatingExpenses")), 0D)
            NetIncome = If(report("netIncome") IsNot Nothing, CDec(report("netIncome")), 0D)
            TotalAssets = If(report("totalAssets") IsNot Nothing, CDec(report("totalAssets")), 0D)
            TotalEquity = If(report("totalShareholderEquity") IsNot Nothing, CDec(report("totalShareholderEquity")), 0D)
            EBITDA = If(report("ebitda") IsNot Nothing, CDec(report("ebitda")), 0D)
            CurrentAssets = If(report("currentAssets") IsNot Nothing, CDec(report("currentAssets")), 0D)
            CurrentLiabilities = If(report("currentLiabilities") IsNot Nothing, CDec(report("currentLiabilities")), 0D)
            TotalLiabilities = If(report("totalLiabilities") IsNot Nothing, CDec(report("totalLiabilities")), 0D)
            InterestExpense = If(report("interestExpense") IsNot Nothing, CDec(report("interestExpense")), 0D)
            ' Other fields can be populated similarly
        End If
    End Function

    ' DELETET THIS WHEN DONE WITH TESTING
    Public Sub PrintFinancialData()
        Console.WriteLine("Financial Data:")
        Console.WriteLine($"Revenue: {Revenue}")
        Console.WriteLine($"Cost of Goods Sold: {CostOfGoodsSold}")
        Console.WriteLine($"Operating Expenses: {OperatingExpenses}")
        Console.WriteLine($"Net Income: {NetIncome}")
        Console.WriteLine($"Total Assets: {TotalAssets}")
        Console.WriteLine($"Total Equity: {TotalEquity}")
        Console.WriteLine($"EBITDA: {EBITDA}")
        Console.WriteLine($"Current Assets: {CurrentAssets}")
        Console.WriteLine($"Current Liabilities: {CurrentLiabilities}")
        Console.WriteLine($"Total Liabilities: {TotalLiabilities}")
        Console.WriteLine($"Interest Expense: {InterestExpense}")
    End Sub

    Public Function GrossProfitMargin() As Decimal
        Return (Revenue - CostOfGoodsSold) / Revenue
    End Function

    Public Function OperatingProfitMargin() As Decimal
        Return (Revenue - CostOfGoodsSold - OperatingExpenses) / Revenue
    End Function

    Public Function NetProfitMargin() As Decimal
        Return NetIncome / Revenue
    End Function

    Public Function ReturnOnAssets() As Decimal
        Return NetIncome / TotalAssets
    End Function

    Public Function ReturnOnEquity() As Decimal
        Return NetIncome / TotalEquity
    End Function

    Public Function CurrentRatio() As Decimal
        Return CurrentAssets / CurrentLiabilities
    End Function

    Public Function DebtToEquityRatio() As Decimal
        Return TotalLiabilities / TotalEquity
    End Function
    Public Function InterestCoverageRatio() As Decimal
        Return EBITDA / InterestExpense
    End Function
    Public Function ContributionMargin() As Decimal
        Return (SalesRevenuePerUnit - VariableCostPerUnit) / SalesRevenuePerUnit
    End Function

    Public Function BreakEvenPoint() As Decimal
        If (SalesRevenuePerUnit - VariableCostPerUnit) > 0 Then
            Return FixedCosts / (SalesRevenuePerUnit - VariableCostPerUnit)
        Else
            Return 0
        End If
    End Function
End Class