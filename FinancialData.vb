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
        Debug.WriteLine("Financial Data:")
        Debug.WriteLine($"Revenue: {Revenue}")
        Debug.WriteLine($"Cost of Goods Sold: {CostOfGoodsSold}")
        Debug.WriteLine($"Operating Expenses: {OperatingExpenses}")
        Debug.WriteLine($"Net Income: {NetIncome}")
        Debug.WriteLine($"Total Assets: {TotalAssets}")
        Debug.WriteLine($"Total Equity: {TotalEquity}")
        Debug.WriteLine($"EBITDA: {EBITDA}")
        Debug.WriteLine($"Current Assets: {CurrentAssets}")
        Debug.WriteLine($"Current Liabilities: {CurrentLiabilities}")
        Debug.WriteLine($"Total Liabilities: {TotalLiabilities}")
        Debug.WriteLine($"Interest Expense: {InterestExpense}")
    End Sub

    'TO DO: IMPLEMENT API FINANCIAL DATA FOR BALANCE SHEET
    Public Function GrossProfitMargin(ByVal revenue As Double, ByVal costOfGoodsSold As Double) As Double
        If revenue = 0 Then
            Return 0
        End If
        Return (revenue - costOfGoodsSold) / revenue
    End Function

    Public Function OperatingProfitMargin(ByVal revenue As Double, ByVal costOfGoodsSold As Double, ByVal operatingExpenses As Double) As Double
        If revenue = 0 Then
            Return 0
        End If
        Return (revenue - costOfGoodsSold - operatingExpenses) / revenue
    End Function

    Public Function NetProfitMargin(ByVal revenue As Double, ByVal costOfGoodsSold As Double, ByVal operatingExpenses As Double, ByVal netIncome As Double) As Double
        If revenue = 0 Then
            Return 0
        End If
        Return netIncome / revenue
    End Function

    Public Function ReturnOnAssets(ByVal netIncome As Double, ByVal totalAssets As Double) As Double
        If totalAssets = 0 Then
            Return 0
        End If
        Return netIncome / totalAssets
    End Function

    Public Function ReturnOnEquity(ByVal netIncome As Double, ByVal totalEquity As Double) As Double
        If totalEquity = 0 Then
            Return 0
        End If
        Return netIncome / totalEquity
    End Function

    Public Function CurrentRatio(ByVal currentAssets As Double, ByVal currentLiabilities As Double) As Double
        If currentLiabilities = 0 Then
            Return 0
        End If
        Return currentAssets / currentLiabilities
    End Function

    Public Function DebtToEquityRatio(ByVal totalLiabilities As Double, ByVal totalEquity As Double) As Double
        If totalEquity = 0 Then
            Return 0
        End If
        Return totalLiabilities / totalEquity
    End Function

    Public Function InterestCoverageRatio(ByVal ebitda As Double, ByVal interestExpense As Double) As Double
        If interestExpense = 0 Then
            Return 0
        End If
        Return ebitda / interestExpense
    End Function

    Public Function ContributionMarginRatio(ByVal salesRevenuePerUnit As Double, ByVal variableCostPerUnit As Double) As Double
        If salesRevenuePerUnit = 0 Then
            Return 0
        End If
        Return (salesRevenuePerUnit - variableCostPerUnit) / salesRevenuePerUnit
    End Function

    Public Function BreakEvenPoint(ByVal fixedCosts As Double, ByVal salesRevenuePerUnit As Double, ByVal variableCostPerUnit As Double) As Double
        If (salesRevenuePerUnit - variableCostPerUnit) > 0 Then
            Return fixedCosts / (salesRevenuePerUnit - variableCostPerUnit)
        Else
            Return 0
        End If
    End Function

End Class