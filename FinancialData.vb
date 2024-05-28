﻿Public Class FinancialData
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

    Public Sub New()
        ' Default constructor
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

    'TO DO: Implement methods to calculate the respective metrics
End Class