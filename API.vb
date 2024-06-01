Imports System.Net.Http
Imports Newtonsoft.Json.Linq

Public Class API
    Private api_key As String
    Private company_symbol As String
    Private fiscal_year As Integer
    Private income_statement_url As String
    Private income_statement_Data As JObject
    Private income_statement_report As JObject
    Private result As String
    Private balancesheet_url As String
    Private balancesheet_Data As JObject
    Private balancesheet_report As JObject

    'Initialize the API class variables with a default constructor
    Public Sub New()
        'Retrieve API key from configuration file KN8N1PLOV3JJ8TCB.
        api_key = "KN8N1PLOV3JJ8TCB"
        income_statement_Data = New JObject()
        income_statement_report = New JObject()
        balancesheet_Data = New JObject()
        balancesheet_report = New JObject()
        result = ""
        fiscal_year = 0
        company_symbol = ""
    End Sub

    Public Property Symbol() As String
        Get
            Return company_symbol
        End Get
        Set(ByVal value As String)
            company_symbol = value
            UpdateIncomeStatementUrl()
        End Set
    End Property

    Public Property FiscalYear() As Integer
        Get
            Return fiscal_year
        End Get
        Set(ByVal value As Integer)
            fiscal_year = value
            UpdateIncomeStatementUrl()
        End Set
    End Property

    Public ReadOnly Property IncomeReport() As JObject
        Get
            Return income_statement_report
        End Get
    End Property

    ' Update the income statement URL based on the current symbol and API key
    Private Sub UpdateIncomeStatementUrl()
        income_statement_url = $"https://www.alphavantage.co/query?function=INCOME_STATEMENT&symbol={company_symbol}&apikey={api_key}"
    End Sub

    'Using the AV API to get the income statement data
    Public Async Function LoadIncomeStatement_API(selectedFiscalYearIndex As Integer) As Task
        If String.IsNullOrEmpty(company_symbol) Then
            Throw New InvalidOperationException("Company symbol is not set.")
        End If

        Using client As New HttpClient()
            Dim income_statement_Json As String = Await client.GetStringAsync(income_statement_url)
            income_statement_Data = JObject.Parse(income_statement_Json)

            ' Select the report based on the selected index
            Dim reports As JArray = CType(income_statement_Data("annualReports"), JArray)

            If selectedFiscalYearIndex >= 0 AndAlso selectedFiscalYearIndex < reports.Count Then
                income_statement_report = CType(reports(selectedFiscalYearIndex), JObject)
            Else
                Throw New IndexOutOfRangeException("Selected fiscal year index is out of range.")
            End If
        End Using
    End Function



    'Using the AV API to get the balance sheet data
    Public Async Function LoadBalanceSheet_API() As Task
        If String.IsNullOrEmpty(company_symbol) Then
            Throw New InvalidOperationException("Company symbol is not set.")
        End If

        balancesheet_url = $"https://www.alphavantage.co/query?function=BALANCE_SHEET&symbol={company_symbol}&apikey={api_key}"
        Using client As New HttpClient()
            Dim balancesheet_Json As String = Await client.GetStringAsync(balancesheet_url)
            balancesheet_Data = JObject.Parse(balancesheet_Json)
            ' Filter the report by fiscal year if needed, otherwise get the first report
            If String.IsNullOrEmpty(fiscal_year) Then
                balancesheet_report = CType(balancesheet_Data("annualReports")(0), JObject)
            Else
                For Each report As JObject In balancesheet_Data("annualReports")
                    If report("fiscalDateEnding").ToString().Contains(fiscal_year) Then
                        balancesheet_report = report
                        Exit For
                    End If
                Next
            End If
        End Using
    End Function
End Class
