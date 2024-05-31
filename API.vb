Imports System.Net.Http
Imports Newtonsoft.Json.Linq

Public Class API
    Private api_key As String
    Private company_symbol As String
    Private fiscal_year As String
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
        fiscal_year = ""
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

    Public Property FiscalYear() As String
        Get
            Return fiscal_year
        End Get
        Set(ByVal value As String)
            fiscal_year = value
            UpdateIncomeStatementUrl()
        End Set
    End Property

    Private Sub UpdateIncomeStatementUrl()
        ' Update the income statement URL based on the current symbol and API key
        income_statement_url = $"https://www.alphavantage.co/query?function=INCOME_STATEMENT&symbol={company_symbol}&apikey={api_key}"
    End Sub
    Public Async Function LoadIncomeStatement_API() As Task
        If String.IsNullOrEmpty(company_symbol) Then
            Throw New InvalidOperationException("Company symbol is not set.")
        End If

        Using client As New HttpClient()
            Dim income_statement_Json As String = Await client.GetStringAsync(income_statement_url)
            income_statement_Data = JObject.Parse(income_statement_Json)
            ' Filter the report by fiscal year if needed, otherwise get the first report
            If String.IsNullOrEmpty(fiscal_year) Then
                income_statement_report = CType(income_statement_Data("annualReports")(0), JObject)
            Else
                For Each report As JObject In income_statement_Data("annualReports")
                    If report("fiscalDateEnding").ToString().Contains(fiscal_year) Then
                        income_statement_report = report
                        Exit For
                    End If
                Next
            End If
        End Using
    End Function
End Class
