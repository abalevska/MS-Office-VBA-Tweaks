
Attribute VB_Name = "CalendarAccountsConstants"

Public Const DefaultEmail = "your.email@yourorganization.com"
Public Const SecondaryEmail = "your.email@clientorganization.com"

Function Accounts() As String()

    Dim returnVal(1) As String
    returnVal(0) = "your.email@yourorganization.com"
    returnVal(1) = "your.email@yourorganization.com"
    
    Accounts = returnVal
End Function