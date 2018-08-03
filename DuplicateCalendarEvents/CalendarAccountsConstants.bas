Attribute VB_Name = "CalendarAccountsConstants"

Public Const DefaultEmail = ""
Public Const SecondaryEmail = ""

Function Accounts() As String()

    Dim returnVal(1) As String
    returnVal(0) = ""
    returnVal(1) = ""
    
    Accounts = returnVal
End Function

