Attribute VB_Name = "CalendarAccountsConstants"

Public Const DefaultEmail = "avasileva@objectsystems.com"
Public Const SecondaryEmail = "avvasileva.cw@mmm.com"

Function Accounts() As String()

    Dim returnVal(1) As String
    returnVal(0) = "avasileva@objectsystems.com"
    returnVal(1) = "avvasileva.cw@mmm.com"
    
    Accounts = returnVal
End Function

