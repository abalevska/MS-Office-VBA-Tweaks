Attribute VB_Name = "CleanUpDeletedByCategory"
Sub CleanUpCopiedItemsFromDeleted()
    
    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")
   
    Dim folder As Outlook.folder
    Dim FilteredItems
    
    Dim Accounts
    Accounts = CalendarAccountsConstants.Accounts
        
    For Each Account In Accounts
        Set folder = NS.Folders(Account).Folders(CalendarActionsCommons.DeletedItemsFolderName)
        Set FilteredItems = folder.Items.Restrict(CalendarActionsCommons.FilterItemsCategoryCopied)
    
        For Each objAppointment In FilteredItems
            objAppointment.Delete
        Next
    Next
    
    Set NS = Nothing
End Sub

