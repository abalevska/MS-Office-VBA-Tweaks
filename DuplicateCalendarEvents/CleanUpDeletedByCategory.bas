Attribute VB_Name = "CleanUpDeletedByCategory"
Sub CleanUpCopiedItemsFromDeleted()
    
    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")
   
    Dim Folder As Outlook.Folder
    Dim FilteredItems
    
    Dim Accounts
    Accounts = CalendarAccountsConstants.Accounts
        
    For Each Account In Accounts
        Set Folder = NS.Folders(Account).Folders(CalendarActionsCommons.DeletedItemsFolderName)
        Set FilteredItems = Folder.Items.Restrict(CalendarActionsCommons.FilterItemsCategoryCopied)
    
        For Each objAppointment In FilteredItems
            objAppointment.Delete
        Next
    Next
    
    Set NS = Nothing
End Sub

