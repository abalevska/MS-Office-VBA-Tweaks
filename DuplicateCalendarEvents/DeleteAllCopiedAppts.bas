Attribute VB_Name = "DeleteAllCopiedAppts"
Sub DeleteAllCopiedAppointmentItems()
    
    Dim Accounts
    Accounts = CalendarAccountsConstants.Accounts
    
    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")
    
    Dim folder As Outlook.folder
    Dim FilteredItems
        
    For Each Account In Accounts
        Set folder = NS.Folders(Account).Folders(CalendarActionsCommons.CalendarFolderName)
        Set FilteredItems = folder.Items.Restrict(CalendarActionsCommons.FilterItemsCategoryCopied)
    
        For Each objAppointment In FilteredItems
            objAppointment.Delete
        Next
        
        Set folder = NS.Folders(Account).Folders(CalendarActionsCommons.DeletedItemsFolderName)
        Set FilteredItems = folder.Items.Restrict(CalendarActionsCommons.FilterItemsCategoryCopied)
        
        For Each objAppointment In FilteredItems
            objAppointment.Delete
        Next
        
    Next
    
    Set NS = Nothing
End Sub

