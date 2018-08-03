Attribute VB_Name = "DeleteAllCopiedAppts"
Sub DeleteAllCopiedAppointmentItems()
    
    Dim Accounts
    Accounts = CalendarAccountsConstants.Accounts
    
    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")
    
    Dim Folder As Outlook.Folder
    Dim FilteredItems
        
    For Each Account In Accounts
        Set Folder = NS.Folders(Account).Folders(CalendarActionsCommons.CalendarFolderName)
        Set FilteredItems = Folder.Items.Restrict(CalendarActionsCommons.FilterItemsCategoryCopied)
    
        For Each objAppointment In FilteredItems
            objAppointment.Delete
        Next
        
        Set Folder = NS.Folders(Account).Folders(CalendarActionsCommons.DeletedItemsFolderName)
        Set FilteredItems = Folder.Items.Restrict(CalendarActionsCommons.FilterItemsCategoryCopied)
        
        For Each objAppointment In FilteredItems
            objAppointment.Delete
        Next
        
    Next
    
    Set NS = Nothing
End Sub

