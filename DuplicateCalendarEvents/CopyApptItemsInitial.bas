Attribute VB_Name = "CopyApptItemsInitial"
   
Dim calendarFolderDefault As Outlook.Folder
Dim calendarFolderSecondary As Outlook.Folder


Sub CopyApptItemsInitial()
    
    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")
    
    Set calendarFolderSecondary = NS.Folders(CalendarAccountsConstants.SecondaryEmail).Folders(CalendarActionsCommons.CalendarFolderName)
    Set calendarFolderDefault = NS.Folders(CalendarAccountsConstants.DefaultEmail).Folders(CalendarActionsCommons.CalendarFolderName)
    
    Call CloneCalendar(calendarFolderSecondary, calendarFolderDefault)
    Call CloneCalendar(calendarFolderDefault, calendarFolderSecondary)
    
    Set NS = Nothing
End Sub

Sub CloneCalendar(ByRef sourceFolder As Outlook.Folder, ByRef destinationFolder As Outlook.Folder)
    Set FilteredItems = sourceFolder.Items.Restrict(CalendarActionsCommons.FilterItemsCategoryNotCopied)
        
    For Each objAppointment In FilteredItems
        Call CalendarActionsCommons.CloneItem(objAppointment, destinationFolder)
    Next
End Sub
