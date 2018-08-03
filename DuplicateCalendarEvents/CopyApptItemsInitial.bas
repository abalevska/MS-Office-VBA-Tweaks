Attribute VB_Name = "CopyApptItemsInitial"
   
Dim calendarFolderDefault As Outlook.folder
Dim calendarFolderSecondary As Outlook.folder


Sub CopyApptItemsInitial()
    
    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")
    
    Set calendarFolderSecondary = NS.Folders(CalendarAccountsConstants.SecondaryEmail).Folders(CalendarActionsCommons.CalendarFolderName)
    Set calendarFolderDefault = NS.Folders(CalendarAccountsConstants.DefaultEmail).Folders(CalendarActionsCommons.CalendarFolderName)
    
    Call CloneCalendar(calendarFolderSecondary, calendarFolderDefault)
    Call CloneCalendar(calendarFolderDefault, calendarFolderSecondary)
    
    Set NS = Nothing
End Sub

Sub CloneCalendar(ByRef sourceFolder As Outlook.folder, ByRef DestinationFolder As Outlook.folder)
    Set FilteredItems = sourceFolder.Items.Restrict(CalendarActionsCommons.FilterItemsCategoryNotCopiedAndCurrent)
        
    For Each objAppointment In FilteredItems
        Call CalendarActionsCommons.CloneItem(objAppointment, DestinationFolder)
    Next
End Sub
