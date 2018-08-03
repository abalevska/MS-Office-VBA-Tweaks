VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim calendarFolderDefault As Outlook.folder
Dim calendarFolderSecondary As Outlook.folder

Dim WithEvents calendarItemsDefault As Items
Attribute calendarItemsDefault.VB_VarHelpID = -1
Dim WithEvents calendarItemsSeconday As Items
Attribute calendarItemsSeconday.VB_VarHelpID = -1

Dim WithEvents deletedItemsDefault As Items
Attribute deletedItemsDefault.VB_VarHelpID = -1
Dim WithEvents deletedItemsSecondary As Items
Attribute deletedItemsSecondary.VB_VarHelpID = -1

Private Sub Application_Startup()
    
    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")
    
    Set calendarFolderDefault = NS.Folders(CalendarAccountsConstants.DefaultEmail).Folders(CalendarActionsCommons.CalendarFolderName)
    Set calendarFolderSecondary = NS.Folders(CalendarAccountsConstants.SecondaryEmail).Folders(CalendarActionsCommons.CalendarFolderName)
    
    Set calendarItemsDefault = calendarFolderDefault.Items
    Set calendarItemsSeconday = calendarFolderSecondary.Items
    
    Set deletedItemsDefault = NS.Folders(CalendarAccountsConstants.DefaultEmail).Folders(CalendarActionsCommons.DeletedItemsFolderName).Items
    Set deletedItemsSecondary = NS.Folders(CalendarAccountsConstants.SecondaryEmail).Folders(CalendarActionsCommons.DeletedItemsFolderName).Items
        
    Set NS = Nothing
End Sub
  
Private Sub calendarItemsDefault_ItemAdd(ByVal Item As Object)
    Call ItemAdd(Item, calendarFolderSecondary)
End Sub

Private Sub calendarItemsSeconday_ItemAdd(ByVal Item As Object)
    Call ItemAdd(Item, calendarFolderDefault)
End Sub

Private Sub calendarItemsDefault_ItemChange(ByVal Item As Object)
    Call ItemChange(Item, calendarFolderSecondary)
End Sub

Private Sub calendarItemsSeconday_ItemChange(ByVal Item As Object)
    Call ItemChange(Item, calendarFolderDefault)
End Sub

Private Sub deletedItemsDefault_ItemAdd(ByVal Item As Object)
    Call ItemDelete(Item, calendarFolderSecondary)
End Sub

Private Sub deletedItemsSecondary_ItemAdd(ByVal Item As Object)
    Call ItemDelete(Item, calendarFolderDefault)
End Sub

Private Sub ItemAdd(ByVal Item As Object, ByVal DestinationFolder As Outlook.folder)
    On Error GoTo EndSub_ItemAdd ' ugly but it works *(1.1) go to ReadMe module for more details;
    If Item.Categories = CalendarActionsCommons.CopiesCategory Then Exit Sub
    Call CalendarActionsCommons.CloneItem(Item, DestinationFolder)
EndSub_ItemAdd:
End Sub

Private Sub ItemChange(ByVal Item As Object, ByVal DestinationFolder As Outlook.folder)
    On Error GoTo EndSub_ItemChange ' ugly but it works *(1.1) go to ReadMe module for more details;
    If Item.Categories = CalendarActionsCommons.CopiesCategory Then Exit Sub
    Call CalendarActionsCommons.DeleteItemClone(Item, DestinationFolder) ' delete and clone again *(1.2) go to ReadMe module for more details;
    Call CalendarActionsCommons.CloneItem(Item, DestinationFolder)
EndSub_ItemChange:
End Sub

Private Sub ItemDelete(ByVal Item As Object, ByVal DestinationFolder As Outlook.folder)
    On Error Resume Next
    If Item.MessageClass <> "IPM.Appointment" Then Exit Sub ' only apply to appointments
    If Item.Categories = CalendarActionsCommons.CopiesCategory Then
        Item.Delete ' Clean up deleted items forlder *(1.3) go to ReadMe module for more details
        Exit Sub
    End If
    Call CalendarActionsCommons.DeleteItemClone(Item, DestinationFolder)
End Sub

