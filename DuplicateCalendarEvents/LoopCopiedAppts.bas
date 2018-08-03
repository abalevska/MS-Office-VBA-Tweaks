Attribute VB_Name = "LoopCopiedAppts"
Sub LoopsCopiedAppointmentItems()
    
    Dim Accounts
    Accounts = Array("avasileva@objectsystems.com", "avvasileva.cw@mmm.com")
    
    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")
    
    Dim CalendarFolder
    CalendarFolder = "Calendar"
    
    Dim filter As String
    filter = "[Categories] = 'Automatic Copy'"
    
    Dim folder As Outlook.folder
    Dim FilteredItems
        
    For Each Account In Accounts
        Set folder = NS.Folders(Account).Folders(CalendarFolder)
        Set FilteredItems = folder.Items.Restrict(filter)
    
        For Each objAppointment In FilteredItems
            ' add whatever you want to do here
            ' RemoveReminders (objAppointment)
        Next
    Next
    
    Set NS = Nothing
End Sub

Sub RemoveReminders(ByVal objAppointment As Outlook.AppointmentItem)
    If objAppointment.ReminderSet = True Then
       objAppointment.ReminderSet = False
       objAppointment.Save
    End If
End Sub

Sub AddReminders(ByVal objAppointment As Outlook.AppointmentItem)
    If objAppointment.ReminderSet = False Then
       objAppointment.ReminderSet = True
       objAppointment.Save
    End If
End Sub
