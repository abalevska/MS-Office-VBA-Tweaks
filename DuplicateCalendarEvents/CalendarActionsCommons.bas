Attribute VB_Name = "CalendarActionsCommons"
  
Public Const SubjectPrefix = "C:"
Public Const CopiesCategory = "Automatic Copy"

Public Const CalendarFolderName = "Calendar"
Public Const DeletedItemsFolderName = "Deleted Items"

Sub CloneItem(ByVal Item As Object, ByRef DestinationFolder As Outlook.folder)
    Dim cAppt As AppointmentItem
    Dim moveCal As AppointmentItem
     
    On Error Resume Next
       
    ' Check if item exists *(2)
    Set cAppt = FindAppointment(Item.globalAppointmentID, DestinationFolder)
    
    If Not cAppt Is Nothing Then
        Call DeleteItemClone(Item, DestinationFolder)
    End If
    
    Set cAppt = Item.Copy ' using Copy and Move, instead of Items.Add because of recurring events *(1)
    Set cAppt = cAppt.Move(DestinationFolder)

    With cAppt
        .Subject = SubjectPrefix & Item.Subject
        .Body = Item.globalAppointmentID
        .Categories = CopiesCategory
        .ReminderSet = False
        .Save
    End With

End Sub

Sub DeleteItemClone(ByVal Item As Object, ByRef DestinationFolder As Outlook.folder)

    Dim cAppt As AppointmentItem
    Dim objAppointment As AppointmentItem
    
    On Error Resume Next
    
    Set cAppt = FindAppointment(Item.globalAppointmentID, DestinationFolder)
    cAppt.Delete ' we assume it is just one item - should we?

End Sub

Function FindAppointment(ByVal globalAppointmentID As String, ByVal DestinationFolder As Outlook.folder) As AppointmentItem
    
    Dim filter As String
    filter = ItemsFilter()
    
    Dim FilteredItems
    Set FilteredItems = DestinationFolder.Items.Restrict(filter) ' performance optimization *(3) go to ReadMe module for more details;

    For Each objAppointment In FilteredItems
        If InStr(1, objAppointment.Body, globalAppointmentID) Then
            Set FindAppointment = objAppointment
            Exit Function
        End If
    Next
End Function

Function LastMonday(pdat As Date) As Date
    LastMonday = DateAdd("ww", -1, pdat - (Weekday(pdat, vbMonday) - 1))
End Function

Function ItemsFilter() As String
   Dim FilterCategory
   FilterCategory = "[Categories] = " & CopiesCategory
   Dim FilterDate
   FilterDate = "[Start] >= '" & Format(LastMonday(Now), "ddddd h:nn AMPM") & "'"
   
   ItemsFilter = FilterCategory & " And " & FilterDate
End Function

Function FilterItemsCategoryCopied() As String
    FilterItemsCategoryCopied = "[Categories] = '" & CopiesCategory & "'"
End Function

Function FilterItemsCategoryNotCopied() As String
    FilterItemsCategoryNotCopied = "[Categories] <> '" & CopiesCategory & "'"
End Function
