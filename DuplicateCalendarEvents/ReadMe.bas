Attribute VB_Name = "ReadMe"


' (1)
' Using copy and move instead of Items.Add and then setting propeprites
' Reason 1: There are too many variables to think of and there is a chance something will get missed
' Reason 2: recurring events. Setting recurrence properties via code does not work OK, but works great when event is copied
'           This also handles then case when a specific ocurrence is modified, which cannot be handled in the other case
' (1.1)
' When Item is copied event is triggered for addition, that becomes invalid after item is moved to the proper calendar
' The moved item becomes invalid and properties cannot be read, so "Item.Categories" throws an exception
' On Error GoTo end of sub, handles that exception
' (1.2)
' Follow up the above mentioned reason
' Reason 1: There are too many variables to think of and there is a chance something will get missed
' When Item is changes it is deleted and then copied.
' This is the only way to handle modification of occurence of a recurring appointment
' (1.3)
' Deleting an item on every change sends it to the Deleted items folder and this introduces overload of leftovers, so we clean them up

' (2)
' This is needed, because when an invitation is received the item is added to the calendar,
' however when meeting is accepted calendarItemsDefault_ItemAdd is triggered once more

' (3)
' takes a lot of time to go trough all events, so:
'   * we disregard any events older then a week ago
'   * and also only loop through the proper categoy

