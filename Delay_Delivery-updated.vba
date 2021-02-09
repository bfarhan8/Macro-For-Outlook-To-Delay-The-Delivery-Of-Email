Dim obj As Object
Dim Mail As Outlook.MailItem
Dim MinNow As Integer
Dim SendHour As Integer
Dim SendDate As Date
Dim SendNow As String
Dim UserDeferOption As Integer
Dim WkDay As Integer

Function getActiveMessage() As Outlook.MailItem
Dim insp As Outlook.Inspector
If TypeOf Application.ActiveWindow Is Outlook.Inspector Then
        Set insp = Application.ActiveWindow
    End If
If insp Is Nothing Then
        Dim inline As Object
        Set inline = Application.ActiveExplorer.ActiveInlineResponse
        If inline Is Nothing Then Exit Function
Set getActiveMessage = inline
    Else
       Set insp = Application.ActiveInspector
       If insp.CurrentItem.Class = olMail Then
          Set getActiveMessage = insp.CurrentItem
       Else
         Exit Function
       End If
End If
End Function
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'On Error GoTo ErrorHandler
'This sub used to delay the sending of an email from send time to the next work day at 9am.
'Set Variables
SendDate = Now()
SendHour = Hour(Now)
MinNow = Minute(Now)
WkDay = Weekday(Now)
SendNow = "Y"


If WkDay <> 1 And WkDay <> 7 Then

'Check if Before 7 PM
If SendHour > 9 And SendHour < 19 Then
SendHour = 9 - SendHour
SendDate = DateAdd("h", SendHour, SendDate)
SendDate = DateAdd("n", -MinNow, SendDate)
SendNow = "Y"
End If

'Check if after 7PM
If SendHour >= 19 Then 'After 7 PM
SendHour = 33 - SendHour 'Send a 9 am next day
SendDate = DateAdd("h", SendHour, SendDate)
SendDate = DateAdd("n", -MinNow, SendDate)
SendNow = "N"
End If


'Check if after 12 AM
If SendHour >= 0 And SendHour < 9 Then 'After 12 AM
SendHour = 9 - SendHour 'Send a 9 am next day
SendDate = DateAdd("h", SendHour, SendDate)
SendDate = DateAdd("n", -MinNow, SendDate)
SendNow = "N"
End If
End If

'Check if after 7PM
If WkDay = 1 And SendHour >= 0 Then
SendHour = 33 - SendHour 'Send a 9 am next day
SendDate = DateAdd("h", SendHour, SendDate)
SendDate = DateAdd("n", -MinNow, SendDate)
SendNow = "N"
End If

'Check if Saturday
If WkDay = 7 And SendHour > 19 Then
SendDate = Now()
SendHour = Hour(Now)
SendDate = DateAdd("d", 2, SendDate)
SendDate = DateAdd("h", 9 - SendHour, SendDate)
SendDate = DateAdd("n", -MinNow, SendDate)
SendNow = "N"
End If

'Send the Email
Set obj = getActiveMessage()
If obj Is Nothing Then
'Do nothing - as this is likely a calendar issue
'MsgBox "No active inspector"
Else
If TypeOf obj Is Outlook.MailItem Then
Set Mail = obj
'Check if we need to delay delivery
If SendNow = "N" Then
UserDeferOption = MsgBox("Do you want to postpone sending until work hours (" & SendDate & ")?", vbYesNo + vbQuestion, "Time to stop working!")
If UserDeferOption = vbYes Then
Mail.DeferredDeliveryTime = SendDate
'MsgBox ("Your mail will be sent at: " & SendDate)
Else
End If
End If
End If
End If
Exit Sub
'ErrorHandler:
' MsgBox "Error!"
End Sub
