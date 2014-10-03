Attribute VB_Name = "ThisOutlookSession1"
Dim olApp As New Outlook.Application
Dim olNameSpace As Outlook.NameSpace
Dim oFolder As Outlook.MAPIFolder
Dim oMail As Outlook.MailItem

Sub Main()
    Set olNameSpace = olApp.GetNamespace("MAPI")
    Set oFolder = olNameSpace.Folders("dvelezschrod@proofpoint.com").Folders("Inbox")
    ManuallyMoveOfficeGossip
End Sub


Sub moveOfficeGossip(item As Outlook.MailItem)
    Dim strNames As String, i As Integer, j As Integer
    Dim olDestFolder As Outlook.MAPIFolder
    
    j = 1
    strNames = item.To
    
    Set olNameSpace = olApp.GetNamespace("MAPI")
    
    For i = 1 To Len(strNames)
        If Mid(strNames, i, 1) = ";" Then
        j = j + 1
        End If
    Next i
    
    If (j >= 2) Then
        Set olDestFolder = olNameSpace.Folders("dvelezschrod@proofpoint.com").Folders("Deleted Items").Folders("Cc")
        item.Move olDestFolder
    End If
End Sub


Sub ManuallyMoveOfficeGossip()
 
 Dim item As Outlook.MailItem
 Dim items As Integer
 Dim processed_items As Integer
 
 items = 0
 processed_items = 0
 
 Set olNameSpace = Application.GetNamespace("MAPI")
 Set oFolder = olNameSpace.Folders("dvelezschrod@proofpoint.com").Folders("Inbox")
 
 items = oFolder.items.Count
 
 While processed_items < items
    For Each item In oFolder.items
       If InStr(LCase(item.To), "velezschrod") Then
           processed_items = processed_items + 1
           ' MsgBox item
           ' item.Display
           moveOfficeGossip item
       End If
    Next
 Wend
 
 ' MsgBox items & " have been moved to Cc."
 
End Sub
