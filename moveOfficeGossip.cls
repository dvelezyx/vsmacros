VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Main()
    'GetFolderNames
End Sub

Sub moveOfficeGossip(item As Outlook.MailItem)

    Dim strNames As String, i As Integer, j As Integer
    Dim olApp As New Outlook.Application
    Dim olNameSpace As Outlook.NameSpace
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
        Set olDestFolder = olNameSpace.Folders("Buzón: Daniel Velez Schrod").Folders("Elementos eliminados").Folders("Cc")
        item.Move olDestFolder
    End If

End Sub
 
