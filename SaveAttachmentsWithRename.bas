Attribute VB_Name = "SaveAttachmentsModule"
Option Explicit

Sub SaveAttachmentsWithUserRename()
    Dim objSelection As Selection
    Dim objMail As MailItem
    Dim objAttachment As Attachment
    Dim newName As String
    Dim savePath As String
    
    ' Change this folder path to wherever you want to save
    savePath = "C:\Attachments\"
    
    Set objSelection = Application.ActiveExplorer.Selection
    
    If objSelection.Count = 0 Then
        MsgBox "No items selected."
        Exit Sub
    End If
    
    On Error Resume Next
    
    ' Loop through the selected items in Outlook
    Dim i As Long
    For i = 1 To objSelection.Count
        If TypeName(objSelection(i)) = "MailItem" Then
            Set objMail = objSelection(i)
            ' Loop through each attachment in the mail
            For Each objAttachment In objMail.Attachments
                ' Prompt for a custom filename (default is the original filename)
                newName = InputBox("Enter a new name for the attachment:", "Rename Attachment", objAttachment.FileName)
                
                ' If user cancels or leaves blank, revert to original filename
                If Trim(newName) = "" Then
                    newName = objAttachment.FileName
                End If
                
                ' Make sure the extension isn't lost, if you want to preserve original extension
                ' If user leaves out the extension, you can automatically append it:
                If InStr(1, newName, ".") = 0 Then
                    newName = newName & "." & Split(objAttachment.FileName, ".")(UBound(Split(objAttachment.FileName, ".")))
                End If

                objAttachment.SaveAsFile savePath & newName
            Next
        End If
    Next i
    
    MsgBox "Attachments saved to " & savePath
End Sub
