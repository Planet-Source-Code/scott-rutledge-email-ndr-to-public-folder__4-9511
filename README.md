<div align="center">

## Email NDR to Public Folder


</div>

### Description

Run as a VBS from a scheduled task, logged in as the desired user, this script will move non-delivery reports (NDR) email to a public folder.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Scott Rutledge](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/scott-rutledge.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/scott-rutledge-email-ndr-to-public-folder__4-9511/archive/master.zip)





### Source Code

```
CONST strServer   = "SERVER"
  CONST strMailbox   = "MAILBOX"
  Dim objSession
  Dim objMessages
  Dim objOneMessage
  Dim objInfoStores
  Dim objInfoStore
  Dim objTopFolder
  Dim objFolders
  Dim objInbox
  Dim objSubFolder
  Dim objTargetFolder
  Dim strProfileInfo
  Dim bstrPublicRootID
  Dim i
  strProfileInfo = strServer & vblf & strMailbox
  Set objSession = CreateObject("MAPI.Session")
  objSession.Logon , , False, , , True, strProfileInfo
  Set objInfoStores = objSession.InfoStores
  For i = 1 To objInfoStores.Count
  	If objInfoStores.Item(i)= "Public Folders" Then
    		Set objInfoStore=objInfoStores.Item(i)
    		Exit For
  	End If
  Next
  bstrPublicRootID = objInfoStore.Fields.Item( &H66310102 ).Value
  Set objTopFolder = objSession.GetFolder(bstrPublicRootID, _
        objInfoStore.ID)
  Set objFolders = objTopFolder.Folders
  Set objFolder = objFolders.GetFirst()
  i = 0
  Do Until objFolder.Name = "Public Folder Name"
  	i = i + 1
  	If i > 100 Then 'kill the search
 		Exit Do
  	End If
  	Set objFolder=objFolders.GetNext()
  Loop
  For i = 1 to 3 '3 passes enough to grab everything
  	Set objInbox = objSession.Inbox
  	Set objMessages = objInbox.Messages
  	For Each objOneMessage in objMessages
   		If objOneMessage.Type = "REPORT.IPM.Note.NDR" Then
  			Set objCopyMsg = objOneMessage.MoveTo(objFolder.ID)
		End If
  	Next
  Next
  objSession.Logoff
  Set objOneMessage = Nothing
  Set objMessages = Nothing
  Set objFolder = Nothing
  Set objTopFolder = Nothing
  Set objSession = Nothing
```

