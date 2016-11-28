# HonoursAutoReply
Outlook 2010 Addin

Compiled binaries are in APP folder.

Prepare source code:
Visual Studio 2010-2015
Create New Project: Visual Basic/Office/Outlook Add-in
In Project Properties / Publish: 
- Set Publishing Folder to C:\TEMP\PUBLISH
- Set Installation Folder URL to \\localhost\PUBLISH

Note: 
The source code will add the new Auto-reply tab to the ribbon to both the Email Read template and Email Compose template so should appear when opening a new email or an existing email.

With the default source code or published binary installed:
- Open HP Records Manager or HP RM Desktop
- Check 
- Right-click a document and select Send To / Mail
- When prompted, ensure "HPE Records Manager Record reference" is selected and OK
- An Outlook email appears and you'll note the Auto-reply tab is visible on the ribbon
- Send the email
- Back in HP RM, right-click the same or another document and note that the right-click menu does not seem to open.  Right-clicking again makes the menu appear behind the HP RM window.  Select the Help drop-down menu in the top-right and note that it does not display correctly.
- Close HP RM or Outlook and re-open and confirm that HP RM again works OK.

In Visual Studio
Change ReplyRibbonDesigner.vb @ line 112 from:
Me.RibbonType = "Microsoft.Outlook.Mail.Read,Microsoft.Outlook.Mail.Compose"
to:
Me.RibbonType = "Microsoft.Outlook.Mail.Read"

Publish the project
Under C:\TEMP\PUBLISH, run SETUP.EXE to update the add-in
In Outlook, confirm the Auto-reply tab now only appears on Read Email template, not the Compose Email template.
Repeat above steps and confirm the issue no longer appears after conducting a Sent To / Mail operation.

