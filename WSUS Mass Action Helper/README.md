## Microsoft's WSUS Mass Action Helper
This is a Powershell script that helps to do mass action with Updates: accept license agreement, approve for computer group.


Actual release 0.9.0

Tested on:
- Production mode: Windows Server 2008 R2 SP1, WSUS 3.0 SP1, Powershell 2;



###How to use

Just start .ps1-script, choose computer with some unapproved updates, check updates that you need to approve, and groups to which updates must be approved.
Then click "Approve to..". 

**Note** You can try to use context menu with Computers/Updates lists to copy come information to clipboard;