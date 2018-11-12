# Readonly Mailbox Access / Folder Based Permissions
## How to manage read only access to exchange mailboxes

What do you do if a co-worker quits and someone must continue working with his/her emails?

* export to pst? ... which means duplicating data and then maybe loose it somewhere?
* give the person access to the mailbox? ... probably, but how much access?

If you are managing Microsoft Exchange you may already have discovered that there is a "ReadPermission" parameter for the Add-MailboxPermission cmdlet, but it does not work. At least I never made it work and found a lot of evidence on the internet that this parameter doesn't, indeed.

What you CAN do, however, is give folder base permissions instead. This works just fine.

Paul Cunningham wrote a [nice article](https://practical365.com/exchange-server/grant-read-access-exchange-mailbox/) about the topic along with a script doing just that (Thank you!). Unfortunately, his script only works for mailboxes used in an all english environment. In my environment, however, I have to deal with English, German, Russion, Chinese, you name it. And you don't want to put it all necessary folders in all languages.

Lucky us, Microsoft implemented a folder type attribute, which always is in english. So I took Mr Cunningham's scripts and rewrote them to my use, using folder types instead of names (leaving out those not relevant for a user) and upgraded it a bit with additional checks, bells and whistles.

You will find 3 scripts here (no pipeline support):

### Add-MailboxUserFolderPermissions.ps1 -Mailbox user -User user/group -Access accesstype

... to add permissions for a user or (universal security distribution (!)) group to a mailbox. The accesstype matches the permission types of the standard Add-MailboxFolderPermission cmdlet.

### Get-MailboxUserFolderPermissions.ps1 -Mailbox user

... to show the current state of user permissions on the mailbox's folders. I recommend piping the output to Out-Gridview since some people have A LOT of folders!

### Remove-MailboxUserFolderPermissions.ps1 -Mailbox user -User user/group

... to remove any folder permissions set for the give user or group.



## Two important things to take care of:

### 1) Universal Security Distribution Groups

Folder Permissions for groups require a mail-enables universal security group. This is mandatory on Exchange 2016+ and only one way to get it:

* Create a universal security group in Active-Directory (NO Distribution Group setting here!)
* run a Get-Group | Enable-DistributionGroup in your Exchange Management console for the newly created group

Creating a universal distribution group directly in Exchange **does not work**! It must be security first!


### 2) "/" in mailbox folder names

Back in the days of Outook 2000/XP and before, whenever you added a "/" in the name of a folder in Outlook, it would split the name at that position and create a subfolder from the second half. Fortunately, this behavior is gone by know. Nevertheless, Exchange **still** cannot process foldernames with "/", at least not in the way I have written the script (state of 2018-11-12). Maybe you can exchange the character before setting the permissions, but I haven't tried, yet. I will update this, as soon as I know.


Good luck!
Max
