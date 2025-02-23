# wua_offline
This script uses "offline" database for **Windows Update Agent** (WUA) and can be helpful in the case of broken Windows Update service.

But it still requires online connection to download the updates from official Microsoft repository (http://download.windowsupdate.com).

Running on Windows XP
---------------------

SHA-256 encryption is not supported by Windows XP.

By this reason, under Windows XP the script will not work with the latest version of Windows Update Agent (WUA) base file.

Last supported version of base file signed with SHA-1 key was issued 14.07.2020 and can be found on Internet Archive by the following links:
* http://web.archive.org/web/20200722185106/http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab
* http://web.archive.org/web/20200811034917/http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab

Running the script on Windows XP:
* Download the base file ``wsusscn2.cab`` from the Internet Archive site (its size is about 880 MB).
* Create subfolder ``.\cache\`` and place base file inside it: ``.\cache\wsusscn2.cab``.
* Run the update script: ``start_update``.

Windows Update Restored
-----------------------

Author also recommends to visit the **Windows Update Restored** service:
* http://windowsupdaterestored.com/
