About this project
------------------
This script uses "offline" database for **Windows Update Agent** (WUA) and can be helpful in the case of broken **Windows Update** service.

But it still requires online connection to download the updates from official Microsoft repository (http://download.windowsupdate.com).

Running on Windows XP
---------------------
SHA-256 encryption is not supported by Windows XP.
By this reason, the script executed on Windows XP will not work with the latest version of WUA database.

Last supported version of the database file signed with SHA-1 key was issued on 14.07.2020 and can be found on **Internet Archive** by the following links:
* http://web.archive.org/web/20200722185106/http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab
* http://web.archive.org/web/20200811034917/http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab

Running on Windows XP:
* Download the database file ``wsusscn2.cab`` from the Internet Archive site (its size is about 880 MB).
* Create a subfolder ``.\cache\`` and place the database file inside it: ``.\cache\wsusscn2.cab``.
* Run the update script: ``start_update``.

Microsoft Update Catalog
------------------------
Separate installers for each update can be found on official **Microsoft Update Catalog**:
* https://www.catalog.update.microsoft.com/

Windows Update Restored
-----------------------
Author also recommends to visit the **Windows Update Restored** project:
* http://windowsupdaterestored.com/
