RegUpdateAllUsers
=================

Modify HKCU registry key(s) for ALL users on a system

Have you ever needed to update a registry key that is stored in the HKEY_CURRENT_USER hive? Have you also ever needed to update it for ALL users on the system, as well as make it the default setting when a new user profile is created?

That can be a bit of a daunting task. One solution is to add the registry key update to the user’s logon script.

However, there is another way.  The idea is to:

1. Update the currently logged on user’s HKCU (that’s easy enough)
2. Then you must enumerate every profile on the system
3. Find their ntuser.dat file (ntuser.dat contains the contents of the user’s HKCU hive)
4. Load ntuser.dat into a temporary key in the HKLM hive (programmatically or using reg.exe)
5. I use ‘HKLM\TempHive’ as the temporary key
6. Then when you write to “HKLM\TempHive” you are actually editing that user’s HKCU hive.
7. If you load and ntuser.dat for the “Default” user, the settings will take effect for any NEW user profile created on the system
8. If more than 1 user is currently logged on, you can edit their HKCU hive by looking the user up by their SID under HKEY_USERS and writing to it at that location.


It’s a bit of a tedious job, so I wrote a VBScript that takes care of all of the steps listed above. This script has been tested on Windows XP and Windows 7 (x64), but should work on Windows 2000 and newer. It relies on “reg.exe” which ships with all versions of Windows.

UPDATED April 23, 2013 – The script can now set REG_BINARY keys as long as they are in the format used by a regedit.exe export. For example:

   [HKEY_CURRENT_USER\Software\_Test\MyTestBinarySubkey]
   "My Test Binary Value"=hex:23,00,41,00,43,00,42,00,6c,00


To set this binary value using the script, you would modify line 60 to be:

   SetBinaryRegKeys sRegistryRootToUse, strRegPathParent03, "My Test Binary Value","hex:23,00,41,00,43,00,42,00,6c,00"
   

LICENSE: BSD 3-clause
