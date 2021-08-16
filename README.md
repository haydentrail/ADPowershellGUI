# Active Directory Powershell GUI

I developed this gui to make searching across multiple domains easier.  The ability to save queries and enter raw powershell commands makes this a powerful tool in any administrators toolset.

Requires the Microsoft Remote Server Administration tools installed for the active directory powershell commands. Optionally you can install the ImportExcel module to allow exporting directly to excel. 

## Tips 

 * When you first run the app there will be no domains configured and it will prompt to enter a domain or forest. If you enter a forest it will enumerate domains in the forest.
 * When you close the form it saves all of the current runtime data to powershellGUI.config and loads it back in on the next run.
 * The window state is restored to the same location and size as it was when you last closed the form.

### Written by. 
 * Author  : Hayden Trail
 * Email   : hayden@tailoredit.co.nz
 * Company : Tailored IT Solutions

### VERSION CONTROL:

* 0.1.1:
  * Initial Release
* 0.1.3:
  * Added Forest, Trusts, sites, subnets to query options
  * Added ability to add Domain or server 
  * Removed all company references
  * Added retrieve domains from forest
  * Changed form to start un-maximised.  Too many issues with resize
  * Added Raw command entry
  * Added ability to save and load RAW queries
  * Added form returns to the last location and size
  * Added version control
  * Added self updating process using updater.ps1

## Future Enhancements 

 * Plugins. You will be able to create new tabs with your own plugins
 * Tab Pinning.  Ability to lock a tab so that if a command is run against the same domain it does not overwrite the data

## Known issues 

 * Sometimes the form doesnt allow interaction, not sure why this happens.
 * The form freezes when a long running command has been issued.  I tried to use powershell jobs and runspace but found that powershell would randomly fail to execute even basic commands like write-host.  Decided to stick to using a single threaded model for now.
 * I dont account for all attribute types so some return values (mainly in raw commands) will be somewhat useless.

