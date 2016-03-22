## WSUS Miner 
This is a little Powershell script that fetch metric's values from Microsoft WSUS 3.0 Server.

Actual release 1.3.0

Tested on:
- Production mode: Windows Server 2008 R2 SP1, WSUS 3.0 SP1, Powershell 2;
- Non-production mode: Windows Server 2012 R2, WSUS 6, Powershell 4.


Support objects:
- _Info_                   - WSUS informaton;
- _Status_                 - WSUS status (number of Approved/Declined/Expired/etc updates, full/partially/unsuccess updated clients and so);
- _Database_               - WSUS database related info;
- _Configuration_          - WSUS configuration info;
- _ComputerGroup_          - Virtual object to taking computer group statistic;
- _LastSynchronization_    - Last Synchronization data;
- _SynchronizationProcess_ - Synchronization process info.

Actions:
- _Discovery_ - Make Zabbix's LLD JSON;
- _Get_       - Get object metric's value;
- _Count_     - Take number of objects in collection (selected with _ComputerGroup's_ _ComputersUpToDate_ virtual key, for example).

Zabbix's LLD available to:
- _ComputerGroup_.

Virtual keys for _ComputerGroup_ object is:
- _ComputerTargetsWithUpdateErrors_ - Computers updated with errors;
- _ComputerTargetsNeedingUpdates_   - Partially updated computers;
- _ComputersUpToDate_               - Full updated computers;
- _ComputerTargetsUnknown_          - Computers without update information.

Virtual keys for _LastSynchronization_ object is:
- _NotSyncInDays_                   - Now much days was not running Synchronization process.

Virtual keys for _SynchronizationProcess_ object:
- _Status_                          - Synchronization process status

###How to use standalone

    # Show all metrics of 'Configuration' object and see verbose messages
    powershell -NoProfile -ExecutionPolicy "RemoteSigned" -File "wsus_miner.ps1" -Action "Get" -Object "Configuration" -Verbose

    # Make Zabbix's LLD JSON for 'ComputerGroup' object. Group names contains Russian Cyrillic symbols
    ... "wsfc.ps1" -Action "Discovery" -Object "ComputerGroup" -consoleCP CP866

    # Get number of days passed since the last synchronization
    ... "wsfc.ps1" -Action "Get" -Object "LastSynchronization" -Key "NotSyncInDays"

    # Get number of computers that have update errors and placed in ComputerGroup with ID=e4b8b165-4e29-42ec-ac40-66178600ca9b
    ..."wsus_miner.ps1" -Action "Count" -Object "ComputerGroup" -Key "ComputerTargetsWithUpdateErrors" -Id "e4b8b165-4e29-42ec-ac40-66178600ca9b"

###How to use with Zabbix
1. Just include [zbx\_wsus\_miner.conf](https://github.com/zbx-sadman/wsus_miner/tree/master/Zabbix_Templates/zbx_wsus_miner.conf) to Zabbix Agent config;
2. Put _wsus\_miner.ps1_ to _C:\zabbix\scripts_ dir. If you want to place script to other directory, you must edit _wsus\_miner.ps1_ to properly set script's path; 
3. Set Zabbix Agent's / Server's _Timeout_ to more that 3 sec (may be 10 or 30);
4. Import [template](https://github.com/zbx-sadman/wsus_miner/tree/master/Zabbix_Templates) to Zabbix Server;
5. Be sure that Zabbix Agent worked in Active mode - in template used 'Zabbix agent(active)' poller type. Otherwise - change its to 'Zabbix agent' and increase value of server's StartPollers parameter;
6. Enjoy.

**Note**
Do not try import Zabbix v2.4 template to Zabbix _pre_ v2.4. You need to edit .xml file and make some changes at discovery_rule - filter tags area and change _#_ to _<>_ in trigger expressions. I will try to make template to old Zabbix.

###Hints
- To see available metrics, run script without "-Key" option: _powershell -File C:\zabbix\scripts\wsus\_miner.ps1 -Action "Get" -Object "Status"_;
- To measure script runtime use _Verbose_ command line switch;
- To get on Zabbix Server side properly UTF-8 output when have non-english (for example Russian Cyrillic) symbols in Computer Group's names, use  _-consoleCP **your_native_codepage**_ command line option. For example to convert from Russian Cyrillic codepage (CP866): _...wsus\_miner.ps1 ... -consoleCP CP866_;
- If u need additional symbol escaping in LLD JSON - just add one more calls of _$InObject = $InObject.Replace(...)_  in _Prepare-ToZabbix_ function;
- Running the script with PowerShell 3 and above may be require to enable PowerShell 2 compatible mode;
- To measure script runtime use _Verbose_ command line switch,

Beware: frequent connections to WSUS may be nuke host server, make over 9000% CPU utilization and yours requests will be processeed slowly. To avoid it - don't use small update intervals with Zabbix's Data Items and disable unused.
