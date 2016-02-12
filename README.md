## WSUS Miner 
This is a little Powershell script that fetch metric's values from WSUS 3.0 Server.

Support objects:
- _Status_ - WSUS Status (Number of Approved/Declined/Expired/etc updates, full/partially/unsuccess updated clients and so);
- _Info_ - Some WSUS Settings;
- _synchronizationProcess_ - Synchronization process status;
- _lastSynchronization_ - Last Synchronization data;
- _database_ - WSUS Database related info;
- _configuration_ - WSUS configuration info;
- _computerGroup_ - Virtual object to taking computer group statistic (keys: _ComputerTargetsWithUpdateErrorsCount_, _ComputerTargetsNeedingUpdatesCount_, _ComputersUpToDateCount_, _ComputerTargetsUnknownCount_).

Zabbix's LLD available to:
- _computerGroup_ 

How to use:
- Just add to Zabbix Agent config, which run on WSUS host this string: _UserParameter=wsus.miner[*], powershell -File C:\zabbix\scripts\wsus_miner.ps1 -Action "$1" -Object "$2" -Key "$3" -Id "$4"_ 
- Put _wsus_miner.ps1_ to _C:\zabbix\scripts_ dir;
- Make unsigned .ps1 script executable with _Set-ExecutionPolicy RemoteSigned_;
- Set Zabbix Agent's / Server's _Timeout_ to more that 3 sec (may be 10 or 30);
- Import [template](https://github.com/zbx-sadman/wsus_miner/tree/master/Zabbix_Templates) to Zabbix Server;
- Enjoy.

**Note**
Do not try import Zabbix v2.4 template to Zabbix _pre_ v2.4. You need to edit .xml file and make some changes at discovery_rule - filter tags area and change _#_ to _<>_ in trigger expressions. I will try to make template to old Zabbix.

Hints:
- To see keys, run script without "-Key" option: _powershell -File C:\zabbix\scripts\wsus_miner.ps1 -Action "Get" -Object "**Object**"_
- If you use non-english (for example Russian Cyrillic) symbols into Computer Group's names and want to get correct UTF-8 on Zabbix Server side, then you must add _-consoleCP **your_native_codepage**_ parameter to command line. For example to convert from Russian Cyrillic codepage (CP866), use _powershell -File C:\zabbix\scripts\wsus_miner.ps1 -Action "$1" -Object "$2" -Key "$3" -Id "$4" -consoleCP CP866_.

Beware: frequent connections to WSUS may be nuke host server and yours requests will be processeed slowly. To avoid it - don't use small update intervals with Zabbix's Data Items and disable unused.
