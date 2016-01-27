## WSUS Miner 
This is a little Powershell script to mine metrics from WSUS 3.0 Server.

How to use:
- Just add to Zabbix Agent config, which run on WSUS host this string: _UserParameter=wsus.miner[*], powershell -File C:\zabbix\scripts\wsus_miner.ps1 -Action "$1" -Object "$2" -Key "$3"_
- Put _wsus_miner.ps1_ to _C:\zabbix\scripts_ dir;
- Make unsigned .ps1 script executable with _Set-ExecutionPolicy RemoteSigned_;
- Set Zabbix Agent's _Timeout_ to more that 3 sec (may be 10 or 30);
- Import [template](https://github.com/zbx-sadman/wsus_miner/tree/master/Zabbix_Templates) to Zabbix Server.
- Enjoy.


Beware: connection to WSUS is slow, don't use small timeouts with Data Items.

