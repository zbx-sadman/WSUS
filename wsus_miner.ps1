<#
    .SYNOPSIS  
        Return WSUS metrics values, count selected objects, make LLD-JSON for Zabbix

    .DESCRIPTION
        Return WSUS metrics values, count selected objects, make LLD-JSON for Zabbix

    .NOTES  
        Version: 1.3.2
        Name: Microsoft's WSUS Miner
        Author: zbx.sadman@gmail.com
        DateCreated: 27JAN2016
        DateModified: 26FEB2018
        Testing environment: Windows Server 2008R2 SP1, WSUS 3 SP2, Powershell 2
        Non-production testing environment: Windows Server 2012 R2, WSUS 6, PowerShell 4

    .LINK  
        https://github.com/zbx-sadman

    .PARAMETER Action
        What need to do with collection or its item:
            Discovery - Make Zabbix's LLD JSON;
            Get - get metric from collection item
            Count - count collection items

    .PARAMETER ObjectType
        Define rule to make collection:
            Info                    - WSUS informaton
            Status                  - WSUS status (number of Approved/Declined/Expired/etc updates, full/partially/unsuccess updated clients and so)
            Database                - WSUS database related info
            Configuration           - WSUS configuration info
            ComputerGroup           - Virtual object to taking computer group statistic
            LastSynchronization     - Last Synchronization data
            SynchronizationProcess  - Synchronization process status (haven't keys)

    .PARAMETER Key
        Define "path" to collection item's metric 

        Virtual keys for 'ComputerGroup' object:
            ComputerTargetsWithUpdateErrorsCount - Computers updated with errors 
            ComputerTargetsNeedingUpdatesCount   - Partially updated computers
            ComputersUpToDateCount               - Full updated computers
            ComputerTargetsUnknownCount          - Computers without update information 
            ComputerTargetsNotReportedWithin     - Computers that have not been reported for several days
            ComputerTargetsNotUpdatedWithin      - Computers that did not report for several days

        Virtual keys for 'LastSynchronization' object:
            NotSyncInDays                        - Now much days was not running Synchronization process;

    .PARAMETER Value
        Key specific parameter:
            For 'ComputerTargetsNotReportedWithin' parameter - number of days
            For 'ComputerTargetsNotUpdatedWithin'  parameter - number of days too.


    .PARAMETER Id
        Used to select only one item from collection

    .PARAMETER ErrorCode
        What must be returned if any process error will be reached

    .PARAMETER ConsoleCP
        Codepage of Windows console. Need to properly convert output to UTF-8

    .PARAMETER DefaultConsoleWidth
        Say to leave default console width and not grow its to $CONSOLE_WIDTH

    .PARAMETER Verbose
        Enable verbose messages

    .EXAMPLE 
        wsus_miner.ps1 -Action "Discovery" -ObjectType "ComputerGroup" -ConsoleCP CP866

        Description
        -----------  
        Make Zabbix's LLD JSON for object "ComputerGroup". Output converted from CP866 to UTF-8.

    .EXAMPLE 
        wsus_miner.ps1 -Action "Count" -ObjectType "ComputerGroup" -Key "ComputerTargetsNeedingUpdatesCount" -Id "020a3aa4-c231-4ffa-a2ff-ff4cc2e95ad0" -defaultConsoleWidth

        Description
        -----------  
        Return number of computers that needing updates places in group with id "020a3aa4-c231-4ffa-a2ff-ff4cc2e95ad0"

    .EXAMPLE 
        wsus_miner.ps1 -Action "Get" -ObjectType "Status" -defaultConsoleWidth -Verbose

        Description
        -----------  
        Show formatted list of 'Status' object metrics. Verbose messages is enabled
#>

Param (
   [Parameter(Mandatory = $False)] 
   [ValidateSet('Discovery', 'Get', 'Count')]
   [string]$Action,
   [Parameter(Mandatory = $False)]
   [ValidateSet('Info', 'Status', 'Database', 'Configuration', 'ComputerGroup', 'LastSynchronization', 'SynchronizationProcess')]
   [Alias('Object')]
   [string]$ObjectType,
   [Parameter(Mandatory = $False)]
   [string]$Key,
   [Parameter(Mandatory = $False)]
   [string]$Id,
   [Parameter(Mandatory = $False)]
   [Int32]$Value,
   [Parameter(Mandatory = $False)]
   [String]$ErrorCode,
   [Parameter(Mandatory = $False)]
   [string]$ConsoleCP,
   [Parameter(Mandatory = $False)]
   [switch]$DefaultConsoleWidth
)

#Set-StrictMode –Version Latest

# Set US locale to properly formatting float numbers while converting to string
[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"

# Width of console to stop breaking JSON lines
Set-Variable -Name "CONSOLE_WIDTH" -Value 255 -Option Constant

####################################################################################################################################
#
#                                                  Function block
#    
####################################################################################################################################
#
#  Select object with Property that equal Value if its given or with Any Property in another case
#
Function PropertyEqualOrAny {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [PSObject]$Property,
      [PSObject]$Value
   );
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         # IsNullorEmpty used because !$Value give a erong result with $Value = 0 (True).
         # But 0 may be right ID  
         If (($Object.$Property -Eq $Value) -Or ([string]::IsNullorEmpty($Value))) { $Object }
      }
   } 
}

#
#  Prepare string to using with Zabbix 
#
Function PrepareTo-Zabbix {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [String]$ErrorCode,
      [Switch]$NoEscape,
      [Switch]$JSONCompatible
   );
   Begin {
      # Add here more symbols to escaping if you need
      $EscapedSymbols = @('\', '"');
      $UnixEpoch = Get-Date -Date "01/01/1970";
   }
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         If ($Null -Eq $Object) {
           # Put empty string or $ErrorCode to output  
           If ($ErrorCode) { $ErrorCode } Else { "" }
           Continue;
         }
         # Need add doublequote around string for other objects when JSON compatible output requested?
         $DoQuote = $False;
         Switch (($Object.GetType()).FullName) {
            'System.Boolean'  { $Object = [int]$Object; }
            'System.DateTime' { $Object = (New-TimeSpan -Start $UnixEpoch -End $Object).TotalSeconds; }
            Default           { $DoQuote = $True; }
         }
         # Normalize String object
         $Object = $( If ($JSONCompatible) { $Object.ToString().Trim() } else { Out-String -InputObject (Format-List -InputObject $Object -Property *) });
         
         If (!$NoEscape) { 
            ForEach ($Symbol in $EscapedSymbols) { 
               $Object = $Object.Replace($Symbol, "\$Symbol");
            }
         }

         # Doublequote object if adherence to JSON standart requested
         If ($JSONCompatible -And $DoQuote) { 
            "`"$Object`"";
         } else {
            $Object;
         }
      }
   }
}

#
#  Convert incoming object's content to UTF-8
#
Function ConvertTo-Encoding ([String]$From, [String]$To){  
   Begin   {  
      $encFrom = [System.Text.Encoding]::GetEncoding($from)  
      $encTo = [System.Text.Encoding]::GetEncoding($to)  
   }  
   Process {  
      $bytes = $encTo.GetBytes($_)  
      $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)  
      $encTo.GetString($bytes)  
   }  
}

#
#  Make & return JSON, due PoSh 2.0 haven't Covert-ToJSON
#
Function Make-JSON {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [array]$ObjectProperties, 
      [Switch]$Pretty
   ); 
   Begin   {
      [String]$Result = "";
      # Pretty json contain spaces, tabs and new-lines
      If ($Pretty) { $CRLF = "`n"; $Tab = "    "; $Space = " "; } Else { $CRLF = $Tab = $Space = ""; }
      # Init JSON-string $InObject
      $Result += "{$CRLF$Space`"data`":[$CRLF";
      # Take each Item from $InObject, get Properties that equal $ObjectProperties items and make JSON from its
      $itFirstObject = $True;
   } 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) {
         # Skip object when its $Null
         If ($Null -Eq $Object) { Continue; }

         If (-Not $itFirstObject) { $Result += ",$CRLF"; }
         $itFirstObject=$False;
         $Result += "$Tab$Tab{$Space"; 
         $itFirstProperty = $True;
         # Process properties. No comma printed after last item
         ForEach ($Property in $ObjectProperties) {
            If (-Not $itFirstProperty) { $Result += ",$Space" }
            $itFirstProperty = $False;
            $Result += "`"{#$Property}`":$Space$(PrepareTo-Zabbix -InputObject $Object.$Property -JSONCompatible)";
         }
         # No comma printed after last string
         $Result += "$Space}";
      }
   }
   End {
      # Finalize and return JSON
      "$Result$CRLF$Tab]$CRLF}";
   }
}

#
#  Return value of object's metric defined by key-chain from $Keys Array
#
Function Get-Metric { 
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [Array]$Keys
   ); 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
        If ($Null -Eq $Object) { Continue; }
        # Expand all metrics related to keys contained in array step by step
        ForEach ($Key in $Keys) {              
           If ($Key) {
              $Object = Select-Object -InputObject $Object -ExpandProperty $Key -ErrorAction SilentlyContinue;
              If ($Error) { Break; }
           }
        }
        $Object;
      }
   }
}

#
#  Exit with specified ErrorCode or Warning message
#
Function Exit-WithMessage { 
   Param (
      [Parameter(Mandatory = $True, ValueFromPipeline = $True)] 
      [String]$Message, 
      [String]$ErrorCode 
   ); 
   If ($ErrorCode) { 
      $ErrorCode;
   } Else {
      Write-Warning ($Message);
   }
   Exit;
}

####################################################################################################################################
#
#                                                 Main code block
#    
####################################################################################################################################

# May be use WSUS 6.3 cmdlets?
$UseNativeCmdLets = $True;
# UpdateServices module is loaded from C:\Windows\System32\WindowsPowerShell\v1.0\Modules\UpdateServices\? (WSUS 6.x on Windows Server 2012)
Write-Verbose "$(Get-Date) Test 'UpdateServices' module state"
If (-Not (Get-Module -List -Name UpdateServices -Verbose:$False)) {
   # No loaded PowerShell module found
   # May be run with WSUS 3.0? Try to load assembly from file
   Write-Verbose "$(Get-Date) Loading 'Microsoft.UpdateServices.Administration' assembly from file"
   Try {
      Add-Type -Path "$Env:ProgramFiles\Update Services\Api\Microsoft.UpdateServices.Administration.dll";
   } Catch {
      Throw ("Error loading the required assemblies to use the WSUS API from {0}" -f "$Env:ProgramFiles\Update Services\Api")
   }
   $UseNativeCmdLets = $False;
}

Write-Verbose "$(Get-Date) Trying to connect to local WSUS Server"
$WSUS = $( if ($UseNativeCmdLets) {
                 Get-WsusServer;
              } else {
                 [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer();
              }
);

If ($Null -Eq $WSUS) {
   Exit-WithMessage -Message "Connection failed" -ErrorCode $ErrorCode;
}
Write-Verbose "$(Get-Date) Connection established";

# split key to subkeys
$Keys = $Key.Split(".");

Write-Verbose "$(Get-Date) Creating collection of specified object: '$ObjectType'";
$Objects =  $(
   Switch ($ObjectType) {
      'Info' { 
         $WSUSObjectCopy = Select-Object -InputObject $WSUS -First 1; 
         Add-Member -Force -InputObject $WSUSObjectCopy -MemberType NoteProperty -Name "FullVersion" -Value $WSUSObjectCopy.Version.ToString();
         $WSUSObjectCopy;
      }
      'Status' {
         $WSUS.GetStatus(); 
      }
      'Database' { 
         $WSUS.GetDatabaseConfiguration(); 
      }
      'Configuration' { 
         $WSUS.GetConfiguration(); 
      }
      'ComputerGroup' {
         $ComputerTargetGroups = PropertyEqualOrAny -InputObject $WSUS.GetComputerTargetGroups() -Property ID -Value $Id;
         Switch ($Key) {
            'ComputerTarget' {
                # Include all computers
                # Sort -Unique ?
                $ComputerTargetGroups | % { $_.GetTotalSummaryPerComputerTarget() } ;
            }
            'ComputerTargetsWithUpdateErrors' {
                # Include all failed (property FailedCount <> 0) computers
                $ComputerTargetGroups | % { $_.GetTotalSummaryPerComputerTarget() } |
                   Where { 0 -ne $_.FailedCount } ;
            }                                                                                                                                       
            'ComputerTargetsNeedingUpdates' {
                # Include no failed, but not installed, downloaded, pending reboot computers
                $ComputerTargetGroups | % { $_.GetTotalSummaryPerComputerTarget() } |
                   Where { (0 -eq $_.FailedCount) -And (0 -ne ($_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount)) };
            }                                                         
            'ComputersUpToDate' {
                # include only no failed, unknown, not installed, downloaded, pending reboot
                $ComputerTargetGroups | % { $_.GetTotalSummaryPerComputerTarget() } |
                   Where { 0 -eq ($_.FailedCount+$_.UnknownCount+$_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount) };
            }                    
            'ComputerTargetsUnknown' {
                # include only unknown, but no failed, not installed, downloaded, pending reboot
                $ComputerTargetGroups | % { $_.GetTotalSummaryPerComputerTarget() } |
                   Where { (0 -ne $_.UnknownCount) -And (0 -eq ($_.FailedCount+$_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount))};
            }
            'ComputerTargetsNotReportedWithin' {
                $Today = Get-Date;
                $ComputerTargetGroups | % { $_.GetComputerTargets() } | Sort-Object -Property Id -Unique | 
#                      % { (New-TimeSpan -Start $_.LastReportedStatusTime -End $Today).Days }
                   Where { (New-TimeSpan -Start $_.LastReportedStatusTime -End $Today).Days -gt $Value } 
            }
            'ComputerTargetsNotUpdatedWithin' {
                $Today = Get-Date;
                $ComputerTargetGroups | % { $_.GetTotalSummaryPerComputerTarget() } |
                   Where { (New-TimeSpan -Start $_.LastUpdated -End $Today).Days -gt $Value }
            }
            Default { $ComputerTargetGroups; }
         }
      }
      'LastSynchronization' { 
          # | Select-Object make copy which used for properly Add-Member work
          $LastSynchronizationInfo = $WSUS.GetSubscription().GetLastSynchronizationInfo() | Select-Object;
          # Just add new Property for Virtual key "NotSyncInDays"
          Add-Member -InputObject $LastSynchronizationInfo -MemberType 'NoteProperty' -Name 'NotSyncInDays' -Value (New-TimeSpan -Start $LastSynchronizationInfo.StartTime.DateTime -End (Get-Date)).Days;
          $LastSynchronizationInfo;
      }
      'SynchronizationProcess' { 
          # SynchronizationStatus contain one value
         New-Object PSObject -Property @{"Status" = $WSUS.GetSubscription().GetSynchronizationStatus()};
      }
   }  
);

Write-Verbose "$(Get-Date) Collection created, begin processing its with action: '$Action'";

$Result = $(
   # if no object in collection: 1) JSON must be empty; 2) 'Get' must be able to return ErrorCode
   Switch ($Action) {
      'Discovery' {
         # Discovery given object, make json for zabbix
         Switch ($ObjectType) {
           'ComputerGroup' { $ObjectProperties = @("NAME", "ID"); }
         }
         Write-Verbose "$(Get-Date) Generating LLD JSON";
         Make-JSON -InputObject $Objects -ObjectProperties $ObjectProperties -Pretty;
      }
      'Get' {
         # Get metrics or metric list
         Write-Verbose "$(Get-Date) Getting metric related to key: '$Key'";
         PrepareTo-Zabbix -InputObject (Get-Metric -InputObject $Objects -Keys $Keys) -ErrorCode $ErrorCode;
      }
      'Count' { 
         Write-Verbose "$(Get-Date) Counting objects";  
         # ++ must be faster that .Count, due don't enumerate object list
         $Count = 0;
         ForEach ($Object in $Objects) {
            If ($Null -Ne $Object) { $Count++; } 
         }
         $Count;
      }
   }
);

# Convert string to UTF-8 if need (For Zabbix LLD-JSON with Cyrillic chars for example)
if ($consoleCP) { 
   Write-Verbose "$(Get-Date) Converting output data to UTF-8";
   $Result = $Result | ConvertTo-Encoding -From $consoleCP -To UTF-8; 
}

# Break lines on console output fix - buffer format to 255 chars width lines 
if (!$defaultConsoleWidth) { 
   Write-Verbose "$(Get-Date) Changing console width to $CONSOLE_WIDTH";
   mode con cols=$CONSOLE_WIDTH; 
}

Write-Verbose "$(Get-Date) Finishing";

$Result;
