<#
    .SYNOPSIS  
        Return WSUS metrics values, count selected objects, make LLD-JSON for Zabbix

    .DESCRIPTION
        Return WSUS metrics values, count selected objects, make LLD-JSON for Zabbix

    .NOTES  
        Version: 1.3.4
        Name: Microsoft's WSUS Miner
        Author: zbx.sadman@gmail.com
        DateCreated: 27JAN2016
        DateModified: 26FEB2018
        DateModified: 09AUG2019
        DateModified: 25SEP2019
        Testing environment: Windows Server 2008R2 SP1, WSUS 3 SP2, Powershell 2
        Non-production testing environment: Windows Server 2012 R2, WSUS 6, PowerShell 4

    .LINK  
        https://github.com/zbx-sadman

    .PARAMETER ServerHost
        Name of host which hosts WSUS server

    .PARAMETER ServerPort
        WSUS Server's TCP port number

    .PARAMETER ServerSSL
        Use SSL when connect to WSUS server

    .PARAMETER Action
        What need to do with collection or its item:
            Discovery - Make Zabbix's LLD JSON;
            Get - get metric from collection item
            Count - count collection items

    .PARAMETER ObjectType
        Specify "rule" to make collection:
            Info                    - WSUS informaton
            Status                  - WSUS status (number of Approved/Declined/Expired/etc updates, full/partially/unsuccess updated clients and so)
            Database                - WSUS database related info
            Configuration           - WSUS configuration info
            ComputerGroup           - Virtual object to taking computer group statistic
            LastSynchronization     - Last Synchronization data
            SynchronizationProcess  - Synchronization process status (haven't keys)

    .PARAMETER Key
        Specify "path" to collection item's metric 

        If virtual key ends with '*' part, JSON formatted object's properties will be returned.

        Virtual keys for 'ComputerGroup' object:
            ComputerTargetsWithUpdateErrorsCount - Computers updated with errors 
            ComputerTargetsNeedingUpdatesCount   - Partially updated computers
            ComputersUpToDateCount               - Full updated computers
            ComputerTargetsUnknownCount          - Computers without update information 
            ComputerTargetsNotReportedWithin     - Computers that have not been reported for several days
            ComputerTargetsNotUpdatedWithin      - Computers that did not report for several days

        Note: Metrics will be contain object collection if ShowAsSystemObject switch used, and number of collection items otherwise

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

    .PARAMETER ShowAsSystemObject
        Metric's properties will be return as Formatted-List cmdlet result

    .PARAMETER Verbose
        Enable verbose messages

    .EXAMPLE 
        wsus_miner.ps1 -Action "Discovery" -ObjectType "ComputerGroup" -ConsoleCP CP866

        Description
        -----------  
        Make Zabbix's LLD JSON for object "ComputerGroup". Output converted from CP866 to UTF-8.

    .EXAMPLE 
        wsus_miner.ps1 -Action "Count" -ObjectType "ComputerGroup" -Key "ComputerTargetsNeedingUpdatesCount" -Id "020a3aa4-c231-4ffa-a2ff-ff4cc2e95ad0" -defaultConsoleWidth -ShowAsSystemObject
          OR
        wsus_miner.ps1 -Action "Get" -ObjectType "ComputerGroup" -Key "ComputerTargetsNeedingUpdatesCount" -Id "020a3aa4-c231-4ffa-a2ff-ff4cc2e95ad0" -defaultConsoleWidth

        Description
        -----------  
        Return number of computers that needing updates places in group with id "020a3aa4-c231-4ffa-a2ff-ff4cc2e95ad0"

    .EXAMPLE 
        wsus_miner.ps1 -Action "Get" -ObjectType "Status" -defaultConsoleWidth -Verbose -ShowAsSystemObject
          OR
        wsus_miner.ps1 -Action "Get" -ObjectType "Status" -Key "*" -defaultConsoleWidth -Verbose

        Description
        -----------  
        Show formatted list of 'Status' object metrics. Verbose messages is enabled
#>

Param (
   [Parameter(Mandatory = $False)] 
   [string]$ServerHost,
   [Parameter(Mandatory = $False)] 
   [Int32]$ServerPort,
   [Parameter(Mandatory = $False)] 
   [switch]$ServerSSL,
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
   [switch]$DefaultConsoleWidth,
   [Parameter(Mandatory = $False)]
   [switch]$ShowAsSystemObject
)

#Set-StrictMode –Version Latest

# Set US locale to properly formatting float numbers while converting to string
[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"

# Width of console to stop breaking JSON lines
Set-Variable -Name "CONSOLE_WIDTH" -Value 255 -Option Constant

# 
Set-Variable -Name "WSUS_DEFAULT_HOSTNAME"  -Value "localhost" -Option Constant
Set-Variable -Name "WSUS_DEFAULT_PORT_SSL"  -Value 8530 -Option Constant
Set-Variable -Name "WSUS_DEFAULT_PORT_HTTP" -Value 80 -Option Constant

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
         # IsNullOrEmpty used because !$Value give a erong result with $Value = 0 (True).
         # But 0 may be right ID  
         If (($Object.$Property -Eq $Value) -Or ([string]::IsNullOrEmpty($Value))) { $Object }
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
      [Switch]$JSONCompatible,
      [Switch]$ShowAsObject
   );
   # Add here more symbols to escaping if you need
   $EscapedSymbols = @('\', '"');
   $UnixEpoch = Get-Date -Date "01/01/1970";

   $Result = $Null;
   # Need add doublequote around string for other objects when JSON compatible output requested?
   $DoQuote = $False;

   If ($Null -Eq $InputObject) { 
      $Result = $(If ($ErrorCode) { $ErrorCode });
   } Else {
     Switch -Wildcard (($InputObject.GetType()).Name) {
         'Int*'           { $Result = $InputObject;}
         'Boolean'        { $Result = $(If ($InputObject) {"true"} Else {"false"}); }
         'DateTime'       { $Result = (New-TimeSpan -Start $UnixEpoch -End $InputObject).TotalSeconds; }
          Default         { $Result = $InputObject; $DoQuote = $True; }
       }

     # Normalize String object
     $Result = $Result.ToString().Trim();             
     If (!$NoEscape) { ForEach ($Symbol in $EscapedSymbols) { $Result = $Result.Replace($Symbol, "\$Symbol"); } }
   }

   # Doublequote object if adherence to JSON standart requested
   If ($JSONCompatible -And $DoQuote) { 
      "`"$Result`"";
   } else {
      $Result;
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

Function Make-LLD-JSON {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [array]$ObjectProperties, 
      [Switch]$Pretty
   ); 
   Begin   {
      [String]$Result = "";
      # Init JSON-string $InObject
      $Result += "{`n `"data`":[`n";
   } 
   Process {
      $Result += (Make-JSON -InputObject $InputObject -ObjectProperties $ObjectProperties -Pretty -LLDFormat);
   }
   End {
      # Finalize and return JSON
      "$Result`n    ]`n}";
   }
}


#
#  Make & return JSON, due PoSh 2.0 haven't Covert-ToJSON
#
Function Make-JSON {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [Array]$ObjectProperties, 
      [Switch]$Pretty,
      [Switch]$LLDFormat
   ); 
   Begin   {
      [String]$Result = "";
      # Pretty json contain spaces, tabs and new-lines
      If ($Pretty) { $CRLF = "`n"; $Tab = "    "; $Space = " "; } Else { $CRLF = $Tab = $Space = ""; }
      If ($LLDFormat) { $KeyPrefix = "{#"; $KeyPostfix = "}"; } Else { $KeyPrefix = $KeyPostfix = ""; }

      # Take each Item from $InObject, get Properties that equal $ObjectProperties items and make JSON from its
      $itFirstObject = $True;
      If ($Null -Eq $ObjectProperties) { $ObjectProperties = @($InputObject.PSObject.Properties | Select-Object -Expand Name)}
   } 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) {
         # Skip object when its $Null
         If ($Null -Eq $Object) { Continue; }

         If (-Not $itFirstObject) { $Result += ",$CRLF"; }
         $itFirstObject=$False;
         $Result += $(If ($LLDFormat) {"$Tab$Tab"} )+"{$Space"; 
         $itFirstProperty = $True;
         # Process properties. No comma printed after last item
         ForEach ($Property in $ObjectProperties) {
            $Chunk = PrepareTo-Zabbix -InputObject $Object.$Property -JSONCompatible; 
            if ($Null -Eq $Chunk) { Continue; }
            If (-Not $itFirstProperty) { $Result += ","+$(If (-Not $LLDFormat) { "$CRLF$Tab" } Else { $Space }) }
            $itFirstProperty = $False;
            #Write-Host ($Object.$Property).GetType().Name;
            $Result += "`"$KeyPrefix$Property$KeyPostfix`":$Chunk";
         }
         # No comma printed after last string
         $Result += $(If (-Not $LLDFormat) { "$CRLF" } Else { $Space }) + "}";
      }
   }
   End {
      # Finalize and return JSON
      $Result;
   }
}

#
#  Return value of object's metric defined by key-chain from $Keys Array
#
Function Get-Objects { 
   Param (
      [Parameter(Mandatory = $False, ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [Array]$Keys
   ); 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
        If ($Null -Eq $Object) { Continue; }
        # Expand all metrics related to keys contained in array step by step
        ForEach ($Key in $Keys) {              
           If ("*" -Eq $Key) { Break; }
           If ($Key) {
              $Object = Select-Object -InputObject $Object -ExpandProperty $Key -ErrorAction SilentlyContinue;
              If ($Error) { Break; }
           }
        }
#        $Object;
      }
   }
   End {
      # Finalize and return JSON
     $Object;
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

Function Count-Collection { 
   Param (
      [Parameter(Mandatory = $False, ValueFromPipeline = $True)] 
      [PSObject]$InputObject
   ); 
   Begin {
      $Count = 0;
   }
   Process {
     ForEach ($Object in $InputObject) {
        If ($Null -Ne $Object) { $Count++; } 
     }
   }
   End {
      $Count;
   }
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

If ([string]::IsNullOrEmpty($ServerHost)) { $ServerHost = $WSUS_DEFAULT_HOSTNAME; }
If (($Null -Eq $ServerPort) -Or (0 -Eq $ServerPort)) { $ServerPort = $( if ($ServerSSL) { $WSUS_DEFAULT_PORT_SSL } else { $WSUS_DEFAULT_PORT_HTTP }); }

Write-Verbose "$(Get-Date) Trying to connect to WSUS Server: $($ServerHost):$($ServerPort)";
$WSUS = $( if ($UseNativeCmdLets) {
                $( if ($ServerSSL) { Get-WsusServer -Name $ServerHost -PortNumber $ServerPort -UseSSL } else { Get-WsusServer -Name $ServerHost -PortNumber $ServerPort } );
              } else {
                 [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($ServerHost, $ServerSSL, $ServerPort);
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
         If ('Discovery'-Eq $Action) {
           $ComputerTargetGroups
         } Else {
         $Today = Get-Date;
         $ComputerTarget = $ComputerTargetsWithUpdateErrors = $ComputerTargetsNeedingUpdates = $ComputersUpToDate = $ComputerTargetsUnknown = $ComputerTargetsNotReportedWithin = $ComputerTargetsNotUpdatedWithin = @();
         $ComputerTargetGroups | % { 
           $TotalSummary = $_.GetTotalSummaryPerComputerTarget();
           $TotalSummary | % {
              #Write-Host $_;
              $ComputerTarget += $_;
              If (0 -ne $_.FailedCount) { $ComputerTargetsWithUpdateErrors += $_; } 
              If ((0 -eq $_.FailedCount) -And (0 -ne ($_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount))) { $ComputerTargetsNeedingUpdates += $_; } 
              If (0 -eq ($_.FailedCount+$_.UnknownCount+$_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount)) { $ComputersUpToDate += $_; } 
              If ((0 -ne $_.UnknownCount) -And (0 -eq ($_.FailedCount+$_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount))) { $ComputerTargetsUnknown += $_; } 
              #If ((New-TimeSpan -Start $_.LastUpdated -End $Today).Days -gt $Value) { $ComputerTargetsNotUpdatedWithin += $_; } 
              #If ((New-TimeSpan -Start $_.LastReportedStatusTime -End $Today).Days -gt $Value) { $ComputerTargetsNotReportedWithin += $_; } 
           }
         };         

         $ComputerTarget = $ComputerTarget | Group-Object 'ComputerTargetId' | %{ $_.Group | Select 'ComputerTargetId' -First 1} ;
         $ComputerTargetsWithUpdateErrors = $ComputerTargetsWithUpdateErrors | Group-Object 'ComputerTargetId' | %{ $_.Group | Select 'ComputerTargetId' -First 1} ;
         $ComputerTargetsNeedingUpdates = $ComputerTargetsNeedingUpdates | Group-Object 'ComputerTargetId' | %{ $_.Group | Select 'ComputerTargetId' -First 1} ;
         $ComputersUpToDate = $ComputersUpToDate | Group-Object 'ComputerTargetId' | %{ $_.Group | Select 'ComputerTargetId' -First 1} ;
         $ComputerTargetsUnknown = $ComputerTargetsUnknown | Group-Object 'ComputerTargetId' | %{ $_.Group | Select 'ComputerTargetId' -First 1} ;

          If (-Not $ShowAsSystemObject) {
            # Autocast to Int32
            $ComputerTarget                   = $ComputerTarget | Count-Collection;
            $ComputerTargetsWithUpdateErrors  = $ComputerTargetsWithUpdateErrors | Count-Collection;
            $ComputerTargetsNeedingUpdates    = $ComputerTargetsNeedingUpdates | Count-Collection;
            $ComputersUpToDate                = $ComputersUpToDate | Count-Collection;
            $ComputerTargetsUnknown           = $ComputerTargetsUnknown | Count-Collection;
            $ComputerTargetsNotReportedWithin = $ComputerTargetsNotReportedWithin | Count-Collection;
            $ComputerTargetsNotUpdatedWithin  = $ComputerTargetsNotUpdatedWithin | Count-Collection;
         }

        $ComputerTargetGroupsData = New-Object PSObject -Property @{"ComputerTarget" = $ComputerTarget;
                                                                     "ComputerTargetsWithUpdateErrors" = $ComputerTargetsWithUpdateErrors;
                                                                     "ComputerTargetsNeedingUpdates" = $ComputerTargetsNeedingUpdates;
                                                                     "ComputersUpToDate" = $ComputersUpToDate;
                                                                     "ComputerTargetsUnknown" = $ComputerTargetsUnknown;
                                                                     "ComputerTargetsNotReportedWithin" = $ComputerTargetsNotReportedWithin;
                                                                     "ComputerTargetsNotUpdatedWithin" = $ComputerTargetsNotUpdatedWithin
         };
         $ComputerTargetGroupsData;
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
         Make-LLD-JSON -InputObject $Objects -ObjectProperties $ObjectProperties -Pretty;
      }
      'Get' {
         # Get metrics or metric list
         Write-Verbose "$(Get-Date) Getting metric related to key: '$Key'";
         $Objects = Get-Objects -InputObject $Objects -Keys $Keys;
         if ($ShowAsSystemObject) {
            Out-String -InputObject (Format-List -InputObject $Objects -Property *)
         } ElseIf ("*" -Eq $Keys[-1]) { 
            Make-JSON -InputObject $Objects -Pretty;
         } Else {
            PrepareTo-Zabbix -InputObject $Objects -ErrorCode $ErrorCode;
         }
      }
      'Count' { 
         Write-Verbose "$(Get-Date) Counting objects";  
         # ++ must be faster that .Count, due don't enumerate object list
         $Objects = Get-Objects -InputObject $Objects -Keys $Keys;
         $Objects | Count-Collection;
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
