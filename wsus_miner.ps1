<#
    .SYNOPSIS  
        Return WSUS metrics values, count selected objects, make LLD-JSON for Zabbix

    .DESCRIPTION
        Return WSUS metrics values, count selected objects, make LLD-JSON for Zabbix

    .NOTES  
        Version: 1.0.1
        Name: Microsoft's WSUS Miner
        Author: zbx.sadman@gmail.com
        DateCreated: 28FEB2016
        Testing environment: Windows Server 2008R2 SP1, WSUS 3.0 SP2, Powershell 2.0

    .LINK  
        https://github.com/zbx-sadman

    .PARAMETER Action
        What need to do with collection or its item:
            Discovery - Make Zabbix's LLD JSON;
            Get - get metric from collection item
            Count - count collection items

    .PARAMETER Object
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

        Virtual keys for 'LastSynchronization' object:
            NotSyncInDays                        - Now much days was not running Synchronization process;

    .PARAMETER Id
        Used to select only one item from collection

    .PARAMETER ConsoleCP
        Codepage of Windows console. Need to properly convert output to UTF-8

    .PARAMETER DefaultConsoleWidth
        Say to leave default console width and not grow its to $CONSOLE_WIDTH

    .PARAMETER Verbose
        Enable verbose messages

    .EXAMPLE 
        wsus_miner.ps1 -Action "Discovery" -Object "ComputerGroup" -ConsoleCP CP866

        Description
        -----------  
        Make Zabbix's LLD JSON for object "ComputerGroup". Output converted from CP866 to UTF-8.

    .EXAMPLE 
        wsus_miner.ps1 -Action "Count" -Object "ComputerGroup" -Key "ComputerTargetsNeedingUpdatesCount" -Id "020a3aa4-c231-4ffa-a2ff-ff4cc2e95ad0" -defaultConsoleWidth

        Description
        -----------  
        Return number of computers that needing updates places in group with id "020a3aa4-c231-4ffa-a2ff-ff4cc2e95ad0"

    .EXAMPLE 
        wsus_miner.ps1 -Action "Get" -Object "Status" -defaultConsoleWidth -Verbose

        Description
        -----------  
        Show formatted list of 'Status' object metrics. Verbose messages is enabled
#>

Param (
        [Parameter(Mandatory = $True)] 
        [string]$Action,
        [Parameter(Mandatory = $True)]
        [string]$Object,
        [Parameter(Mandatory = $False)]
        [string]$Key,
        [Parameter(Mandatory = $False)]
        [string]$Id,
        [Parameter(Mandatory = $False)]
        [string]$ConsoleCP,
        [Parameter(Mandatory = $False)]
        [switch]$DefaultConsoleWidth
      )

# Set US locale to properly formatting float numbers while converting to string
[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"

# Width of console to stop breaking JSON lines
Set-Variable -Name "CONSOLE_WIDTH" -Value 255 -Option Constant -Scope Global

####################################################################################################################################
#
#                                                  Function block
#    
####################################################################################################################################

#
#  Select object with ID if its given or with Any ID in another case
#
filter IDEqualOrAny($Id) { if (($_.Id -Eq $Id) -Or (!$Id)) { $_ } }

#
#  Prepare string to using with Zabbix 
#
Function Prepare-ToZabbix {
  Param (
     [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
     [PSObject]$InObject
  );
  $InObject = ($InObject.ToString());

  $InObject = $InObject.Replace("`"", "\`"");

  $InObject;
}

#
#  Convert incoming object's content to UTF-8
#
function ConvertTo-Encoding ([string]$From, [string]$To){  
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
#  Return value of object's metric defined by key-chain from $Keys Array
#
Function Get-Metric { 
   Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
      [PSObject]$InObject, 
      [array]$Keys
   ); 
   # Expand all metrics related to keys contained in array step by step
   $Keys | % { if ($_) { $InObject = $InObject | Select -Expand $_ }};
   $InObject;
}

#
#  Convert Windows DateTime to Unix timestamp and return its
#
Function ConvertTo-UnixTime { 
   Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
      [PSObject]$EndDate
   ); 

   Begin   { 
      $StartDate = Get-Date -Date "01/01/1970"; 
   }  

   Process { 
      # Return unix timestamp
      (New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds; 
   }  
}


#
#  Make & return JSON, due PoSh 2.0 haven't Covert-ToJSON
#
Function Make-JSON {
   Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
      [PSObject]$InObject, 
      [array]$ObjectProperties, 
      [switch]$Pretty
   ); 
   Begin   {
               # Pretty json contain spaces, tabs and new-lines
               if ($Pretty) { $CRLF = "`n"; $Tab = "    "; $Space = " "; } else {$CRLF = $Tab = $Space = "";}
               # Init JSON-string $InObject
               $Result += "{$CRLF$Space`"data`":[$CRLF";
               # Take each Item from $InObject, get Properties that equal $ObjectProperties items and make JSON from its
               $itFirstObject = $True;
           } 
   Process {
               ForEach ($Object in $InObject) {
                  if (-Not $itFirstObject) { $Result += ",$CRLF"; }
                  $itFirstObject=$False;
                  $Result += "$Tab$Tab{$Space"; 
                  $itFirstProperty = $True;
                  # Process properties. No comma printed after last item
                  ForEach ($Property in $ObjectProperties) {
                     if (-Not $itFirstProperty) { $Result += ",$Space" }
                     $itFirstProperty = $False;
                     $Result += "`"{#$Property}`":$Space`"$($Object.$Property | Prepare-ToZabbix)`""
                  }
                  # No comma printed after last string
                  $Result += "$Space}";
               }
           }
  End      {
               # Finalize and return JSON
               "$Result$CRLF$Tab]$CRLF}";
           }
}

#
#  Return collection of GetComputerTargetGroups (all or selected by ID) or GetTotalSummaryPerComputerTarget (full or shrinked with condition)
#
Function Get-WSUSComputerTargetGroupInfo  { 
   Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
      [PSObject]$WSUS, 
      [string]$Action,
      [string]$Key,
      [string]$Id
   ); 

   Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] Retrieving target groups(s)"
   # Take all computer Groups with specific or Any ID
   $ComputerTargetGroups = $WSUS.GetComputerTargetGroups() | IDEqualOrAny $Id;
   If (-Not $ComputerTargetGroups) {
      Write-Error "Group with ID = '$Id' does not exist in WSUS!";
      Exit;
   }

   if ('Discovery' -eq $Action) {
      Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] Returning GetComputerTargetGroups collection"
      # Return all computer Groups 
      $ComputerTargetGroups;
   } else {
      Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] Taking group's GetTotalSummaryPerComputerTarget collection"
      # If no ID given - change local $Key value to any for calling switch's default section
      $ComputerTargets = $ComputerTargetGroups | % { $_.GetTotalSummaryPerComputerTarget() };
      # Analyzing Key and count how much computers present into collection from selection 
      Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] Filtering and return collection..."
      switch ($Key) {
         'ComputerTarget' {
             # Include all computers
             $ComputerTargets;
         }               
         'ComputerTargetsWithUpdateErrors' {
             # $ComputerTargets | Where { $_.FailedCount -gt 0 };
             # Include all failed (property FailedCount <> 0) computers
             $ComputerTargets | Where { 0 -ne $_.FailedCount };
         }                                                                                                                                       
         'ComputerTargetsNeedingUpdates' {
             #$ComputerTargets | Where { ($_.NotInstalledCount -gt 0 -Or $_.DownloadedCount -gt 0 -Or $_.InstalledPendingRebootCount -gt 0) -And $_.FailedCount -le 0}; 
             # Include no failed, but not installed, downloaded, pending reboot computers
             $ComputerTargets | Where { (0 -eq $_.FailedCount) -And (0 -ne ($_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount)) };
         }                                                         
         'ComputersUpToDate' {
#             $ComputerTargets | Where { $_.UnknownCount -eq 0 -And $_.NotInstalledCount -eq 0 -And $_.DownloadedCount -le 0 -And $_.InstalledPendingRebootCount -le 0 -And $_.FailedCount -le 0 };
             # include only no failed, unknown, not installed, downloaded, pending reboot
             $ComputerTargets | Where { 0 -eq ($_.FailedCount+$_.UnknownCount+$_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount) };
         }                    
         'ComputerTargetsUnknown' {
             #$ComputerTargets | Where { $_.UnknownCount -gt 0 -And $_.NotInstalledCount -le 0 -And $_.DownloadedCount -le 0 -And $_.InstalledPendingRebootCount -le 0 -And $_.FailedCount -le 0 };
             # include only unknown, but no failed, not installed, downloaded, pending reboot
             $ComputerTargets | Where { (0 -ne $_.UnknownCount) -And (0 -eq ($_.FailedCount+$_.NotInstalledCount+$_.DownloadedCount+$_.InstalledPendingRebootCount))};
         }
         default { $False; }
      }
   }
}

####################################################################################################################################
#
#                                                 Main code block
#    
####################################################################################################################################
Write-Verbose "$(Get-Date) Loading 'Microsoft.UpdateServices.Administration' Assembly"
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null
Write-Verbose "$(Get-Date) Trying to connect to local WSUS Server"
# connect on Local WSUS
$WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer();
if (!$WSUS)
   {
     Write-Warning "$(Get-Date) Connection failed";
     Exit;
   }
Write-Verbose "$(Get-Date) Connection established";

# if needProcess is False - $Result is not need to convert to string and etc 
$doAction = $True;
# split 
$Keys = $Key.split(".");

Write-Verbose "$(Get-Date) Creating collection of specified object: '$Object'";
switch ($Object) {
   'Info'                   { $Objects = $WSUS; }
   'Status'                 { $Objects = $WSUS.GetStatus(); }
   'Database'               { $Objects = $WSUS.GetDatabaseConfiguration(); }
   'Configuration'          { $Objects = $WSUS.GetConfiguration(); }
                              # $Key variable need here, not $Keys array
   'ComputerGroup'          { $Objects = $WSUS | Get-WSUSComputerTargetGroupInfo -Action $Action -Key $key -Id $Id; }
                              # | Select-Object make copy which used for properly Add-Member work
   'LastSynchronization'    { $Objects = ($WSUS.GetSubscription()).GetLastSynchronizationInfo() | Select-Object;
                              # Just add new Property for Virtual key "NotSyncInDays"
                              $Objects | % { $_ | Add-Member -MemberType NoteProperty -Name "NotSyncInDays" -Value (New-TimeSpan -Start $_.StartTime.DateTime -End (Get-Date)).Days }
                            }
                              # SynchronizationStatus contain one value
   'SynchronizationProcess' { $doAction = $False; $Result = ($WSUS.GetSubscription()).GetSynchronizationStatus(); }
   default                  { 
                              Write-Error "Unknown object: '$Object'";
                              Exit;
                            }
}  

Write-Verbose "$(Get-Date) Collection created";
#$Objects 

if ($doAction) { 
   Write-Verbose "$(Get-Date) Processing collection with action: '$Action'";
   switch ($Action) {
      # Discovery given object, make json for zabbix
      'Discovery' {
          switch ($Object) {
             'ComputerGroup' { $ObjectProperties = @("NAME", "ID"); }
          }
          Write-Verbose "$(Get-Date) Generating LLD JSON";
          $Result = $Objects | Make-JSON -ObjectProperties $ObjectProperties -Pretty;
      }
      # Get metrics or metric list
      'Get' {
         if ($Keys) { 
            Write-Verbose "$(Get-Date) Getting metric related to key: '$Key'";
            $Result = $Objects | Get-Metric -Keys $Keys;
         } else { 
            Write-Verbose "$(Get-Date) Getting metric list due metric's Key not specified";
            $Result = $Objects | fl *;
        };
      }
      # Count selected objects
      'Count' { 
          Write-Verbose "$(Get-Date) Counting objects";  
          # if result not null, False or 0 - return .Count
          $Result = $(if ($Objects) { @($Objects).Count } else { 0 } ); 
      }
      default  { 
          Write-Error "Unknown action: '$Action'";
          Exit;
      }
   }  
}

Write-Verbose "$(Get-Date) Converting Windows DataTypes to equal Unix's / Zabbix's";
switch (($Result.GetType()).Name) {
   'Boolean'  { $Result = [int]$Result; }
   'DateTime' { $Result = $Result | ConvertTo-UnixTime; }
   'Object[]' { $Result = $Result | Out-String; }
}

# Normalize String object
$Result = $Result.ToString().Trim();

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

"$Result";
