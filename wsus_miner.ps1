<#
    .SYNOPSIS  
        Return WSUS metrics values, count selected objects, make LLD-JSON for Zabbix

    .DESCRIPTION
        Return WSUS metrics values, count selected objects, make LLD-JSON for Zabbix

    .NOTES  
        Version: 1.0
        Name: Microsoft's WSUS Miner
        Author: zbx.sadman@gmail.com
        DateCreated: 26FEB2016

    .LINK  
        https://github.com/zbx-sadman

    .PARAMETER Action
        What need to do with collection or its item:
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

   Process {
      # Expand all metrics related to keys contained in array step by step
      $Keys | % { if ($_) { $InObject = $InObject | Select -Expand $_ }};
   }

   End     { 
      $InObject;
   }
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
#  Prepare string to using with Zabbix 
#
Function Prepare-ToZabbix {
  Param (
     [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
     [PSObject]$InObject
  );
  $InObject = ($InObject.ToString());
  $InObject.Replace("`"", "\`"");
}


#
#  Return LastSynchronizationInfo object or object with metric, that named as virtual key
#
Function Get-WSUSLastSynchronizationInfo { 
   Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
      [PSObject]$WSUS, 
      [string]$Key
   ); 
   Write-Verbose "$(Get-Date) [Get-WSUSLastSynchronizationInfo] Taking LastSynchronizationInfo object"
   $Result = ($WSUS.GetSubscription()).GetLastSynchronizationInfo();

   # check for Key existience
   switch ($Key) {
      'NotSyncInDays'  { 
          Write-Verbose "$(Get-Date) [Get-WSUSLastSynchronizationInfo] Return new object with 'NotSyncInDays' property"
          New-Object PSObject -Property @{ $key = (New-TimeSpan -Start $Result.StartTime.DateTime -End (Get-Date)).Days };
      }
      # Otherwise - just return object
      default { 
         Write-Verbose "$(Get-Date) [Get-WSUSLastSynchronizationInfo] Return object"
         $Result; 
      }
   }  
}

#
#  Return collection of GetComputerTargetGroups (all or selected by ID) or GetTotalSummaryPerComputerTarget (full or shrinked with condition)
#
Function Get-WSUSComputerTargetGroupInfo  { 
   Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)] 
      [PSObject]$WSUS, 
      [string]$Key,
      [string]$Id
   ); 

   Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] Taking GetComputerTargetGroups collection"
   # Take all computer Groups with specific or Any ID
   $ComputerTargetGroups = $WSUS.GetComputerTargetGroups() | IDEqualOrAny $Id;

   if (!$Key) {
      Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] No Key specified, return collection"
      $ComputerTargetGroups;
   } else {
      Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] Taking GetTotalSummaryPerComputerTarget collection"
      $ComputerTargets = $ComputerTargetGroups.GetTotalSummaryPerComputerTarget();
      # Analyzing Key and count how much computers present into collection from selection 
      Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] Filtering..."
      switch ($key) {
         'ComputerTarget' {
             # All object must be counted
             #$ComputerTargets = $ComputerTargets;
         }               
         'ComputerTargetsWithUpdateErrors' {
             # Select and count all computers with property FailedCount > 0
             $ComputerTargets = $ComputerTargets | Where { $_.FailedCount -gt 0 };
         }
         'ComputerTargetsNeedingUpdates' {
             $ComputerTargets = $ComputerTargets | Where { ($_.NotInstalledCount -gt 0 -Or $_.DownloadedCount -gt 0 -Or $_.InstalledPendingRebootCount -gt 0) -And $_.FailedCount -le 0};
         }                    
         'ComputersUpToDate' {
             $ComputerTargets = $ComputerTargets | Where { $_.UnknownCount -eq 0 -And $_.NotInstalledCount -eq 0 -And $_.DownloadedCount -le 0 -And $_.InstalledPendingRebootCount -le 0 -And $_.FailedCount -le 0 };
         }                    
         'ComputerTargetsUnknown' {
             $ComputerTargets = $ComputerTargets | Where { $_.UnknownCount -gt 0 -And $_.NotInstalledCount -le 0 -And $_.DownloadedCount -le 0 -And $_.InstalledPendingRebootCount -le 0 -And $_.FailedCount -le 0 };
         }
         default { 
             $ComputerTargets = @();
         }
      }
      Write-Verbose "$(Get-Date) [Get-WSUSComputerTargetGroupInfo] Return collection"
      $ComputerTargets
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
$objWSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer();
if (!$objWSUS)
   {
     Write-Verbose "$(Get-Date) Connecting error reached";
     exit;
   }
Write-Verbose "$(Get-Date) Connected OK";

# if needProcess is False - $Result is not need to convert to string and etc 
$doAction = $True;
# split 
$Keys = $Key.split(".");

Write-Verbose "$(Get-Date) Create collection of specified object: '$Object'";
switch ($Object) {
   'Info'                   { $Objects = $objWSUS; }
   'Status'                 { $Objects = $objWSUS.GetStatus(); }
   'Database'               { $Objects = $objWSUS.GetDatabaseConfiguration(); }
   'Configuration'          { $Objects = $objWSUS.GetConfiguration(); }
                              # $Key variable need here, not $Keys array
   'ComputerGroup'          { $Objects = $objWSUS | Get-WSUSComputerTargetGroupInfo -Key $key -Id $Id; }
                              # $Key variable need here, not $Keys array
   'LastSynchronization'    { $Objects = $objWSUS | Get-WSUSLastSynchronizationInfo -Key $Key; }
                              # SynchronizationStatus contain one value
   'SynchronizationProcess' { $doAction = $False; $Result = ($objWSUS.GetSubscription()).GetSynchronizationStatus(); }
   default                  { $doAction = $False; $Result = "Incorrect object: '$Object'";}
}  

Write-Verbose "$(Get-Date) Collection created";
if ($doAction) { 
   Write-Verbose "$(Get-Date) Processeed collection with action: '$Action' ";
   switch ($Action) {
      #
      # Discovery given object, make json for zabbix
      #
      'Discovery' {
          switch ($Object) {
             'ComputerGroup' { $ObjectProperties = @("NAME", "ID"); }
          }
          Write-Verbose "$(Get-Date) Generating LLD JSON";
          $Result = $Objects | Make-JSON -ObjectProperties $ObjectProperties -Pretty;
      }
      'Get' {
         if ($Keys) { 
            Write-Verbose "$(Get-Date) Get metric related to key: '$Key'";
            $Result = $Objects | Get-Metric -Keys $Keys;
         } else { 
            Write-Verbose "$(Get-Date) Get metric list due metric's Key not specified";
            $Result = $Objects | fl *;
        };
      }
      'Count' { 
          Write-Verbose "$(Get-Date) Count objects";
          $Result = @($Objects).Count; 
      }
      default  { 
        $Result = "Incorrect action: '$Action'"; 
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
   Write-Verbose "$(Get-Date) Change console width to $CONSOLE_WIDTH";
   mode con cols=$CONSOLE_WIDTH; 
}

Write-Verbose "$(Get-Date) Finished";

"$Result";
