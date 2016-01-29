#
# WSUS Miner
# zbx.sadman@gmail.com, 2016
#

Param (
[string]$Action,
[string]$Object,
[string]$Key,
[string]$Id,
[string]$consoleCP
)

function ConvertTo-Encoding ([string]$From, [string]$To){  
    Begin{  
        $encFrom = [System.Text.Encoding]::GetEncoding($from)  
        $encTo = [System.Text.Encoding]::GetEncoding($to)  
    }  
    Process{  
        $bytes = $encTo.GetBytes($_)  
        $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)  
        $encTo.GetString($bytes)  
    }  
}

Function ConvertTo-UnixTime { 
    Begin{  
        $StartDate = Get-Date -Date "01/01/1970";
    }  
    Process{  
       (New-TimeSpan -Start $StartDate -End $_).TotalSeconds;
    }  
}


Function Make-JSON {
  Param ([PSObject]$InObject, [array]$ObjectProperties, [boolean]$Pretty);
  # Pretty json contain spaces, tabs and new-lines
  if ($Pretty) { $CRLF = "`n"; $Tab = "    "; $Space = " "; } else {$CRLF = $Tab = $Space = "";}
  # Init JSON-string $InObject
  $Result = "{$CRLF$Space`"data`":[$CRLF";
  # Take each Item from $InObject, get Properties that equal $ObjectProperties items and make JSON from its
  $k = 0;
  ForEach ($Object in $InObject) {$k++;
     $Result += "$Tab$Tab{$Space";
     # Process properties. No comma printed after last item
     $ObjectProperties | % {$i = 0} {$i++; $Result += "`"{#$_}`":$Space`""+$Object.$_+"`""+(&{if ($i -lt $ObjectProperties.Count) {",$Space"} })}
     # No comma printed after last string
     $Result += " }"+(&{if ($k -lt $InObject.Count) {",$Space"} })+"$CRLF";
  }
  # Finalize and return JSON
  "$Result$Space]$CRLF}";
}


Function Get-WSUSLastSynchronizationInfo ([PSObject]$WSUS, [string]$key) { 
  $Result = ($WSUS.GetSubscription()).GetLastSynchronizationInfo();
  # check for Key existience
  switch ($key) {
     # Key is 'StartTime'. Need to make Unix timestamp from DateTime object and replace 'StartTime' Property with it
     ('StartTime') { New-Object PSObject -Property @{ $key = ($Result.$key.DateTime | ConvertTo-UnixTime) }}; 
     # Otherwise - just return object
     default {$Result;}
  }  
  
}

Function Get-WSUSTotalSummaryPerGroupTarget ([PSObject]$WSUS, [string]$key, [string]$Id) { 
  # Take all computers into Group with specific ID
  $ComputerTargets = ($objWSUS.GetComputerTargetGroup($Id)).GetTotalSummaryPerComputerTarget()

  # Analyzing Key and count how much computers present into collection from selection 
  switch ($key) {
     ('ComputerTargetsWithUpdateErrorsCount') {
         # Select and count all computers with property FailedCount > 0
         $Result = (($ComputerTargets | Where { $_.FailedCount -gt 0 }) | measure).Count
     }
     ('ComputerTargetsNeedingUpdatesCount') {
        $Result = (($ComputerTargets | Where { ($_.NotInstalledCount -gt 0 -Or $_.DownloadedCount -gt 0 -Or $_.InstalledPendingRebootCount -gt 0) -And $_.FailedCount -le 0}) | measure).Count
     }                    
     ('ComputersUpToDateCount') {
         # why .Count without `measure` not give 1 with UnknownCount = 0?  ("only warpig in select" case)
         $Result = (($ComputerTargets | Where { $_.UnknownCount -eq 0 -And $_.NotInstalledCount -eq 0 -And $_.DownloadedCount -le 0 -And $_.InstalledPendingRebootCount -le 0 -And $_.FailedCount -le 0 }) | measure).Count   
     }                    
     ('ComputerTargetsUnknownCount') {
         $Result = (($ComputerTargets | Where { $_.UnknownCount -gt 0 -And $_.NotInstalledCount -le 0 -And $_.DownloadedCount -le 0 -And $_.InstalledPendingRebootCount -le 0 -And $_.FailedCount -le 0 }) | measure).Count
     }               
  }

  # If no computers is selected - collection will be empty and $Result will be undefined. Need to make undefined to 0
  $Result = (&{if ($Result) { $Result } else { 0 }});
  # return new object with property that named as key and contain number of computers into selection
  New-Object PSObject -Property @{ $key = $Result }; 
}


[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null
# connect on Local WSUS
$objWSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer();

# if needProcess is False - $Result is not need to convert to string and etc 
$needProcess = $True;

switch ($Action) {
     #
     # Discovery given object, make json for zabbix
     #
     ('Discovery') {
         $needProcess = $False;
         switch ($Object) {
            ('ComputerGroup')                  { $ObjectProperties = @("NAME", "ID"); $InObject = $objWSUS.GetComputerTargetGroups(); }
         }
         $Result = Make-JSON -InObject $InObject -ObjectProperties $ObjectProperties -Pretty $True;
     }
     #
     # Get metrics from object (real or virtual)
     #
     ('Get') {
        switch ($Object) {
            ('Info')                   { $Result = $objWSUS; }
            ('Status')                 { $Result = $objWSUS.GetStatus(); }
            ('Database')               { $Result = $objWSUS.GetDatabaseConfiguration(); }
            ('Configuration')          { $Result = $objWSUS.GetConfiguration(); }
            ('ComputerGroup')          { $Result = Get-WSUSTotalSummaryPerGroupTarget -WSUS $objWSUS -Key $key -Id $Id; }
            ('LastSynchronization')    { $Result = Get-WSUSLastSynchronizationInfo -WSUS $objWSUS -Key $key ; }
            ('SynchronizationProcess') { $needProcess = $False; $Result = ($objWSUS.GetSubscription()).GetSynchronizationStatus(); }
            default                    { $needProcess = $False; $Result = "Incorrect object: '$Object'";}
        }  
        if ($needProcess -And $Key) { $Result = ($Result.$Key).ToString(); }
     }
     #
     # Error
     #
     default  { $Result = "Incorrect action: '$Action'"; }
}  

# Normalize String object
$Result = ($Result | Out-String).trim();

# Convert String to UTF-8 if need (For Zabbix LLD-JSON with Cyrillic for example)
if ($consoleCP) { $Result = $Result | ConvertTo-Encoding -From $consoleCP -To UTF-8 }

Write-Host $Result;

