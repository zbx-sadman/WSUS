#
# WSUS Miner
# zbx.sadman@gmail.com, 2016
#

Param (
[string]$Action,
[string]$Object,
[string]$Key
)

Function Convert-ToUnixTime { 
  Param ($dateTime);
  (New-TimeSpan -Start (Get-Date -Date "01/01/1970") -End $dateTime).TotalSeconds;
}

Function Get-WSUSLastSynchronizationInfo { 
  Param ($key, $WSUS);
  $Result = ($WSUS.GetSubscription()).GetLastSynchronizationInfo();

  if ($Key) { 
    $Result = $Result.$Key; 
    switch ($key) {
       ('StartTime') {
          Convert-ToUnixTime -DateTime $Result.DateTime;
       }
       default  { $Result; }
    }  
  } else {
    $Result;
  }
  
}

[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
# connect on Local WSUS
$objWSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer();
$needProcessKey = $True;

switch ($Object) {
    ('Status') {
        $Result = $objWSUS.GetStatus();
    }

    ('Database') {
        $Result = $objWSUS.GetDatabaseConfiguration();
    }

    ('Info') {
        $Result = $objWSUS;
    }

    ('LastSynchronization') {
        $Result = Get-WSUSLastSynchronizationInfo -Key $key -WSUS $objWSUS;
    }

    ('SynchronizationProcess') {
          $needProcessKey = $False;
          $Result = ($objWSUS.GetSubscription()).GetSynchronizationStatus();
    }

    default  { 
          $needProcessKey = $False;
          $Result = "Incorrect object: '$Object'"; 
    }
}  

if ($needProcessKey -And $Key) { $Result=($Result.$Key).ToString(); }
$Result = ($Result | Out-String).trim();
Write-Host $Result;

