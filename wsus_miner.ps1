#
# WSUS Miner
# zbx.sadman@gmail.com, 2016
#

Param (
[string]$Action = 'Get',
[string]$Object = 'Status',
# try "UpdateCount"
[string]$Key = 'UpdateCount'
)

[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null

Function Connect-WSUSServer
{ 
  # connect on Local WSUS
  $WSUS = ([Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer());
  $WSUS;
}

Function Get-WSUSStatus
{ 
  Param ($key, $WSUS);
  ($WSUS.getstatus()).$Key;
}

Function Get-WSUSInfo
{ 
  Param ($key, $WSUS);
  ($WSUS.$Key) -join '.';
}


switch ($Object) 
  {
    ('Status')  
        {
          $objWSUS = Connect-WSUSServer;
          $Result = Get-WSUSStatus -Key $key -WSUS $objWSUS;
        }
    ('Info')  
        {
          $objWSUS = Connect-WSUSServer;
          $Result = Get-WSUSInfo -Key $key -WSUS $objWSUS;
        }
    default  { $Result = "Incorrect object: '$Object'"; }
  }  

$Result = ($Result | Out-String).trim();
Write-Host $Result;

