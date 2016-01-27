#
# WSUS Miner
# zbx.sadman@gmail.com, 2016
#

Param (
[string]$Action='Get',
[string]$Object='Status',
# try "UpdateCount"
[string]$Key='UpdateCount'
)

[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null

Function Get-WSUSStatus
{ 
  Param ($Key);
  # connect on Local WSUS
  (([Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()).getstatus()).$Key;
}

switch ($Object) 
  {
    ('Status')  
        {
          $Result = Get-WSUSStatus($Key);
        }
    default  { $Result = "Incorrect object: '" + $Object + "'"; }
  }  

$Result = ($Result | Out-String).trim();
Write-Host $Result;
