<#
    .SYNOPSIS  
        Helps to do mass action with Updates: accept license agreement, approve for computer group.
        GUI used.

    .DESCRIPTION
        Helps to do mass action with Updates: accept license agreement, approve for computer group
        GUI used.

    .NOTES  
        Version: 0.9.1
        Name: Microsoft's WSUS Mass Action Helper
        Author: zbx.sadman@gmail.com
        DateCreated: 04AUG2016
        Testing environment: Windows Server 2008R2 SP1, WSUS 3 SP2, Powershell 2

    .LINK  
        https://github.com/zbx-sadman

#>

Set-StrictMode –Version Latest

# Set US locale to properly formatting float numbers while converting to string
[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"

function ApproveUpdatesToGroupButton_onClick()
{
  if (0 -lt @($UpdatesListView.CheckedItems).Count) { 
     $Weight = 100 / @($UpdatesListView.CheckedItems).Count; $Count = 0; 
     ForEach ($CheckedUpdate in $UpdatesListView.CheckedItems) {
        $Update = $WSUS.GetUpdate([GUID] $CheckedUpdate.Tag);
        if ($True -Eq $Update.RequiresLicenseAgreementAcceptance) {
           $Update.AcceptLicenseAgreement();
        }

        ForEach ($CheckedGroup in $ComputerGroupsListView.CheckedItems) {
           $Group = $WSUS.GetComputerTargetGroup([GUID] $CheckedGroup.Tag)
           $Update.Approve([Microsoft.UpdateServices.Administration.UpdateApprovalAction]::Install, $Group);
         }
        $Count += $Weight; if (100 -lt $Count) {$Count = 100;}
        $ProgressBar.Value = $Count;
    }
   $ProgressBar.Value = 100;
  }
}

function onResize_Form() 
{

$MainUtilityForm.SuspendLayout();

$ProgressBar.Location = New-Object System.Drawing.Point(4, ($MainUtilityForm.Height - 36))
$ProgressBar.Height = 16;
$ProgressBar.Width = ($MainUtilityForm.Width - 28);

# 50pix on form bottom - button place
$ButtonTop = ($MainUtilityForm.Height - 70);
$NewGroupBoxHeight = ($ButtonTop - 20) / 3

$ApproveUpdatesToGroupButton.Location = New-Object System.Drawing.Point(4, $ButtonTop);
$ExitButton.Location = New-Object System.Drawing.Point(($MainUtilityForm.Width - $ExitButton.Width - 16), $ButtonTop);

# GroupBoxes
$UpdatesOwnerGroupBox.Location = New-Object System.Drawing.Point(4, 4)
$UpdatesOwnerGroupBox.Width = ($UpdatesOwnerGroupBox.Parent.Width - 16);
$UpdatesOwnerGroupBox.Height = $NewGroupBoxHeight;

$UpdatesGroupBox.Location = New-Object System.Drawing.Point(4, ($UpdatesOwnerGroupBox.Bottom + 4))

$UpdatesGroupBox.Width = ($UpdatesGroupBox.Parent.Width - 16);
$UpdatesGroupBox.Height = $NewGroupBoxHeight;

$ComputerGroupsGroupBox.Location = New-Object System.Drawing.Point(4, ($UpdatesGroupBox.Bottom + 4));
$ComputerGroupsGroupBox.Width = ($ComputerGroupsGroupBox.Parent.Width - 16);
$ComputerGroupsGroupBox.Height = $NewGroupBoxHeight;

# GroupBoxes content
$ComputerGroupsTree.Location = New-Object System.Drawing.Point(4, 16);
$ComputerGroupsTree.Height = ($ComputerGroupsTree.Parent.ClientSize.Height - 20);
$ComputerGroupsTree.Width = $ComputerGroupsTree.Parent.Width / 3;

$ComputersListView.Location = New-Object System.Drawing.Point(($ComputersListView.Parent.Width/3 + 8), 16);
$ComputersListView.Height = ($ComputersListView.Parent.Height - 20);
$ComputersListView.Width = (($ComputersListView.Parent.Width/3*2) - 12);


$UpdatesListView.Location = New-Object System.Drawing.Point(4, 16);
$UpdatesListView.Height = ($UpdatesListView.Parent.Height - $UpdateInfoLabel.Height - 24);
$UpdatesListView.Width = ($UpdatesListView.Parent.Width - 8);

$UpdateInfoLabel.Location = New-Object System.Drawing.Point(4, ($UpdatesListView.Bottom + 4))

$ComputerGroupsListView.Location = New-Object System.Drawing.Point(4, 16);
$ComputerGroupsListView.Height = ($ComputerGroupsListView.Parent.Height - 20);
$ComputerGroupsListView.Width = ($ComputerGroupsListView.Parent.Width - 8);

$MainUtilityForm.ResumeLayout();
}

function UpdatesContextMenuItem_onClick()
{
  $Update = $WSUS.GetUpdate([GUID] $UpdatesListView.SelectedItems[0].Tag);
  Write-Host $([String]$Update.AdditionalInformationUrls);
  [System.Diagnostics.Process]::Start([String]$Update.AdditionalInformationUrls);
}

function UpdatesContextMenuItemOpenURL_onClick()
{
  $Update = $WSUS.GetUpdate([GUID] $UpdatesListView.SelectedItems[0].Tag);
  [System.Diagnostics.Process]::Start([String]$Update.AdditionalInformationUrls);
}

function UpdatesContextMenuItemCopyURLToClipboard_onClick()
{
  if ($Null -ne $UpdatesListView.SelectedItems) {
    $Update = $WSUS.GetUpdate([GUID] $UpdatesListView.SelectedItems[0].Tag);
#    [String]$Update.AdditionalInformationUrls | clip.exe;
    $Update | clip.exe;
  }
}

function UpdatesContextMenuItemCopyUpdatePropertiesToClipboard_onClick()
{
  if ($Null -ne $UpdatesListView.SelectedItems) {
    $WSUS.GetUpdate([GUID] $UpdatesListView.SelectedItems[0].Tag) | clip.exe;
  }
}

function ComputersListViewCopyComputerPropertiesToClipboard_onClick()
{
  if ($Null -ne $ComputersListView.SelectedItems) {
    $WSUS.GetComputerTarget([GUID] $ComputersListView.SelectedItems[0].Tag) | clip.exe;
  }
}


function ComputersListView_onMouseClick()
{
  If ("Left" -ne $_.Button) { Return };
  $ProgressBar.Value = 0;
  $UpdatesListView.BeginUpdate();
  $hit = $ComputersListView.HitTest($_.Location)
  if ($hit.Item) {
     [void] $UpdateScope.ApprovedComputerTargetGroups.Clear();
     $NotInstalledUpdateIDs = @();
     $ApprovedUpdateIDs = @();
     $UnapprovedUpdateIDs = @();
     $ComputerTarget = $WSUS.GetComputerTarget([GUID] $hit.Item.Tag);
     $ComputerGroup = $WSUS.GetComputerTargetGroup([GUID] $ComputerGroupsTree.SelectedNode.Tag);
     # item clicked
     # Get not-installed updates for selected computer
     $UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled;
     $NotInstalledUpdateIDs = $ComputerTarget.GetUpdateInstallationInfoPerUpdate($UpdateScope) | % { $_.UpdateID };
#     Write-Host $($ComputerTarget).FullDomainName
#     Write-Host $($ComputerGroup).Name
     if ($Null -ne $NotInstalledUpdateIDs) {
        # Prepare group list for GetUpdateApprovals() request 
        # Add to list computer's group and its parent groups
        While ($computerGroup -ne $rootComputerGroup) {
           [void] $UpdateScope.ApprovedComputerTargetGroups.Add($computerGroup);
           $computerGroup = $computerGroup.GetParentTargetGroup();
        } 
        # finally add root group "All Computers"
        [void] $updateScope.ApprovedComputerTargetGroups.Add($rootComputerGroup);

        # Take all approved updates  
        $UpdateScope.ApprovedStates="Any";
        $ApprovedUpdateIDs = $WSUS.GetUpdateApprovals($UpdateScope) | % { $_.UpdateID.UpdateID };
        if ($Null -ne $ApprovedUpdateIDs) {
           # Select all updates that exist in $NotInstalledUpdateIDs, but not exist in $ApprovedUpdateIDs
           $UnapprovedUpdateIDs = Compare-Object $ApprovedUpdateIDs $NotInstalledUpdateIDs | ? { $_.SideIndicator -eq "=>" } | %{ [GUID] $_.InputObject }
        }
     }

     $UpdateInfoLabel.Text ="Unapproved: $(@($UnapprovedUpdateIDs).Count)";
     $UpdatesListView.Items.Clear();
     if (0 -lt @($UnapprovedUpdateIDs).Count) { 
        $Weight = 100 / @($UnapprovedUpdateIDs).Count; $Count = 0; 
        ForEach ($Id in $UnapprovedUpdateIDs) {
           if ([String]::IsNullorEmpty($Id)) { Continue; }

           $Update = $WSUS.GetUpdate($Id);
           $Item = New-Object System.Windows.Forms.ListViewItem([String]$Update.Title)
           $Item.SubItems.Add([String]$Update.State);
           $Item.SubItems.Add([String]$Update.AdditionalInformationUrls);
           $Item.SubItems.Add([String]$Update.HasLicenseAgreement);
#           $Item.SubItems.Add([String]$Update.IsApproved);
#           $Item.SubItems.Add([String]$Update.IsDeclined);
           $Item.Tag = [String] $Id;
           [void] $UpdatesListView.Items.Add($Item);
           $Count += $Weight; if (100 -lt $Count) {$Count = 100;}
           $ProgressBar.Value = $Count;
        }
        $ProgressBar.Value = 100;
      }
   }
  $UpdatesListView.EndUpdate();

  ForEach ($Item in $ComputerGroupsListView.Items) {
     if ($Item.Tag -eq $ComputerGroupsTree.SelectedNode.Tag) {
        $Item.Checked = $True;
     }
  }

}


function ComputerGroupsTreeNode_onNodeMouseClick()
{
  ComputerGroupsTreeNode_SelectAction -GroupID $_.Node.Tag
}

function ComputerGroupsTreeNode_onKeyDown()
{
  if ($_.KeyCode -eq "Enter") {
     ComputerGroupsTreeNode_SelectAction -GroupID $ComputerGroupsTree.SelectedNode.Tag;
#     $UpdateInfoLabel.Text = $ComputerGroupsTree.SelectedNode.Tag;
  }
}

function ComputerGroupsTreeNode_SelectAction()
{
   Param (
      [GUID]$GroupID
   ); 

  $ProgressBar.Value = 0;
  $UpdatesListView.Items.Clear();
  $ComputersListView.BeginUpdate()
  
  $ComputersListView.Items.Clear();
  $ComputerTargets = $($WSUS.GetComputerTargetGroup($GroupID)).GetComputerTargets();
  if (0 -lt @($ComputerTargets).Count) { 
     $Weight = 100 / @($ComputerTargets).Count; $Count = 0; 
     ForEach ($ComputerTarget in $ComputerTargets) {
        $Item = New-Object System.Windows.Forms.ListViewItem([String]$ComputerTarget.FullDomainName)
        $Item.SubItems.Add([String]$ComputerTarget.IPAddress);
        $Item.SubItems.Add([String]$ComputerTarget.OSDescription);
        $Item.Tag = $ComputerTarget.Id;
        [void] $ComputersListView.Items.Add($Item);
        $Count += $Weight; if (100 -lt $Count) {$Count = 100;}
        $ProgressBar.Value = $Count;
     }
     $ProgressBar.Value = 100;
   }
   $ComputersListView.EndUpdate();

}


function MakeComputerGroupsTree()
{
   Param (
      [Microsoft.UpdateServices.Internal.BaseApi.ComputerTargetGroup]$Group,
      [Object]$TreeNode 
   ); 
   If ($Null -eq $Group) {
     $Group = $WSUS.GetComputerTargetGroups() | Where {$_.Name –eq "All Computers"}
     $TreeNode = $ComputerGroupsTree.Nodes.Add($Group.Name);
     $TreeNode.Tag = $Group.ID;
   }
   $Childs = $Group.GetChildTargetGroups();
   if ($Childs) {
      ForEach ($Child in $Childs) {
         $Item = $TreeNode.Nodes.Add($Child.Name);
         $Item.Tag = ([string] $Child.Id);
         MakeComputerGroupsTree -Group $Child -TreeNode $Item;
      }
   }
}

function MakeComputerGroupsList()
{
   Param (
      [Microsoft.UpdateServices.Internal.BaseApi.ComputerTargetGroup]$Group,
      [Int32]$Level = 0
   ); 
   If ($Null -eq $Group) {
     $Group = $WSUS.GetComputerTargetGroups() | Where {$_.Name –eq "All Computers"}
     $Item = $ComputerGroupsListView.Items.Add($Group.Name);
     $Item.Tag = $Group.ID;
   }
   $Childs = $Group.GetChildTargetGroups()
   if ($Childs) {
      ForEach ($Child in $Childs) {
#         $Item = New-Object System.Windows.Forms.ListViewItem([String] $Child.Name);         
         $Item = $ComputerGroupsListView.Items.Add($Child.Name);
         $Item.Tag = $Child.Id;
         MakeComputerGroupsList -Group $Child -Level ($Level+1);
      }
   }
}




Add-Type -assembly System.Windows.Forms

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
   Write-Error "$(Get-Date) Connection failed";
}

Write-Verbose "$(Get-Date) Connection established";
$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled
$rootComputerGroup = $WSUS.GetComputerTargetGroups() | ? { $_.Name -eq 'All Computers'}


$MainUtilityForm = New-Object System.Windows.Forms.Form
$MainUtilityForm.StartPosition  = "WindowsDefaultLocation"
$MainUtilityForm.Text = "WSUS Utility"
$MainUtilityForm.Width = 700
$MainUtilityForm.Height = 700
$MainUtilityForm.ControlBox = $True;
$MainUtilityForm.KeyPreview = $True;
#$MainUtilityForm.AutoSize = $True
#$MainUtilityForm.AutoSizeMode = "GrowAndShrink"

$ComputerGroupsTree = New-Object System.Windows.Forms.TreeView;
$UpdateInfoLabel = New-Object System.Windows.Forms.Label;
#$CheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$UpdatesListView = New-Object System.Windows.Forms.ListView;
$ComputersListView = New-Object System.Windows.Forms.ListView;
$ComputerGroupsListView = New-Object System.Windows.Forms.ListView;
$UpdatesOwnerGroupBox = New-Object System.Windows.Forms.GroupBox;
$UpdatesGroupBox = New-Object System.Windows.Forms.GroupBox;
$UpdatesGroupBox = New-Object System.Windows.Forms.GroupBox;
$ComputerGroupsGroupBox = New-Object System.Windows.Forms.GroupBox;
$ProgressBar = New-Object System.Windows.Forms.ProgressBar;
$UpdatesContextMenu = New-Object System.Windows.Forms.ContextMenu;
$ComputersListViewContextMenu = New-Object System.Windows.Forms.ContextMenu;


$ComputersListViewCopyComputerPropertiesToClipboard = New-Object System.Windows.Forms.MenuItem;
$ComputersListViewCopyComputerPropertiesToClipboard.Text = "Copy Computer properties to clipboard";
$ComputersListViewCopyComputerPropertiesToClipboard.add_Click({ ComputersListViewCopyComputerPropertiesToClipboard_onClick });
[void] $ComputersListViewContextMenu.MenuItems.Add($ComputersListViewCopyComputerPropertiesToClipboard);
$ComputersListView.ContextMenu = $ComputersListViewContextMenu;

$UpdatesContextMenuItemOpenURL = New-Object System.Windows.Forms.MenuItem;
$UpdatesContextMenuItemOpenURL.Text = "Open informational link";
$UpdatesContextMenuItemOpenURL.add_Click({ UpdatesContextMenuItemOpenURL_onClick });

$UpdatesContextMenuItemCopyURLToClipboard = New-Object System.Windows.Forms.MenuItem;
$UpdatesContextMenuItemCopyURLToClipboard.Text = "Copy informational link to clipboard";
$UpdatesContextMenuItemCopyURLToClipboard.add_Click({ UpdatesContextMenuItemCopyURLToClipboard_onClick });

$UpdatesContextMenuItemCopyUpdatePropertiesToClipboard = New-Object System.Windows.Forms.MenuItem;
$UpdatesContextMenuItemCopyUpdatePropertiesToClipboard.Text = "Copy Update properties to clipboard";
$UpdatesContextMenuItemCopyUpdatePropertiesToClipboard.add_Click({ UpdatesContextMenuItemCopyUpdatePropertiesToClipboard_onClick });

[void] $UpdatesContextMenu.MenuItems.Add($UpdatesContextMenuItemCopyUpdatePropertiesToClipboard);
[void] $UpdatesContextMenu.MenuItems.Add($UpdatesContextMenuItemCopyURLToClipboard);
[void] $UpdatesContextMenu.MenuItems.Add($UpdatesContextMenuItemOpenURL);
$UpdatesListView.ContextMenu = $UpdatesContextMenu;

$UpdatesOwnerGroupBox.Text = "Updates owner";
$UpdatesGroupBox.Text = "Updates";
$ComputerGroupsGroupBox.Text = "Groups to action";
#$MoveToGroupButton = New-Object System.Windows.Forms.Button;
$ApproveUpdatesToGroupButton = New-Object System.Windows.Forms.Button;
$ApproveUpdatesToGroupButton.Text = 'Approve to..';
$ExitButton = New-Object System.Windows.Forms.Button;
$ExitButton.Text = 'Exit';

#$UpdateInfoLabel.Text ='Not installed: 0    Unapproved: 0';
$UpdateInfoLabel.Text ='Unapproved: 0';

$UpdatesListView.View = 'Details';
$ComputersListView.View = 'Details';
$ComputerGroupsListView.View = 'Details';

$MainUtilityForm.MaximizeBox = $True;

$UpdateInfoLabel.AutoSize = $True;

$UpdatesOwnerGroupBox.Controls.Add($ComputerGroupsTree);
$UpdatesOwnerGroupBox.Controls.Add($ComputersListView);
$MainUtilityForm.Controls.Add($UpdatesOwnerGroupBox);

$UpdatesGroupBox.Controls.Add($UpdatesListView);
$UpdatesGroupBox.Controls.Add($UpdateInfoLabel);

$MainUtilityForm.Controls.Add($UpdatesGroupBox);

$ComputerGroupsGroupBox.Controls.Add($ComputerGroupsListView);
$MainUtilityForm.Controls.Add($ComputerGroupsGroupBox);

$MainUtilityForm.Controls.add($ProgressBar);

$MainUtilityForm.Controls.Add($ApproveUpdatesToGroupButton);
$MainUtilityForm.Controls.Add($ExitButton);
$UpdatesListView.CheckBoxes = $True;
$UpdatesListView.GridLines = $True;
$UpdatesListView.FullRowSelect = $True;

$ComputersListView.Sorting = "Ascending";
$ComputersListView.CheckBoxes = $False;
$ComputersListView.GridLines = $True;
$ComputersListView.FullRowSelect = $True;
$ComputersListView.HideSelection = $False;
[void] $ComputersListView.Columns.Add('Name', 150);
[void] $ComputersListView.Columns.Add('IP', 100);
[void] $ComputersListView.Columns.Add('OS', 200);
#$ComputersListView.Sorted = $True;
$ComputersListView.MultiSelect = $False;

[void] $UpdatesListView.Columns.Add('Title', 450);
[void] $UpdatesListView.Columns.Add('State', 50);
[void] $UpdatesListView.Columns.Add('URI', 150);
[void] $UpdatesListView.Columns.Add('Lic', 50);
#[void] $UpdatesListView.Columns.Add('Approved', 50);
#[void] $UpdatesListView.Columns.Add('Declined', 50);

$UpdatesListView.HideSelection = $False;

[void] $ComputerGroupsListView.Columns.Add('Name', 500);
$ComputerGroupsListView.HideSelection = $False;

$ComputerGroupsTree.Sorted = $True;
$ComputerGroupsTree.HideSelection = $False;

$ComputerGroupsTree.Add_NodeMouseClick( { ComputerGroupsTreeNode_onNodeMouseClick } );
# Add_KeyDown, not Add_KeyPress
$ComputerGroupsTree.Add_KeyDown( { ComputerGroupsTreeNode_onKeyDown } );

$ComputersListView.Add_MouseClick( { ComputersListView_onMouseClick } );
$ExitButton.Add_Click( { $MainUtilityForm.Close() } );
$ApproveUpdatesToGroupButton.Add_Click( { ApproveUpdatesToGroupButton_onClick } );
   

$ComputerGroupsListView.CheckBoxes = $True;
$ComputerGroupsListView.GridLines = $True;
$ComputerGroupsListView.FullRowSelect = $True;

$MainUtilityForm.Add_Resize({ onResize_Form });
#$MainUtilityForm.Add_KeyDown({if ($_.KeyCode -eq "Escape")  {$MainUtilityForm.Close()}})

MakeComputerGroupsTree -TreeNode $ComputerGroupsTree;
MakeComputerGroupsList;
onResize_Form;
[void] $MainUtilityForm.ShowDialog()
