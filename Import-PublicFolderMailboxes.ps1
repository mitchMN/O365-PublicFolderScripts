# .SYNOPSIS
# Import-PublicFolderMailboxes.ps1
#    Import the public folder mailboxes as mail enabled users from cloud to on-premise 
#
# Example input to the script:
#
# Import-PublicFolderMailboxes.ps1 -ConnectionUri <cloud url> -Credential <credential> 
#
# The above example imports public folder mailbox objects from cloud as mail enabled users to on-premise.
#
# .DESCRIPTION
#
# Copyright (c) 2012 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID",

    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [PSCredential] $Credential
    )

## Create a tenant PSSession.
function GetTenantSession([string] $uri, [PSCredential] $cred)
{
    $sessionOption = (New-PSSessionOption -SkipCACheck);
    $session = New-PSSession -ConnectionURI:$uri `
                             -ConfigurationName:Microsoft.Exchange `
                             -AllowRedirection `
                             -Authentication:"Basic" `
                             -SessionOption:$sessionOption `
                             -Credential:$cred `
                             -ErrorAction:SilentlyContinue;
    return $session;
}

## Writes a dated information message to console
function WriteInfoMessage()
{
    param ($message)
    Write-Host "[$($(Get-Date).ToString())]" $message;
}

## Execute command
function ExecuteCommand(
    [string] $cmd,
    [System.Management.Automation.Runspaces.PSSession] $session)
{
    # This isn't vulnerable to PowerShell injection or isn't interesting to exploit. Justification: There is no user input params used here.
    $scriptBlock = $executioncontext.invokecommand.NewScriptBlock($cmd);
    if ($session -ne $null)
    {
        return (Invoke-Command -Session $session -ScriptBlock $scriptBlock);
    }

    return (Invoke-Command -ScriptBlock $scriptBlock);
}

## Retrieve public folder mailboxes
function GetPublicFolderMailBoxes(
    [System.Management.Automation.Runspaces.PSSession] $session)
{
    $publicFolderMailboxes = ExecuteCommand "Get-Mailbox -PublicFolder -ResultSize:Unlimited -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue" $session;

    # Return the results     
    if ($publicFolderMailboxes -eq $null -or ([array]($publicFolderMailboxes)).Count -lt 1)
    {
        return $null;
    }

    return $publicFolderMailboxes;
}

## Sync public folder mailboxes from cloud to on-prem.
function SyncPublicFolderMailboxes(
    [object[]] $publicFolderMailboxes)
{
    $validExternalEmailAddresses = @();

    if ($publicFolderMailboxes -ne $null)
    {
        $hasPublicFolderServingHierarchy = $false;
        foreach ($publicFolderMailbox in $publicFolderMailboxes)
        {
            if ($publicFolderMailbox.IsExcludedFromServingHierarchy -eq $false)
            {  
                $hasPublicFolderServingHierarchy = $true;
                $displayName = $publicFolderMailbox.Name.ToString().Trim();
                $name = "RemotePfMbx-" + $displayName + "-" + [guid]::NewGuid();
                $externalEmailAddress = $publicFolderMailbox.PrimarySmtpAddress.ToString();

                WriteInfoMessage ($LocalizedStrings.SyncingPublicFolderMailbox -f $displayName);

                $mailUser = Get-MailUser $externalEmailAddress -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue;

                if ($mailUser -eq $null)
                {
                    WriteInfoMessage ($LocalizedStrings.CreatingMailUser -f $displayName);
                    $mailUser = New-MailUser -Name $name -ExternalEmailAddress $externalEmailAddress -DisplayName $displayName;
                }
                else
                {
                    WriteInfoMessage ($LocalizedStrings.MailUserExists -f $mailUser);
                }

                WriteInfoMessage ($LocalizedStrings.ConfiguringMailUser -f $mailUser);

                Set-OrganizationConfig -RemotePublicFolderMailboxes @{Add=$mailUser};

                $validExternalEmailAddresses += $mailUser.ExternalEmailAddress;
                WriteInfoMessage ($LocalizedStrings.DoneSyncingPublicFolderMailbox -f $displayName);
                Write-Host "";
            }
        }
    }

    if (-not $hasPublicFolderServingHierarchy)
    {
        WriteInfoMessage ($LocalizedStrings.NoHierarchyPublicFolderMailbox);
        Write-Host "";
    }

    WriteInfoMessage ($LocalizedStrings.DeletingMailUsersInfo);
    $remoteMailboxes = Get-OrganizationConfig | select RemotePublicFolderMailboxes

    foreach($adObjectId in $remoteMailboxes.RemotePublicFolderMailboxes)
    {
        $mailUser = Get-MailUser $adObjectId -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue

        if ($mailUser -ne $null)
        {
            if (-not $validExternalEmailAddresses.Contains($mailUser.ExternalEmailAddress))
            {
                WriteInfoMessage ($LocalizedStrings.RemovingMailUsers -f $mailUser);
                Set-OrganizationConfig -RemotePublicFolderMailboxes @{Remove=$mailUser}

                WriteInfoMessage ($LocalizedStrings.DeleteMailUser -f $mailUser);
                Remove-MailUser $mailUser -Confirm:$false 
            }
        }
    }
}

#load hashtable of localized string
Import-LocalizedData -BindingVariable LocalizedStrings -FileName ImportPublicFolderMailboxes.strings.psd1

# Create a PSSession
$session = GetTenantSession -uri:$ConnectionUri -cred:$Credential;
if ($session -eq $null)
{
    WriteInfoMessage ($LocalizedStrings.IncorrectCredentials);
    exit;
}

WriteInfoMessage ($LocalizedStrings.StartedPublicFolderMailboxImport);
Write-Host "";

# Get mail enabled public folders in the organization
$publicFolderMailboxes = GetPublicFolderMailBoxes $session;

# Create mail enabled users for remote public folder mailboxes
SyncPublicFolderMailboxes $publicFolderMailboxes;

Write-Host "";
WriteInfoMessage ($LocalizedStrings.CompletedPublicFolderMailboxImport);

# Terminate the PSSession
Remove-PSSession $session;