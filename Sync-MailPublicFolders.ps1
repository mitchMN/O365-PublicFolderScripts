# .SYNOPSIS
#    Syncs mail-enabled public folder objects from the local Exchange deployment into O365. It uses the local Exchange deployment
#    as master to determine what changes need to be applied to O365. The script will create, update or delete mail-enabled public
#    folder objects on O365 Active Directory when appropriate.
#
# .DESCRIPTION
#    The script must be executed from an Exchange 2007 or 2010 Management Shell window providing access to mail public folders in
#    the local Exchange deployment. Then, using the credentials provided, the script will create a session against Exchange Online,
#    which will be used to manipulate O365 Active Directory objects remotely.
#
#    Copyright (c) 2014 Microsoft Corporation. All rights reserved.
#
#    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# .PARAMETER Credential
#    Exchange Online user name and password.
#
# .PARAMETER CsvSummaryFile
#    The file path where sync operations and errors will be logged in a CSV format.
#
# .PARAMETER ConnectionUri
#    The Exchange Online remote PowerShell connection uri. If you are an Office 365 operated by 21Vianet customer in China, use "https://partner.outlook.cn/PowerShell".
#
# .PARAMETER Confirm
#    The Confirm switch causes the script to pause processing and requires you to acknowledge what the script will do before processing continues. You don't have to specify
#    a value with the Confirm switch.
#
# .PARAMETER Force
#    Force the script execution and bypass validation warnings.
#
# .PARAMETER WhatIf
#    The WhatIf switch instructs the script to simulate the actions that it would take on the object. By using the WhatIf switch, you can view what changes would occur
#    without having to apply any of those changes. You don't have to specify a value with the WhatIf switch.
#
# .EXAMPLE
#    .\Sync-MailPublicFolders.ps1 -Credential (Get-Credential) -CsvSummaryFile:sync_summary.csv
#    
#    This example shows how to sync mail-public folders from your local deployment to Exchange Online. Note that the script outputs a CSV file listing all operations executed, and possibly errors encountered, during sync.
#
# .EXAMPLE
#    .\Sync-MailPublicFolders.ps1 -Credential (Get-Credential) -CsvSummaryFile:sync_summary.csv -ConnectionUri:"https://partner.outlook.cn/PowerShell"
#    
#    This example shows how to use a different URI to connect to Exchange Online and sync mail-public folders from your local deployment.
#
param(
    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential] $Credential,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string] $CsvSummaryFile,
    
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID",

    [Parameter(Mandatory=$false)]
    [bool] $Confirm = $true,

    [Parameter(Mandatory=$false)]
    [switch] $Force = $false,

    [Parameter(Mandatory=$false)]
    [switch] $WhatIf = $false
)

# Writes a dated information message to console
function WriteInfoMessage()
{
    param ($message)
    Write-Host "[$($(Get-Date).ToString())]" $message;
}

# Writes a dated warning message to console
function WriteWarningMessage()
{
    param ($message)
    Write-Warning ("[{0}] {1}" -f (Get-Date),$message);
}

# Writes a verbose message to console
function WriteVerboseMessage()
{
    param ($message)
    Write-Host "[VERBOSE] $message" -ForegroundColor Green -BackgroundColor Black;
}

# Writes an error importing a mail public folder to the CSV summary
function WriteErrorSummary()
{
    param ($folder, $operation, $errorMessage, $commandtext)

    WriteOperationSummary $folder.Guid $operation $errorMessage $commandtext;
    $script:errorsEncountered++;
}

# Writes the operation executed and its result to the output CSV
function WriteOperationSummary()
{
    param ($folder, $operation, $result, $commandtext)

    $columns = @(
        (Get-Date).ToString(),
        $folder.Guid,
        $operation,
        (EscapeCsvColumn $result),
        (EscapeCsvColumn $commandtext)
    );

    Add-Content $CsvSummaryFile -Value ("{0},{1},{2},{3},{4}" -f $columns);
}

#Escapes a column value based on RFC 4180 (http://tools.ietf.org/html/rfc4180)
function EscapeCsvColumn()
{
    param ([string]$text)

    if ($text -eq $null)
    {
        return $text;
    }

    $hasSpecial = $false;
    for ($i=0; $i -lt $text.Length; $i++)
    {
        $c = $text[$i];
        if ($c -eq $script:csvEscapeChar -or
            $c -eq $script:csvFieldDelimiter -or
            $script:csvSpecialChars -contains $c)
        {
            $hasSpecial = $true;
            break;
        }
    }

    if (-not $hasSpecial)
    {
        return $text;
    }
    
    $ch = $script:csvEscapeChar.ToString([System.Globalization.CultureInfo]::InvariantCulture);
    return $ch + $text.Replace($ch, $ch + $ch) + $ch;
}

# Writes the current progress
function WriteProgress()
{
    param($statusFormat, $statusProcessed, $statusTotal)
    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity `
        -Status ($statusFormat -f $statusProcessed,$statusTotal) `
        -PercentComplete (100 * ($script:itemsProcessed + $statusProcessed)/$script:totalItems);
}

# Create a tenant PSSession against Exchange Online.
function InitializeExchangeOnlineRemoteSession()
{
    WriteInfoMessage $LocalizedStrings.CreatingRemoteSession;

    $oldWarningPreference = $WarningPreference;
    $oldVerbosePreference = $VerbosePreference;

    try
    {
        $VerbosePreference = $WarningPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue;
        $sessionOption = (New-PSSessionOption -SkipCACheck);
        $script:session = New-PSSession -ConnectionURI:$ConnectionUri `
            -ConfigurationName:Microsoft.Exchange `
            -AllowRedirection `
            -Authentication:"Basic" `
            -SessionOption:$sessionOption `
            -Credential:$Credential `
            -ErrorAction:SilentlyContinue;
        
        if ($script:session -eq $null)
        {
            Write-Error ($LocalizedStrings.FailedToCreateRemoteSession -f $error[0].Exception.Message);
            Exit;
        }
        else
        {
            $result = Import-PSSession -Session $script:session `
                -Prefix "EXO" `
                -AllowClobber;

            if (-not $?)
            {
                Write-Error ($LocalizedStrings.FailedToImportRemoteSession -f $error[0].Exception.Message);
                Remove-PSSession $script:session;
                Exit;
            }
        }
    }
    finally
    {
        $WarningPreference = $oldWarningPreference;
        $VerbosePreference = $oldVerbosePreference;
    }

    WriteInfoMessage $LocalizedStrings.RemoteSessionCreatedSuccessfully;
}

# Invokes New-SyncMailPublicFolder to create a new MEPF object on AD
function NewMailEnabledPublicFolder()
{
    param ($localFolder)

    if ($localFolder.PrimarySmtpAddress.ToString() -eq "")
    {
        $errorMsg = ($LocalizedStrings.FailedToCreateMailPublicFolderEmptyPrimarySmtpAddress -f $localFolder.Guid);
        Write-Error $errorMsg;
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $errorMsg "";
        return;
    }

    # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddressesArray = $localFolder.EmailAddresses.ToStringArray() + ("x500:" + $localFolder.LegacyExchangeDN);
           
    $newParams = @{};
    AddNewOrSetCommonParameters $localFolder $emailAddressesArray $newParams;

    [string]$commandText = (FormatCommand $script:NewSyncMailPublicFolderCommand $newParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        $result = &$script:NewSyncMailPublicFolderCommand @newParams;
        WriteOperationSummary $localFolder $LocalizedStrings.CreateOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsCreated++;
        }
    }
    catch
    {
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $error[0].Exception.Message $commandText;
        Write-Error $_;
    }
}

# Invokes Remove-SyncMailPublicFolder to remove a MEPF from AD
function RemoveMailEnabledPublicFolder()
{
    param ($remoteFolder)    

    $removeParams = @{};
    $removeParams.Add("Identity", $remoteFolder.DistinguishedName);
    $removeParams.Add("Confirm", $false);
    $removeParams.Add("WarningAction", [System.Management.Automation.ActionPreference]::SilentlyContinue);
    $removeParams.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop);

    if ($WhatIf)
    {
        $removeParams.Add("WhatIf", $true);
    }
    
    [string]$commandText = (FormatCommand $script:RemoveSyncMailPublicFolderCommand $removeParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }
    
    try
    {
        &$script:RemoveSyncMailPublicFolderCommand @removeParams;
        WriteOperationSummary $remoteFolder $LocalizedStrings.RemoveOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsDeleted++;
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.RemoveOperationName $_.Exception.Message $commandText;
        Write-Error $_;
    }
}

# Invokes Set-MailPublicFolder to update the properties of an existing MEPF
function UpdateMailEnabledPublicFolder()
{
    param ($localFolder, $remoteFolder)

    $localEmailAddresses = $localFolder.EmailAddresses.ToStringArray();
    $localEmailAddresses += ("x500:" + $localFolder.LegacyExchangeDN); # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddresses = ConsolidateEmailAddresses $localEmailAddresses $remoteFolder.EmailAddresses $remoteFolder.LegacyExchangeDN;

    $setParams = @{};
    $setParams.Add("Identity", $remoteFolder.DistinguishedName);

    if ($script:mailEnabledSystemFolders.Contains($localFolder.Guid))
    {
        $setParams.Add("IgnoreMissingFolderLink", $true);
    }

    AddNewOrSetCommonParameters $localFolder $emailAddresses $setParams;

    [string]$commandText = (FormatCommand $script:SetMailPublicFolderCommand $setParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        &$script:SetMailPublicFolderCommand @setParams;
        WriteOperationSummary $remoteFolder $LocalizedStrings.UpdateOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsUpdated++;
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.UpdateOperationName $_.Exception.Message $commandText;
        Write-Error $_;
    }
}

# Adds the common set of parameters between New and Set cmdlets to the given dictionary
function AddNewOrSetCommonParameters()
{
    param ($localFolder, $emailAddresses, [System.Collections.IDictionary]$parameters)

    $windowsEmailAddress = $localFolder.WindowsEmailAddress.ToString();
    if ($windowsEmailAddress -eq "")
    {
        $windowsEmailAddress = $localFolder.PrimarySmtpAddress.ToString();      
    }

    $parameters.Add("Alias", $localFolder.Alias.Trim());
    $parameters.Add("DisplayName", $localFolder.DisplayName.Trim());
    $parameters.Add("EmailAddresses", $emailAddresses);
    $parameters.Add("ExternalEmailAddress", $localFolder.PrimarySmtpAddress.ToString());
    $parameters.Add("HiddenFromAddressListsEnabled", $localFolder.HiddenFromAddressListsEnabled);
    $parameters.Add("Name", $localFolder.Name.Trim());
    $parameters.Add("OnPremisesObjectId", $localFolder.Guid);
    $parameters.Add("WindowsEmailAddress", $windowsEmailAddress);
    $parameters.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop);

    if ($WhatIf)
    {
        $parameters.Add("WhatIf", $true);
    }
}

# Finds out the cloud-only email addresses and merges those with the values current persisted in the on-premises object
function ConsolidateEmailAddresses()
{
    param($localEmailAddresses, $remoteEmailAddresses, $remoteLegDN)

    # Check if the email address in the existing cloud object is present on-premises; if it is not, then the address was either:
    # 1. Deleted on-premises and must be removed from cloud
    # 2. or it is a cloud-authoritative address and should be kept
    $remoteAuthoritative = @();
    foreach ($remoteAddress in $remoteEmailAddresses)
    {
        if ($remoteAddress.StartsWith("SMTP:", [StringComparison]::InvariantCultureIgnoreCase))
        {
            $found = $false;
            $remoteAddressParts = $remoteAddress.Split($script:proxyAddressSeparators); # e.g. SMTP:alias@domain
            if ($remoteAddressParts.Length -ne 3)
            {
                continue; # Invalid SMTP proxy address (it will be removed)
            }

            foreach ($localAddress in $localEmailAddresses)
            {
                # note that the domain part of email addresses is case insensitive while the alias part is case sensitive
                $localAddressParts = $localAddress.Split($script:proxyAddressSeparators);
                if ($localAddressParts.Length -eq 3 -and
                    $remoteAddressParts[0].Equals($localAddressParts[0], [StringComparison]::InvariantCultureIgnoreCase) -and
                    $remoteAddressParts[1].Equals($localAddressParts[1], [StringComparison]::InvariantCulture) -and
                    $remoteAddressParts[2].Equals($localAddressParts[2], [StringComparison]::InvariantCultureIgnoreCase))
                {
                    $found = $true;
                    break;
                }
            }

            if (-not $found)
            {
                foreach ($domain in $script:authoritativeDomains)
                {
                    if ($remoteAddressParts[2] -eq $domain)
                    {
                        $found = $true;
                        break;
                    }
                }

                if (-not $found)
                {
                    # the address on the remote object is from a cloud authoritative domain and should not be removed
                    $remoteAuthoritative += $remoteAddress;
                }
            }
        }
        elseif ($remoteAddress.StartsWith("X500:", [StringComparison]::InvariantCultureIgnoreCase) -and
            $remoteAddress.Substring(5) -eq $remoteLegDN)
        {
            $remoteAuthoritative += $remoteAddress;
        }
    }

    return $localEmailAddresses + $remoteAuthoritative;
}

# Formats the command and its parameters to be printed on console or to file
function FormatCommand()
{
    param ([string]$command, [System.Collections.IDictionary]$parameters)

    $commandText = New-Object System.Text.StringBuilder;
    [void]$commandText.Append($command);
    foreach ($name in $parameters.Keys)
    {
        [void]$commandText.AppendFormat(" -{0}:",$name);

        $value = $parameters[$name];
        if ($value -isnot [Array])
        {
            [void]$commandText.AppendFormat("`"{0}`"", $value);
        }
        elseif ($value.Length -eq 0)
        {
            [void]$commandText.Append("@()");
        }
        else
        {
            [void]$commandText.Append("@(");
            foreach ($subValue in $value)
            {
                [void]$commandText.AppendFormat("`"{0}`",",$subValue);
            }
            
            [void]$commandText.Remove($commandText.Length - 1, 1);
            [void]$commandText.Append(")");
        }
    }

    return $commandText.ToString();
}

################ DECLARING GLOBAL VARIABLES ################
$script:session = $null;
$script:verbose = $VerbosePreference -eq [System.Management.Automation.ActionPreference]::Continue;

$script:csvSpecialChars = @("`r", "`n");
$script:csvEscapeChar = '"';
$script:csvFieldDelimiter = ',';

$script:ObjectsCreated = $script:ObjectsUpdated = $script:ObjectsDeleted = 0;
$script:NewSyncMailPublicFolderCommand = "New-EXOSyncMailPublicFolder";
$script:SetMailPublicFolderCommand = "Set-EXOMailPublicFolder";
$script:RemoveSyncMailPublicFolderCommand = "Remove-EXOSyncMailPublicFolder";
[char[]]$script:proxyAddressSeparators = ':','@';
$script:errorsEncountered = 0;
$script:authoritativeDomains = $null;
$script:mailEnabledSystemFolders = New-Object 'System.Collections.Generic.HashSet[Guid]'; 
$script:WellKnownSystemFolders = @(
    "\NON_IPM_SUBTREE\EFORMS REGISTRY",
    "\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK",
    "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
    "\NON_IPM_SUBTREE\schema-root",
    "\NON_IPM_SUBTREE\Events Root");

#load hashtable of localized string
Import-LocalizedData -BindingVariable LocalizedStrings -FileName SyncMailPublicFolders.strings.psd1

#minimum supported exchange version to run this script
$minSupportedVersion = 8
################ END OF DECLARATION #################

if (Test-Path $CsvSummaryFile)
{
    Remove-Item $CsvSummaryFile -Confirm:$Confirm -Force;
}

# Write the output CSV headers
$csvFile = New-Item -Path $CsvSummaryFile -ItemType File -Force -ErrorAction:Stop -Value ("#{0},{1},{2},{3},{4}`r`n" -f $LocalizedStrings.TimestampCsvHeader,
    $LocalizedStrings.IdentityCsvHeader,
    $LocalizedStrings.OperationCsvHeader,
    $LocalizedStrings.ResultCsvHeader,
    $LocalizedStrings.CommandCsvHeader);

$localServerVersion = (Get-ExchangeServer $env:COMPUTERNAME -ErrorAction:Stop).AdminDisplayVersion;
# This script can run from Exchange 2007 Management shell and above
if ($localServerVersion.Major -lt $minSupportedVersion)
{
    Write-Error ($LocalizedStrings.LocalServerVersionNotSupported -f $localServerVersion) -ErrorAction:Continue;
    Exit;
}

try
{
    InitializeExchangeOnlineRemoteSession;

    WriteInfoMessage $LocalizedStrings.LocalMailPublicFolderEnumerationStart;

    # During finalization, Public Folders deployment is locked for migration, which means the script cannot invoke
    # Get-PublicFolder as that operation would fail. In that case, the script cannot determine which mail public folder
    # objects are linked to system folders under the NON_IPM_SUBTREE.
    $lockedForMigration = (Get-OrganizationConfig).PublicFoldersLockedForMigration;
    $allSystemFoldersInAD = @();
    if (-not $lockedForMigration)
    {
        # See https://technet.microsoft.com/en-us/library/bb397221(v=exchg.141).aspx#Trees
        # Certain WellKnownFolders in pre-E15 are created with prefix such as OWAScratchPad, StoreEvents.
        # For instance, StoreEvents folders have the following pattern: "\NON_IPM_SUBTREE\StoreEvents{46F83CF7-2A81-42AC-A0C6-68C7AA49FF18}\internal1"
        $storeEventAndOwaScratchPadFolders = @(Get-PublicFolder \NON_IPM_SUBTREE -GetChildren -ResultSize:Unlimited | ?{$_.Name -like "StoreEvents*" -or $_.Name -like "OWAScratchPad*"});
        $allSystemFolderParents = $storeEventAndOwaScratchPadFolders + @($script:WellKnownSystemFolders | Get-PublicFolder -ErrorAction:SilentlyContinue);
        $allSystemFoldersInAD = @($allSystemFolderParents | Get-PublicFolder -Recurse -ResultSize:Unlimited | Get-MailPublicFolder -ErrorAction:SilentlyContinue);

        foreach ($systemFolder in $allSystemFoldersInAD)
        {
            [void]$script:mailEnabledSystemFolders.Add($systemFolder.Guid);
        }
    }
    else
    {
        WriteWarningMessage $LocalizedStrings.UnableToDetectSystemMailPublicFolders;
    }

    if ($script:verbose)
    {
        WriteVerboseMessage ($LocalizedStrings.SystemFoldersSkipped -f $script:mailEnabledSystemFolders.Count);
        $allSystemFoldersInAD | Sort Alias | ft -a | Out-String | Write-Host -ForegroundColor Green -BackgroundColor Black;
    }

    $localFolders = @(Get-MailPublicFolder -ResultSize:Unlimited -IgnoreDefaultScope | Sort Guid);
    WriteInfoMessage ($LocalizedStrings.LocalMailPublicFolderEnumerationCompleted -f $localFolders.Length);

    if ($localFolders.Length -eq 0 -and $Force -eq $false)
    {
        WriteWarningMessage $LocalizedStrings.ForceParameterRequired;
        Exit;
    }

    WriteInfoMessage $LocalizedStrings.RemoteMailPublicFolderEnumerationStart;
    $remoteFolders = @(Get-EXOMailPublicFolder -ResultSize:Unlimited | Sort OnPremisesObjectId);
    WriteInfoMessage ($LocalizedStrings.RemoteMailPublicFolderEnumerationCompleted -f $remoteFolders.Length);

    $missingOnPremisesGuid = @();
    $pendingRemoves = @();
    $pendingUpdates = @{};
    $pendingAdds = @{};

    $localIndex = 0;
    $remoteIndex = 0;
    while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $remoteFolders.Length)
    {
        $local = $localFolders[$localIndex];
        $remote = $remoteFolders[$remoteIndex];

        if ($remote.OnPremisesObjectId -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote;
            $remoteIndex++;
        }
        elseif ($local.Guid.ToString() -eq $remote.OnPremisesObjectId)
        {
            $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local=$local; Remote=$remote }));
            $localIndex++;
            $remoteIndex++;
        }
        elseif ($local.Guid.ToString() -lt $remote.OnPremisesObjectId)
        {
            if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
            {
                $pendingAdds.Add($local.Guid, $local);
            }

            $localIndex++;
        }
        else
        {
            $pendingRemoves += $remote;
            $remoteIndex++;
        }
    }

    # Remaining folders on $localFolders collection must be added to Exchange Online
    while ($localIndex -lt $localFolders.Length)
    {
        $local = $localFolders[$localIndex];

        if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
        {
            $pendingAdds.Add($local.Guid, $local);
        }

        $localIndex++;
    }

    # Remaining folders on $remoteFolders collection must be removed from Exchange Online
    while ($remoteIndex -lt $remoteFolders.Length)
    {
        $remote = $remoteFolders[$remoteIndex];
        if ($remote.OnPremisesObjectId  -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote;
        }
        else
        {
            $pendingRemoves += $remote;
        }
        
        $remoteIndex++;
    }

    if ($missingOnPremisesGuid.Length -gt 0)
    {
        # Process remote objects missing the OnPremisesObjectId using the PrimarySmtpAddress as a key instead.
        $missingOnPremisesGuid = @($missingOnPremisesGuid | Sort PrimarySmtpAddress);
        $localFolders = @($localFolders | Sort PrimarySmtpAddress);

        $localIndex = 0;
        $remoteIndex = 0;
        while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $local = $localFolders[$localIndex];
            $remote = $missingOnPremisesGuid[$remoteIndex];

            if ($local.PrimarySmtpAddress.ToString() -eq $remote.PrimarySmtpAddress.ToString())
            {
                # Make sure the PrimarySmtpAddress has no duplicate on-premises; otherwise, skip updating all objects with duplicate address
                $j = $localIndex + 1;
                while ($j -lt $localFolders.Length)
                {
                    $next = $localFolders[$j];
                    if ($local.PrimarySmtpAddress.ToString() -ne $next.PrimarySmtpAddress.ToString())
                    {
                        break;
                    }

                    WriteErrorSummary $next $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByAnotherFolder -f $local.PrimarySmtpAddress,$local.Guid) "";

                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($next.Guid);
                    $pendingUpdates.Remove($next.Guid);
                    $j++;
                }

                $duplicatesFound = $j - $localIndex - 1;
                if ($duplicatesFound -gt 0)
                {
                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($local.Guid);
                    $pendingUpdates.Remove($local.Guid);
                    $localIndex += $duplicatesFound + 1;

                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByOtherFolders -f $local.PrimarySmtpAddress,$duplicatesFound) "";
                    WriteWarningMessage ($LocalizedStrings.SkippingFoldersWithDuplicateAddress -f ($duplicatesFound + 1),$local.PrimarySmtpAddress);
                }
                elseif ($pendingUpdates.Contains($local.Guid))
                {
                    # If we get here, it means two different remote objects match the same local object (one by OnPremisesObjectId and another by PrimarySmtpAddress).
                    # Since that is an ambiguous resolution, let's skip updating the remote objects.
                    $ambiguousRemoteObj = $pendingUpdates[$local.Guid].Remote;
                    $pendingUpdates.Remove($local.Guid);

                    $errorMessage = ($LocalizedStrings.AmbiguousLocalMailPublicFolderResolution -f $local.Guid,$ambiguousRemoteObj.Guid,$remote.Guid);
                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName $errorMessage "";
                    WriteWarningMessage $errorMessage;
                }
                else
                {
                    # Since there was no match originally using OnPremisesObjectId, the local object was treated as an add to Exchange Online.
                    # In this way, since we now found a remote object (by PrimarySmtpAddress) to update, we must first remove the local object from the add list.
                    $pendingAdds.Remove($local.Guid);
                    $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local=$local; Remote=$remote }));
                }

                $localIndex++;
                $remoteIndex++;
            }
            elseif ($local.PrimarySmtpAddress.ToString() -gt $remote.PrimarySmtpAddress.ToString())
            {
                # There are no local objects using the remote object's PrimarySmtpAddress
                $pendingRemoves += $remote;
                $remoteIndex++;
            }
            else
            {
                $localIndex++;
            }
        }

        # All objects remaining on the $missingOnPremisesGuid list no longer exist on-premises
        while ($remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $pendingRemoves += $missingOnPremisesGuid[$remoteIndex];
            $remoteIndex++;
        }
    }

    $script:totalItems = $pendingRemoves.Length + $pendingUpdates.Count + $pendingAdds.Count;

    # At this point, we know all changes that need to be synced to Exchange Online. Let's prompt the admin for confirmation before proceeding.
    if ($Confirm -eq $true -and $script:totalItems -gt 0)
    {
        $title = $LocalizedStrings.ConfirmationTitle;
        $message = ($LocalizedStrings.ConfirmationQuestion -f $pendingAdds.Count,$pendingUpdates.Count,$pendingRemoves.Length);
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationYesOption, `
            $LocalizedStrings.ConfirmationYesOptionHelp;

        $no = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationNoOption, `
            $LocalizedStrings.ConfirmationNoOptionHelp;

        [System.Management.Automation.Host.ChoiceDescription[]]$options = $no,$yes;
        $confirmation = $host.ui.PromptForChoice($title, $message, $options, 0);
        if ($confirmation -eq 0)
        {
            Exit;
        }
    }

    # Find out the authoritative AcceptedDomains on-premises so that we don't accidently remove cloud-only email addresses during updates
    $script:authoritativeDomains = @(Get-AcceptedDomain | ?{$_.DomainType -eq "Authoritative" } | foreach {$_.DomainName.ToString()});
    
    # Finally, let's perfom the actual operations against Exchange Online
    $script:itemsProcessed = 0;
    for ($i = 0; $i -lt $pendingRemoves.Length; $i++)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusRemoving $i $pendingRemoves.Length;
        RemoveMailEnabledPublicFolder $pendingRemoves[$i];
    }

    $script:itemsProcessed += $pendingRemoves.Length;
    $updatesProcessed = 0;
    foreach ($folderPair in $pendingUpdates.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusUpdating $updatesProcessed $pendingUpdates.Count;
        UpdateMailEnabledPublicFolder $folderPair.Local $folderPair.Remote;
        $updatesProcessed++;
    }

    $script:itemsProcessed += $pendingUpdates.Count;
    $addsProcessed = 0;
    foreach ($localFolder in $pendingAdds.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusCreating $addsProcessed $pendingAdds.Count;
        NewMailEnabledPublicFolder $localFolder;
        $addsProcessed++;
    }

    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity -Status ($LocalizedStrings.ProgressBarStatusCreating -f $pendingAdds.Count,$pendingAdds.Count) -Completed;
    WriteInfoMessage ($LocalizedStrings.SyncMailPublicFolderObjectsComplete -f $script:ObjectsCreated,$script:ObjectsUpdated,$script:ObjectsDeleted);

    if ($script:errorsEncountered -gt 0)
    {
        WriteWarningMessage ($LocalizedStrings.ErrorsFoundDuringImport -f $script:errorsEncountered,(Get-Item $CsvSummaryFile).FullName);
    }
}
finally
{
    if ($script:session -ne $null)
    {
        Remove-PSSession $script:session;
    }
}
# SIG # Begin signature block
# MIIdtQYJKoZIhvcNAQcCoIIdpjCCHaICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHZbIYGfoB0qYph1U53JyqTk9
# B9agghhlMIIEwzCCA6ugAwIBAgITMwAAAMp9MhZ8fv0FAwAAAAAAyjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwOTA3MTc1ODU1
# WhcNMTgwOTA3MTc1ODU1WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjcyOEQtQzQ1Ri1GOUVCMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAj3CeDl2ll7S4
# 96ityzOt4bkPI1FucwjpTvklJZLOYljFyIGs/LLi6HyH+Czg8Xd/oDQYFzmJTWac
# A0flGdvk8Yj5OLMEH4yPFFgQsZA5Wfnz/Cg5WYR2gmsFRUFELCyCbO58DvzOQQt1
# k/tsTJ5Ns5DfgCb5e31m95yiI44v23FVpKnTY9CUJbIr8j28O3biAhrvrVxI57GZ
# nzkUM8GPQ03o0NGCY1UEpe7UjY22XL2Uq816r0jnKtErcNqIgglXIurJF9QFJrvw
# uvMbRjeTBTCt5o12D4b7a7oFmQEDgg+koAY5TX+ZcLVksdgPNwbidprgEfPykXiG
# ATSQlFCEXwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFGb30hxaE8ox6QInbJZnmt6n
# G7LKMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAGyg/1zQebvX564G4LsdYjFr9ptnqO4KaD0lnYBECEjMqdBM
# 4t+rNhN38qGgERoc+ns5QEGrrtcIW30dvMvtGaeQww5sFcAonUCOs3OHR05QII6R
# XYbxtAMyniTUPwacJiiCSeA06tLg1bebsrIY569mRQHSOgqzaO52EzJlOtdLrGDk
# Ot1/eu8E2zN9/xetZm16wLJVCJMb3MKosVFjFZ7OlClFTPk6rGyN9jfbKKDsDtNr
# jfAiZGVhxrEqMiYkj4S4OyvJ2uhw/ap7dbotTCfZu1yO57SU8rE06K6j8zWB5L9u
# DmtgcqXg3ckGvdmWVWBrcWgnmqNMYgX50XSzffQwggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhEwggP5
# oAMCAQICEzMAAACOh5GkVxpfyj4AAAAAAI4wDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNjExMTcyMjA5MjFaFw0xODAy
# MTcyMjA5MjFaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDQh9RCK36d2cZ61KLD4xWS
# 0lOdlRfJUjb6VL+rEK/pyefMJlPDwnO/bdYA5QDc6WpnNDD2Fhe0AaWVfIu5pCzm
# izt59iMMeY/zUt9AARzCxgOd61nPc+nYcTmb8M4lWS3SyVsK737WMg5ddBIE7J4E
# U6ZrAmf4TVmLd+ArIeDvwKRFEs8DewPGOcPUItxVXHdC/5yy5VVnaLotdmp/ZlNH
# 1UcKzDjejXuXGX2C0Cb4pY7lofBeZBDk+esnxvLgCNAN8mfA2PIv+4naFfmuDz4A
# lwfRCz5w1HercnhBmAe4F8yisV/svfNQZ6PXlPDSi1WPU6aVk+ayZs/JN2jkY8fP
# AgMBAAGjggGAMIIBfDAfBgNVHSUEGDAWBgorBgEEAYI3TAgBBggrBgEFBQcDAzAd
# BgNVHQ4EFgQUq8jW7bIV0qqO8cztbDj3RUrQirswUgYDVR0RBEswSaRHMEUxDTAL
# BgNVBAsTBE1PUFIxNDAyBgNVBAUTKzIzMDAxMitiMDUwYzZlNy03NjQxLTQ0MWYt
# YmM0YS00MzQ4MWU0MTVkMDgwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1
# ApUwVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jcmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEF
# BQcBAQRVMFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9w
# a2lvcHMvY2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNV
# HRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBEiQKsaVPzxLa71IxgU+fKbKhJ
# aWa+pZpBmTrYndJXAlFq+r+bltumJn0JVujc7SV1eqVHUqgeSxZT8+4PmsMElSnB
# goSkVjH8oIqRlbW/Ws6pAR9kRqHmyvHXdHu/kghRXnwzAl5RO5vl2C5fAkwJnBpD
# 2nHt5Nnnotp0LBet5Qy1GPVUCdS+HHPNIHuk+sjb2Ns6rvqQxaO9lWWuRi1XKVjW
# kvBs2mPxjzOifjh2Xt3zNe2smjtigdBOGXxIfLALjzjMLbzVOWWplcED4pLJuavS
# Vwqq3FILLlYno+KYl1eOvKlZbiSSjoLiCXOC2TWDzJ9/0QSOiLjimoNYsNSa5jH6
# lEeOfabiTnnz2NNqMxZQcPFCu5gJ6f/MlVVbCL+SUqgIxPHo8f9A1/maNp39upCF
# 0lU+UK1GH+8lDLieOkgEY+94mKJdAw0C2Nwgq+ZWtd7vFmbD11WCHk+CeMmeVBoQ
# YLcXq0ATka6wGcGaM53uMnLNZcxPRpgtD1FgHnz7/tvoB3kH96EzOP4JmtuPe7Y6
# vYWGuMy8fQEwt3sdqV0bvcxNF/duRzPVQN9qyi5RuLW5z8ME0zvl4+kQjOunut6k
# LjNqKS8USuoewSI4NQWF78IEAA1rwdiWFEgVr35SsLhgxFK1SoK3hSoASSomgyda
# Qd691WZJvAuceHAJvDCCB3owggVioAMCAQICCmEOkNIAAAAAAAMwDQYJKoZIhvcN
# AQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAw
# BgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEx
# MB4XDTExMDcwODIwNTkwOVoXDTI2MDcwODIxMDkwOVowfjELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUg
# U2lnbmluZyBQQ0EgMjAxMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AKvw+nIQHC6t2G6qghBNNLrytlghn0IbKmvpWlCquAY4GgRJun/DDB7dN2vGEtgL
# 8DjCmQawyDnVARQxQtOJDXlkh36UYCRsr55JnOloXtLfm1OyCizDr9mpK656Ca/X
# llnKYBoF6WZ26DJSJhIv56sIUM+zRLdd2MQuA3WraPPLbfM6XKEW9Ea64DhkrG5k
# NXimoGMPLdNAk/jj3gcN1Vx5pUkp5w2+oBN3vpQ97/vjK1oQH01WKKJ6cuASOrdJ
# Xtjt7UORg9l7snuGG9k+sYxd6IlPhBryoS9Z5JA7La4zWMW3Pv4y07MDPbGyr5I4
# ftKdgCz1TlaRITUlwzluZH9TupwPrRkjhMv0ugOGjfdf8NBSv4yUh7zAIXQlXxgo
# tswnKDglmDlKNs98sZKuHCOnqWbsYR9q4ShJnV+I4iVd0yFLPlLEtVc/JAPw0Xpb
# L9Uj43BdD1FGd7P4AOG8rAKCX9vAFbO9G9RVS+c5oQ/pI0m8GLhEfEXkwcNyeuBy
# 5yTfv0aZxe/CHFfbg43sTUkwp6uO3+xbn6/83bBm4sGXgXvt1u1L50kppxMopqd9
# Z4DmimJ4X7IvhNdXnFy/dygo8e1twyiPLI9AN0/B4YVEicQJTMXUpUMvdJX3bvh4
# IFgsE11glZo+TzOE2rCIF96eTvSWsLxGoGyY0uDWiIwLAgMBAAGjggHtMIIB6TAQ
# BgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUSG5k5VAF04KqFzc3IrVtqMp1ApUw
# GQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB
# /wQFMAMBAf8wHwYDVR0jBBgwFoAUci06AjGQQ7kUBU7h6qfHMdEjiTQwWgYDVR0f
# BFMwUTBPoE2gS4ZJaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJv
# ZHVjdHMvTWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNybDBeBggrBgEFBQcB
# AQRSMFAwTgYIKwYBBQUHMAKGQmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kv
# Y2VydHMvTWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNydDCBnwYDVR0gBIGX
# MIGUMIGRBgkrBgEEAYI3LgMwgYMwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvZG9jcy9wcmltYXJ5Y3BzLmh0bTBABggrBgEFBQcC
# AjA0HjIgHQBMAGUAZwBhAGwAXwBwAG8AbABpAGMAeQBfAHMAdABhAHQAZQBtAGUA
# bgB0AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAZ/KGpZjgVHkaLtPYdGcimwuWEeFj
# kplCln3SeQyQwWVfLiw++MNy0W2D/r4/6ArKO79HqaPzadtjvyI1pZddZYSQfYtG
# UFXYDJJ80hpLHPM8QotS0LD9a+M+By4pm+Y9G6XUtR13lDni6WTJRD14eiPzE32m
# kHSDjfTLJgJGKsKKELukqQUMm+1o+mgulaAqPyprWEljHwlpblqYluSD9MCP80Yr
# 3vw70L01724lruWvJ+3Q3fMOr5kol5hNDj0L8giJ1h/DMhji8MUtzluetEk5CsYK
# wsatruWy2dsViFFFWDgycScaf7H0J/jeLDogaZiyWYlobm+nt3TDQAUGpgEqKD6C
# PxNNZgvAs0314Y9/HG8VfUWnduVAKmWjw11SYobDHWM2l4bf2vP48hahmifhzaWX
# 0O5dY0HjWwechz4GdwbRBrF1HxS+YWG18NzGGwS+30HHDiju3mUv7Jf2oVyW2ADW
# oUa9WfOXpQlLSBCZgB/QACnFsZulP0V3HjXG0qKin3p6IvpIlR+r+0cjgPWe+L9r
# t0uX4ut1eBrs6jeZeRhL/9azI2h15q/6/IvrC4DqaTuv/DDtBEyO3991bWORPdGd
# Vk5Pv4BXIqF4ETIheu9BCrE/+6jMpF3BoYibV3FWTkhFwELJm3ZbCoBIa/15n8G9
# bW1qyVJzEw16UM0xggS6MIIEtgIBATCBlTB+MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5n
# IFBDQSAyMDExAhMzAAAAjoeRpFcaX8o+AAAAAACOMAkGBSsOAwIaBQCggc4wGQYJ
# KoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQB
# gjcCARUwIwYJKoZIhvcNAQkEMRYEFB0xzxKxeuFtJ8PqZCbFgyIPaG59MG4GCisG
# AQQBgjcCAQwxYDBeoDaANABTAHkAbgBjAC0ATQBhAGkAbABQAHUAYgBsAGkAYwBG
# AG8AbABkAGUAcgBzAC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQA1yP50BCfGe8eBwAYBDYwDKTpJ
# 7eRpliUVZ1sRQjAGynhkNgLH9sOHdC/EP4vyn4SOAyExE3n85dITosDoLouDV97F
# ld4lGAmCi+Jjvrrju+8pErh11qPgIT6TPah7K0/o4Fb/WAoq9bTx6sh55FwiIWgA
# yPcC/X76vN+RPS9Nlha/VkF/12QxHgkCdbhL/Z9mfVRGPdSn9FhzwAqs/3ps4jLP
# c8ZnxfgyH+8fv1MwbhsBIIPeKAlkR+k4hs+/KJq+IuXDFKktMegiBbRdNoUnAGkg
# rLPuu8ztTkSPO4p6zqxcv3XrVGYExT6dBQ0qgJqCWY2hYHsHiWcZ+4+8CZwsoYIC
# KDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFBDQQITMwAAAMp9MhZ8fv0FAwAAAAAAyjAJBgUrDgMCGgUAoF0wGAYJKoZI
# hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTcwMzMxMTQ0OTMx
# WjAjBgkqhkiG9w0BCQQxFgQURVbwhQhKGtMojwZzjReg70JKkqUwDQYJKoZIhvcN
# AQEFBQAEggEAjQ2QfkWmHZClaLWF3dGobLmVwxpJAm1pvlHm1ezMBAWW2ng5a3w0
# Gjce8vFZGtqjDc1azsNgMhPo7sB/xTCljM4N1ZQO486+anY8Th8H1fYy4syUWm10
# 60Hyob4lpSaHQZS1P/712qR76iX1fofRizXXy1xe9rGt/BHxs/DbXBMYLVr170NQ
# X5m+QJqnz2cfr1ua9NPTGk8O8GAtnzRDih+19TJEOB6vBu3ubIpN0mihWyPNhApL
# 7d9s1CTBvtVG4zJFjC6zwKEUP9HBgkVtEPu2pvIXDNbgB13Alf7ejCf1PIiPCogs
# AK0wBocoA6BDZwAa1WyxUQ8sPJN7AxyVAQ==
# SIG # End signature block
