# .SYNOPSIS
# Sync-MailPublicFoldersCloudToOnprem.ps1
#    This script imports the new mail public folders as sync mail public folders from Exchange Online to on-premise.
#	 And also synchronizes the properites of existing mail-enabled public folders from Exchange Online to on-premises (thereby overriding the mail public folder properties in on-premise).
#
# Example input to the script:
#
# Sync-MailPublicFoldersCloudToOnprem.ps1 -ConnectionUri <cloud url> -Credential <credential> -CsvSummaryFile <path for the summary file>
#
# The above example imports new mail public folders objects from Exchange Online as sync mail public folders to on-premise.
#
# .DESCRIPTION
#
# Copyright (c) 2016 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID",

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string] $CsvSummaryFile,

    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [PSCredential] $Credential
    )

# Writes a dated information message to console
function WriteInfoMessage()
{
    param ($message)
    Write-Host "[$($(Get-Date).ToString())]" $message;
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

## Formats the command and its parameters to be printed on console or to file
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

## Retrieve mail enabled public folders from EXO
function GetMailPublicFolders(
    [System.Management.Automation.Runspaces.PSSession] $session)
{    
    $mailPublicFolders = ExecuteCommand "Get-MailPublicFolder -ResultSize:Unlimited -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue" $session;
    
    # Return the results
    if ($mailPublicFolders -eq $null -or ([array]($mailPublicFolders)).Count -eq 0)
    {
        return $null;
    }

    return $mailPublicFolders;
}

## Sync mail public folders from cloud to on-premise.
function SyncMailPublicFolders(
    [object[]] $mailPublicFolders)
{
    $validExternalEmailAddresses = @();

    if ($mailPublicFolders -ne $null)
    {        
        foreach ($mailPublicFolder in $mailPublicFolders)
        {
            # extracting properties
            $alias = $mailPublicFolder.Alias.Trim();
            $externalEmailAddress = $mailPublicFolder.PrimarySmtpAddress.ToString();
            $entryId = $mailPublicFolder.EntryId.ToString();
            $name = $mailPublicFolder.Name.Trim();
            $displayName = $mailPublicFolder.DisplayName.Trim()
            $hiddenFromAddressListsEnabled = $mailPublicFolder.HiddenFromAddressListsEnabled;

            $windowsEmailAddress = $mailPublicFolder.WindowsEmailAddress.ToString();
            if ($windowsEmailAddress -eq "")
            {
                $windowsEmailAddress = $externalEmailAddress;
            }

            # extracting all the EmailAddress
            $emailAddress = @();
            foreach($address in $mailPublicFolder.EmailAddresses)
            {
                $emailAddress += $address.ToString();
            }

            # preserve the ability to reply via Outlook's nickname cache post-migration
            $emailAddress +=  ("X500:" + $mailPublicFolder.LegacyExchangeDN);

            WriteInfoMessage ($LocalizedStrings.SyncingMailPublicFolder -f $alias);

            $syncPublicFolder = ExecuteCommand "Get-MailPublicFolder -Identity '$externalEmailAddress' -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue";

            if ($syncPublicFolder -eq $null)
            {
                WriteInfoMessage ($LocalizedStrings.CreatingSyncMailPublicFolder -f $alias);
                try
                {
                    $newParams = @{};
                    $newParams.Add("Name", $name);
                    $newParams.Add("ExternalEmailAddress", $externalEmailAddress);
                    $newParams.Add("Alias", $alias);
                    $newParams.Add("EntryId", $entryId);
                    $newParams.Add("WindowsEmailAddress", $windowsEmailAddress);
                    $newParams.Add("WarningAction", "SilentlyContinue");
                    $newParams.Add("ErrorAction", "Stop");

                    [string]$createSyncPublicFolder = (FormatCommand $script:NewSyncMailPublicFolderCommand $newParams);
                    
                    # Creating new sync mail public folder
                    $newMailPublicFolder = &$script:NewSyncMailPublicFolderCommand @newParams;

                    WriteOperationSummary $mailPublicFolder $LocalizedStrings.CreateOperationName $LocalizedStrings.CsvSuccessResult $createSyncPublicFolder;

                    $setParams = @{};
                    $setParams.Add("Identity", $name);
                    $setParams.Add("EmailAddresses", $emailAddress);
                    $setParams.Add("DisplayName", $displayName);
                    $setParams.Add("HiddenFromAddressListsEnabled", $hiddenFromAddressListsEnabled);
                    $setParams.Add("WarningAction", "SilentlyContinue");
                    $setParams.Add("ErrorAction", "Stop");

                    [string]$setOtherProperties = (FormatCommand $script:SetMailPublicFolderCommand $setParams);

                    # Setting other properties to the newly created sync mail public folder
                    &$script:SetMailPublicFolderCommand @setParams;
                    
                    WriteOperationSummary $mailPublicFolder $LocalizedStrings.SetOperationName $LocalizedStrings.CsvSuccessResult $setOtherProperties;

                    $validExternalEmailAddresses += $newMailPublicFolder.ExternalEmailAddress;
                    $script:CreatedPublicFoldersCount++;
                }

                catch
                { 
                    WriteErrorSummary $mailPublicFolder $LocalizedStrings.CreateOperationName $_.Exception.Message "";
                    Write-Error $_;
                }
            }

            else
            {
                WriteInfoMessage ($LocalizedStrings.UpdatingSyncMailPublicFolder -f $syncPublicFolder);
                try
                {
                    $updateParams = @{};
                    $updateParams.Add("Identity", $syncPublicFolder);
                    $updateParams.Add("EmailAddresses", $emailAddress);
                    $updateParams.Add("HiddenFromAddressListsEnabled", $hiddenFromAddressListsEnabled);
                    $updateParams.Add("DisplayName", $displayName);
                    $updateParams.Add("Name", $name);
                    $updateParams.Add("ExternalEmailAddress", $externalEmailAddress);
                    $updateParams.Add("Alias", $alias);
                    $updateParams.Add("WindowsEmailAddress", $windowsEmailAddress);
                    $updateParams.Add("WarningAction", "SilentlyContinue");
                    $updateParams.Add("ErrorAction", "Stop");

                    [string]$updateProperties = (FormatCommand $script:SetMailPublicFolderCommand $updateParams);

                    # Setting properties to the existing sync mail public folder
                    &$script:SetMailPublicFolderCommand @updateParams;

                    WriteOperationSummary $mailPublicFolder $LocalizedStrings.UpdateOperationName $LocalizedStrings.CsvSuccessResult $updateProperties;

                    $validExternalEmailAddresses += $syncPublicFolder.ExternalEmailAddress;
                    $script:UpdatedPublicFoldersCount++;
               }

               catch
               {
                    WriteErrorSummary $mailPublicFolder $LocalizedStrings.UpdateOperationName $_.Exception.Message $updateProperties;
                    Write-Error $_;
               }

            }

            WriteInfoMessage ($LocalizedStrings.DoneSyncingMailPublicFolder -f $alias);
            Write-Host "";
        }
    }

    else
    {
        WriteInfoMessage ($LocalizedStrings.NoMailPublicFoldersToSync);
        Write-Host "";
    }

    WriteInfoMessage ($LocalizedStrings.DeleteSyncMailPublicFolderTitle);

    $localMailPublicFolders = ExecuteCommand "Get-MailPublicFolder -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue";

    foreach($syncPublicFolder in $localMailPublicFolders)
    {
        if (-not $validExternalEmailAddresses.Contains($syncPublicFolder.ExternalEmailAddress))
        {
            WriteInfoMessage ($LocalizedStrings.DeletingSyncMailPublicFolder -f $syncPublicFolder);
            try
            {
                $deleteParams = @{};
                $deleteParams.Add("Identity", $syncPublicFolder);
                $deleteParams.Add("Confirm", $false);

                [string]$disableMailPublicFolder = (FormatCommand $script:DeletePublicFolderCommand $deleteParams);

                # Deleting sync mail public folder
                &$script:DeletePublicFolderCommand @deleteParams;

                WriteOperationSummary $syncPublicFolder $LocalizedStrings.DeleteOperationName $LocalizedStrings.CsvSuccessResult $disableMailPublicFolder;
                $script:RemovedPublicFoldersCount++;
            }
            catch
            {
                WriteErrorSummary $syncPublicFolder $LocalizedStrings.DeleteOperationName $_.Exception.Message $disableMailPublicFolder;
                Write-Error $_;
            }
        }
    }
}

################ DECLARING GLOBAL VARIABLES ################
$script:session = $null;

$script:csvSpecialChars = @("`r", "`n");
$script:csvEscapeChar = '"';
$script:csvFieldDelimiter = ',';
$script:NewSyncMailPublicFolderCommand = "New-SyncMailPublicFolder";
$script:SetMailPublicFolderCommand = "Set-MailPublicFolder";
$script:DeletePublicFolderCommand = "Disable-MailPublicFolder";
$script:CreatedPublicFoldersCount = 0;
$script:UpdatedPublicFoldersCount = 0;
$script:RemovedPublicFoldersCount = 0;

#load hashtable of localized string
Import-LocalizedData -BindingVariable LocalizedStrings -FileName SyncMailPublicFoldersCloudToOnprem.strings.psd1

#minimum supported exchange version to run this script
$minSupportedVersion = 15
################ END OF DECLARATION #########################

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
# This script can run from Exchange 2013 Management shell and above
if ($localServerVersion.Major -lt $minSupportedVersion)
{
    Write-Error ($LocalizedStrings.LocalServerVersionNotSupported -f $localServerVersion) -ErrorAction:Continue;
    Exit;
}

# Create a PSSession
WriteInfoMessage ($LocalizedStrings.CreatingRemoteSession);

$session = GetTenantSession -uri:$ConnectionUri -cred:$Credential;
if ($session -eq $null)
{
    WriteInfoMessage ($LocalizedStrings.FailedToCreateRemoteSession -f $_.Exception.Message);
    exit;
}
WriteInfoMessage ($LocalizedStrings.RemoteSessionCreatedSuccessfully);

# Get mail enabled public folders in cloud
WriteInfoMessage ($LocalizedStrings.StartedImportingMailPublicFolders);
Write-Host "";

$mailPublicFoldersEXO = GetMailPublicFolders $session;

# Create sync mail public folders in on-premise
SyncMailPublicFolders $mailPublicFoldersEXO;

Write-Host "";
WriteInfoMessage ($LocalizedStrings.CompletedImportingMailPublicFolders);
WriteInfoMessage ($LocalizedStrings.CompletedStatsCount -f  $script:CreatedPublicFoldersCount, $script:UpdatedPublicFoldersCount, $script:RemovedPublicFoldersCount);

# Terminate the PSSession
Remove-PSSession $session;