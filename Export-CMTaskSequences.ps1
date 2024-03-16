<#
.SYNOPSIS
    This PowerShell script automates Task Sequence exporting from Configuration Manager, including package details, references, and sizes; additionally, it supports Git commit and push for version control.

.DESCRIPTION
    This PowerShell script automates the process of exporting Task Sequences from Configuration Manager.
    It takes input parameters such as the repository path, the output path, the use of Git to commit and push changes, the custom commit message, and the excluded paths. 
    The script retrieves task sequences and checks for exclusions based on the TS paths. 
    Then it fetches task sequence details like package ID, boot image, operating system image, referenced packages, their size, and applications. 
    It then writes the gathered information into a file, one per task sequence. 
    If Git is enabled, it automatically commits and pushes changes to the remote repository using the specified branch and push settings. 
    In case of errors or if no task sequences have been modified since a specified date, the script writes error messages or status updates to an error log file.

.NOTES
    Some functions are copied and slightly modified from the excellent project: https://github.com/paulwetter/DocumentConfigMgrCB.
    Visit the link to find out more.

.LINK
    https://roenlond.github.io/

.PARAMETER SMSProvider
   This is a mandatory parameter. You need to provide the FQDN of a SMS provider.

.PARAMETER RepoPath
   This is an optional path to your local repository. By default, it is the directory from which the script is being run. 
   NOTE: 'FileSystem::' will always be prepended to this path since we want to use the path from the ConfigMgr Drive.

.PARAMETER OutputPath
   This is an optional string parameter to the output directory of your exported task sequences. 
   By default, it is a folder joined with RepoPath as a ChildPath (i.e. in the root of RepoPath) with the name CMTaskSequences.
   NOTE: 'FileSystem::' will always be prepended to this path since we want to use the path from the ConfigMgr Drive.

.PARAMETER ExclusionPath
   This is a parameter that can either be a string or a comma-separated array of strings in the format as shown under Task Sequences in ConfigMgr. 
   For example: "Client/Test/Windows 11". Any task sequence found in the specific excluded path will not be exported.
   This parameter accepts wildcards, so if you supply "*Test*", any task sequence found in a path where the word Test exists would be excluded from export.
    
.PARAMETER FullExport
   This is a switch parameter that determines whether to do a full export of all task sequences regardless of when the script was run last time.
   It is mutually exclusive with SpecificTaskSequence

.PARAMETER SpecificTaskSequence
   This is a string parameter that can be specified to export a specific task sequence instead of basing the export on dates.
   If this is used, no "last run" date will be set and any Date parameters will be ignored. It needs to be an exact name and does not support wildcards.
   It is mutually exclusive with FullExport
   
.PARAMETER AfterDate
   This is a DateTime parameter. 
   By default, the script will export all modified task sequences since last run (or all TS:es if it's the first run).
   With this parameter you can manually override this functionality to export all task sequences modified after your specified date.
   If you supply the parameter without a date, the current date -7 days will be the default.

   The script is tested with the Swedish 'MM/dd/yyyy hh:mm:ss' format, your mileage may vary.

.PARAMETER UseGit
   This is a switch parameter that determines whether to push changes to Git. By default, it is set to false.

.PARAMETER CustomCommitMessage
    This is a string parameter that sets the commit message when using Git. It only has an effect if useGit is set to true.
    If you do not supply a custom commit message, the default will be a newline-separated list of all exported task sequences this run.

.PARAMETER GitBranch
   This is a string parameter that sets the git branch to checkout when using Git. It only has an effect if useGit is set to true. 
   By default, it is 'main'.

.EXAMPLE
    Default behaviour without pushing to git:
    .\Export-CMTaskSequences.ps1 -SMSProvider "SMSProvider.domain.com" 

    Default behaviour and pushing to git:
    .\Export-CMTaskSequences.ps1 -SMSProvider "SMSProvider.domain.com" -UseGit

    Exporting a single task sequence only:
    .\Export-CMTaskSequences.ps1 -SMSProvider "SMSProvider.domain.com" -SpecificTaskSequence "Windows 10 (23H2)"

    Exporting everything to a specific path:
    .\Export-CMTaskSequences.ps1 -SMSProvider "SMSProvider.domain.com" -FullExport -OutputPath "C:\ConfigMgrTools\CMTaskSequences"

    Everything at once:
    .\Export-CMTaskSequences.ps1 -SMSProvider "SMSProvider.domain.com" -RepoPath "\\networkShare\Scripts\Export-CMTaskSequences" -ExclusionPath "*Test*" -useGit $true -AfterDate "2024-02-11 16:00:00" -customCommitMessage "My Commit Message" -GitBranch "master"

    And if you'd like to splat the parameters:
    $params = @{
        SMSProvider = 'SMSProvider.domain.com'
        SpecificTaskSequence = 'Windows 10 (23H2)'
        UseGit = $true
    }

    .\Export-CMTaskSequences.ps1 @params

.AUTHOR 
    Patrik Rönnlund 2024-01-11
    https://roenlond.github.io/
    
.MODIFIED
    2024-01-11 - Version 0.0.1 - Initial version
    2024-01-12 - Version 0.0.2 - Modified to export Task sequences to the same structure as ConfigMgr
    2024-03-01 - Version 0.0.3 - Added calculated ScriptHash on embedded Run Powershell scripts
                                 General cleanup and small fixes
    2024-03-11 - Version 0.1.0 - Parameterised for release, helptext added.
    2024-03-12 - Version 0.1.1 - Added git parameters, exclusion path, fullexport, custom output path and specific task sequence parameters 
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="The full FQDN of a SMS Provider")]
    [ValidateNotNullOrEmpty()]
    [string]$SMSProvider,
    
    [ValidateNotNullOrEmpty()]
    [string]$RepoPath = (Join-Path "FileSystem::" -ChildPath $PSScriptRoot),

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-path $RepoPath -ChildPath "CMTaskSequences"),

    [ValidateNotNullOrEmpty()]
    [string[]]$ExclusionPaths,

    [switch]$FullExport = $false,

    [ValidateNotNullOrEmpty()]
    [string]$SpecificTaskSequence,
    
    [ValidateNotNullOrEmpty()]
    [DateTime]$AfterDate = (Get-Date).AddDays(-7),

    [Parameter(ParameterSetName='Git')]
    [switch]$UseGit = $false,

    [Parameter(ParameterSetName='Git')]
    [string]$CustomCommitMessage,

    [Parameter(ParameterSetName='Git')]
    [ValidateScript({
        $branchName = $_
        $remoteBranches = git branch -r | ForEach-Object { $_.trim().replace('origin/', '') }
        if ($branchName -in $remoteBranches) { $true }
        else { throw "`n Branch <$branchName> does not exist in the repo." }
    })]
    [string]$GitBranch = 'main'
)

# Check mutual exclusivity for full or specific export
if ($FullExport -and $SpecificTaskSequence) {
    throw "Parameters FullExport and SpecificTaskSequence cannot both be specified."
}

# Configure script paths
$LastUpdatedTSPath = Join-Path $RepoPath -ChildPath "LastUpdatedTS.txt"
$LastRunPath = Join-Path $RepoPath -ChildPath "LastRun.txt"
$LastErrorPath = Join-Path $RepoPath -ChildPath "LastError.txt"

# Set the base Date. This is modified later if needed.
$Date = [DateTime]::MinValue

# Load custom functions
try {
    $CustomFunctionsScript = Join-Path $RepoPath -ChildPath "CMReportingFunctions.ps1"
    . $CustomfunctionsScript
} catch {
    throw "The custom functions script could not be loaded. Make sure the RepoPath is correct and files exist."
}

# Get sitecode and load CM cmdlets
$SiteCode = Get-SiteCode

# Import the ConfigurationManager.psd1 module 
if($null -eq (Get-Module ConfigurationManager)) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" 
}

# Connect to the site's drive if it is not already present
if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SMSProvider 
}
$CMPSSuppressFastNotUsedCheck = $true

Set-Location "$($SiteCode):"

# Clear any previous errors
$error.clear()

##############################
# Task sequence script start #
##############################
# Handle date parameters
if ($FullExport -eq $False -and $SpecificTaskSequence -eq $false) {
    if ([string]::IsNullOrEmpty($AfterDate)) {
        try {
            if ([string]::IsNullOrWhiteSpace($SpecificTaskSequence)) {
                $Date = Get-content $LastRunPath -ErrorAction stop
                Write-Host "Fetching all ConfigMgr task sequences modified after: $($Date)" -ForegroundColor Cyan
            }
        } catch [System.Management.Automation.ActionPreferenceStopException] {
            Write-Host "INFO: Found no records of a previous export. Running a full export." -ForegroundColor Cyan            
        } catch {
            Write-Warning "An unknown error occurred getting the last run time. Running a full export."
        }
    } else {
        $Date = $AfterDate
        Write-Host "Getting task sequences modified after: $($Date)" -ForegroundColor Cyan
    }
} elseif ($FullExport -eq $true) {
    Write-Host "Full Export specificied. Ignoring last run date, will fetch every task sequence." -ForegroundColor Cyan
}


# Fetch every TS that has been modified since a specified date above unless specific TS is requested
if ([string]::IsNullOrWhiteSpace($SpecificTaskSequence)) {
    Set-content -Path $LastRunPath -Value $(Get-date)
    $TaskSequences = Get-CMTaskSequence | Where-Object {$_.LastRefreshTime -gt $Date}
} else {
    Write-Host "INFO: Specific task sequence is requested, will not get or set a Last Run date" -ForegroundColor Yellow
    $TaskSequences = Get-CMTaskSequence -name $SpecificTaskSequence
}

# Initialize counts
$count = $Tasksequences.Count
$currentCount = 0;
Write-Host "Found " -NoNewline; Write-host $($count) -NoNewline -ForegroundColor Yellow; Write-host " task sequences to process"

# Loop through all retrieved task sequences
if (-not [string]::IsNullOrEmpty($TaskSequences)){

    # Reset the LastUpdatedTS file, this will be appended with all updated task sequences as we go.
    Set-content -Path $LastUpdatedTSPath -Value "Updated: $(Get-date)"

    # Loop through all retrieved task sequences. The label is used to skip excluded task sequences.
    :TSLoop foreach ($TaskSequence in $TaskSequences){
        # Increment the count to show the progress, print the TS name
        $currentCount++
        Write-Host "$($currentCount)/$($count) - $($TaskSequence.Name)" -ForegroundColor Cyan

        ##############################
        # Task sequence Path         #
        ##############################
        Write-Host "$($currentCount)/$($count) - Fetching path"
        $PackageID = $($Tasksequence.PackageID)
        $TSPath = ((Get-WmiObject -ComputerName $SMSProvider -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_TaskSequencePackage WHERE PackageID = '$PackageID'").objectpath) -replace "/","\"       

        # Check if the path is excluded, skip this TS if it is
        if ([string]::IsNullOrWhiteSpace($ExclusionPaths) -eq $false) {
            foreach ($ExclusionPath in $ExclusionPaths) {
                if ($TSPath -like "$ExclusionPath") {
                    Write-Host "$($currentCount)/$($count) - Current task sequence is in an excluded path, will not export" -ForegroundColor Yellow
                    Write-Host "-------------------------------------------------"
                    Write-Host ""
                    continue TSLoop
                }
            }
        } 

        ##############################
        # General details            #
        ##############################
        Write-Host "$($currentCount)/$($count) - Fetching general details"                         
        $TSDetails = @()
        $TSDetails += "Package ID: $($TaskSequence.PackageID)"
        $TSBootImage = $TaskSequence.BootImageID
        If([string]::IsNullOrEmpty($TSBootImage)){
            $BootImage="None"
        } else{                    
            $BootImage = (Get-CMBootImage -id $TSBootImage -ErrorAction Ignore).Name
        }
        $TSDetails += "Task Sequence Boot Image: $BootImage"
        $TSRefs = $TaskSequence.References.Package

        $OSImages = Get-CMOperatingSystemImage
        If([string]::IsNullOrEmpty($TSRefs) -or [string]::IsNullOrEmpty($OSImages)){
            $TSOSImage="None"
        }else{
            foreach ($Ref in $TSRefs){
                If($Ref -in $OSImages.PackageID){
                    $TSOSImage = (Get-CMOperatingSystemImage -id $Ref).Name
                }
            }
        }
        If([string]::IsNullOrEmpty($TSOSImage)){$TSOSImage="None"}
        $TSDetails += "Task Sequence Operating System Image: $TSOSImage"            

        ##############################
        # Task sequence references   #
        ##############################
        Write-Host "$($currentCount)/$($count) - Fetching references"
        $TSReferences = Get-WmiObject -Namespace "Root\SMS\site_$($SiteCode)" -Query "SELECT * FROM SMS_TaskSequencePackageReference_Flat where PackageID = `'$($TaskSequence.PackageID)`'" -ComputerName $SMSProvider
        $TotalSize = 0
        $TotalTsRefs = 0
        $TotalTsRefs = $TSReferences.count
        $TSRefInfo = @()
        foreach($ref in $TSReferences){
            switch ($ref.ObjectType){
                0 {
                    $PackageType = "Package"
                }
                3 {
                    $PackageType = "Driver Package"
                }
                4 {
                    $PackageType = "Task Sequence"
                }
                257{
                    $PackageType = "OS Image"
                }
                258{
                    $PackageType = "Boot Image"
                }
                259{
                    $PackageType = "OS Upgrade Package"
                }
                512{
                    $PackageType = "Application"
                }
                default{
                    $PackageType = $ref.ObjectType
                }
            }
            $SizeMB = [math]::Round($ref.SourceSize/1024)
            $TotalSize = $TotalSize + $SizeMB
            If ($SizeMB -eq 0){
                $SizeMB = "<1 MB"
            } else {
                $SizeMB = "$SizeMB MB"
            }
            $TSRefInfo += New-Object -TypeName PSObject -Property @{'Name'="$($ref.ObjectName)";'Type'="$PackageType";'Package ID'="$($ref.RefPackageID)";'Size'="$SizeMB"}
        }

        $TSRefInfo = $TSRefInfo | Select-Object Type,Name,'Package ID',Size            

        ##############################
        # Output refs to file        #
        ##############################
        Write-Host "$($currentCount)/$($count) - Writing references to file"

        # If the output path is not the default we need to append the filesystem tag to support running from the CMDrive
        if ($OutputPath -ne (Join-path $RepoPath -ChildPath "CMTaskSequences")) {
            $ExecutionContext.SessionState.Path.Combine('FileSystem::', $OutputPath) | out-null
        }

        $ExportFilePath = Join-Path -Path (Join-Path -Path $OutPutPath -ChildPath $TSPath) -ChildPath "$($TaskSequence.Name).txt"
        $mainOutFileParams = @{
            FilePath = $ExportFilePath
            Append = $true
            Encoding = 'utf8'
        }

        # Create the export path if it does not exist
        if (-not(test-path $ExportFilePath)) {
            New-item -ItemType Directory -Path (Join-path $OutPutPath -ChildPath $TSPath) -ErrorAction SilentlyContinue | out-null
        }

        Set-content $ExportFilePath -Value "Task Sequence Exported at $(Get-date)" -Force
        
        Out-File -InputObject "$($TaskSequence.Name)" @mainOutFileParams
        Out-File -InputObject $TSDetails @mainOutFileParams

        Out-File -InputObject "" @mainOutFileParams
        Out-File -InputObject "Task Sequence References" @mainOutFileParams

        Out-File -InputObject @("This task sequence references $TotalTsRefs packages.","The referenced packages total $TotalSize MB.") @mainOutFileParams
        If ($TotalTsRefs -ne 0){
            $tsrefinfo | Out-File @mainOutFileParams -Width 1024               
        }

        ##############################
        # Task sequence steps        #
        ##############################
        Write-Host "$($currentCount)/$($count) - Fetching and writing steps"              
        Out-File -InputObject "" @mainOutFileParams
        Out-File -InputObject "Task Sequence Steps"  @mainOutFileParams
        $Sequence = $Null
        $Sequence = ([xml]$TaskSequence.Sequence).sequence

        # This adds "step" to the output, but if we have steps it will break subsequent commits because each order change would make everything below it a new change
        #$c = 0
        #foreach ($Step in $AllSteps){$c++;$Step|Add-Member -MemberType NoteProperty -Name 'Step' -Value $c}
        #$AllSteps = $AllSteps | Select-Object 'Step','Group Name','Step Name','Status','Continue on Error','Conditions','Action','Description'

        $AllSteps = Read-TSSteps -Sequence $Sequence
        $AllSteps = $AllSteps | Select-Object 'Group Name','Step Name','Status','Continue on Error','Conditions','Action','Description', 'ScriptHash'            
        If ($AllSteps.count -gt 0){
            $allsteps = $allsteps | ForEach-Object {
                if ($_.conditions) {
                    $_.conditions = Convert-HTMLTags -s $_.conditions
                }
                if ($_.description) {
                    $_.description = Convert-HTMLTags -s $_.description
                }                    
                $_                
            }   
            $AllSteps | Format-PropertyValues -MaxWidth 100 | Format-Table -Wrap | Out-File @mainOutFileParams -width 800         
        }else{
            Out-File -InputObject "This task sequence contains no steps" @mainOutFileParams
        }
        "$($TaskSequence.Name) was last updated $($TaskSequence.LastRefreshTime)" | out-file $LastUpdatedTSPath -Encoding UTF8 -Append
        Remove-Variable TSOSImage -ErrorAction Ignore
        Remove-Variable BootImage -ErrorAction Ignore
        Write-Host "$($currentCount)/$($count) - Finished with $($TaskSequence.Name)" -ForegroundColor Green
        Write-Host "-------------------------------------------------"
        Write-Host ""
    } 

    Write-Host "Finished looping through all $($count) task sequence(s)" -ForegroundColor Green
    ###############################
    # Push to bitbucket if needed #
    ###############################
    if ($UseGit -eq $True) {
        Write-Host "Will try to push to remote using git push -u $($GitPushRemote)"
        Set-Location $RepoPath
        git checkout $GitBranch         
        git add --all
        if ($CustomCommitMessage) {
            $CommitMessage = $CustomCommitMessage
        } else {
            $CommitMessage = ($Tasksequences.name -join "`n")
        }    
        git commit -m "$CommitMessage"
        git push
    } else {
        Write-Host "INFO: UseGit not specified, will not push to remote" -ForegroundColor Yellow
    }
} else{
    Set-content -Path $LastErrorPath -Value "$(Get-date)"        
    if ($error) {
        $error[-1].Exception.ErrorRecord | out-file $LastErrorPath -Append -Encoding UTF8
    } else {
        out-file $LastErrorPath -Append -Encoding UTF8 -InputObject "No task sequence has been modified since specified date: $Date"
    }       
}

set-location $RepoPath
