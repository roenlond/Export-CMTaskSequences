function Format-PropertyValues {
    param(
        [Parameter(ValueFromPipeline=$true)]
        $InputObject,
        [int]$MaxWidth
    )

    process {
        $InputObject | ForEach-Object {
            $item = $_
            $item.PSObject.Properties | ForEach-Object {
                if ($_.Value -is [string] -and $_.Value.Length -gt $MaxWidth) {
                    $_.Value = $_.Value.Substring(0, $MaxWidth)
                }
            }
            $item
        }
    }
}

Function Get-SiteCode
{
  $wqlQuery = 'SELECT * FROM SMS_ProviderLocation'
  $a = Get-WmiObject -Query $wqlQuery -Namespace 'root\sms' -ComputerName $SMSProvider -ErrorAction Stop
  $a | ForEach-Object {
    if($_.ProviderForLocalSite)
    {
      $script:SiteCode = $_.SiteCode
    }
  }
  return $SiteCode
}


Function Convert-HTMLTags {
    <#
    .SYNOPSIS
    Converts custom formatted HTML tags to standard PowerShell formatting
    .PARAMETER InputString
    The string or array of strings of text to parse
    .EXAMPLE
    Convert-HTMLTags -InputString 'This is --B--bold text--/B--'
    .NOTES
    ========== Change Log History ==========
    - 2021/02/02 by Chad@ChadsTech.net / Chad.Simmons@CatapultSystems.com - Moved to function
    - 2018/03/07 by Paul Wetter - Created
    #>
    param ([Parameter(Mandatory = $true)][Alias('InputString')][string[]]$S)
    $S = $S -replace '--CRLF--',"`n" 
    $S = $S -replace '--TAB--',"`t"
    $S = $S -replace '--B--','`e[1m'
    $S = $S -replace '--/B--','`e[0m'
    $S = $S -replace '--I--','`e[3m'
    $S = $S -replace '--/I--','`e[0m'
    $S = $S -replace '--U--','`e[4m'
    $S = $S -replace '--/U--','`e[0m'
    $S = $S -replace '--CBOX--','[X]'
    $S = $S -replace '--UNCBOX--','[ ]'
    $S = $S -replace '--BULLET--','-'
    Return $S
}

#small function that will convert utc time to local time, option to ignore daylight savings time.
function Convert-UTCtoLocal{
    param(
        [parameter(Mandatory=$true)]
        [String]$UTCTimeString,
        [parameter(Mandatory=$false)]
        [Switch]$IgnoreDST
    )
    $UTCTime = ($UTCTimeString.Split('.'))[0]
    $dt = ([datetime]::ParseExact($UTCTime,'yyyyMMddhhmmss',$null))
    if ($IgnoreDST){
        $dt+([System.TimeZoneInfo]::Local).BaseUtcOffset
    }else{
        $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
        $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
        [System.TimeZoneInfo]::ConvertTimeFromUtc($dt, $TZ)
    }
}


##Recursively processes through all the conditions in a task sequence step.
function Read-TSConditions{
    param ($condition,$Level = 0)
    $prefix = ""
    for ($x=0; $x -lt $Level; $x++){$prefix="--TAB--" + $prefix}
    If($condition.osConditionGroup){
        $OSCondition = $condition.osConditionGroup.osExpressionGroup.name -join ", $($condition.osConditionGroup.type) "
        "$($prefix)Operating System Equals: $OSCondition"
        Remove-Variable OSCondition -ErrorAction Ignore
    }
    If($condition.expression){
        $expressions = $condition.expression
        foreach ($expression in $expressions){
            #$expression
            switch ($expression.type){
                'SMS_TaskSequence_WMIConditionExpression' {
                    foreach ($pair in $expression.variable){
                        if ($pair.name -eq 'Query'){
                            if ($pair.'#text' -like 'SELECT OsLanguage FROM Win32_OperatingSystem WHERE OsLanguage*'){
                                $lang=[int](($pair.'#text').Split('='))[1].Trim("`'")
                                "$($prefix)Operating System Language: $(([System.Globalization.Cultureinfo]::GetCultureInfo($lang)).DisplayName) ($lang)"
                            }else{
                                "$($prefix)WMI Query: " + $pair.'#text'
                            }
                        }
                    }
                }
                'SMS_TaskSequence_VariableConditionExpression'{
                    foreach ($pair in $expression.variable){
                        if ($pair.name -eq 'Operator'){
                            $ExpOperator = $pair.'#text'
                        }
                        if ($pair.name -eq 'Value'){
                            $ExpValue = $pair.'#text'
                        }
                        if ($pair.name -eq 'Variable'){
                            $ExpVariable = $pair.'#text'
                        }
                    }
                    "$($prefix)Task Sequence Variable: $ExpVariable $ExpOperator $ExpValue"
                    Remove-Variable ExpVariable,ExpOperator,ExpValue -ErrorAction Ignore
                }
                'SMS_TaskSequence_FileConditionExpression'{
                    If(('Path' -in ($expression.variable).name) -and ('DateTimeOperator' -notin ($expression.variable).name) -and ('VersionOperator' -notin ($expression.variable).name)){
                        "$($prefix)File Exists: " + ($expression.variable).'#text'
                    }else{
                        foreach ($pair in $expression.variable){
                            switch ($pair.name){
                                'DateTime'{$FileDate = Convert-UTCtoLocal($pair.'#text')}
                                'DateTimeOperator'{$FileDateOperator = $pair.'#text'}
                                'Path'{$FilePath = $pair.'#text'}
                                'Version'{$FileVersion = $pair.'#text'}
                                'VersionOperator'{$FileVersionOperator = $pair.'#text'}
                            }
                        }
                        #'DateTimeOperator' -in ($expression.variable).name
                        #'VersionOperator' -in ($expression.variable).name
                        "$($prefix)File: $FilePath     File Version: $FileVersionOperator $FileVersion     File Date: $FileDateOperator $FileDate"
                        Remove-Variable FileDate,FileDateOperator,FilePath,FileVersion,FileVersionOperator -ErrorAction Ignore
                    }
                }
                'SMS_TaskSequence_FolderConditionExpression'{
                    If(('Path' -in ($expression.variable).name) -and ('DateTimeOperator' -notin ($expression.variable).name)){
                        "$($prefix)Folder Exists: " + ($expression.variable).'#text'
                    }else{
                        foreach ($pair in $expression.variable){
                            switch ($pair.name){
                                'DateTime'{$FolderDate = Convert-UTCtoLocal($pair.'#text')}
                                'DateTimeOperator'{$FolderDateOperator = $pair.'#text'}
                                'Path'{$FolderPath = $pair.'#text'}
                            }
                        }
                        #'DateTimeOperator' -in ($expression.variable).name
                        #'VersionOperator' -in ($expression.variable).name
                        "$($prefix)Folder: $FolderPath     Folder Date: $FolderDateOperator $FolderDate"
                        Remove-Variable FolderPath,FolderDateOperator,FolderDate -ErrorAction Ignore
                    }
                }
                'SMS_TaskSequence_RegistryConditionExpression'{
                    foreach ($pair in $expression.variable){
                        Switch ($pair.name){
                            'Operator'{$RegOperator = $pair.'#text'}
                            'KeyPath'{$RegKeyPath = $pair.'#text'}
                            'Data'{$RegData = $pair.'#text'}
                            'Value'{$RegValue = $pair.'#text'}
                            'Type'{$RegType = $pair.'#text'}
                        }
                    }
                    "$($prefix)Registry Value: $RegKeyPath $RegValue ($RegType) $RegOperator $RegData"
                    Remove-Variable RegKeyPath,RegValue,RegType,RegOperator,RegData -ErrorAction Ignore
                }
                'SMS_TaskSequence_SoftwareConditionExpression'{
                    foreach ($pair in $expression.variable){
                        Switch ($pair.name){
                            'Operator'{$AppOperator = $pair.'#text'}
                            'ProductCode'{$AppProductCode = $pair.'#text'}
                            'ProductName'{$AppProductName = $pair.'#text'}
                            #'UpgradeCode'{$AppUpgradeCode = $pair.'#text'}
                            'Version'{$AppVersion = $pair.'#text'}
                        }
                    }
                    If ($AppOperator -eq 'AnyVersion'){
                        "$($prefix)Installed Software: Any Version of `"$AppProductName`""
                    }else{
                        "$($prefix)Installed Software: Exact Version of `"$AppProductName`", Version: $AppVersion, Product Code: $AppProductCode"
                    }
                    Remove-Variable AppOperator,AppProductCode,AppProductName,AppUpgradeCode,AppVersion -ErrorAction Ignore
                }
            }
        }
    }
    If($condition.operator){
        Switch($condition.operator.type){
        'or'{"$($prefix)-If any of these conditions are true"}
        'and'{"$($prefix)-If all of these conditions are true"}
        'not'{"$($prefix)-If none of these conditions are true"}
        }
        $Level = $Level + 1
        Read-TSConditions -condition $condition.operator -Level $Level
    }
}


##Recursively processes through all the steps in a task sequence
Function Read-TSSteps {
    param ($Sequence, $GroupName)
    foreach ($node in $Sequence.ChildNodes) {
        switch ($node.localname) {
            'step' {
                if (-not [string]::IsNullOrEmpty($node.Description)) {
                    $StepDescription = "$($node.Description)"
                }
                if ($node.condition) {
                    $Conditions = (Read-TSConditions -condition $node.condition) -join "--CRLF--"
                }
                try {
                    if (-not [string]::IsNullOrEmpty($node.disable)) {
                        $StepStatus = 'Disabled'
                    }
                    else {
                        $StepStatus = 'Enabled'
                    }
                }   
                catch [System.Management.Automation.PropertyNotFoundException] {
                    $StepStatus = 'Enabled'
                }
                If ($node.continueOnError -eq "true") {
                    $StepContinueError = "Yes"
                }
                else {
                    $StepContinueError = "No"
                }

                if ($node.Action -eq "OSDRunPowershellScript.exe") {
                    try { 
                        $ScriptHash = (($node.defaultVarList.variable | Where-Object name -eq "OSDRunPowerShellScriptSourceScript").'#Text').GetHashCode()
                    }
                    catch {}
                }

                if ($GroupName) {
                    $TSStep = New-Object -TypeName psobject -Property @{
                        'Group Name' = "$GroupName"; 
                        'Step Name' = "$($node.Name)"; 
                        'Description' = "$StepDescription"; 
                        'Action' = "$($node.Action)"; 
                        'Continue on Error' = "$StepContinueError"; 
                        'Status' = "$StepStatus"; 
                        'Conditions' = "$Conditions"; 
                        'ScriptHash' = "$ScriptHash"}
                }
                else {
                    $TSStep = New-Object -TypeName psobject -Property @{
                    'Group Name' = "N/A";
                    'Step Name' = "$($node.Name)";
                    'Description' = "$StepDescription"; 
                    'Action' = "$($node.Action)"; 
                    'Continue on Error' = "$StepContinueError"; 
                    'Status' = "$StepStatus"; 
                    'Conditions' = "$Conditions"; 
                    'ScriptHash' = "$ScriptHash" 
                }
                }
                Remove-Variable Conditions -ErrorAction Ignore
                Remove-Variable StepDescription -ErrorAction Ignore
                Remove-Variable ScriptHash -ErrorAction Ignore
                $TSStep
            }
            'subtasksequence' {
                foreach ($item in $node.defaultVarList.variable) {
                    Switch ($item.property) {
                        'TsName' { $NestTSName = $item.'#text' }
                        'TsPackageID' { $NestTSPackage = $item.'#text' }
                    }
                }
                if (-not [string]::IsNullOrEmpty($node.Description)) {
                    $StepDescription = "$($node.Description)"
                }
                if ($node.condition) {
                    $Conditions = (Read-TSConditions -condition $node.condition) -join "--CRLF--"
                }
                try {
                    if (-not [string]::IsNullOrEmpty($node.disable)) {
                        $StepStatus = 'Disabled'
                    }
                    else {
                        $StepStatus = 'Enabled'
                    }
                }   
                catch [System.Management.Automation.PropertyNotFoundException] {
                    $StepStatus = 'Enabled'
                }
                If ($node.continueOnError -eq "true") {
                    $StepContinueError = "Yes"
                }
                else {
                    $StepContinueError = "No"
                }

                if ($node.Action -eq "OSDRunPowershellScript.exe") {
                    try { 
                        $ScriptHash = (($node.step.defaultVarList.variable | Where-Object name -eq "OSDRunPowerShellScriptSourceScript").'#Text').GetHashCode()
                    }
                    catch {}
                }

                if ($GroupName) {
                    #"$($GroupName):  $($node.name) $($node.action)"
                    $TSStep = New-Object -TypeName psobject -Property @{
                        'Group Name' = "$GroupName"; 
                        'Step Name' = "$($node.Name)"; 
                        'Description' = "$StepDescription"; 
                        'Action' = "Run Task Sequence ($($node.Action)):--CRLF--$NestTSName ($NestTSPackage)"; 
                        'Continue on Error' = "$StepContinueError"; 
                        'Status' = "$StepStatus"; 
                        'Conditions' = "$Conditions"; 
                        'ScriptHash' = "" 
                    }
                }
                else {
                    $TSStep = New-Object -TypeName psobject -Property @{
                        'Group Name' = "N/A"; 
                        'Step Name' = "$($node.Name)"; 
                        'Description' = "$StepDescription"; 
                        'Action' = "Run Task Sequence ($($node.Action)):--CRLF--$NestTSName ($NestTSPackage)"; 
                        'Continue on Error' = "$StepContinueError"; 
                        'Status' = "$StepStatus"; 
                        'Conditions' = "$Conditions"; 
                        'ScriptHash' = "" 
                    }
                    #"$($node.name) $($node.action)"
                }
                Remove-Variable Conditions -ErrorAction Ignore
                Remove-Variable StepDescription -ErrorAction Ignore
                Remove-Variable ScriptHash -ErrorAction Ignore
                $TSStep
            }
            'group' {
                $TSStepNumber++
                if ($node.condition) {
                    $Conditions = (Read-TSConditions -condition $node.condition) -join "--CRLF--"
                }
                if (-not [string]::IsNullOrEmpty($node.Description)) {
                    $StepDescription = "$($node.Description)"
                }
                try {
                    if (-not [string]::IsNullOrEmpty($node.disable)) {
                        $StepStatus = 'Disabled'
                    }
                    else {
                        $StepStatus = 'Enabled'
                    }
                }   
                catch [System.Management.Automation.PropertyNotFoundException] {
                    $StepStatus = 'Enabled'
                }
                If ($node.continueOnError -eq "true") {
                    $StepContinueError = "Yes"
                }
                else {
                    $StepContinueError = "No"
                }

                if ($node.Action -eq "OSDRunPowershellScript.exe") {
                    try { 
                        $ScriptHash = (($node.step.defaultVarList.variable | Where-Object name -eq "OSDRunPowerShellScriptSourceScript").'#Text').GetHashCode()
                    }
                    catch {}
                }

                #"Group: $($node.Name)"
                $TSStep = New-Object -TypeName psobject -Property @{
                    'Group Name' = "$($node.Name)"; 
                    'Step Name' = "N/A"; 
                    'Description' = "$StepDescription"; 
                    'Action' = "N/A"; 
                    'Continue on Error' = "$StepContinueError"; 
                    'Status' = "$StepStatus"; 
                    'Conditions' = "$Conditions"; 
                    'ScriptHash' = "" }
                Remove-Variable Conditions -ErrorAction Ignore
                Remove-Variable StepDescription -ErrorAction Ignore
                Remove-Variable ScriptHash -ErrorAction Ignore
                $TSStep
                Read-TSSteps -Sequence $node -GroupName "$($node.Name)" -TSSteps $TSSteps -StepCounter $TSStepNumber
            }
            default {}
        }
    }
}



